//
// C ABI - workbook protection and per-sheet protected ranges. Both are
// views over XML the model retains; see `io/protection_xml.h`.

#include "protection.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "io/protection_xml.h"
#include "io/sqref_text.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_rect_list;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_rects;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::TextStore;

namespace {

formulon::ProtectedRange ToModel(const fm_protected_range_t& range) {
  formulon::ProtectedRange out;
  out.name = text_or_empty(range.name);
  std::vector<formulon::MergeRange> ranges;
  ranges.reserve(range.range_count);
  for (std::size_t i = 0; i < range.range_count; ++i) {
    const fm_cell_rect_t& r = range.ranges[i];
    ranges.push_back({r.first_row, r.first_col, r.last_row, r.last_col});
  }
  out.sqref = formulon::io::format_sqref_text(ranges);
  out.password = text_or_empty(range.legacy_password_hash);
  out.algorithm_name = text_or_empty(range.algorithm_name);
  out.hash_value = text_or_empty(range.hash_value);
  out.salt_value = text_or_empty(range.salt_value);
  out.spin_count = range.spin_count;
  out.security_descriptor = text_or_empty(range.security_descriptor);
  return out;
}

// Every text is UTF-8 and a name is unique per sheet; `skip` is the entry being replaced.
fm_status_t CheckRange(const std::vector<formulon::ProtectedRange>& ranges, const formulon::ProtectedRange& range,
                       std::size_t skip) {
  if (auto rc = check_utf8({range.name.c_str(), range.password.c_str(), range.algorithm_name.c_str(),
                            range.hash_value.c_str(), range.salt_value.c_str(), range.security_descriptor.c_str()});
      rc != 0) {
    return rc;
  }
  if (range.name.empty()) {
    return invalid_argument("empty name");
  }
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    if (i != skip && ranges[i].name == range.name) {
      return invalid_argument("name already in use", "name=" + range.name);
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_workbook_protection(const fm_workbook_t* wb, fm_workbook_protection_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::WorkbookProtection p =
      formulon::io::parse_workbook_protection(wb->workbook().workbook_protection_xml());
  TextStore& store = fresh_read_scratch(wb);
  out->lock_structure = p.lock_structure ? 1 : 0;
  out->lock_windows = p.lock_windows ? 1 : 0;
  out->lock_revision = p.lock_revision ? 1 : 0;
  out->workbook_spin_count = p.workbook_spin_count;
  out->revisions_spin_count = p.revisions_spin_count;
  out->workbook_legacy_password_hash = store_cstr(store, p.workbook_password);
  out->workbook_algorithm_name = store_cstr(store, p.workbook_algorithm_name);
  out->workbook_hash_value = store_cstr(store, p.workbook_hash_value);
  out->workbook_salt_value = store_cstr(store, p.workbook_salt_value);
  out->revisions_legacy_password_hash = store_cstr(store, p.revisions_password);
  out->revisions_algorithm_name = store_cstr(store, p.revisions_algorithm_name);
  out->revisions_hash_value = store_cstr(store, p.revisions_hash_value);
  out->revisions_salt_value = store_cstr(store, p.revisions_salt_value);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_workbook_protection(fm_workbook_t* wb,
                                                           const fm_workbook_protection_t* protection) {
  clear_last_error();
  if (wb == nullptr || protection == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*protection); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({protection->workbook_legacy_password_hash, protection->workbook_algorithm_name,
                            protection->workbook_hash_value, protection->workbook_salt_value,
                            protection->revisions_legacy_password_hash, protection->revisions_algorithm_name,
                            protection->revisions_hash_value, protection->revisions_salt_value});
      rc != 0) {
    return rc;
  }
  formulon::WorkbookProtection p;
  p.lock_structure = protection->lock_structure != 0;
  p.lock_windows = protection->lock_windows != 0;
  p.lock_revision = protection->lock_revision != 0;
  p.workbook_password = text_or_empty(protection->workbook_legacy_password_hash);
  p.workbook_algorithm_name = text_or_empty(protection->workbook_algorithm_name);
  p.workbook_hash_value = text_or_empty(protection->workbook_hash_value);
  p.workbook_salt_value = text_or_empty(protection->workbook_salt_value);
  p.workbook_spin_count = protection->workbook_spin_count;
  p.revisions_password = text_or_empty(protection->revisions_legacy_password_hash);
  p.revisions_algorithm_name = text_or_empty(protection->revisions_algorithm_name);
  p.revisions_hash_value = text_or_empty(protection->revisions_hash_value);
  p.revisions_salt_value = text_or_empty(protection->revisions_salt_value);
  p.revisions_spin_count = protection->revisions_spin_count;
  formulon::Workbook& book = wb->workbook();
  book.set_workbook_protection_xml(formulon::io::apply_workbook_protection(book.workbook_protection_xml(), p));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_protected_range_count(const fm_workbook_t* wb, size_t sheet_index,
                                                          size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = formulon::io::read_protected_ranges(wb->workbook().sheet(sheet_index).raw_extensions()).size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_protected_range(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                    fm_protected_range_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto ranges = formulon::io::read_protected_ranges(wb->workbook().sheet(sheet_index).raw_extensions());
  if (auto rc = check_index(index, ranges.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ProtectedRange& r = ranges[index];
  TextStore& store = fresh_read_scratch(wb);
  out->spin_count = r.spin_count;
  out->name = store_cstr(store, r.name);
  std::vector<formulon::MergeRange> rects;
  formulon::io::parse_selection_sqref_text(r.sqref, rects);
  out->ranges = publish_rects(wb, rects);
  out->range_count = static_cast<std::uint32_t>(rects.size());
  out->legacy_password_hash = store_cstr(store, r.password);
  out->algorithm_name = store_cstr(store, r.algorithm_name);
  out->hash_value = store_cstr(store, r.hash_value);
  out->salt_value = store_cstr(store, r.salt_value);
  out->security_descriptor = store_cstr(store, r.security_descriptor);
  return 0;
}

extern "C" fm_status_t fm_sheet_add_protected_range(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_protected_range_t* range, size_t* out_index) {
  clear_last_error();
  if (range == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*range); rc != 0) {
    return rc;
  }
  if (auto rc = check_rect_list(range->ranges, range->range_count); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet_index).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  formulon::ProtectedRange added = ToModel(*range);
  if (auto rc = CheckRange(ranges, added, ranges.size()); rc != 0) {
    return rc;
  }
  ranges.push_back(std::move(added));
  formulon::io::write_protected_ranges(raw, ranges);
  *out_index = ranges.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_sheet_update_protected_range(fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                       const fm_protected_range_t* range) {
  clear_last_error();
  if (range == nullptr) {
    return null_argument("range");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*range); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet_index).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  if (auto rc = check_index(index, ranges.size(), "index"); rc != 0) {
    return rc;
  }
  // NULL with a zero count keeps the stored text, unparsed tokens included.
  const bool keep_ranges = range->ranges == nullptr && range->range_count == 0;
  if (!keep_ranges) {
    if (auto rc = check_rect_list(range->ranges, range->range_count); rc != 0) {
      return rc;
    }
  }
  formulon::ProtectedRange updated = ToModel(*range);
  if (keep_ranges) {
    updated.sqref = ranges[index].sqref;
  }
  if (auto rc = CheckRange(ranges, updated, index); rc != 0) {
    return rc;
  }
  ranges[index] = std::move(updated);
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_protected_range(fm_workbook_t* wb, size_t sheet_index, size_t index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet_index).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  if (auto rc = check_index(index, ranges.size(), "index"); rc != 0) {
    return rc;
  }
  ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(index));
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}
