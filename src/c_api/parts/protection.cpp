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
#include "io/protection_xml.h"
#include "io/sqref_text.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::store_cstr;

namespace {

std::string Text(const char* text) {
  return text == nullptr ? std::string() : std::string(text);
}

formulon::ProtectedRange ToModel(const fm_protected_range_t& range) {
  formulon::ProtectedRange out;
  out.name = Text(range.name);
  out.sqref = Text(range.sqref);
  out.password = Text(range.password);
  out.algorithm_name = Text(range.algorithm_name);
  out.hash_value = Text(range.hash_value);
  out.salt_value = Text(range.salt_value);
  out.spin_count = range.spin_count;
  out.security_descriptor = Text(range.security_descriptor);
  return out;
}

// A name is unique per sheet; `skip` is the entry being replaced.
fm_status_t CheckRange(const std::vector<formulon::ProtectedRange>& ranges, const formulon::ProtectedRange& range,
                       std::size_t skip, const char* fn) {
  std::vector<formulon::MergeRange> parsed;
  if (range.name.empty() || !formulon::io::parse_sqref_text(range.sqref, parsed)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(fn) + ": empty name or unparseable sqref").c_str(), "sqref=" + range.sqref);
  }
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    if (i != skip && ranges[i].name == range.name) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(fn) + ": name already in use").c_str(), "name=" + range.name);
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_protection(fm_workbook_t* wb, fm_workbook_protection_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_protection: NULL argument");
  }
  const formulon::WorkbookProtection p =
      formulon::io::parse_workbook_protection(wb->workbook().workbook_protection_xml());
  wb->read_scratch.clear();
  auto& store = wb->read_scratch;
  out->lock_structure = p.lock_structure ? 1 : 0;
  out->lock_windows = p.lock_windows ? 1 : 0;
  out->lock_revision = p.lock_revision ? 1 : 0;
  out->workbook_spin_count = p.workbook_spin_count;
  out->revisions_spin_count = p.revisions_spin_count;
  out->workbook_password = store_cstr(store, p.workbook_password);
  out->workbook_algorithm_name = store_cstr(store, p.workbook_algorithm_name);
  out->workbook_hash_value = store_cstr(store, p.workbook_hash_value);
  out->workbook_salt_value = store_cstr(store, p.workbook_salt_value);
  out->revisions_password = store_cstr(store, p.revisions_password);
  out->revisions_algorithm_name = store_cstr(store, p.revisions_algorithm_name);
  out->revisions_hash_value = store_cstr(store, p.revisions_hash_value);
  out->revisions_salt_value = store_cstr(store, p.revisions_salt_value);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_protection(fm_workbook_t* wb, const fm_workbook_protection_t* protection) {
  clear_last_error();
  if (wb == nullptr || protection == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_protection: NULL argument");
  }
  formulon::WorkbookProtection p;
  p.lock_structure = protection->lock_structure != 0;
  p.lock_windows = protection->lock_windows != 0;
  p.lock_revision = protection->lock_revision != 0;
  p.workbook_password = Text(protection->workbook_password);
  p.workbook_algorithm_name = Text(protection->workbook_algorithm_name);
  p.workbook_hash_value = Text(protection->workbook_hash_value);
  p.workbook_salt_value = Text(protection->workbook_salt_value);
  p.workbook_spin_count = protection->workbook_spin_count;
  p.revisions_password = Text(protection->revisions_password);
  p.revisions_algorithm_name = Text(protection->revisions_algorithm_name);
  p.revisions_hash_value = Text(protection->revisions_hash_value);
  p.revisions_salt_value = Text(protection->revisions_salt_value);
  p.revisions_spin_count = protection->revisions_spin_count;
  formulon::Workbook& book = wb->workbook();
  book.set_workbook_protection_xml(formulon::io::apply_workbook_protection(book.workbook_protection_xml(), p));
  return 0;
}

extern "C" fm_status_t fm_sheet_protected_range_count(fm_workbook_t* wb, uint32_t sheet, uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_protected_range_count: NULL out_count");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_protected_range_count"); rc != 0) {
    return rc;
  }
  *out_count =
      static_cast<uint32_t>(formulon::io::read_protected_ranges(wb->workbook().sheet(sheet).raw_extensions()).size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_protected_range(fm_workbook_t* wb, uint32_t sheet, uint32_t index,
                                                    fm_protected_range_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_protected_range: NULL out");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_protected_range"); rc != 0) {
    return rc;
  }
  const auto ranges = formulon::io::read_protected_ranges(wb->workbook().sheet(sheet).raw_extensions());
  if (auto rc = check_index(index, ranges.size(), "fm_sheet_get_protected_range", "index"); rc != 0) {
    return rc;
  }
  const formulon::ProtectedRange& r = ranges[index];
  wb->read_scratch.clear();
  auto& store = wb->read_scratch;
  out->spin_count = r.spin_count;
  out->name = store_cstr(store, r.name);
  out->sqref = store_cstr(store, r.sqref);
  out->password = store_cstr(store, r.password);
  out->algorithm_name = store_cstr(store, r.algorithm_name);
  out->hash_value = store_cstr(store, r.hash_value);
  out->salt_value = store_cstr(store, r.salt_value);
  out->security_descriptor = store_cstr(store, r.security_descriptor);
  return 0;
}

extern "C" fm_status_t fm_sheet_add_protected_range(fm_workbook_t* wb, uint32_t sheet,
                                                    const fm_protected_range_t* range) {
  clear_last_error();
  if (range == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_add_protected_range: NULL range");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_protected_range"); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  formulon::ProtectedRange added = ToModel(*range);
  if (auto rc = CheckRange(ranges, added, ranges.size(), "fm_sheet_add_protected_range"); rc != 0) {
    return rc;
  }
  ranges.push_back(std::move(added));
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}

extern "C" fm_status_t fm_sheet_update_protected_range(fm_workbook_t* wb, uint32_t sheet, uint32_t index,
                                                       const fm_protected_range_t* range) {
  clear_last_error();
  if (range == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_update_protected_range: NULL range");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_update_protected_range"); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  if (auto rc = check_index(index, ranges.size(), "fm_sheet_update_protected_range", "index"); rc != 0) {
    return rc;
  }
  formulon::ProtectedRange updated = ToModel(*range);
  if (auto rc = CheckRange(ranges, updated, index, "fm_sheet_update_protected_range"); rc != 0) {
    return rc;
  }
  ranges[index] = std::move(updated);
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_protected_range(fm_workbook_t* wb, uint32_t sheet, uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_protected_range"); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  if (auto rc = check_index(index, ranges.size(), "fm_sheet_remove_protected_range", "index"); rc != 0) {
    return rc;
  }
  ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(index));
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}
