//
// C ABI - workbook protection and per-sheet protected ranges. Both are
// views over XML the model retains; see `io/protection_xml.h`.

#include "protection.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "io/protection_xml.h"
#include "io/sqref_text.h"
#include "sheet.h"
#include "sheet_name.h"
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
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::publish_rects;
using formulon::c_api::parts::set_binding_error;
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

// Every text is UTF-8 and a name is non-empty and collides with no other on the sheet (`sheet_names::write_collides`).
fm_status_t CheckRange(const std::vector<formulon::ProtectedRange>& ranges, const formulon::ProtectedRange& range) {
  if (auto rc = check_utf8({range.name.c_str(), range.password.c_str(), range.algorithm_name.c_str(),
                            range.hash_value.c_str(), range.salt_value.c_str(), range.security_descriptor.c_str()});
      rc != 0) {
    return rc;
  }
  if (range.name.empty()) {
    return invalid_argument("empty name");
  }
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    if (formulon::sheet_names::write_collides(ranges[i].name, range.name)) {
      return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "name already in use",
                               "name=" + range.name);
    }
  }
  return 0;
}

// An absent (empty) attribute reads NULL.
const char* optional_cstr(TextStore& store, const std::string& text) {
  return text.empty() ? nullptr : store_cstr(store, text);
}

// The position of the range `name` resolves to (`sheet_names::resolve_named`), or `ranges.size()` when none.
std::size_t FindRange(const std::vector<formulon::ProtectedRange>& ranges, const char* name) {
  const std::size_t found = formulon::sheet_names::resolve_named(ranges, name);
  return found == formulon::sheet_names::kNotFound ? ranges.size() : found;
}

fm_status_t RangeNotFound(const char* name) {
  return formulon::c_api::parts::not_found("no protected range has that name", "name=" + std::string(name));
}

// Fills `out` from `r`; every text view lives in the read scratch.
void FillRange(const fm_workbook_t* wb, const formulon::ProtectedRange& r, fm_protected_range_t* out) {
  TextStore& store = fresh_read_scratch(wb);
  out->spin_count = r.spin_count;
  out->name = store_cstr(store, r.name);
  std::vector<formulon::MergeRange> rects;
  formulon::io::parse_selection_sqref_text(r.sqref, rects);
  out->ranges = publish_rects(wb, rects);
  out->range_count = static_cast<std::uint32_t>(rects.size());
  out->legacy_password_hash = optional_cstr(store, r.password);
  out->algorithm_name = optional_cstr(store, r.algorithm_name);
  out->hash_value = optional_cstr(store, r.hash_value);
  out->salt_value = optional_cstr(store, r.salt_value);
  out->security_descriptor = optional_cstr(store, r.security_descriptor);
}

// One string field of `fm_protected_range_patch_t`: its op, value and model member.
struct StringPatch {
  fm_patch_op_t op;
  const char* value;
  std::string formulon::ProtectedRange::*member;
  bool clearable;
};

}  // namespace

extern "C" fm_status_t fm_workbook_get_workbook_protection(const fm_workbook_t* wb, int32_t* out_present,
                                                           fm_workbook_protection_t* out) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::string& xml = wb->workbook().workbook_protection_xml();
  if (xml.empty()) {
    const std::uint32_t struct_size = out->struct_size;
    *out = FM_STRUCT_INIT(fm_workbook_protection_t);
    out->struct_size = struct_size;
    *out_present = 0;
    return 0;
  }
  *out_present = 1;
  const formulon::WorkbookProtection p = formulon::io::parse_workbook_protection(xml);
  TextStore& store = fresh_read_scratch(wb);
  out->lock_structure = p.lock_structure ? 1 : 0;
  out->lock_windows = p.lock_windows ? 1 : 0;
  out->lock_revision = p.lock_revision ? 1 : 0;
  out->workbook_spin_count = p.workbook_spin_count;
  out->revisions_spin_count = p.revisions_spin_count;
  out->workbook_legacy_password_hash = optional_cstr(store, p.workbook_password);
  out->workbook_algorithm_name = optional_cstr(store, p.workbook_algorithm_name);
  out->workbook_hash_value = optional_cstr(store, p.workbook_hash_value);
  out->workbook_salt_value = optional_cstr(store, p.workbook_salt_value);
  out->revisions_legacy_password_hash = optional_cstr(store, p.revisions_password);
  out->revisions_algorithm_name = optional_cstr(store, p.revisions_algorithm_name);
  out->revisions_hash_value = optional_cstr(store, p.revisions_hash_value);
  out->revisions_salt_value = optional_cstr(store, p.revisions_salt_value);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_workbook_protection(fm_workbook_t* wb,
                                                           const fm_workbook_protection_t* protection) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (protection == nullptr) {
    wb->workbook().set_workbook_protection_xml(std::string());
    return 0;
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
  std::string xml = formulon::io::apply_workbook_protection(book.workbook_protection_xml(), p);
  if (xml.empty()) {
    // An element with every lock off is still present.
    xml = "<workbookProtection/>";
  }
  book.set_workbook_protection_xml(std::move(xml));
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

extern "C" fm_status_t fm_sheet_get_protected_range_at(const fm_workbook_t* wb, size_t sheet_index,
                                                       size_t protected_range_index, fm_protected_range_t* out) {
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
  if (auto rc = check_index(protected_range_index, ranges.size(), "protected_range_index"); rc != 0) {
    return rc;
  }
  FillRange(wb, ranges[protected_range_index], out);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_protected_range(const fm_workbook_t* wb, size_t sheet_index, const char* name,
                                                    int32_t* out_present, fm_protected_range_t* out) {
  clear_last_error();
  if (name == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument(name == nullptr ? "name" : (out_present == nullptr ? "out_present" : "out"));
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  const auto ranges = formulon::io::read_protected_ranges(wb->workbook().sheet(sheet_index).raw_extensions());
  const std::size_t i = FindRange(ranges, name);
  if (i == ranges.size()) {
    const std::uint32_t struct_size = out->struct_size;
    *out = FM_STRUCT_INIT(fm_protected_range_t);
    out->struct_size = struct_size;
    *out_present = 0;
    return 0;
  }
  FillRange(wb, ranges[i], out);
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_sheet_add_protected_range(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_protected_range_t* range, const char** out_name) {
  clear_last_error();
  if (out_name != nullptr) {
    *out_name = nullptr;
  }
  if (range == nullptr) {
    return null_argument("range");
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
  if (auto rc = CheckRange(ranges, added); rc != 0) {
    return rc;
  }
  ranges.push_back(std::move(added));
  formulon::io::write_protected_ranges(raw, ranges);
  if (out_name != nullptr) {
    *out_name = store_cstr(fresh_read_scratch(wb), ranges.back().name);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_update_protected_range(fm_workbook_t* wb, size_t sheet_index, const char* name,
                                                       const fm_protected_range_patch_t* patch) {
  clear_last_error();
  if (name == nullptr || patch == nullptr) {
    return null_argument(name == nullptr ? "name" : "patch");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  using formulon::ProtectedRange;
  const StringPatch strings[] = {
      {patch->name_op, patch->name, &ProtectedRange::name, false},
      {patch->legacy_password_hash_op, patch->legacy_password_hash, &ProtectedRange::password, true},
      {patch->algorithm_name_op, patch->algorithm_name, &ProtectedRange::algorithm_name, true},
      {patch->hash_value_op, patch->hash_value, &ProtectedRange::hash_value, true},
      {patch->salt_value_op, patch->salt_value, &ProtectedRange::salt_value, true},
      {patch->security_descriptor_op, patch->security_descriptor, &ProtectedRange::security_descriptor, true},
  };
  for (const StringPatch& f : strings) {
    if (!(f.clearable ? patch_op_ok(f.op) : patch_op_ok_without_clear(f.op))) {
      return invalid_argument("protected range patch op out of range", "op=" + std::to_string(f.op));
    }
    if (f.op == FM_PATCH_OP_SET) {
      if (f.value == nullptr) {
        return invalid_argument("SET string is NULL");
      }
      if (auto rc = check_utf8({f.value}); rc != 0) {
        return rc;
      }
    }
  }
  if (patch->name_op == FM_PATCH_OP_SET && patch->name[0] == '\0') {
    return invalid_argument("empty name");
  }
  if (!patch_op_ok_without_clear(patch->spin_count_op) || !patch_op_ok_without_clear(patch->ranges_op)) {
    return invalid_argument("protected range patch op out of range");
  }
  if (patch->ranges_op == FM_PATCH_OP_SET) {
    if (auto rc = check_rect_list(patch->ranges, patch->range_count); rc != 0) {
      return rc;
    }
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet_index).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  const std::size_t protected_range_index = FindRange(ranges, name);
  if (protected_range_index == ranges.size()) {
    return RangeNotFound(name);
  }
  ProtectedRange updated = ranges[protected_range_index];
  for (const StringPatch& f : strings) {
    if (f.op == FM_PATCH_OP_SET) {
      updated.*f.member = f.value;
    } else if (f.op == FM_PATCH_OP_CLEAR) {
      (updated.*f.member).clear();
    }
  }
  if (patch->spin_count_op == FM_PATCH_OP_SET) {
    updated.spin_count = patch->spin_count;
  }
  if (patch->ranges_op == FM_PATCH_OP_SET) {
    std::vector<formulon::MergeRange> rects;
    rects.reserve(patch->range_count);
    for (std::size_t i = 0; i < patch->range_count; ++i) {
      rects.push_back(formulon::c_api::parts::to_merge_range(patch->ranges[i]));
    }
    updated.sqref = formulon::io::format_sqref_text(rects);
  }
  if (patch->name_op == FM_PATCH_OP_SET) {
    for (std::size_t i = 0; i < ranges.size(); ++i) {
      if (i != protected_range_index && formulon::sheet_names::write_collides(ranges[i].name, updated.name)) {
        return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "name already in use",
                                 "name=" + updated.name);
      }
    }
  }
  ranges[protected_range_index] = std::move(updated);
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_protected_range(fm_workbook_t* wb, size_t sheet_index, const char* name) {
  clear_last_error();
  if (name == nullptr) {
    return null_argument("name");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  formulon::WorksheetRawExtensions& raw = wb->workbook().sheet(sheet_index).mutable_raw_extensions();
  auto ranges = formulon::io::read_protected_ranges(raw);
  const std::size_t protected_range_index = FindRange(ranges, name);
  if (protected_range_index == ranges.size()) {
    return RangeNotFound(name);
  }
  ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(protected_range_index));
  formulon::io::write_protected_ranges(raw, ranges);
  return 0;
}
