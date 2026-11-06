//
// C ABI - text comparison and value classification under a profile's
// collation. Input values go through `value_from_fm`, so the validation rules
// match every other value-taking entry point.
//
// @size-budget: 3 KB

#include "eval/collation.h"

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/dynamic_array/common.h"
#include "excel_profile.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::value_from_fm;

namespace {

fm_status_t check_strength(std::int32_t strength, formulon::eval::CollationStrength* out) {
  if (strength != FM_COLLATION_COMPARE && strength != FM_COLLATION_DYNAMIC && strength != FM_COLLATION_CLASSIC_EXACT) {
    return invalid_argument("strength out of range");
  }
  *out = static_cast<formulon::eval::CollationStrength>(strength);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_compare_text(const char* a, const char* b, const char* profile_id, int32_t strength,
                                       int32_t* out_order) {
  clear_last_error();
  if (a == nullptr || b == nullptr || profile_id == nullptr || out_order == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({a, b}); rc != 0) {
    return rc;
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  formulon::eval::CollationStrength fold{};
  if (auto rc = check_strength(strength, &fold); rc != 0) {
    return rc;
  }
  const int order = formulon::eval::collate(a, b, profile, fold);
  *out_order = order < 0 ? -1 : (order > 0 ? 1 : 0);
  return 0;
}

extern "C" fm_status_t fm_get_collation_classes(const fm_value_t* values, size_t count, const char* profile_id,
                                                int32_t strength, uint32_t* out_class_ids, uint32_t* out_class_count) {
  clear_last_error();
  if (profile_id == nullptr || out_class_count == nullptr ||
      (count > 0 && (values == nullptr || out_class_ids == nullptr))) {
    return null_argument();
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  formulon::eval::CollationStrength fold{};
  if (auto rc = check_strength(strength, &fold); rc != 0) {
    return rc;
  }
  if (count > std::numeric_limits<std::uint32_t>::max()) {
    return invalid_argument("count too large");
  }
  std::vector<formulon::Value> converted;
  converted.reserve(count);
  for (size_t i = 0; i < count; ++i) {
    formulon::Value v = formulon::Value::blank();
    if (auto rc = value_from_fm(values[i], &v); rc != 0) {
      return rc;
    }
    converted.push_back(std::move(v));
  }
  *out_class_count =
      formulon::eval::dynamic_array::collation_classes(converted.data(), count, profile, fold, out_class_ids);
  return 0;
}
