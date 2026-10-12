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
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::value_from_fm;

namespace {

fm_status_t null_argument(const char* api) {
  return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                           (std::string(api) + ": NULL argument").c_str());
}

fm_status_t check_strength(fm_collation_strength_t strength, const char* api, formulon::eval::CollationStrength* out) {
  if (strength != FM_COLLATION_COMPARE && strength != FM_COLLATION_DYNAMIC && strength != FM_COLLATION_CLASSIC_EXACT) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": strength out of range").c_str());
  }
  *out = static_cast<formulon::eval::CollationStrength>(strength);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_text_compare(const char* a, const char* b, const char* profile_id,
                                       fm_collation_strength_t strength, int32_t* out_order) {
  static constexpr const char* kApi = "fm_text_compare";
  clear_last_error();
  if (a == nullptr || b == nullptr || profile_id == nullptr || out_order == nullptr) {
    return null_argument(kApi);
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, kApi, &profile); rc != 0) {
    return rc;
  }
  formulon::eval::CollationStrength fold{};
  if (auto rc = check_strength(strength, kApi, &fold); rc != 0) {
    return rc;
  }
  const int order = formulon::eval::collate(a, b, profile, fold);
  *out_order = order < 0 ? -1 : (order > 0 ? 1 : 0);
  return 0;
}

extern "C" fm_status_t fm_collation_classes(const fm_value_t* values, size_t count, const char* profile_id,
                                            fm_collation_strength_t strength, uint32_t* out_class_ids,
                                            uint32_t* out_class_count) {
  static constexpr const char* kApi = "fm_collation_classes";
  clear_last_error();
  if (profile_id == nullptr || out_class_count == nullptr ||
      (count > 0 && (values == nullptr || out_class_ids == nullptr))) {
    return null_argument(kApi);
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, kApi, &profile); rc != 0) {
    return rc;
  }
  formulon::eval::CollationStrength fold{};
  if (auto rc = check_strength(strength, kApi, &fold); rc != 0) {
    return rc;
  }
  if (count > std::numeric_limits<std::uint32_t>::max()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_collation_classes: count too large");
  }
  std::vector<formulon::Value> converted;
  converted.reserve(count);
  for (size_t i = 0; i < count; ++i) {
    formulon::Value v = formulon::Value::blank();
    if (auto rc = value_from_fm(values[i], &v, kApi); rc != 0) {
      return rc;
    }
    converted.push_back(std::move(v));
  }
  *out_class_count =
      formulon::eval::dynamic_array::collation_classes(converted.data(), count, profile, fold, out_class_ids);
  return 0;
}
