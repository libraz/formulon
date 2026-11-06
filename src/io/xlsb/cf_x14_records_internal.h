#ifndef FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_
#define FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_

// Declarations shared by the x14 conditional-format record units
// (cf_x14_records.cpp, cf_x14_icon_set.cpp).

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "io/cf_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

inline constexpr std::uint16_t kEndCfRule14 = 1049;
inline constexpr std::size_t kRule14GuidOffset = 50;
// BrtCFVO14 type, indexed by `cf::CfvoType`; 7 (formula) is not written.
inline constexpr std::array<std::uint32_t, 8> kCfvo14Types = {1, 4, 5, 2, 3, 7, 8, 9};

std::optional<cf::CFRule> DecodeIconRule14(const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                                           const MergeRange& base, const FeatureFormulaReadContext& ctx);

Expected<void, Error> EmitIconRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const MergeRange& base,
                                     const FeatureFormulaWriteContext& ctx);

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_
