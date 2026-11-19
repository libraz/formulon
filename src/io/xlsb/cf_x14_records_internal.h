#ifndef FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_
#define FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_

// Declarations shared by the x14 conditional-format record units
// (cf_x14_records.cpp, cf_x14_icon_set.cpp).

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
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
inline constexpr std::uint16_t kBeginColorScale14 = 1157;
inline constexpr std::uint16_t kEndColorScale14 = 1158;
inline constexpr std::size_t kRule14GuidOffset = 50;
// BrtCFVO14 type, indexed by `cf::CfvoType`.
inline constexpr std::array<std::uint32_t, 8> kCfvo14Types = {1, 4, 5, 2, 3, 7, 8, 9};

/// iType / iTemplate of a standalone colour scale, data bar and icon set rule.
inline constexpr std::uint32_t kRule14TypeColorScale = 3;
inline constexpr std::uint32_t kRule14TypeDataBar = 4;
inline constexpr std::uint32_t kRule14TypeIconSet = 6;
inline constexpr std::uint32_t kRule14TemplateColorScale = 2;
inline constexpr std::uint32_t kRule14TemplateDataBar = 3;
inline constexpr std::uint32_t kRule14TemplateIconSet = 4;

/// Writes the 70-byte BrtBeginCFRule14 of a standalone visual rule: no FRT formulas, dxfId none, every other
/// field zero, an empty (not null) strParam, then the GUID and fGuid = 1.
void EmitRule14Head(std::vector<std::uint8_t>& dst, std::uint32_t itype, std::uint32_t tmpl, std::int32_t priority,
                    const std::array<std::uint8_t, 16>& guid);

/// The GUID bytes of `id` for a standalone visual rule, or the refusal when the rule carries something the x14
/// records do not hold: stopIfTrue, raw extension XML, a non-positive priority or an id that is not a GUID.
Expected<std::array<std::uint8_t, 16>, Error> VisualRule14Guid(const cf::CFRule& rule, const std::string& id);

/// The priority and GUID of `head` read as a standalone visual BrtBeginCFRule14 of kind (`itype`, `tmpl`) in the
/// shape `EmitRule14Head` writes; false for any other head.
bool ParseVisualRule14Head(ByteSpan head, std::uint32_t itype, std::uint32_t tmpl, std::int32_t& priority,
                           std::string& id);

/// One BrtCFVO14 of an icon set (`icon`: fSaveGTE 1, fGTE the threshold's gte) or of a data bar or colour scale
/// (fSaveGTE 0, fGTE 0, which `gte = true` stands for).
Expected<void, Error> EmitCfvo14(std::vector<std::uint8_t>& dst, const cf::CfValueObject& v, bool icon,
                                 const MergeRange& base, const FeatureFormulaWriteContext& ctx);

/// The threshold `EmitCfvo14` writes as `p`, or nothing for any other shape.
std::optional<cf::CfValueObject> DecodeCfvo14(ByteSpan p, bool icon, const MergeRange& base,
                                              const FeatureFormulaReadContext& ctx);

std::optional<cf::CFRule> DecodeIconRule14(const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                                           const MergeRange& base, const FeatureFormulaReadContext& ctx);

Expected<void, Error> EmitIconRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const std::string& id,
                                     const MergeRange& base, const FeatureFormulaWriteContext& ctx);

/// What a BrtBeginCFRule14 payload names: its rule type, or nothing for a head that cannot be read or a kind outside
/// the supported set, and its GUID, empty unless the head ends with a GUID and fGuid = 1.
struct Rule14Kind {
  std::optional<cf::RuleType> type;
  std::string id;
};

Rule14Kind ReadRule14Kind(ByteSpan head);

/// Whether a BrtCFVO14 payload is a formula threshold.
bool IsFormulaCfvo14(ByteSpan payload);

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_X14_RECORDS_INTERNAL_H_
