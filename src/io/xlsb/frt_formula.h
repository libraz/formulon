//
// Future-record formula framing shared by the x14 records the model owns
// (BrtCFVO14, BrtSparkline, BrtBeginSparklineGroup): [MS-XLSB] 2.5.59
// FRTFormula is a u32 flag word (reserved2 = 1, so 0x02), then an
// FRTParsedFormula laid out as `cce`, `cb`, `rgce`, `rgcb` -- the size
// fields first, unlike the `cce` / `rgce` / `cb` / `rgcb` order of a
// legacy feature formula.

#ifndef FORMULON_IO_XLSB_FRT_FORMULA_H_
#define FORMULON_IO_XLSB_FRT_FORMULA_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// FRTHeader flag bits ([MS-XLSB] 2.5.61).
inline constexpr std::uint32_t kFrtHeaderSqrefs = 0x02U;
inline constexpr std::uint32_t kFrtHeaderFormulas = 0x04U;
/// The FRTFormula / FRTSqref flag word Excel writes (fDoAdjust / reserved2).
inline constexpr std::uint32_t kFrtItemFlags = 0x02U;

/// Reads one FRTFormula from `cursor` and decodes it anchored at
/// (`base_row`, `base_col`), without a leading `=`.
Expected<std::string, Error> read_frt_formula(ByteSpan& cursor, std::uint32_t base_row, std::uint32_t base_col,
                                              const FeatureFormulaReadContext& ctx);

/// Appends `formula` (as `encode_feature_formula` returns it) as an FRTFormula.
void emit_frt_formula(std::vector<std::uint8_t>& dst, const EncodedFeatureFormula& formula);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_FRT_FORMULA_H_
