#include "io/xlsb/frt_formula.h"

#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

std::uint32_t LoadU32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

Error Truncated() {
  return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "future-record formula truncated",
                    "context=xlsb_frt_formula");
}

}  // namespace

Expected<std::string, Error> read_frt_formula(ByteSpan& cursor, std::uint32_t base_row, std::uint32_t base_col,
                                              const FeatureFormulaReadContext& ctx) {
  if (cursor.size < 12U || LoadU32(cursor.data) != kFrtItemFlags) {
    return Truncated();
  }
  const std::size_t cce = LoadU32(cursor.data + 4);
  const std::size_t cb = LoadU32(cursor.data + 8);
  if (cursor.size - 12U < cce || cursor.size - 12U - cce < cb) {
    return Truncated();
  }
  // Re-frame as cce, rgce, cb, rgcb for the feature-formula decoder.
  std::vector<std::uint8_t> legacy;
  emit_u32(legacy, static_cast<std::uint32_t>(cce));
  legacy.insert(legacy.end(), cursor.data + 12, cursor.data + 12 + cce);
  emit_u32(legacy, static_cast<std::uint32_t>(cb));
  legacy.insert(legacy.end(), cursor.data + 12 + cce, cursor.data + 12 + cce + cb);
  cursor.data += 12U + cce + cb;
  cursor.size -= 12U + cce + cb;
  ByteSpan inner{legacy.data(), legacy.size()};
  return read_feature_formula(inner, base_row, base_col, ctx);
}

void emit_frt_formula(std::vector<std::uint8_t>& dst, const EncodedFeatureFormula& formula) {
  const std::vector<std::uint8_t>& b = formula.bytes;
  const std::size_t cce = b.size() >= 4U ? LoadU32(b.data()) : 0U;
  const std::size_t cb = b.size() >= 8U + cce ? LoadU32(b.data() + 4 + cce) : 0U;
  emit_u32(dst, kFrtItemFlags);
  emit_u32(dst, static_cast<std::uint32_t>(cce));
  emit_u32(dst, static_cast<std::uint32_t>(cb));
  dst.insert(dst.end(), b.begin() + 4, b.begin() + 4 + static_cast<std::ptrdiff_t>(cce));
  dst.insert(dst.end(), b.begin() + 8 + static_cast<std::ptrdiff_t>(cce), b.end());
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
