//
// Single source of truth for what a retained XLSB part's
// `PassthroughPart::model_fingerprint` means. The XLSB writer never
// regenerates a pivot table, pivot cache, or styles part from the model —
// it always retains the bytes read in verbatim (see `writer.cpp`'s
// passthrough step) — so this fingerprint is the only signal that tells an
// unmodified retained part apart from one whose model twin has since been
// mutated. The reader computes it once at load time; `write_xlsb`
// recomputes it from the current model immediately before emitting the
// part and fails closed on a mismatch. Both call through
// `current_retained_part_fingerprint` so neither end can drift from the
// other's notion of "what does this part depend on".
//
// Each fingerprint hashes the same canonical serialisation the OOXML
// writer already produces for that model type (`write_pivot_table_definition`,
// `write_pivot_cache_definition` / `write_pivot_cache_records`,
// `write_styles`): a model field this fingerprint would miss is also a
// field the OOXML round trip already fails to preserve.

#ifndef FORMULON_IO_XLSB_RETAINED_PART_FINGERPRINT_H_
#define FORMULON_IO_XLSB_RETAINED_PART_FINGERPRINT_H_

#include <cstdint>
#include <optional>

#include "passthrough_part.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The fingerprint `part` would need to carry for its retained bytes to
/// still match the piece of `wb`'s model named by `part.retained_origin`.
/// Returns `nullopt` when there is nothing to check (`kNone`) or the
/// referenced model object no longer exists (an out-of-range sheet/pivot
/// index, or a cache id `wb.find_pivot_cache` cannot resolve) — both cases
/// a caller should treat as "not fresh".
std::optional<std::uint64_t> current_retained_part_fingerprint(const Workbook& wb, const PassthroughPart& part);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RETAINED_PART_FINGERPRINT_H_
