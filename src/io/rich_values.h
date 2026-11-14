//
// Rich values an .xlsx package carries beyond the modelled `_error` ones (an
// in-cell picture, a linked data type), read into `OpaqueRichValue`s kept on
// their cells: structure type, ordered keys and text, and each `_rvRel:` key
// resolved to the part it reaches through `richValueRel`. The save lists them
// with the rich errors in one table and regenerates every rich-value part.
//
// A read keeps `richValueRel.xml` and its rels as passthrough: the save
// supersedes them, and their targets name the media only rich values used,
// so one nothing references any more is dropped. Supporting-property bags
// and `richStyles` stay verbatim, never renumbered, so indices into them hold.
// Anything the table cannot carry keeps the whole book's parts verbatim, its
// `vm`s dropped, and the save reports why (`rich_value_fallback_reason`).

#ifndef FORMULON_IO_RICH_VALUES_H_
#define FORMULON_IO_RICH_VALUES_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "io/rich_value_errors.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"

namespace formulon {
class Workbook;
namespace io {

inline constexpr std::string_view kRichValueRelPartPath = "xl/richData/richValueRel.xml";
inline constexpr std::string_view kRichValueRelRelsPartPath = "xl/richData/_rels/richValueRel.xml.rels";
inline constexpr std::string_view kRelRichValueRel =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
inline constexpr std::string_view kCtRichValueRel = "application/vnd.ms-excel.richvaluerel+xml";
inline constexpr std::string_view kNsRichValueRel =
    "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";

/// Why a book keeps its rich-value parts verbatim.
inline constexpr std::string_view kRichFallbackMetadataMissing = "metadata_missing";
inline constexpr std::string_view kRichFallbackMetadataType = "metadata_type";    ///< not XLDAPR / XLRICHVALUE
inline constexpr std::string_view kRichFallbackMetadataBlock = "metadata_block";  ///< a block or entry not as measured
inline constexpr std::string_view kRichFallbackRvbAttribute = "rvb_attribute";    ///< an `rvb` carrying more than `i`
inline constexpr std::string_view kRichFallbackRichDataPart = "rich_data_part";   ///< an unknown `xl/richData/` part
inline constexpr std::string_view kRichFallbackValueParts = "value_parts";        ///< values or structures unreadable
inline constexpr std::string_view kRichFallbackValueShape = "value_shape";    ///< an element or attribute not carried
inline constexpr std::string_view kRichFallbackRvAttribute = "rv_attribute";  ///< an `rv` attribute other than `s`
inline constexpr std::string_view kRichFallbackKeyType = "key_type";          ///< a key type the table cannot carry
inline constexpr std::string_view kRichFallbackErrorValue = "error_value";    ///< an `_error` value not as measured
inline constexpr std::string_view kRichFallbackRelPart = "rel_part";  ///< a `_rvRel:` index that does not resolve
inline constexpr std::string_view kRichFallbackRelExternal = "rel_external";  ///< an external `_rvRel:` target
inline constexpr std::string_view kRichFallbackVmIndex = "vm_index";          ///< a `vm` naming no value
inline constexpr std::string_view kRichFallbackPackage = "package";  ///< an opaque sheet, or an external link with `vm`

/// True for a key holding a `richValueRel` index.
bool is_rich_value_rel_key(std::string_view name);

/// True when `parts` keeps `xl/metadata.xml` or the rich values themselves
/// verbatim, so a save generates neither.
bool rich_value_parts_kept_verbatim(const std::vector<PassthroughPart>& parts);

/// True for the `xl/richData/` parts a consumed book keeps as passthrough:
/// `richValueRel` with its rels, the supporting-property bags, `richStyles`.
bool is_retained_rich_data_part(std::string_view path);

/// The rich-value parts of one package; null where absent.
struct RichValuePackage {
  const std::vector<std::uint8_t>* metadata = nullptr;
  const std::vector<std::uint8_t>* values = nullptr;
  const std::vector<std::uint8_t>* structures = nullptr;
  const std::vector<std::uint8_t>* rel = nullptr;
  const std::vector<std::uint8_t>* rel_rels = nullptr;
  /// The first `xl/richData/` part neither read nor retained, or empty.
  std::string unknown_part;
  /// A workbook relationship of type `kRelRichValueRel` names `kRichValueRelPartPath`.
  bool rel_registered = false;
};

/// True when `rels` holds the internal workbook relationship to `richValueRel`.
bool rich_value_rel_registered(const std::vector<UnknownRelationship>& rels);

/// A read set: the errors as `restore_rich_error_cells` takes them (an opaque
/// value's slot there is unused) and, per rich value, its opaque value.
struct RichValueSet {
  RichErrorSet errors;
  std::vector<std::optional<OpaqueRichValue>> opaque;
};

struct RichValueRead {
  /// Empty when every part was understood; `set` is then the book's values.
  std::string_view fallback_reason;
  RichValueSet set;
};

/// Reads `package`. Values and `_error` values follow the strict rules of
/// `parse_rich_errors`; any other value is kept opaque when its layout is one
/// the table writes back unchanged.
RichValueRead read_rich_value_parts(const RichValuePackage& package);

/// Records on `sheet` the opaque value of each of `cells` that caches an error
/// and returns the cells whose `vm` names anything else.
std::vector<RichErrorCell> restore_opaque_rich_values(Sheet& sheet, const std::vector<RichErrorCell>& cells,
                                                      const RichValueSet& set);

/// Why `wb` keeps rich values verbatim, or empty when it keeps none.
std::string_view rich_value_fallback_reason(const Workbook& wb);

/// What a save does with `richValueRel`: the parts it writes for `table`
/// (none without `_rvRel:` keys) and the passthrough parts it leaves out, the
/// read `richValueRel` pair and each media part only those referenced.
struct RichValueRelPlan {
  std::vector<PassthroughPart> parts;
  std::unordered_set<std::string> skipped;
};

RichValueRelPlan plan_rich_value_rels(const Workbook& wb, const RichErrorTable& table);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_RICH_VALUES_H_
