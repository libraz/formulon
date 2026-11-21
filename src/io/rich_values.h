//
// Rich values an .xlsx or .xlsb package carries beyond the modelled `_error` ones. An
// in-cell picture in one of the two measured layouts becomes a `Value::picture`;
// any other (a linked data type, a picture in another layout) is read into an
// `OpaqueRichValue` kept on its cell: structure type, ordered keys and text, and each
// `_rvRel:` key resolved to the part it reaches through `richValueRel`. The save lists
// pictures (derived from the cached values) and opaque values with the rich errors in
// one table and regenerates every rich-value part.
//
// A read keeps `richValueRel.xml` and its rels as passthrough: the save
// supersedes them, and their targets name the media only rich values used,
// so one nothing references any more is dropped. Supporting-property bags
// and `richStyles` stay verbatim, never renumbered, so indices into them hold.
// Anything the table cannot carry keeps the whole book's parts verbatim, each
// cell's `vm` kept as `OpaqueRichValue::verbatim_vm` and saved as read, and the
// save reports why (`rich_value_fallback_reason`). A save in the other format
// converts the metadata part (`convert_verbatim_metadata`), or refuses while a
// cell still names a part it cannot convert.

#ifndef FORMULON_IO_RICH_VALUES_H_
#define FORMULON_IO_RICH_VALUES_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "io/rich_value_errors.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook_format.h"

namespace formulon {
class Workbook;
namespace io {
class ZipReader;

inline constexpr std::string_view kRichValueRelPartPath = "xl/richData/richValueRel.xml";
inline constexpr std::string_view kRichValueRelRelsPartPath = "xl/richData/_rels/richValueRel.xml.rels";
inline constexpr std::string_view kRelRichValueRel =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
inline constexpr std::string_view kCtRichValueRel = "application/vnd.ms-excel.richvaluerel+xml";
inline constexpr std::string_view kNsRichValueRel =
    "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";

/// The in-cell picture rich value: structure type, key names and the image relationship type.
inline constexpr std::string_view kPictureStructure = "_localImage";
inline constexpr std::string_view kImageKey = "_rvRel:LocalImageIdentifier";
inline constexpr std::string_view kOriginKey = "CalcOrigin";
inline constexpr std::string_view kTextKey = "Text";
inline constexpr std::string_view kImageRelType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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

/// True when `parts` keeps a metadata part (`xl/metadata.xml` or
/// `xl/metadata.bin`) or the rich values themselves verbatim, so a save
/// generates neither.
bool rich_value_parts_kept_verbatim(const std::vector<PassthroughPart>& parts);

/// `rich_value_parts_kept_verbatim` over `wb`'s passthrough parts.
bool rich_values_kept_verbatim(const Workbook& wb);

/// What a save does with a metadata part kept verbatim in the other format.
struct VerbatimMetadataConversion {
  /// The part in the target format, when the kept part converts.
  std::optional<std::vector<std::uint8_t>> bytes;
  /// The kept part is left out unconverted; no cell names it.
  bool dropped = false;
};

/// The metadata part `wb` keeps verbatim, converted for a save as `target`. It
/// converts only when the kept part parses to a model, the model reproduces the
/// part's bytes, and the target bytes parse back to the same model; XLDAPR is
/// appended when the book has a dynamic-array formula and the part no such
/// type. No kept part, or one already in `target`'s format, gives neither bytes
/// nor a drop. A part that does not convert is dropped when no cell names it,
/// else the save fails with `kIoFeatureNotWritable`.
Expected<VerbatimMetadataConversion, Error> convert_verbatim_metadata(const Workbook& wb, WorkbookFormat target);

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

/// The rich-value parts one consume reads, owning the bytes `package` points into.
struct RichValueSource {
  /// Metadata and every value, structure and type part to mark consumed.
  std::vector<std::string> parts;
  std::vector<std::vector<std::uint8_t>> bytes;
  RichValuePackage package;
};

/// Scans `xl/richData/` and loads the metadata part at `metadata_path` with the value, structure and
/// `richValueRel` parts into `*source`. `source` must not be moved afterwards.
void load_rich_value_source(const ZipReader& zip, std::string_view metadata_path,
                            const std::vector<UnknownRelationship>& rels, RichValueSource* source);

/// A read set: the errors as `restore_rich_error_cells` takes them (an opaque
/// value's slot there is unused) and, per rich value, its opaque value.
struct RichValueSet {
  RichErrorSet errors;
  std::vector<std::optional<OpaqueRichValue>> opaque;
  /// The package being read, set by its reader once the set is committed; a
  /// picture's media bytes are read from it.
  const ZipReader* zip = nullptr;
  /// Media parts read from `zip` while restoring pictures, by path; `take_package_part` hands them on.
  std::unordered_map<std::string, std::vector<std::uint8_t>> media;
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

/// As above for a package whose metadata part is not `xl/metadata.xml`
/// (`package.metadata` is unused): `metadata` is what that part says.
RichValueRead read_rich_value_parts(const RichValuePackage& package, RichValueMetadata metadata);

/// When `read` understood every part, moves its set into `*set`, marks `parts` consumed and drops their
/// workbook relationships from `rels`; returns false (touching nothing) otherwise.
bool commit_rich_value_read(RichValueRead& read, const std::vector<std::string>& parts, RichValueSet* set,
                            std::unordered_set<std::string>& consumed_parts, std::vector<UnknownRelationship>& rels);

/// What a reader's consume of the rich-value parts did.
struct RichValueConsumption {
  bool consumed = false;
  /// When not consumed: the metadata part's value-entry count, nullopt when
  /// it cannot be counted or there is no metadata part.
  std::optional<std::size_t> verbatim_value_entries;
};

/// Reads the cells of sheet `sheet_index` of `wb` that cache an error and
/// carry a `vm` naming an opaque value. An in-cell picture (`_localImage`
/// with an image relationship, an integer CalcOrigin in 0..255 and an optional
/// Text, no `style`) becomes a `Value::picture` interned in `wb`'s picture
/// table, its bytes read from `set.zip`; any other value is recorded on the sheet
/// as opaque. Returns the cells whose `vm` names anything else.
std::vector<RichErrorCell> restore_opaque_rich_values(Workbook& wb, std::size_t sheet_index,
                                                      const std::vector<RichErrorCell>& cells, RichValueSet& set);

/// For a book whose rich-value parts stay verbatim: records each cell's `vm`
/// as an `OpaqueRichValue` holding only `verbatim_vm`, when it names one of
/// the `value_entries` value-metadata entries, the cell caches #VALUE! and the
/// sheet is not kept opaque. Any other `vm` is dropped.
void restore_verbatim_vms(Workbook& wb, std::size_t sheet_index, const std::vector<RichErrorCell>& cells,
                          std::size_t value_entries);

/// The bytes of package part `path`: moved out of `set.media` when a picture read them, else read from `zip`,
/// so a media part is extracted and charged once.
Expected<std::vector<std::uint8_t>, Error> take_package_part(const ZipReader& zip, RichValueSet& set,
                                                             const std::string& path);

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
