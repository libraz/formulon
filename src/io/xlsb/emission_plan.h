//
// Emission plan of the XLSB writer and the package-XML builders that read it.
// Internal to `src/io/xlsb/`; not part of the public API.

#ifndef FORMULON_IO_XLSB_EMISSION_PLAN_H_
#define FORMULON_IO_XLSB_EMISSION_PLAN_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "io/ooxml/emission_plan.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/table_records.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;
class Sheet;

namespace io {
namespace xlsb {

struct EmissionPlan {
  std::vector<const PassthroughPart*> passthrough_kept;
  // Drawing parts written without the adjust guides their presets do not define
  // (`drawing_guides::without_undefined_guides`), listed in `passthrough_kept`
  // in place of the read ones; reserved before it is filled, so its entries
  // never move.
  std::vector<PassthroughPart> repaired_drawing_parts;
  // Shared membership index over the finalized kept list. The list and its
  // pointed-to parts remain immutable for the rest of the write.
  PassthroughPathIndex passthrough_path_index;
  /// Source Default registrations, lower-case by extension and unique. They
  /// remain available even when a source archive contains no current part of
  /// a given extension: keeping the registry avoids changing the meaning of
  /// any Default-typed part a caller may add before the next save.
  std::vector<DefaultContentType> default_content_types;
  /// When a source `bin` Default describes an embedded OLE/VBA payload (or
  /// any other non-workbook binary), the generated workbook must override
  /// that Default with the canonical XLSB workbook type.
  bool workbook_bin_override = false;
  bool has_text_cells = false;  // gates emission of xl/sharedStrings.bin
  bool has_generated_styles = false;
  /// The writer generates `xl/metadata.bin` (the XLDAPR entry dynamic-array
  /// anchors name, the XLRICHVALUE blocks rich errors name) and, with
  /// `has_generated_rich_values`, the `xl/richData/` parts.
  bool has_generated_metadata = false;
  bool has_generated_rich_values = false;
  /// The generated `xl/metadata.bin` is a verbatim `xl/metadata.xml` converted; that part is superseded, not lost.
  bool converted_metadata = false;
  /// External link parts generated from the model, in written order.
  std::size_t external_link_count = 0;
  /// Table parts generated from the model, per sheet.
  std::vector<std::vector<XlsbTablePart>> sheet_tables;
  /// A pivot cache the save writes, retained or generated from the model: its id, its two parts and their bytes.
  struct PivotCachePart {
    std::uint32_t cache_id = 0;
    std::string definition_path;
    std::string records_path;
    std::vector<std::uint8_t> definition;
    std::vector<std::uint8_t> records;
  };
  /// A pivot table the save writes, related from its sheet by `rel_id` and drawing from `pivot_caches[cache]`.
  struct PivotTablePart {
    std::string path;
    std::string rel_id;
    std::size_t cache = 0;
    /// The table's index among its sheet's pivot tables.
    std::size_t index = 0;
    std::vector<std::uint8_t> bytes;
    /// The bytes are the part as read, which may name dxfs the save does not renumber.
    bool verbatim = false;
  };
  /// Pivot caches in written order; the workbook relationship of the `k`-th is keyed in `BuildWorkbookRels`'s id
  /// map by `pivot_cache_rel_key(k)`.
  std::vector<PivotCachePart> pivot_caches;
  /// Pivot table parts the save writes, per sheet (`plan_pivot_parts`).
  std::vector<std::vector<PivotTablePart>> sheet_pivots;
  /// Relationship id of each sheet's linked drawing; empty when not linked.
  std::vector<std::string> sheet_drawing_rids;
  /// Sheets whose model drawing replaces retained drawing relationships.
  std::vector<bool> sheet_owns_drawing;
  /// Each sheet's comments, VML drawing and threaded-comment parts
  /// (`plan_comment_parts`); `numeric_id` is 0 for a sheet with none.
  std::vector<io::EmissionPlan::CommentsPlan> comments_by_sheet;
  /// The relationship ids naming each sheet's planned comment parts.
  struct CommentRelIds {
    std::string comments;
    std::string vml;
    std::string threaded;
  };
  std::vector<CommentRelIds> sheet_comment_rids;
  /// The writer generates the persons part the threads name.
  bool generated_persons = false;
  /// The canonical folder of `Workbook::location()`, or empty without one;
  /// local hyperlink targets are written relative to it (`saved_hyperlink`).
  /// Set by `write_xlsb` from the value its sheet bodies were written with.
  std::string book_directory;
};

/// The id `BuildWorkbookRels` maps to the written relationship of generated pivot cache `k`; no source id is
/// spelled so.
std::string pivot_cache_rel_key(std::size_t k);

/// Fails a retained part (pivot table, pivot cache, styles) whose model state
/// changed since load. Untracked passthrough parts always pass.
Expected<void, Error> CheckRetainedPartFreshness(const Workbook& workbook, const PassthroughPart& part);

/// True for a chartsheet, dialogsheet or macrosheet read from an .xlsb: its
/// binary part and relationships are written back as they were read, in
/// place of a worksheet part.
bool is_retained_binary_sheet(const Sheet& sheet);

/// The binary part an .xlsb save writes from a chartsheet read from an .xlsx
/// (its `.xml` path ending in `.bin`), or empty for any other sheet.
std::string converted_chartsheet_bin_path(const Sheet& sheet);

/// Package paths of the 1-based `position`-th worksheet part, its rels, and
/// the rels of the `position`-th external link part.
std::string sheet_part_path(std::size_t position);
std::string sheet_rels_part_path(std::size_t position);
std::string ExternalLinkRelsPath(std::size_t position);

/// True when the save writes styles.bin from the model rather than passing a
/// retained one through: there is no retained part, or it is stale since load.
/// A retained part that still matches the model passes through byte for byte
/// whatever the profile, and the parts written beside it follow its stored
/// names and indices (`stored_cell_style_save_plan`). Only a regenerated
/// styles.bin takes the save plans' names, faces and compacted cell formats.
bool xlsb_styles_regenerated(const Workbook& wb);

/// True when a retained styles.bin that passes through must have its dxf list
/// renumbered in place: the model's dxf references
/// (`current_dxf_reference_fingerprint`) differ from those at load. The rest
/// of the part stays byte for byte; an equal set leaves the part untouched.
bool xlsb_dxf_references_changed(const Workbook& wb);

/// `comments_by_sheet` is the comment-parts plan the sheet bodies were written against; the plan takes it.
Expected<EmissionPlan, Error> BuildEmissionPlan(const Workbook& wb, bool sst_present, bool generate_metadata,
                                                bool generate_rich_values, std::size_t external_link_count,
                                                std::vector<io::EmissionPlan::CommentsPlan> comments_by_sheet,
                                                WriteDiagnostics* diagnostics);

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan);

/// `comments` and `comment_rids` (null: none) relate the sheet's generated notes, VML drawing and threads.
std::string BuildSheetRels(const Sheet& sheet, const EmissionPlan& plan, const std::vector<XlsbTablePart>& tables,
                           const std::vector<std::string>& hyperlink_rids, bool owns_drawing,
                           const std::string& drawing_rid, WriteDiagnostics* diagnostics,
                           const io::EmissionPlan::CommentsPlan* comments = nullptr,
                           const EmissionPlan::CommentRelIds* comment_rids = nullptr,
                           const std::vector<EmissionPlan::PivotTablePart>* pivots = nullptr);

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst, const EmissionPlan& plan, const Workbook& wb,
                              const std::vector<std::string>& link_rel_ids, WriteDiagnostics* diagnostics,
                              std::vector<std::pair<std::string, std::string>>* id_map);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_EMISSION_PLAN_H_
