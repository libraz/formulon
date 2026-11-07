//
// Emission plan of the XLSB writer and the package-XML builders that read it.
// Internal to `src/io/xlsb/`; not part of the public API.

#ifndef FORMULON_IO_XLSB_EMISSION_PLAN_H_
#define FORMULON_IO_XLSB_EMISSION_PLAN_H_

#include <cstddef>
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
  bool has_generated_dynamic_metadata = false;
  /// External link parts generated from the model, in written order.
  std::size_t external_link_count = 0;
  /// Table parts generated from the model, per sheet.
  std::vector<std::vector<XlsbTablePart>> sheet_tables;
  /// Relationship id of each sheet's linked drawing; empty when not linked.
  std::vector<std::string> sheet_drawing_rids;
  /// Sheets whose model drawing replaces retained drawing relationships.
  std::vector<bool> sheet_owns_drawing;
};

/// Fails a retained part (pivot table, pivot cache, styles) whose model state
/// changed since load. Untracked passthrough parts always pass.
Expected<void, Error> CheckRetainedPartFreshness(const Workbook& workbook, const PassthroughPart& part);

/// Package paths of the 1-based `position`-th worksheet part, its rels, and
/// the rels of the `position`-th external link part.
std::string sheet_part_path(std::size_t position);
std::string sheet_rels_part_path(std::size_t position);
std::string ExternalLinkRelsPath(std::size_t position);

Expected<EmissionPlan, Error> BuildEmissionPlan(const Workbook& wb, bool sst_present, bool generate_dynamic_metadata,
                                                std::size_t external_link_count, WriteDiagnostics* diagnostics);

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan);

std::string BuildSheetRels(const Sheet& sheet, const EmissionPlan& plan, const std::vector<XlsbTablePart>& tables,
                           const std::vector<std::string>& hyperlink_rids, bool owns_drawing,
                           const std::string& drawing_rid, WriteDiagnostics* diagnostics);

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst, const EmissionPlan& plan, const Workbook& wb,
                              const std::vector<std::string>& link_rel_ids, WriteDiagnostics* diagnostics,
                              std::vector<std::pair<std::string, std::string>>* id_map);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_EMISSION_PLAN_H_
