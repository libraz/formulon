//
// Emission plan and package-XML builders for the XLSB writer: which passthrough
// parts survive collision detection, plus `[Content_Types].xml` and the
// workbook / worksheet relationship parts derived from that plan.

#include "io/xlsb/emission_plan.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <tuple>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "hyperlink_address.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/comments_records.h"
#include "io/xlsb/external_link_writer.h"
#include "io/xlsb/part_paths.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/slicer_records.h"
#include "io/xlsb/table_records.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

namespace {

// The reader accepts both `application/vnd.ms-excel.sheet.binary.macroEnabled.main`
// (used by `.xlsm` and the `.xlsb` corpus xlwings emits on macOS) and
// `application/vnd.ms-excel.sheet.macroEnabled.main` (the alternative
// some non-macro xlsb writers ship). We emit the first form because
// (a) it's what the reader's primary fixture uses and (b) it's the
// content type Excel for Mac actually writes — the other variant is
// accepted for compatibility on input only.
// Worksheet part content type. This is the worksheet data part, NOT the
// binary-index part (`application/vnd.ms-excel.binIndexWs`); mislabelling it as
// the index type makes Excel treat the package as having zero worksheets and
// reject it (-50).
constexpr std::string_view kCtWorksheetXlsb = "application/vnd.ms-excel.worksheet";
constexpr std::string_view kCtSharedStringsXlsb = "application/vnd.ms-excel.sharedStrings";
constexpr std::string_view kCtStylesXlsb = "application/vnd.ms-excel.styles";
constexpr std::string_view kCtSheetMetadataXlsb = "application/vnd.ms-excel.sheetMetadata";
constexpr std::string_view kCtExternalLinkXlsb = "application/vnd.ms-excel.externalLink";
constexpr std::string_view kCtCommentsXlsb = "application/vnd.ms-excel.comments";
constexpr std::string_view kCtVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing";
constexpr std::string_view kCtThreadedComments = "application/vnd.ms-excel.threadedcomments+xml";
constexpr std::string_view kCtPerson = "application/vnd.ms-excel.person+xml";

}  // namespace

/// Fails a retained part (pivot table, pivot cache, styles) whose model
/// state changed since load, including the model object having gone.
/// Untracked passthrough parts always pass.  A stale styles part is not
/// written through this check: the plan regenerates it from the model.
Expected<void, Error> CheckRetainedPartFreshness(const Workbook& workbook, const PassthroughPart& part) {
  if (part.retained_origin == PassthroughPart::RetainedOrigin::kNone) {
    return Expected<void, Error>::Ok();
  }
  const std::optional<std::uint64_t> current = current_retained_part_fingerprint(workbook, part);
  if (!current.has_value() || !part.model_fingerprint.has_value() || *current != *part.model_fingerprint) {
    return make_error(FormulonErrorCode::kIoXlsbRetainedPartStale,
                      "retained XLSB part no longer matches the model state it was loaded from",
                      std::string(kWriteXlsbContext) + " part=" + part.path);
  }
  return Expected<void, Error>::Ok();
}

namespace {

// True for a worksheet binary-index part (`xl/worksheets/binaryIndex<N>.bin`).
// These describe the original sheet bodies' row layout and are stale once we
// regenerate the sheets; Excel opens without them.
bool IsBinaryIndexPart(const std::string& path) {
  constexpr std::string_view kPrefix = "xl/worksheets/binaryIndex";
  return path.size() > kPrefix.size() && path.compare(0, kPrefix.size(), kPrefix) == 0;
}

// `xl/metadata.xml` is an XLSX-only dynamic-array metadata part.  Its XLSB
// counterpart is `xl/metadata.bin`; copying XML bytes into a binary workbook
// makes Excel repair the file on open.  Dynamic-array metadata that the
// model can represent is regenerated below; other XLSX-only metadata remains
// deliberately out of the XLSB package.
bool IsXlsxOnlyMetadataPart(const std::string& path) {
  return path == "xl/metadata.xml";
}

// An external link part or its rels. The writer generates every link from the
// model, so a source package's own parts are stale.
bool IsExternalLinkPart(const std::string& path) {
  constexpr std::string_view kPrefix = "xl/externalLinks/";
  return strings::starts_with(path, kPrefix);
}

}  // namespace

/// Package paths of the 1-based `position`-th worksheet part and its rels.
std::string sheet_part_path(std::size_t position) {
  return "xl/worksheets/sheet" + std::to_string(position) + ".bin";
}

std::string sheet_rels_part_path(std::size_t position) {
  return "xl/worksheets/_rels/sheet" + std::to_string(position) + ".bin.rels";
}

std::string ExternalLinkRelsPath(std::size_t position) {
  return "xl/externalLinks/_rels/externalLink" + std::to_string(position) + ".bin.rels";
}

namespace {

/// The retained source `xl/styles.bin`, or null when the workbook has none.
const PassthroughPart* FindRawStylesPart(const Workbook& wb) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == kStylesBinPath)
      return &part;
  }
  return nullptr;
}

bool HasModelledStyles(const Workbook& wb) {
  const StylesTable& styles = wb.styles();
  // An empty table is the unstyled shape a package without a styles part
  // parses into; both workbook factories seed one.  Any populated collection,
  // including an explicitly-created default XF, needs a styles part because
  // worksheet iStyleRef values resolve only through its relationship.
  return !styles.fonts.empty() || !styles.fills.empty() || !styles.borders.empty() || !styles.num_fmts.empty() ||
         !styles.cell_xfs.empty() || !styles.cell_style_xfs.empty() || !styles.cell_styles.empty() ||
         !styles.dxfs.empty();
}

std::unordered_set<std::string> BuildGeneratedPathSet(const Workbook& wb, bool emit_sst_part, bool emit_styles_part,
                                                      bool emit_metadata_part, bool emit_rich_value_parts,
                                                      std::size_t external_link_count,
                                                      const std::vector<io::EmissionPlan::CommentsPlan>& comments) {
  std::unordered_set<std::string> paths;
  for (const io::EmissionPlan::CommentsPlan& c : comments) {
    for (const std::string* path : {&c.comments_path, &c.vml_path, &c.threaded_path}) {
      if (!path->empty()) {
        paths.insert(*path);
      }
    }
  }
  if (!wb.persons().empty()) {
    paths.insert(std::string(kPersonsPartPath));
  }
  paths.insert("[Content_Types].xml");
  paths.insert("_rels/.rels");
  paths.insert(std::string(kWorkbookBinPath));
  paths.insert(std::string(kWorkbookBinRelsPath));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    paths.insert(sheet_part_path(i + 1U));
    paths.insert(sheet_rels_part_path(i + 1U));
  }
  if (emit_sst_part) {
    paths.insert(std::string(kSharedStringsBinPath));
  }
  if (emit_styles_part) {
    paths.insert(std::string(kStylesBinPath));
  }
  if (emit_metadata_part) {
    paths.insert(std::string(kMetadataBinPath));
  }
  if (emit_rich_value_parts) {
    paths.insert(std::string(kRdRichValuePartPath));
    paths.insert(std::string(kRdRichValueStructurePartPath));
    paths.insert(std::string(kRdRichValueTypesPartPath));
  }
  for (std::size_t i = 1; i <= external_link_count; ++i) {
    paths.insert(external_link_part_path(i));
    paths.insert(ExternalLinkRelsPath(i));
  }
  return paths;
}

}  // namespace

bool xlsb_styles_regenerated(const Workbook& wb) {
  const PassthroughPart* raw_styles = FindRawStylesPart(wb);
  const bool raw_styles_fresh = raw_styles != nullptr && static_cast<bool>(CheckRetainedPartFreshness(wb, *raw_styles));
  if (raw_styles_fresh) {
    // A note run in a font the stored table lacks needs the table written with that font appended.
    std::vector<const std::vector<CellComment>*> notes;
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      notes.push_back(&wb.sheet(i).comments());
    }
    return !NoteRunFonts(wb.styles().fonts, notes).appended().empty();
  }
  return raw_styles != nullptr || HasModelledStyles(wb);
}

bool xlsb_dxf_references_changed(const Workbook& wb) {
  const PassthroughPart* raw_styles = FindRawStylesPart(wb);
  return raw_styles != nullptr && raw_styles->dxf_reference_fingerprint != current_dxf_reference_fingerprint(wb);
}

Expected<EmissionPlan, Error> BuildEmissionPlan(const Workbook& wb, bool sst_present, bool generate_metadata,
                                                bool generate_rich_values, std::size_t external_link_count,
                                                std::vector<io::EmissionPlan::CommentsPlan> comments_by_sheet,
                                                WriteDiagnostics* diagnostics) {
  EmissionPlan plan;
  plan.comments_by_sheet = std::move(comments_by_sheet);
  plan.generated_persons = !wb.persons().empty();
  plan.has_text_cells = sst_present;
  // A retained styles.bin that still matches the model is emitted verbatim,
  // which keeps BrtXF fMergeCell/fSxButton and dxf properties the table does
  // not represent.  Without a raw part (XLSX/native source) or once the model
  // styles changed since load, styles.bin is generated from the model and the
  // stale raw part is dropped as a generated-path collision.
  plan.has_generated_styles = xlsb_styles_regenerated(wb);
  // Decided by `BuildMetadataPlan`, which the sheet bodies were already
  // emitted against.
  plan.has_generated_metadata = generate_metadata;
  plan.has_generated_rich_values = generate_rich_values;
  plan.external_link_count = external_link_count;

  const std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, sst_present, plan.has_generated_styles, plan.has_generated_metadata,
                            plan.has_generated_rich_values, external_link_count, plan.comments_by_sheet);

  // The reader rejects conflicting defaults, but Workbook is also a public
  // construction surface. Validate that hand-built workbooks cannot produce
  // ambiguous extension semantics on write.
  std::unordered_map<std::string, std::string> source_defaults;
  source_defaults.reserve(wb.default_content_types().size());
  for (const DefaultContentType& source : wb.default_content_types()) {
    if (source.extension.empty() || source.content_type.empty()) {
      continue;
    }
    const std::string extension = strings::to_ascii_lower(source.extension);
    auto [it, inserted] = source_defaults.emplace(extension, source.content_type);
    if (!inserted && it->second != source.content_type) {
      return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                        "workbook has conflicting Default content types for extension " + extension,
                        std::string(kWriteXlsbContext) + " extension=" + extension);
    }
  }

  std::unordered_set<std::string> kept_paths;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (!ooxml::is_safe_part_name(part.path)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "passthrough part name escapes package root; refusing to write",
                        std::string(kWriteXlsbContext) + " part=" + part.path);
    }
    if (generated.count(part.path) != 0U) {
      RecordDroppedPart("xlsb.writer.passthrough_collision", part.path, "generated_path_wins", diagnostics);
      continue;
    }
    // Drop parts that would be stale or orphaned after we regenerate the
    // worksheets. The binary-index parts (`xl/worksheets/binaryIndex*.bin`)
    // describe the ROW layout of the original sheet bodies; once we re-emit
    // sheets they no longer match, and Excel opens fine without them. The
    // calc chain likewise references the original formula graph; a stale
    // chain is rejected, and Excel rebuilds it on load when absent (same
    // policy as the OOXML writer). Neither carries a workbook relationship
    // here, so keeping them would leave dangling / mismatched parts. A
    // source external link part the generated set does not replace (an
    // `.xml` one, or a number past the model's links) would be an orphan.
    if (IsBinaryIndexPart(part.path) || part.path == "xl/calcChain.bin" || IsXlsxOnlyMetadataPart(part.path) ||
        is_xml_slicer_part(part.path) || IsExternalLinkPart(part.path)) {
      if (IsXlsxOnlyMetadataPart(part.path)) {
        StructuredLog("xlsb.writer.deferred")
            .field("kind", std::string_view("xlsx_metadata"))
            .field("count", static_cast<std::int64_t>(1))
            .warn();
      }
      continue;
    }
    if (!kept_paths.insert(part.path).second) {
      RecordDroppedPart("xlsb.writer.passthrough_collision", part.path, "duplicate_passthrough_path", diagnostics);
      continue;
    }
    if (part.content_type.empty()) {
      const std::string extension = ooxml::extension_of_part(part.path);
      // `rels` and `xml` are always registered, with the canonical type when the source has none.
      const bool always_registered = extension == "rels" || extension == "xml";
      if (extension.empty() || (!always_registered && source_defaults.count(extension) == 0U)) {
        return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                          "Default-typed passthrough part has no matching Default registration",
                          std::string(kWriteXlsbContext) + " part=" + part.path + " extension=" + extension);
      }
    }
    plan.passthrough_kept.push_back(&part);
  }
  for (const auto& [extension, content_type] : source_defaults) {
    plan.default_content_types.push_back(DefaultContentType{extension, content_type});
    if (extension == "bin" && content_type != kCtWorkbookXlsb) {
      plan.workbook_bin_override = true;
    }
  }
  sort_by_index(plan.default_content_types, [](const DefaultContentType& lhs, const DefaultContentType& rhs) {
    return lhs.extension < rhs.extension;
  });
  plan.passthrough_path_index.reset(plan.passthrough_kept);
  return plan;
}

// ---------------------------------------------------------------------------
// XML part builders: the package envelope is XML even in xlsb.
// ---------------------------------------------------------------------------

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 128 + plan.passthrough_kept.size() * 128);
  out.append(kXmlDecl);
  out.append("<Types");
  append_xml_ns_decl(out, "", kNsContentTypes);
  out.append(">\n");
  std::string_view rels_default = kCtPackageRels;
  std::string_view xml_default = kCtXml;
  // `bin` is commonly registered as the workbook type, but a real package
  // may use the same extension for OLE/VBA payloads. Preserve the source
  // Default when a kept binary passthrough needs it, and add a canonical
  // workbook Override so `xl/workbook.bin` remains unambiguous.
  std::string_view bin_default = kCtWorkbookXlsb;
  for (const DefaultContentType& def : plan.default_content_types) {
    if (def.extension == "rels") {
      rels_default = def.content_type;
    } else if (def.extension == "xml") {
      xml_default = def.content_type;
    } else if (def.extension == "bin") {
      bin_default = def.content_type;
    }
  }
  AppendDefault(out, "rels", rels_default);
  AppendDefault(out, "xml", xml_default);
  AppendDefault(out, "bin", bin_default);
  bool vml_default = false;
  for (const DefaultContentType& def : plan.default_content_types) {
    if (def.extension == "rels" || def.extension == "xml" || def.extension == "bin") {
      continue;
    }
    vml_default = vml_default || def.extension == "vml";
    AppendDefault(out, def.extension, def.content_type);
  }
  // A VML drawing is typed by extension, as in xlsx; the comments and thread parts by Override below.
  const bool any_comment_parts =
      std::any_of(plan.comments_by_sheet.begin(), plan.comments_by_sheet.end(),
                  [](const io::EmissionPlan::CommentsPlan& c) { return c.numeric_id != 0U; });
  if (any_comment_parts && !vml_default) {
    AppendDefault(out, "vml", kCtVmlDrawing);
  }
  if (plan.workbook_bin_override) {
    AppendOverride(out, kWorkbookBinPath, kCtWorkbookXlsb);
  }
  // A source `rels` Default may use a vendor-specific type. Generated
  // relationship parts still need their canonical OPC semantics, so pin
  // those generated paths with Overrides while leaving the source Default
  // available for retained opaque parts of the same extension. The special
  // `[Content_Types].xml` part intentionally remains Default-typed when the
  // source registry uses a noncanonical `xml` type; adding an Override for
  // that part would change the source registry's meaning.
  if (rels_default != kCtPackageRels) {
    AppendOverride(out, "_rels/.rels", kCtPackageRels);
    AppendOverride(out, kWorkbookBinRelsPath, kCtPackageRels);
    for (std::size_t i = 1; i <= plan.external_link_count; ++i) {
      AppendOverride(out, ExternalLinkRelsPath(i), kCtPackageRels);
    }
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      bool has_emitted_sheet_rels = (i < plan.sheet_tables.size() && !plan.sheet_tables[i].empty()) ||
                                    !plan.sheet_drawing_rids[i].empty() ||
                                    std::any_of(wb.sheet(i).hyperlinks().begin(), wb.sheet(i).hyperlinks().end(),
                                                [](const Hyperlink& hyperlink) { return !hyperlink.target.empty(); });
      for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
        if (plan.sheet_owns_drawing[i] && rel.type == kRelDrawing) {
          continue;
        }
        if (rel.target_external) {
          has_emitted_sheet_rels = true;
          break;
        }
        if (plan.passthrough_path_index.contains(plan.passthrough_kept, rel.target)) {
          has_emitted_sheet_rels = true;
          break;
        }
      }
      if (has_emitted_sheet_rels) {
        AppendOverride(out, sheet_rels_part_path(i + 1U), kCtPackageRels);
      }
    }
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    AppendOverride(out, sheet_part_path(i + 1U), kCtWorksheetXlsb);
  }
  if (plan.has_text_cells) {
    AppendOverride(out, kSharedStringsBinPath, kCtSharedStringsXlsb);
  }
  if (plan.has_generated_metadata) {
    AppendOverride(out, kMetadataBinPath, kCtSheetMetadataXlsb);
  }
  if (plan.has_generated_rich_values) {
    AppendOverride(out, kRdRichValuePartPath, kCtRdRichValue);
    AppendOverride(out, kRdRichValueStructurePartPath, kCtRdRichValueStructure);
    AppendOverride(out, kRdRichValueTypesPartPath, kCtRdRichValueTypes);
  }
  if (plan.has_generated_styles) {
    AppendOverride(out, kStylesBinPath, kCtStylesXlsb);
  }
  for (std::size_t i = 1; i <= plan.external_link_count; ++i) {
    AppendOverride(out, external_link_part_path(i), kCtExternalLinkXlsb);
  }
  for (const std::vector<XlsbTablePart>& tables : plan.sheet_tables) {
    for (const XlsbTablePart& table : tables) {
      AppendOverride(out, table.path, kCtTableXlsb);
    }
  }
  // Passthrough overrides: only for entries that carried an explicit
  // ContentType in the source archive. Default-typed parts (empty
  // content_type) must NOT appear as Overrides.
  for (const io::EmissionPlan::CommentsPlan& c : plan.comments_by_sheet) {
    if (!c.comments_path.empty()) {
      AppendOverride(out, c.comments_path, kCtCommentsXlsb);
    }
    if (!c.threaded_path.empty()) {
      AppendOverride(out, c.threaded_path, kCtThreadedComments);
    }
  }
  if (plan.generated_persons) {
    AppendOverride(out, kPersonsPartPath, kCtPerson);
  }
  AppendPassthroughOverrides(out, plan.passthrough_kept);
  out.append("</Types>\n");
  return out;
}

// Builds `xl/worksheets/_rels/sheet<N>.bin.rels`, or an empty string when the
// sheet has nothing to relate to.
//
// Retained relationship ids remain verbatim so opaque worksheet-tail records
// (drawings, tables, and similar features) continue to resolve. Model-owned
// BrtHLink records are emitted separately and get ids from the same collision
// set; a relationship whose target part is not in the package is dropped so
// the emitted rels never dangle. A sheet owning a model drawing writes only
// the drawing relationship named by `drawing_rid`, if any.
std::string BuildSheetRels(const Sheet& sheet, const EmissionPlan& plan, const std::vector<XlsbTablePart>& tables,
                           const std::vector<std::string>& hyperlink_rids, bool owns_drawing,
                           const std::string& drawing_rid, WriteDiagnostics* diagnostics,
                           const io::EmissionPlan::CommentsPlan* comments,
                           const EmissionPlan::CommentRelIds* comment_rids) {
  const bool owns_notes = comments != nullptr && comments->numeric_id != 0U;
  std::string entries;
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    if (owns_drawing && rel.type == kRelDrawing) {
      continue;
    }
    // Generated notes replace any retained notes drawing and parts.
    if (owns_notes && (rel.type == kRelComments || rel.type == kRelVmlDrawing || rel.type == kRelThreadedComment)) {
      continue;
    }
    if (!rel.target_external && !plan.passthrough_path_index.contains(plan.passthrough_kept, rel.target)) {
      // Counted alongside the package- and workbook-scope drops: a dropped
      // sheet relationship is a dropped relationship, and OOXML having no
      // sheet-scope equivalent must not make it invisible.
      RecordDroppedRelationship("xlsb.writer.sheet_rel_skipped", rel.type, rel.target, diagnostics);
      continue;
    }
    const std::string target = rel.target_external ? rel.target : TargetRelativeToWorksheet(rel.target);
    AppendRelationship(entries, rel.id, rel.type, target, rel.target_external, /*escape_target=*/true);
  }
  std::unordered_set<std::string> emitted_hyperlink_rids;
  emitted_hyperlink_rids.reserve(hyperlink_rids.size());
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& hyperlink = sheet.hyperlinks()[i];
    if (hyperlink.target.empty() || i >= hyperlink_rids.size() || hyperlink_rids[i].empty()) {
      continue;
    }
    if (!emitted_hyperlink_rids.insert(hyperlink_rids[i]).second) {
      continue;
    }
    AppendRelationship(entries, hyperlink_rids[i], kRelHyperlink,
                       EncodeTargetSpaces(saved_hyperlink(hyperlink, plan.book_directory).target),
                       /*target_external=*/true, /*escape_target=*/true);
  }
  for (const XlsbTablePart& table : tables) {
    AppendRelationship(entries, table.rel_id, kRelTable, TargetRelativeToWorksheet(table.path),
                       /*target_external=*/false, /*escape_target=*/true);
  }
  if (!drawing_rid.empty()) {
    AppendRelationship(entries, drawing_rid, kRelDrawing, TargetRelativeToWorksheet(sheet.drawing_rel_target()),
                       /*target_external=*/false, /*escape_target=*/true);
  }
  if (owns_notes && comment_rids != nullptr) {
    for (const auto& [rid, type, path] :
         {std::tuple{&comment_rids->comments, kRelComments, &comments->comments_path},
          std::tuple{&comment_rids->vml, kRelVmlDrawing, &comments->vml_path},
          std::tuple{&comment_rids->threaded, kRelThreadedComment, &comments->threaded_path}}) {
      if (!rid->empty() && !path->empty()) {
        AppendRelationship(entries, *rid, type, TargetRelativeToWorksheet(*path), /*target_external=*/false,
                           /*escape_target=*/true);
      }
    }
  }
  if (entries.empty()) {
    return {};
  }
  std::string out;
  out.reserve(entries.size() + 192);
  out.append(kXmlDecl);
  BeginRelationships(out);
  out.append(entries);
  EndRelationships(out);
  return out;
}

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst, const EmissionPlan& plan, const Workbook& wb,
                              const std::vector<std::string>& link_rel_ids, WriteDiagnostics* diagnostics,
                              std::vector<std::pair<std::string, std::string>>* id_map) {
  std::string out;
  out.reserve(256 + sheet_count * 192 + wb.unknown_workbook_rels().size() * 192);
  out.append(kXmlDecl);
  BeginRelationships(out);
  std::uint32_t next_rid = 1;
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const std::string target = "worksheets/sheet" + std::to_string(i + 1U) + ".bin";
    AppendRelationship(out, next_rid++, kRelWorksheet, target);
  }
  // The ids `BrtSupBookSrc` names each link by, numbered right after the sheets.
  for (std::size_t i = 0; i < link_rel_ids.size(); ++i) {
    AppendRelationship(out, link_rel_ids[i], kRelExternalLink,
                       TargetRelativeToWorkbook(external_link_part_path(i + 1U)),
                       /*target_external=*/false, /*escape_target=*/true);
  }
  next_rid += static_cast<std::uint32_t>(link_rel_ids.size());
  if (emit_sst) {
    AppendRelationship(out, next_rid++, kRelSharedStrings, "sharedStrings.bin", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  // The styles / theme / metadata parts ride the passthrough path, but the
  // reader (and Excel) locate them only through these workbook relationships.
  // Without them a styled cell's iStyleRef dangles against an empty style
  // table, and the theme / metadata parts are treated as orphans. Emit a rel
  // for each part that is actually present in the package.
  if (plan.has_generated_styles || plan.passthrough_path_index.contains(plan.passthrough_kept, kStylesBinPath)) {
    AppendRelationship(out, next_rid++, kRelStyles, "styles.bin", /*target_external=*/false, /*escape_target=*/true);
  }
  if (plan.passthrough_path_index.contains(plan.passthrough_kept, "xl/theme/theme1.xml")) {
    AppendRelationship(out, next_rid++, kRelTheme, "theme/theme1.xml", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  if (plan.has_generated_metadata || plan.passthrough_path_index.contains(plan.passthrough_kept, kMetadataBinPath)) {
    AppendRelationship(out, next_rid++, kRelSheetMetadata, "metadata.bin", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  if (plan.generated_persons) {
    AppendRelationship(out, next_rid++, kRelPerson, TargetRelativeToWorkbook(kPersonsPartPath));
  }
  if (plan.has_generated_rich_values) {
    for (const auto& [type, path] : {std::pair{kRelRdRichValue, kRdRichValuePartPath},
                                     std::pair{kRelRdRichValueStructure, kRdRichValueStructurePartPath},
                                     std::pair{kRelRdRichValueTypes, kRdRichValueTypesPartPath}}) {
      AppendRelationship(out, next_rid++, type, TargetRelativeToWorkbook(path));
    }
  }
  // Preserve relationships to raw XLSB parts that the reader does not model
  // (drawings, VBA, custom XML, etc.).  Internal targets are stored as
  // package paths on Workbook, while workbook relationships are relative to
  // `xl/`.  Do not duplicate relationships which the generated package
  // already owns above.
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    // The source's link relationships are replaced by the generated ones above.
    if (rel.type == kRelExternalLink) {
      continue;
    }
    if (!rel.target_external && !plan.passthrough_path_index.contains(plan.passthrough_kept, rel.target)) {
      RecordDroppedRelationship("xlsb.writer.workbook_rel_skipped", rel.type, rel.target, diagnostics);
      continue;
    }
    if (!rel.target_external && ((rel.type == kRelTheme && rel.target == "xl/theme/theme1.xml") ||
                                 (rel.type == kRelSheetMetadata && rel.target == kMetadataBinPath))) {
      continue;
    }
    const std::string target = rel.target_external ? rel.target : TargetRelativeToWorkbook(rel.target);
    if (id_map != nullptr) {
      id_map->emplace_back(rel.id, "rId" + std::to_string(next_rid));
    }
    AppendRelationship(out, next_rid++, rel.type, target, rel.target_external, /*escape_target=*/true);
  }
  EndRelationships(out);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
