//
// Loss / fallback counters accumulated while reading or writing a
// workbook package. Both container formats feed the same two structs, so
// a caller can ask "did this round-trip lose anything?" without first
// knowing whether it opened an `.xlsx` or an `.xlsb`.
//
// These live in their own header because the increment sites are spread
// across `io/` (`cf_reader`, `sheet_reader`, `ooxml/package_validator`,
// `ooxml/emission_plan`, `ooxml/workbook_xml_builder`, `xlsb/writer`) and
// every one of them would otherwise have to include the reader or writer
// entry point it ultimately feeds.
//
// Field names match `fm_read_diagnostics_info_t` / `fm_save_diagnostics_info_t` in
// `src/c_api/formulon_c.h` one-for-one; the C ABI layer copies straight
// across. Counters are `uint32_t` rather than `size_t` so the C ABI
// structs have the same layout on native and wasm32.
//
// Coverage is partial and deliberately so: these counters cover the
// part-, relationship- and feature-level losses of the two package
// writers and readers. Three of the library's structured-log emission
// sites stay log-only, because each needs its own semantics rather than
// a slot in an existing field: an unrepresentable cached cell value
// (`xlsb.writer.unsupported_cached_value`), unsupported worksheet format
// metadata (`xlsb.reader.unsupported_ws_format_metadata`), and a failed
// recalc worker launch (`recalc.worker.launch_failed`).
//
// `diagnostics` pointers are optional throughout: a reader or writer
// handed NULL discards the counters and behaves identically otherwise.
// Both halves follow that rule, so no caller of either can crash by
// declining to collect.

#ifndef FORMULON_IO_PACKAGE_DIAGNOSTICS_H_
#define FORMULON_IO_PACKAGE_DIAGNOSTICS_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {
namespace io {

/// Counters a package read accumulates.
///
/// Each field lists the structured-log events that feed it. A field whose
/// events cannot occur for the container actually loaded stays at zero;
/// the field's meaning does not change with the container.
struct ReadDiagnostics {
  /// Formula cells whose stored formula could not be decoded; the cached
  /// value survives but the formula does not.
  /// Fed by: `xlsb.formula.not_decoded`.
  std::uint32_t undecoded_formula_count = 0;
  /// Defined names skipped for the same reason.
  /// Fed by: `xlsb.defined_name.not_decoded`.
  std::uint32_t undecoded_defined_name_count = 0;
  /// Package parts whose content type could not be resolved, so they were
  /// not decoded and are absent from the loaded model. OOXML never
  /// contributes: it carries every unmodelled part through passthrough
  /// instead of dropping it.
  /// Fed by: `xlsb.package.parts_dropped`.
  std::uint32_t undecoded_part_count = 0;
  /// Presentation-overlay entries dropped because their reference was
  /// missing or unparseable. Counts one per dropped entry for merges,
  /// hyperlinks and data validations, and one per dropped
  /// `<conditionalFormatting>` block (a block carries several rules, so a
  /// single increment can cost more than one rule). Also counts one per
  /// `<f t="dataTable">` whose `ref` does not parse: the cell keeps its
  /// cached value but no data table is modelled. Also counts one per table
  /// column whose header dxf or header cell style was dropped (Excel refuses
  /// a file that carries either, in .xlsx and .xlsb alike), and one per .xlsb
  /// table part dropped, with its relationship, because it could not be read
  /// far enough to clear those slots.
  /// Also counts one per .xlsx sparkline group dropped because an item's
  /// location is not one range.
  /// Also counts one per defined name whose `xlm`, `function` or
  /// `vbProcedure` flag an .xlsx load drops because the book has no macro
  /// sheet, and one per .xlsb name whose lone `fFunc` flag the load drops.
  /// Fed by: `io.sheet.overlay.skip`, `io.cf.skip`, `io.sheet.formula.data_table_skip`,
  /// `io.table.column_header.skip`, `io.defined_name.macro_flag_dropped`,
  /// `io.sparkline.skip`, `xlsb.defined_name.function_flag_dropped`, `xlsb.table.unreadable_dropped`.
  std::uint32_t skipped_feature_count = 0;
  /// Workbook parts whose declared content type was unrecognised, so the
  /// read fell back to `WorkbookKind::kXlsx`. At most 1 per read.
  /// Fed by: `ooxml.reader.unknown_workbook_content_type`.
  std::uint32_t unknown_content_type_count = 0;
};

/// Counters one package write accumulates.
///
/// Same rule as `ReadDiagnostics`: a field means the same thing whichever
/// writer ran, and where an event exists in only one container the field
/// still counts it.
///
/// `dropped_part_count` and `dropped_relationship_count` deliberately
/// co-fire for one logical loss -- see the note on
/// `dropped_relationship_count`.
struct WriteDiagnostics {
  /// Formula cells emitted as their cached literal because the AST could
  /// not be lowered to the container's encoding. OOXML never contributes:
  /// it emits formula text verbatim.
  /// Fed by: `xlsb.writer.formula_downgraded`,
  /// `xlsb.writer.array_formula_downgraded`.
  std::uint32_t downgraded_formula_count = 0;
  /// Modelled features the writer could not carry, one per feature.
  /// XLSB: sheet features not lowered to records, workbook settings the
  /// records cannot hold, custom table styles, custom views (and sheet
  /// views without their workbook view), rich values kept verbatim.
  /// OOXML: rich values kept verbatim, an XLSB-source PivotTable the
  /// reader could not model, a custom view of an XLSB source with no
  /// OOXML form, an XLSB-source print record family the reader could not
  /// state exactly, a macro name (an .xlsx without a macro sheet cannot
  /// hold one). Rich values kept verbatim count once, and only when a rich
  /// error is saved plain or their metadata part is dropped: each cell
  /// otherwise keeps its vm. A table naming a removed sheet fails the save
  /// with `kIoWriteFailed` before any count is made.
  /// Fed by: `xlsb.writer.deferred`, `xlsb.writer.rich_values_kept_verbatim`,
  /// `ooxml_writer.rich_values_kept_verbatim`,
  /// `ooxml_writer.macro_name_left_out`, and the OOXML emission plan's
  /// PivotTable and custom-view counts.
  std::uint32_t deferred_feature_count = 0;
  /// Passthrough parts dropped: their package path collided with a path
  /// the writer generates itself (the generated copy wins), two passthrough
  /// entries claimed the same path, or the target format cannot hold the
  /// part (an xlsb binary part in an .xlsx, a metadata part kept verbatim
  /// that does not convert).
  /// Fed by: `ooxml_writer.passthrough_collision`,
  /// `ooxml_writer.passthrough_dropped`, `xlsb.writer.passthrough_collision`,
  /// `xlsb.writer.passthrough_dropped`.
  std::uint32_t dropped_part_count = 0;
  /// Round-tripped relationships left out with a part counted in
  /// `dropped_part_count`. A relationship takes the classification of its
  /// target: one whose part the save regenerates, leaves out as derived, or
  /// reports as a deferred feature is not counted here either.
  ///
  /// This is an *observation of* a loss, not an independent loss: both
  /// counters rise for one dropped part with a relationship to it, so they
  /// must not be summed. Settled once per save from the two lists below.
  /// Fed by: `ooxml_writer.package_rel_skipped`,
  /// `ooxml_writer.workbook_rel_skipped`,
  /// `xlsb.writer.package_rel_skipped`,
  /// `xlsb.writer.workbook_rel_skipped`,
  /// `xlsb.writer.sheet_rel_skipped`.
  std::uint32_t dropped_relationship_count = 0;
  /// Parts emitted under a writer-assigned id rather than the id the
  /// model carried, because the model's id was unusable. The part still
  /// ships; an external reference to the old id does not survive. XLSB
  /// never contributes: the binary writer does not reassign part ids.
  /// Fed by: `ooxml_writer.table_id_fallback`.
  std::uint32_t renumbered_part_count = 0;
  /// Paths counted in `dropped_part_count`, and the targets of the
  /// relationships the save left out; bookkeeping for
  /// `dropped_relationship_count`, not reported.
  std::vector<std::string> dropped_part_paths;
  std::vector<std::string> unwritten_relationship_targets;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PACKAGE_DIAGNOSTICS_H_
