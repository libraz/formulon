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
// Field names match `fm_read_diagnostics_t` / `fm_save_diagnostics_t` in
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
  /// single increment can cost more than one rule).
  /// Fed by: `io.sheet.overlay.skip`, `io.cf.skip`.
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
  /// Modelled sheet features the writer could not lower to records at
  /// all. OOXML never contributes: it represents every modelled feature,
  /// and rejects the one case it cannot (a table naming a removed sheet)
  /// with `kIoWriteFailed` before any part is built, so the caller loses
  /// the save rather than the table.
  /// Fed by: `xlsb.writer.deferred`.
  std::uint32_t deferred_feature_count = 0;
  /// Passthrough parts dropped, either because their package path
  /// collided with a path the writer generates itself (the generated copy
  /// wins) or because two passthrough entries claimed the same path.
  /// Fed by: `ooxml_writer.passthrough_collision`,
  /// `xlsb.writer.passthrough_collision`.
  std::uint32_t dropped_part_count = 0;
  /// Round-tripped relationships not re-emitted because the part they
  /// target is no longer in the package. Emitting them anyway would leave
  /// a dangling relationship and Excel would open the package in repair
  /// mode, so dropping is the correct write.
  ///
  /// This is an *observation of* a loss, not an independent loss. When a
  /// part is dropped and a relationship pointed at it, both this counter
  /// and `dropped_part_count` rise by one for that single part: the two
  /// are separate observations of the same event and must not be summed.
  /// The counter is not suppressed in that case because a relationship
  /// can also dangle for a part that was never captured at all, and the
  /// caller cannot tell those apart from a suppressed count.
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
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PACKAGE_DIAGNOSTICS_H_
