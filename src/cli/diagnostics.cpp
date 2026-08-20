
#include "cli/diagnostics.h"

#include <cstdint>
#include <ostream>
#include <string_view>

#include "c_api/formulon_c.h"

namespace formulon {
namespace cli {

namespace {

// Appends `; name=value` for a counter that is actually non-zero. A zero
// counter is the normal case and printing it would bury the one number
// the user needs to act on.
void emit_counter(std::ostream& err, const char* name, std::uint32_t value) {
  if (value != 0U) {
    err << "; " << name << '=' << value;
  }
}

// The XLSB half of the load counters: lossy Ptg recovery and parts the
// reader could not decode. Zero for every `.xlsx` load.
void emit_xlsb_read_diagnostics(std::ostream& err, std::string_view subcommand, const fm_read_diagnostics_t& d) {
  if (d.undecoded_formula_count == 0U && d.undecoded_defined_name_count == 0U && d.undecoded_part_count == 0U) {
    return;
  }
  err << "formulon: " << subcommand << ": warning: XLSB read diagnostics";
  emit_counter(err, "undecoded_formula_count", d.undecoded_formula_count);
  emit_counter(err, "undecoded_defined_name_count", d.undecoded_defined_name_count);
  emit_counter(err, "undecoded_part_count", d.undecoded_part_count);
  err << '\n';
}

// The OOXML half of the same load counters: presentation-overlay entries
// dropped for an unusable reference, and an unrecognised workbook content
// type. Zero for every `.xlsb` load.
void emit_ooxml_read_diagnostics(std::ostream& err, std::string_view subcommand, const fm_read_diagnostics_t& d) {
  if (d.skipped_feature_count == 0U && d.unknown_content_type_count == 0U) {
    return;
  }
  err << "formulon: " << subcommand << ": warning: OOXML read diagnostics";
  emit_counter(err, "skipped_feature_count", d.skipped_feature_count);
  emit_counter(err, "unknown_content_type_count", d.unknown_content_type_count);
  err << '\n';
}

}  // namespace

fm_status_t emit_read_diagnostics(const fm_workbook_t* wb, std::ostream& err, std::string_view subcommand) {
  fm_read_diagnostics_t diagnostics{};
  if (auto rc = fm_workbook_read_diagnostics(wb, &diagnostics); rc != 0) {
    return rc;
  }
  emit_xlsb_read_diagnostics(err, subcommand, diagnostics);
  emit_ooxml_read_diagnostics(err, subcommand, diagnostics);
  return 0;
}

void emit_write_diagnostics(std::ostream& err, std::string_view subcommand, fm_workbook_format_t format,
                            const fm_save_diagnostics_t& d) {
  if (d.downgraded_formula_count == 0U && d.deferred_feature_count == 0U && d.dropped_part_count == 0U &&
      d.dropped_relationship_count == 0U && d.renumbered_part_count == 0U) {
    return;
  }
  err << "formulon: " << subcommand << ": warning: " << (format == FM_WORKBOOK_FORMAT_XLSB ? "XLSB" : "OOXML")
      << " write diagnostics";
  emit_counter(err, "downgraded_formula_count", d.downgraded_formula_count);
  emit_counter(err, "deferred_feature_count", d.deferred_feature_count);
  emit_counter(err, "dropped_part_count", d.dropped_part_count);
  emit_counter(err, "dropped_relationship_count", d.dropped_relationship_count);
  emit_counter(err, "renumbered_part_count", d.renumbered_part_count);
  // A dropped part also drops the relationship pointing at it, so the two
  // numbers above can describe one loss twice. Say so rather than leaving
  // the reader to add them up.
  if (d.dropped_part_count != 0U && d.dropped_relationship_count != 0U) {
    err << " (a dropped part also drops its relationship; these may describe the same loss)";
  }
  err << '\n';
}

}  // namespace cli
}  // namespace formulon
