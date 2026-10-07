#include "io/ooxml/workbook_part_reader.h"

#include <cstdint>
#include <string_view>

#include "calc_settings.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {

void apply_workbook_part_settings(const pugi::xml_node& wb_root, Workbook& wb) {
  // <calcPr> — workbook-level calc mode and iterative-calc options. The
  // element is optional; absence means Excel defaults (auto + iterative
  // off). When present, accept the documented `calcMode` values
  // (`auto` / `manual` / `autoNoTable`) and the iterative trio
  // (`iterate`, `iterateCount`, `iterateDelta`). Unknown calcMode
  // strings fall back to `auto` rather than failing the load.
  if (pugi::xml_node calc_pr = wb_root.child("calcPr"); calc_pr) {
    const std::string_view calc_mode_attr = calc_pr.attribute("calcMode").value();
    if (calc_mode_attr == "manual") {
      wb.set_calc_mode(Workbook::CalcMode::kManual);
    } else if (calc_mode_attr == "autoNoTable") {
      wb.set_calc_mode(Workbook::CalcMode::kAutoNoTable);
    } else {
      wb.set_calc_mode(Workbook::CalcMode::kAuto);
    }
    IterativeOptions opts;
    if (pugi::xml_attribute iterate = calc_pr.attribute("iterate"); iterate) {
      opts.enabled = parse_xml_bool(iterate.value());
    }
    if (pugi::xml_attribute count = calc_pr.attribute("iterateCount"); count) {
      const long long parsed = count.as_llong(static_cast<long long>(kDefaultMaxIterations));
      // Clamp into Excel's own dialog range. The upper bound matters more
      // than the lower one: the file decides how much work the first
      // `recalc()` performs, the solver has no wall-clock limit, and the
      // cancellation callback is null unless the host opted in. Without
      // this, `iterateCount="4294967295"` on a two-cell cycle is an
      // unrecoverable hang rather than a slow load.
      opts.max_iterations = parsed < 1                                           ? 1U
                            : parsed > static_cast<long long>(kMaxIterationsCap) ? kMaxIterationsCap
                                                                                 : static_cast<std::uint32_t>(parsed);
    }
    // The solver stops once the largest change falls below `max_change`.
    // A NaN tolerance makes that comparison false forever, so the workbook
    // silently burns the whole iteration budget and reports the
    // unconverged values; a negative one can never be reached either.
    double delta_value = 0.0;
    if (parse_xsd_nonneg_double(attr_str(calc_pr, "iterateDelta"), &delta_value)) {
      opts.max_change = delta_value;
    }
    wb.set_iterative_options(opts);
  }

  // Workbook-level elements captured raw for verbatim re-emission:
  // `<fileVersion>`, `<fileSharing>`, `<workbookPr>`, `<workbookProtection>`,
  // `<bookViews>`, and trailing `<extLst>`. Without this,
  // the writer regenerates only `<sheets>` / `<definedNames>` / `<calcPr>`
  // / `<pivotCaches>` and silently drops the date system, tab-selection
  // state, and workbook protection. `<workbookPr date1904>` additionally
  // seeds the model-level `date1904` flag the date-serial conversions read.
  if (pugi::xml_node file_version = wb_root.child("fileVersion"); file_version) {
    wb.set_file_version_xml(raw_xml(file_version));
  }
  if (pugi::xml_node file_sharing = wb_root.child("fileSharing"); file_sharing) {
    wb.set_file_sharing_xml(raw_xml(file_sharing));
  }
  if (pugi::xml_node workbook_pr = wb_root.child("workbookPr"); workbook_pr) {
    wb.set_workbook_pr_xml(raw_xml(workbook_pr));
    // Excel emits the attribute as `date1904`; some legacy producers use
    // the bare `1904` spelling. Both default to false when absent.
    const bool from_date1904 = read_xsd_bool(workbook_pr, "date1904", false);
    const bool from_legacy = read_xsd_bool(workbook_pr, "1904", false);
    wb.set_date1904(from_date1904 || from_legacy);
  }
  if (pugi::xml_node workbook_protection = wb_root.child("workbookProtection"); workbook_protection) {
    wb.set_workbook_protection_xml(raw_xml(workbook_protection));
  }
  if (pugi::xml_node book_views = wb_root.child("bookViews"); book_views) {
    wb.set_book_views_xml(raw_xml(book_views));
  }
  if (pugi::xml_node ext_lst = wb_root.child("extLst"); ext_lst) {
    wb.set_workbook_ext_lst_xml(raw_xml(ext_lst));
  }
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
