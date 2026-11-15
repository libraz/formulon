// The table-side pivot records with no xlsx counterpart in the base record set: SXFORMATS (`formats`), SXCRTFORMATS
// (`chartFormats`), the x14 field and data-field FRT records (`fillDownLabels`, `pivotShowAs`), the SXView14 block
// that carries `hideValuesRow`, and the drop of an unknown table `ext`. Each variant's `.xlsx` (Excel 16.115's save of
// an authored source) is written as `.xlsb` and its pivot table records held against Excel's own `.xlsb` save, masked
// by Excel's re-save of that `.xlsb` (tests/fixtures/excel/xlsb_parity/pivot).

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <set>
#include <string>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/read_options.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;
namespace ids = pivot_records;
using tc::Bytes;

/// The records this unit writes, plus the empty end records that pin where they sit in the part.
std::set<std::uint16_t> List1Types() {
  return {
      ids::kBeginSXFormats,   ids::kBeginSXFormat,  ids::kEndSXFormat,      ids::kEndSXFormats, ids::kBeginSXCrtFormats,
      ids::kBeginSXCrtFormat, ids::kEndSXCrtFormat, ids::kEndSXCrtFormats,  ids::kBeginPRule,   ids::kEndPRule,
      ids::kBeginPRFilters,   ids::kEndPRFilters,   ids::kBeginPRFilter,    ids::kEndPRFilter,  ids::kBeginPRFItem,
      ids::kEndPRFItem,       ids::kSXVD14,         ids::kSXDI14,           kFrtBegin,          kFrtEnd,
      ids::kBeginSXView14,    ids::kEndSXView14,    ids::kBeginSXView16,    ids::kEndSXView16,  ids::kEndSXVD,
      ids::kEndSXDI,          ids::kEndSXDIs,       ids::kTableStyleClient, ids::kEndSXView,
  };
}

bool IsPivotTablePart(const std::string& part) {
  return part.rfind("xl/pivotTables/", 0) == 0 && tc::IsRecordPart(part);
}

/// The selected record types of every pivot table part, one line per part, for a readable sequence diff.
std::string TypeSequences(const char* label, const Bytes& package, const std::set<std::uint16_t>& types) {
  std::string out;
  for (const auto& [part, records] : tc::ExtractRecords(package, types)) {
    if (!IsPivotTablePart(part)) {
      continue;
    }
    out += std::string("  ") + label + " " + part + ":";
    for (const tc::IndexedRecord& rec : records) {
      out += " " + std::to_string(rec.type);
    }
    out += "\n";
  }
  return out;
}

/// Writes `<variant>.xlsx` as `.xlsb` and holds its pivot table records of `types` against `<variant>.xlsb`.
void ExpectExcelsRecords(const std::string& variant, const std::set<std::uint16_t>& types) {
  const std::string root = FORMULON_FIXTURES_DIR;
  const std::string dir = root + "/excel/xlsb_parity/pivot/";
  const Bytes excel = tc::ReadBytes(dir + variant + ".xlsb");
  const Bytes resave = tc::ReadBytes(dir + variant + ".2.xlsb");
  ASSERT_FALSE(excel.empty()) << "missing fixture " << variant << ".xlsb";
  ASSERT_FALSE(resave.empty()) << "missing fixture " << variant << ".2.xlsb";

  // Mac Excel saved every twin, so the table records carry its host bits.
  ReadOptions mac;
  mac.profile = mac_365_en_us_profile();
  const tc::Converted ours = tc::ConvertToXlsb(root, "excel/xlsb_parity/pivot/" + variant + ".xlsx", mac);
  ASSERT_TRUE(ours.error.empty()) << ours.error;

  tc::ByteMask mask = tc::BuildTwinMask(excel, resave);
  tc::AddUninitialisedFields(excel, tc::DefaultUninitialisedHooks(), &mask);
  // Only the pivot table parts are compared: FRT and pivot-rule records also occur in other parts.
  for (const Bytes* package : {&excel, &ours.bytes}) {
    for (const std::string& part : tc::ListParts(*package)) {
      if (!IsPivotTablePart(part)) {
        mask.whole_parts.insert(part);
      }
    }
  }

  const std::vector<std::string> diffs = tc::CompareRecordSets(ours.bytes, excel, types, mask);
  EXPECT_TRUE(diffs.empty()) << [&] {
    std::string all;
    for (const std::string& d : diffs) {
      all += "  " + d + "\n";
    }
    return all + TypeSequences("ours ", ours.bytes, types) + TypeSequences("Excel", excel, types);
  }();
}

// Five `format` blocks over the pivot-area shapes: a selected item reference, a label-only reference, a row button,
// the grand row and the grand column.
TEST(XlsbParityPivotFormats, FormatsAreas) {
  ExpectExcelsRecords("formats_areas", List1Types());
}

// Two `format` blocks whose dxf ids Excel renumbers: `type="all"` and `collapsedLevelsAreSubtotals`.
TEST(XlsbParityPivotFormats, FormatsDxfOrder) {
  ExpectExcelsRecords("formats_dxf_order", List1Types());
}

// Two `chartFormat` blocks over data-field references; the root `chartFormat` attribute sits in BrtBeginSXView.
TEST(XlsbParityPivotFormats, ChartFormats) {
  std::set<std::uint16_t> types = List1Types();
  types.insert(ids::kBeginSXView);
  ExpectExcelsRecords("chart_formats", types);
}

// x14 `<pivotField fillDownLabels="1">` on the row field.
TEST(XlsbParityPivotFormats, X14FillDownLabels) {
  ExpectExcelsRecords("x14_fill_down", List1Types());
}

// x14 `<dataField pivotShowAs="percentOfParentRow">` on the first data field.
TEST(XlsbParityPivotFormats, X14PivotShowAs) {
  ExpectExcelsRecords("x14_show_as", List1Types());
}

// `hideValuesRow="0"` with two data fields and Values on columns: Excel keeps it by writing no SXView14 block.
TEST(XlsbParityPivotFormats, HideValuesRowOff) {
  ExpectExcelsRecords("hide_values_row", List1Types());
}

// A table `ext` with a URI Excel does not know: kept in Excel's xlsx save, written to no record in its xlsb.
TEST(XlsbParityPivotFormats, UnknownExtDropped) {
  ExpectExcelsRecords("unknown_ext", List1Types());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
