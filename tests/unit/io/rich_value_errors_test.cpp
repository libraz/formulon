// Rich-value errors in .xlsx against Excel's own files: the parts and cells a
// recalculated book writes, the errors and details a load restores without
// recalculating, a resave of that load, the blocked pivot anchor, and the
// measured (function, argument position) table of propagated #VALUE!.

#include "io/rich_value_errors.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <iostream>
#include <iterator>
#include <map>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "parser/ast.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

constexpr const char* kRichParts[] = {"xl/metadata.xml", "xl/richData/rdrichvalue.xml",
                                      "xl/richData/rdrichvaluestructure.xml"};
constexpr const char* kTypesPart = "xl/richData/rdRichValueTypes.xml";

std::vector<std::uint8_t> ReadSample(std::string_view name) {
  std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/rich_errors/" + std::string(name) + ".xlsx");
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

/// The text of `part`, or nullopt when the package has none.
std::optional<std::string> Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  std::string body;
  if (!test::extract_part(test::span_of(package), part, &body)) {
    return std::nullopt;
  }
  return body;
}

void ParseXml(const std::string& xml, pugi::xml_document* doc) {
  EXPECT_TRUE(doc->load_buffer(xml.data(), xml.size()));
}

/// The worksheet part of the sheet named `sheet`, resolved through the workbook relationships.
std::string SheetXml(const std::vector<std::uint8_t>& package, std::string_view sheet) {
  pugi::xml_document wb;
  pugi::xml_document rels;
  ParseXml(Part(package, "xl/workbook.xml").value_or(""), &wb);
  ParseXml(Part(package, "xl/_rels/workbook.xml.rels").value_or(""), &rels);
  std::string rid;
  for (pugi::xml_node s : wb.child("workbook").child("sheets").children("sheet")) {
    if (sheet == s.attribute("name").value()) {
      rid = s.attribute("r:id").value();
    }
  }
  for (pugi::xml_node r : rels.child("Relationships").children("Relationship")) {
    if (rid == r.attribute("Id").value()) {
      const std::string target = r.attribute("Target").value();
      return Part(package, target.front() == '/' ? target.substr(1) : "xl/" + target).value_or("");
    }
  }
  ADD_FAILURE() << "no sheet " << sheet;
  return {};
}

struct CellRef {
  std::string sheet;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string a1;
};

CellRef ParseRef(std::string_view ref) {
  CellRef out;
  const std::size_t bang = ref.find('!');
  out.sheet = std::string(ref.substr(0, bang));
  out.a1 = std::string(ref.substr(bang + 1U));
  EXPECT_TRUE(a1::parse_a1_ref(out.a1, &out.row, &out.col)) << ref;
  return out;
}

/// The attributes a rich error bears on its cell.
struct CellShape {
  bool found = false;
  std::string t;
  std::string cm;
  std::string vm;
  std::string v;
};

CellShape ShapeAt(const std::string& sheet_xml, std::string_view a1) {
  CellShape out;
  pugi::xml_document doc;
  ParseXml(sheet_xml, &doc);
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      if (a1 == c.attribute("r").value()) {
        out.found = true;
        out.t = c.attribute("t").value();
        out.cm = c.attribute("cm").value();
        out.vm = c.attribute("vm").value();
        out.v = c.child("v").text().get();
        return out;
      }
    }
  }
  return out;
}

/// Every `Sheet!A1` of `package` whose cell carries `vm`, sheets in tab order.
std::vector<std::string> VmCells(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  pugi::xml_document wb;
  ParseXml(Part(package, "xl/workbook.xml").value_or(""), &wb);
  for (pugi::xml_node s : wb.child("workbook").child("sheets").children("sheet")) {
    const std::string name = s.attribute("name").value();
    pugi::xml_document sheet;
    ParseXml(SheetXml(package, name), &sheet);
    for (pugi::xml_node row : sheet.child("worksheet").child("sheetData").children("row")) {
      for (pugi::xml_node c : row.children("c")) {
        if (c.attribute("vm")) {
          out.push_back(name + "!" + c.attribute("r").value());
        }
      }
    }
  }
  return out;
}

void ExpectSameCell(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual,
                    std::string_view ref) {
  SCOPED_TRACE(std::string(ref));
  const CellRef r = ParseRef(ref);
  const CellShape want = ShapeAt(SheetXml(expected, r.sheet), r.a1);
  const CellShape got = ShapeAt(SheetXml(actual, r.sheet), r.a1);
  ASSERT_TRUE(want.found);
  ASSERT_TRUE(got.found);
  EXPECT_EQ(got.t, want.t);
  EXPECT_EQ(got.cm, want.cm);
  EXPECT_EQ(got.vm, want.vm);
  EXPECT_EQ(got.v, want.v);
}

void ExpectSamePart(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual,
                    std::string_view part) {
  SCOPED_TRACE(std::string(part));
  const std::optional<std::string> want = Part(expected, part);
  const std::optional<std::string> got = Part(actual, part);
  ASSERT_EQ(got.has_value(), want.has_value());
  if (want.has_value()) {
    EXPECT_EQ(*got, *want);
  }
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

/// A new workbook holding `src`'s sheets, names, merges, tables, formulas and
/// non-error constants, entered through the Workbook API. Spilled cells and
/// the cells in `skip` are left out. The formulas in `staged` are entered
/// last, in that order, each followed by a recalculation, as Excel keeps the
/// spill entered first.
Workbook Rebuild(const Workbook& src, const std::vector<const char*>& skip, const std::vector<const char*>& staged) {
  Workbook wb = Workbook::create_empty();
  for (std::size_t s = 0; s < src.sheet_count(); ++s) {
    wb.add_sheet(src.sheet(s).name());
  }
  for (const DefinedName& name : src.defined_names()) {
    if (name.local_sheet_id < 0) {
      EXPECT_TRUE(static_cast<bool>(
          wb.set_defined_name(name.name, name.formula.front() == '=' ? name.formula : "=" + name.formula)));
    }
  }
  for (std::size_t s = 0; s < src.sheet_count(); ++s) {
    const Sheet& sheet = src.sheet(s);
    for (const MergeRange& merge : sheet.merges()) {
      EXPECT_TRUE(static_cast<bool>(wb.add_merge(s, merge)));
    }
    std::vector<std::uint32_t> rows;
    for (const auto& entry : sheet.rows()) {
      rows.push_back(entry.first);
    }
    std::sort(rows.begin(), rows.end());
    for (const std::uint32_t row : rows) {
      for (const RowCells::Entry e : sheet.rows().at(row)) {
        const std::string ref = sheet.name() + "!" + a1::encode_a1(row, e.col);
        bool skipped = sheet.spill_region_covering(row, e.col) != nullptr;
        for (const char* cell : skip) {
          skipped = skipped || ref == cell;
        }
        for (const char* cell : staged) {
          skipped = skipped || ref == cell;
        }
        if (skipped) {
          continue;
        }
        if (const auto block = src.array_formula_at(s, row, e.col)) {
          if (block->anchor_row == row && block->anchor_col == e.col) {
            EXPECT_TRUE(
                static_cast<bool>(wb.set_array_formula(s, row, e.col, block->rows, block->cols, e.cell.formula_text)))
                << ref;
          }
          continue;
        }
        const Value& v = e.cell.cached_value;
        if (!e.cell.formula_text.empty()) {
          EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(s, row, e.col, e.cell.formula_text))) << ref;
        } else if (v.is_text()) {
          EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(s, row, e.col, v.as_text()))) << ref;
        } else if (v.is_number() || v.is_boolean()) {
          EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(s, row, e.col, v))) << ref;
        }
      }
    }
  }
  wb.set_tables(src.tables());
  for (const char* cell : staged) {
    const CellRef r = ParseRef(cell);
    const Sheet* sheet = src.sheet_by_name(r.sheet);
    std::size_t index = 0;
    while (index < src.sheet_count() && &src.sheet(index) != sheet) {
      ++index;
    }
    const Cell* formula = sheet != nullptr ? sheet->cell_at(r.row, r.col) : nullptr;
    if (formula == nullptr || formula->formula_text.empty()) {
      ADD_FAILURE() << "no staged formula at " << cell;
      continue;
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(index, r.row, r.col, formula->formula_text))) << cell;
    EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry()))) << cell;
  }
  return wb;
}

// ---------------------------------------------------------------------------
// Writer: a book rebuilt and recalculated writes Excel's cells and parts.
// ---------------------------------------------------------------------------

struct WriterCase {
  const char* sample = nullptr;
  std::vector<const char*> cells;   // compared on t, cm, vm and <v>
  std::vector<const char*> skip;    // engine-divergent cells left out of the rebuild
  std::vector<const char*> staged;  // formulas entered last, one recalculation each, in Excel's entry order
  bool rebuild = true;              // false: the sample itself, loaded and recalculated
  bool compare_parts = true;        // false: a left-out cell carries a rich value in Excel's file
  bool compare_types = true;        // false: Excel's file has no types part
};

WriterCase Golden(const char* sample, std::vector<const char*> cells, std::vector<const char*> skip = {}) {
  WriterCase c;
  c.sample = sample;
  c.cells = std::move(cells);
  c.skip = std::move(skip);
  return c;
}

void ExpectWriterGolden(const WriterCase& c) {
  SCOPED_TRACE(c.sample);
  const std::vector<std::uint8_t> excel = ReadSample(c.sample);
  Workbook loaded = Load(excel);
  Workbook wb = c.rebuild ? Rebuild(loaded, c.skip, c.staged) : std::move(loaded);
  auto recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc)) << recalc.error().message;
  const std::vector<std::uint8_t> saved = Save(wb);
  ASSERT_FALSE(saved.empty());
  for (const char* cell : c.cells) {
    ExpectSameCell(excel, saved, cell);
  }
  if (c.compare_parts) {
    for (const char* part : kRichParts) {
      ExpectSamePart(excel, saved, part);
    }
  }
  if (c.compare_types) {
    ExpectSamePart(excel, saved, kTypesPart);
  }
}

TEST(RichValueErrorWriter, SValue) {
  ExpectWriterGolden(Golden("s_value", {"Sheet1!C1", "Sheet1!E1"}, {}));
}

TEST(RichValueErrorWriter, SMerge) {
  ExpectWriterGolden(Golden("s_merge", {"Sheet1!C1", "Sheet1!F1"}, {}));
}

TEST(RichValueErrorWriter, STable) {
  ExpectWriterGolden(Golden("s_table", {"Sheet1!C1", "Sheet1!F1"}, {}));
}

TEST(RichValueErrorWriter, SGrid) {
  ExpectWriterGolden(Golden("s_grid", {"Sheet1!E1", "Sheet1!C1048575"}, {}));
}

TEST(RichValueErrorWriter, CEmpty) {
  ExpectWriterGolden(Golden("c_empty", {"Sheet1!C1", "Sheet1!E1"}, {}));
}

TEST(RichValueErrorWriter, CNested) {
  ExpectWriterGolden(Golden("c_nested", {"Sheet1!C1", "Sheet1!E1"}, {}));
}

TEST(RichValueErrorWriter, CLambda) {
  ExpectWriterGolden(Golden("c_lambda", {"Sheet1!C1", "Sheet1!E1"}, {}));
}

TEST(RichValueErrorWriter, PlainValue) {
  ExpectWriterGolden(Golden("plain_value", {"Sheet1!C1", "Sheet1!E1", "Sheet1!C2"}, {}));
}

TEST(RichValueErrorWriter, Combo) {
  ExpectWriterGolden(Golden("combo",
                            {"Sheet1!C1", "Sheet1!D1", "Sheet1!E1", "Sheet1!F1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4",
                             "Sheet1!C10", "Sheet1!D10"},
                            {}));
}

TEST(RichValueErrorWriter, Multi) {
  ExpectWriterGolden(
      Golden("multi", {"Sheet1!C1", "Sheet1!E1", "Sheet1!C5", "Sheet2!C1", "Sheet2!D1", "Sheet2!E1"}, {}));
}

TEST(RichValueErrorWriter, Notypes) {
  WriterCase c = Golden("notypes", {"Sheet1!C1", "Sheet1!E1"}, {});
  c.compare_types = false;
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, SCse) {
  ExpectWriterGolden(Golden("s_cse", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, SDtable) {
  WriterCase c = Golden("s_dtable", {"Sheet1!F8"}, {});
  c.rebuild = false;
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, SPivot) {
  WriterCase c = Golden("s_pivot", {"Sheet2!G1"}, {});
  c.rebuild = false;
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, SPhantom) {
  // D3 was entered first and keeps its spill; E1 then meets D3's phantom E3.
  WriterCase c = Golden("s_phantom", {"Sheet1!E1"}, {});
  c.staged = {"Sheet1!D3", "Sheet1!E1"};
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, SMergeval) {
  ExpectWriterGolden(Golden("s_mergeval", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, SMixMergefirst) {
  ExpectWriterGolden(Golden("s_mix_mergefirst", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, SMixRowcol) {
  ExpectWriterGolden(Golden("s_mix_rowcol", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, SMixValfirst) {
  ExpectWriterGolden(Golden("s_mix_valfirst", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, CCalc) {
  // C6:C9 differ from Excel in value and are not compared, but they stay in the book: C6 blocks C5's spill.
  ExpectWriterGolden(
      Golden("c_calc", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!C4", "Sheet1!C5", "Sheet1!C10"}, {}));
}

TEST(RichValueErrorWriter, Propag) {
  // E1 `=INDEX(C1#,1)` is left out: the engine reads a spill reference to a blocked anchor as #REF!, Excel as
  // #SPILL!. Excel's E1 shares the propagated value of E2:E5, so the parts stay comparable.
  ExpectWriterGolden(
      Golden("propag", {"Sheet1!C1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!E5"}, {"Sheet1!E1"}));
}

TEST(RichValueErrorWriter, Numbering) {
  ExpectWriterGolden(Golden("numbering", {"Sheet2!C1", "Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, ECalc2) {
  ExpectWriterGolden(Golden("e_calc2", {"Sheet1!C2", "Sheet1!C3", "Sheet1!C4", "Sheet1!C6", "Sheet1!C8", "Sheet1!C15"},
                            {"Sheet1!C1"}));
}

TEST(RichValueErrorWriter, ESpill) {
  ExpectWriterGolden(
      Golden("e_spill", {"Sheet1!C1", "Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!J2"}, {"Sheet1!F1"}));
}

TEST(RichValueErrorWriter, EEdges) {
  ExpectWriterGolden(Golden("e_edges", {"Sheet1!XFC1", "Sheet1!C1048574", "Sheet1!E1048574"}, {}));
}

TEST(RichValueErrorWriter, EPass) {
  ExpectWriterGolden(Golden("e_pass", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C10"}, {}));
}

TEST(RichValueErrorWriter, PvpageBlocked) {
  WriterCase c = Golden("pvpage_blocked", {"Sheet2!A3"}, {});
  c.rebuild = false;
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, PivotBlocked) {
  WriterCase c = Golden("pivot_blocked", {"Sheet2!A3", "Sheet1!H8"}, {});
  c.rebuild = false;
  ExpectWriterGolden(c);
}

TEST(RichValueErrorWriter, B3Dep) {
  ExpectWriterGolden(Golden("b3_dep", {"Sheet1!C1", "Sheet1!E1", "Sheet1!E2"}, {}));
}

TEST(RichValueErrorWriter, B3Phantom) {
  ExpectWriterGolden(Golden("b3_phantom", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3"}, {}));
}

TEST(RichValueErrorWriter, B3Mixed) {
  ExpectWriterGolden(Golden("b3_mixed", {"Sheet1!E1", "Sheet1!E2", "Sheet1!C5"}, {}));
}

TEST(RichValueErrorWriter, B3Bycol) {
  ExpectWriterGolden(Golden("b3_bycol", {"Sheet1!C1"}, {}));
}

TEST(RichValueErrorWriter, B4Value) {
  ExpectWriterGolden(Golden("b4_value",
                            {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5", "Sheet1!E6",
                             "Sheet1!E7", "Sheet1!E8", "Sheet1!E9", "Sheet1!E10", "Sheet1!E11", "Sheet1!E12",
                             "Sheet1!E13", "Sheet1!E14", "Sheet1!E15", "Sheet1!E16"},
                            {}));
}

TEST(RichValueErrorWriter, B4Calc) {
  ExpectWriterGolden(Golden(
      "b4_calc", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5", "Sheet1!E6"}, {}));
}

TEST(RichValueErrorWriter, B4Spill) {
  ExpectWriterGolden(Golden(
      "b4_spill", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5", "Sheet1!E6"}, {}));
}

TEST(RichValueErrorWriter, VValue) {
  // IMAGE and TRANSLATE are stubs: C4's {2, 10} stays out, so the parts differ by that value.
  WriterCase c = Golden("v_value", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!C6", "Sheet1!C7", "Sheet1!C8"},
                        {"Sheet1!C4", "Sheet1!C5"});
  c.compare_parts = false;
  ExpectWriterGolden(c);
}

// ---------------------------------------------------------------------------
// Reader: a load restores each vm cell's real error and detail, no recalc.
// ---------------------------------------------------------------------------

using RvKeys = std::map<std::string, std::uint32_t>;

/// Excel's rich value behind `vm` (1-based; valueMetadata n names rv n-1 in every capture).
RvKeys ExcelRichValue(const std::vector<std::uint8_t>& package, std::uint32_t vm) {
  pugi::xml_document values;
  pugi::xml_document structures;
  ParseXml(Part(package, "xl/richData/rdrichvalue.xml").value_or(""), &values);
  ParseXml(Part(package, "xl/richData/rdrichvaluestructure.xml").value_or(""), &structures);
  std::vector<std::vector<std::string>> keys;
  for (pugi::xml_node s : structures.child("rvStructures").children("s")) {
    keys.emplace_back();
    for (pugi::xml_node k : s.children("k")) {
      keys.back().push_back(k.attribute("n").value());
    }
  }
  RvKeys out;
  std::uint32_t index = 1;
  for (pugi::xml_node rv : values.child("rvData").children("rv")) {
    if (index++ != vm) {
      continue;
    }
    const std::vector<std::string>& names = keys.at(rv.attribute("s").as_uint());
    std::size_t i = 0;
    for (pugi::xml_node v : rv.children("v")) {
      out[names.at(i++)] = v.text().as_uint();
    }
  }
  return out;
}

ErrorCode CodeOfType(std::uint32_t type) {
  return type == 8U ? ErrorCode::Spill : type == 13U ? ErrorCode::Calc : ErrorCode::Value;
}

const pivot::PivotTable* PivotAt(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  for (const std::unique_ptr<pivot::PivotTable>& table : sheet.pivot_tables()) {
    if (table != nullptr && table->anchor_row() == row && table->anchor_col() == col) {
      return table.get();
    }
  }
  return nullptr;
}

void ExpectRestoredCell(const Workbook& wb, const std::vector<std::uint8_t>& excel, std::string_view ref) {
  SCOPED_TRACE(std::string(ref));
  const CellRef r = ParseRef(ref);
  const Sheet* sheet = wb.sheet_by_name(r.sheet);
  ASSERT_NE(sheet, nullptr);
  const CellShape shape = ShapeAt(SheetXml(excel, r.sheet), r.a1);
  const RvKeys rv = ExcelRichValue(excel, static_cast<std::uint32_t>(std::strtoul(shape.vm.c_str(), nullptr, 10)));
  ASSERT_FALSE(rv.empty());
  const ErrorCode code = CodeOfType(rv.at("errorType"));
  const Cell* cell = sheet->cell_at(r.row, r.col);
  if (rv.count("colOffset") != 0U) {
    const std::uint32_t sub = rv.at("subType");
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->cached_value.is_error() && cell->cached_value.as_error() == ErrorCode::Spill);
    if (!cell->formula_text.empty()) {
      const std::optional<BlockedSpillFootprint> fp = sheet->blocked_spill_footprint_at(r.row, r.col);
      ASSERT_TRUE(fp.has_value());
      EXPECT_EQ(fp->rows, rv.at("rwOffset") + 1U);
      EXPECT_EQ(fp->cols, rv.at("colOffset") + 1U);
      EXPECT_EQ(fp->cause, sub == 3U   ? SpillCause::kBeyondEdge
                           : sub == 6U ? SpillCause::kMerged
                                       : SpillCause::kOccupied);
      return;
    }
    const pivot::PivotTable* table = PivotAt(*sheet, r.row, r.col);
    ASSERT_NE(table, nullptr);
    ASSERT_TRUE(table->blocked().has_value());
    EXPECT_EQ(table->blocked()->rows - table->blocked()->rows_above, rv.at("rwOffset") + 1U);
    EXPECT_EQ(table->blocked()->cols, rv.at("colOffset") + 1U);
    EXPECT_EQ(table->blocked()->past_edge, sub == 3U);
    return;
  }
  if (const SpillRegion* region = sheet->spill_region_covering(r.row, r.col)) {
    const std::size_t i =
        static_cast<std::size_t>(r.row - region->anchor_row) * region->cols + (r.col - region->anchor_col);
    ASSERT_LT(i, region->cells.size());
    EXPECT_TRUE(region->cells[i].is_error() && region->cells[i].as_error() == code);
    const std::optional<CellErrorDetail> anchor = sheet->error_detail_at(region->anchor_row, region->anchor_col);
    ASSERT_TRUE(anchor.has_value());
    EXPECT_TRUE(anchor->elements);
    EXPECT_EQ(anchor->code, code);
    return;
  }
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->cached_value.is_error() && cell->cached_value.as_error() == code);
  const std::optional<CellErrorDetail> detail = sheet->error_detail_at(r.row, r.col);
  ASSERT_TRUE(detail.has_value());
  EXPECT_EQ(detail->code, code);
  using Shape = CellErrorDetail::Shape;
  if (rv.count("ptg") != 0U) {
    EXPECT_EQ(detail->shape, Shape::kPropagatedPtg);
    EXPECT_EQ(detail->ptg, rv.at("ptg"));
  } else if (rv.count("propagated") != 0U && rv.count("iftab") != 0U) {
    EXPECT_EQ(detail->shape, Shape::kPropagatedIftab);
    EXPECT_EQ(detail->iftab, rv.at("iftab"));
  } else if (rv.count("propagated") != 0U) {
    EXPECT_EQ(detail->shape, Shape::kPropagated);
  } else if (rv.count("iftab") != 0U) {
    EXPECT_EQ(detail->shape, Shape::kIftab);
    EXPECT_EQ(detail->iftab, rv.at("iftab"));
    EXPECT_EQ(detail->sub_type, rv.at("subType"));
  } else {
    EXPECT_EQ(detail->shape, Shape::kSubType);
    EXPECT_EQ(detail->sub_type, rv.at("subType"));
  }
}

void ExpectReaderRestores(const char* sample) {
  SCOPED_TRACE(sample);
  const std::vector<std::uint8_t> excel = ReadSample(sample);
  const Workbook wb = Load(excel);
  const std::vector<std::string> cells = VmCells(excel);
  ASSERT_FALSE(cells.empty());
  for (const std::string& cell : cells) {
    ExpectRestoredCell(wb, excel, cell);
  }
  // The consumed parts are regenerated, not kept.
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    EXPECT_NE(part.path, "xl/metadata.xml");
    EXPECT_NE(part.path.rfind("xl/richData/", 0), 0U) << part.path;
  }
}

TEST(RichValueErrorReader, SValue) {
  ExpectReaderRestores("s_value");
}

TEST(RichValueErrorReader, SGrid) {
  ExpectReaderRestores("s_grid");
}

TEST(RichValueErrorReader, CLambda) {
  ExpectReaderRestores("c_lambda");
}

TEST(RichValueErrorReader, Combo) {
  ExpectReaderRestores("combo");
}

TEST(RichValueErrorReader, Multi) {
  ExpectReaderRestores("multi");
}

TEST(RichValueErrorReader, Notypes) {
  ExpectReaderRestores("notypes");
}

TEST(RichValueErrorReader, NotypesPivot) {
  ExpectReaderRestores("notypes_pivot");
}

TEST(RichValueErrorReader, CCalc) {
  ExpectReaderRestores("c_calc");
}

TEST(RichValueErrorReader, ESpill) {
  ExpectReaderRestores("e_spill");
}

TEST(RichValueErrorReader, EEdges) {
  ExpectReaderRestores("e_edges");
}

TEST(RichValueErrorReader, ECalc2) {
  ExpectReaderRestores("e_calc2");
}

TEST(RichValueErrorReader, PvpageBlocked) {
  ExpectReaderRestores("pvpage_blocked");
}

TEST(RichValueErrorReader, PivotBlocked) {
  ExpectReaderRestores("pivot_blocked");
}

TEST(RichValueErrorReader, B3Phantom) {
  ExpectReaderRestores("b3_phantom");
}

TEST(RichValueErrorReader, B4Value) {
  ExpectReaderRestores("b4_value");
}

TEST(RichValueErrorReader, B5Fn) {
  ExpectReaderRestores("b5_fn");
}

TEST(RichValueErrorReader, VValue) {
  // IMAGE's {2, 10} has no engine producer; it is kept as a loaded detail.
  ExpectReaderRestores("v_value");
}

TEST(RichValueErrorReader, ClassicErrorsStayPlain) {
  const Workbook wb = Load(ReadSample("plain_value"));
  const Sheet& sheet = wb.sheet(0);
  EXPECT_FALSE(sheet.has_error_details());
  EXPECT_EQ(sheet.cell_at(0U, 2U)->cached_value.as_error(), ErrorCode::Div0);
  EXPECT_EQ(sheet.cell_at(1U, 2U)->cached_value.as_error(), ErrorCode::Value);
}

TEST(RichValueErrorReader, AForeignRichValueKeepsEveryPartVerbatim) {
  // An image-in-cell structure is no `_error`: the set is foreign, so the
  // cells stay #VALUE! and the parts pass through.
  const std::string values =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"1\">"
      "<rv s=\"0\"><v>0</v><v>5</v></rv></rvData>";
  const std::string structures =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<rvStructures xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"1\">"
      "<s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/></s>"
      "</rvStructures>";
  const std::vector<std::optional<RichError>> parsed =
      parse_rich_errors(std::vector<std::uint8_t>(values.begin(), values.end()),
                        std::vector<std::uint8_t>(structures.begin(), structures.end()));
  ASSERT_EQ(parsed.size(), 1U);
  EXPECT_FALSE(parsed[0].has_value());
}

TEST(RichValueErrorReader, UnmeasuredValuesAreForeign) {
  const auto parse = [](const char* rv, const char* keys) {
    const std::string values = std::string("<rvData><rv s=\"0\">") + rv + "</rv></rvData>";
    const std::string structures = std::string("<rvStructures><s t=\"_error\">") + keys + "</s></rvStructures>";
    return parse_rich_errors(std::vector<std::uint8_t>(values.begin(), values.end()),
                             std::vector<std::uint8_t>(structures.begin(), structures.end()))
        .at(0);
  };
  const char* spill =
      "<k n=\"colOffset\" t=\"i\"/><k n=\"errorType\" t=\"i\"/><k n=\"rwOffset\" t=\"i\"/>"
      "<k n=\"subType\" t=\"i\"/>";
  EXPECT_TRUE(parse("<v>0</v><v>8</v><v>2</v><v>1</v>", spill).has_value());
  EXPECT_FALSE(parse("<v>0</v><v>8</v><v>2</v><v>2</v>", spill).has_value());
  EXPECT_FALSE(parse("<v>0</v><v>13</v><v>2</v><v>1</v>", spill).has_value());
  const char* sub = "<k n=\"errorType\" t=\"i\"/><k n=\"subType\" t=\"i\"/>";
  EXPECT_TRUE(parse("<v>13</v><v>1</v>", sub).has_value());
  EXPECT_TRUE(parse("<v>2</v><v>135</v>", sub).has_value());
  EXPECT_FALSE(parse("<v>13</v><v>5</v>", sub).has_value());
  EXPECT_FALSE(parse("<v>8</v><v>1</v>", sub).has_value());
  const char* ptg = "<k n=\"errorType\" t=\"i\"/><k n=\"propagated\" t=\"b\"/><k n=\"ptg\" t=\"i\"/>";
  EXPECT_TRUE(parse("<v>2</v><v>1</v><v>14</v>", ptg).has_value());
  EXPECT_FALSE(parse("<v>2</v><v>1</v><v>15</v>", ptg).has_value());
  EXPECT_FALSE(parse("<v>13</v><v>1</v><v>3</v>", ptg).has_value());
  EXPECT_FALSE(parse("<v>2</v><v>0</v><v>3</v>", ptg).has_value());
}

// ---------------------------------------------------------------------------
// Round trip: a load saved again without recalculation writes the same parts.
// ---------------------------------------------------------------------------

void ExpectRoundTrip(const char* sample, const char* expected_sample = nullptr) {
  SCOPED_TRACE(sample);
  const std::vector<std::uint8_t> excel = ReadSample(sample);
  const std::vector<std::uint8_t> expected = expected_sample != nullptr ? ReadSample(expected_sample) : excel;
  const std::vector<std::uint8_t> saved = Save(Load(excel));
  ASSERT_FALSE(saved.empty());
  for (const std::string& cell : VmCells(expected)) {
    ExpectSameCell(expected, saved, cell);
  }
  for (const char* part : kRichParts) {
    ExpectSamePart(expected, saved, part);
  }
  ExpectSamePart(expected, saved, kTypesPart);
}

TEST(RichValueErrorRoundTrip, SValue) {
  ExpectRoundTrip("s_value");
}

TEST(RichValueErrorRoundTrip, SGrid) {
  ExpectRoundTrip("s_grid");
}

TEST(RichValueErrorRoundTrip, CLambda) {
  ExpectRoundTrip("c_lambda");
}

TEST(RichValueErrorRoundTrip, Combo) {
  ExpectRoundTrip("combo");
}

TEST(RichValueErrorRoundTrip, Multi) {
  ExpectRoundTrip("multi");
}

TEST(RichValueErrorRoundTrip, Notypes) {
  ExpectRoundTrip("notypes", "notypes_resaved");
}

TEST(RichValueErrorRoundTrip, NotypesPivot) {
  ExpectRoundTrip("notypes_pivot", "notypes_pivot_resaved");
}

TEST(RichValueErrorRoundTrip, CCalc) {
  ExpectRoundTrip("c_calc");
}

TEST(RichValueErrorRoundTrip, ESpill) {
  ExpectRoundTrip("e_spill");
}

TEST(RichValueErrorRoundTrip, EEdges) {
  ExpectRoundTrip("e_edges");
}

TEST(RichValueErrorRoundTrip, PvpageBlocked) {
  ExpectRoundTrip("pvpage_blocked");
}

TEST(RichValueErrorRoundTrip, PivotBlocked) {
  ExpectRoundTrip("pivot_blocked");
}

TEST(RichValueErrorRoundTrip, B3Phantom) {
  ExpectRoundTrip("b3_phantom");
}

TEST(RichValueErrorRoundTrip, B4Value) {
  ExpectRoundTrip("b4_value");
}

TEST(RichValueErrorRoundTrip, B5Fn) {
  ExpectRoundTrip("b5_fn");
}

TEST(RichValueErrorRoundTrip, VValue) {
  ExpectRoundTrip("v_value");
}

// ---------------------------------------------------------------------------
// The measured (function, argument position) table.
// ---------------------------------------------------------------------------

CellErrorDetail ThroughFunction(const char* name, std::uint8_t argpos) {
  CellErrorDetail d;
  d.code = ErrorCode::Value;
  d.shape = CellErrorDetail::Shape::kPropagatedIftab;
  d.root_function = name;
  d.root_argpos = argpos;
  return d;
}

TEST(RichValueErrorIftab, AFunctionThatCoercesTheErrorCarriesItsId) {
  const std::optional<RichError> abs = rich_error_of_detail(ThroughFunction("ABS", 1U));
  ASSERT_TRUE(abs.has_value());
  EXPECT_EQ(abs->kind, RichError::Kind::kPropagatedIftab);
  EXPECT_EQ(abs->iftab, 24U);
  EXPECT_EQ(abs->error_type, 2U);
  // A newer function's id is the internal one Excel writes.
  const std::optional<RichError> sequence = rich_error_of_detail(ThroughFunction("SEQUENCE", 1U));
  ASSERT_TRUE(sequence.has_value());
  EXPECT_EQ(sequence->iftab, 632U);
  const std::optional<RichError> sort = rich_error_of_detail(ThroughFunction("SORT", 1U));
  ASSERT_TRUE(sort.has_value());
  EXPECT_EQ(sort->iftab, 615U);
}

TEST(RichValueErrorIftab, APositionIsMeasuredOnItsOwn) {
  // IF carries iftab at its condition only; NPV at its rate only.
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("IF", 1U))->kind, RichError::Kind::kPropagatedIftab);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("IF", 2U))->kind, RichError::Kind::kPropagated);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("NPV", 1U))->iftab, 11U);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("NPV", 2U))->kind, RichError::Kind::kPropagated);
  // Past the sweep, a variadic position folds into its last measured group.
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("NPV", 12U))->kind, RichError::Kind::kPropagated);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("SUM", 1U))->kind, RichError::Kind::kPropagated);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("NOSUCHFUNCTION", 1U))->kind, RichError::Kind::kPropagated);
  EXPECT_EQ(rich_error_of_detail(ThroughFunction("ABS", 0U))->kind, RichError::Kind::kPropagated);
}

TEST(RichValueErrorIftab, APlainPositionWritesNoRichValue) {
  EXPECT_FALSE(rich_error_of_detail(ThroughFunction("SUMIFS", 1U)).has_value());
  EXPECT_FALSE(rich_error_of_detail(ThroughFunction("MAXIFS", 1U)).has_value());
  EXPECT_TRUE(rich_error_of_detail(ThroughFunction("SUMIFS", 2U)).has_value());
}

TEST(RichValueErrorIftab, ARootOperatorWritesItsPtgId) {
  CellErrorDetail d;
  d.code = ErrorCode::Value;
  d.shape = CellErrorDetail::Shape::kPropagatedPtg;
  const std::pair<parser::BinOp, std::uint32_t> cases[] = {
      {parser::BinOp::Add, 3U}, {parser::BinOp::Sub, 4U},    {parser::BinOp::Mul, 5U}, {parser::BinOp::Div, 6U},
      {parser::BinOp::Pow, 7U}, {parser::BinOp::Concat, 8U}, {parser::BinOp::Lt, 9U},  {parser::BinOp::LtEq, 10U},
      {parser::BinOp::Eq, 11U}, {parser::BinOp::GtEq, 12U},  {parser::BinOp::Gt, 13U}, {parser::BinOp::NotEq, 14U},
  };
  for (const auto& [op, ptg] : cases) {
    d.root_binop = static_cast<std::uint8_t>(static_cast<std::uint8_t>(op) + 1U);
    const std::optional<RichError> v = rich_error_of_detail(d);
    ASSERT_TRUE(v.has_value());
    EXPECT_EQ(v->kind, RichError::Kind::kPropagatedPtg);
    EXPECT_EQ(v->ptg, ptg) << static_cast<int>(op);
  }
}

TEST(RichValueErrorIftab, TheTypesPartIsExcelsConstant) {
  EXPECT_EQ(rich_value_types_xml().size(), 1187U);
  EXPECT_EQ(std::string(rich_value_types_xml()), Part(ReadSample("combo"), kTypesPart).value_or(""));
}

// ---------------------------------------------------------------------------
// The sweep's shown values against the engine's, listed for review.
// ---------------------------------------------------------------------------

struct IftabSweepRow {
  const char* fn;
  std::uint32_t argpos;
  std::uint32_t row;
  const char* formula;
  const char* shown;
};

#include "rich_value_iftab_sweep_rows.inc"

/// True when the engine's `v` reads as Excel's displayed `shown`.
bool ShowsAs(const Value& v, std::string_view shown) {
  if (v.is_error()) {
    return display_name(v.as_error()) == shown;
  }
  if (v.is_boolean()) {
    return shown == (v.as_boolean() ? "TRUE" : "FALSE");
  }
  if (v.is_text()) {
    return v.as_text() == shown;
  }
  if (v.is_blank()) {
    return shown.empty() || shown == "0";
  }
  if (v.is_number()) {
    const std::string text(shown);
    char* end = nullptr;
    const double want = std::strtod(text.c_str(), &end);
    if (end == text.c_str() || *end != '\0') {
      return false;
    }
    return std::fabs(v.as_number() - want) <= 1e-6 * std::fmax(1.0, std::fabs(want));
  }
  return false;
}

// A review listing, not a gate: run with --gtest_also_run_disabled_tests.
TEST(RichValueErrorIftabSweep, DISABLED_ListsShownValueMismatches) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 51U, Value::number(r + 1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 52U, Value::number((r + 1.0) * 10.0))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4U, 2U, "=GROUPBY(AZ1:AZ3,BA1:BA3,1)")));
  std::size_t mismatches = 0;
  for (const IftabSweepRow& row : kIftabSweepRows) {
    const std::uint32_t r = row.row - 1U;
    if (!wb.set_cell_formula(0, r, 4U, row.formula) || !wb.recalc(eval::default_registry())) {
      std::cout << "iftab sweep not evaluated: " << row.formula << "\n";
      ++mismatches;
      continue;
    }
    const Value v = wb.sheet(0).resolve_cell_value(r, 4U);
    if (!ShowsAs(v, row.shown)) {
      ++mismatches;
      std::cout << "iftab sweep mismatch: " << row.fn << " pos " << row.argpos << " " << row.formula << "  Excel "
                << row.shown << "  engine "
                << (v.is_error()     ? std::string(display_name(v.as_error()))
                    : v.is_text()    ? std::string(v.as_text())
                    : v.is_number()  ? std::to_string(v.as_number())
                    : v.is_boolean() ? std::string(v.as_boolean() ? "TRUE" : "FALSE")
                                     : std::string("(blank)"))
                << "\n";
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 4U, Value::blank())));
  }
  std::cout << "iftab sweep mismatches: " << mismatches << " of " << std::size(kIftabSweepRows) << "\n";
}

}  // namespace
}  // namespace io
}  // namespace formulon
