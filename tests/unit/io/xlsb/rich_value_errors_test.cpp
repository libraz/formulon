// Rich-value errors in .xlsb against Excel's own files: the metadata part,
// richData parts and per-cell value metadata a recalculated book writes, the
// errors and details a load restores without recalculating, a resave of that
// load, and the blocked pivot anchor of a retained pivot part.

#include "io/rich_value_errors.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
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
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "unknown_relationship.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kRichParts[] = {"xl/richData/rdrichvalue.xml", "xl/richData/rdrichvaluestructure.xml",
                                      "xl/richData/rdRichValueTypes.xml"};
constexpr const char* kMetadataPart = "xl/metadata.bin";

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

std::vector<std::uint8_t> ReadSample(std::string_view name, const char* dir = "rich_errors_xlsb",
                                     const char* extension = "xlsb") {
  std::vector<std::uint8_t> bytes = test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + dir + "/" +
                                                          std::string(name) + "." + extension);
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

/// The bytes of `part`, or nullopt when the package has none.
std::optional<std::vector<std::uint8_t>> Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  ZipReader zip;
  if (!zip.open(SpanOf(package)) || !zip.has_entry(part)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(part);
  if (!bytes) {
    return std::nullopt;
  }
  return std::move(bytes.value());
}

std::string TextOf(const std::vector<std::uint8_t>& package, std::string_view part) {
  const auto bytes = Part(package, part);
  return bytes.has_value() ? std::string(bytes->begin(), bytes->end()) : std::string();
}

using RecordList = std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>>;

RecordList RecordsOf(const std::vector<std::uint8_t>& part) {
  RecordList out;
  ByteSpan cursor = SpanOf(part);
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    out.emplace_back(record.value().type,
                     std::vector<std::uint8_t>(record.value().payload.data,
                                               record.value().payload.data + record.value().payload.size));
  }
  return out;
}

std::uint32_t U32(const std::vector<std::uint8_t>& payload, std::size_t at = 0) {
  return static_cast<std::uint32_t>(payload[at]) | (static_cast<std::uint32_t>(payload[at + 1U]) << 8U) |
         (static_cast<std::uint32_t>(payload[at + 2U]) << 16U) | (static_cast<std::uint32_t>(payload[at + 3U]) << 24U);
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

/// Each sheet's name and worksheet part, in tab order, read from the workbook
/// part and its relationships (Excel numbers the parts by creation, not tab, order).
std::vector<std::pair<std::string, std::string>> SheetParts(const std::vector<std::uint8_t>& package) {
  std::vector<std::pair<std::string, std::string>> out;
  const auto book = Part(package, "xl/workbook.bin");
  const std::string rels_text = TextOf(package, "xl/_rels/workbook.bin.rels");
  if (!book.has_value()) {
    return out;
  }
  pugi::xml_document rels;
  EXPECT_TRUE(rels.load_buffer(rels_text.data(), rels_text.size()));
  constexpr std::uint16_t kBundleSheet = 156;
  for (const auto& [type, payload] : RecordsOf(*book)) {
    if (type != kBundleSheet) {
      continue;
    }
    ByteSpan cursor{payload.data() + 8U, payload.size() - 8U};  // hsState, iTabID
    auto rel_id = read_xlwidestring(cursor);
    auto name = read_xlwidestring(cursor);
    if (!rel_id || !name) {
      ADD_FAILURE() << "unreadable BrtBundleSh";
      continue;
    }
    for (pugi::xml_node r : rels.child("Relationships").children("Relationship")) {
      if (rel_id.value() == r.attribute("Id").value()) {
        const std::string target = r.attribute("Target").value();
        out.emplace_back(name.value(), target.front() == '/' ? target.substr(1) : "xl/" + target);
      }
    }
  }
  return out;
}

std::string SheetPart(const std::vector<std::uint8_t>& package, std::string_view sheet) {
  for (const auto& [name, part] : SheetParts(package)) {
    if (name == sheet) {
      return part;
    }
  }
  ADD_FAILURE() << "no sheet " << sheet;
  return {};
}

/// What the records ahead of one cell say about it.
struct CellRecords {
  bool found = false;
  std::uint32_t cm = 0;
  std::uint32_t vm = 0;
  std::uint16_t type = 0;
  int error_byte = -1;
};

constexpr std::uint16_t kRowHeader = 0;
constexpr std::uint16_t kFirstCell = 1;
constexpr std::uint16_t kLastCell = 11;
constexpr std::uint16_t kCellMeta = 49;
constexpr std::uint16_t kValueMeta = 50;

CellRecords ScanCell(const std::vector<std::uint8_t>& sheet_bin, std::uint32_t row, std::uint32_t col) {
  CellRecords out;
  std::uint32_t current_row = 0;
  std::uint32_t cm = 0;
  std::uint32_t vm = 0;
  for (const auto& [type, payload] : RecordsOf(sheet_bin)) {
    if (type == kRowHeader) {
      current_row = U32(payload);
      cm = vm = 0;
    } else if (type == kCellMeta) {
      cm = U32(payload);
    } else if (type == kValueMeta) {
      vm = U32(payload);
    } else if (type >= kFirstCell && type <= kLastCell) {
      if (current_row == row && U32(payload) == col) {
        out.found = true;
        out.cm = cm;
        out.vm = vm;
        out.type = type;
        if (type == 3U || type == 11U) {
          out.error_byte = payload[8];
        }
        return out;
      }
      cm = vm = 0;
    }
  }
  return out;
}

/// Every `Sheet!A1` of `package` whose cell record follows a value-metadata record.
std::vector<std::string> VmCells(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  for (const auto& [sheet_name, part] : SheetParts(package)) {
    const auto bin = Part(package, part);
    if (!bin.has_value()) {
      continue;
    }
    std::uint32_t row = 0;
    std::uint32_t vm = 0;
    for (const auto& [type, payload] : RecordsOf(*bin)) {
      if (type == kRowHeader) {
        row = U32(payload);
        vm = 0;
      } else if (type == kValueMeta) {
        vm = U32(payload);
      } else if (type >= kFirstCell && type <= kLastCell) {
        if (vm != 0U) {
          out.push_back(sheet_name + "!" + a1::encode_a1(row, U32(payload)));
        }
        vm = 0;
      }
    }
  }
  return out;
}

void ExpectSameCell(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual,
                    std::string_view ref) {
  SCOPED_TRACE(std::string(ref));
  const CellRef r = ParseRef(ref);
  const auto want_bin = Part(expected, SheetPart(expected, r.sheet));
  const auto got_bin = Part(actual, SheetPart(actual, r.sheet));
  ASSERT_TRUE(want_bin.has_value());
  ASSERT_TRUE(got_bin.has_value());
  const CellRecords want = ScanCell(*want_bin, r.row, r.col);
  const CellRecords got = ScanCell(*got_bin, r.row, r.col);
  ASSERT_TRUE(want.found);
  ASSERT_TRUE(got.found);
  EXPECT_EQ(got.type, want.type);
  EXPECT_EQ(got.cm, want.cm);
  EXPECT_EQ(got.vm, want.vm);
  EXPECT_EQ(got.error_byte, want.error_byte);
}

void ExpectSamePart(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual,
                    std::string_view part) {
  SCOPED_TRACE(std::string(part));
  const auto want = Part(expected, part);
  const auto got = Part(actual, part);
  ASSERT_EQ(got.has_value(), want.has_value());
  if (want.has_value()) {
    EXPECT_EQ(*got, *want);
  }
}

/// The metadata part is compared record by record, so the framing of a
/// record's length does not matter.
void ExpectSameMetadata(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual) {
  const auto want = Part(expected, kMetadataPart);
  const auto got = Part(actual, kMetadataPart);
  ASSERT_TRUE(want.has_value());
  ASSERT_TRUE(got.has_value());
  EXPECT_EQ(RecordsOf(*got), RecordsOf(*want));
}

void ExpectSameRichParts(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual) {
  for (const char* part : kRichParts) {
    ExpectSamePart(expected, actual, part);
  }
  ExpectSameMetadata(expected, actual);
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto loaded = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = write_xlsb(wb);
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
// Writer: a book rebuilt and recalculated writes Excel's parts and records.
// ---------------------------------------------------------------------------

struct WriterCase {
  const char* sample = nullptr;
  std::vector<const char*> cells;   // compared on record kind, cm, vm and error byte
  std::vector<const char*> skip;    // engine-divergent cells left out of the rebuild
  std::vector<const char*> staged;  // formulas entered last, one recalculation each, in Excel's entry order
  bool rebuild = true;              // false: the sample itself, loaded
  bool recalc = true;               // false: saved as loaded, since a recalc re-plans a retained pivot part
  bool compare_parts = true;        // false: a left-out cell carries a rich value in Excel's file
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
  if (c.recalc) {
    auto recalc = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(recalc)) << recalc.error().message;
  }
  const std::vector<std::uint8_t> saved = Save(wb);
  ASSERT_FALSE(saved.empty());
  for (const char* cell : c.cells) {
    ExpectSameCell(excel, saved, cell);
  }
  if (c.compare_parts) {
    ExpectSameRichParts(excel, saved);
  }
}

TEST(XlsbRichValueErrorWriter, SValue) {
  ExpectWriterGolden(Golden("s_value", {"Sheet1!C1", "Sheet1!E1"}));
}

TEST(XlsbRichValueErrorWriter, SMerge) {
  ExpectWriterGolden(Golden("s_merge", {"Sheet1!C1", "Sheet1!F1"}));
}

TEST(XlsbRichValueErrorWriter, STable) {
  ExpectWriterGolden(Golden("s_table", {"Sheet1!C1", "Sheet1!F1"}));
}

TEST(XlsbRichValueErrorWriter, SGrid) {
  ExpectWriterGolden(Golden("s_grid", {"Sheet1!E1", "Sheet1!C1048575"}));
}

TEST(XlsbRichValueErrorWriter, CEmpty) {
  ExpectWriterGolden(Golden("c_empty", {"Sheet1!C1", "Sheet1!E1"}));
}

TEST(XlsbRichValueErrorWriter, CNested) {
  ExpectWriterGolden(Golden("c_nested", {"Sheet1!C1", "Sheet1!E1"}));
}

TEST(XlsbRichValueErrorWriter, CLambda) {
  ExpectWriterGolden(Golden("c_lambda", {"Sheet1!C1", "Sheet1!E1"}));
}

TEST(XlsbRichValueErrorWriter, PlainValue) {
  // No rich value anywhere: Excel writes no richData part, and neither does a save.
  const std::vector<std::uint8_t> excel = ReadSample("plain_value");
  Workbook wb = Rebuild(Load(excel), {}, {});
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const std::vector<std::uint8_t> saved = Save(wb);
  for (const char* cell : {"Sheet1!C1", "Sheet1!E1", "Sheet1!C2"}) {
    ExpectSameCell(excel, saved, cell);
  }
  for (const char* part : kRichParts) {
    EXPECT_FALSE(Part(saved, part).has_value()) << part;
  }
}

TEST(XlsbRichValueErrorWriter, Combo) {
  ExpectWriterGolden(Golden("combo", {"Sheet1!C1", "Sheet1!D1", "Sheet1!E1", "Sheet1!F1", "Sheet1!E2", "Sheet1!E3",
                                      "Sheet1!E4", "Sheet1!C10", "Sheet1!D10"}));
}

TEST(XlsbRichValueErrorWriter, Multi) {
  ExpectWriterGolden(Golden("multi", {"Sheet1!C1", "Sheet1!E1", "Sheet1!C5", "Sheet2!C1", "Sheet2!D1", "Sheet2!E1"}));
}

TEST(XlsbRichValueErrorWriter, SCse) {
  ExpectWriterGolden(Golden("s_cse", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, SDtable) {
  WriterCase c = Golden("s_dtable", {"Sheet1!F8"});
  c.rebuild = false;
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, SPivot) {
  WriterCase c = Golden("s_pivot", {"Sheet2!G1"});
  c.rebuild = false;
  c.recalc = false;
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, SPhantom) {
  // D3 was entered first and keeps its spill; E1 then meets D3's phantom E3.
  WriterCase c = Golden("s_phantom", {"Sheet1!E1"});
  c.staged = {"Sheet1!D3", "Sheet1!E1"};
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, SMergeval) {
  ExpectWriterGolden(Golden("s_mergeval", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, SMixMergefirst) {
  ExpectWriterGolden(Golden("s_mix_mergefirst", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, SMixRowcol) {
  ExpectWriterGolden(Golden("s_mix_rowcol", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, SMixValfirst) {
  ExpectWriterGolden(Golden("s_mix_valfirst", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, CCalc) {
  // C6:C9 differ from Excel in value and are not compared, but they stay in the book: C6 blocks C5's spill.
  ExpectWriterGolden(Golden("c_calc", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!C4", "Sheet1!C5", "Sheet1!C10"}));
}

TEST(XlsbRichValueErrorWriter, Propag) {
  // E1 `=INDEX(C1#,1)` is left out: the engine reads a spill reference to a blocked anchor as #REF!, Excel as
  // #SPILL!. Excel's E1 shares the propagated value of E2:E5, so the parts stay comparable.
  ExpectWriterGolden(
      Golden("propag", {"Sheet1!C1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!E5"}, {"Sheet1!E1"}));
}

TEST(XlsbRichValueErrorWriter, Numbering) {
  ExpectWriterGolden(Golden("numbering", {"Sheet2!C1", "Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, ECalc2) {
  ExpectWriterGolden(Golden("e_calc2", {"Sheet1!C2", "Sheet1!C3", "Sheet1!C4", "Sheet1!C6", "Sheet1!C8", "Sheet1!C15"},
                            {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, ESpill) {
  ExpectWriterGolden(
      Golden("e_spill", {"Sheet1!C1", "Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!J2"}, {"Sheet1!F1"}));
}

TEST(XlsbRichValueErrorWriter, EEdges) {
  ExpectWriterGolden(Golden("e_edges", {"Sheet1!XFC1", "Sheet1!C1048574", "Sheet1!E1048574"}));
}

TEST(XlsbRichValueErrorWriter, EPass) {
  ExpectWriterGolden(Golden("e_pass", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C10"}));
}

TEST(XlsbRichValueErrorWriter, PvpageBlocked) {
  // A recalculation re-plans a retained pivot part, so the book is saved as loaded.
  WriterCase c = Golden("pvpage_blocked", {"Sheet2!A3"});
  c.rebuild = false;
  c.recalc = false;
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, PivotBlocked) {
  WriterCase c = Golden("pivot_blocked", {"Sheet2!A3", "Sheet1!H8"});
  c.rebuild = false;
  c.recalc = false;
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, B3Dep) {
  ExpectWriterGolden(Golden("b3_dep", {"Sheet1!C1", "Sheet1!E1", "Sheet1!E2"}));
}

TEST(XlsbRichValueErrorWriter, B3Phantom) {
  ExpectWriterGolden(Golden("b3_phantom", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3"}));
}

TEST(XlsbRichValueErrorWriter, B3Mixed) {
  ExpectWriterGolden(Golden("b3_mixed", {"Sheet1!E1", "Sheet1!E2", "Sheet1!C5"}));
}

TEST(XlsbRichValueErrorWriter, B3Bycol) {
  ExpectWriterGolden(Golden("b3_bycol", {"Sheet1!C1"}));
}

TEST(XlsbRichValueErrorWriter, B4Value) {
  ExpectWriterGolden(Golden("b4_value", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5",
                                         "Sheet1!E6", "Sheet1!E7", "Sheet1!E8", "Sheet1!E9", "Sheet1!E10", "Sheet1!E11",
                                         "Sheet1!E12", "Sheet1!E13", "Sheet1!E14", "Sheet1!E15", "Sheet1!E16"}));
}

TEST(XlsbRichValueErrorWriter, B4Calc) {
  ExpectWriterGolden(
      Golden("b4_calc", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5", "Sheet1!E6"}));
}

TEST(XlsbRichValueErrorWriter, B4Spill) {
  ExpectWriterGolden(
      Golden("b4_spill", {"Sheet1!E1", "Sheet1!E2", "Sheet1!E3", "Sheet1!E4", "Sheet1!C5", "Sheet1!E5", "Sheet1!E6"}));
}

TEST(XlsbRichValueErrorWriter, VValue) {
  // IMAGE and TRANSLATE are stubs: C4's {2, 10} stays out, so the parts differ by that value.
  WriterCase c = Golden("v_value", {"Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!C6", "Sheet1!C7", "Sheet1!C8"},
                        {"Sheet1!C4", "Sheet1!C5"});
  c.compare_parts = false;
  ExpectWriterGolden(c);
}

TEST(XlsbRichValueErrorWriter, ABlockedSpillNamesBothMetadataEntries) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=SEQUENCE(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(9.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const std::vector<std::uint8_t> saved = Save(wb);

  const auto metadata = Part(saved, kMetadataPart);
  ASSERT_TRUE(metadata.has_value());
  EXPECT_EQ(*metadata, build_metadata_bin(true, 1U));
  const auto sheet = Part(saved, "xl/worksheets/sheet1.bin");
  ASSERT_TRUE(sheet.has_value());
  const CellRecords anchor = ScanCell(*sheet, 0U, 0U);
  ASSERT_TRUE(anchor.found);
  EXPECT_EQ(anchor.cm, 1U);
  EXPECT_EQ(anchor.vm, 1U);
  EXPECT_EQ(anchor.type, 11U);
  EXPECT_EQ(anchor.error_byte, 0x0F);
  EXPECT_EQ(ScanCell(*sheet, 1U, 0U).vm, 0U);

  const std::string rels = TextOf(saved, "xl/_rels/workbook.bin.rels");
  for (const char* target : {"richData/rdrichvalue.xml", "richData/rdrichvaluestructure.xml",
                             "richData/rdRichValueTypes.xml", "metadata.bin"}) {
    EXPECT_NE(rels.find(std::string("Target=\"") + target + "\""), std::string::npos) << target;
  }
  const std::string types = TextOf(saved, "[Content_Types].xml");
  for (const std::string_view type : {kCtRdRichValue, kCtRdRichValueStructure, kCtRdRichValueTypes}) {
    EXPECT_NE(types.find(std::string(type)), std::string::npos) << type;
  }
}

TEST(XlsbRichValueErrorWriter, ABlockedPivotWithoutItsPartIsWrittenPlain) {
  // A pivot loaded from .xlsx has no XLSB pivot part to write, so its anchor
  // is the plain #VALUE! and the rich values number the remaining cells alone.
  const std::vector<std::uint8_t> xlsx = ReadSample("pivot_blocked", "rich_errors", "xlsx");
  auto loaded = read_ooxml(test::span_of(xlsx));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  const Sheet* sheet2 = wb.sheet_by_name("Sheet2");
  ASSERT_NE(sheet2, nullptr);
  ASSERT_FALSE(sheet2->pivot_tables().empty());
  ASSERT_TRUE(sheet2->pivot_tables().front()->blocked().has_value());

  const std::vector<std::uint8_t> saved = Save(wb);
  ASSERT_FALSE(saved.empty());
  const auto bin = Part(saved, SheetPart(saved, "Sheet2"));
  ASSERT_TRUE(bin.has_value());
  const CellRecords anchor =
      ScanCell(*bin, sheet2->pivot_tables().front()->anchor_row(), sheet2->pivot_tables().front()->anchor_col());
  ASSERT_TRUE(anchor.found);
  EXPECT_EQ(anchor.vm, 0U);
  EXPECT_EQ(anchor.error_byte, 0x0F);
}

// ---------------------------------------------------------------------------
// Reader: a load restores each vm cell's real error and detail, no recalc.
// ---------------------------------------------------------------------------

using RvKeys = std::map<std::string, std::uint32_t>;

/// Excel's rich value behind `vm` (1-based; valueMetadata n names rv n-1 in every capture).
RvKeys ExcelRichValue(const std::vector<std::uint8_t>& package, std::uint32_t vm) {
  pugi::xml_document values;
  pugi::xml_document structures;
  const auto value_bytes = Part(package, kRichParts[0]);
  const auto structure_bytes = Part(package, kRichParts[1]);
  EXPECT_TRUE(value_bytes.has_value() && structure_bytes.has_value());
  if (!value_bytes.has_value() || !structure_bytes.has_value()) {
    return {};
  }
  EXPECT_TRUE(values.load_buffer(value_bytes->data(), value_bytes->size()));
  EXPECT_TRUE(structures.load_buffer(structure_bytes->data(), structure_bytes->size()));
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
  const auto bin = Part(excel, SheetPart(excel, r.sheet));
  ASSERT_TRUE(bin.has_value());
  const RvKeys rv = ExcelRichValue(excel, ScanCell(*bin, r.row, r.col).vm);
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
    EXPECT_NE(part.path, kMetadataPart);
    EXPECT_NE(part.path.rfind("xl/richData/", 0), 0U) << part.path;
  }
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    EXPECT_NE(rel.target, kMetadataPart);
    EXPECT_NE(rel.target.rfind("xl/richData/", 0), 0U) << rel.target;
  }
}

TEST(XlsbRichValueErrorReader, SValue) {
  ExpectReaderRestores("s_value");
}

TEST(XlsbRichValueErrorReader, SGrid) {
  ExpectReaderRestores("s_grid");
}

TEST(XlsbRichValueErrorReader, CLambda) {
  ExpectReaderRestores("c_lambda");
}

TEST(XlsbRichValueErrorReader, Combo) {
  ExpectReaderRestores("combo");
}

TEST(XlsbRichValueErrorReader, Multi) {
  ExpectReaderRestores("multi");
}

TEST(XlsbRichValueErrorReader, CCalc) {
  ExpectReaderRestores("c_calc");
}

TEST(XlsbRichValueErrorReader, ESpill) {
  ExpectReaderRestores("e_spill");
}

TEST(XlsbRichValueErrorReader, EEdges) {
  ExpectReaderRestores("e_edges");
}

TEST(XlsbRichValueErrorReader, ECalc2) {
  ExpectReaderRestores("e_calc2");
}

TEST(XlsbRichValueErrorReader, PvpageBlocked) {
  ExpectReaderRestores("pvpage_blocked");
}

TEST(XlsbRichValueErrorReader, PivotBlocked) {
  ExpectReaderRestores("pivot_blocked");
}

TEST(XlsbRichValueErrorReader, B3Phantom) {
  ExpectReaderRestores("b3_phantom");
}

TEST(XlsbRichValueErrorReader, B4Value) {
  ExpectReaderRestores("b4_value");
}

TEST(XlsbRichValueErrorReader, B5Fn) {
  ExpectReaderRestores("b5_fn");
}

TEST(XlsbRichValueErrorReader, VValue) {
  // IMAGE's {2, 10} has no engine producer; it is kept as a loaded detail.
  ExpectReaderRestores("v_value");
}

TEST(XlsbRichValueErrorReader, ClassicErrorsStayPlain) {
  const Workbook wb = Load(ReadSample("plain_value"));
  const Sheet& sheet = wb.sheet(0);
  EXPECT_FALSE(sheet.has_error_details());
  EXPECT_EQ(sheet.cell_at(0U, 2U)->cached_value.as_error(), ErrorCode::Div0);
  EXPECT_EQ(sheet.cell_at(1U, 2U)->cached_value.as_error(), ErrorCode::Value);
}

TEST(XlsbRichValueErrorReader, ABlockedPivotWithItsRetainedPartLoadsBlocked) {
  // The pivot part is the XLSB one, kept verbatim: the anchor's rich value is
  // what marks the report blocked, and the part is written back as it came.
  const std::vector<std::uint8_t> excel = ReadSample("pivot_blocked");
  const Workbook wb = Load(excel);
  const Sheet* sheet = wb.sheet_by_name("Sheet2");
  ASSERT_NE(sheet, nullptr);
  ASSERT_EQ(sheet->pivot_tables().size(), 1U);
  EXPECT_TRUE(sheet->pivot_tables().front()->blocked().has_value());
  bool part_kept = false;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    part_kept = part_kept || part.retained_origin == PassthroughPart::RetainedOrigin::kPivotTable;
  }
  EXPECT_TRUE(part_kept);

  const std::vector<std::uint8_t> saved = Save(wb);
  const auto want = Part(excel, "xl/pivotTables/pivotTable1.bin");
  const auto got = Part(saved, "xl/pivotTables/pivotTable1.bin");
  ASSERT_TRUE(want.has_value() && got.has_value());
  EXPECT_EQ(*got, *want);
  const CellRecords anchor =
      ScanCell(*Part(saved, SheetPart(saved, "Sheet2")), sheet->pivot_tables().front()->anchor_row(),
               sheet->pivot_tables().front()->anchor_col());
  EXPECT_NE(anchor.vm, 0U);
}

TEST(XlsbRichValueErrorReader, AForeignRichValueBlockIsNotConsumable) {
  // A rich-value block of another set (first word non-zero) is foreign, so
  // the reader leaves the metadata part and the richData parts passthrough.
  auto metadata = Part(ReadSample("c_lambda"), kMetadataPart);
  ASSERT_TRUE(metadata.has_value());
  EXPECT_TRUE(parse_metadata_bin(SpanOf(*metadata)).consumable);
  // The first rich-value block's 5003 payload is `{0, 0}`: its first word becomes 1.
  const std::vector<std::uint8_t> block(8U, 0U);
  const auto at = std::search(metadata->begin(), metadata->end(), block.begin(), block.end());
  ASSERT_NE(at, metadata->end());
  *at = 1U;
  EXPECT_FALSE(parse_metadata_bin(SpanOf(*metadata)).consumable);
}

// ---------------------------------------------------------------------------
// Round trip: a load saved again without recalculation writes the same parts.
// ---------------------------------------------------------------------------

void ExpectRoundTrip(const char* sample) {
  SCOPED_TRACE(sample);
  const std::vector<std::uint8_t> excel = ReadSample(sample);
  const Workbook wb = Load(excel);
  const std::vector<std::uint8_t> saved = Save(wb);
  ASSERT_FALSE(saved.empty());
  for (const std::string& cell : VmCells(excel)) {
    ExpectSameCell(excel, saved, cell);
  }
  ExpectSameRichParts(excel, saved);
}

TEST(XlsbRichValueErrorRoundTrip, SValue) {
  ExpectRoundTrip("s_value");
}

TEST(XlsbRichValueErrorRoundTrip, SGrid) {
  ExpectRoundTrip("s_grid");
}

TEST(XlsbRichValueErrorRoundTrip, CLambda) {
  ExpectRoundTrip("c_lambda");
}

TEST(XlsbRichValueErrorRoundTrip, Combo) {
  ExpectRoundTrip("combo");
}

TEST(XlsbRichValueErrorRoundTrip, Multi) {
  ExpectRoundTrip("multi");
}

TEST(XlsbRichValueErrorRoundTrip, CCalc) {
  ExpectRoundTrip("c_calc");
}

TEST(XlsbRichValueErrorRoundTrip, ESpill) {
  ExpectRoundTrip("e_spill");
}

TEST(XlsbRichValueErrorRoundTrip, EEdges) {
  ExpectRoundTrip("e_edges");
}

TEST(XlsbRichValueErrorRoundTrip, PvpageBlocked) {
  ExpectRoundTrip("pvpage_blocked");
}

TEST(XlsbRichValueErrorRoundTrip, PivotBlocked) {
  ExpectRoundTrip("pivot_blocked");
}

TEST(XlsbRichValueErrorRoundTrip, B3Phantom) {
  ExpectRoundTrip("b3_phantom");
}

TEST(XlsbRichValueErrorRoundTrip, B4Value) {
  ExpectRoundTrip("b4_value");
}

TEST(XlsbRichValueErrorRoundTrip, B5Fn) {
  ExpectRoundTrip("b5_fn");
}

TEST(XlsbRichValueErrorRoundTrip, VValue) {
  ExpectRoundTrip("v_value");
}

TEST(XlsbRichValueErrorRoundTrip, ABlockedSpillSurvivesAnXlsbSave) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=SEQUENCE(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(9.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Workbook back = Load(Save(wb));
  const Sheet& sheet = back.sheet(0);
  const Cell* anchor = sheet.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_TRUE(anchor->cached_value.is_error() && anchor->cached_value.as_error() == ErrorCode::Spill);
  const std::optional<BlockedSpillFootprint> footprint = sheet.blocked_spill_footprint_at(0U, 0U);
  ASSERT_TRUE(footprint.has_value());
  EXPECT_EQ(footprint->rows, 3U);
  EXPECT_EQ(footprint->cols, 1U);
  EXPECT_EQ(footprint->cause, SpillCause::kOccupied);
}

// ---------------------------------------------------------------------------
// The pivot parts of the two blocked samples, decoded against their `.xlsx`
// twins: a shared-item array, a shared string kept though the source no longer
// holds it (`u="1"`), and a missing item (`m="1"`).
// ---------------------------------------------------------------------------

struct PivotParts {
  std::vector<std::uint8_t> definition;
  std::vector<std::uint8_t> records;
  std::vector<std::uint8_t> table;
};

PivotParts PivotPartsOf(const char* sample) {
  const std::vector<std::uint8_t> package = ReadSample(sample);
  PivotParts out;
  out.definition = Part(package, "xl/pivotCache/pivotCacheDefinition1.bin").value_or(std::vector<std::uint8_t>());
  out.records = Part(package, "xl/pivotCache/pivotCacheRecords1.bin").value_or(std::vector<std::uint8_t>());
  out.table = Part(package, "xl/pivotTables/pivotTable1.bin").value_or(std::vector<std::uint8_t>());
  EXPECT_FALSE(out.definition.empty() || out.records.empty() || out.table.empty()) << sample;
  return out;
}

/// Sets the byte `at` bytes after the first occurrence of `pattern` in `bytes` to `value`.
void PatchAfter(std::vector<std::uint8_t>& bytes, const std::vector<std::uint8_t>& pattern, std::size_t at,
                std::uint8_t value) {
  const auto hit = std::search(bytes.begin(), bytes.end(), pattern.begin(), pattern.end());
  ASSERT_NE(hit, bytes.end());
  *(hit + static_cast<std::ptrdiff_t>(at)) = value;
}

Expected<pivot::PivotCache, Error> ReadCache(const PivotParts& p) {
  return read_pivot_cache_bin(SpanOf(p.definition), SpanOf(p.records));
}

TEST(XlsbPivotSamples, ANumericSharedItemArrayDecodesLikeTheTwinsNumberItems) {
  // pvpage_blocked.xlsx: Qty `<sharedItems count="5"><n v="1"/> ... <n v="5"/>`.
  const PivotParts parts = PivotPartsOf("pvpage_blocked");
  auto cache = ReadCache(parts);
  ASSERT_TRUE(static_cast<bool>(cache)) << cache.error().message;
  const pivot::PivotCacheField& qty = cache.value().fields().at(3);
  ASSERT_EQ(qty.shared_items.size(), 5U);
  for (std::size_t i = 0; i < 5U; ++i) {
    EXPECT_EQ(qty.shared_items[i], Value::number(static_cast<double>(i + 1U)));
  }
  ASSERT_EQ(cache.value().records().size(), 5U);
  EXPECT_TRUE(cache.value().records()[4].cell_is_index[3]);
  EXPECT_EQ(cache.value().records()[4].cells[3], Value::number(4.0));
}

TEST(XlsbPivotSamples, AStringArrayAndUnusedItemsDecodeLikeTheTwinsSharedItems) {
  // pivot_blocked.xlsx: Region holds N1..N5 and then North, South, East, each `u="1"`.
  const PivotParts parts = PivotPartsOf("pivot_blocked");
  auto cache = ReadCache(parts);
  ASSERT_TRUE(static_cast<bool>(cache)) << cache.error().message;
  const pivot::PivotCacheField& region = cache.value().fields().at(0);
  ASSERT_EQ(region.shared_items.size(), 8U);
  EXPECT_EQ(region.shared_items[0], Value::text("N1"));
  EXPECT_EQ(region.shared_items[4], Value::text("N5"));
  EXPECT_EQ(region.shared_items[5], Value::text("North"));
  EXPECT_EQ(region.shared_items[7], Value::text("East"));
  ASSERT_EQ(region.shared_item_is_missing.size(), 8U);
  for (std::size_t i = 0; i < 8U; ++i) {
    EXPECT_EQ(static_cast<bool>(region.shared_item_is_missing[i]), i >= 5U) << i;
  }
}

TEST(XlsbPivotSamples, AMissingItemFlagBecomesTheTwinsMissingAttribute) {
  // pivot_blocked.xlsx: `<item m="1" x="7"/><item m="1" x="5"/><item m="1" x="6"/><item x="0"/> ...`.
  const PivotParts parts = PivotPartsOf("pivot_blocked");
  auto table = read_pivot_table_bin(SpanOf(parts.table));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  const std::vector<pivot::PivotItem>& items = table.value().fields().at(0).items;
  ASSERT_EQ(items.size(), 8U);
  const std::uint32_t expected_index[] = {7U, 5U, 6U, 0U, 1U, 2U, 3U, 4U};
  for (std::size_t i = 0; i < items.size(); ++i) {
    EXPECT_EQ(items[i].cache_index, expected_index[i]) << i;
    if (i < 3U) {
      ASSERT_EQ(items[i].passthrough_attrs.size(), 1U) << i;
      EXPECT_EQ(items[i].passthrough_attrs[0], (std::pair<std::string, std::string>{"m", "1"}));
    } else {
      EXPECT_TRUE(items[i].passthrough_attrs.empty()) << i;
    }
  }
}

TEST(XlsbPivotSamples, TheValuesPseudoFieldIsKeptOutOfTheColumnOrder) {
  // pivot_blocked.xlsx: `<colFields count="2"><field x="-2"/><field x="1"/></colFields>`.
  const PivotParts parts = PivotPartsOf("pivot_blocked");
  auto table = read_pivot_table_bin(SpanOf(parts.table));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  EXPECT_EQ(table.value().col_field_order(), std::vector<std::uint32_t>{1U});
  ASSERT_TRUE(table.value().col_values_position().has_value());
  EXPECT_EQ(*table.value().col_values_position(), 0U);
  EXPECT_FALSE(table.value().row_values_position().has_value());
}

TEST(XlsbPivotSamples, LocationOffsetsAreRelativeToTheAnchorLikeTheTwins) {
  // The twins' `<location ref="A3" firstHeaderRow="0" firstDataRow="0" firstDataCol="0"/>`.
  for (const char* sample : {"pivot_blocked", "pvpage_blocked"}) {
    const PivotParts parts = PivotPartsOf(sample);
    auto table = read_pivot_table_bin(SpanOf(parts.table));
    ASSERT_TRUE(static_cast<bool>(table)) << sample << ": " << table.error().message;
    EXPECT_EQ(table.value().anchor_row(), 2U) << sample;
    EXPECT_EQ(table.value().anchor_col(), 0U) << sample;
    EXPECT_EQ(table.value().location_first_header_row(), std::optional<std::uint32_t>(0U)) << sample;
    EXPECT_EQ(table.value().location_first_data_row(), std::optional<std::uint32_t>(0U)) << sample;
    EXPECT_EQ(table.value().location_first_data_col(), std::optional<std::uint32_t>(0U)) << sample;
  }
}

TEST(XlsbPivotSamples, WhatTheTwinDoesNotPinIsRefused) {
  PivotParts parts = PivotPartsOf("pivot_blocked");
  // A shared-string attribute block other than `u="1"`.
  {
    PivotParts changed = parts;
    const std::vector<std::uint8_t> north = {'N', 0, 'o', 0, 'r', 0, 't', 0, 'h', 0};
    PatchAfter(changed.definition, north, north.size(), 2U);
    EXPECT_FALSE(static_cast<bool>(ReadCache(changed)));
  }
  // A shared-item array of an element type that is neither doubles nor strings.
  {
    PivotParts changed = parts;
    PatchAfter(changed.definition, {0xBF, 0x01, 0x2E, 0x02, 0x00}, 3U, 3U);
    EXPECT_FALSE(static_cast<bool>(ReadCache(changed)));
  }
  // An item flag besides the missing bit.
  {
    PivotParts changed = parts;
    PatchAfter(changed.table, {0x9A, 0x02, 0x07, 0x00, 0x08, 0x00, 0x07}, 4U, 0x18U);
    EXPECT_FALSE(static_cast<bool>(read_pivot_table_bin(SpanOf(changed.table))));
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
