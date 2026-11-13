//
// An xlsb save writes only the dxfs some written record names, renumbers
// every reference (modelled or kept verbatim), and leaves the model's dxf
// table as it was. Each dxf below carries a distinct number format code, so a
// reference is checked by the record it resolves to after a reload.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/dxf_record_slots.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "styles.h"
#include "table.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// The `index`-th test dxf: a number format whose code is `index + 1` zeros.
DifferentialFormat Tagged(std::size_t index) {
  DifferentialFormat dxf;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = static_cast<std::uint16_t>(164U + index);
  dxf.num_fmt_code = std::string(index + 1U, '0');
  return dxf;
}

/// What a reference names in `wb`: the tag's code, "none" or "dangling".
std::string Resolve(const Workbook& wb, std::optional<std::uint32_t> id) {
  if (!id.has_value()) {
    return "none";
  }
  return *id < wb.styles().dxfs.size() ? wb.styles().dxfs[*id].num_fmt_code : "dangling";
}

std::string Tag(std::size_t index) {
  return Tagged(index).num_fmt_code;
}

Workbook BookWithDxfs(std::size_t count, std::size_t sheets = 1U) {
  Workbook wb = Workbook::create_empty();
  for (std::size_t i = 0; i < sheets; ++i) {
    wb.add_sheet("Sheet" + std::to_string(i + 1U));
  }
  for (std::size_t i = 0; i < count; ++i) {
    wb.mutable_styles().dxfs.push_back(Tagged(i));
  }
  return wb;
}

void AddRule(Workbook& wb, std::size_t sheet, std::int32_t priority, std::optional<std::uint32_t> dxf,
             const std::string& formula = "$A$1>0") {
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 2}, {4, 2}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = priority;
  rule.formula1 = formula;
  rule.dxf_id = dxf;
  block.rules.push_back(std::move(rule));
  wb.sheet(sheet).mutable_conditional_formats().push_back(std::move(block));
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto written = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(written)) << (written ? "" : written.error().message);
  return written ? std::move(written.value()) : std::vector<std::uint8_t>{};
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Entry(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry)) << path;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

/// Every dxf slot of every worksheet, table and styles part of `package`
/// names a saved dxf.
void ExpectNoDanglingSlot(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  std::size_t saved = 0;
  for (const FramedRecord& rec : split_records(Entry(package, "xl/styles.bin"))) {
    saved += rec.type == kBrtDxf ? 1U : 0U;
  }
  for (const std::string& name : zip.list_entries()) {
    const bool scanned =
        name == "xl/styles.bin" || name.rfind("xl/worksheets/sheet", 0) == 0U || name.rfind("xl/tables/", 0) == 0U;
    if (!scanned || name.size() < 4U || name.compare(name.size() - 4U, 4U, ".bin") != 0) {
      continue;
    }
    auto scan = scan_dxf_slots(Entry(package, name));
    ASSERT_TRUE(static_cast<bool>(scan)) << name;
    for (const DxfRecordSlot& slot : scan.value().slots) {
      EXPECT_LT(slot.value, saved) << name << " record " << slot.record_type;
    }
  }
}

/// The dxf slot values of `type` records in `stream`, in order.
std::vector<std::uint32_t> SlotValues(const std::vector<std::uint8_t>& stream, std::uint16_t type) {
  std::vector<std::uint32_t> out;
  auto scan = scan_dxf_slots(stream);
  EXPECT_TRUE(static_cast<bool>(scan));
  if (scan) {
    for (const DxfRecordSlot& slot : scan.value().slots) {
      if (slot.record_type == type) {
        out.push_back(slot.value);
      }
    }
  }
  return out;
}

/// The value of attribute `name` in raw attribute text, as a dxf reference.
std::optional<std::uint32_t> AttrId(const std::string& text, const std::string& name) {
  const std::string key = " " + name + "=\"";
  const std::size_t at = text.find(key);
  if (at == std::string::npos) {
    return std::nullopt;
  }
  return static_cast<std::uint32_t>(std::stoul(text.substr(at + key.size())));
}

std::optional<std::uint32_t> FirstRuleDxf(const Workbook& wb, std::size_t sheet) {
  const auto& formats = wb.sheet(sheet).conditional_formats();
  return formats.empty() || formats[0].rules.empty() ? std::nullopt : formats[0].rules[0].dxf_id;
}

/// A table on Sheet1 whose own data dxf, second column's data dxf, colour
/// filter and sort condition name model dxfs 1, 3, 5 and 7.
TableMetadata TableNamingOddDxfs() {
  TableMetadata t;
  t.id = 1;
  t.name = "T";
  t.display_name = "T";
  t.ref = "A1:B5";
  t.sheet_index = 0;
  t.header_row = true;
  t.root_extra_attrs = " dataDxfId=\"1\"";
  TableColumn a(1, "A", "", "", "");
  TableColumn b(2, "B", "", "", "");
  b.extra_attrs = " dataDxfId=\"3\"";
  t.columns = {a, b};
  t.auto_filter_xml.set(
      auto_filter_from_xml("<autoFilter ref=\"A1:B5\"><filterColumn colId=\"0\"><colorFilter dxfId=\"5\"/>"
                           "</filterColumn></autoFilter>"));
  t.sort_state_xml =
      "<sortState ref=\"A2:B5\"><sortCondition sortBy=\"cellColor\" ref=\"B2:B5\" dxfId=\"7\"/></sortState>";
  return t;
}

/// `part` with a future-record block the model has no field for, which keeps
/// the part verbatim on load.
std::vector<std::uint8_t> WithUnknownFutureRecord(const std::vector<std::uint8_t>& part) {
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == 344U) {  // BrtEndList
      emit_record(patched, 35, std::vector<std::uint8_t>{0x02, 0x0E, 0x00, 0x00});
      emit_record(patched, 5000, std::vector<std::uint8_t>{0x01, 0x01, 0x00});
      emit_record(patched, 36, ByteSpan{});
    }
    append_record(patched, part, rec);
  }
  return patched;
}

TEST(DxfCompactionXlsb, OrphanIsDroppedAndALegacyCfRuleIsRenumbered) {
  Workbook wb = BookWithDxfs(3);
  AddRule(wb, 0, 1, 1U);
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(FirstRuleDxf(back, 0), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(Resolve(back, FirstRuleDxf(back, 0)), Tag(1));
  ExpectNoDanglingSlot(package);
}

TEST(DxfCompactionXlsb, AnX14OnlyCfRuleIsRenumbered) {
  Workbook wb = BookWithDxfs(3, 2U);
  AddRule(wb, 1, 1, 2U, "Sheet1!$A$1>0");
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(Resolve(back, FirstRuleDxf(back, 1)), Tag(2));
  ExpectNoDanglingSlot(package);
}

// A retained x14 rule is matched to its model rule by model index, then
// written with the saved one.
TEST(DxfCompactionXlsb, ARetainedX14HeadIsComparedInModelSpaceThenMapped) {
  Workbook wb = BookWithDxfs(2, 2U);
  AddRule(wb, 0, 1, 0U);
  AddRule(wb, 1, 2, 1U, "Sheet1!$A$1>0");
  Workbook loaded = Load(Save(wb));
  ASSERT_EQ(Resolve(loaded, FirstRuleDxf(loaded, 1)), Tag(1));
  loaded.sheet(0).mutable_conditional_formats().clear();
  const std::vector<std::uint8_t> package = Save(loaded);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(FirstRuleDxf(back, 1), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(Resolve(back, FirstRuleDxf(back, 1)), Tag(1));
  ExpectNoDanglingSlot(package);
}

TEST(DxfCompactionXlsb, TableColumnFilterAndSortReferencesAreRenumbered) {
  Workbook wb = BookWithDxfs(9);
  wb.mutable_tables().push_back(TableNamingOddDxfs());
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 4U);
  ASSERT_EQ(back.tables().size(), 1U);
  const TableMetadata& t = back.tables()[0];
  EXPECT_EQ(Resolve(back, AttrId(t.root_extra_attrs, "dataDxfId")), Tag(1));
  EXPECT_EQ(Resolve(back, AttrId(t.columns.at(1).extra_attrs, "dataDxfId")), Tag(3));
  ASSERT_FALSE(t.auto_filter_xml.empty());
  ASSERT_FALSE(t.auto_filter_xml.get()->columns.empty());
  EXPECT_EQ(Resolve(back, t.auto_filter_xml.get()->columns[0].color.dxf_id), Tag(5));
  EXPECT_EQ(Resolve(back, AttrId(t.sort_state_xml, "dxfId")), Tag(7));
  ExpectNoDanglingSlot(package);
}

TEST(DxfCompactionXlsb, AnIdNamingNoModelDxfIsOmitted) {
  Workbook wb = BookWithDxfs(1);
  AddRule(wb, 0, 1, 4U);
  TableMetadata t = TableNamingOddDxfs();
  t.root_extra_attrs = " dataDxfId=\"0\"";
  t.columns[1].extra_attrs = " dataDxfId=\"9\"";
  wb.mutable_tables().push_back(std::move(t));
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_FALSE(FirstRuleDxf(back, 0).has_value());
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_EQ(Resolve(back, AttrId(back.tables()[0].root_extra_attrs, "dataDxfId")), Tag(0));
  EXPECT_FALSE(AttrId(back.tables()[0].columns.at(1).extra_attrs, "dataDxfId").has_value());
  ASSERT_FALSE(back.tables()[0].auto_filter_xml.empty());
  ASSERT_FALSE(back.tables()[0].auto_filter_xml.get()->columns.empty());
  EXPECT_FALSE(back.tables()[0].auto_filter_xml.get()->columns[0].color.dxf_id.has_value());
  ExpectNoDanglingSlot(package);
}

TEST(DxfCompactionXlsb, SavingLeavesTheModelDxfTableAlone) {
  Workbook wb = BookWithDxfs(3);
  AddRule(wb, 0, 1, 2U);
  const std::vector<std::uint8_t> package = Save(wb);
  ASSERT_FALSE(package.empty());
  ASSERT_EQ(wb.styles().dxfs.size(), 3U);
  for (std::size_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(wb.styles().dxfs[i].num_fmt_code, Tag(i));
  }
  EXPECT_EQ(FirstRuleDxf(wb, 0), std::optional<std::uint32_t>(2U));
}

// A save with no reference writes no dxf; the model keeps them.
TEST(DxfCompactionXlsb, NoReferenceWritesNoDxf) {
  Workbook wb = BookWithDxfs(2);
  const Workbook back = Load(Save(wb));
  EXPECT_TRUE(back.styles().dxfs.empty());
  EXPECT_EQ(wb.styles().dxfs.size(), 2U);
}

/// A package whose table part (naming dxf 2) is kept verbatim on load, and
/// whose two CF rules name dxfs 0 and 1.
std::vector<std::uint8_t> PackageWithRetainedTable() {
  Workbook wb = BookWithDxfs(3);
  AddRule(wb, 0, 1, 0U);
  AddRule(wb, 0, 2, 1U);
  TableMetadata t = TableNamingOddDxfs();
  t.root_extra_attrs = " dataDxfId=\"2\"";
  t.columns[1].extra_attrs.clear();
  t.auto_filter_xml.reset();
  t.sort_state_xml.clear();
  wb.mutable_tables().push_back(std::move(t));
  const std::vector<std::uint8_t> package = Save(wb);
  const std::string path = "xl/tables/table1.bin";
  return RepackWithEntries(package, {{path, WithUnknownFutureRecord(Entry(package, path))}});
}

TEST(DxfCompactionXlsb, ARetainedTablePartAndTheRawStylesPartAreRenumbered) {
  const std::vector<std::uint8_t> source = PackageWithRetainedTable();
  Workbook wb = Load(source);
  ASSERT_TRUE(wb.tables().empty());
  ASSERT_EQ(wb.styles().dxfs.size(), 3U);
  wb.sheet(0).mutable_conditional_formats().clear();
  const std::vector<std::uint8_t> saved = Save(wb);

  EXPECT_EQ(SlotValues(Entry(saved, "xl/tables/table1.bin"), kBrtBeginList), std::vector<std::uint32_t>{0U});
  const Workbook back = Load(saved);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(back.styles().dxfs[0].num_fmt_code, Tag(2));
  ExpectNoDanglingSlot(saved);

  // The fresh raw styles part keeps every record but the dxf list byte for byte.
  const std::vector<std::uint8_t> before = Entry(source, "xl/styles.bin");
  const std::vector<std::uint8_t> after = Entry(saved, "xl/styles.bin");
  std::vector<std::vector<std::uint8_t>> kept_before;
  std::vector<std::vector<std::uint8_t>> dxfs_before;
  for (const FramedRecord& rec : split_records(before)) {
    std::vector<std::uint8_t> bytes(before.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                                    before.begin() + static_cast<std::ptrdiff_t>(rec.end));
    (rec.type == kBrtDxf ? dxfs_before : kept_before).push_back(std::move(bytes));
  }
  std::vector<std::vector<std::uint8_t>> kept_after;
  std::vector<std::vector<std::uint8_t>> dxfs_after;
  for (const FramedRecord& rec : split_records(after)) {
    std::vector<std::uint8_t> bytes(after.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                                    after.begin() + static_cast<std::ptrdiff_t>(rec.end));
    if (rec.type == kBrtBeginDxfs) {
      EXPECT_EQ(load_u32(rec.payload.data), 1U);
      store_u32(bytes.data() + (static_cast<std::size_t>(rec.payload.data - after.data()) - rec.begin), 3U);
    }
    (rec.type == kBrtDxf ? dxfs_after : kept_after).push_back(std::move(bytes));
  }
  ASSERT_EQ(dxfs_before.size(), 3U);
  ASSERT_EQ(dxfs_after.size(), 1U);
  EXPECT_EQ(dxfs_after[0], dxfs_before[2]);
  EXPECT_EQ(kept_after, kept_before);
}

TEST(DxfCompactionXlsb, AnIdentityPlanKeepsTheRawStylesPartAndRetainedPartBytes) {
  const std::vector<std::uint8_t> source = PackageWithRetainedTable();
  const std::vector<std::uint8_t> saved = Save(Load(source));
  EXPECT_EQ(Entry(saved, "xl/styles.bin"), Entry(source, "xl/styles.bin"));
  EXPECT_EQ(Entry(saved, "xl/tables/table1.bin"), Entry(source, "xl/tables/table1.bin"));
}

TEST(DxfCompactionXlsb, ARetainedSheetAutoFilterIsRenumbered) {
  Workbook wb = BookWithDxfs(3);
  AddRule(wb, 0, 1, 0U);
  AddRule(wb, 0, 2, 1U);
  // The first save writes the CF dxfs reversed (model 2 at index 0); the
  // injected colour filter names saved index 2, which holds model dxf 0.
  AddRule(wb, 0, 3, 2U);
  const std::vector<std::uint8_t> package = Save(wb);
  std::vector<std::uint8_t> block;
  ASSERT_EQ(
      emit_auto_filter_block(block, auto_filter_from_xml("<autoFilter ref=\"A1:A5\"><filterColumn colId=\"0\">"
                                                         "<colorFilter dxfId=\"2\"/></filterColumn></autoFilter>")),
      0U);
  // AUTOFILTER precedes the merged cells and CF blocks, right after the cell table.
  const std::string path = "xl/worksheets/sheet1.bin";
  const std::vector<std::uint8_t> sheet = Entry(package, path);
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(sheet)) {
    append_record(patched, sheet, rec);
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData)) {
      patched.insert(patched.end(), block.begin(), block.end());
    }
  }
  Workbook loaded = Load(RepackWithEntries(package, {{path, patched}}));
  ASSERT_EQ(SlotValues(loaded.sheet(0).xlsb_tail().before_merges, kBrtColorFilter), std::vector<std::uint32_t>{2U});
  loaded.sheet(0).mutable_conditional_formats().clear();
  const std::vector<std::uint8_t> saved = Save(loaded);

  EXPECT_EQ(SlotValues(Entry(saved, path), kBrtColorFilter), std::vector<std::uint32_t>{0U});
  const Workbook back = Load(saved);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(back.styles().dxfs[0].num_fmt_code, Tag(0));
  ExpectNoDanglingSlot(saved);
}

/// A table on Sheet1 whose only dxf reference is its own data dxf.
TableMetadata TableNamingDxf(std::uint32_t dxf) {
  TableMetadata t = TableNamingOddDxfs();
  t.root_extra_attrs = " dataDxfId=\"" + std::to_string(dxf) + "\"";
  t.columns[1].extra_attrs.clear();
  t.auto_filter_xml.reset();
  t.sort_state_xml.clear();
  return t;
}

/// The model dxf each saved BrtDXF of `package` carries, as tags.
std::vector<std::string> SavedTags(const std::vector<std::uint8_t>& package) {
  const Workbook back = Load(package);
  std::vector<std::string> out;
  for (const DifferentialFormat& dxf : back.styles().dxfs) {
    out.push_back(dxf.num_fmt_code);
  }
  return out;
}

// Excel writes a record per user: a CF rule and a table attribute naming one
// model dxf each get their own record, and each slot resolves to its own.
TEST(DxfCompactionXlsb, OneDxfUsedByACfRuleAndATableIsSavedOncePerUser) {
  Workbook wb = BookWithDxfs(2);
  AddRule(wb, 0, 1, 0U);
  wb.mutable_tables().push_back(TableNamingDxf(0U));
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 2U);
  EXPECT_EQ(back.styles().dxfs[0].num_fmt_code, Tag(0));
  EXPECT_EQ(back.styles().dxfs[1].num_fmt_code, Tag(0));
  std::vector<std::vector<std::uint8_t>> records;
  const std::vector<std::uint8_t> styles = Entry(package, "xl/styles.bin");
  for (const FramedRecord& rec : split_records(styles)) {
    if (rec.type == kBrtDxf) {
      records.emplace_back(styles.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                           styles.begin() + static_cast<std::ptrdiff_t>(rec.end));
    }
  }
  ASSERT_EQ(records.size(), 2U);
  EXPECT_EQ(records[0], records[1]);
  // The table registers before CF, so it takes the later record.
  EXPECT_EQ(SlotValues(Entry(package, "xl/tables/table1.bin"), kBrtBeginList), std::vector<std::uint32_t>{1U});
  EXPECT_EQ(FirstRuleDxf(back, 0), std::optional<std::uint32_t>(0U));
  ExpectNoDanglingSlot(package);
}

// CF rules of one sheet naming the same dxf take a record each.
TEST(DxfCompactionXlsb, CfRulesOfOneSheetSharingADxfTakeARecordEach) {
  Workbook wb = BookWithDxfs(2);
  AddRule(wb, 0, 1, 0U, "$A$1>0");
  AddRule(wb, 0, 2, 0U, "$A$1>1");
  const std::vector<std::uint8_t> package = Save(wb);
  const Workbook back = Load(package);
  ASSERT_EQ(back.styles().dxfs.size(), 2U);
  const auto& formats = back.sheet(0).conditional_formats();
  ASSERT_EQ(formats.size(), 2U);
  ASSERT_EQ(formats[0].rules.size(), 1U);
  ASSERT_EQ(formats[1].rules.size(), 1U);
  EXPECT_EQ(formats[0].rules[0].dxf_id, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(formats[1].rules[0].dxf_id, std::optional<std::uint32_t>(0U));
  ExpectNoDanglingSlot(package);
}

// Registration: the table's dxf, then CF by tab order; records are reversed.
TEST(DxfCompactionXlsb, RecordsFollowExcelRegistrationOrderReversed) {
  Workbook wb = BookWithDxfs(3, 2U);
  AddRule(wb, 0, 1, 0U);
  AddRule(wb, 1, 1, 1U, "Sheet1!$A$1>0");
  wb.mutable_tables().push_back(TableNamingDxf(2U));
  const std::vector<std::uint8_t> package = Save(wb);
  EXPECT_EQ(SavedTags(package), (std::vector<std::string>{Tag(1), Tag(0), Tag(2)}));
  const Workbook back = Load(package);
  EXPECT_EQ(FirstRuleDxf(back, 0), std::optional<std::uint32_t>(1U));
  EXPECT_EQ(FirstRuleDxf(back, 1), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(SlotValues(Entry(package, "xl/tables/table1.bin"), kBrtBeginList), std::vector<std::uint32_t>{2U});
  ExpectNoDanglingSlot(package);
}

// CF rules register by priority, not by document order.
TEST(DxfCompactionXlsb, CfRulesOfOneBlockRegisterByPriority) {
  Workbook wb = BookWithDxfs(2);
  AddRule(wb, 0, 2, 0U, "$A$1>0");
  AddRule(wb, 0, 1, 1U, "$A$1>1");
  auto& formats = wb.sheet(0).mutable_conditional_formats();
  formats[0].rules.push_back(formats[1].rules[0]);
  formats.pop_back();
  const std::vector<std::uint8_t> package = Save(wb);
  // Priority 1 (model 1) registers first, so it is written last.
  EXPECT_EQ(SavedTags(package), (std::vector<std::string>{Tag(0), Tag(1)}));
  const Workbook back = Load(package);
  ASSERT_EQ(back.sheet(0).conditional_formats().size(), 1U);
  for (const auto& rule : back.sheet(0).conditional_formats()[0].rules) {
    ASSERT_TRUE(rule.dxf_id.has_value());
    EXPECT_EQ(Resolve(back, rule.dxf_id), rule.priority == 1 ? Tag(1) : Tag(0));
  }
  ExpectNoDanglingSlot(package);
}

TEST(DxfCompactionXlsb, CfBlocksRegisterInDocumentOrder) {
  Workbook wb = BookWithDxfs(2);
  AddRule(wb, 0, 2, 0U, "$A$1>0");
  AddRule(wb, 0, 1, 1U, "$A$1>1");
  // The first block registers first though its rule has the lower priority.
  EXPECT_EQ(SavedTags(Save(wb)), (std::vector<std::string>{Tag(1), Tag(0)}));
}

// No reader decodes a retained pivot table part's dxf records, so the table
// is written whole and every id stays as it was.
TEST(DxfCompactionXlsb, ARetainedPivotPartKeepsEveryDxf) {
  const std::vector<std::uint8_t> source =
      ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_base.xlsb");
  const Workbook original = Load(source);
  Workbook wb = Load(source);
  bool has_pivot_part = false;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    has_pivot_part = has_pivot_part || part.retained_origin == PassthroughPart::RetainedOrigin::kPivotTable;
  }
  ASSERT_TRUE(has_pivot_part);
  wb.mutable_styles().dxfs.push_back(Tagged(0));
  const Workbook back = Load(Save(wb));
  ASSERT_EQ(back.styles().dxfs.size(), original.styles().dxfs.size() + 1U);
  EXPECT_EQ(back.styles().dxfs.back().num_fmt_code, Tag(0));
  for (std::size_t s = 0; s < original.sheet_count(); ++s) {
    const auto& expected = original.sheet(s).conditional_formats();
    const auto& actual = back.sheet(s).conditional_formats();
    ASSERT_EQ(actual.size(), expected.size());
    for (std::size_t f = 0; f < expected.size(); ++f) {
      ASSERT_EQ(actual[f].rules.size(), expected[f].rules.size());
      for (std::size_t r = 0; r < expected[f].rules.size(); ++r) {
        EXPECT_EQ(actual[f].rules[r].dxf_id, expected[f].rules[r].dxf_id);
      }
    }
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
