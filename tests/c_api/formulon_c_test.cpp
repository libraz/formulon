//
// Stable C ABI (`src/c_api/formulon_c.h`) end-to-end tests.
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface declared in
// `formulon_c.h`. The same surface drives the CLI binary, the WASM
// embind layer, and any external language binding, so this file is the
// regression harness for the contract.

#include "c_api/formulon_c.h"

#include <algorithm>
#include <atomic>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/format_detect.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace {

// RAII wrapper so the workbook handle is released even on test failure.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// RAII wrapper for buffers returned by `fm_workbook_save`.
struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

void MarkZipEntriesEncrypted(std::vector<std::uint8_t>& bytes) {
  for (std::size_t i = 0; i + 8U <= bytes.size(); ++i) {
    const bool local = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 3U && bytes[i + 3U] == 4U;
    const bool central = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 1U && bytes[i + 3U] == 2U;
    if (local || central) {
      const std::size_t flag_offset = i + (local ? 6U : 8U);
      bytes[flag_offset] = static_cast<std::uint8_t>(bytes[flag_offset] | 0x01U);
    }
  }
}

std::vector<std::uint8_t> AppendEmptyZipEntry(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  const std::uint8_t signature[] = {0x50, 0x4b, 0x05, 0x06};
  const auto eocd_it = std::find_end(bytes.begin(), bytes.end(), std::begin(signature), std::end(signature));
  EXPECT_NE(eocd_it, bytes.end());
  if (eocd_it == bytes.end()) {
    return {};
  }
  const std::size_t eocd = static_cast<std::size_t>(eocd_it - bytes.begin());
  const auto read16 = [&](std::size_t offset) {
    return static_cast<std::uint16_t>(bytes[offset] | (static_cast<std::uint16_t>(bytes[offset + 1U]) << 8U));
  };
  const auto read32 = [&](std::size_t offset) {
    return static_cast<std::uint32_t>(bytes[offset] | (static_cast<std::uint32_t>(bytes[offset + 1U]) << 8U) |
                                      (static_cast<std::uint32_t>(bytes[offset + 2U]) << 16U) |
                                      (static_cast<std::uint32_t>(bytes[offset + 3U]) << 24U));
  };
  const std::uint16_t count = read16(eocd + 10U);
  const std::uint32_t central_size = read32(eocd + 12U);
  const std::uint32_t central_offset = read32(eocd + 16U);
  const auto put16 = [](std::vector<std::uint8_t>& out, std::uint16_t value) {
    out.push_back(static_cast<std::uint8_t>(value));
    out.push_back(static_cast<std::uint8_t>(value >> 8U));
  };
  const auto put32 = [](std::vector<std::uint8_t>& out, std::uint32_t value) {
    out.push_back(static_cast<std::uint8_t>(value));
    out.push_back(static_cast<std::uint8_t>(value >> 8U));
    out.push_back(static_cast<std::uint8_t>(value >> 16U));
    out.push_back(static_cast<std::uint8_t>(value >> 24U));
  };
  const std::string encoded_name(name);
  std::vector<std::uint8_t> local;
  put32(local, 0x04034b50U);
  put16(local, 20);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put32(local, 0);
  put32(local, 0);
  put32(local, 0);
  put16(local, static_cast<std::uint16_t>(encoded_name.size()));
  put16(local, 0);
  local.insert(local.end(), encoded_name.begin(), encoded_name.end());
  std::vector<std::uint8_t> central;
  put32(central, 0x02014b50U);
  put16(central, 20);
  put16(central, 20);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, 0);
  put32(central, 0);
  put16(central, static_cast<std::uint16_t>(encoded_name.size()));
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, central_offset);
  central.insert(central.end(), encoded_name.begin(), encoded_name.end());
  std::vector<std::uint8_t> out;
  out.insert(out.end(), bytes.begin(), bytes.begin() + central_offset);
  out.insert(out.end(), local.begin(), local.end());
  out.insert(out.end(), bytes.begin() + central_offset, bytes.begin() + central_offset + central_size);
  out.insert(out.end(), central.begin(), central.end());
  put32(out, 0x06054b50U);
  put16(out, 0);
  put16(out, 0);
  put16(out, count + 1U);
  put16(out, count + 1U);
  put32(out, central_size + static_cast<std::uint32_t>(central.size()));
  put32(out, central_offset + static_cast<std::uint32_t>(local.size()));
  put16(out, 0);
  return out;
}

}  // namespace

TEST(FormulonCApi, CreateAndDestroy) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NE(wb.handle, nullptr);
  // create() always seeds a single Sheet1.
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_sheet_name(wb.handle, 0, &name), 0);
  ASSERT_NE(name, nullptr);
  EXPECT_STREQ(name, "Sheet1");
}

TEST(FormulonCApi, TableCreateUpdateRemoveRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;
  ASSERT_EQ(
      fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "TableStyleMedium2", 1, 0, &index),
      0);
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(fm_workbook_table_count(wb.handle), 1U);

  const char* name = nullptr;
  const char* display_name = nullptr;
  const char* ref = nullptr;
  size_t sheet = 99;
  ASSERT_EQ(fm_workbook_table_at(wb.handle, index, &name, &display_name, &ref, &sheet), 0);
  EXPECT_STREQ(name, "Sales");
  EXPECT_STREQ(display_name, "Sales");
  EXPECT_STREQ(ref, "A1:B3");
  EXPECT_EQ(sheet, 0U);

  ASSERT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B4", "TableStyleLight9", 1, 1), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_table_at(loaded.handle, 0, &name, &display_name, &ref, &sheet), 0);
  EXPECT_STREQ(ref, "A1:B4");

  ASSERT_EQ(fm_workbook_table_remove(loaded.handle, 0), 0);
  EXPECT_EQ(fm_workbook_table_count(loaded.handle), 0U);
}

TEST(FormulonCApi, TableRangeMustMatchTheColumnList) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;

  // Two headers cannot describe a three-column range, and a range that is
  // not a plain A1 area has no width to check against at all.
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "A1:C3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "Sheet1!A1:B3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "$A$1:$B$3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_EQ(fm_workbook_table_count(wb.handle), 0U);

  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  // Growing rows is fine; growing columns would orphan the column list.
  EXPECT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B9", "", 1, 0), 0);
  EXPECT_NE(fm_workbook_table_update(wb.handle, index, "A1:C9", "", 1, 0), 0);

  const char* duplicate_columns[] = {"Product", "product"};
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "D1:E3", "Other", "Other", duplicate_columns, 2, "", 1, 0, &index),
            0);
}

TEST(FormulonCApi, TableUpdatePreservesRawMetadataAcrossSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;
  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "CustomStyle", 1, 1, &index),
            0);

  // Seed payload that the C ABI deliberately does not model. The update must
  // rewrite only the opening autoFilter ref and retain every raw payload.
  auto& table = wb.handle->wb->mutable_tables()[index];
  table.auto_filter_xml =
      "<autoFilter ref=\"A1:B3\"><filterColumn colId=\"0\"><filters><filter val=\"West\"/></filters></filterColumn>"
      "</autoFilter>";
  table.sort_state_xml = "<sortState ref=\"A1:B3\"><sortCondition ref=\"B2:B3\" descending=\"1\"/></sortState>";
  table.table_style_info_xml = "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>";
  table.ext_lst_xml = "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>";

  ASSERT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B4", nullptr, -1, -1), 0);
  EXPECT_EQ(table.ref, "A1:B4");
  EXPECT_TRUE(table.header_row);
  EXPECT_TRUE(table.totals_row);
  EXPECT_NE(table.auto_filter_xml.find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(table.auto_filter_xml.find("filterColumn"), std::string::npos);
  EXPECT_NE(table.sort_state_xml.find("ref=\"A1:B3\""), std::string::npos);
  EXPECT_NE(table.sort_state_xml.find("sortCondition"), std::string::npos);
  EXPECT_EQ(table.table_style_info_xml, "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>");
  EXPECT_EQ(table.ext_lst_xml,
            "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  const auto& reloaded = loaded.handle->wb->tables()[0];
  EXPECT_EQ(reloaded.ref, "A1:B4");
  EXPECT_TRUE(reloaded.header_row);
  EXPECT_TRUE(reloaded.totals_row);
  EXPECT_NE(reloaded.auto_filter_xml.find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(reloaded.auto_filter_xml.find("filterColumn"), std::string::npos);
  EXPECT_NE(reloaded.auto_filter_xml.find("West"), std::string::npos);
  EXPECT_EQ(reloaded.sort_state_xml,
            "<sortState ref=\"A1:B3\"><sortCondition ref=\"B2:B3\" descending=\"1\"/></sortState>");
  EXPECT_EQ(reloaded.table_style_info_xml, "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>");
  EXPECT_EQ(reloaded.ext_lst_xml,
            "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>");

  formulon::io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{saved.data, saved.len})));
  auto table_part_or = zip.read_entry("xl/tables/table1.xml");
  ASSERT_TRUE(static_cast<bool>(table_part_or)) << table_part_or.error().message;
  const std::string table_xml(table_part_or.value().begin(), table_part_or.value().end());
  const std::size_t auto_filter_pos = table_xml.find("<autoFilter");
  const std::size_t sort_state_pos = table_xml.find("<sortState");
  const std::size_t table_columns_pos = table_xml.find("<tableColumns");
  const std::size_t style_info_pos = table_xml.find("<tableStyleInfo");
  const std::size_t ext_lst_pos = table_xml.find("<extLst");
  ASSERT_NE(auto_filter_pos, std::string::npos);
  ASSERT_NE(sort_state_pos, std::string::npos);
  ASSERT_NE(table_columns_pos, std::string::npos);
  ASSERT_NE(style_info_pos, std::string::npos);
  ASSERT_NE(ext_lst_pos, std::string::npos);
  EXPECT_LT(auto_filter_pos, sort_state_pos);
  EXPECT_LT(sort_state_pos, table_columns_pos);
  EXPECT_LT(table_columns_pos, style_info_pos);
  EXPECT_LT(style_info_pos, ext_lst_pos);
  EXPECT_NE(table_xml.find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(table_xml.find("filterColumn"), std::string::npos);
  EXPECT_NE(table_xml.find("sortCondition ref=\"B2:B3\""), std::string::npos);
  EXPECT_NE(table_xml.find("futureTableData value=\"kept\""), std::string::npos);
}

TEST(FormulonCApi, PaginationSnapshotExposesBreaksAndUsedRangeFallback) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 48, 0, 2.0), 0);

  fm_pagination_t* pagination = nullptr;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &pagination), 0);
  ASSERT_NE(pagination, nullptr);
  EXPECT_EQ(fm_pagination_page_count(pagination), 2U);
  // No explicit _xlnm.Print_Area is reported even though pagination falls
  // back to the used range internally.
  EXPECT_EQ(fm_pagination_print_area_count(pagination), 0U);
  ASSERT_EQ(fm_pagination_horizontal_break_count(pagination), 1U);
  std::uint32_t row = 0;
  ASSERT_EQ(fm_pagination_horizontal_break_at(pagination, 0, &row), 0);
  EXPECT_EQ(row, 44U);
  EXPECT_EQ(fm_pagination_vertical_break_count(pagination), 0U);
  EXPECT_NE(fm_pagination_horizontal_break_at(pagination, 1, &row), 0);
  fm_pagination_destroy(pagination);

  EXPECT_NE(fm_workbook_paginate(nullptr, 0, &pagination), 0);
  EXPECT_NE(fm_workbook_paginate(wb.handle, 0, nullptr), 0);
}

TEST(FormulonCApi, CellPhoneticCanBeReadClearedAndRejectsInvalidArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);

  const char* phonetic = nullptr;
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "かんじ");

  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, ""), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "");

  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "文字列"), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "");

  EXPECT_NE(fm_workbook_set_cell_phonetic(nullptr, 0, 0, 0, "x"), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, nullptr), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic(wb.handle, 0, formulon::Sheet::kMaxRows, 0, "x"), 0);
  EXPECT_NE(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApi, LoadMapsCorruptAndEncryptedContainersToIoErrors) {
  fm_workbook_t* loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  const std::vector<std::uint8_t> garbage = {0x01U, 0x02U, 0x03U, 0x04U};
  EXPECT_EQ(fm_workbook_load(garbage.data(), garbage.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipCorrupt));
  EXPECT_EQ(loaded, nullptr);

  loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  const std::vector<std::uint8_t> cdfv2 = {0xD0U, 0xCFU, 0x11U, 0xE0U, 0xA1U, 0xB1U, 0x1AU, 0xE1U};
  EXPECT_EQ(fm_workbook_load(cdfv2.data(), cdfv2.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipEncrypted));
  EXPECT_EQ(loaded, nullptr);

  formulon::Workbook source = formulon::Workbook::create();
  auto saved = source.save();
  ASSERT_TRUE(static_cast<bool>(saved));
  std::vector<std::uint8_t> encrypted = saved.value();
  MarkZipEntriesEncrypted(encrypted);
  loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  EXPECT_EQ(fm_workbook_load(encrypted.data(), encrypted.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipEncrypted));
  EXPECT_EQ(loaded, nullptr);
}

TEST(FormulonCApi, CreateEmptyAndAddSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 0U);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Data"), 0);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Stats"), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);

  const char* n0 = nullptr;
  const char* n1 = nullptr;
  ASSERT_EQ(fm_workbook_sheet_name(wb.handle, 0, &n0), 0);
  ASSERT_EQ(fm_workbook_sheet_name(wb.handle, 1, &n1), 0);
  EXPECT_STREQ(n0, "Data");
  EXPECT_STREQ(n1, "Stats");
}

TEST(FormulonCApi, AddSheetValidatesNameAndRejectsDuplicate) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Data"), 0);
  // Duplicate (case-insensitive), forbidden character, and empty name are
  // all rejected now that the public add surface shares the rename
  // validator instead of silently accepting anything.
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "data"), 0);
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "a/b"), 0);
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, ""), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
  // A valid, distinct name still succeeds.
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Stats"), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);
}

TEST(FormulonCApi, NumberLiteralRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.5), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 42.5);
}

TEST(FormulonCApi, BoolAndBlankSetters) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_bool(wb.handle, 0, 0, 0, 1), 0);
  ASSERT_EQ(fm_workbook_set_bool(wb.handle, 0, 1, 0, 0), 0);
  ASSERT_EQ(fm_workbook_set_blank(wb.handle, 0, 2, 0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);

  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 0);

  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
}

TEST(FormulonCApi, ErrorSetterStoresStaticErrorLiteral) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_error(wb.handle, 0, 0, 0, 1), 0);  // ErrorCode::Div0
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  EXPECT_EQ(v.u.error_code, 1);
}

TEST(FormulonCApi, ErrorSetterRejectsInvalidErrorCode) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_workbook_set_error(wb.handle, 0, 0, 0, -1), 0);
  EXPECT_NE(fm_workbook_set_error(wb.handle, 0, 0, 0, 999), 0);
}

TEST(FormulonCApi, FormulaNumericResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 10.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=A1*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 20.0);
}

TEST(FormulonCApi, FormulaTextResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=UPPER(\"hello\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "HELLO");
}

TEST(FormulonCApi, TextSetterRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Use a stack buffer that goes out of scope before recalc to confirm
  // the handle interns the bytes.
  {
    char tmp[] = "hello";
    ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, tmp), 0);
    // Mutate the source buffer to prove we're not aliasing.
    tmp[0] = 'X';
    (void)tmp[0];
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "hello");
  // Verify the returned pointer is actually NUL-terminated by inspecting
  // strlen against the documented length.
  EXPECT_EQ(std::strlen(v.u.text), 5U);
}

TEST(FormulonCApi, FormulaErrorSurfacesAsValueError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1/0"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  // Excel's `#DIV/0!` is `ErrorCode::Div0`, ordinal 1.
  EXPECT_EQ(v.u.error_code, 1);
}

TEST(FormulonCApi, NullWorkbookSetsBindingError) {
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(nullptr, 0, 0, 0, &v);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);

  // Also exercise the create-side NULL path.
  rc = fm_workbook_create(nullptr);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);
}

TEST(FormulonCApi, OutOfRangeSheetIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_status_t rc = fm_workbook_set_number(wb.handle, 99, 0, 0, 1.0);
  EXPECT_NE(rc, 0);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  // Context should reference the offending sheet index.
  std::string ctx = fm_last_error_context();
  EXPECT_NE(ctx.find("sheet_index"), std::string::npos) << "context=" << ctx;
}

TEST(FormulonCApi, OutOfGridCoordinateIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // A column near the top of the u32 range would otherwise resize a row
  // vector to billions of cells. It must be rejected before the storage
  // layer, not turned into a multi-GB allocation.
  const std::uint32_t kBadCol = 4'000'000'000U;
  fm_status_t rc = fm_workbook_set_number(wb.handle, 0, 0, kBadCol, 1.0);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  std::string ctx = fm_last_error_context();
  EXPECT_NE(ctx.find("col"), std::string::npos) << "context=" << ctx;

  // Row at the Excel ceiling (kMaxRows) is one past the last addressable
  // row and must also be rejected.
  EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, formulon::Sheet::kMaxRows, 0, "=1"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_set_blank(wb.handle, 0, 0, formulon::Sheet::kMaxCols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // The last in-grid cell still round-trips.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols - 1U, 3.0),
            0);
}

TEST(FormulonCApi, SuccessClearsPreviousLastError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // First, force an error so the thread-local diagnostic is populated.
  fm_value_t v{};
  ASSERT_NE(fm_workbook_get_value(nullptr, 0, 0, 0, &v), 0);
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);
  // Now drive a success and confirm the diagnostic was cleared.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  EXPECT_EQ(std::strlen(fm_last_error_message()), 0U);
  EXPECT_EQ(std::strlen(fm_last_error_context()), 0U);
}

TEST(FormulonCApi, SaveLoadRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=A1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(loaded.handle), 2U);

  const char* sheet0 = nullptr;
  ASSERT_EQ(fm_workbook_sheet_name(loaded.handle, 0, &sheet0), 0);
  EXPECT_STREQ(sheet0, "Sheet1");

  // The literal A1=7 must round-trip; the formula B1=A1+1 may need a
  // recalc on the loaded workbook to populate its cached value.
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t a1{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &a1), 0);
  EXPECT_EQ(a1.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(a1.u.number, 7.0);
}

TEST(FormulonCApi, SaveExXlsxMatchesSave) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard xlsx_buf;
  ASSERT_EQ(fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &xlsx_buf.data, &xlsx_buf.len), 0);
  ASSERT_NE(xlsx_buf.data, nullptr);
  EXPECT_GT(xlsx_buf.len, 0U);

  // `FM_WORKBOOK_FORMAT_XLSX` must produce an OOXML container, so the
  // C ABI's own format sniff (used by `fm_workbook_load`) reports it as
  // such rather than xlsb.
  formulon::io::ByteSpan xlsx_span{xlsx_buf.data, xlsx_buf.len};
  EXPECT_EQ(formulon::io::detect_workbook_format(xlsx_span), formulon::io::WorkbookFormat::Ooxml);
}

TEST(FormulonCApi, SaveExXlsbProducesLoadableXlsbContainer) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard xlsb_buf;
  ASSERT_EQ(fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb_buf.data, &xlsb_buf.len), 0);
  ASSERT_NE(xlsb_buf.data, nullptr);
  EXPECT_GT(xlsb_buf.len, 0U);

  // The bytes must be a real MS-XLSB package (declares `xl/workbook.bin`,
  // not `xl/workbook.xml`), and must load back through the byte-only
  // C ABI, which auto-detects the container from its contents.
  formulon::io::ByteSpan xlsb_span{xlsb_buf.data, xlsb_buf.len};
  EXPECT_EQ(formulon::io::detect_workbook_format(xlsb_span), formulon::io::WorkbookFormat::Xlsb);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(xlsb_buf.data, xlsb_buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 42.0);
}

TEST(FormulonCApi, SaveXlsbWithResultReportsUnsupportedFormulaDowngrade) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=@A1:A10"), 0);

  BufferGuard xlsb_buf;
  size_t downgraded = 0;
  ASSERT_EQ(fm_workbook_save_xlsb_with_result(wb.handle, &xlsb_buf.data, &xlsb_buf.len, &downgraded), 0);
  ASSERT_NE(xlsb_buf.data, nullptr);
  EXPECT_GT(xlsb_buf.len, 0U);
  EXPECT_EQ(downgraded, 1U);
}

TEST(FormulonCApi, SaveExWithDiagnosticsReportsBothXlsbCountersAndKeepsLegacyShapes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=@A1:A10"), 0);

  auto& sheet = wb.handle->workbook().sheet(0);
  formulon::Hyperlink hyperlink;
  hyperlink.target = "https://example.com";
  sheet.mutable_hyperlinks().push_back(std::move(hyperlink));
  sheet.mutable_validations().push_back(formulon::DataValidation{});
  sheet.set_auto_filter_xml("<autoFilter ref=\"A1:B2\"/>");

  BufferGuard xlsx_buf;
  size_t downgraded = 99U;
  size_t deferred = 99U;
  ASSERT_EQ(fm_workbook_save_ex_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &xlsx_buf.data, &xlsx_buf.len,
                                                 &downgraded, &deferred),
            0);
  EXPECT_EQ(downgraded, 0U);
  EXPECT_EQ(deferred, 0U);

  BufferGuard xlsb_buf;
  downgraded = 0U;
  deferred = 0U;
  ASSERT_EQ(fm_workbook_save_ex_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb_buf.data, &xlsb_buf.len,
                                                 &downgraded, &deferred),
            0);
  EXPECT_EQ(downgraded, 1U);
  EXPECT_EQ(deferred, 3U);

  BufferGuard legacy_buf;
  size_t legacy_downgraded = 0U;
  ASSERT_EQ(fm_workbook_save_xlsb_with_result(wb.handle, &legacy_buf.data, &legacy_buf.len, &legacy_downgraded), 0);
  EXPECT_EQ(legacy_downgraded, downgraded);
  EXPECT_GT(legacy_buf.len, 0U);
}

TEST(FormulonCApi, XlsbReadDiagnosticsAreZeroForCleanRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.0), 0);

  BufferGuard bytes;
  ASSERT_EQ(fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &bytes.data, &bytes.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(bytes.data, bytes.len, &loaded.handle), 0);

  size_t formulas = 99U;
  size_t names = 99U;
  ASSERT_EQ(fm_workbook_xlsb_read_diagnostics(loaded.handle, &formulas, &names), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);

  size_t dropped = 99U;
  ASSERT_EQ(fm_workbook_xlsb_read_diagnostics_ex(loaded.handle, &formulas, &names, &dropped), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  EXPECT_EQ(dropped, 0U);
}

TEST(FormulonCApi, XlsbReadDiagnosticsExReportsDroppedPartAndInitializesOutputsOnFailure) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  size_t formulas = 99U;
  size_t names = 99U;
  size_t dropped = 99U;
  ASSERT_EQ(fm_workbook_xlsb_read_diagnostics_ex(wb.handle, &formulas, &names, &dropped), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  EXPECT_EQ(dropped, 0U);

  uint8_t* bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  size_t len = 99U;
  size_t downgraded = 99U;
  size_t deferred = 99U;
  EXPECT_NE(
      fm_workbook_save_ex_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_UNKNOWN, &bytes, &len, &downgraded, &deferred),
      0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(downgraded, 0U);
  EXPECT_EQ(deferred, 0U);

  formulas = 99U;
  names = 99U;
  dropped = 99U;
  EXPECT_NE(fm_workbook_xlsb_read_diagnostics_ex(nullptr, &formulas, &names, &dropped), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  EXPECT_EQ(dropped, 0U);
}

TEST(FormulonCApi, XlsbReadDiagnosticsPreservesUnknownPartFromDeterministicFixture) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  BufferGuard bytes;
  ASSERT_EQ(fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &bytes.data, &bytes.len), 0);
  const std::vector<std::uint8_t> fixture =
      AppendEmptyZipEntry(std::vector<std::uint8_t>(bytes.data, bytes.data + bytes.len), "xl/dropped.bin");
  ASSERT_FALSE(fixture.empty());
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(fixture.data(), fixture.size(), &loaded.handle), 0);
  size_t formulas = 0;
  size_t names = 0;
  size_t dropped = 0;
  ASSERT_EQ(fm_workbook_xlsb_read_diagnostics_ex(loaded.handle, &formulas, &names, &dropped), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  // Unknown package parts are retained as passthrough data, so loading the
  // fixture is lossless and the dropped-part diagnostic remains zero.
  EXPECT_EQ(dropped, 0U);
}

TEST(FormulonCApi, DiagnosticFailuresNameTheInvokedLegacyOrExtendedSymbol) {
  size_t formulas = 77U;
  size_t names = 77U;
  EXPECT_NE(fm_workbook_xlsb_read_diagnostics(nullptr, &formulas, &names), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_xlsb_read_diagnostics: NULL argument");

  formulas = 77U;
  names = 77U;
  size_t dropped = 77U;
  EXPECT_NE(fm_workbook_xlsb_read_diagnostics_ex(nullptr, &formulas, &names, &dropped), 0);
  EXPECT_EQ(formulas, 0U);
  EXPECT_EQ(names, 0U);
  EXPECT_EQ(dropped, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_xlsb_read_diagnostics_ex: NULL argument");

  uint8_t* bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  size_t len = 77U;
  EXPECT_NE(fm_workbook_save_ex(nullptr, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save_ex: NULL argument");

  size_t downgraded = 77U;
  EXPECT_NE(fm_workbook_save_xlsb_with_result(nullptr, &bytes, &len, &downgraded), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(downgraded, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save_xlsb_with_result: NULL argument");

  size_t deferred = 77U;
  bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  len = 77U;
  EXPECT_NE(
      fm_workbook_save_ex_with_diagnostics(nullptr, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len, &downgraded, &deferred), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(downgraded, 0U);
  EXPECT_EQ(deferred, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save_ex_with_diagnostics: NULL argument");
}

TEST(FormulonCApi, MemoryUsageTracksTheWorkbookAndRejectsNulls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  size_t empty = 0U;
  ASSERT_EQ(fm_workbook_memory_usage(wb.handle, &empty), 0);
  EXPECT_GT(empty, 0U);

  for (uint32_t row = 0; row < 200U; ++row) {
    ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, row, 0, "=1+2"), 0);
  }
  size_t filled = 0U;
  ASSERT_EQ(fm_workbook_memory_usage(wb.handle, &filled), 0);
  EXPECT_GT(filled, empty);

  // Both NULL arguments are rejected rather than silently reporting 0,
  // which is what lets a binding distinguish "empty" from "broken".
  size_t scratch = 0U;
  EXPECT_NE(fm_workbook_memory_usage(nullptr, &scratch), 0);
  EXPECT_NE(fm_workbook_memory_usage(wb.handle, nullptr), 0);
}

TEST(FormulonCApi, SaveExRejectsUnknownFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const std::int32_t invalid_formats[] = {99, std::numeric_limits<std::int32_t>::min(),
                                          std::numeric_limits<std::int32_t>::max()};
  for (const std::int32_t raw : invalid_formats) {
    BufferGuard buf;
    EXPECT_EQ(fm_workbook_save_ex(wb.handle, raw, &buf.data, &buf.len),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    EXPECT_EQ(buf.data, nullptr);
    EXPECT_EQ(buf.len, 0U);
    std::size_t downgraded = 99;
    std::size_t deferred = 99;
    EXPECT_EQ(fm_workbook_save_ex_with_diagnostics(wb.handle, raw, &buf.data, &buf.len, &downgraded, &deferred),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    EXPECT_EQ(buf.data, nullptr);
    EXPECT_EQ(buf.len, 0U);
    EXPECT_EQ(downgraded, 0U);
    EXPECT_EQ(deferred, 0U);
  }
}

TEST(FormulonCApi, LoadRoutesXlsbBytesToXlsbReader) {
  // Build a minimal `.xlsb` byte stream via the engine writer, then load
  // it through the byte-only C ABI. The load boundary must detect the
  // xlsb container and route to `read_xlsb` rather than failing in the
  // OOXML reader with a "missing xl/workbook.xml" diagnostic.
  formulon::Workbook src = formulon::Workbook::create_empty();
  formulon::Sheet& s = src.add_sheet("S");
  s.set_cell_value(0U, 0U, formulon::Value::number(123.5));
  auto xlsb_or = formulon::io::xlsb::write_xlsb(src);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  const std::vector<std::uint8_t>& xlsb = xlsb_or.value();

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(xlsb.data(), xlsb.size(), &loaded.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(loaded.handle), 1U);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 123.5);
}

TEST(FormulonCApi, SaveLoadFormulaTextResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=UPPER(\"world\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "WORLD");
}

TEST(FormulonCApi, SaveLoadFormulaBoolResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 0, "=TRUE()"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 2, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);
}

TEST(FormulonCApi, SaveLoadFormulaErrorResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 3, 0, "=1/0"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 3, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  EXPECT_EQ(v.u.error_code, 1);
}

TEST(FormulonCApi, IterativeOptionsConverge) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, 1, 100, 1e-6), 0);

  // A1 = 0.5 * (A1 + 2) converges to 2 from the blank initial cache
  // value. Unlike Newton's method, this intentionally needs no separate
  // numeric seed, because setting a formula replaces that cell's cache.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=0.5*(A1+2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_NEAR(v.u.number, 2.0, 1e-3);
}

TEST(FormulonCApi, ThreadLocalLastErrorIsolation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Establish a known last-error on the main thread.
  fm_value_t v{};
  ASSERT_NE(fm_workbook_get_value(nullptr, 0, 0, 0, &v), 0);
  const std::string main_msg = fm_last_error_message();
  EXPECT_GT(main_msg.size(), 0U);

  // A worker thread triggers a *different* error path; the main
  // thread's last error must remain untouched.
  std::atomic<fm_status_t> worker_rc{0};
  std::thread worker([&]() {
    // Out-of-range sheet on a freshly created workbook on this thread.
    fm_workbook_t* local = nullptr;
    if (fm_workbook_create(&local) != 0) {
      return;
    }
    worker_rc.store(fm_workbook_set_number(local, 99, 0, 0, 1.0));
    fm_workbook_destroy(local);
  });
  worker.join();
  EXPECT_NE(worker_rc.load(), 0);

  // Main thread's diagnostic survived the worker's run.
  EXPECT_EQ(std::string(fm_last_error_message()), main_msg);
}

TEST(FormulonCApi, StatusStringCoversKnownCodes) {
  // Spot-check a handful of band-spanning codes; the source-of-truth is
  // `formulon::to_cstring`, which the C API forwards to.
  EXPECT_STREQ(fm_status_string(0), "kOk");
  EXPECT_STREQ(fm_status_string(2), "kInvalidArgument");
  EXPECT_STREQ(fm_status_string(7000), "kBindingInvalidHandle");
  EXPECT_STREQ(fm_status_string(7001), "kBindingNullPointer");
  // Unknown numeric values must still return a non-NULL fallback.
  const char* unknown = fm_status_string(123456);
  ASSERT_NE(unknown, nullptr);
  EXPECT_GT(std::strlen(unknown), 0U);
}

TEST(FormulonCApi, ErrorDisplayNameUsesExcelLiterals) {
  EXPECT_STREQ(fm_error_display_name(0), "#NULL!");
  EXPECT_STREQ(fm_error_display_name(1), "#DIV/0!");
  EXPECT_STREQ(fm_error_display_name(3), "#REF!");
  EXPECT_STREQ(fm_error_display_name(-1), "#UNKNOWN!");
  EXPECT_STREQ(fm_error_display_name(999), "#UNKNOWN!");
}

TEST(FormulonCApi, VersionStringNonEmpty) {
  const char* v = fm_version_string();
  ASSERT_NE(v, nullptr);
  EXPECT_GT(std::strlen(v), 0U);
}
