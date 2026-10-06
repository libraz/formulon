//
// Stable C ABI tests for cell display text and ad-hoc value formatting.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "value.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

fm_value_t Number(double v) {
  fm_value_t value{};
  value.kind = FM_VAL_NUMBER;
  value.u.number = v;
  return value;
}

/// Renders through `fm_workbook_format_value` and returns `(text, status)`.
std::pair<std::string, int32_t> Format(fm_workbook_t* wb, const fm_value_t& value, const char* code) {
  const char* text = nullptr;
  int32_t status = -1;
  EXPECT_EQ(fm_workbook_format_value(wb, &value, code, &text, &status), 0);
  return {text == nullptr ? std::string() : std::string(text), status};
}

using Bytes = std::vector<std::uint8_t>;

Bytes SavePackage(fm_workbook_t* wb) {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  EXPECT_EQ(fm_workbook_save(wb, &data, &len), 0);
  Bytes result;
  if (data != nullptr) {
    result.assign(data, data + len);
  }
  fm_buffer_free(data);
  return result;
}

void LoadPackage(const Bytes& bytes, WorkbookGuard* out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), 0);
}

/// Rewrites a package entry while preserving every unrelated OOXML part.
/// This is the same fixture shape used by the effective-style C API tests.
Bytes ReplaceEntries(const Bytes& package, const std::map<std::string, std::string>& replacements) {
  formulon::io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(formulon::io::ByteSpan{package.data(), package.size()})));
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  std::size_t replaced = 0;
  for (const std::string& name : input.list_entries()) {
    auto body = input.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(body)) << name;
    if (!body) {
      continue;
    }
    Bytes content = body.value();
    if (auto it = replacements.find(name); it != replacements.end()) {
      content.assign(it->second.begin(), it->second.end());
      ++replaced;
    }
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), content.data(), content.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE);
  }
  EXPECT_EQ(replaced, replacements.size());
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  Bytes result(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return result;
}

Bytes PackageWithSheetXml(fm_workbook_t* wb, std::string sheet_xml) {
  return ReplaceEntries(SavePackage(wb), {{"xl/worksheets/sheet1.xml", std::move(sheet_xml)}});
}

Bytes PackageWithColumnPercentSpill(fm_workbook_t* wb, uint32_t column_xf) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<cols><col min=\"2\" max=\"2\" style=\"" +
      std::to_string(column_xf) +
      "\"/></cols><sheetData><row r=\"1\"><c r=\"A1\" cm=\"1\"><f>SEQUENCE(1,2)</f><v>1</v></c>"
      "</row></sheetData></worksheet>";
  return PackageWithSheetXml(wb, sheet);
}

Bytes SaveXlsbPackage(fm_workbook_t* wb) {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  EXPECT_EQ(fm_workbook_save_as(wb, FM_WORKBOOK_FORMAT_XLSB, &data, &len), 0);
  Bytes result;
  if (data != nullptr) {
    result.assign(data, data + len);
  }
  fm_buffer_free(data);
  return result;
}

Bytes PackageWithNonGeneralDefaultXf(fm_workbook_t* wb, std::string sheet_xml) {
  const Bytes package = SavePackage(wb);
  formulon::io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(formulon::io::ByteSpan{package.data(), package.size()})));
  auto styles_entry = input.read_entry("xl/styles.xml");
  EXPECT_TRUE(static_cast<bool>(styles_entry));
  if (!styles_entry) {
    return {};
  }
  std::string styles(styles_entry.value().begin(), styles_entry.value().end());
  const std::size_t cell_xfs_begin = styles.find("<cellXfs");
  EXPECT_NE(cell_xfs_begin, std::string::npos);
  const std::size_t first_xf = styles.find("<xf", cell_xfs_begin);
  EXPECT_NE(first_xf, std::string::npos);
  const std::size_t first_xf_end = styles.find('>', first_xf);
  EXPECT_NE(first_xf_end, std::string::npos);
  const std::size_t general_id = styles.find("numFmtId=\"0\"", first_xf);
  EXPECT_NE(general_id, std::string::npos);
  if (cell_xfs_begin == std::string::npos || first_xf == std::string::npos || first_xf_end == std::string::npos ||
      general_id == std::string::npos || general_id > first_xf_end) {
    return {};
  }
  styles.replace(general_id, std::string("numFmtId=\"0\"").size(), "numFmtId=\"2\"");
  return ReplaceEntries(package,
                        {{"xl/styles.xml", std::move(styles)}, {"xl/worksheets/sheet1.xml", std::move(sheet_xml)}});
}

}  // namespace

TEST(FormulonCApiDisplayText, CellUsesItsNumberFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  xf.num_fmt_id = 2;  // built-in 0.00
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_index), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_index), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 1, 0, "plain"), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 0, &text, &status), 0);
  EXPECT_STREQ(text, "1.50");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 0, &text, &status), 0);
  EXPECT_STREQ(text, "plain");
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 5, 5, &text, &status), 0);
  EXPECT_STREQ(text, "");
  EXPECT_EQ(status, FM_DISPLAY_OK);
}

TEST(FormulonCApiDisplayText, SpillCellsDisplayResolvedValuesWithTheirOwnFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_EQ(value.u.number, 4.0);
  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4");
  EXPECT_EQ(status, FM_DISPLAY_OK);

  fm_cell_xf xf{};
  xf.num_fmt_id = 2;  // 0.00
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_index), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1, 1, xf_index), 0);
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4.00");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 1, &value), 0);
  EXPECT_EQ(value.u.number, 4.0);

  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "");
}

TEST(FormulonCApiDisplayText, SpillPhantomsUseColumnRowAndCellStylePrecedence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  fm_cell_xf decimal{};
  decimal.num_fmt_id = 2;  // 0.00
  uint32_t row_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, decimal, &row_xf), 0);
  fm_cell_xf explicit_cell{};
  explicit_cell.num_fmt_id = 9;  // 0%
  uint32_t cell_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, explicit_cell, &cell_xf), 0);

  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,2)"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 2, 0.125), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1, 2, cell_xf), 0);
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<cols><col min=\"2\" max=\"2\" style=\"" +
      std::to_string(column_xf) +
      "\"/></cols><sheetData>"
      "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f>SEQUENCE(2,2)</f><v>1</v></c></row>"
      "<row r=\"2\" s=\"" +
      std::to_string(row_xf) + "\" customFormat=\"1\"><c r=\"C2\" s=\"" + std::to_string(cell_xf) +
      "\"><v>0.125</v></c></row>"
      "</sheetData></worksheet>";

  WorkbookGuard loaded;
  LoadPackage(PackageWithSheetXml(wb.handle, sheet), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4.00");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  // C2 carries an explicit xf and therefore beats row 2's inherited decimal
  // format, even though the row and the cell both sit outside the spill.
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 1, 2, &text, &status), 0);
  EXPECT_STREQ(text, "13%");
  EXPECT_EQ(status, FM_DISPLAY_OK);

  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 2);  // column
  EXPECT_STREQ(style.num_fmt_code, "0.00%");
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 1);  // customFormat row
  EXPECT_STREQ(style.num_fmt_code, "0.00");
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 2, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit cell
  EXPECT_STREQ(style.num_fmt_code, "0%");
}

TEST(FormulonCApiDisplayText, SpillPhantomWithoutLayoutUsesDefaultXfZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // An unassigned non-default XF must not leak.
  uint32_t percent_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &percent_xf), 0);
  ASSERT_NE(percent_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 3);  // default xf 0
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, SpillPhantomUsesNonGeneralDefaultXfZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,2)"), 0);
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData><row r=\"1\"><c r=\"A1\" cm=\"1\"><f>SEQUENCE(2,2)</f><v>1</v></c></row></sheetData>"
      "</worksheet>";
  WorkbookGuard loaded;
  LoadPackage(PackageWithNonGeneralDefaultXf(wb.handle, sheet), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4.00");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 3);  // the non-General default xf 0
  EXPECT_STREQ(style.num_fmt_code, "0.00");
}

TEST(FormulonCApiDisplayText, InvalidRowXfFallsBackToDefaultRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,2)"), 0);
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetData><row r=\"1\"><c r=\"A1\" cm=\"1\"><f>SEQUENCE(2,2)</f><v>1</v></c></row>"
      "<row r=\"2\" s=\"9999\" customFormat=\"1\"/></sheetData>"
      "</worksheet>";
  WorkbookGuard loaded;
  LoadPackage(PackageWithNonGeneralDefaultXf(wb.handle, sheet), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 1, 1, &text, &status), 0);
  EXPECT_STREQ(text, "4.00");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 1);     // the row override still supplies precedence
  EXPECT_EQ(style.xf_index, 0U);  // its raw index is out of range
  EXPECT_STREQ(style.num_fmt_code, "0.00");
}

TEST(FormulonCApiDisplayText, MaterializedBlankGapDoesNotShadowColumnStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=B1"), 0);
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<cols><col min=\"2\" max=\"2\" style=\"" +
      std::to_string(column_xf) +
      "\"/></cols><sheetData>"
      "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f>SEQUENCE(1,2)</f><v>1</v></c>"
      "<c r=\"C1\"><f>B1</f><v>2</v></c></row>"
      "</sheetData></worksheet>";
  WorkbookGuard loaded;
  LoadPackage(PackageWithSheetXml(wb.handle, sheet), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 2, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 2);  // column, despite the materialized blank gap
  EXPECT_STREQ(style.num_fmt_code, "0.00%");
}

TEST(FormulonCApiDisplayText, ExplicitDefaultCellXfOverridesInheritedColumnFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  EXPECT_EQ(status, FM_DISPLAY_OK);

  ASSERT_EQ(fm_cell_set_xf_index(loaded.handle, 0, 0, 1, 0), 0);
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit cell, even though its xf is the default record
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");

  const Bytes saved = SavePackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  ASSERT_EQ(fm_workbook_get_display_text(reloaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, ExplicitDefaultRangeXfOverridesInheritedColumnFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  ASSERT_EQ(fm_sheet_set_range_xf_index(loaded.handle, 0, 0, 1, 0, 1, 0), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit range cell
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, XlsbPreservesNonDefaultXfOnBlankCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  fm_cell_xf decimal{};
  decimal.num_fmt_id = 2;  // 0.00
  uint32_t cell_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, decimal, &cell_xf), 0);
  ASSERT_NE(cell_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  // B2 is outside the one-row spill, so this blank cell exercises ordinary
  // cell serialization while inheriting the column format when unstyled.
  ASSERT_EQ(fm_cell_set_xf_index(loaded.handle, 0, 1, 1, cell_xf), 0);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, cell_xf);
  EXPECT_STREQ(style.num_fmt_code, "0.00");

  const Bytes saved = SaveXlsbPackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, cell_xf);
  EXPECT_STREQ(style.num_fmt_code, "0.00");
}

TEST(FormulonCApiDisplayText, XlsbPreservesExplicitDefaultXfOnBlankCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  // An explicit default XF must block inherited column formatting even though
  // B2 remains a blank cell outside the spill footprint.
  ASSERT_EQ(fm_cell_set_xf_index(loaded.handle, 0, 1, 1, 0), 0);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");

  const Bytes saved = SaveXlsbPackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 1, 1, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, XlsbKeepsInheritedFormatOnUnstyledSpillPhantom) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);

  const Bytes saved = SaveXlsbPackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  ASSERT_EQ(fm_workbook_get_display_text(reloaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  ASSERT_EQ(fm_workbook_get_value(reloaded.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // serialized effective XF on the phantom cell
  EXPECT_EQ(style.xf_index, column_xf);
  EXPECT_STREQ(style.num_fmt_code, "0.00%");
}

TEST(FormulonCApiDisplayText, XlsbKeepsExplicitDefaultXfOnSpillPhantom) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  ASSERT_EQ(fm_cell_set_xf_index(loaded.handle, 0, 0, 1, 0), 0);
  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(loaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(loaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit cell XF0
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");

  const Bytes saved = SaveXlsbPackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  ASSERT_EQ(fm_workbook_get_display_text(reloaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  ASSERT_EQ(fm_workbook_get_value(reloaded.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit cell XF0 survived the phantom record
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, XlsxKeepsInheritedFormatOnUnstyledSpillPhantom) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);

  const Bytes saved = SavePackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(reloaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "200.00%");
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // serialized effective XF on the phantom cell
  EXPECT_EQ(style.xf_index, column_xf);
  EXPECT_STREQ(style.num_fmt_code, "0.00%");
}

TEST(FormulonCApiDisplayText, XlsxKeepsExplicitDefaultXfOnSpillPhantom) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf percent{};
  percent.num_fmt_id = 10;  // 0.00%
  uint32_t column_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, percent, &column_xf), 0);
  ASSERT_NE(column_xf, 0U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)"), 0);

  WorkbookGuard loaded;
  LoadPackage(PackageWithColumnPercentSpill(wb.handle, column_xf), &loaded);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  ASSERT_EQ(fm_cell_set_xf_index(loaded.handle, 0, 0, 1, 0), 0);

  const Bytes saved = SavePackage(loaded.handle);
  ASSERT_FALSE(saved.empty());
  WorkbookGuard reloaded;
  LoadPackage(saved, &reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(reloaded.handle, 0, 0, 1, &text, &status), 0);
  EXPECT_STREQ(text, "2");
  fm_effective_style style{};
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 1, &style), 0);
  EXPECT_EQ(style.source, 0);  // explicit cell XF0 survived the phantom cell
  EXPECT_EQ(style.xf_index, 0U);
  EXPECT_STREQ(style.num_fmt_code, "General");
}

TEST(FormulonCApiDisplayText, ReferencesPreserveValuesAndUseDestinationFormats) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  xf.num_fmt_id = 2;  // 0.00
  uint32_t decimal_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &decimal_xf), 0);
  xf.num_fmt_id = 10;  // 0.00%
  uint32_t percent_xf = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &percent_xf), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 0.125), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, decimal_xf), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 1, 0, "0012.345"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 3, 0, "=1/0"), 0);
  for (uint32_t row = 0; row < 4; ++row) {
    const std::string formula = "=A" + std::to_string(row + 1);
    ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, row, 1, formula.c_str()), 0);
    ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, row, 1, percent_xf), 0);
  }
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=A1*100"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 0, &text, &status), 0);
  EXPECT_STREQ(text, "0.13");
  const char* expected[] = {"12.50%", "0012.345", "0.00%", "#DIV/0!"};
  for (uint32_t row = 0; row < 4; ++row) {
    ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, row, 1, &text, &status), 0);
    EXPECT_STREQ(text, expected[row]);
    EXPECT_EQ(status, FM_DISPLAY_OK);
  }
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_EQ(value.u.number, 0.125);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_TEXT);
  EXPECT_STREQ(value.u.text, "0012.345");
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_EQ(value.u.number, 0.0);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 3, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_ERROR);
  EXPECT_EQ(value.u.error_code, static_cast<int32_t>(formulon::ErrorCode::Div0));
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 2, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_EQ(value.u.number, 12.5);
}

TEST(FormulonCApiDisplayText, OutOfCalendarDateOverflows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const auto rendered = Format(wb.handle, Number(-1.0), "yyyy/m/d");
  EXPECT_EQ(rendered.first, "########");
  EXPECT_EQ(rendered.second, FM_DISPLAY_OVERFLOW);
}

TEST(FormulonCApiDisplayText, FormatValueCoversEveryScalarKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(Format(wb.handle, Number(1234.5), "#,##0").first, "1,235");
  EXPECT_EQ(Format(wb.handle, Number(0.25), "").first, "0.25");

  fm_value_t boolean{};
  boolean.kind = FM_VAL_BOOL;
  boolean.u.boolean = 1;
  EXPECT_EQ(Format(wb.handle, boolean, "General").first, "TRUE");

  fm_value_t text{};
  text.kind = FM_VAL_TEXT;
  text.u.text = "abc";
  EXPECT_EQ(Format(wb.handle, text, "General").first, "abc");

  fm_value_t error{};
  error.kind = FM_VAL_ERROR;
  error.u.error_code = 1;  // #DIV/0!
  EXPECT_EQ(Format(wb.handle, error, "0.00").first, "#DIV/0!");

  fm_value_t blank{};
  blank.kind = FM_VAL_BLANK;
  EXPECT_EQ(Format(wb.handle, blank, "0.00").first, "");
}

TEST(FormulonCApiDisplayText, FormatValueFollowsWorkbookDateSystem) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(Format(wb.handle, Number(1.0), "yyyy-mm-dd").first, "1900-01-01");
  wb.handle->workbook().set_date1904(true);
  EXPECT_EQ(Format(wb.handle, Number(1.0), "yyyy-mm-dd").first, "1904-01-02");
}

TEST(FormulonCApiDisplayText, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* text = nullptr;
  int32_t status = 0;
  EXPECT_EQ(fm_workbook_get_display_text(wb.handle, 3, 0, 0, &text, &status), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 16384U, &text, &status), kInvalidArgument);
  EXPECT_NE(fm_workbook_get_display_text(wb.handle, 0, 0, 0, nullptr, &status), 0);

  fm_value_t value = Number(std::numeric_limits<double>::quiet_NaN());
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_ARRAY;
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_ERROR;
  value.u.error_code = 9999;
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_TEXT;
  value.u.text = nullptr;
  EXPECT_NE(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), 0);
  value = Number(1.0);
  EXPECT_NE(fm_workbook_format_value(wb.handle, &value, nullptr, &text, &status), 0);
  EXPECT_NE(fm_workbook_format_value(nullptr, &value, "0", &text, &status), 0);
}
