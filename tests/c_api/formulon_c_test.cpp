//
// Stable C ABI (`src/c_api/formulon_c.h`) end-to-end tests.
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface declared in
// `formulon_c.h`. The same surface drives the CLI binary, the WASM
// embind layer, and any external language binding, so this file is the
// regression harness for the contract.

#include "c_api/formulon_c.h"

#include <atomic>
#include <cstdint>
#include <cstring>
#include <string>
#include <thread>
#include <vector>

#include "gtest/gtest.h"
#include "io/format_detect.h"
#include "io/xlsb/writer.h"
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

TEST(FormulonCApi, SaveExRejectsUnknownFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  BufferGuard buf;
  const fm_status_t rc = fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_UNKNOWN, &buf.data, &buf.len);
  EXPECT_NE(rc, 0);
  EXPECT_EQ(buf.data, nullptr);
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

TEST(FormulonCApi, VersionStringNonEmpty) {
  const char* v = fm_version_string();
  ASSERT_NE(v, nullptr);
  EXPECT_GT(std::strlen(v), 0U);
}
