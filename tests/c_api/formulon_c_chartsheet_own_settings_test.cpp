// Stable C ABI: the settings a chartsheet's part states itself (protection, margins, page setup, header and
// footer) are read, set and written back. The fixtures are Excel-saved twins of the worksheet `Sheet1` then a
// chartsheet, each in both containers.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "utils/error.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr std::size_t kChartsheet = 1;
constexpr const char* kExts[] = {"xlsx", "xlsb"};

std::vector<std::uint8_t> Fixture(const std::string& name, const std::string& ext) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/chartsheet_twins/" + name + "." + ext;
  std::ifstream in(path, std::ios::binary);
  EXPECT_TRUE(in.good()) << path;
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

void Load(const std::string& name, const std::string& ext, WorkbookGuard* out) {
  const std::vector<std::uint8_t> bytes = Fixture(name, ext);
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &out->handle), 0) << fm_get_last_error_message();
}

std::vector<std::uint8_t> Save(const WorkbookGuard& wb, const std::string& ext) {
  BufferGuard buffer;
  const int32_t format = ext == "xlsx" ? FM_WORKBOOK_FORMAT_XLSX : FM_WORKBOOK_FORMAT_XLSB;
  EXPECT_EQ(fm_workbook_save(wb.handle, format, &buffer.data, &buffer.len), 0) << fm_get_last_error_message();
  return std::vector<std::uint8_t>(buffer.data, buffer.data + buffer.len);
}

// The chartsheet part of a package in its own container.
std::vector<std::uint8_t> ChartsheetPart(const std::vector<std::uint8_t>& package, const std::string& ext) {
  formulon::io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{package.data(), package.size()})));
  auto entry = zip.read_entry(ext == "xlsx" ? "xl/chartsheets/sheet1.xml" : "xl/chartsheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(entry));
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

// The part the book saves, in the container it was loaded from.
std::vector<std::uint8_t> SavedPart(const WorkbookGuard& wb, const std::string& ext) {
  return ChartsheetPart(Save(wb, ext), ext);
}

std::vector<std::uint8_t> TwinPart(const std::string& name, const std::string& ext) {
  return ChartsheetPart(Fixture(name, ext), ext);
}

fm_chartsheet_protection_t ProtectionRecord() {
  return FM_STRUCT_INIT(fm_chartsheet_protection_t);
}

struct Protection {
  int32_t present = 0;
  std::string algorithm;
  std::string hash;
  std::string salt;
  std::uint32_t spin = 0;
  std::string legacy;
  int32_t content = 0;
  int32_t objects = 0;
  bool operator==(const Protection& o) const {
    return present == o.present && algorithm == o.algorithm && hash == o.hash && salt == o.salt && spin == o.spin &&
           legacy == o.legacy && content == o.content && objects == o.objects;
  }
};

Protection ReadProtection(const WorkbookGuard& wb) {
  Protection out;
  fm_chartsheet_protection_t record = ProtectionRecord();
  EXPECT_EQ(fm_sheet_get_chartsheet_protection(wb.handle, kChartsheet, &out.present, &record), 0);
  const auto text = [](const char* value) { return value == nullptr ? std::string() : std::string(value); };
  out.algorithm = text(record.algorithm_name);
  out.hash = text(record.hash_value);
  out.salt = text(record.salt_value);
  out.spin = record.spin_count;
  out.legacy = text(record.legacy_password_hash);
  out.content = record.allow_edit_content;
  out.objects = record.allow_edit_objects;
  return out;
}

// Applies the record the getter reads back to the setter.
fm_status_t SetFromRead(const WorkbookGuard& wb) {
  fm_chartsheet_protection_t record = ProtectionRecord();
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_chartsheet_protection(wb.handle, kChartsheet, &present, &record), 0);
  return fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record);
}

void ExpectRefusal(fm_status_t rc, const char* message, const char* context) {
  EXPECT_EQ(rc, kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(), message);
  EXPECT_STREQ(fm_get_last_error_context(), context);
}

}  // namespace

TEST(FormulonCApiChartsheetOwnSettings, ProtectionSetTakesBackWhatGetReads) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    for (const char* name : {"cs_protect_hash", "cs_protect_content_only"}) {
      SCOPED_TRACE(name);
      WorkbookGuard wb;
      Load(name, ext, &wb);
      ASSERT_NE(wb.handle, nullptr);
      const Protection before = ReadProtection(wb);
      ASSERT_EQ(before.present, 1);
      EXPECT_EQ(SetFromRead(wb), 0) << fm_get_last_error_message();
      EXPECT_EQ(ReadProtection(wb), before);
      // The saved part is the twin Excel wrote, unchanged.
      EXPECT_EQ(SavedPart(wb, ext), TwinPart(name, ext));
    }
  }
}

TEST(FormulonCApiChartsheetOwnSettings, ProtectionSetWritesWhatExcelWrites) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard wb;
    Load("cs_control", ext, &wb);
    ASSERT_NE(wb.handle, nullptr);
    fm_chartsheet_protection_t record = ProtectionRecord();
    record.allow_edit_content = 0;
    record.allow_edit_objects = 1;
    ASSERT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record), 0) << fm_get_last_error_message();
    const Protection read = ReadProtection(wb);
    EXPECT_EQ(read.present, 1);
    EXPECT_EQ(read.content, 0);
    EXPECT_EQ(read.objects, 1);
    EXPECT_EQ(SavedPart(wb, ext), TwinPart("cs_protect_content_only", ext));
  }
}

TEST(FormulonCApiChartsheetOwnSettings, ProtectionIsNeverRemoved) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard wb;
    Load("cs_control", ext, &wb);
    ASSERT_NE(wb.handle, nullptr);
    fm_chartsheet_protection_t none = ProtectionRecord();
    none.allow_edit_content = 1;
    none.allow_edit_objects = 1;

    // Nothing to remove on an unprotected chartsheet.
    EXPECT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, nullptr), 0);
    EXPECT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &none), 0);
    EXPECT_EQ(ReadProtection(wb).present, 0);

    fm_chartsheet_protection_t locked = ProtectionRecord();
    ASSERT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &locked), 0);
    const Protection before = ReadProtection(wb);
    ASSERT_EQ(before.present, 1);
    ExpectRefusal(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, nullptr),
                  "Chartsheet protection cannot be removed", "sheet_index=1");
    ExpectRefusal(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &none),
                  "Chartsheet protection cannot be removed", "sheet_index=1");
    EXPECT_EQ(ReadProtection(wb), before);
  }
}

TEST(FormulonCApiChartsheetOwnSettings, ProtectionArgumentsAreChecked) {
  WorkbookGuard wb;
  Load("cs_control", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  fm_chartsheet_protection_t record = ProtectionRecord();

  record.hash_value = "";
  ExpectRefusal(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record), "hash_value must not be empty",
                "field=hash_value");
  record.hash_value = nullptr;
  record.salt_value = "";
  ExpectRefusal(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record), "salt_value must not be empty",
                "field=salt_value");
  record.salt_value = nullptr;
  record.allow_edit_content = 2;
  EXPECT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record), kInvalidArgument);
  record.allow_edit_content = 0;
  record.struct_size = 1;
  EXPECT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, kChartsheet, &record), kInvalidArgument);
  EXPECT_EQ(ReadProtection(wb).present, 0);

  EXPECT_EQ(fm_sheet_set_chartsheet_protection(nullptr, kChartsheet, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_set_chartsheet_protection(wb.handle, 99U, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  // A worksheet has no chartsheet protection to set.
  ExpectRefusal(fm_sheet_set_chartsheet_protection(wb.handle, 0, nullptr),
                "sheet_index is a worksheet, but this call needs a chartsheet", "sheet_index=0 kind=worksheet");
}

TEST(FormulonCApiChartsheetOwnSettings, PaperSizeAloneWritesPortraitOrientation) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard wb;
    Load("cs_control", ext, &wb);
    ASSERT_NE(wb.handle, nullptr);
    fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
    patch.paper_size_op = FM_PATCH_OP_SET;
    patch.paper_size = 5;
    ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch), 0) << fm_get_last_error_message();

    fm_page_setup_t read = FM_STRUCT_INIT(fm_page_setup_t);
    ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, kChartsheet, &read), 0);
    EXPECT_EQ(read.paper_size, 5U);
    EXPECT_EQ(read.has_paper_size, 1);

    // BrtCsPageSetup in .xlsb, `<pageSetup paperSize="5" orientation="portrait"/>` in .xlsx.
    EXPECT_EQ(SavedPart(wb, ext), TwinPart("cs_page_setup_legal", ext));
  }
}

TEST(FormulonCApiChartsheetOwnSettings, PaperSizeAndOrientationApplyByPatchAndByXml) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard patched;
    Load("cs_control", ext, &patched);
    ASSERT_NE(patched.handle, nullptr);
    fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
    patch.paper_size_op = FM_PATCH_OP_SET;
    patch.paper_size = 9;
    patch.orientation_op = FM_PATCH_OP_SET;
    patch.orientation = FM_PAGE_ORIENTATION_LANDSCAPE;
    ASSERT_EQ(fm_sheet_update_page_setup(patched.handle, kChartsheet, &patch), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(patched, ext), TwinPart("cs_page_setup_landscape_a4", ext));

    WorkbookGuard by_xml;
    Load("cs_control", ext, &by_xml);
    ASSERT_NE(by_xml.handle, nullptr);
    ASSERT_EQ(fm_sheet_set_page_setup_xml(by_xml.handle, kChartsheet,
                                          "<pageSetup orientation=\"landscape\" paperSize=\"9\"/>"),
              0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(by_xml, ext), TwinPart("cs_page_setup_landscape_a4", ext));

    // Removing the element returns the part to the control.
    ASSERT_EQ(fm_sheet_set_page_setup_xml(by_xml.handle, kChartsheet, nullptr), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(by_xml, ext), TwinPart("cs_control", ext));
  }
}

TEST(FormulonCApiChartsheetOwnSettings, MarginsAndHeaderFooterApply) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard margins;
    Load("cs_control", ext, &margins);
    ASSERT_NE(margins.handle, nullptr);
    fm_page_margins_patch_t patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
    patch.left_in_op = patch.right_in_op = patch.top_in_op = FM_PATCH_OP_SET;
    patch.bottom_in_op = patch.header_in_op = patch.footer_in_op = FM_PATCH_OP_SET;
    patch.left_in = 0.5;
    patch.right_in = 0.75;
    patch.top_in = 1.0;
    patch.bottom_in = 1.25;
    patch.header_in = 0.25;
    patch.footer_in = 0.375;
    ASSERT_EQ(fm_sheet_update_page_margins(margins.handle, kChartsheet, &patch), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(margins, ext), TwinPart("cs_margins", ext));

    WorkbookGuard header;
    Load("cs_control", ext, &header);
    ASSERT_NE(header.handle, nullptr);
    fm_header_footer_patch_t hf = FM_STRUCT_INIT(fm_header_footer_patch_t);
    hf.odd_header_op = FM_PATCH_OP_SET;
    hf.odd_header = "&RR";
    hf.odd_footer_op = FM_PATCH_OP_SET;
    hf.odd_footer = "&LL&P&CC";
    ASSERT_EQ(fm_sheet_update_header_footer(header.handle, kChartsheet, &hf), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(header, ext), TwinPart("cs_header_footer", ext));
  }
}

// Rule 26: each record read from a chartsheet goes back through its setter and changes nothing.
TEST(FormulonCApiChartsheetOwnSettings, EveryPrintRecordSetsBackWhatItReads) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    WorkbookGuard wb;
    Load("cs_header_footer_different_first", ext, &wb);
    ASSERT_NE(wb.handle, nullptr);
    WorkbookGuard setup_book;
    Load("cs_page_setup_landscape_a4", ext, &setup_book);
    ASSERT_NE(setup_book.handle, nullptr);
    WorkbookGuard margins_book;
    Load("cs_margins", ext, &margins_book);
    ASSERT_NE(margins_book.handle, nullptr);

    fm_header_footer_t hf = FM_STRUCT_INIT(fm_header_footer_t);
    ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, kChartsheet, &hf), 0);
    fm_header_footer_patch_t hf_patch = FM_STRUCT_INIT(fm_header_footer_patch_t);
    // Only what the record states goes back, as a binding passes it.
    const auto section = [](const char* text, fm_patch_op_t* op, const char** value) {
      if (text != nullptr) {
        *op = FM_PATCH_OP_SET;
        *value = text;
      }
    };
    section(hf.odd_header, &hf_patch.odd_header_op, &hf_patch.odd_header);
    section(hf.odd_footer, &hf_patch.odd_footer_op, &hf_patch.odd_footer);
    section(hf.even_header, &hf_patch.even_header_op, &hf_patch.even_header);
    section(hf.even_footer, &hf_patch.even_footer_op, &hf_patch.even_footer);
    section(hf.first_header, &hf_patch.first_header_op, &hf_patch.first_header);
    section(hf.first_footer, &hf_patch.first_footer_op, &hf_patch.first_footer);
    const auto flag = [](int32_t has, int32_t value, fm_patch_op_t* op, int32_t* out) {
      if (has != 0) {
        *op = FM_PATCH_OP_SET;
        *out = value;
      }
    };
    flag(hf.has_different_odd_even, hf.different_odd_even, &hf_patch.different_odd_even_op,
         &hf_patch.different_odd_even);
    flag(hf.has_different_first, hf.different_first, &hf_patch.different_first_op, &hf_patch.different_first);
    flag(hf.has_scale_with_doc, hf.scale_with_doc, &hf_patch.scale_with_doc_op, &hf_patch.scale_with_doc);
    flag(hf.has_align_with_margins, hf.align_with_margins, &hf_patch.align_with_margins_op,
         &hf_patch.align_with_margins);
    ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, kChartsheet, &hf_patch), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(wb, ext), TwinPart("cs_header_footer_different_first", ext));

    fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
    ASSERT_EQ(fm_sheet_get_page_setup(setup_book.handle, kChartsheet, &setup), 0);
    fm_page_setup_patch_t setup_patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
    setup_patch.orientation_op = setup_patch.paper_size_op = setup_patch.scale_pct_op = FM_PATCH_OP_SET;
    setup_patch.fit_to_width_op = setup_patch.fit_to_height_op = setup_patch.fit_to_page_op = FM_PATCH_OP_SET;
    setup_patch.orientation = setup.orientation;
    setup_patch.paper_size = setup.paper_size;
    setup_patch.scale_pct = setup.scale_pct;
    setup_patch.fit_to_width = setup.fit_to_width;
    setup_patch.fit_to_height = setup.fit_to_height;
    setup_patch.fit_to_page = setup.fit_to_page;
    ASSERT_EQ(fm_sheet_update_page_setup(setup_book.handle, kChartsheet, &setup_patch), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(setup_book, ext), TwinPart("cs_page_setup_landscape_a4", ext));

    const char* xml = nullptr;
    ASSERT_EQ(fm_sheet_get_page_setup_xml(setup_book.handle, kChartsheet, &xml), 0);
    const std::string setup_xml = xml != nullptr ? xml : "";
    EXPECT_EQ(fm_sheet_set_page_setup_xml(setup_book.handle, kChartsheet, setup_xml.c_str()), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(setup_book, ext), TwinPart("cs_page_setup_landscape_a4", ext));

    fm_page_margins_t read = FM_STRUCT_INIT(fm_page_margins_t);
    ASSERT_EQ(fm_sheet_get_page_margins(margins_book.handle, kChartsheet, &read), 0);
    fm_page_margins_patch_t margins_patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
    margins_patch.left_in_op = margins_patch.right_in_op = margins_patch.top_in_op = FM_PATCH_OP_SET;
    margins_patch.bottom_in_op = margins_patch.header_in_op = margins_patch.footer_in_op = FM_PATCH_OP_SET;
    margins_patch.left_in = read.left_in;
    margins_patch.right_in = read.right_in;
    margins_patch.top_in = read.top_in;
    margins_patch.bottom_in = read.bottom_in;
    margins_patch.header_in = read.header_in;
    margins_patch.footer_in = read.footer_in;
    ASSERT_EQ(fm_sheet_update_page_margins(margins_book.handle, kChartsheet, &margins_patch), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(margins_book, ext), TwinPart("cs_margins", ext));
    ASSERT_EQ(fm_sheet_get_page_margins_xml(margins_book.handle, kChartsheet, &xml), 0);
    const std::string margins_xml = xml != nullptr ? xml : "";
    EXPECT_EQ(fm_sheet_set_page_margins_xml(margins_book.handle, kChartsheet, margins_xml.c_str()), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(margins_book, ext), TwinPart("cs_margins", ext));
  }
}

TEST(FormulonCApiChartsheetOwnSettings, PageSetupRefusesWhatAChartsheetDoesNotStore) {
  WorkbookGuard wb;
  Load("cs_page_setup_landscape_a4", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);

  patch.scale_pct_op = FM_PATCH_OP_SET;
  patch.scale_pct = 50;
  ExpectRefusal(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch),
                "scale_pct states the scale, which a chartsheet does not store", "sheet_index=1 field=scale_pct");
  patch.scale_pct_op = FM_PATCH_OP_KEEP;
  patch.fit_to_width_op = FM_PATCH_OP_SET;
  patch.fit_to_width = 3;
  ExpectRefusal(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch),
                "fit_to_width states the fit to width, which a chartsheet does not store",
                "sheet_index=1 field=fit_to_width");
  patch.fit_to_width_op = FM_PATCH_OP_KEEP;
  patch.fit_to_height_op = FM_PATCH_OP_SET;
  patch.fit_to_height = 3;
  ExpectRefusal(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch),
                "fit_to_height states the fit to height, which a chartsheet does not store",
                "sheet_index=1 field=fit_to_height");
  patch.fit_to_height_op = FM_PATCH_OP_KEEP;
  patch.fit_to_page_op = FM_PATCH_OP_SET;
  patch.fit_to_page = 1;
  ExpectRefusal(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch),
                "fit_to_page states the fit to page, which a chartsheet does not store",
                "sheet_index=1 field=fit_to_page");

  // A refused patch changes nothing, paper size and orientation included.
  patch.paper_size_op = FM_PATCH_OP_SET;
  patch.paper_size = 5;
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch), kInvalidArgument);
  fm_page_setup_t read = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, kChartsheet, &read), 0);
  EXPECT_EQ(read.paper_size, 9U);

  const std::pair<const char*, const char*> unstored[] = {
      {"<pageSetup scale=\"50\"/>", "xml states the scale, which a chartsheet does not store"},
      {"<pageSetup fitToWidth=\"2\"/>", "xml states the fit to width, which a chartsheet does not store"},
      {"<pageSetup fitToHeight=\"2\"/>", "xml states the fit to height, which a chartsheet does not store"},
      {"<pageSetup pageOrder=\"overThenDown\"/>", "xml states the page order, which a chartsheet does not store"},
      {"<pageSetup cellComments=\"atEnd\"/>", "xml states the cell comments, which a chartsheet does not store"},
      {"<pageSetup errors=\"blank\"/>", "xml states the error display, which a chartsheet does not store"},
  };
  for (const auto& [xml, message] : unstored) {
    SCOPED_TRACE(xml);
    EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, kChartsheet, xml), kInvalidArgument);
    EXPECT_STREQ(fm_get_last_error_message(), message);
    EXPECT_EQ(std::strncmp(fm_get_last_error_context(), "sheet_index=1 field=", 20), 0);
  }
  // A value equal to the default restates it.
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, kChartsheet, "<pageSetup paperSize=\"9\" scale=\"100\"/>"), 0)
      << fm_get_last_error_message();
  patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  patch.scale_pct_op = FM_PATCH_OP_SET;
  patch.scale_pct = 100;
  EXPECT_EQ(fm_sheet_update_page_setup(wb.handle, kChartsheet, &patch), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiChartsheetOwnSettings, PageSetupKeepsWhatCannotBeChanged) {
  WorkbookGuard wb;
  Load("cs_control", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  const std::pair<const char*, const char*> fixed[] = {
      {"<pageSetup firstPageNumber=\"5\" useFirstPageNumber=\"1\"/>", "first page number"},
      {"<pageSetup copies=\"2\"/>", "copies"},
      {"<pageSetup horizontalDpi=\"300\"/>", "resolution"},
      {"<pageSetup draft=\"1\"/>", "draft"},
      {"<pageSetup blackAndWhite=\"1\"/>", "black and white"},
      {"<pageSetup usePrinterDefaults=\"0\"/>", "printer defaults"},
  };
  for (const auto& [xml, word] : fixed) {
    SCOPED_TRACE(xml);
    EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, kChartsheet, xml), kInvalidArgument);
    EXPECT_EQ(std::string(fm_get_last_error_message()),
              std::string("xml states the ") + word + ", which cannot be changed on a chartsheet");
    EXPECT_EQ(std::strncmp(fm_get_last_error_context(), "sheet_index=1 field=", 20), 0);
  }
  ExpectRefusal(fm_sheet_set_page_setup_xml(wb.handle, kChartsheet, "<pageSetup copies=\"2\"/>"),
                "xml states the copies, which cannot be changed on a chartsheet", "sheet_index=1 field=copies");
  // The current value, stated or defaulted, is accepted.
  EXPECT_EQ(fm_sheet_set_page_setup_xml(wb.handle, kChartsheet, "<pageSetup copies=\"1\" draft=\"0\"/>"), 0)
      << fm_get_last_error_message();
}

TEST(FormulonCApiChartsheetOwnSettings, PrintOptionsStayWorksheetOnly) {
  WorkbookGuard wb;
  Load("cs_control", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  fm_print_options_patch_t patch = FM_STRUCT_INIT(fm_print_options_patch_t);
  ExpectRefusal(fm_sheet_update_print_options(wb.handle, kChartsheet, &patch),
                "sheet_index is a chartsheet, but this call needs a worksheet", "sheet_index=1 kind=chartsheet");
}

TEST(FormulonCApiChartsheetOwnSettings, PageSetupWithoutPaperSizeWritesWhatExcelWrites) {
  for (const char* ext : kExts) {
    SCOPED_TRACE(ext);
    // Orientation lands after `firstPageNumber`, and no paper size is added.
    WorkbookGuard first_page;
    Load("cs_page_setup_first_page", ext, &first_page);
    ASSERT_NE(first_page.handle, nullptr);
    ASSERT_EQ(fm_sheet_set_page_setup_xml(first_page.handle, kChartsheet,
                                          "<pageSetup firstPageNumber=\"3\" useFirstPageNumber=\"1\"/>"),
              0)
        << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(first_page, ext), TwinPart("cs_page_setup_first_page", ext));

    WorkbookGuard landscape;
    Load("cs_control", ext, &landscape);
    ASSERT_NE(landscape.handle, nullptr);
    fm_page_setup_patch_t patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
    patch.orientation_op = FM_PATCH_OP_SET;
    patch.orientation = FM_PAGE_ORIENTATION_LANDSCAPE;
    ASSERT_EQ(fm_sheet_update_page_setup(landscape.handle, kChartsheet, &patch), 0) << fm_get_last_error_message();
    EXPECT_EQ(SavedPart(landscape, ext), TwinPart("cs_page_setup_orientation_only", ext));
  }
}

TEST(FormulonCApiChartsheetOwnSettings, SheetPrSetWritesWhatExcelWrites) {
  const std::pair<const char*, const char*> cases[] = {
      {"cs_code_name", "<sheetPr codeName=\"ChartA\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
      {"cs_unpublished", "<sheetPr published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
      {"cs_code_name_unpublished",
       "<sheetPr codeName=\"ChartA\" published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
  };
  for (const char* ext : kExts) {
    for (const auto& [twin, xml] : cases) {
      SCOPED_TRACE(std::string(twin) + "." + ext);
      WorkbookGuard wb;
      Load("cs_control", ext, &wb);
      ASSERT_NE(wb.handle, nullptr);
      ASSERT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet, xml), 0) << fm_get_last_error_message();
      EXPECT_EQ(SavedPart(wb, ext), TwinPart(twin, ext));

      // Restating the control's element takes both back.
      ASSERT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet, "<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>"), 0)
          << fm_get_last_error_message();
      EXPECT_EQ(SavedPart(wb, ext), TwinPart("cs_control", ext));
    }
  }
}

TEST(FormulonCApiChartsheetOwnSettings, SheetPrTabColorAppliesAsTheColorSetterDoes) {
  WorkbookGuard wb;
  Load("cs_code_name", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  ASSERT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet,
                                      "<sheetPr codeName=\"ChartA\"><tabColor theme=\"4\"/></sheetPr>"),
            0)
      << fm_get_last_error_message();
  fm_color_t color = fm_color_t{};
  ASSERT_EQ(fm_sheet_get_tab_color(wb.handle, kChartsheet, &color), 0);
  EXPECT_EQ(color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(color.theme, 4U);
  ASSERT_EQ(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet, "<sheetPr codeName=\"ChartA\"/>"), 0);
  ASSERT_EQ(fm_sheet_get_tab_color(wb.handle, kChartsheet, &color), 0);
  EXPECT_EQ(color.kind, FM_COLOR_KIND_NONE);
}

TEST(FormulonCApiChartsheetOwnSettings, SheetPrRefusesWhatAChartsheetDoesNotStore) {
  WorkbookGuard wb;
  Load("cs_code_name", "xlsx", &wb);
  ASSERT_NE(wb.handle, nullptr);
  ExpectRefusal(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet, "<sheetPr><outlinePr summaryBelow=\"0\"/></sheetPr>"),
                "xml states the outlinePr, which a chartsheet does not store", "sheet_index=1 field=outlinePr");
  ExpectRefusal(fm_sheet_set_sheet_pr_xml(wb.handle, kChartsheet, "<sheetPr codeName=\"ChartA\" filterMode=\"1\"/>"),
                "xml states the filterMode, which a chartsheet does not store", "sheet_index=1 field=filterMode");
  // A refusal leaves the part as it was.
  EXPECT_EQ(SavedPart(wb, "xlsx"), TwinPart("cs_code_name", "xlsx"));
}
