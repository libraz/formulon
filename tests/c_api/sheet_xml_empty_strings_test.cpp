//
// Empty-value rules of the raw sheet-XML setters, the cell phonetic setter and
// the header/footer section patch: an empty string is refused where it is not
// one element, and means "no value" where Excel stores one state for both.

#include <cstddef>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

using XmlGetter = fm_status_t (*)(const fm_workbook_t*, size_t, const char**);
using XmlSetter = fm_status_t (*)(fm_workbook_t*, size_t, const char*);

struct XmlSetterCase {
  const char* name;
  XmlGetter get;
  XmlSetter set;
  const char* fragment;
};

const XmlSetterCase kXmlSetters[] = {
    {"auto_filter", &fm_sheet_get_auto_filter_xml, &fm_sheet_set_auto_filter_xml, "<autoFilter ref=\"A1:C10\"/>"},
    {"page_setup", &fm_sheet_get_page_setup_xml, &fm_sheet_set_page_setup_xml,
     "<pageSetup paperSize=\"9\" orientation=\"landscape\"/>"},
    {"page_margins", &fm_sheet_get_page_margins_xml, &fm_sheet_set_page_margins_xml,
     "<pageMargins left=\"1\" right=\"1\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"},
    {"print_options", &fm_sheet_get_print_options_xml, &fm_sheet_set_print_options_xml,
     "<printOptions gridLines=\"1\"/>"},
    {"header_footer", &fm_sheet_get_header_footer_xml, &fm_sheet_set_header_footer_xml,
     "<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>"},
    {"sheet_pr", &fm_sheet_get_sheet_pr_xml, &fm_sheet_set_sheet_pr_xml,
     "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>"},
};

// The element as the handle reports it; "" when the sheet has none.
std::string ReadXml(const XmlSetterCase& c, fm_workbook_t* wb) {
  const char* xml = nullptr;
  EXPECT_EQ(c.get(wb, 0, &xml), 0) << c.name;
  return xml == nullptr ? std::string() : std::string(xml);
}

TEST(SheetXmlEmptyStrings, EmptyStringIsRefusedAndChangesNothing) {
  for (const XmlSetterCase& c : kXmlSetters) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(c.set(wb.handle, 0, c.fragment), 0) << c.name;
    const std::string before = ReadXml(c, wb.handle);
    ASSERT_FALSE(before.empty()) << c.name;

    EXPECT_EQ(c.set(wb.handle, 0, ""), kInvalidArgument) << c.name;
    EXPECT_EQ(ReadXml(c, wb.handle), before) << c.name;
  }
}

TEST(SheetXmlEmptyStrings, EmptyStringOnAbsentElementIsRefusedAndChangesNothing) {
  for (const XmlSetterCase& c : kXmlSetters) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before = ReadXml(c, wb.handle);
    EXPECT_EQ(c.set(wb.handle, 0, ""), kInvalidArgument) << c.name;
    EXPECT_EQ(ReadXml(c, wb.handle), before) << c.name;
  }
}

TEST(SheetXmlEmptyStrings, NullRemovesTheElement) {
  for (const XmlSetterCase& c : kXmlSetters) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(c.set(wb.handle, 0, c.fragment), 0) << c.name;
    ASSERT_FALSE(ReadXml(c, wb.handle).empty()) << c.name;

    EXPECT_EQ(c.set(wb.handle, 0, nullptr), 0) << c.name;
    EXPECT_EQ(ReadXml(c, wb.handle), "") << c.name;
  }
}

TEST(SheetXmlEmptyStrings, NullOnAbsentElementIsANoOp) {
  for (const XmlSetterCase& c : kXmlSetters) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(c.set(wb.handle, 0, nullptr), 0) << c.name;
    EXPECT_EQ(ReadXml(c, wb.handle), "") << c.name;
  }
}

TEST(SheetXmlEmptyStrings, NullWorkbookIsStillANullPointer) {
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  for (const XmlSetterCase& c : kXmlSetters) {
    EXPECT_EQ(c.set(nullptr, 0, nullptr), kNullPointer) << c.name;
  }
}

TEST(SheetXmlEmptyStrings, NullPageSetupResetsTheStructuredView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup orientation=\"landscape\" scale=\"50\"/>"), 0);

  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, ""), kInvalidArgument);
  fm_page_setup_t kept = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &kept), 0);
  EXPECT_EQ(kept.orientation, FM_PAGE_ORIENTATION_LANDSCAPE);
  EXPECT_EQ(kept.scale_pct, 50U);

  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, nullptr), 0);
  fm_page_setup_t reset = FM_STRUCT_INIT(fm_page_setup_t);
  ASSERT_EQ(fm_sheet_get_page_setup(wb.handle, 0, &reset), 0);
  EXPECT_EQ(reset.orientation, FM_PAGE_ORIENTATION_PORTRAIT);
  EXPECT_EQ(reset.scale_pct, 100U);
}

// Reads the cell's phonetic guide; false when the cell has none.
bool ReadPhonetic(fm_workbook_t* wb, std::string* text) {
  int32_t present = 0;
  const char* reading = nullptr;
  EXPECT_EQ(fm_sheet_get_cell_phonetic(wb, 0, 0, 0, &present, &reading), 0);
  if (present != 0 && reading != nullptr) {
    *text = reading;
  }
  return present != 0;
}

TEST(SheetXmlEmptyStrings, PhoneticEmptyStringAndNullBothClear) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "kanji"), 0);

  std::string text;
  ASSERT_EQ(fm_sheet_set_cell_phonetic(wb.handle, 0, 0, 0, "kana"), 0);
  ASSERT_TRUE(ReadPhonetic(wb.handle, &text));
  EXPECT_EQ(text, "kana");

  EXPECT_EQ(fm_sheet_set_cell_phonetic(wb.handle, 0, 0, 0, ""), 0);
  EXPECT_FALSE(ReadPhonetic(wb.handle, &text));

  ASSERT_EQ(fm_sheet_set_cell_phonetic(wb.handle, 0, 0, 0, "kana"), 0);
  ASSERT_TRUE(ReadPhonetic(wb.handle, &text));
  EXPECT_EQ(fm_sheet_set_cell_phonetic(wb.handle, 0, 0, 0, nullptr), 0);
  EXPECT_FALSE(ReadPhonetic(wb.handle, &text));
}

TEST(SheetXmlEmptyStrings, PhoneticNullWorkbookIsStillANullPointer) {
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  EXPECT_EQ(fm_sheet_set_cell_phonetic(nullptr, 0, 0, 0, nullptr), kNullPointer);
}

// Section text of the given header/footer slot; nullptr when the section is absent.
const char* const* SectionSlot(fm_header_footer_t& read, bool header) {
  return header ? &read.odd_header : &read.odd_footer;
}

TEST(SheetXmlEmptyStrings, HeaderFooterSectionEmptyStringAndClearBothRemove) {
  for (const bool header : {true, false}) {
    for (const bool use_clear : {false, true}) {
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

      fm_header_footer_patch_t set = FM_STRUCT_INIT(fm_header_footer_patch_t);
      set.odd_header_op = FM_PATCH_OP_SET;
      set.odd_header = "&CHead";
      set.odd_footer_op = FM_PATCH_OP_SET;
      set.odd_footer = "&CFoot";
      ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &set), 0);

      fm_header_footer_patch_t remove = FM_STRUCT_INIT(fm_header_footer_patch_t);
      fm_patch_op_t& op = header ? remove.odd_header_op : remove.odd_footer_op;
      const char*& text = header ? remove.odd_header : remove.odd_footer;
      if (use_clear) {
        op = FM_PATCH_OP_CLEAR;
      } else {
        op = FM_PATCH_OP_SET;
        text = "";
      }
      ASSERT_EQ(fm_sheet_update_header_footer(wb.handle, 0, &remove), 0);

      fm_header_footer_t read = FM_STRUCT_INIT(fm_header_footer_t);
      ASSERT_EQ(fm_sheet_get_header_footer(wb.handle, 0, &read), 0);
      EXPECT_EQ(*SectionSlot(read, header), nullptr)
          << (header ? "header" : "footer") << (use_clear ? " clear" : " empty");
      // The other section is untouched.
      const char* other = *SectionSlot(read, !header);
      ASSERT_NE(other, nullptr);
      EXPECT_STREQ(other, header ? "&CFoot" : "&CHead");
    }
  }
}

}  // namespace
