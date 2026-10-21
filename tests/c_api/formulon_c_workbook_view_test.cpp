// Stable C ABI active-sheet, show-phonetic and phonetic-properties presence tests.

#include "formulon_c_pivot_test_helpers.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

// A package of `sheets` empty worksheets; `book_views` is spliced into workbook.xml and `sheet1_xml` replaces the
// first worksheet when non-empty.
std::vector<std::uint8_t> BuildPackage(std::string_view book_views, std::string_view sheet1_xml = {}) {
  std::string workbook =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
  workbook += book_views;
  workbook +=
      "<sheets><sheet name=\"A\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"B\" sheetId=\"2\" r:id=\"rId2\"/>"
      "<sheet name=\"C\" sheetId=\"3\" r:id=\"rId3\"/></sheets></workbook>";
  const std::string content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet2.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet3.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>";
  const std::string package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/></Relationships>";
  std::string workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
  for (int i = 1; i <= 3; ++i) {
    workbook_rels += "<Relationship Id=\"rId" + std::to_string(i) +
                     "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
                     "Target=\"worksheets/sheet" +
                     std::to_string(i) + ".xml\"/>";
  }
  workbook_rels += "</Relationships>";
  const std::string empty_sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>";
  const std::string first_sheet = sheet1_xml.empty() ? empty_sheet : std::string(sheet1_xml);
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", first_sheet},
      {"xl/worksheets/sheet2.xml", empty_sheet},
      {"xl/worksheets/sheet3.xml", empty_sheet},
  });
}

std::vector<std::uint8_t> Save(fm_workbook_t* wb, fm_workbook_format_t format = FM_WORKBOOK_FORMAT_XLSX) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save_as(wb, format, &saved.data, &saved.len), 0) << fm_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

void Load(const std::vector<std::uint8_t>& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0) << fm_last_error_message();
}

std::uint32_t ActiveSheet(const fm_workbook_t* wb) {
  std::uint32_t index = 99;
  EXPECT_EQ(fm_workbook_get_active_sheet(wb, &index), 0) << fm_last_error_message();
  return index;
}

}  // namespace

TEST(FormulonCApiWorkbookView, ActiveSheetDefaultsToFirstAndSynthesizesBookViews) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Third"), 0);
  EXPECT_EQ(ActiveSheet(wb.handle), 0U);

  ASSERT_EQ(fm_workbook_set_active_sheet(wb.handle, 2), 0) << fm_last_error_message();
  EXPECT_EQ(ActiveSheet(wb.handle), 2U);

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_NE(ExtractZipEntry(saved, "xl/workbook.xml").find("<bookViews><workbookView activeTab=\"2\"/></bookViews>"),
            std::string::npos);
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  EXPECT_EQ(ActiveSheet(reloaded.handle), 2U);
}

TEST(FormulonCApiWorkbookView, SetActiveSheetKeepsOtherAttributesAndClampsFirstSheet) {
  const std::vector<std::uint8_t> bytes = BuildPackage(
      "<bookViews><workbookView xWindow=\"10\" firstSheet=\"2\" activeTab=\"2\" tabRatio=\"600\"/></bookViews>");
  WorkbookGuard wb;
  Load(bytes, wb);
  EXPECT_EQ(ActiveSheet(wb.handle), 2U);

  ASSERT_EQ(fm_workbook_set_active_sheet(wb.handle, 1), 0);
  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/workbook.xml");
  EXPECT_NE(xml.find("xWindow=\"10\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("tabRatio=\"600\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("activeTab=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("firstSheet=\"1\""), std::string::npos) << xml;
  EXPECT_EQ(ActiveSheet(wb.handle), 1U);
}

TEST(FormulonCApiWorkbookView, ActiveSheetRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_set_active_sheet(wb.handle, 1), kInvalidArgument);
  EXPECT_EQ(ActiveSheet(wb.handle), 0U);
  std::uint32_t index = 7;
  EXPECT_NE(fm_workbook_get_active_sheet(nullptr, &index), 0);
  EXPECT_NE(fm_workbook_get_active_sheet(wb.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_set_active_sheet(nullptr, 0), 0);
}

TEST(FormulonCApiWorkbookView, SetActiveSheetMovesTabSelectionAndSurvivesXlsb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Third"), 0);
  ASSERT_EQ(fm_sheet_set_tab_selected(wb.handle, 0, 1), 0);
  ASSERT_EQ(fm_workbook_set_active_sheet(wb.handle, 2), 0) << fm_last_error_message();
  for (std::size_t i = 0; i < 3; ++i) {
    fm_sheet_view_t view{};
    ASSERT_EQ(fm_sheet_get_view(wb.handle, i, &view), 0);
    EXPECT_EQ(view.tab_selected != 0, i == 2U) << "sheet " << i;
  }

  WorkbookGuard reloaded;
  Load(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), reloaded);
  EXPECT_EQ(ActiveSheet(reloaded.handle), 2U);
}

TEST(FormulonCApiWorkbookView, SetActiveSheetRejectsHiddenSheets) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "VeryHidden"), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_HIDDEN), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 2, FM_SHEET_VERY_HIDDEN), 0);
  EXPECT_EQ(fm_workbook_set_active_sheet(wb.handle, 1), kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_active_sheet(wb.handle, 2), kInvalidArgument);
  EXPECT_EQ(ActiveSheet(wb.handle), 0U);
}

TEST(FormulonCApiWorkbookView, SheetEditsKeepTheActiveSheetValid) {
  WorkbookGuard wb;
  Load(BuildPackage("<bookViews><workbookView activeTab=\"2\"/></bookViews>"), wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 2, 0), 0);
  EXPECT_EQ(ActiveSheet(wb.handle), 0U);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 0, 2), 0);
  EXPECT_EQ(ActiveSheet(wb.handle), 2U);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 2), 0);
  EXPECT_LT(ActiveSheet(wb.handle), fm_workbook_sheet_count(wb.handle));
  EXPECT_EQ(ActiveSheet(wb.handle), 1U);
}

TEST(FormulonCApiWorkbookView, ShowPhoneticSurvivesValueWritesAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  int32_t show = -1;
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 0, &show), 0);
  EXPECT_EQ(show, 0);

  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);
  ASSERT_EQ(fm_workbook_set_cell_show_phonetic(wb.handle, 0, 0, 0, 1), 0);
  // A flag on a cell with no content is kept as well.
  ASSERT_EQ(fm_workbook_set_cell_show_phonetic(wb.handle, 0, 4, 2, 1), 0);

  // Re-entering the value clears the readings but not the display flag.
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  const char* phonetic = nullptr;
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  EXPECT_STREQ(phonetic, "");
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 0, &show), 0);
  EXPECT_EQ(show, 1);

  for (const fm_workbook_format_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<std::uint8_t> saved = Save(wb.handle, format);
    if (format == FM_WORKBOOK_FORMAT_XLSX) {
      EXPECT_NE(ExtractZipEntry(saved, "xl/worksheets/sheet1.xml").find("ph=\"1\""), std::string::npos);
    }
    WorkbookGuard reloaded;
    Load(saved, reloaded);
    ASSERT_EQ(fm_workbook_get_cell_show_phonetic(reloaded.handle, 0, 0, 0, &show), 0);
    EXPECT_EQ(show, 1) << "format " << format;
    ASSERT_EQ(fm_workbook_get_cell_show_phonetic(reloaded.handle, 0, 4, 2, &show), 0);
    EXPECT_EQ(show, 1) << "format " << format;
    ASSERT_EQ(fm_workbook_get_cell_show_phonetic(reloaded.handle, 0, 1, 1, &show), 0);
    EXPECT_EQ(show, 0);
  }

  ASSERT_EQ(fm_workbook_set_cell_show_phonetic(wb.handle, 0, 0, 0, 0), 0);
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 0, &show), 0);
  EXPECT_EQ(show, 0);
  EXPECT_EQ(fm_workbook_set_cell_show_phonetic(wb.handle, 9, 0, 0, 1), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiWorkbookView, RowColumnAndCellPhoneticFlagsAreReadAndPreserved) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<cols><col min=\"3\" max=\"3\" width=\"12\" customWidth=\"1\" phonetic=\"1\"/></cols>"
      "<sheetData><row r=\"1\" ph=\"1\">"
      "<c r=\"A1\" t=\"inlineStr\" ph=\"1\"><is><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh></is></c>"
      "<c r=\"B1\" t=\"inlineStr\"><is><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh>"
      "<phoneticPr fontId=\"0\" type=\"halfwidthKatakana\" alignment=\"noControl\"/></is></c>"
      "</row><row r=\"2\"><c r=\"A2\" ph=\"1\"/></row></sheetData></worksheet>";
  WorkbookGuard wb;
  Load(BuildPackage("", sheet), wb);

  int32_t show = -1;
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 0, &show), 0);
  EXPECT_EQ(show, 1);
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 0, 1, &show), 0);
  EXPECT_EQ(show, 0);
  ASSERT_EQ(fm_workbook_get_cell_show_phonetic(wb.handle, 0, 1, 0, &show), 0);
  EXPECT_EQ(show, 1);

  // An absent `<phoneticPr>` is the inferred all-zero triple; an element spelling it out is explicit.
  int32_t present = -1;
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 0, &present), 0);
  EXPECT_EQ(present, 0);
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 1, &present), 0);
  EXPECT_EQ(present, 1);

  const std::string saved_sheet = ExtractZipEntry(Save(wb.handle), "xl/worksheets/sheet1.xml");
  EXPECT_NE(saved_sheet.find("<row r=\"1\" ph=\"1\""), std::string::npos) << saved_sheet;
  EXPECT_NE(saved_sheet.find("phonetic=\"1\""), std::string::npos) << saved_sheet;
  EXPECT_NE(saved_sheet.find("<c r=\"A2\" ph=\"1\"/>"), std::string::npos) << saved_sheet;
}

TEST(FormulonCApiWorkbookView, PhoneticPropertiesPresenceTracksExplicitTriples) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);

  int32_t present = -1;
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 0, &present), 0);
  EXPECT_EQ(present, 0);
  // An absent cell reports "not present" rather than failing.
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 7, 7, &present), 0);
  EXPECT_EQ(present, 0);

  ASSERT_EQ(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 0, 0, 0), 0);
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 0, &present), 0);
  EXPECT_EQ(present, 1);

  // The all-zero triple is written out and reads back explicit.
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(reloaded.handle, 0, 0, 0, &present), 0);
  EXPECT_EQ(present, 1);
  uint32_t font_id = 9;
  uint32_t type = 9;
  uint32_t alignment = 9;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(reloaded.handle, 0, 0, 0, &font_id, &type, &alignment), 0);
  EXPECT_EQ(font_id, 0U);
  EXPECT_EQ(type, 0U);
  EXPECT_EQ(alignment, 0U);

  // A value write drops the guide, the triple and its presence together.
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  ASSERT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 0, &present), 0);
  EXPECT_EQ(present, 0);

  EXPECT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 0, 0, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_has_cell_phonetic_properties(wb.handle, 5, 0, 0, &present), kInvalidArgument);
}

TEST(FormulonCApiWorkbookView, SheetViewScrollAndSelectionRoundTripFromAPackage) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetViews><sheetView tabSelected=\"1\" topLeftCell=\"C4\" workbookViewId=\"0\">"
      "<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B20\" activePane=\"bottomRight\" state=\"frozen\"/>"
      "<selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/>"
      "<selection pane=\"bottomLeft\" activeCell=\"A3\" sqref=\"A3\"/>"
      "<selection pane=\"bottomRight\" activeCell=\"C22\" sqref=\"C22:D24 F30 3:4\"/>"
      "</sheetView></sheetViews><sheetData/></worksheet>";
  WorkbookGuard wb;
  Load(BuildPackage("", sheet), wb);

  fm_sheet_view_t view{};
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &view), 0);
  EXPECT_EQ(view.top_left_row, 3U);
  EXPECT_EQ(view.top_left_col, 2U);
  EXPECT_EQ(view.pane_top_left_row, 19U);
  EXPECT_EQ(view.pane_top_left_col, 1U);

  uint32_t row = 99;
  uint32_t col = 99;
  size_t total = 0;
  ASSERT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, nullptr, 0, &total), 0);
  EXPECT_EQ(row, 21U);
  EXPECT_EQ(col, 2U);
  ASSERT_EQ(total, 3U);
  fm_merge_range ranges[3]{};
  size_t again = 0;
  ASSERT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, ranges, 3, &again), 0);
  EXPECT_EQ(again, 3U);
  EXPECT_EQ(ranges[0].first_row, 21U);
  EXPECT_EQ(ranges[0].last_col, 3U);
  // Whole rows `3:4` widen to the full sheet width.
  EXPECT_EQ(ranges[2].first_row, 2U);
  EXPECT_EQ(ranges[2].last_row, 3U);
  EXPECT_EQ(ranges[2].first_col, 0U);
  EXPECT_EQ(ranges[2].last_col, formulon::Sheet::kMaxCols - 1U);
  // A short buffer is filled as far as it goes and the total still reports every range.
  fm_merge_range one[1]{};
  ASSERT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, one, 1, &again), 0);
  EXPECT_EQ(again, 3U);
  EXPECT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, nullptr, 1, &again),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  // Nothing is dropped on save: origin, pane, and every pane's selection.
  const std::string saved = ExtractZipEntry(Save(wb.handle), "xl/worksheets/sheet1.xml");
  EXPECT_NE(saved.find("topLeftCell=\"C4\""), std::string::npos) << saved;
  EXPECT_NE(saved.find("topLeftCell=\"B20\" activePane=\"bottomRight\" state=\"frozen\""), std::string::npos) << saved;
  EXPECT_NE(saved.find("<selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/>"), std::string::npos) << saved;
  EXPECT_NE(saved.find("<selection pane=\"bottomLeft\" activeCell=\"A3\" sqref=\"A3\"/>"), std::string::npos) << saved;
  EXPECT_NE(saved.find("sqref=\"C22:D24 F30 3:4\""), std::string::npos) << saved;
}

TEST(FormulonCApiWorkbookView, SelectionReadSkipsInvertedAndOffGridWholeRowsAndColumns) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetViews><sheetView workbookViewId=\"0\">"
      "<selection activeCell=\"B2\" sqref=\"C:A XFE:XFF 5:3 1048577:1048578 B2\"/>"
      "</sheetView></sheetViews><sheetData/></worksheet>";
  WorkbookGuard wb;
  Load(BuildPackage("", sheet), wb);
  uint32_t row = 99;
  uint32_t col = 99;
  size_t total = 0;
  fm_merge_range ranges[5]{};
  ASSERT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, ranges, 5, &total), 0);
  ASSERT_EQ(total, 1U);
  EXPECT_EQ(ranges[0].first_row, 1U);
  EXPECT_EQ(ranges[0].first_col, 1U);
  EXPECT_EQ(ranges[0].last_row, 1U);
  EXPECT_EQ(ranges[0].last_col, 1U);
}

TEST(FormulonCApiWorkbookView, UnfrozenSplitViewSavesOnlyTheTopLeftSelection) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetViews><sheetView workbookViewId=\"0\">"
      "<pane xSplit=\"2000\" ySplit=\"1500\" topLeftCell=\"D10\" activePane=\"bottomRight\"/>"
      "<selection activeCell=\"A2\" sqref=\"A2\"/>"
      "<selection pane=\"bottomRight\" activeCell=\"E12\" sqref=\"E12\"/>"
      "</sheetView></sheetViews><sheetData/></worksheet>";
  WorkbookGuard wb;
  Load(BuildPackage("", sheet), wb);
  const std::string saved = ExtractZipEntry(Save(wb.handle), "xl/worksheets/sheet1.xml");
  EXPECT_NE(saved.find("<selection activeCell=\"A2\" sqref=\"A2\"/>"), std::string::npos) << saved;
  EXPECT_EQ(saved.find("E12"), std::string::npos) << saved;
}

TEST(FormulonCApiWorkbookView, SettersWriteTheFrozenPaneTheDesktopWayAndRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, 2, 1), 0);
  // Without a stored pane origin the first unfrozen cell is reported and written.
  fm_sheet_view_t view{};
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &view), 0);
  EXPECT_EQ(view.pane_top_left_row, 2U);
  EXPECT_EQ(view.pane_top_left_col, 1U);

  ASSERT_EQ(fm_sheet_set_top_left(wb.handle, 0, 5, 0), 0);
  ASSERT_EQ(fm_sheet_set_pane_top_left(wb.handle, 0, 40, 6), 0);
  EXPECT_EQ(fm_sheet_set_pane_top_left(wb.handle, 0, 1, 6), kInvalidArgument);  // inside the frozen rows
  EXPECT_EQ(fm_sheet_set_pane_top_left(wb.handle, 0, 4, 0), kInvalidArgument);  // inside the frozen columns
  const fm_merge_range ranges[] = {{41, 7, 43, 8}, {50, 9, 50, 9}};
  ASSERT_EQ(fm_sheet_set_selection(wb.handle, 0, 41, 7, ranges, 2), 0) << fm_last_error_message();
  EXPECT_EQ(fm_sheet_set_selection(wb.handle, 0, 0, 0, ranges, 0), 0);  // replaces: active cell alone
  ASSERT_EQ(fm_sheet_set_selection(wb.handle, 0, 41, 7, ranges, 2), 0);
  const fm_merge_range inverted{3, 3, 1, 1};
  EXPECT_EQ(fm_sheet_set_selection(wb.handle, 0, 0, 0, &inverted, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_selection(wb.handle, 0, formulon::Sheet::kMaxRows, 0, nullptr, 0), kInvalidArgument);

  for (const fm_workbook_format_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<std::uint8_t> bytes = Save(wb.handle, format);
    if (format == FM_WORKBOOK_FORMAT_XLSX) {
      const std::string xml = ExtractZipEntry(bytes, "xl/worksheets/sheet1.xml");
      EXPECT_NE(xml.find("topLeftCell=\"A6\""), std::string::npos) << xml;
      EXPECT_NE(
          xml.find("<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"G41\" activePane=\"bottomRight\" state=\"frozen\"/>"),
          std::string::npos)
          << xml;
      EXPECT_NE(xml.find("<selection pane=\"bottomRight\" activeCell=\"H42\" sqref=\"H42:I44 J51\"/>"),
                std::string::npos)
          << xml;
    }
    WorkbookGuard reloaded;
    Load(bytes, reloaded);
    ASSERT_EQ(fm_sheet_get_view(reloaded.handle, 0, &view), 0);
    EXPECT_EQ(view.top_left_row, 5U) << "format " << format;
    EXPECT_EQ(view.pane_top_left_row, 40U) << "format " << format;
    EXPECT_EQ(view.pane_top_left_col, 6U) << "format " << format;
    EXPECT_EQ(view.freeze_rows, 2U);
    if (format == FM_WORKBOOK_FORMAT_XLSX) {
      uint32_t row = 0;
      uint32_t col = 0;
      size_t count = 0;
      ASSERT_EQ(fm_sheet_get_selection(reloaded.handle, 0, &row, &col, nullptr, 0, &count), 0);
      EXPECT_EQ(row, 41U);
      EXPECT_EQ(col, 7U);
      EXPECT_EQ(count, 2U);
    }
  }

  // Changing the freeze re-derives the pane and carries the active selection along.
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, 0, 3), 0);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &view), 0);
  EXPECT_EQ(view.pane_top_left_row, 0U);
  EXPECT_EQ(view.pane_top_left_col, 3U);
  uint32_t row = 0;
  uint32_t col = 0;
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_selection(wb.handle, 0, &row, &col, nullptr, 0, &count), 0);
  EXPECT_EQ(row, 41U);
  EXPECT_EQ(count, 2U);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, 0, 0), 0);
  EXPECT_EQ(fm_sheet_set_pane_top_left(wb.handle, 0, 0, 0), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_top_left(wb.handle, 9, 0, 0), kInvalidArgument);
}
