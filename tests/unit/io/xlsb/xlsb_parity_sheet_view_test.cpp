// The worksheet view records (BrtBeginWsView, BrtPane, BrtSel inside BrtBeginWsViews) held against Excel 16.115's
// own xlsb and xlsx saves of the same sources (tests/fixtures/excel/xlsb_parity/views), and the normalisation both
// readers apply to what Excel normalises on open: an unknown view mode, a repeated window and a pivotSelection.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using test::Bytes;

constexpr std::uint16_t kBrtBeginWsView = 137;
constexpr std::uint16_t kBrtEndWsView = 138;
constexpr std::uint16_t kBrtSel = 152;
const std::set<std::uint16_t> kViewRecords = {133, 134, kBrtBeginWsView, kBrtEndWsView, 151, kBrtSel};

// Sources Excel opened and saved in both formats, its xlsb re-saved as well.
const char* const kOpened[] = {"view_extlst", "unknown_view", "window_id_control", "window_id_repeat"};
// Pivot books Excel opened, read but not converted here.
const char* const kOpenedPivots[] = {"pivot_selection_min_rid", "pivot_selection_native"};

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/views/" + name;
}

Workbook ReadXlsx(const std::string& name) {
  auto read = read_ooxml(reader_test_support::SpanOf(test::ReadBytes(Fixture(name))));
  EXPECT_TRUE(static_cast<bool>(read)) << name << ": " << (read ? "" : read.error().message);
  return std::move(read.value().workbook);
}

XlsbReadResult ReadXlsb(const Bytes& bytes) {
  auto read = read_xlsb(reader_test_support::SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return std::move(read.value());
}

Bytes WriteXlsb(const Workbook& workbook) {
  auto written = write_xlsb_with_result(workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << (written ? "" : written.error().message);
  return written ? written.value().bytes : Bytes{};
}

/// Each worksheet part's view records as (type, payload), in stream order.
std::vector<std::pair<std::string, std::vector<std::pair<std::uint16_t, Bytes>>>> ViewRecords(const Bytes& package) {
  std::vector<std::pair<std::string, std::vector<std::pair<std::uint16_t, Bytes>>>> out;
  for (const auto& [part, records] : test::ExtractRecords(package, kViewRecords)) {
    // The sheet parts only: binary index parts carry no view records.
    if (part.rfind("xl/worksheets/sheet", 0) != 0) {
      continue;
    }
    std::vector<std::pair<std::uint16_t, Bytes>> list;
    for (const test::IndexedRecord& record : records) {
      list.emplace_back(record.type, record.payload);
    }
    out.emplace_back(part, std::move(list));
  }
  return out;
}

void ExpectSameView(const SheetView& got, const SheetView& want) {
  EXPECT_EQ(got.zoom_scale, want.zoom_scale);
  EXPECT_EQ(got.freeze_rows, want.freeze_rows);
  EXPECT_EQ(got.freeze_cols, want.freeze_cols);
  EXPECT_EQ(got.split_x, want.split_x);
  EXPECT_EQ(got.split_y, want.split_y);
  EXPECT_EQ(got.tab_selected, want.tab_selected);
  EXPECT_EQ(got.show_grid_lines, want.show_grid_lines);
  EXPECT_EQ(got.view_mode, want.view_mode);
  EXPECT_EQ(got.top_left_cell, want.top_left_cell);
  EXPECT_EQ(got.extra_views_xml, want.extra_views_xml);
  ASSERT_EQ(got.selections.size(), want.selections.size());
  for (std::size_t i = 0; i < want.selections.size(); ++i) {
    EXPECT_EQ(SheetView::pane_key(got.selections[i].pane), SheetView::pane_key(want.selections[i].pane));
    EXPECT_EQ(got.selections[i].active_cell, want.selections[i].active_cell);
    EXPECT_EQ(got.selections[i].sqref, want.selections[i].sqref);
  }
}

TEST(XlsbParitySheetViewTwins, ConvertedMatchesExcel) {
  for (const char* name : kOpened) {
    SCOPED_TRACE(name);
    EXPECT_EQ(ViewRecords(WriteXlsb(ReadXlsx(std::string(name) + ".xlsx"))),
              ViewRecords(test::ReadBytes(Fixture(std::string(name) + ".xlsb"))));
  }
}

TEST(XlsbParitySheetViewTwins, ReadMatchesXlsxTwin) {
  std::vector<std::string> names(std::begin(kOpened), std::end(kOpened));
  names.insert(names.end(), std::begin(kOpenedPivots), std::end(kOpenedPivots));
  for (const std::string& name : names) {
    SCOPED_TRACE(name);
    const XlsbReadResult xlsb = ReadXlsb(test::ReadBytes(Fixture(name + ".xlsb")));
    EXPECT_EQ(xlsb.skipped_feature_count, 0U);
    const Workbook xlsx = ReadXlsx(name + ".xlsx");
    ASSERT_EQ(xlsb.workbook.sheet_count(), xlsx.sheet_count());
    for (std::size_t i = 0; i < xlsx.sheet_count(); ++i) {
      SCOPED_TRACE(i);
      ExpectSameView(xlsb.workbook.sheet(i).view(), xlsx.sheet(i).view());
      // Excel keeps the view extLst in xlsx and writes no record for it.
      EXPECT_TRUE(xlsb.workbook.sheet(i).view().tail_xml.empty());
      EXPECT_EQ(xlsx.sheet(i).view().tail_xml.empty(), name != "view_extlst");
    }
  }
}

TEST(XlsbParitySheetViewTwins, ResaveMatchesExcel) {
  for (const char* name : kOpened) {
    SCOPED_TRACE(name);
    const XlsbReadResult xlsb = ReadXlsb(test::ReadBytes(Fixture(std::string(name) + ".xlsb")));
    EXPECT_EQ(ViewRecords(WriteXlsb(xlsb.workbook)),
              ViewRecords(test::ReadBytes(Fixture(std::string(name) + ".2.xlsb"))));
  }
}

TEST(XlsbParitySheetView, NormaliseUnknownViewMode) {
  const Workbook src = ReadXlsx("unknown_view_src.xlsx");
  const Workbook twin = ReadXlsx("unknown_view.xlsx");
  ASSERT_EQ(src.sheet_count(), 3U);
  for (std::size_t i = 0; i < src.sheet_count(); ++i) {
    SCOPED_TRACE(i);
    EXPECT_EQ(src.sheet(i).view().view_mode, "");
    EXPECT_EQ(src.sheet(i).view().view_mode, twin.sheet(i).view().view_mode);
  }
}

TEST(XlsbParitySheetView, NormaliseRepeatedWindowKeepsTheLaterView) {
  const Workbook src = ReadXlsx("window_id_repeat_src.xlsx");
  const Workbook twin = ReadXlsx("window_id_repeat.xlsx");
  ASSERT_EQ(src.sheet_count(), 2U);
  const SheetView& view = src.sheet(1).view();
  ASSERT_EQ(view.selections.size(), 1U);
  EXPECT_EQ(view.selections[0].active_cell, "B2");
  EXPECT_TRUE(view.extra_views_xml.empty()) << view.extra_views_xml;
  const SheetView& want = twin.sheet(1).view();
  ASSERT_EQ(want.selections.size(), 1U);
  EXPECT_EQ(view.selections[0].active_cell, want.selections[0].active_cell);
  EXPECT_EQ(view.selections[0].sqref, want.selections[0].sqref);
}

TEST(XlsbParitySheetView, NormalisePivotSelectionIsNotRead) {
  for (const char* name : {"pivot_selection_src.xlsx", "pivot_selection_min_rid_src.xlsx",
                           "pivot_selection_min_src.xlsx", "pivot_selection_multi_src.xlsx"}) {
    SCOPED_TRACE(name);
    const Workbook book = ReadXlsx(name);
    for (std::size_t i = 0; i < book.sheet_count(); ++i) {
      EXPECT_TRUE(book.sheet(i).view().tail_xml.empty()) << book.sheet(i).view().tail_xml;
    }
  }
  const Workbook src = ReadXlsx("pivot_selection_min_rid_src.xlsx");
  ExpectSameView(src.sheet(0).view(), ReadXlsx("pivot_selection_min_rid.xlsx").sheet(0).view());
}

TEST(XlsbParitySheetView, NormaliseViewExtensionIsLeftOutOfXlsb) {
  const Workbook src = ReadXlsx("view_extlst_src.xlsx");
  EXPECT_NE(src.sheet(0).view().tail_xml.find("fm:marker"), std::string::npos);
  std::vector<std::uint16_t> types;
  for (const auto& [part, records] : ViewRecords(WriteXlsb(src))) {
    for (const auto& record : records) {
      types.push_back(record.first);
    }
  }
  EXPECT_EQ(types, (std::vector<std::uint16_t>{133, kBrtBeginWsView, kBrtSel, kBrtEndWsView, 134}));
}

TEST(XlsbParitySheetView, NormaliseWindowPastTheBookViewsReadsButRefusesXlsb) {
  for (const char* name : {"window_id_past_src.xlsx", "bad_window_id_src.xlsx"}) {
    SCOPED_TRACE(name);
    const Workbook book = ReadXlsx(name);
    EXPECT_NE(book.sheet(0).view().extra_views_xml.find("workbookViewId=\"5\""), std::string::npos);
    auto written = write_xlsb_with_result(book);
    ASSERT_FALSE(static_cast<bool>(written));
    EXPECT_EQ(written.error().code, FormulonErrorCode::kIoFeatureNotWritable);
    EXPECT_EQ(written.error().context, "context=write_xlsb kind=sheet_view sheet_index=0");
  }
  // The repeated window on bad_window_id's second sheet folds as in window_id_repeat.
  const Workbook bad = ReadXlsx("bad_window_id_src.xlsx");
  const SheetView& repeated = bad.sheet(1).view();
  EXPECT_TRUE(repeated.extra_views_xml.empty());
  ASSERT_EQ(repeated.selections.size(), 1U);
  EXPECT_EQ(repeated.selections[0].active_cell, "B2");
}

/// `package` with its first sheet's record stream rebuilt by `edit`.
template <typename Edit>
Bytes WithFirstSheet(const Bytes& package, Edit edit) {
  std::vector<reader_test_support::PartFile> parts;
  for (const std::string& name : test::ListParts(package)) {
    const std::optional<Bytes> body = test::ReadPart(package, name);
    EXPECT_TRUE(body.has_value()) << name;
    if (name != "xl/worksheets/sheet1.bin") {
      parts.push_back({name, body.value_or(Bytes{})});
      continue;
    }
    Bytes rebuilt;
    for (const test::IndexedRecord& record : test::SplitPart(body.value_or(Bytes{}))) {
      edit(record, rebuilt);
    }
    parts.push_back({name, std::move(rebuilt)});
  }
  return reader_test_support::BuildZip(parts);
}

/// A BrtBeginWsView payload with its iWbkView replaced.
Bytes WithWindow(Bytes payload, std::uint32_t window) {
  for (std::size_t i = 0; i < 4U; ++i) {
    payload[payload.size() - 4U + i] = static_cast<std::uint8_t>((window >> (8U * i)) & 0xFFU);
  }
  return payload;
}

/// A BrtSel of the top-left pane selecting C3 alone.
Bytes SelectC3() {
  Bytes p;
  for (const std::uint32_t v : {3U, 2U, 2U, 0U, 1U, 2U, 2U, 2U, 2U}) {
    reader_test_support::AppendU32(p, v);
  }
  return p;
}

TEST(XlsbParitySheetView, NormaliseXlsbReaderAligned) {
  const Bytes control = test::ReadBytes(Fixture("window_id_control.xlsb"));
  const auto append = [](const test::IndexedRecord& record, Bytes& out) {
    reader_test_support::AppendRecord(out, record.type, record.payload);
  };

  // An xlView past the three Excel writes reads as normal and is counted.
  const XlsbReadResult unknown_mode = ReadXlsb(WithFirstSheet(control, [&](const test::IndexedRecord& r, Bytes& out) {
    if (r.type != kBrtBeginWsView) {
      append(r, out);
      return;
    }
    Bytes payload = r.payload;
    payload[2] = 3U;  // xlView, after the u16 flags
    reader_test_support::AppendRecord(out, r.type, payload);
  }));
  EXPECT_EQ(unknown_mode.skipped_feature_count, 1U);
  EXPECT_EQ(unknown_mode.workbook.sheet(0).view().view_mode, "");

  // A second view: window 5 of a one-window book is skipped; window 0 replaces the first view.
  const auto with_second_view = [&](std::uint32_t window) {
    Bytes begin;
    return WithFirstSheet(control, [&, window](const test::IndexedRecord& r, Bytes& out) {
      append(r, out);
      if (r.type == kBrtBeginWsView) {
        begin = r.payload;
      } else if (r.type == kBrtEndWsView) {
        reader_test_support::AppendRecord(out, kBrtBeginWsView, WithWindow(begin, window));
        reader_test_support::AppendRecord(out, kBrtSel, SelectC3());
        reader_test_support::AppendRecord(out, kBrtEndWsView, {});
      }
    });
  };
  const XlsbReadResult past = ReadXlsb(with_second_view(5U));
  EXPECT_EQ(past.skipped_feature_count, 1U);
  EXPECT_TRUE(past.workbook.sheet(0).view().extra_views_xml.empty());
  ASSERT_EQ(past.workbook.sheet(0).view().selections.size(), 1U);
  EXPECT_EQ(past.workbook.sheet(0).view().selections[0].active_cell, "B2");

  const XlsbReadResult repeated = ReadXlsb(with_second_view(0U));
  EXPECT_EQ(repeated.skipped_feature_count, 0U);
  EXPECT_TRUE(repeated.workbook.sheet(0).view().extra_views_xml.empty());
  ASSERT_EQ(repeated.workbook.sheet(0).view().selections.size(), 1U);
  EXPECT_EQ(repeated.workbook.sheet(0).view().selections[0].active_cell, "C3");
  EXPECT_TRUE(repeated.workbook.sheet(0).view().tab_selected);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
