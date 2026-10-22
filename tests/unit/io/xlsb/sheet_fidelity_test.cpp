// Sheet-level XLSB fidelity: view mode and active-range index, extra workbook
// windows' views, and protected ranges, through the model and the records.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/protection_xml.h"
#include "io/sheet_view_xml.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/view_records.h"
#include "io/xlsb/writer.h"
#include "protection.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"
#include "writer_test_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

Workbook RoundTrip(const Workbook& wb, std::uint32_t* deferred = nullptr) {
  auto written = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(written)) << written.error().message;
  if (deferred != nullptr) {
    *deferred = written.value().diagnostics.deferred_feature_count;
  }
  auto read = read_xlsb(SpanOf(written.value().bytes));
  EXPECT_TRUE(static_cast<bool>(read)) << read.error().message;
  EXPECT_EQ(read.value().dropped_record_count, 0U);
  return std::move(read.value().workbook);
}

std::uint32_t DeferredCount(const Workbook& wb) {
  auto written = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(written)) << written.error().message;
  return written.value().diagnostics.deferred_feature_count;
}

std::vector<std::uint16_t> RecordTypes(const std::vector<std::uint8_t>& body) {
  std::vector<std::uint16_t> types;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size != 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      break;
    }
    types.push_back(rec.value().type);
  }
  return types;
}

std::size_t CountOf(const std::vector<std::uint16_t>& types, std::uint16_t type) {
  return static_cast<std::size_t>(std::count(types.begin(), types.end(), type));
}

std::vector<std::uint8_t> EmitSheetBody(const Sheet& sheet) {
  SstBuilder sst;
  auto body = emit_sheet(sheet, sst, {sheet.name()}, SheetRangeTable{}, NameTable{});
  EXPECT_TRUE(static_cast<bool>(body)) << body.error().message;
  return body.value();
}

/// Decodes `body` as sheet 0 of `wb`.
SheetDecodeState DecodeBody(const std::vector<std::uint8_t>& body, Workbook& wb) {
  std::deque<std::string> text_storage;
  std::uint32_t undecoded = 0;
  auto state = DecodeSheetBin(body, 0, wb, {}, {}, {}, text_storage, {wb.sheet(0).name()}, {}, {}, {}, &undecoded);
  EXPECT_TRUE(static_cast<bool>(state)) << state.error().message;
  return std::move(state.value());
}

/// `body` with `records` spliced in after the first BrtEndWsView.
std::vector<std::uint8_t> AfterMainView(const std::vector<std::uint8_t>& body,
                                        const std::vector<std::uint8_t>& records) {
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size != 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      break;
    }
    if (rec.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsView)) {
      const auto at = static_cast<std::ptrdiff_t>(cursor.data - body.data());
      std::vector<std::uint8_t> out(body.begin(), body.begin() + at);
      out.insert(out.end(), records.begin(), records.end());
      out.insert(out.end(), body.begin() + at, body.end());
      return out;
    }
  }
  ADD_FAILURE() << "no BrtEndWsView";
  return body;
}

TEST(XlsbSheetFidelity, ViewModeAndActiveRangeIndexSurvive) {
  Workbook wb = Workbook::create_empty();
  SheetView& view = wb.sheet(wb.add_sheet("View")).mutable_view();
  view.view_mode = "pageBreakPreview";
  view.selections.push_back(SheetSelection{"topLeft", "D4", "A1:B2 D4:E5", 1U});

  const Workbook back = RoundTrip(wb);
  const SheetView& v = back.sheet(0).view();
  EXPECT_EQ(v.view_mode, "pageBreakPreview");
  ASSERT_EQ(v.selections.size(), 1U);
  EXPECT_EQ(v.selections[0].active_cell, "D4");
  EXPECT_EQ(v.selections[0].sqref, "A1:B2 D4:E5");
  EXPECT_EQ(v.selections[0].active_cell_id, 1U);

  view.view_mode = "pageLayout";
  EXPECT_EQ(RoundTrip(wb).sheet(0).view().view_mode, "pageLayout");
}

TEST(XlsbSheetFidelity, UnknownViewModeAndViewExtensionsAreDeferred) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("View");
  const std::uint32_t baseline = DeferredCount(wb);
  SheetView& view = wb.sheet(0).mutable_view();
  view.view_mode = "futureMode";
  view.tail_xml = "<extLst/>";
  EXPECT_EQ(DeferredCount(wb), baseline + 2U);
  EXPECT_EQ(RoundTrip(wb).sheet(0).view().view_mode, "");
}

TEST(XlsbSheetFidelity, ExtraWindowViewDecodesIntoExtraViewsXml) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Windows");
  SheetView extra;
  extra.zoom_scale = 80U;
  extra.show_grid_lines = false;
  extra.set_freeze(2U, 0U);
  extra.selections.push_back(SheetSelection{"bottomLeft", "A5", "A5", 0U});
  std::vector<std::uint8_t> extra_records;
  emit_ws_view(extra_records, extra, 1U);
  const std::vector<std::uint8_t> body = AfterMainView(EmitSheetBody(wb.sheet(0)), extra_records);

  Workbook loaded = Workbook::create_empty();
  loaded.add_sheet("Windows");
  const SheetDecodeState state = DecodeBody(body, loaded);
  EXPECT_EQ(state.dropped_records, 0U);
  const SheetView& main = loaded.sheet(0).view();
  EXPECT_EQ(main.zoom_scale, SheetView::kDefaultZoomScale);
  EXPECT_FALSE(main.is_frozen());

  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(main.extra_views_xml.c_str())) << main.extra_views_xml;
  const pugi::xml_node node = doc.child("sheetView");
  ASSERT_TRUE(node) << main.extra_views_xml;
  EXPECT_FALSE(node.next_sibling("sheetView"));
  EXPECT_STREQ(node.attribute("workbookViewId").value(), "1");
  SheetView reread;
  read_sheet_view_element(node, reread);
  EXPECT_EQ(reread.zoom_scale, 80U);
  EXPECT_FALSE(reread.show_grid_lines);
  EXPECT_EQ(reread.freeze_rows, 2U);
  ASSERT_EQ(reread.selections.size(), 1U);
  EXPECT_EQ(reread.selections[0].pane, "bottomLeft");
  EXPECT_EQ(reread.selections[0].active_cell, "A5");

  // The workbook part writes one window, so the second sheet view has none
  // to attach to: it is not written and is reported.
  const std::vector<std::uint16_t> types = RecordTypes(EmitSheetBody(loaded.sheet(0)));
  EXPECT_EQ(CountOf(types, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsView)), 1U);
  EXPECT_EQ(unwritten_view_part_count(main), 1U);
}

TEST(XlsbSheetFidelity, ExtraWindowViewFromXlsxIsReportedWhenNotWritten) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Windows");
  const std::uint32_t baseline = DeferredCount(wb);
  wb.sheet(0).mutable_view().extra_views_xml = "<sheetView workbookViewId=\"1\" zoomScale=\"80\"/>";
  std::uint32_t deferred = 0;
  const Workbook back = RoundTrip(wb, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);
  EXPECT_TRUE(back.sheet(0).view().extra_views_xml.empty());
}

std::vector<ProtectedRange> SampleRanges() {
  ProtectedRange legacy;
  legacy.name = "Inputs";
  legacy.sqref = "A1:B2 D4";
  legacy.password = "CC1A";
  ProtectedRange hashed;
  hashed.name = "Totals";
  hashed.sqref = "F1:F10";
  hashed.algorithm_name = "SHA-512";
  hashed.hash_value = "q83vEjRWeJA=";
  hashed.salt_value = "AAECAwQFBgc=";
  hashed.spin_count = 100000U;
  ProtectedRange open;
  open.name = "Free";
  open.sqref = "H1";
  return {legacy, hashed, open};
}

void ExpectSameRanges(const std::vector<ProtectedRange>& got, const std::vector<ProtectedRange>& want) {
  ASSERT_EQ(got.size(), want.size());
  for (std::size_t i = 0; i < want.size(); ++i) {
    SCOPED_TRACE(want[i].name);
    EXPECT_EQ(got[i].name, want[i].name);
    EXPECT_EQ(got[i].sqref, want[i].sqref);
    EXPECT_EQ(got[i].password, want[i].password);
    EXPECT_EQ(got[i].algorithm_name, want[i].algorithm_name);
    EXPECT_EQ(got[i].hash_value, want[i].hash_value);
    EXPECT_EQ(got[i].salt_value, want[i].salt_value);
    EXPECT_EQ(got[i].spin_count, want[i].spin_count);
    EXPECT_EQ(got[i].security_descriptor, want[i].security_descriptor);
  }
}

TEST(XlsbSheetFidelity, ProtectedRangesSurviveXlsbAndXlsbToXlsb) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Ranges");
  const std::uint32_t baseline = DeferredCount(wb);
  write_protected_ranges(wb.sheet(0).mutable_raw_extensions(), SampleRanges());

  const std::vector<std::uint16_t> types = RecordTypes(EmitSheetBody(wb.sheet(0)));
  EXPECT_EQ(CountOf(types, kBrtRangeProtection), 3U);
  EXPECT_EQ(CountOf(types, kBrtRangeProtectionIso), 1U);

  std::uint32_t deferred = 0;
  Workbook once = RoundTrip(wb, &deferred);
  EXPECT_EQ(deferred, baseline);
  ExpectSameRanges(read_protected_ranges(once.sheet(0).raw_extensions()), SampleRanges());
  EXPECT_TRUE(once.sheet(0).xlsb_tail().empty());
  const Workbook twice = RoundTrip(once);
  ExpectSameRanges(read_protected_ranges(twice.sheet(0).raw_extensions()), SampleRanges());
}

TEST(XlsbSheetFidelity, ProtectedRangeWithSecurityDescriptorIsSkippedAndReported) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Ranges");
  const std::uint32_t baseline = DeferredCount(wb);
  std::vector<ProtectedRange> ranges = SampleRanges();
  ranges[0].security_descriptor = "O:WDG:WDD:(A;;CC;;;WD)";
  write_protected_ranges(wb.sheet(0).mutable_raw_extensions(), ranges);
  std::uint32_t deferred = 0;
  const Workbook back = RoundTrip(wb, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);
  ExpectSameRanges(read_protected_ranges(back.sheet(0).raw_extensions()), {SampleRanges()[1], SampleRanges()[2]});
}

std::vector<std::uint8_t> RangeProtectionWithDescriptor() {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, 0U);
  emit_u32(payload, 1U);
  emit_rfx(payload, MergeRange{0U, 0U, 3U, 3U});
  emit_xlwidestring(payload, "Team");
  const std::vector<std::uint8_t> descriptor = {1U, 0U, 4U, 0x80U, 0U, 0U, 0U, 0U, 0U, 0U,
                                                0U, 0U, 0U, 0U,    0U, 0U, 0U, 0U, 0U, 0U};
  emit_u32(payload, static_cast<std::uint32_t>(descriptor.size()));
  payload.insert(payload.end(), descriptor.begin(), descriptor.end());
  std::vector<std::uint8_t> record;
  emit_record(record, kBrtRangeProtection, payload);
  return record;
}

/// `body` with `records` spliced in before the first record of `type`.
std::vector<std::uint8_t> Before(const std::vector<std::uint8_t>& body, std::uint16_t type,
                                 const std::vector<std::uint8_t>& records) {
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size != 0U) {
    const std::uint8_t* at = cursor.data;
    auto rec = read_record(cursor);
    if (!rec) {
      break;
    }
    if (rec.value().type == type) {
      const auto offset = static_cast<std::ptrdiff_t>(at - body.data());
      std::vector<std::uint8_t> out(body.begin(), body.begin() + offset);
      out.insert(out.end(), records.begin(), records.end());
      out.insert(out.end(), body.begin() + offset, body.end());
      return out;
    }
  }
  ADD_FAILURE() << "no record " << type;
  return body;
}

TEST(XlsbSheetFidelity, RangeProtectionWithDescriptorIsRetainedVerbatim) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Ranges");
  write_protected_ranges(wb.sheet(0).mutable_raw_extensions(), {SampleRanges()[0]});
  const std::vector<std::uint8_t> raw = RangeProtectionWithDescriptor();
  const std::vector<std::uint8_t> body =
      Before(EmitSheetBody(wb.sheet(0)), static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), raw);

  Workbook loaded = Workbook::create_empty();
  loaded.add_sheet("Ranges");
  const SheetDecodeState state = DecodeBody(body, loaded);
  EXPECT_EQ(state.dropped_records, 0U);
  ExpectSameRanges(read_protected_ranges(loaded.sheet(0).raw_extensions()), {SampleRanges()[0]});
  const std::vector<std::uint8_t> out = EmitSheetBody(loaded.sheet(0));
  EXPECT_NE(std::search(out.begin(), out.end(), raw.begin(), raw.end()), out.end());
  EXPECT_EQ(CountOf(RecordTypes(out), kBrtRangeProtection), 2U);
}

TEST(XlsbSheetFidelity, IsoRangeProtectionWithoutItsCompanionIsRetained) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Ranges");
  std::vector<std::uint8_t> records;
  ASSERT_TRUE(static_cast<bool>(emit_range_protections(records, {SampleRanges()[1]})));
  // Keep the Iso record alone.
  ByteSpan cursor{records.data(), records.size()};
  ASSERT_TRUE(static_cast<bool>(read_record(cursor)));
  const std::vector<std::uint8_t> iso(records.cbegin(), records.cbegin() + (cursor.data - records.data()));
  const std::vector<std::uint8_t> body =
      Before(EmitSheetBody(wb.sheet(0)), static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), iso);

  Workbook loaded = Workbook::create_empty();
  loaded.add_sheet("Ranges");
  const SheetDecodeState state = DecodeBody(body, loaded);
  EXPECT_EQ(state.dropped_records, 0U);
  EXPECT_TRUE(read_protected_ranges(loaded.sheet(0).raw_extensions()).empty());
  EXPECT_FALSE(loaded.sheet(0).xlsb_tail().empty());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
