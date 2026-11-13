//
// Unit tests for the save-time dxf plan, the raw-XML dxf attribute scanner
// and the xlsb record dxf-slot table.

#include "io/dxf_save_plan.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <set>
#include <string>
#include <tuple>
#include <vector>

#include "gtest/gtest.h"
#include "io/dxf_xml_ids.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/dxf_record_slots.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"

namespace formulon::io {
namespace {

std::vector<DxfUse> Uses(std::initializer_list<std::uint32_t> indices) {
  std::vector<DxfUse> out;
  for (const std::uint32_t index : indices) {
    DxfUse use;
    use.model_index = index;
    out.push_back(use);
  }
  return out;
}

TEST(DxfSavePlan, OrphanIsDroppedAndUsersRenumbered) {
  const DxfSavePlan plan = build_dxf_save_plan(4, Uses({3, 0, 2, 3}), false);
  EXPECT_EQ(plan.order(), (std::vector<std::uint32_t>{0, 2, 3}));
  EXPECT_EQ(plan.saved_count(), 3U);
  EXPECT_FALSE(plan.identity());
  EXPECT_EQ(plan.map(0), 0U);
  EXPECT_EQ(plan.map(1), std::nullopt);
  EXPECT_EQ(plan.map(2), 1U);
  EXPECT_EQ(plan.map(3), 2U);
}

TEST(DxfSavePlan, OutOfRangeIndexMapsToNothing) {
  const DxfSavePlan plan = build_dxf_save_plan(3, Uses({1, 9, 0xFFFFFFFFU}), false);
  EXPECT_EQ(plan.order(), (std::vector<std::uint32_t>{1}));
  EXPECT_EQ(plan.map(3), std::nullopt);
  EXPECT_EQ(plan.map(9), std::nullopt);
  EXPECT_EQ(plan.map(0xFFFFFFFFU), std::nullopt);
  EXPECT_EQ(identity_dxf_plan(3).map(3), std::nullopt);
}

TEST(DxfSavePlan, BlockedYieldsIdentity) {
  const DxfSavePlan plan = build_dxf_save_plan(3, Uses({1}), true);
  EXPECT_EQ(plan.order(), (std::vector<std::uint32_t>{0, 1, 2}));
  EXPECT_TRUE(plan.identity());
  for (std::uint32_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(plan.map(i), i);
  }
}

TEST(DxfSavePlan, EmptyUsesOverNonEmptyTableIsEmpty) {
  const DxfSavePlan plan = build_dxf_save_plan(2, {}, false);
  EXPECT_TRUE(plan.order().empty());
  EXPECT_EQ(plan.saved_count(), 0U);
  EXPECT_FALSE(plan.identity());
  EXPECT_EQ(plan.map(0), std::nullopt);
}

TEST(DxfSavePlan, EveryDxfUsedIsIdentity) {
  EXPECT_TRUE(build_dxf_save_plan(2, Uses({1, 0}), false).identity());
  EXPECT_TRUE(build_dxf_save_plan(0, {}, false).identity());
  EXPECT_EQ(build_dxf_save_plan(0, {}, false).saved_count(), 0U);
}

TEST(DxfXmlIds, CollectsDxfIdAndTheSixTableNames) {
  const std::string xml =
      "<table headerRowDxfId=\"1\" dataDxfId=\"2\" totalsRowDxfId=\"3\" headerRowBorderDxfId=\"4\""
      " tableBorderDxfId=\"5\" totalsRowBorderDxfId=\"6\"><tableColumn dxfId='7'/></table>";
  auto ids = visit_dxf_xml_ids(xml);
  ASSERT_TRUE(ids);
  const std::vector<std::string> names = {
      "headerRowDxfId",       "dataDxfId", "totalsRowDxfId", "headerRowBorderDxfId", "tableBorderDxfId",
      "totalsRowBorderDxfId", "dxfId"};
  ASSERT_EQ(ids.value().size(), names.size());
  for (std::size_t i = 0; i < names.size(); ++i) {
    EXPECT_EQ(ids.value()[i].attribute, names[i]);
    EXPECT_EQ(ids.value()[i].value, i + 1U);
  }
  EXPECT_EQ(ids.value()[0].element, "table");
  EXPECT_EQ(ids.value()[6].element, "tableColumn");
}

TEST(DxfXmlIds, RewriteSplicesDigitsOnly) {
  const DxfSavePlan plan = build_dxf_save_plan(8, Uses({3, 7}), false);
  auto out = rewrite_dxf_xml_ids("<a dxfId = '003' x=\"1\"/>\n<b\tdxfId=\"7\">t</b>", plan);
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), "<a dxfId = '0' x=\"1\"/>\n<b\tdxfId=\"1\">t</b>");
}

TEST(DxfXmlIds, UnmappedIdRemovesAttributeAndLeadingWhitespace) {
  const DxfSavePlan plan = build_dxf_save_plan(3, Uses({0}), false);
  auto out = rewrite_dxf_xml_ids("<cfRule type=\"x\"  dxfId=\"2\" priority=\"1\"/><cfRule dxfId=\"0\"/>", plan);
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), "<cfRule type=\"x\" priority=\"1\"/><cfRule dxfId=\"0\"/>");
  auto last = rewrite_dxf_xml_ids("<cfRule dxfId=\"5\"/>", plan);
  ASSERT_TRUE(last);
  EXPECT_EQ(last.value(), "<cfRule/>");
}

TEST(DxfXmlIds, LeavesOtherNamesTextCommentsCdataAndPisUntouched) {
  const std::string xml =
      "<a fooDxfId=\"3\" x14:dxfId=\"4\" dxfIds=\"5\">dxfId=\"3\" &lt;b dxfId=\"6\"/&gt;</a>"
      "<!-- <b dxfId=\"1\"/> --><![CDATA[<c dxfId=\"2\"/>]]><?pi dxfId=\"5\"?></a>";
  auto ids = visit_dxf_xml_ids(xml);
  ASSERT_TRUE(ids);
  EXPECT_TRUE(ids.value().empty());
  auto out = rewrite_dxf_xml_ids(xml, build_dxf_save_plan(8, {}, false));
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), xml);
}

TEST(DxfXmlIds, NonIntegerValueIsNotAReference) {
  const std::string xml =
      "<a dxfId=\"-1\"/><a dxfId=\"\"/><a dxfId=\"1a\"/><a dxfId=\" 1\"/><a dxfId=\"4294967296\"/>"
      "<a dxfId=\"&#49;\"/>";
  auto ids = visit_dxf_xml_ids(xml);
  ASSERT_TRUE(ids);
  EXPECT_TRUE(ids.value().empty());
  auto out = rewrite_dxf_xml_ids(xml, build_dxf_save_plan(2, {}, false));
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), xml);
  auto max = visit_dxf_xml_ids("<a dxfId=\"4294967295\"/>");
  ASSERT_TRUE(max);
  ASSERT_EQ(max.value().size(), 1U);
  EXPECT_EQ(max.value()[0].value, 0xFFFFFFFFU);
}

TEST(DxfXmlIds, MultiRootFragmentWithPrefixedElements) {
  const DxfSavePlan plan = build_dxf_save_plan(3, Uses({1, 2}), false);
  const std::string xml = "<x14:cfRule dxfId=\"1\"/><x:sortCondition descending=\"1\" dxfId=\"2\"></x:sortCondition>";
  auto ids = visit_dxf_xml_ids(xml);
  ASSERT_TRUE(ids);
  ASSERT_EQ(ids.value().size(), 2U);
  EXPECT_EQ(ids.value()[0].element, "x14:cfRule");
  EXPECT_EQ(ids.value()[1].element, "x:sortCondition");
  auto out = rewrite_dxf_xml_ids(xml, plan);
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), "<x14:cfRule dxfId=\"0\"/><x:sortCondition descending=\"1\" dxfId=\"1\"></x:sortCondition>");
}

TEST(DxfXmlIds, SlicerAndTimelineStyleElementsAreIgnored) {
  const std::string xml =
      "<x14:slicerStyleElement type=\"wholeTable\" dxfId=\"0\"/><slicerStyleElement dxfId=\"1\"/>"
      "<x15:timelineStyleElement dxfId=\"2\"/><timelineStyleElement dxfId=\"3\"/><tableStyleElement dxfId=\"4\"/>";
  auto ids = visit_dxf_xml_ids(xml);
  ASSERT_TRUE(ids);
  ASSERT_EQ(ids.value().size(), 1U);
  EXPECT_EQ(ids.value()[0].element, "tableStyleElement");
  auto out = rewrite_dxf_xml_ids(xml, build_dxf_save_plan(5, {}, false));
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(),
            "<x14:slicerStyleElement type=\"wholeTable\" dxfId=\"0\"/><slicerStyleElement dxfId=\"1\"/>"
            "<x15:timelineStyleElement dxfId=\"2\"/><timelineStyleElement dxfId=\"3\"/><tableStyleElement/>");
}

TEST(DxfXmlIds, MalformedMarkupIsAnError) {
  const DxfSavePlan plan = identity_dxf_plan(2);
  for (const char* xml : {"<a dxfId=\"1\"", "<a dxfId=\"1/>", "<a dxfId=1/>", "<a dxfId/>", "<!-- open",
                          "<![CDATA[ open", "<?pi open", "</a", "<>", "<a / dxfId=\"1\">"}) {
    EXPECT_FALSE(visit_dxf_xml_ids(xml)) << xml;
    EXPECT_FALSE(rewrite_dxf_xml_ids(xml, plan)) << xml;
  }
}

TEST(DxfXmlIds, AttributeListVisitAndRewrite) {
  const std::string attrs = " dataDxfId=\"2\" dataCellStyle=\"S\" totalsRowDxfId=\"1\" xr3:uid=\"{X}\"";
  auto ids = visit_dxf_attr_ids(attrs);
  ASSERT_TRUE(ids);
  ASSERT_EQ(ids.value().size(), 2U);
  EXPECT_TRUE(ids.value()[0].element.empty());
  EXPECT_EQ(ids.value()[0].attribute, "dataDxfId");
  EXPECT_EQ(ids.value()[1].value, 1U);
  auto out = rewrite_dxf_attr_ids(attrs, build_dxf_save_plan(3, Uses({2}), false));
  ASSERT_TRUE(out);
  EXPECT_EQ(out.value(), " dataDxfId=\"0\" dataCellStyle=\"S\" xr3:uid=\"{X}\"");
  auto empty = rewrite_dxf_attr_ids("", identity_dxf_plan(1));
  ASSERT_TRUE(empty);
  EXPECT_TRUE(empty.value().empty());
}

TEST(DxfXmlIds, MalformedAttributeListIsAnError) {
  for (const char* attrs : {" dataDxfId=\"2", " dataDxfId", " a=\"1\"/>", " a=\"1\">"}) {
    EXPECT_FALSE(visit_dxf_attr_ids(attrs)) << attrs;
  }
}

}  // namespace

namespace xlsb {
namespace {

void EmitNullStrings(std::vector<std::uint8_t>& p, int count) {
  for (int i = 0; i < count; ++i) {
    emit_u32(p, 0xFFFFFFFFU);
  }
}

/// One record of each slot-bearing type, plus two BrtDXF and an unrelated record.
struct Fixture {
  std::vector<std::uint8_t> stream;
  /// Expected (record type, payload offset, value) of every reported slot.
  std::vector<std::tuple<std::uint16_t, std::size_t, std::uint32_t>> slots;
};

Fixture BuildFixture() {
  Fixture f;
  std::vector<std::uint8_t>& s = f.stream;
  std::vector<std::uint8_t> p;
  emit_record(s, kBrtDxf, std::vector<std::uint8_t>(10, 0));
  emit_record(s, kBrtDxf, std::vector<std::uint8_t>(10, 0));
  // BrtBeginList: RfX, lt, id, crwHeader, crwTotals, flags, six dxf ids, connection id, six strings.
  emit_rfx(p, MergeRange{0, 0, 3, 2});
  for (const std::uint32_t v : {0U, 1U, 1U, 0U, 1U}) {
    emit_u32(p, v);
  }
  for (const std::uint32_t v : {0U, kNoDxf, 1U, 2U, 3U, 4U}) {
    emit_u32(p, v);
  }
  emit_u32(p, 0U);
  EmitNullStrings(p, 6);
  emit_record(s, kBrtBeginList, p);
  f.slots.insert(f.slots.end(), {{kBrtBeginList, 36, 0U},
                                 {kBrtBeginList, 44, 1U},
                                 {kBrtBeginList, 48, 2U},
                                 {kBrtBeginList, 52, 3U},
                                 {kBrtBeginList, 56, 4U}});
  // BrtBeginListCol: id, ilta, header / data / totals dxf ids, qsi, strings.
  p.clear();
  for (const std::uint32_t v : {1U, 0U, 5U, 6U, kNoDxf, 0U}) {
    emit_u32(p, v);
  }
  EmitNullStrings(p, 6);
  emit_record(s, kBrtBeginListCol, p);
  f.slots.insert(f.slots.end(), {{kBrtBeginListCol, 8, 5U}, {kBrtBeginListCol, 12, 6U}});
  // Legacy BrtBeginCFRule: iType, iTemplate, dxfId, ...
  p.clear();
  for (const std::uint32_t v : {1U, 0U, 7U, 1U, 0U, 0U, 0U}) {
    emit_u32(p, v);
  }
  emit_u16(p, 0U);
  EmitNullStrings(p, 1);
  emit_record(s, kBrtBeginCFRule, p);
  f.slots.push_back({kBrtBeginCFRule, 8, 7U});
  // BrtBeginCFRule14 with one FRTFormula (3-byte rgce): dxfId after the list, iType and iTemplate.
  p.clear();
  emit_u32(p, kFrtHeaderFormulas);
  emit_u32(p, 1U);
  emit_u32(p, kFrtItemFlags);
  emit_u32(p, 3U);
  emit_u32(p, 0U);
  for (const std::uint8_t b : std::initializer_list<std::uint8_t>{0x1E, 0x01, 0x00}) {
    emit_u8(p, b);
  }
  for (const std::uint32_t v : {1U, 0U, 8U, 1U}) {
    emit_u32(p, v);
  }
  emit_record(s, kBeginCfRule14, p);
  f.slots.push_back({kBeginCfRule14, 31, 8U});
  // BrtBeginCFRule14 without formulas: dxfId at 12.
  p.clear();
  for (const std::uint32_t v : {0U, 1U, 0U, 9U, 1U}) {
    emit_u32(p, v);
  }
  emit_record(s, kBeginCfRule14, p);
  f.slots.push_back({kBeginCfRule14, 12, 9U});
  // BrtColorFilter: dxfId, fCellColor.
  p.clear();
  emit_u32(p, 10U);
  emit_u32(p, 1U);
  emit_record(s, kBrtColorFilter, p);
  f.slots.push_back({kBrtColorFilter, 0, 10U});
  // BrtBeginSortCond by cell colour, by value and by icon: only the first carries a dxf id.
  for (const std::uint16_t sort_on : std::initializer_list<std::uint16_t>{1, 0, 3, 2}) {
    p.clear();
    emit_u16(p, static_cast<std::uint16_t>(sort_on << 1U));
    emit_rfx(p, MergeRange{1, 0, 3, 0});
    emit_u32(p, sort_on == 2U ? 3U : 2U);
    emit_u32(p, 0U);
    EmitNullStrings(p, 1);
    emit_record(s, kBrtBeginSortCond, p);
    if (sort_on == 1U || sort_on == 2U) {
      f.slots.push_back({kBrtBeginSortCond, 18, sort_on == 2U ? 3U : 2U});
    }
  }
  emit_record(s, kBrtEndSortCond, std::vector<std::uint8_t>{});
  return f;
}

TEST(XlsbDxfRecordSlots, VisitsEachRowsOffset) {
  const Fixture f = BuildFixture();
  auto scan = scan_dxf_slots(f.stream);
  ASSERT_TRUE(scan);
  EXPECT_EQ(scan.value().dxf_records, 2U);
  EXPECT_TRUE(scan.value().undecodable.empty());
  const std::vector<FramedRecord> recs = split_records(f.stream);
  ASSERT_EQ(scan.value().slots.size(), f.slots.size());
  for (std::size_t i = 0; i < f.slots.size(); ++i) {
    const DxfRecordSlot& slot = scan.value().slots[i];
    const auto& [type, offset, value] = f.slots[i];
    EXPECT_EQ(slot.record_type, type) << i;
    EXPECT_EQ(slot.value, value) << i;
    const std::size_t payload_begin = static_cast<std::size_t>(recs[slot.record_index].payload.data - f.stream.data());
    EXPECT_EQ(slot.offset - payload_begin, offset) << i;
  }
}

TEST(XlsbDxfRecordSlots, RewriteMapsEverySlotAndKeepsOtherBytes) {
  const Fixture f = BuildFixture();
  // Model dxf 3 is unused: everything above it moves down one, and its slot is cleared.
  std::vector<DxfUse> uses;
  for (std::uint32_t i = 0; i < 11U; ++i) {
    if (i != 3U) {
      uses.push_back(DxfUse{i, DxfUseKind::kTable, 0, 0, 0});
    }
  }
  const DxfSavePlan plan = build_dxf_save_plan(11, uses, false);
  auto out = rewrite_dxf_slots(f.stream, plan);
  ASSERT_TRUE(out);
  ASSERT_EQ(out.value().size(), f.stream.size());
  auto scan = scan_dxf_slots(f.stream);
  ASSERT_TRUE(scan);
  std::set<std::size_t> slot_bytes;
  for (const DxfRecordSlot& slot : scan.value().slots) {
    const std::uint32_t expected = slot.value == 3U ? kNoDxf : (slot.value < 3U ? slot.value : slot.value - 1U);
    EXPECT_EQ(load_u32(out.value().data() + slot.offset), expected) << slot.value;
    for (std::size_t b = 0; b < 4U; ++b) {
      slot_bytes.insert(slot.offset + b);
    }
  }
  for (std::size_t i = 0; i < f.stream.size(); ++i) {
    if (slot_bytes.count(i) == 0U) {
      EXPECT_EQ(out.value()[i], f.stream[i]) << i;
    }
  }
  auto identity = rewrite_dxf_slots(f.stream, identity_dxf_plan(11));
  ASSERT_TRUE(identity);
  EXPECT_EQ(identity.value(), f.stream);
}

TEST(XlsbDxfRecordSlots, TableStyleElementSlotIsVisitedAndRewritten) {
  std::vector<std::uint8_t> s;
  emit_record(s, kBrtBeginTableStyle, std::vector<std::uint8_t>(12, 0));
  std::vector<std::uint8_t> p;
  for (const std::uint32_t v : {0U, 1U, 2U}) {
    emit_u32(p, v);
  }
  emit_record(s, kBrtTableStyleElement, p);
  emit_record(s, kBrtEndTableStyle, std::vector<std::uint8_t>{});
  auto scan = scan_dxf_slots(s);
  ASSERT_TRUE(scan);
  EXPECT_TRUE(scan.value().undecodable.empty());
  ASSERT_EQ(scan.value().slots.size(), 1U);
  const DxfRecordSlot& slot = scan.value().slots[0];
  EXPECT_EQ(slot.record_type, kBrtTableStyleElement);
  EXPECT_EQ(slot.value, 2U);
  const std::vector<FramedRecord> recs = split_records(s);
  EXPECT_EQ(slot.offset - static_cast<std::size_t>(recs[slot.record_index].payload.data - s.data()), 8U);
  auto out = rewrite_dxf_slots(s, build_dxf_save_plan(3, {DxfUse{2, DxfUseKind::kTableStyleElement, -1, 0, 0}}, false));
  ASSERT_TRUE(out);
  std::vector<std::uint8_t> expected = s;
  store_u32(expected.data() + slot.offset, 0U);
  EXPECT_EQ(out.value(), expected);
}

TEST(XlsbDxfRecordSlots, OutOfRangeSlotBecomesNoDxf) {
  std::vector<std::uint8_t> s;
  std::vector<std::uint8_t> p;
  emit_u32(p, 5U);
  emit_u32(p, 0U);
  emit_record(s, kBrtColorFilter, p);
  auto out = rewrite_dxf_slots(s, identity_dxf_plan(2));
  ASSERT_TRUE(out);
  auto scan = scan_dxf_slots(out.value());
  ASSERT_TRUE(scan);
  EXPECT_TRUE(scan.value().slots.empty());
}

TEST(XlsbDxfRecordSlots, SlotBearingRecordsAreDecodable) {
  for (const std::uint16_t type : {kBrtBeginList, kBrtBeginListCol, kBrtBeginCFRule, kBeginCfRule14, kBrtColorFilter,
                                   kBrtBeginSortCond, kBrtTableStyleElement, kBrtDxf}) {
    EXPECT_FALSE(is_undecodable_dxf_record(type)) << type;
  }
}

TEST(XlsbDxfRecordSlots, UnlocatableSlotIsAnError) {
  const auto one = [](std::uint16_t type, std::vector<std::uint8_t> payload) {
    std::vector<std::uint8_t> s;
    emit_record(s, type, payload);
    return s;
  };
  EXPECT_FALSE(scan_dxf_slots(one(kBrtBeginList, std::vector<std::uint8_t>(56, 0))));
  EXPECT_FALSE(scan_dxf_slots(one(kBrtBeginListCol, std::vector<std::uint8_t>(16, 0))));
  EXPECT_FALSE(scan_dxf_slots(one(kBrtColorFilter, std::vector<std::uint8_t>(3, 0))));
  std::vector<std::uint8_t> sqref_header(20, 0);
  sqref_header[0] = static_cast<std::uint8_t>(kFrtHeaderSqrefs);
  EXPECT_FALSE(scan_dxf_slots(one(kBeginCfRule14, sqref_header)));
  std::vector<std::uint8_t> bad_sort(30, 0);
  bad_sort[0] = 5U << 1U;
  EXPECT_FALSE(scan_dxf_slots(one(kBrtBeginSortCond, bad_sort)));
  std::vector<std::uint8_t> torn = one(kBrtColorFilter, std::vector<std::uint8_t>(8, 0));
  torn.pop_back();
  EXPECT_FALSE(scan_dxf_slots(torn));
  EXPECT_FALSE(rewrite_dxf_slots(torn, identity_dxf_plan(1)));
  auto empty = scan_dxf_slots({});
  ASSERT_TRUE(empty);
  EXPECT_TRUE(empty.value().slots.empty());
}

}  // namespace
}  // namespace xlsb
}  // namespace formulon::io
