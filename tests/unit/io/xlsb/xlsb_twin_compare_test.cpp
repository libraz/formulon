// Self-tests of the twin comparison helper: equality, mask application, mask construction from two Excel saves,
// and the uninitialised-field hook.

#include "xlsb_twin_compare.h"

#include <cstdint>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb::test {
namespace {

const std::string kWorkbookPart = "xl/workbook.bin";

Bytes NotesBase(const char* suffix) {
  const Bytes bytes = ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/notes/base" + suffix);
  EXPECT_FALSE(bytes.empty()) << suffix;
  return bytes;
}

/// `package` with byte 0 of the first non-empty record of the workbook part flipped; `index` receives the record.
Bytes FlipFirstPayloadByte(const Bytes& package, std::size_t* index) {
  const std::optional<Bytes> part = ReadPart(package, kWorkbookPart);
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return {};
  }
  const std::vector<FramedRecord> records = split_records(*part);
  for (std::size_t i = 0; i < records.size(); ++i) {
    if (records[i].payload.size == 0U) {
      continue;
    }
    *index = i;
    Bytes patched = *part;
    patched[static_cast<std::size_t>(records[i].payload.data - part->data())] ^= 0xFFU;
    return RepackWithEntries(package, {{kWorkbookPart, patched}});
  }
  ADD_FAILURE() << "no record with a payload in " << kWorkbookPart;
  return {};
}

TEST(XlsbTwinCompare, IdenticalInputsCompareEqual) {
  const Bytes base = NotesBase(".xlsb");
  EXPECT_FALSE(ExtractRecords(base, {}).empty());
  EXPECT_TRUE(CompareRecordSets(base, base, {}, ByteMask{}).empty());
}

TEST(XlsbTwinCompare, MaskOfOneSaveAgainstItselfIsEmpty) {
  const Bytes base = NotesBase(".xlsb");
  EXPECT_TRUE(BuildTwinMask(base, base).empty());
}

TEST(XlsbTwinCompare, UnmaskedByteIsReported) {
  const Bytes base = NotesBase(".xlsb");
  std::size_t index = 0;
  const Bytes patched = FlipFirstPayloadByte(base, &index);
  ASSERT_FALSE(patched.empty());
  const std::vector<std::string> mismatches = CompareRecordSets(patched, base, {}, ByteMask{});
  ASSERT_EQ(mismatches.size(), 1U);
  EXPECT_NE(mismatches[0].find(kWorkbookPart + " record #" + std::to_string(index) + " "), std::string::npos)
      << mismatches[0];
  EXPECT_NE(mismatches[0].find("offset 0:"), std::string::npos) << mismatches[0];
}

TEST(XlsbTwinCompare, MaskedByteIsIgnored) {
  const Bytes base = NotesBase(".xlsb");
  std::size_t index = 0;
  const Bytes patched = FlipFirstPayloadByte(base, &index);
  ASSERT_FALSE(patched.empty());
  ByteMask by_hand;
  by_hand.bytes[kWorkbookPart][index].insert(0U);
  EXPECT_TRUE(CompareRecordSets(patched, base, {}, by_hand).empty());

  const ByteMask built = BuildTwinMask(base, patched);
  EXPECT_EQ(built, by_hand);
  EXPECT_EQ(built.ByteCount(), 1U);
  EXPECT_TRUE(CompareRecordSets(patched, base, {}, built).empty());
}

TEST(XlsbTwinCompare, TypeFilterSkipsOtherRecords) {
  const Bytes base = NotesBase(".xlsb");
  std::size_t index = 0;
  const Bytes patched = FlipFirstPayloadByte(base, &index);
  ASSERT_FALSE(patched.empty());
  const std::optional<Bytes> part = ReadPart(base, kWorkbookPart);
  ASSERT_TRUE(part.has_value());
  const std::uint16_t flipped_type = SplitPart(*part)[index].type;
  std::set<std::uint16_t> others;
  for (const IndexedRecord& rec : SplitPart(*part)) {
    if (rec.type != flipped_type) {
      others.insert(rec.type);
    }
  }
  ASSERT_FALSE(others.empty());
  EXPECT_TRUE(CompareRecordSets(patched, base, others, ByteMask{}).empty());
  EXPECT_FALSE(CompareRecordSets(patched, base, {flipped_type}, ByteMask{}).empty());
}

TEST(XlsbTwinCompare, NotesTwinMaskIsDeterministic) {
  const Bytes first = NotesBase(".xlsb");
  const Bytes second = NotesBase(".2.xlsb");
  const ByteMask mask = BuildTwinMask(first, second);
  EXPECT_EQ(BuildTwinMask(first, second), mask);
  EXPECT_EQ(BuildTwinMask(second, first), mask);
  RecordProperty("masked_bytes", static_cast<int>(mask.ByteCount()));
  RecordProperty("masked_whole_parts", static_cast<int>(mask.whole_parts.size()));

  const bool records_differ = !CompareRecordSets(second, first, {}, ByteMask{}).empty();
  std::size_t masked_record_parts = 0;
  for (const std::string& part : mask.whole_parts) {
    masked_record_parts += IsRecordPart(part) ? 1U : 0U;
  }
  EXPECT_EQ(records_differ, mask.ByteCount() + mask.whole_records.size() + masked_record_parts > 0U);
  EXPECT_TRUE(CompareRecordSets(second, first, {}, mask).empty());
}

TEST(XlsbTwinCompare, LeadingAttrSemiDataWordIsMasked) {
  // BrtFmlaNum: cell header, IEEE value, grbit, cce = 4, rgce = PtgAttrSemi + data word, cb = 0.
  Bytes payload(8U + 8U + 2U, 0U);
  const Bytes formula = {0x04, 0x00, 0x00, 0x00, 0x19, 0x01, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00};
  payload.insert(payload.end(), formula.begin(), formula.end());
  std::vector<std::size_t> offsets;
  MaskLeadingAttrSemi(9U, payload, &offsets);
  EXPECT_EQ(offsets, (std::vector<std::size_t>{24U, 25U}));

  offsets.clear();
  payload[22] = 0x1E;  // PtgInt: no data word to mask.
  MaskLeadingAttrSemi(9U, payload, &offsets);
  EXPECT_TRUE(offsets.empty());

  offsets.clear();
  MaskLeadingAttrSemi(7U, payload, &offsets);
  EXPECT_TRUE(offsets.empty());
}

}  // namespace
}  // namespace formulon::io::xlsb::test
