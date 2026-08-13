//
// Integration test: a workbook carrying `<externalReferences>` plus
// the matching `xl/_rels/workbook.xml.rels` and per-link rels files
// must survive a full writer -> reader cycle without losing the
// references. The body parts themselves ride through passthrough.

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/external_links.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// Builds a minimal externalLink body part. The content is what Excel
// itself emits for a freshly-linked workbook with one cached sheet
// reference; the test only asserts on the metadata round-trip, so the
// exact body bytes are mostly there to keep Excel and our reader from
// rejecting the package.
std::vector<std::uint8_t> MakeExternalBookBody(const std::string& body_rid) {
  std::string s;
  s.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
  s.append(
      "<externalLink "
      "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  s.append("  <externalBook r:id=\"");
  s.append(body_rid);
  s.append("\"><sheetNames><sheetName val=\"Sheet1\"/></sheetNames></externalBook>\n");
  s.append("</externalLink>\n");
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

// The miniz writer has no password/encryption mode. Marking one entry's
// general-purpose bit in both the local and central headers gives ZipReader
// its existing deterministic encrypted-entry failure without changing the
// ZIP reader or relying on any aggregate archive cap.
std::size_t MarkZipEntryEncrypted(std::vector<std::uint8_t>& bytes, std::string_view entry_name) {
  const auto read16 = [&bytes](std::size_t offset) -> std::uint16_t {
    return static_cast<std::uint16_t>(bytes[offset]) |
           static_cast<std::uint16_t>(static_cast<std::uint16_t>(bytes[offset + 1U]) << 8U);
  };
  std::size_t matches = 0;
  for (std::size_t i = 0; i + 4U <= bytes.size(); ++i) {
    const bool local = bytes[i] == 0x50U && bytes[i + 1U] == 0x4bU && bytes[i + 2U] == 0x03U && bytes[i + 3U] == 0x04U;
    const bool central =
        bytes[i] == 0x50U && bytes[i + 1U] == 0x4bU && bytes[i + 2U] == 0x01U && bytes[i + 3U] == 0x02U;
    if (!local && !central) {
      continue;
    }
    const std::size_t flag_offset = i + (local ? 6U : 8U);
    const std::size_t name_length_offset = i + (local ? 26U : 28U);
    const std::size_t name_offset = i + (local ? 30U : 46U);
    if (name_offset > bytes.size() || name_length_offset + 2U > bytes.size()) {
      continue;
    }
    const std::size_t name_length = read16(name_length_offset);
    if (name_offset + name_length > bytes.size()) {
      continue;
    }
    const std::string_view actual_name(reinterpret_cast<const char*>(bytes.data() + name_offset), name_length);
    if (actual_name != entry_name) {
      continue;
    }
    bytes[flag_offset] = static_cast<std::uint8_t>(bytes[flag_offset] | 0x01U);
    ++matches;
  }
  return matches;
}

Workbook MakeSingleExternalLinkWorkbook(std::vector<std::uint8_t> body_bytes) {
  Workbook src = Workbook::create();

  io::PassthroughPart body;
  body.path = "xl/externalLinks/externalLink1.xml";
  body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
  body.bytes = std::move(body_bytes);
  std::vector<io::PassthroughPart> parts;
  parts.push_back(std::move(body));
  src.set_passthrough_parts(std::move(parts));

  io::ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId7";
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "file:///Users/example/RemoteBook.xlsx";
  rec.target_external = true;
  rec.kind = io::ExternalLinkRecord::Kind::kExternalBook;
  std::vector<io::ExternalLinkRecord> links;
  links.push_back(std::move(rec));
  src.set_external_links(std::move(links));
  return src;
}

TEST(ExternalLinksRoundTrip, PreservesSingleExternalBook) {
  Workbook src = Workbook::create();

  // 1. Body part lives in passthrough; it round-trips verbatim.
  io::PassthroughPart body;
  body.path = "xl/externalLinks/externalLink1.xml";
  body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
  body.bytes = MakeExternalBookBody("rId1");

  std::vector<io::PassthroughPart> parts;
  parts.push_back(std::move(body));
  src.set_passthrough_parts(std::move(parts));

  // 2. Metadata: one external book pointing at a remote file URL.
  io::ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId7";  // Reader replaces this on round-trip; value here is irrelevant.
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "file:///Users/example/RemoteBook.xlsx";
  rec.target_external = true;
  rec.kind = io::ExternalLinkRecord::Kind::kExternalBook;
  std::vector<io::ExternalLinkRecord> links;
  links.push_back(std::move(rec));
  src.set_external_links(std::move(links));

  // 3. Save / reload.
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  // 4. Assert the metadata survives.
  const auto& rt = load_or.value().workbook.external_links();
  ASSERT_EQ(rt.size(), 1U);
  EXPECT_EQ(rt[0].index, 1U);
  EXPECT_EQ(rt[0].part_path, "xl/externalLinks/externalLink1.xml");
  EXPECT_EQ(rt[0].body_rel_id, "rId1");
  EXPECT_EQ(rt[0].target, "file:///Users/example/RemoteBook.xlsx");
  EXPECT_TRUE(rt[0].target_external);
  EXPECT_EQ(rt[0].kind, io::ExternalLinkRecord::Kind::kExternalBook);
  // The body part must still be present in passthrough — the reader
  // should have left it alone (it consumes only the per-link rels file).
  bool body_present = false;
  for (const io::PassthroughPart& p : load_or.value().workbook.passthrough_parts()) {
    if (p.path == "xl/externalLinks/externalLink1.xml") {
      body_present = true;
      break;
    }
  }
  EXPECT_TRUE(body_present);
}

TEST(ExternalLinksRoundTrip, PreservesDocumentOrderAcrossMultipleLinks) {
  Workbook src = Workbook::create();

  std::vector<io::PassthroughPart> parts;
  for (int i = 1; i <= 3; ++i) {
    io::PassthroughPart body;
    body.path = "xl/externalLinks/externalLink" + std::to_string(i) + ".xml";
    body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    body.bytes = MakeExternalBookBody("rId1");
    parts.push_back(std::move(body));
  }
  src.set_passthrough_parts(std::move(parts));

  std::vector<io::ExternalLinkRecord> links;
  for (int i = 1; i <= 3; ++i) {
    io::ExternalLinkRecord rec;
    rec.index = static_cast<std::uint32_t>(i);
    rec.rel_id = "rId" + std::to_string(100 + i);
    rec.part_path = "xl/externalLinks/externalLink" + std::to_string(i) + ".xml";
    rec.body_rel_id = "rId1";
    rec.target = "file:///remote/book" + std::to_string(i) + ".xlsx";
    rec.target_external = true;
    rec.kind = io::ExternalLinkRecord::Kind::kExternalBook;
    links.push_back(std::move(rec));
  }
  src.set_external_links(std::move(links));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));

  const auto& rt = load_or.value().workbook.external_links();
  ASSERT_EQ(rt.size(), 3U);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(rt[i].index, i + 1U);
    EXPECT_EQ(rt[i].part_path, "xl/externalLinks/externalLink" + std::to_string(i + 1) + ".xml");
    EXPECT_EQ(rt[i].target, "file:///remote/book" + std::to_string(i + 1) + ".xlsx");
    EXPECT_EQ(rt[i].kind, io::ExternalLinkRecord::Kind::kExternalBook);
  }
}

TEST(ExternalLinksRoundTrip, SuccessfullyReadMalformedBodyRemainsUnknown) {
  // XML parsing is intentionally best-effort for external-link metadata. A
  // body that extracts successfully but is malformed must still leave the
  // workbook load usable and classify the record as kUnknown.
  Workbook src = MakeSingleExternalLinkWorkbook({'<', 'e', 'x', 't', 'e', 'r', 'n', 'a', 'l', 'L', 'i', 'n', 'k'});
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or))
      << "malformed external-link XML must remain compatible: " << load_or.error().message;
  const auto& links = load_or.value().workbook.external_links();
  ASSERT_EQ(links.size(), 1U);
  EXPECT_EQ(links[0].kind, io::ExternalLinkRecord::Kind::kUnknown);
  EXPECT_EQ(links[0].target, "file:///Users/example/RemoteBook.xlsx");
}

TEST(ExternalLinksRoundTrip, BodyReadFailurePropagatesUnchanged) {
  Workbook src = MakeSingleExternalLinkWorkbook(MakeExternalBookBody("rId1"));
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  std::vector<std::uint8_t> bytes = std::move(save_or.value());
  EXPECT_EQ(MarkZipEntryEncrypted(bytes, "xl/externalLinks/externalLink1.xml"), 2U);

  auto load_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_FALSE(static_cast<bool>(load_or));
  EXPECT_EQ(load_or.error().code, FormulonErrorCode::kIoZipEncrypted);
  EXPECT_NE(load_or.error().context.find("entry=xl/externalLinks/externalLink1.xml"), std::string::npos);
}

TEST(ExternalLinksRoundTrip, RelsReadFailurePropagatesUnchanged) {
  Workbook src = MakeSingleExternalLinkWorkbook(MakeExternalBookBody("rId1"));
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  std::vector<std::uint8_t> bytes = std::move(save_or.value());
  EXPECT_EQ(MarkZipEntryEncrypted(bytes, "xl/externalLinks/_rels/externalLink1.xml.rels"), 2U);

  auto load_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_FALSE(static_cast<bool>(load_or));
  EXPECT_EQ(load_or.error().code, FormulonErrorCode::kIoZipEncrypted);
  EXPECT_NE(load_or.error().context.find("entry=xl/externalLinks/_rels/externalLink1.xml.rels"), std::string::npos);
}

TEST(ExternalLinksRoundTrip, EmptyWorkbookEmitsNoExternalReferencesBlock) {
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  EXPECT_TRUE(load_or.value().workbook.external_links().empty());
}

}  // namespace
}  // namespace formulon
