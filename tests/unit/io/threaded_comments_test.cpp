#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <ctime>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/threaded_comments_io.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kThread = "{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}";
constexpr std::string_view kReply = "{8E1F3C55-0A9B-4D27-B6E4-31C7D2A9F0E8}";
constexpr std::string_view kMention = "{5C6D7E8F-9A0B-4C1D-8E2F-3A4B5C6D7E8F}";
constexpr std::string_view kAlice = "{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}";
constexpr std::string_view kBob = "{22222222-AAAA-4BBB-8CCC-EEEEEEEEEEEE}";
constexpr std::string_view kNoteUid = "{0F0E0D0C-0B0A-4908-8706-050403020100}";

constexpr std::string_view kContentTypes =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>)";

constexpr std::string_view kPackageRels =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>)";

constexpr std::string_view kWorkbookXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>)";

constexpr std::string_view kWorkbookRels =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>)";

constexpr std::string_view kSheetXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rId1"/></worksheet>)";

constexpr std::string_view kSheetRels =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>)";

// A1 carries a plain note; B2 carries a resolved thread with one reply that
// mentions Alice. The thread's legacy stub sits beside the note.
constexpr std::string_view kCommentsXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors><author>Carol</author><author>tc={2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}</author></authors><commentList><comment ref="A1" authorId="0" shapeId="0" xr:uid="{0F0E0D0C-0B0A-4908-8706-050403020100}"><text><r><t>Plain note</t></r></text></comment><comment ref="B2" authorId="1" shapeId="0" xr:uid="{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}"><text><t>[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    Check this total
Reply:
    @Alice Smith fixed</t></text></comment></commentList></comments>)";

constexpr std::string_view kThreadedXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><threadedComment ref="B2" dT="2026-09-30T08:15:42.17" personId="{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}" id="{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}" done="1"><text>Check this total</text></threadedComment><threadedComment ref="B2" dT="2026-09-30T09:02:05.80" personId="{22222222-AAAA-4BBB-8CCC-EEEEEEEEEEEE}" id="{8E1F3C55-0A9B-4D27-B6E4-31C7D2A9F0E8}" parentId="{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}"><text>@Alice Smith fixed</text><mentions><mention mentionpersonId="{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}" mentionId="{5C6D7E8F-9A0B-4C1D-8E2F-3A4B5C6D7E8F}" startIndex="0" length="12"/></mentions></threadedComment></ThreadedComments>)";

constexpr std::string_view kPersonsXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><person displayName="Alice Smith" id="{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}" userId="alice@example.com" providerId="AD"/><person displayName="Bob Jones" id="{22222222-AAAA-4BBB-8CCC-EEEEEEEEEEEE}" userId="Bob Jones" providerId="None"/></personList>)";

// Source VML with one shape per commented cell; `alt="source-vml"` marks the
// bytes so a test can tell verbatim re-emission from regeneration.
constexpr std::string_view kVml = R"(<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
 <v:shape id="_x0000_s1025" type="#_x0000_t202" alt="source-vml" style="position:absolute;visibility:hidden" fillcolor="#ffffe1"><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t202" alt="source-vml" style="position:absolute;visibility:hidden" fillcolor="#ffffe1"><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
</xml>)";

std::vector<std::uint8_t> BuildZip(const std::vector<std::pair<const char*, std::string_view>>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.first, p.second.data(), p.second.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

std::vector<std::uint8_t> ExcelFixture() {
  return BuildZip({{"[Content_Types].xml", kContentTypes},
                   {"_rels/.rels", kPackageRels},
                   {"xl/workbook.xml", kWorkbookXml},
                   {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                   {"xl/worksheets/sheet1.xml", kSheetXml},
                   {"xl/worksheets/_rels/sheet1.xml.rels", kSheetRels},
                   {"xl/comments1.xml", kCommentsXml},
                   {"xl/threadedComments/threadedComment1.xml", kThreadedXml},
                   {"xl/persons/person.xml", kPersonsXml},
                   {"xl/drawings/vmlDrawing1.vml", kVml}});
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << read.error().message;
  return std::move(read.value().workbook);
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = wb.save();
  EXPECT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  return std::move(saved.value());
}

/// Text of the package part at `path`, or "" when absent.
std::string Part(const std::vector<std::uint8_t>& package, const std::string& path) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{package.data(), package.size()})));
  if (!zip.has_entry(path)) {
    return {};
  }
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes));
  return std::string(bytes.value().begin(), bytes.value().end());
}

std::size_t Count(std::string_view haystack, std::string_view needle) {
  std::size_t n = 0;
  for (std::size_t at = haystack.find(needle); at != std::string_view::npos; at = haystack.find(needle, at + 1)) {
    ++n;
  }
  return n;
}

ThreadedComment Comment(std::string_view id, std::uint32_t row, std::uint32_t col, std::string_view person,
                        std::string text, std::string_view parent = {}) {
  ThreadedComment c;
  c.id = std::string(id);
  c.row = row;
  c.col = col;
  c.person_id = std::string(person);
  c.created = "2026-10-05T10:00:00.00";
  c.text = std::move(text);
  c.parent_id = std::string(parent);
  return c;
}

Workbook WorkbookWithPersons() {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_person(Person{std::string(kAlice), "Alice Smith", "alice@example.com", "AD"})));
  EXPECT_TRUE(static_cast<bool>(wb.add_person(Person{std::string(kBob), "Bob Jones", "Bob Jones", "None"})));
  return wb;
}

void ExpectFixtureModel(const Workbook& wb) {
  ASSERT_EQ(wb.persons().size(), 2U);
  EXPECT_EQ(wb.persons()[0].id, kAlice);
  EXPECT_EQ(wb.persons()[0].display_name, "Alice Smith");
  EXPECT_EQ(wb.persons()[0].user_id, "alice@example.com");
  EXPECT_EQ(wb.persons()[0].provider_id, "AD");
  EXPECT_EQ(wb.persons()[1].provider_id, "None");

  const std::vector<ThreadedComment>& threads = wb.sheet(0).threaded_comments();
  ASSERT_EQ(threads.size(), 2U);
  EXPECT_EQ(threads[0].id, kThread);
  EXPECT_EQ(threads[0].row, 1U);
  EXPECT_EQ(threads[0].col, 1U);
  EXPECT_EQ(threads[0].person_id, kAlice);
  EXPECT_EQ(threads[0].created, "2026-09-30T08:15:42.17");
  EXPECT_EQ(threads[0].text, "Check this total");
  EXPECT_TRUE(threads[0].parent_id.empty());
  EXPECT_TRUE(threads[0].done);
  EXPECT_TRUE(threads[0].mentions.empty());

  EXPECT_EQ(threads[1].id, kReply);
  EXPECT_EQ(threads[1].parent_id, kThread);
  EXPECT_EQ(threads[1].person_id, kBob);
  EXPECT_EQ(threads[1].row, 1U);
  EXPECT_EQ(threads[1].col, 1U);
  EXPECT_EQ(threads[1].text, "@Alice Smith fixed");
  ASSERT_EQ(threads[1].mentions.size(), 1U);
  EXPECT_EQ(threads[1].mentions[0].person_id, kAlice);
  EXPECT_EQ(threads[1].mentions[0].mention_id, kMention);
  EXPECT_EQ(threads[1].mentions[0].start, 0U);
  EXPECT_EQ(threads[1].mentions[0].length, 12U);
}

TEST(ThreadedComments, RoundTripRepliesMentions) {
  Workbook wb = Load(ExcelFixture());
  ExpectFixtureModel(wb);

  const std::vector<std::uint8_t> saved = Save(wb);
  const std::string content_types = Part(saved, "[Content_Types].xml");
  EXPECT_NE(content_types.find("application/vnd.ms-excel.threadedcomments+xml"), std::string::npos);
  EXPECT_NE(content_types.find("application/vnd.ms-excel.person+xml"), std::string::npos);
  EXPECT_NE(Part(saved, "xl/_rels/workbook.xml.rels").find("/relationships/person"), std::string::npos);
  EXPECT_NE(Part(saved, "xl/worksheets/_rels/sheet1.xml.rels").find("/relationships/threadedComment"),
            std::string::npos);
  // The persons and threaded parts are model-owned, so they are not listed
  // as passthrough after a load.
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    EXPECT_EQ(part.path.find("threadedComments"), std::string::npos) << part.path;
    EXPECT_EQ(part.path.find("persons"), std::string::npos) << part.path;
  }

  Workbook reloaded = Load(saved);
  ExpectFixtureModel(reloaded);
  ASSERT_EQ(reloaded.sheet(0).comments().size(), 1U);
  EXPECT_EQ(reloaded.sheet(0).comments()[0].text, "Plain note");
}

TEST(ThreadedComments, CommentWithNoTimeOmitsDtAndRoundTrips) {
  // A comment that carries no time, as a loaded file can, is written without `dT` and reads back without one.
  ThreadedComment c = Comment(kThread, 0, 0, kAlice, "undated");
  c.created.clear();
  const std::string xml = io::write_threaded_comments({c});
  EXPECT_EQ(xml.find("dT="), std::string::npos);
  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto decoded = io::read_threaded_comments(bytes);
  ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
  ASSERT_EQ(decoded.value().size(), 1U);
  EXPECT_TRUE(decoded.value()[0].created.empty());
}

TEST(ThreadedComments, AddedCommentWithNoTimeIsStamped) {
  Workbook wb = WorkbookWithPersons();
  ThreadedComment c = Comment(kThread, 0, 0, kAlice, "undated");
  c.created.clear();
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, c)));
  Workbook reloaded = Load(Save(wb));
  ASSERT_EQ(reloaded.sheet(0).threaded_comments().size(), 1U);
  EXPECT_FALSE(reloaded.sheet(0).threaded_comments()[0].created.empty());
}

TEST(ThreadedComments, TextSurvivesTwoOoxmlCycles) {
  const std::string payload = "literal _x0041_ & <猫> \rcontrol \x01 🙂";
  std::vector<ThreadedComment> current = {Comment(kThread, 0, 0, kAlice, payload)};

  for (int cycle = 0; cycle < 2; ++cycle) {
    const std::string xml = io::write_threaded_comments(current);
    std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
    auto decoded = io::read_threaded_comments(bytes);
    ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
    ASSERT_EQ(decoded.value().size(), 1U);
    current = std::move(decoded.value());
  }

  EXPECT_EQ(current[0].text, payload);
}

TEST(ThreadedComments, StubNotListedAsNote) {
  Workbook wb = Load(ExcelFixture());
  ASSERT_EQ(wb.sheet(0).comments().size(), 1U);
  EXPECT_EQ(wb.sheet(0).comments()[0].row, 0U);
  EXPECT_EQ(wb.sheet(0).comments()[0].author, "Carol");

  // The saved comments part still carries Excel's stub for older readers,
  // keyed to the thread by author and `xr:uid`.
  const std::string comments = Part(Save(wb), "xl/comments1.xml");
  EXPECT_NE(comments.find("<author>tc={2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}</author>"), std::string::npos);
  EXPECT_NE(comments.find("ref=\"B2\""), std::string::npos);
  EXPECT_NE(comments.find("xr:uid=\"{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23}\""), std::string::npos);
  EXPECT_NE(comments.find("[Threaded comment]"), std::string::npos);
  EXPECT_NE(comments.find("Check this total"), std::string::npos);
  EXPECT_NE(comments.find("Reply:\n    @Alice Smith fixed"), std::string::npos);
}

TEST(ThreadedComments, TcAuthorWithoutThreadStaysNote) {
  std::vector<CellComment> notes = {CellComment(1, 1, "tc=" + std::string(kThread), "orphan stub")};
  io::drop_thread_stubs(notes, {});
  EXPECT_EQ(notes.size(), 1U);
  std::vector<ThreadedComment> elsewhere = {Comment(kThread, 5, 5, kAlice, "x")};
  io::drop_thread_stubs(notes, elsewhere);
  EXPECT_EQ(notes.size(), 1U);
}

TEST(ThreadedComments, BatchStubIndexesPreserveFirstMatchesAndSourceOrder) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<ThreadedComment>& threads = sheet.mutable_threaded_comments();
  threads = {Comment("duplicate", 1U, 1U, kAlice, "first root"),
             Comment("reply-a", 9U, 9U, kBob, "reply one", "duplicate"),
             Comment("duplicate", 2U, 2U, kAlice, "second root"),
             Comment("unique", 3U, 3U, kBob, "unique root"),
             Comment("reply-b", 8U, 8U, kBob, "reply two", "duplicate"),
             Comment("reply-unique", 3U, 3U, kAlice, "unique reply", "unique"),
             Comment("", 10U, 10U, kAlice, "empty-id root")};
  sheet.mutable_comments() = {
      CellComment(1U, 1U, "Carol", "root collision one"),   CellComment(2U, 2U, "Carol", "root collision two"),
      CellComment(3U, 3U, "Carol", "root collision three"), CellComment(10U, 10U, "Carol", "empty-id collision"),
      CellComment(4U, 4U, "tc=missing", "unknown stub"),    CellComment(9U, 9U, "Carol", "reply-cell note"),
      CellComment(5U, 5U, "Carol", "ordinary note")};

  const std::vector<CellComment> generated = io::legacy_comments_with_stubs(sheet);
  ASSERT_EQ(generated.size(), 7U);
  EXPECT_EQ(generated[0].text, "unknown stub");
  EXPECT_EQ(generated[1].text, "reply-cell note");
  EXPECT_EQ(generated[2].text, "ordinary note");
  EXPECT_EQ(generated[3].author, "tc=duplicate");
  EXPECT_EQ(generated[4].author, "tc=duplicate");
  EXPECT_EQ(generated[5].author, "tc=unique");
  EXPECT_EQ(generated[6].author, "tc=");
  EXPECT_EQ(generated[3].text, threaded_comment_stub_text(threads, 0U));
  EXPECT_EQ(generated[4].text, threaded_comment_stub_text(threads, 2U));
  EXPECT_EQ(generated[5].text, threaded_comment_stub_text(threads, 3U));
  EXPECT_EQ(generated[6].text, threaded_comment_stub_text(threads, 6U));
  const std::size_t first_reply = generated[3].text.find("Reply:\n    reply one");
  const std::size_t second_reply = generated[3].text.find("Reply:\n    reply two");
  EXPECT_NE(first_reply, std::string::npos);
  EXPECT_NE(second_reply, std::string::npos);
  EXPECT_LT(first_reply, second_reply);

  std::vector<CellComment> notes = {CellComment(1U, 1U, "tc=duplicate", "first stub"),
                                    CellComment(2U, 2U, "tc=duplicate", "duplicate id at later root"),
                                    CellComment(3U, 3U, "tc=unique", "unique stub"),
                                    CellComment(10U, 10U, "tc=", "empty id stub"),
                                    CellComment(9U, 9U, "tc=reply-a", "reply is not a root"),
                                    CellComment(4U, 4U, "tc=missing", "unknown stub"),
                                    CellComment(5U, 5U, "Carol", "plain note")};
  io::drop_thread_stubs(notes, threads);
  ASSERT_EQ(notes.size(), 4U);
  EXPECT_EQ(notes[0].text, "duplicate id at later root");
  EXPECT_EQ(notes[1].text, "reply is not a root");
  EXPECT_EQ(notes[2].text, "unknown stub");
  EXPECT_EQ(notes[3].text, "plain note");
}

TEST(ThreadedComments, LegacyNoteUidPreserved) {
  Workbook wb = Load(ExcelFixture());
  ASSERT_EQ(wb.sheet(0).comments().size(), 1U);
  EXPECT_EQ(wb.sheet(0).comments()[0].uid, kNoteUid);

  const std::vector<std::uint8_t> saved = Save(wb);
  const std::string comments = Part(saved, "xl/comments1.xml");
  EXPECT_NE(comments.find("xr:uid=\"{0F0E0D0C-0B0A-4908-8706-050403020100}\""), std::string::npos);
  EXPECT_NE(comments.find("xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""),
            std::string::npos);
  Workbook reloaded = Load(saved);
  ASSERT_EQ(reloaded.sheet(0).comments().size(), 1U);
  EXPECT_EQ(reloaded.sheet(0).comments()[0].uid, kNoteUid);
}

TEST(ThreadedComments, VmlRegeneratedWithPassthroughNotes) {
  {
    // Nothing changed: the source VML goes out byte for byte.
    Workbook wb = Load(ExcelFixture());
    EXPECT_EQ(Part(Save(wb), "xl/drawings/vmlDrawing1.vml"), kVml);
  }
  Workbook wb = Load(ExcelFixture());
  ASSERT_TRUE(static_cast<bool>(
      wb.add_threaded_comment(0, Comment("{33333333-1111-4222-8333-444444444444}", 2, 2, kBob, "New thread"))));
  const std::vector<std::uint8_t> saved = Save(wb);
  const std::string vml = Part(saved, "xl/drawings/vmlDrawing1.vml");
  // The plain note keeps its own shape; the thread stubs get the default one.
  EXPECT_EQ(Count(vml, "source-vml"), 1U);
  EXPECT_EQ(Count(vml, "ObjectType=\"Note\""), 3U);
  EXPECT_NE(vml.find("<x:Row>0</x:Row><x:Column>0</x:Column>"), std::string::npos);
  EXPECT_NE(vml.find("<x:Row>1</x:Row>\n   <x:Column>1</x:Column>"), std::string::npos);
  EXPECT_NE(vml.find("<x:Row>2</x:Row>\n   <x:Column>2</x:Column>"), std::string::npos);

  // The passthrough note, the source thread and the new thread all survive.
  Workbook reloaded = Load(saved);
  ASSERT_EQ(reloaded.sheet(0).comments().size(), 1U);
  EXPECT_EQ(reloaded.sheet(0).comments()[0].text, "Plain note");
  EXPECT_EQ(reloaded.sheet(0).threaded_comments().size(), 3U);
  EXPECT_EQ(Count(Part(saved, "xl/comments1.xml"), "<comment "), 3U);
}

TEST(ThreadedComments, CreateWithoutPersonsPartGeneratesParts) {
  Workbook wb = WorkbookWithPersons();
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kThread, 0, 3, kAlice, "Hello"))));
  const std::vector<std::uint8_t> saved = Save(wb);

  EXPECT_NE(Part(saved, "xl/persons/person.xml").find("displayName=\"Alice Smith\""), std::string::npos);
  EXPECT_NE(Part(saved, "xl/threadedComments/threadedComment1.xml").find("ref=\"D1\""), std::string::npos);
  const std::string content_types = Part(saved, "[Content_Types].xml");
  EXPECT_NE(content_types.find("PartName=\"/xl/persons/person.xml\" "
                               "ContentType=\"application/vnd.ms-excel.person+xml\""),
            std::string::npos);
  EXPECT_NE(content_types.find("PartName=\"/xl/threadedComments/threadedComment1.xml\" "
                               "ContentType=\"application/vnd.ms-excel.threadedcomments+xml\""),
            std::string::npos);
  EXPECT_NE(Part(saved, "xl/_rels/workbook.xml.rels").find("Target=\"persons/person.xml\""), std::string::npos);
  EXPECT_NE(
      Part(saved, "xl/worksheets/_rels/sheet1.xml.rels").find("Target=\"../threadedComments/threadedComment1.xml\""),
      std::string::npos);
  EXPECT_NE(Part(saved, "xl/drawings/vmlDrawing1.vml").find("<x:Row>0</x:Row>\n   <x:Column>3</x:Column>"),
            std::string::npos);

  Workbook reloaded = Load(saved);
  EXPECT_EQ(reloaded.persons().size(), 2U);
  ASSERT_EQ(reloaded.sheet(0).threaded_comments().size(), 1U);
  EXPECT_EQ(reloaded.sheet(0).threaded_comments()[0].text, "Hello");
  EXPECT_TRUE(reloaded.sheet(0).comments().empty());
}

TEST(ThreadedComments, ResolveAndUnresolve) {
  Workbook wb = WorkbookWithPersons();
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kThread, 0, 0, kAlice, "Q"))));
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kReply, 9, 9, kBob, "A", kThread))));
  ASSERT_TRUE(static_cast<bool>(wb.set_thread_resolved(0, kThread, true)));
  EXPECT_TRUE(Load(Save(wb)).sheet(0).threaded_comments()[0].done);
  ASSERT_TRUE(static_cast<bool>(wb.set_thread_resolved(0, kThread, false)));
  EXPECT_FALSE(Load(Save(wb)).sheet(0).threaded_comments()[0].done);

  auto on_reply = wb.set_thread_resolved(0, kReply, true);
  ASSERT_FALSE(static_cast<bool>(on_reply));
  EXPECT_EQ(on_reply.error().code, FormulonErrorCode::kInvalidArgument);
  // The reply took its thread's anchor rather than the one it was given.
  EXPECT_EQ(wb.sheet(0).threaded_comments()[1].row, 0U);
  EXPECT_EQ(wb.sheet(0).threaded_comments()[1].col, 0U);
}

TEST(ThreadedComments, DeleteThreadKeepsOtherNotes) {
  Workbook wb = Load(ExcelFixture());
  ASSERT_TRUE(static_cast<bool>(wb.remove_threaded_comment(0, kReply)));
  ASSERT_EQ(wb.sheet(0).threaded_comments().size(), 1U);
  ASSERT_TRUE(static_cast<bool>(wb.remove_threaded_comment(0, kThread)));
  EXPECT_TRUE(wb.sheet(0).threaded_comments().empty());

  const std::vector<std::uint8_t> saved = Save(wb);
  EXPECT_TRUE(Part(saved, "xl/threadedComments/threadedComment1.xml").empty());
  EXPECT_EQ(Part(saved, "xl/comments1.xml").find("tc="), std::string::npos);
  Workbook reloaded = Load(saved);
  ASSERT_EQ(reloaded.sheet(0).comments().size(), 1U);
  EXPECT_EQ(reloaded.sheet(0).comments()[0].text, "Plain note");
  EXPECT_EQ(reloaded.sheet(0).comments()[0].uid, kNoteUid);
  // Persons outlive the threads that named them.
  EXPECT_EQ(reloaded.persons().size(), 2U);
}

TEST(ThreadedComments, AnchorMovesOnRowInsert) {
  Workbook wb = Load(ExcelFixture());
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 2)));
  for (const ThreadedComment& c : wb.sheet(0).threaded_comments()) {
    EXPECT_EQ(c.row, 3U);
    EXPECT_EQ(c.col, 1U);
  }
  EXPECT_EQ(wb.sheet(0).comments()[0].row, 0U);
  const std::vector<std::uint8_t> saved = Save(wb);
  EXPECT_NE(Part(saved, "xl/threadedComments/threadedComment1.xml").find("ref=\"B4\""), std::string::npos);
  EXPECT_NE(Part(saved, "xl/drawings/vmlDrawing1.vml").find("<x:Row>3</x:Row>"), std::string::npos);
}

TEST(ThreadedComments, ThreadDeletedWithItsRow) {
  Workbook wb = Load(ExcelFixture());
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 1, 1)));
  EXPECT_TRUE(wb.sheet(0).threaded_comments().empty());
  ASSERT_EQ(wb.sheet(0).comments().size(), 1U);
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 0, 1)));
  EXPECT_TRUE(wb.sheet(0).comments().empty());
}

TEST(ThreadedComments, XlsbSaveReportsDeferral) {
  Workbook wb = WorkbookWithPersons();
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kThread, 0, 0, kAlice, "Q"))));
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kReply, 0, 0, kBob, "A", kThread))));
  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.deferred_feature_count, 2U);
}

TEST(ThreadedComments, RefusesMalformedInputCollisionsAndUnknownIds) {
  Workbook wb = WorkbookWithPersons();
  wb.sheet(0).mutable_comments().push_back(CellComment(4, 4, "Carol", "note"));
  const auto invalid = [](const auto& r) { return !r && r.error().code == FormulonErrorCode::kInvalidArgument; };
  const auto code = [](const auto& r) { return r ? FormulonErrorCode::kOk : r.error().code; };
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment("2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B23", 0, 0, kAlice, "x"))),
            FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment("{2BD0B1A4-7C3E-4F2A-9D11-6A5E8C0F1B2G}", 0, 0, kAlice, "x"))),
            FormulonErrorCode::kInvalidArgument);
  ThreadedComment bad_time = Comment(kThread, 0, 0, kAlice, "x");
  bad_time.created = "2026-10-05T10:00:00.000";
  EXPECT_TRUE(invalid(wb.add_threaded_comment(0, bad_time)));
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment(kThread, 0, 0, "{99999999-9999-4999-8999-999999999999}", "x"))),
            FormulonErrorCode::kNotFound);
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment(kThread, 4, 4, kAlice, "x"))), FormulonErrorCode::kAlreadyExists);
  EXPECT_TRUE(invalid(wb.add_threaded_comment(0, Comment(kReply, 0, 0, kAlice, "x", kThread))));

  // "日本🙂!" is five UTF-16 units: the emoji is a surrogate pair.
  ThreadedComment mentioned = Comment(kThread, 0, 0, kAlice, "日本🙂!");
  mentioned.mentions.push_back(Mention{std::string(kBob), std::string(kMention), 2, 4});
  EXPECT_TRUE(invalid(wb.add_threaded_comment(0, mentioned)));
  mentioned.mentions[0].length = 3;
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, mentioned)));
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment(kReply, 0, 0, kAlice, "x"))), FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(code(wb.add_threaded_comment(0, Comment(kThread, 1, 1, kAlice, "x"))), FormulonErrorCode::kAlreadyExists);

  EXPECT_EQ(code(wb.remove_person(kBob)), FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(invalid(wb.edit_threaded_comment(0, kThread, "", mentioned.mentions)));
  ASSERT_TRUE(static_cast<bool>(wb.edit_threaded_comment(0, kThread, "plain", {})));
  ASSERT_TRUE(static_cast<bool>(wb.remove_person(kBob)));
  EXPECT_EQ(code(wb.add_person(Person{std::string(kAlice), "Again", "", ""})), FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(code(wb.remove_person(kBob)), FormulonErrorCode::kNotFound);
  EXPECT_EQ(code(wb.remove_threaded_comment(0, kReply)), FormulonErrorCode::kNotFound);
  EXPECT_EQ(code(wb.edit_threaded_comment(0, kReply, "x", {})), FormulonErrorCode::kNotFound);
  EXPECT_EQ(code(wb.set_thread_resolved(0, kReply, true)), FormulonErrorCode::kNotFound);
  ASSERT_EQ(wb.sheet(0).threaded_comments().size(), 1U);
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].text, "plain");
}

TEST(ThreadedComments, GeneratesIdAndStampFromThePinnedClock) {
  const date_time::CivilTime pinned{{2026, 11, 9}, {8, 5, 3}};
  ThreadedComment thread = Comment("", 0, 0, kAlice, "x");
  thread.created.clear();
  Workbook wb = WorkbookWithPersons();
  wb.set_pinned_now(pinned);
  auto id = wb.add_threaded_comment(0, thread);
  ASSERT_TRUE(static_cast<bool>(id));
  EXPECT_TRUE(is_threaded_comment_guid(id.value()));
  ASSERT_EQ(wb.sheet(0).threaded_comments().size(), 1U);
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].id, id.value());
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].created, "2026-11-09T08:05:03.00");

  // The same pinned instant reproduces the same id in another workbook.
  Workbook other = WorkbookWithPersons();
  other.set_pinned_now(pinned);
  auto again = other.add_threaded_comment(0, thread);
  ASSERT_TRUE(static_cast<bool>(again));
  EXPECT_EQ(again.value(), id.value());

  ThreadedComment reply = Comment("", 0, 0, kBob, "y", id.value());
  reply.created.clear();
  auto reply_id = wb.add_threaded_comment(0, reply);
  ASSERT_TRUE(static_cast<bool>(reply_id));
  EXPECT_TRUE(is_threaded_comment_guid(reply_id.value()));
  EXPECT_NE(reply_id.value(), id.value());

  // An explicit id and stamp are kept; the same id again is refused.
  auto kept = wb.add_threaded_comment(0, Comment(kThread, 3, 3, kAlice, "z"));
  ASSERT_TRUE(static_cast<bool>(kept));
  EXPECT_EQ(kept.value(), kThread);
  EXPECT_EQ(wb.sheet(0).threaded_comments().back().created, "2026-10-05T10:00:00.00");
  auto duplicate = wb.add_threaded_comment(0, Comment(kThread, 4, 4, kAlice, "w"));
  ASSERT_FALSE(static_cast<bool>(duplicate));
  EXPECT_EQ(duplicate.error().code, FormulonErrorCode::kAlreadyExists);
}

// Restores the process time zone when a test that changes `TZ` ends.
class ScopedTimeZone {
 public:
  explicit ScopedTimeZone(const char* zone) {
    const char* old = std::getenv("TZ");
    had_ = old != nullptr;
    if (had_) {
      saved_ = old;
    }
    setenv("TZ", zone, 1);
    tzset();
  }
  ~ScopedTimeZone() {
    if (had_) {
      setenv("TZ", saved_.c_str(), 1);
    } else {
      unsetenv("TZ");
    }
    tzset();
  }
  ScopedTimeZone(const ScopedTimeZone&) = delete;
  ScopedTimeZone& operator=(const ScopedTimeZone&) = delete;

 private:
  bool had_ = false;
  std::string saved_;
};

// Seconds since the epoch of a `YYYY-MM-DDTHH:MM:SS.ff` value read as UTC.
std::time_t StampAsUtcSeconds(const std::string& stamp) {
  std::tm tm{};
  tm.tm_year = std::stoi(stamp.substr(0, 4)) - 1900;
  tm.tm_mon = std::stoi(stamp.substr(5, 2)) - 1;
  tm.tm_mday = std::stoi(stamp.substr(8, 2));
  tm.tm_hour = std::stoi(stamp.substr(11, 2));
  tm.tm_min = std::stoi(stamp.substr(14, 2));
  tm.tm_sec = std::stoi(stamp.substr(17, 2));
  return timegm(&tm);
}

std::string UnpinnedStampUnder(const char* zone) {
  ScopedTimeZone tz(zone);
  Workbook wb = WorkbookWithPersons();
  ThreadedComment c = Comment(kThread, 0, 0, kAlice, "x");
  c.created.clear();
  const std::time_t before = std::time(nullptr);
  EXPECT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, c)));
  const std::time_t after = std::time(nullptr);
  const std::string stamp = wb.sheet(0).threaded_comments()[0].created;
  EXPECT_TRUE(is_threaded_comment_timestamp(stamp)) << stamp;
  const std::time_t at = StampAsUtcSeconds(stamp);
  EXPECT_GE(at, before - 2);
  EXPECT_LE(at, after + 2);
  return stamp;
}

TEST(ThreadedComments, UnpinnedStampIsUtcWhateverTheHostZone) {
  const std::string tokyo = UnpinnedStampUnder("Asia/Tokyo");
  const std::string utc = UnpinnedStampUnder("UTC");
  EXPECT_NEAR(static_cast<double>(StampAsUtcSeconds(tokyo)), static_cast<double>(StampAsUtcSeconds(utc)), 2.0);
}

TEST(ThreadedComments, PinnedStampIsWrittenAsIsUnderAnyHostZone) {
  ScopedTimeZone tz("Asia/Tokyo");
  Workbook wb = WorkbookWithPersons();
  wb.set_pinned_now(date_time::CivilTime{{2026, 11, 9}, {17, 4, 48}});
  ThreadedComment c = Comment(kThread, 0, 0, kAlice, "x");
  c.created.clear();
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, c)));
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].created, "2026-11-09T17:04:48.00");
}

TEST(ThreadedComments, InputStampWithZuluSuffixIsStoredWithoutIt) {
  Workbook wb = WorkbookWithPersons();
  ThreadedComment c = Comment(kThread, 0, 0, kAlice, "x");
  c.created = "2026-10-06T09:00:00.00Z";
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, c)));
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].created, "2026-10-06T09:00:00.00");
}

TEST(ThreadedComments, InputStampWithOffsetOrThreeDigitFractionIsRefused) {
  for (const char* bad : {"2026-10-06T09:00:00.00+09:00", "2026-10-06T09:00:00.123", "2026-10-06T09:00:00.123Z"}) {
    Workbook wb = WorkbookWithPersons();
    ThreadedComment c = Comment(kThread, 0, 0, kAlice, "x");
    c.created = bad;
    auto r = wb.add_threaded_comment(0, c);
    ASSERT_FALSE(static_cast<bool>(r)) << bad;
    EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument) << bad;
  }
}

std::vector<ThreadedComment> ReadWithDt(const std::string& dt_attr) {
  const std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\">"
      "<threadedComment ref=\"A1\" " +
      dt_attr + " personId=\"" + std::string(kAlice) + "\" id=\"" + std::string(kThread) +
      "\"><text>t</text></threadedComment></ThreadedComments>";
  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto decoded = io::read_threaded_comments(bytes);
  EXPECT_TRUE(static_cast<bool>(decoded));
  return decoded ? decoded.value() : std::vector<ThreadedComment>{};
}

TEST(ThreadedComments, LoadedStampWithZuluSuffixReadsWithoutIt) {
  const auto list = ReadWithDt("dT=\"2026-10-06T09:00:00.00Z\"");
  ASSERT_EQ(list.size(), 1U);
  EXPECT_EQ(list[0].created, "2026-10-06T09:00:00.00");
}

TEST(ThreadedComments, LoadedStampWithOffsetOrThreeDigitFractionIsAbsentAndOmittedOnSave) {
  for (const char* bad : {"2026-10-06T09:00:00.00+09:00", "2026-10-06T09:00:00.123"}) {
    const auto list = ReadWithDt(std::string("dT=\"") + bad + "\"");
    ASSERT_EQ(list.size(), 1U) << bad;
    EXPECT_TRUE(list[0].created.empty()) << bad;
    EXPECT_EQ(io::write_threaded_comments(list).find("dT="), std::string::npos) << bad;
  }
}

TEST(ThreadedComments, PersonIdsMatchAcrossCase) {
  Workbook wb = WorkbookWithPersons();
  const std::string alice_lower = "{11111111-aaaa-4bbb-8ccc-dddddddddddd}";
  ThreadedComment thread = Comment(kThread, 0, 0, alice_lower, "x");
  thread.mentions.push_back(Mention{"{22222222-aaaa-4bbb-8ccc-eeeeeeeeeeee}", std::string(kMention), 0, 1});
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, thread)));
  // Stored in the listed spelling.
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].person_id, kAlice);
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].mentions[0].person_id, kBob);

  auto refused = wb.add_person(Person{alice_lower, "Alice again", "", ""});
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kAlreadyExists);
  auto in_use = wb.remove_person(alice_lower);
  ASSERT_FALSE(static_cast<bool>(in_use));
  EXPECT_EQ(in_use.error().code, FormulonErrorCode::kPreconditionFailed);
  ASSERT_TRUE(static_cast<bool>(wb.remove_threaded_comment(0, kThread)));
  ASSERT_TRUE(static_cast<bool>(wb.remove_person(alice_lower)));
  ASSERT_EQ(wb.persons().size(), 1U);
  EXPECT_EQ(wb.persons()[0].id, kBob);
}

}  // namespace
TEST(ThreadedComments, GuidIdsAreTakenInEitherCaseStoredUppercaseAndMatchedAcrossCase) {
  Workbook wb = WorkbookWithPersons();
  ASSERT_TRUE(static_cast<bool>(wb.add_person(Person{"{33333333-abcd-4ef0-8abc-0123456789ab}", "Carol", "", ""})));
  EXPECT_EQ(wb.persons().back().id, "{33333333-ABCD-4EF0-8ABC-0123456789AB}");
  const auto code = [](const auto& r) { return r ? FormulonErrorCode::kOk : r.error().code; };
  EXPECT_EQ(code(wb.add_person(Person{"{33333333-ABCD-4EF0-8ABC-0123456789AB}", "Again", "", ""})),
            FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(code(wb.add_person(Person{"{not-a-guid}", "Dan", "", ""})), FormulonErrorCode::kInvalidArgument);

  const std::string lower_thread = "{6f1e0d2c-1111-4111-8111-0000000000aa}";
  ThreadedComment thread = Comment(lower_thread, 0, 0, kAlice, "Q");
  thread.mentions.push_back(Mention{std::string(kBob), "{7a2b3c4d-2222-4222-8222-0000000000bb}", 0, 1});
  auto id = wb.add_threaded_comment(0, thread);
  ASSERT_TRUE(static_cast<bool>(id));
  EXPECT_EQ(id.value(), "{6F1E0D2C-1111-4111-8111-0000000000AA}");
  EXPECT_EQ(wb.sheet(0).threaded_comments()[0].mentions[0].mention_id, "{7A2B3C4D-2222-4222-8222-0000000000BB}");
  EXPECT_EQ(find_threaded_comment(wb.sheet(0).threaded_comments(), lower_thread), 0U);

  // A reply naming its thread in another case binds to the stored spelling.
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, Comment(kReply, 0, 0, kBob, "A", lower_thread))));
  EXPECT_EQ(wb.sheet(0).threaded_comments()[1].parent_id, id.value());
  ASSERT_TRUE(static_cast<bool>(wb.set_thread_resolved(0, lower_thread, true)));
  ASSERT_TRUE(static_cast<bool>(wb.remove_threaded_comment(0, lower_thread)));
  EXPECT_TRUE(wb.sheet(0).threaded_comments().empty());
}

}  // namespace formulon
