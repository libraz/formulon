//
// Integration test: a workbook carrying `<externalReferences>` plus
// the matching `xl/_rels/workbook.xml.rels` and per-link rels files
// must survive a full writer -> reader cycle without losing the
// references. The body parts themselves ride through passthrough.

#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"
#include "workbook_format.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

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

  PassthroughPart body;
  body.path = "xl/externalLinks/externalLink1.xml";
  body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
  body.bytes = std::move(body_bytes);
  std::vector<PassthroughPart> parts;
  parts.push_back(std::move(body));
  src.set_passthrough_parts(std::move(parts));

  ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId7";
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "file:///Users/example/RemoteBook.xlsx";
  rec.target_external = true;
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  std::vector<ExternalLinkRecord> links;
  links.push_back(std::move(rec));
  src.set_external_links(std::move(links));
  return src;
}

TEST(ExternalLinksRoundTrip, PreservesSingleExternalBook) {
  Workbook src = Workbook::create();

  // 1. Body part lives in passthrough; it round-trips verbatim.
  PassthroughPart body;
  body.path = "xl/externalLinks/externalLink1.xml";
  body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
  body.bytes = MakeExternalBookBody("rId1");

  std::vector<PassthroughPart> parts;
  parts.push_back(std::move(body));
  src.set_passthrough_parts(std::move(parts));

  // 2. Metadata: one external book pointing at a remote file URL.
  ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId7";  // Reader replaces this on round-trip; value here is irrelevant.
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "file:///Users/example/RemoteBook.xlsx";
  rec.target_external = true;
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  std::vector<ExternalLinkRecord> links;
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
  EXPECT_EQ(rt[0].kind, ExternalLinkRecord::Kind::kExternalBook);
  // The body part must still be present in passthrough — the reader
  // should have left it alone (it consumes only the per-link rels file).
  bool body_present = false;
  for (const PassthroughPart& p : load_or.value().workbook.passthrough_parts()) {
    if (p.path == "xl/externalLinks/externalLink1.xml") {
      body_present = true;
      break;
    }
  }
  EXPECT_TRUE(body_present);
}

TEST(ExternalLinksRoundTrip, PreservesDocumentOrderAcrossMultipleLinks) {
  Workbook src = Workbook::create();

  std::vector<PassthroughPart> parts;
  for (int i = 1; i <= 3; ++i) {
    PassthroughPart body;
    body.path = "xl/externalLinks/externalLink" + std::to_string(i) + ".xml";
    body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    body.bytes = MakeExternalBookBody("rId1");
    parts.push_back(std::move(body));
  }
  src.set_passthrough_parts(std::move(parts));

  std::vector<ExternalLinkRecord> links;
  for (int i = 1; i <= 3; ++i) {
    ExternalLinkRecord rec;
    rec.index = static_cast<std::uint32_t>(i);
    rec.rel_id = "rId" + std::to_string(100 + i);
    rec.part_path = "xl/externalLinks/externalLink" + std::to_string(i) + ".xml";
    rec.body_rel_id = "rId1";
    rec.target = "file:///remote/book" + std::to_string(i) + ".xlsx";
    rec.target_external = true;
    rec.kind = ExternalLinkRecord::Kind::kExternalBook;
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
    EXPECT_EQ(rt[i].kind, ExternalLinkRecord::Kind::kExternalBook);
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
  EXPECT_EQ(links[0].kind, ExternalLinkRecord::Kind::kUnknown);
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

// `TargetMode` is optional in the packaging relationships, and its
// default is `Internal` — so the writer omits the attribute for an
// in-package target, and the reader must read that omission back as
// in-package. Getting this backwards is not a cosmetic flag error: a
// relationship the source file did NOT declare as external comes back
// marked external, and the next save writes `TargetMode="External"` into
// the package. That turns an in-package reference into one a consumer
// will try to fetch from outside the file.
//
// The whole cycle stays inside our own writer and reader, so a
// disagreement here is a plain round-trip loss, independent of what any
// particular producer happens to emit.
TEST(ExternalLinksRoundTrip, InPackageTargetModeSurvivesRoundTrip) {
  Workbook src = MakeSingleExternalLinkWorkbook(MakeExternalBookBody("rId1"));
  {
    std::vector<ExternalLinkRecord> links = src.external_links();
    ASSERT_EQ(links.size(), 1U);
    links[0].target = "externalLinks/RemoteBook.xlsx";
    links[0].target_external = false;
    src.set_external_links(std::move(links));
  }

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  const auto& rt = load_or.value().workbook.external_links();
  ASSERT_EQ(rt.size(), 1U);
  EXPECT_EQ(rt[0].target, "externalLinks/RemoteBook.xlsx");
  EXPECT_FALSE(rt[0].target_external) << "an in-package target came back marked external";
}

// The common path, pinned alongside the case above so a fix to one is not
// free to break the other.
TEST(ExternalLinksRoundTrip, ExternalTargetModeSurvivesRoundTrip) {
  Workbook src = MakeSingleExternalLinkWorkbook(MakeExternalBookBody("rId1"));
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));

  const auto& rt = load_or.value().workbook.external_links();
  ASSERT_EQ(rt.size(), 1U);
  EXPECT_TRUE(rt[0].target_external);
}

// `part_path` becomes a zip entry name at save time, and it can reach the
// model without passing the reader's traversal check — `set_external_links`
// takes it verbatim. The writer applies the same refusal it already applies
// to passthrough part names, so the model cannot be used to place an entry
// outside the package for whoever extracts the result later.
TEST(ExternalLinksRoundTrip, TraversalShapedPartPathIsRefusedOnWrite) {
  Workbook src = MakeSingleExternalLinkWorkbook(MakeExternalBookBody("rId1"));
  {
    std::vector<ExternalLinkRecord> links = src.external_links();
    ASSERT_EQ(links.size(), 1U);
    links[0].part_path = "../../evil/externalLink1.xml";
    src.set_external_links(std::move(links));
  }

  auto save_or = src.save();
  ASSERT_FALSE(static_cast<bool>(save_or)) << "a traversal-shaped part path was written into the package";
  EXPECT_EQ(save_or.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(ExternalLinksRoundTrip, EmptyWorkbookEmitsNoExternalReferencesBlock) {
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  EXPECT_TRUE(load_or.value().workbook.external_links().empty());
}

// ---------------------------------------------------------------------------
// Authoring a cross-workbook reference, saving it and reading it back.
// ---------------------------------------------------------------------------

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
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

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(f);
  return out;
}

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

std::string EntryText(const std::vector<std::uint8_t>& package, std::string_view name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry_or)) << "missing entry " << name;
  if (!entry_or) {
    return {};
  }
  return std::string(entry_or.value().begin(), entry_or.value().end());
}

// Largest N of any `[N]` (a bracketed run of digits) in `text`; 0 when none.
std::uint32_t MaxBracketIndex(const std::string& text) {
  std::uint32_t max_index = 0;
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] != '[') {
      continue;
    }
    std::size_t j = i + 1U;
    std::uint32_t n = 0;
    while (j < text.size() && text[j] >= '0' && text[j] <= '9' && j - i < 9U) {
      n = n * 10U + static_cast<std::uint32_t>(text[j] - '0');
      ++j;
    }
    if (j > i + 1U && j < text.size() && text[j] == ']' && n > max_index) {
      max_index = n;
    }
  }
  return max_index;
}

// Concatenated content of every `<tag>...</tag>` run, separated by '\n'.
std::string TagTexts(const std::string& xml, const std::string& tag) {
  const std::string open = "<" + tag + ">";
  const std::string close = "</" + tag + ">";
  std::string out;
  std::size_t pos = 0;
  while ((pos = xml.find(open, pos)) != std::string::npos) {
    const std::size_t begin = pos + open.size();
    const std::size_t end = xml.find(close, begin);
    if (end == std::string::npos) {
      break;
    }
    out.append(xml, begin, end - begin);
    out.push_back('\n');
    pos = end + close.size();
  }
  return out;
}

std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value ValueAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  return wb.sheet(sheet).resolve_cell_value(row, col);
}

constexpr const char* kCfFormula = "[Book.xlsx]Sheet1!$A$1>0";
constexpr const char* kNameFormula = "[Book.xlsx]Sheet1!$A$1";

// Cells A1:A6 on the first sheet, one conditional format and one defined name,
// every one of them naming a workbook that is not open.
Workbook AuthorExternalReferences() {
  Workbook wb = Workbook::create();
  const char* formulas[] = {
      "=[Book.xlsx]Sheet1!A1", "='/Users/x/[Path.xlsx]Sheet1'!A1", "=SUM([Book.xlsx]Sheet1:Sheet2!A1)",
      "=Book.xlsx!Total",      "=[Other.xlsx]Sheet1!B2",           "='/Users/x/[Book.xlsx]Sheet1'!A1",
  };
  std::uint32_t row = 0;
  for (const char* f : formulas) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row++, 0, f))) << f;
  }
  cf::ConditionalFormat block;
  block.sqref.push_back(cf::CFCellRange{CellAddress{0, 2}, CellAddress{9, 2}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = kCfFormula;
  block.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));
  wb.bind_external_books(kCfFormula);
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("ExtName", kNameFormula)));
  return wb;
}

Workbook ReloadAs(const Workbook& src, WorkbookFormat format, std::vector<std::uint8_t>* package_out) {
  auto save_or = src.save_as(format);
  EXPECT_TRUE(static_cast<bool>(save_or)) << (save_or ? "" : save_or.error().message);
  if (!save_or) {
    return Workbook::create_empty();
  }
  *package_out = save_or.value();
  if (format == WorkbookFormat::Xlsb) {
    auto load_or = io::xlsb::read_xlsb(SpanOf(*package_out));
    EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
    return load_or ? std::move(load_or.value().workbook) : Workbook::create_empty();
  }
  auto load_or = io::read_ooxml(SpanOf(*package_out));
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  return load_or ? std::move(load_or.value().workbook) : Workbook::create_empty();
}

void ExpectAuthoredReferencesSurvive(WorkbookFormat format) {
  Workbook src = AuthorExternalReferences();
  const std::size_t link_count = src.external_links().size();
  ASSERT_EQ(link_count, 3U) << "Book.xlsx (also spelled with a path), Path.xlsx and Other.xlsx";

  std::vector<std::uint8_t> package;
  Workbook wb = ReloadAs(src, format, &package);
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.external_links().size(), link_count);

  // Book.xlsx is linked once; the path-qualified entry gives that link its absolute path,
  // so every reference to it reads back in the path-qualified spelling.
  EXPECT_EQ(FormulaAt(wb, 0, 0, 0), "='/Users/x/[Book.xlsx]Sheet1'!A1");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 0), "='/Users/x/[Path.xlsx]Sheet1'!A1");
  EXPECT_EQ(FormulaAt(wb, 0, 2, 0), "=SUM('/Users/x/[Book.xlsx]Sheet1:Sheet2'!A1)");
  EXPECT_EQ(FormulaAt(wb, 0, 3, 0), "='/Users/x/Book.xlsx'!Total");
  EXPECT_EQ(FormulaAt(wb, 0, 4, 0), "=[Other.xlsx]Sheet1!B2");
  EXPECT_EQ(FormulaAt(wb, 0, 5, 0), "='/Users/x/[Book.xlsx]Sheet1'!A1");

  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  ASSERT_EQ(wb.sheet(0).conditional_formats()[0].rules.size(), 1U);
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value_or(""), "'/Users/x/[Book.xlsx]Sheet1'!$A$1>0");
  bool name_found = false;
  for (const DefinedName& dn : wb.defined_names()) {
    if (dn.name == "ExtName") {
      name_found = true;
      EXPECT_EQ(dn.formula, "'/Users/x/[Book.xlsx]Sheet1'!$A$1");
    }
  }
  EXPECT_TRUE(name_found);

  // The links carry no sheetData, so every reference reads #REF!. The one
  // exception is the name `Total`, which the supporting book does not declare.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t row = 0; row < 6U; ++row) {
    const Value v = ValueAt(wb, 0, row, 0);
    ASSERT_TRUE(v.is_error()) << "row " << row;
    const ErrorCode expected = row == 3U ? ErrorCode::Name : ErrorCode::Ref;
    EXPECT_EQ(v.as_error(), expected) << "row " << row;
  }

  if (format == WorkbookFormat::Ooxml) {
    // Every `[N]` the package stores must name a link that exists.
    const std::string sheet_xml = EntryText(package, "xl/worksheets/sheet1.xml");
    const std::string workbook_xml = EntryText(package, "xl/workbook.xml");
    EXPECT_FALSE(TagTexts(sheet_xml, "f").empty());
    EXPECT_LE(MaxBracketIndex(sheet_xml), link_count);
    EXPECT_LE(MaxBracketIndex(workbook_xml), link_count);
    EXPECT_GE(MaxBracketIndex(sheet_xml), 1U);
    EXPECT_NE(sheet_xml.find("[1]"), std::string::npos);
    EXPECT_NE(EntryText(package, "xl/externalLinks/_rels/externalLink1.xml.rels").find("/Users/x/Book.xlsx"),
              std::string::npos);
    EXPECT_NE(EntryText(package, "xl/externalLinks/externalLink1.xml").find("absoluteUrl"), std::string::npos);
  }
}

TEST(ExternalLinksAuthoring, XlsxRoundTripKeepsEveryReference) {
  ExpectAuthoredReferencesSurvive(WorkbookFormat::Ooxml);
}

TEST(ExternalLinksAuthoring, XlsbRoundTripKeepsEveryReference) {
  ExpectAuthoredReferencesSurvive(WorkbookFormat::Xlsb);
}

TEST(ExternalLinksAuthoring, NumericBracketIsABookNamedByTheNumber) {
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0, 0, 0, "=[1]Sheet1!A1")));
  ASSERT_EQ(src.external_links().size(), 1U);
  std::vector<std::uint8_t> package;
  Workbook wb = ReloadAs(src, WorkbookFormat::Ooxml, &package);
  EXPECT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(FormulaAt(wb, 0, 0, 0), "=[1]Sheet1!A1");
  EXPECT_LE(MaxBracketIndex(EntryText(package, "xl/worksheets/sheet1.xml")), 1U);
}

// An extensionless book's book-scope name is `'Src2'!Name` in the formula bar and
// `[N]!Name` in both containers.
void ExpectExtensionlessBookScopeSurvives(WorkbookFormat format) {
  Workbook src = Workbook::create();
  std::vector<ExternalLinkRecord> links(1);
  links[0].index = 1;
  links[0].target = "Src2";
  links[0].kind = ExternalLinkRecord::Kind::kExternalBook;
  links[0].book.sheet_names = {"Data"};
  links[0].book.sheet_data = {true};
  ExternalBookName total;
  total.name = "Total";
  total.resolvable = true;
  links[0].book.names = {total};
  links[0].book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::number(42.0);
  src.set_external_links(std::move(links));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0, 0, 0, "='Src2'!Total")));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0, 1, 0, "='Src2'!NoSuch")));

  std::vector<std::uint8_t> package;
  Workbook wb = ReloadAs(src, format, &package);
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(FormulaAt(wb, 0, 0, 0), "='Src2'!Total");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 0), "='Src2'!NoSuch");
  if (format == WorkbookFormat::Ooxml) {
    const std::string sheet_xml = EntryText(package, "xl/worksheets/sheet1.xml");
    EXPECT_NE(sheet_xml.find("[1]!Total"), std::string::npos) << sheet_xml;
    EXPECT_NE(sheet_xml.find("[1]!NoSuch"), std::string::npos) << sheet_xml;
  }
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value total_value = ValueAt(wb, 0, 0, 0);
  ASSERT_TRUE(total_value.is_number());
  EXPECT_DOUBLE_EQ(total_value.as_number(), 42.0);
  const Value missing = ValueAt(wb, 0, 1, 0);
  ASSERT_TRUE(missing.is_error());
  EXPECT_EQ(missing.as_error(), ErrorCode::Name);
}

TEST(ExternalLinksRoundTrip, ExtensionlessBookScopeSurvivesXlsx) {
  ExpectExtensionlessBookScopeSurvives(WorkbookFormat::Ooxml);
}

TEST(ExternalLinksRoundTrip, ExtensionlessBookScopeSurvivesXlsb) {
  ExpectExtensionlessBookScopeSurvives(WorkbookFormat::Xlsb);
}

// ---------------------------------------------------------------------------
// A file Excel wrote: re-saving keeps the stored text and the absolute path.
// ---------------------------------------------------------------------------

TEST(ExternalLinksFixture, ResaveKeepsStoredFormulaTextAndAbsoluteUrl) {
  const std::vector<std::uint8_t> original = ReadFileBytes(FixturePath("external_link_mixed.xlsx"));
  ASSERT_FALSE(original.empty());
  const std::string original_formulas = TagTexts(EntryText(original, "xl/worksheets/sheet1.xml"), "f");
  ASSERT_FALSE(original_formulas.empty());

  auto first_or = io::read_ooxml(SpanOf(original));
  ASSERT_TRUE(static_cast<bool>(first_or)) << first_or.error().message;
  EXPECT_EQ(FormulaAt(first_or.value().workbook, 0, 2, 0),
            "='/Users/libraz/Documents/ext_link_probe/[ExtSource.xlsx]Data'!A1");

  std::vector<std::uint8_t> save1;
  Workbook second = ReloadAs(first_or.value().workbook, WorkbookFormat::Ooxml, &save1);
  EXPECT_EQ(TagTexts(EntryText(save1, "xl/worksheets/sheet1.xml"), "f"), original_formulas);

  std::vector<std::uint8_t> save2;
  Workbook third = ReloadAs(second, WorkbookFormat::Ooxml, &save2);
  EXPECT_EQ(TagTexts(EntryText(save2, "xl/worksheets/sheet1.xml"), "f"), original_formulas);
  ASSERT_EQ(third.external_links().size(), 2U);
  EXPECT_EQ(third.external_links()[0].absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx");

  const std::string rels = EntryText(save2, "xl/externalLinks/_rels/externalLink1.xml.rels");
  EXPECT_NE(rels.find("/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx"), std::string::npos);
  EXPECT_NE(EntryText(save2, "xl/externalLinks/externalLink1.xml").find("absoluteUrl"), std::string::npos);
}

// The reader must find the absolute-path relationship wherever it sits in the
// link's rels file, including after the relationship the body part names.
TEST(ExternalLinksFixture, AbsoluteTargetIsReadWhenItsRelFollowsTheBodyRel) {
  const std::vector<std::uint8_t> original = ReadFileBytes(FixturePath("external_link_mixed.xlsx"));
  ASSERT_FALSE(original.empty());

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(original))));
  const std::string rels_name = "xl/externalLinks/_rels/externalLink1.xml.rels";
  const std::string swapped =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
      "externalLinkPath\" Target=\"ExtSource.xlsx\" TargetMode=\"External\"/>"
      "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
      "externalLinkPath\" Target=\"/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx\" TargetMode=\"External\"/>"
      "</Relationships>";
  ASSERT_NE(EntryText(original, rels_name).find("rId2"), EntryText(original, rels_name).find("rId1"));

  std::vector<std::string> names = zip.list_entries();
  std::vector<std::string> bodies;
  bodies.reserve(names.size());
  for (const std::string& name : names) {
    bodies.push_back(name == rels_name ? swapped : EntryText(original, name));
  }
  std::vector<PartFile> parts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    parts.push_back(PartFile{names[i].c_str(), bodies[i]});
  }

  auto load_or = io::read_ooxml(SpanOf(BuildZip(parts)));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  const auto& links = load_or.value().workbook.external_links();
  ASSERT_EQ(links.size(), 2U);
  EXPECT_EQ(links[0].target, "ExtSource.xlsx");
  EXPECT_EQ(links[0].absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx");
  EXPECT_EQ(links[0].absolute_rel_id, "rId2");
}

}  // namespace
}  // namespace formulon
