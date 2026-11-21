// A book whose rich-value parts stay verbatim saves every cell's `vm`: Excel's
// pivot-record pictures come out in both formats, from either source, with the
// vm, metadata part, richData parts and media Excel itself writes; a format
// change converts the metadata part and appends the dynamic-array type a new
// formula needs; a part that cannot be converted refuses the save while a cell
// still names it, and is dropped and counted once none does.

#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/rich_value_errors.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

using VmTable = std::map<std::pair<std::size_t, std::string>, std::uint32_t>;

// Per Excel pair, the vm of every cell carrying one, by (sheet index, A1), as Excel saved it in both formats.
struct Pair {
  const char* name;
  VmTable vms;
};

const std::vector<Pair>& Pairs() {
  static const std::vector<Pair> pairs = {
      {"one_picture", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
      {"two_pictures", {{{0, "A6"}, 1}, {{0, "A7"}, 2}, {{1, "A3"}, 1}, {{1, "A5"}, 2}}},
      {"picture_and_text", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
      {"seven_items",
       {{{0, "A6"}, 1}, {{0, "A7"}, 2}, {{0, "A10"}, 3}, {{1, "A3"}, 1}, {{1, "A5"}, 2}, {{1, "A8"}, 3}}},
      {"numeric_field", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
  };
  return pairs;
}

enum class Format { kXlsx, kXlsb };

const char* Extension(Format f) {
  return f == Format::kXlsx ? ".xlsx" : ".xlsb";
}

std::vector<std::uint8_t> ReadPair(const char* name, Format format) {
  std::vector<std::uint8_t> bytes = test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) +
                                                          "/excel/pivot_rich_records/" + name + Extension(format));
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

std::optional<std::vector<std::uint8_t>> Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  if (!zip.has_entry(part)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(bytes)) << part;
  return bytes ? std::optional<std::vector<std::uint8_t>>(std::move(bytes.value())) : std::nullopt;
}

std::string Text(const std::optional<std::vector<std::uint8_t>>& bytes) {
  return bytes ? std::string(bytes->begin(), bytes->end()) : std::string();
}

std::vector<std::string> Entries(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  return zip.list_entries();
}

/// `package` with each part of `replaced` swapped for its new bytes.
std::vector<std::uint8_t> Repack(const std::vector<std::uint8_t>& package,
                                 const std::map<std::string, std::vector<std::uint8_t>>& replaced) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const std::string& name : zip.list_entries()) {
    const auto it = replaced.find(name);
    std::vector<std::uint8_t> bytes;
    if (it != replaced.end()) {
      bytes = it->second;
    } else if (auto read = zip.read_entry(name)) {
      bytes = std::move(read.value());
    }
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name.c_str(), bytes.data(), bytes.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << name;
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  mz_zip_writer_end(&writer);
  return out;
}

/// `bytes` with their one occurrence of `from` replaced by `to`.
std::vector<std::uint8_t> ReplaceOnce(const std::vector<std::uint8_t>& bytes, const std::string& from,
                                      const std::string& to) {
  std::string text(bytes.begin(), bytes.end());
  const std::size_t at = text.find(from);
  EXPECT_NE(at, std::string::npos);
  EXPECT_EQ(text.find(from, at + 1U), std::string::npos);
  if (at != std::string::npos) {
    text.replace(at, from.size(), to);
  }
  return std::vector<std::uint8_t>(text.begin(), text.end());
}

Workbook Load(const std::vector<std::uint8_t>& bytes, Format format) {
  if (format == Format::kXlsx) {
    auto loaded = read_ooxml(test::span_of(bytes));
    EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
    return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
  }
  auto loaded = xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

struct Saved {
  std::vector<std::uint8_t> bytes;
  WriteDiagnostics diagnostics;
};

Expected<Saved, Error> TrySave(const Workbook& wb, Format format) {
  if (format == Format::kXlsx) {
    auto saved = write_ooxml_with_result(wb);
    if (!saved) {
      return std::move(saved.error());
    }
    return Saved{std::move(saved.value().bytes), saved.value().diagnostics};
  }
  auto saved = xlsb::write_xlsb_with_result(wb);
  if (!saved) {
    return std::move(saved.error());
  }
  return Saved{std::move(saved.value().bytes), saved.value().diagnostics};
}

Saved Save(const Workbook& wb, Format format) {
  auto saved = TrySave(wb, format);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : Saved{};
}

/// Every opaque rich value of `wb`, by (sheet index, A1): its `verbatim_vm`, with the cell still caching #VALUE!.
VmTable VerbatimVms(const Workbook& wb) {
  VmTable out;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const CellRichValue& v : wb.sheet(s).rich_values()) {
      const Cell* cell = wb.sheet(s).cell_at(v.row, v.col);
      EXPECT_TRUE(cell != nullptr && cell->cached_value.is_error() &&
                  cell->cached_value.as_error() == ErrorCode::Value);
      out[{s, a1::encode_a1(v.row, v.col)}] = v.value.verbatim_vm;
    }
  }
  return out;
}

std::string_view MetadataPath(Format f) {
  return f == Format::kXlsx ? "xl/metadata.xml" : "xl/metadata.bin";
}

std::string_view WorkbookRelsPath(Format f) {
  return f == Format::kXlsx ? "xl/_rels/workbook.xml.rels" : "xl/_rels/workbook.bin.rels";
}

/// The workbook relationships of `package` as (type, target).
std::vector<std::pair<std::string, std::string>> WorkbookRels(const std::vector<std::uint8_t>& package, Format f) {
  std::vector<std::pair<std::string, std::string>> out;
  const std::string xml = Text(Part(package, WorkbookRelsPath(f)));
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size()));
  for (pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
    out.emplace_back(rel.attribute("Type").value(), rel.attribute("Target").value());
  }
  return out;
}

std::set<std::pair<std::string, std::string>> RichDataRels(const std::vector<std::uint8_t>& package, Format f) {
  std::set<std::pair<std::string, std::string>> out;
  for (auto& rel : WorkbookRels(package, f)) {
    if (rel.second.rfind("richData/", 0) == 0) {
      out.insert(std::move(rel));
    }
  }
  return out;
}

std::size_t Occurrences(const std::string& text, std::string_view needle) {
  std::size_t n = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + 1U)) {
    ++n;
  }
  return n;
}

/// The `xl/richData/` and `xl/media/` parts of `package`, by path.
std::map<std::string, std::vector<std::uint8_t>> RichParts(const std::vector<std::uint8_t>& package) {
  std::map<std::string, std::vector<std::uint8_t>> out;
  for (const std::string& name : Entries(package)) {
    if (name.rfind("xl/richData/", 0) == 0 || name.rfind("xl/media/", 0) == 0) {
      out[name] = Part(package, name).value_or(std::vector<std::uint8_t>{});
    }
  }
  return out;
}

std::optional<MetadataModel> MetadataOf(const std::vector<std::uint8_t>& package, Format f) {
  const std::optional<std::vector<std::uint8_t>> bytes = Part(package, MetadataPath(f));
  if (!bytes.has_value()) {
    return std::nullopt;
  }
  return f == Format::kXlsx ? parse_metadata_xml_model(*bytes) : xlsb::parse_metadata_bin_model(test::span_of(*bytes));
}

constexpr Format kFormats[] = {Format::kXlsx, Format::kXlsb};

Format Other(Format f) {
  return f == Format::kXlsx ? Format::kXlsb : Format::kXlsx;
}

TEST(VerbatimRichVm, ExcelPivotRecordBooksSaveLikeExcelInEveryDirection) {
  for (const Pair& pair : Pairs()) {
    for (const Format from : kFormats) {
      for (const Format to : kFormats) {
        SCOPED_TRACE(std::string(pair.name) + Extension(from) + " -> " + Extension(to));
        const Saved saved = Save(Load(ReadPair(pair.name, from), from), to);
        ASSERT_FALSE(saved.bytes.empty());
        const std::vector<std::uint8_t> twin = ReadPair(pair.name, to);
        EXPECT_EQ(VerbatimVms(Load(saved.bytes, to)), pair.vms);
        EXPECT_EQ(Part(saved.bytes, MetadataPath(to)), Part(twin, MetadataPath(to)));
        EXPECT_FALSE(Part(saved.bytes, MetadataPath(Other(to))).has_value());
        EXPECT_EQ(RichParts(saved.bytes), RichParts(twin));
        EXPECT_EQ(Occurrences(Text(Part(saved.bytes, "[Content_Types].xml")), "PartName=\"/xl/metadata."), 1U);
        std::size_t metadata_rels = 0;
        for (const auto& [type, target] : WorkbookRels(saved.bytes, to)) {
          metadata_rels += type == kRelSheetMetadata ? 1U : 0U;
        }
        EXPECT_EQ(metadata_rels, 1U);
        EXPECT_EQ(RichDataRels(saved.bytes, to), RichDataRels(twin, to));
        EXPECT_EQ(saved.diagnostics.deferred_feature_count, 0U);
      }
    }
  }
}

TEST(VerbatimRichVm, EditsMoveAndDropTheVm) {
  const Pair& pair = Pairs()[1];  // two_pictures
  for (const Format from : kFormats) {
    for (const Format to : kFormats) {
      SCOPED_TRACE(std::string(Extension(from)) + " -> " + Extension(to));
      Workbook wb = Load(ReadPair(pair.name, from), from);
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 5U, 0U, Value::number(1.0))));  // A6
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 6U, 0U, Value::blank())));      // A7
      ASSERT_TRUE(static_cast<bool>(wb.insert_rows(1, 0U, 1U)));
      const Saved saved = Save(wb, to);
      ASSERT_FALSE(saved.bytes.empty());
      EXPECT_EQ(VerbatimVms(Load(saved.bytes, to)), (VmTable{{{1, "A4"}, 1}, {{1, "A6"}, 2}}));
    }
  }
}

TEST(VerbatimRichVm, AFormatChangeAddsTheDynamicArrayTypeANewFormulaNeeds) {
  const Pair& pair = Pairs()[1];  // two_pictures
  for (const Format from : kFormats) {
    const Format to = Other(from);
    SCOPED_TRACE(std::string(Extension(from)) + " -> " + Extension(to));
    Workbook wb = Load(ReadPair(pair.name, from), from);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 19U, 3U, "=SEQUENCE(3)")));  // D20
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_TRUE(wb.sheet(0).cell_at(19U, 3U)->dynamic_array);
    const Saved saved = Save(wb, to);
    ASSERT_FALSE(saved.bytes.empty());
    const Workbook back = Load(saved.bytes, to);
    EXPECT_EQ(VerbatimVms(back), pair.vms);
    EXPECT_TRUE(back.sheet(0).cell_at(19U, 3U)->dynamic_array);
    const std::optional<MetadataModel> written = MetadataOf(saved.bytes, to);
    const std::optional<MetadataModel> excel = MetadataOf(ReadPair(pair.name, to), to);
    ASSERT_TRUE(written.has_value());
    ASSERT_TRUE(excel.has_value());
    EXPECT_EQ(excel->dynamic_type, 0U);
    EXPECT_NE(written->dynamic_type, 0U);
    EXPECT_EQ(written->value_entries, excel->value_entries);
    EXPECT_EQ(written->rich_blocks, excel->rich_blocks);
    EXPECT_FALSE(Part(saved.bytes, MetadataPath(from)).has_value());
    EXPECT_EQ(Occurrences(Text(Part(saved.bytes, "[Content_Types].xml")), "PartName=\"/xl/metadata."), 1U);
  }
}

TEST(VerbatimRichVm, ASameFormatSaveLeavesANewFormulaWithoutCellMetadata) {
  // The kept part has no dynamic-array entry to name, and the save does not rewrite it.
  const Pair& pair = Pairs()[1];  // two_pictures
  Workbook wb = Load(ReadPair(pair.name, Format::kXlsx), Format::kXlsx);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 19U, 3U, "=SEQUENCE(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Saved saved = Save(wb, Format::kXlsx);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(Part(saved.bytes, "xl/metadata.xml"), Part(ReadPair(pair.name, Format::kXlsx), "xl/metadata.xml"));
  for (const std::string& name : Entries(saved.bytes)) {
    if (name.rfind("xl/worksheets/sheet", 0) == 0) {
      EXPECT_EQ(Text(Part(saved.bytes, name)).find(" cm=\""), std::string::npos) << name;
    }
  }
  EXPECT_EQ(VerbatimVms(Load(saved.bytes, Format::kXlsx)), pair.vms);
}

// A metadata part the model refuses, per source format.
std::vector<std::uint8_t> Unconvertible(const char* name, Format format) {
  const std::vector<std::uint8_t> excel = ReadPair(name, format);
  if (format == Format::kXlsx) {
    return Repack(excel,
                  {{"xl/metadata.xml", ReplaceOnce(*Part(excel, "xl/metadata.xml"), "<metadataTypes count=\"1\">",
                                                   "<metadataTypes count=\"2\"><metadataType name=\"XLFUTURE\" "
                                                   "minSupportedVersion=\"120000\"/>")}});
  }
  // The first rich-value block's set word, 0 in every Excel part.
  return Repack(excel, {{"xl/metadata.bin", ReplaceOnce(*Part(excel, "xl/metadata.bin"),
                                                        std::string("\x8B\x27\x08\x00\x00\x00\x00\x00", 8),
                                                        std::string("\x8B\x27\x08\x01\x00\x00\x00\x00", 8))}});
}

TEST(VerbatimRichVm, AnUnconvertiblePartRefusesAFormatChangeAndKeepsItsOwnFormat) {
  const Pair& pair = Pairs()[1];  // two_pictures
  for (const Format from : kFormats) {
    SCOPED_TRACE(Extension(from));
    const Workbook wb = Load(Unconvertible(pair.name, from), from);
    const auto refused = TrySave(wb, Other(from));
    ASSERT_FALSE(static_cast<bool>(refused));
    EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoFeatureNotWritable);
    EXPECT_NE(refused.error().context.find("kind=rich_value_metadata"), std::string::npos) << refused.error().context;
    const Saved kept = Save(wb, from);
    ASSERT_FALSE(kept.bytes.empty());
    EXPECT_EQ(VerbatimVms(Load(kept.bytes, from)), pair.vms);
    EXPECT_EQ(kept.diagnostics.deferred_feature_count, 0U);
  }
}

TEST(VerbatimRichVm, AnUnconvertiblePartNoCellNamesIsDroppedAndCounted) {
  const Pair& pair = Pairs()[1];  // two_pictures
  const auto clear = [&pair](Workbook& wb) {
    for (const auto& [at, vm] : pair.vms) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      ASSERT_TRUE(a1::parse_a1_ref(at.second, &row, &col));
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(at.first, row, col, Value::number(1.0))));
    }
  };
  for (const Format from : kFormats) {
    SCOPED_TRACE(Extension(from));
    Workbook control = Load(ReadPair(pair.name, from), from);
    clear(control);
    const Saved converted = Save(control, Other(from));
    Workbook wb = Load(Unconvertible(pair.name, from), from);
    clear(wb);
    const Saved dropped = Save(wb, Other(from));
    ASSERT_FALSE(dropped.bytes.empty());
    EXPECT_FALSE(Part(dropped.bytes, MetadataPath(from)).has_value());
    EXPECT_FALSE(Part(dropped.bytes, MetadataPath(Other(from))).has_value());
    EXPECT_EQ(dropped.diagnostics.dropped_part_count, converted.diagnostics.dropped_part_count + 1U);
    EXPECT_EQ(dropped.diagnostics.deferred_feature_count, converted.diagnostics.deferred_feature_count + 1U);
  }
}

TEST(VerbatimRichVm, ARichErrorSavedPlainIsCountedOnce) {
  const Pair& pair = Pairs()[0];  // one_picture
  for (const Format format : kFormats) {
    SCOPED_TRACE(Extension(format));
    Workbook wb = Load(ReadPair(pair.name, format), format);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 19U, 3U, "=SEQUENCE(3)")));  // D20, blocked by D21
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 20U, 3U, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 19U, 5U, "=SEQUENCE(3)")));  // F20, blocked by F22
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 21U, 5U, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    const Saved saved = Save(wb, format);
    ASSERT_FALSE(saved.bytes.empty());
    EXPECT_EQ(saved.diagnostics.deferred_feature_count, 1U);
    EXPECT_EQ(VerbatimVms(Load(saved.bytes, format)), pair.vms);
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
