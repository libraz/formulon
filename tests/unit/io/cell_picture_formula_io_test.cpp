// Pictures as formula values in a saved workbook: a load and save with no
// recalc keeps the rich-value parts and every cell's `vm` as Excel wrote them,
// a picture a formula returns is written from its cached value in both
// containers, and one the table cannot carry is written as a plain `#VALUE!`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "passthrough_part.h"
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

std::vector<std::uint8_t> ReadPicture(std::string_view name) {
  std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cell_pictures/" + std::string(name) + ".xlsx");
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

/// The text of `part`, or nullopt when the package has none.
std::optional<std::string> Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  ZipReader zip;
  if (!zip.open(test::span_of(package)) || !zip.has_entry(part)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(part);
  if (!bytes) {
    return std::nullopt;
  }
  return std::string(bytes.value().begin(), bytes.value().end());
}

/// `vm` of every `<c>` in `sheet_part` that carries one, by A1 reference.
std::map<std::string, std::string> VmCells(const std::vector<std::uint8_t>& package, std::string_view sheet_part) {
  std::map<std::string, std::string> out;
  const std::string xml = Part(package, sheet_part).value_or("");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << sheet_part;
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      if (c.attribute("vm")) {
        out[c.attribute("r").value()] = c.attribute("vm").value();
      }
    }
  }
  return out;
}

constexpr std::string_view kValuesPart = "xl/richData/rdrichvalue.xml";
constexpr std::string_view kStructuresPart = "xl/richData/rdrichvaluestructure.xml";
constexpr std::string_view kMetadataPart = "xl/metadata.xml";
constexpr std::string_view kSheetPart = "xl/worksheets/sheet1.xml";
constexpr std::string_view kSheetBin = "xl/worksheets/sheet1.bin";
constexpr std::uint8_t kValueErrorWireCode = 0x0F;

std::uint32_t Row(const char* a1) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(a1::parse_a1_ref(a1, &row, &col)) << a1;
  return row;
}

std::uint32_t Col(const char* a1) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(a1::parse_a1_ref(a1, &row, &col)) << a1;
  return col;
}

void PutBe32(std::vector<std::uint8_t>& out, std::uint32_t v) {
  for (int shift = 24; shift >= 0; shift -= 8) {
    out.push_back(static_cast<std::uint8_t>(v >> shift));
  }
}

/// A PNG whose header probes as `w` x `h`; `tag` makes otherwise equal images differ.
std::vector<std::uint8_t> Png(std::uint32_t w, std::uint32_t h, std::uint8_t tag = 0x11) {
  std::vector<std::uint8_t> out = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  PutBe32(out, w);
  PutBe32(out, h);
  out.insert(out.end(), {8, 6, 0, 0, 0, tag, tag, tag, tag});
  return out;
}

/// `package` with each part of `replaced` swapped for its new bytes.
std::vector<std::uint8_t> Repack(const std::vector<std::uint8_t>& package,
                                 const std::map<std::string, std::string>& replaced) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const std::string& name : zip.list_entries()) {
    const auto it = replaced.find(name);
    std::vector<std::uint8_t> bytes;
    if (it != replaced.end()) {
      bytes.assign(it->second.begin(), it->second.end());
    } else {
      auto read = zip.read_entry(name);
      EXPECT_TRUE(static_cast<bool>(read)) << name;
      if (read) {
        bytes = std::move(read.value());
      }
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

/// `source` with the first `from` in `part` replaced by `to`.
std::vector<std::uint8_t> EditPart(const std::vector<std::uint8_t>& source, std::string_view part,
                                   std::string_view from, std::string_view to) {
  std::string text = Part(source, part).value_or("");
  const std::size_t at = text.find(from);
  EXPECT_NE(at, std::string::npos) << from;
  if (at != std::string::npos) {
    text.replace(at, from.size(), to);
  }
  return Repack(source, {{std::string(part), text}});
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto loaded = xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto saved = xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

/// The value cell `a1` of sheet 0 caches (blank when it holds none).
Value Cached(const Workbook& wb, const char* a1) {
  const Cell* cell = wb.sheet(0).cell_at(Row(a1), Col(a1));
  return cell != nullptr ? cell->cached_value : Value::blank();
}

/// The value spill member `index` (row-major, 0 is the anchor) of the region anchored at `anchor` holds.
Value Member(const Workbook& wb, const char* anchor, std::size_t index) {
  const SpillRegion* region = wb.sheet(0).spill_region_at_anchor(Row(anchor), Col(anchor));
  EXPECT_NE(region, nullptr) << anchor;
  return region != nullptr && index < region->cells.size() ? region->cells[index] : Value::blank();
}

/// The picture-table entry `v` names, or null when `v` is no picture.
const PictureEntry* EntryOf(const Workbook& wb, const Value& v) {
  return v.is_picture() ? wb.picture_entry(v.as_picture().picture_id) : nullptr;
}

/// The `t` of every `<c>` in `sheet_part`, by A1 reference.
std::map<std::string, std::string> CellTypes(const std::vector<std::uint8_t>& package, std::string_view sheet_part) {
  std::map<std::string, std::string> out;
  const std::string xml = Part(package, sheet_part).value_or("");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << sheet_part;
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      out[c.attribute("r").value()] = c.attribute("t").value();
    }
  }
  return out;
}

std::vector<std::string> MediaParts(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  std::vector<std::string> out;
  for (const std::string& name : zip.list_entries()) {
    if (name.compare(0, 9, "xl/media/") == 0) {
      out.push_back(name);
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

std::uint32_t U32(const std::vector<std::uint8_t>& payload) {
  return payload.size() < 4U
             ? 0U
             : static_cast<std::uint32_t>(payload[0]) | (static_cast<std::uint32_t>(payload[1]) << 8U) |
                   (static_cast<std::uint32_t>(payload[2]) << 16U) | (static_cast<std::uint32_t>(payload[3]) << 24U);
}

/// The cell record at `(row, col)` of the binary worksheet `part`: its type, the type of the record before it
/// and its first value byte.
struct XlsbCellRecord {
  bool found = false;
  std::uint16_t type = 0;
  std::uint16_t before = 0;
  std::uint8_t value_byte = 0;
};

XlsbCellRecord FindXlsbCell(const std::vector<std::uint8_t>& package, std::string_view part, std::uint32_t row,
                            std::uint32_t col) {
  const std::string text = Part(package, part).value_or("");
  const std::vector<std::uint8_t> bytes(text.begin(), text.end());
  ByteSpan cursor = test::span_of(bytes);
  std::uint32_t current_row = 0;
  std::uint16_t before = 0;
  while (cursor.size > 0U) {
    auto record = xlsb::read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const std::uint16_t type = record.value().type;
    const std::vector<std::uint8_t> payload(record.value().payload.data,
                                            record.value().payload.data + record.value().payload.size);
    if (type == static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtRowHdr)) {
      current_row = U32(payload);
    } else if (type >= static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtCellBlank) &&
               type <= static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtFmlaError) && current_row == row &&
               U32(payload) == col) {
      return {true, type, before, payload.size() > 8U ? payload[8] : std::uint8_t{0}};
    }
    before = type;
  }
  return {};
}

/// Places a picture at E2 of a new book and `=E2` at G2, calculated.
Workbook BookWithPictureAndReference() {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> png = Png(8, 8);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_picture(0, Row("E2"), Col("E2"), png.data(), png.size(), "alt")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, Row("G2"), Col("G2"), "=E2")));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return wb;
}

TEST(CellPictureFormulaIo, XlsxRoundTripIsByteExactForPictureFormulaRefs) {
  // E2 and E4 hold pictures; G2 is `=E2`, cached as E2's picture.
  const std::vector<std::uint8_t> source = ReadPicture("picture_formula_refs");
  auto loaded = read_ooxml(test::span_of(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::vector<std::uint8_t>& out = saved.value();

  for (std::string_view part :
       {"xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/rdrichvaluestructure.xml"}) {
    SCOPED_TRACE(std::string(part));
    const std::optional<std::string> want = Part(source, part);
    ASSERT_TRUE(want.has_value());
    EXPECT_EQ(Part(out, part), want);
  }
  const std::map<std::string, std::string> want = VmCells(source, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(want.count("G2"), 1U);
  EXPECT_EQ(VmCells(out, "xl/worksheets/sheet1.xml"), want);
}

TEST(CellPictureFormulaIo, XlsxLoadCachesPictureOnPlacedAndFormulaCells) {
  // E2 and E4 hold the same picture; G2 is `=E2` and G7 is `=E8`, both cached as their source's picture.
  const Workbook wb = Load(ReadPicture("picture_formula_refs"));
  for (const char* cell : {"E2", "E4", "E6", "E8", "G2", "G7"}) {
    EXPECT_TRUE(Cached(wb, cell).is_picture()) << cell;
    EXPECT_FALSE(wb.sheet(0).rich_value_at(Row(cell), Col(cell)).has_value()) << cell;
  }
  EXPECT_FALSE(wb.sheet(0).has_rich_values());
  EXPECT_EQ(Cached(wb, "E2").as_picture().picture_id, Cached(wb, "E4").as_picture().picture_id);
  EXPECT_EQ(Cached(wb, "E2").picture_origin(), 5U);
  EXPECT_EQ(Cached(wb, "G2").picture_origin(), 4U);
  EXPECT_EQ(Cached(wb, "G2").as_picture().media_id, Cached(wb, "E2").as_picture().media_id);
  const PictureEntry* e2 = EntryOf(wb, Cached(wb, "E2"));
  ASSERT_NE(e2, nullptr);
  EXPECT_EQ(e2->media_path, "xl/media/image1.png");
  EXPECT_FALSE(e2->alt.has_value());
  const PictureEntry* g7 = EntryOf(wb, Cached(wb, "G7"));
  ASSERT_NE(g7, nullptr);
  EXPECT_EQ(g7->media_path, "xl/media/image2.png");
  EXPECT_EQ(g7->alt, std::optional<std::string>("Blue circle on a blue field"));
  // The cells that read as other values stay as they were.
  EXPECT_FALSE(Cached(wb, "G3").is_picture());
}

TEST(CellPictureFormulaIo, XlsxSpillMembersLoadAsPicturesAndKeepTheirRvsOnSave) {
  // AD2:AD4 spills `=E2:E4`: a picture, a blank read as 0, and a picture again.
  const std::vector<std::uint8_t> source = ReadPicture("picture_formula_spill");
  const Workbook wb = Load(source);
  EXPECT_TRUE(Cached(wb, "AD2").is_picture());
  EXPECT_EQ(Cached(wb, "AD2").picture_origin(), 4U);
  EXPECT_TRUE(Member(wb, "AD2", 0).is_picture());
  EXPECT_FALSE(Member(wb, "AD2", 1).is_picture());
  ASSERT_TRUE(Member(wb, "AD2", 2).is_picture());
  EXPECT_EQ(Member(wb, "AD2", 2).picture_origin(), 4U);
  EXPECT_EQ(Member(wb, "AD2", 2).as_picture().picture_id, Member(wb, "AD2", 0).as_picture().picture_id);
  EXPECT_FALSE(wb.sheet(0).has_rich_values());

  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  for (std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart}) {
    SCOPED_TRACE(std::string(part));
    EXPECT_EQ(Part(saved, part), Part(source, part));
  }
  const std::map<std::string, std::string> want = VmCells(source, kSheetPart);
  EXPECT_EQ(want.count("AD4"), 1U);
  EXPECT_EQ(VmCells(saved, kSheetPart), want);
}

TEST(CellPictureFormulaIo, XlsxSortedSpillOfMixedOriginsRoundTripsByteExact) {
  const std::vector<std::uint8_t> source = ReadPicture("picture_formula_origins");
  const Workbook wb = Load(source);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  for (std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart}) {
    SCOPED_TRACE(std::string(part));
    EXPECT_EQ(Part(saved, part), Part(source, part));
  }
  EXPECT_EQ(VmCells(saved, kSheetPart), VmCells(source, kSheetPart));
  EXPECT_EQ(MediaParts(saved), MediaParts(source));
}

TEST(CellPictureFormulaIo, WriterDerivesFormulaRvFromThePictureValue) {
  const Workbook wb = BookWithPictureAndReference();
  ASSERT_TRUE(Cached(wb, "G2").is_picture());
  ASSERT_EQ(Cached(wb, "G2").picture_origin(), 4U);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  ASSERT_FALSE(saved.empty());

  const std::map<std::string, std::string> vm = VmCells(saved, kSheetPart);
  ASSERT_EQ(vm.count("E2"), 1U);
  ASSERT_EQ(vm.count("G2"), 1U);
  EXPECT_NE(vm.at("E2"), vm.at("G2"));
  EXPECT_EQ(CellTypes(saved, kSheetPart).at("G2"), "e");
  // The placed picture is origin 5 and the formula's is origin 4, both of the same image and alt.
  const std::string values = Part(saved, kValuesPart).value_or("");
  EXPECT_NE(values.find("<rv s=\"0\"><v>0</v><v>5</v><v>alt</v></rv><rv s=\"0\"><v>0</v><v>4</v><v>alt</v></rv>"),
            std::string::npos)
      << values;

  const Workbook reloaded = Load(saved);
  ASSERT_TRUE(Cached(reloaded, "G2").is_picture());
  EXPECT_EQ(Cached(reloaded, "G2").picture_origin(), 4U);
  EXPECT_EQ(Cached(reloaded, "E2").picture_origin(), 5U);
  const PictureEntry* g2 = EntryOf(reloaded, Cached(reloaded, "G2"));
  ASSERT_NE(g2, nullptr);
  EXPECT_EQ(g2->alt, std::optional<std::string>("alt"));
  EXPECT_EQ(MediaParts(saved).size(), 1U);
}

TEST(CellPictureFormulaIo, XlsbWriterDerivesFormulaRvFromThePictureValue) {
  const Workbook wb = BookWithPictureAndReference();
  ASSERT_TRUE(Cached(wb, "G2").is_picture());
  const std::vector<std::uint8_t> saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.empty());

  const XlsbCellRecord g2 = FindXlsbCell(saved, kSheetBin, Row("G2"), Col("G2"));
  ASSERT_TRUE(g2.found);
  EXPECT_EQ(g2.type, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtFmlaError));
  EXPECT_EQ(g2.before, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(g2.value_byte, kValueErrorWireCode);
  const XlsbCellRecord e2 = FindXlsbCell(saved, kSheetBin, Row("E2"), Col("E2"));
  ASSERT_TRUE(e2.found);
  EXPECT_EQ(e2.type, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtCellError));
  EXPECT_EQ(e2.before, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtValueMeta));

  const Workbook reloaded = LoadXlsb(saved);
  ASSERT_TRUE(Cached(reloaded, "G2").is_picture());
  EXPECT_EQ(Cached(reloaded, "G2").picture_origin(), 4U);
  EXPECT_EQ(Cached(reloaded, "E2").picture_origin(), 5U);
}

TEST(CellPictureFormulaIo, XlsbSpillMembersSurviveAConversionFromXlsx) {
  // No binary fixture: the book is authored from the .xlsx and checked through its own round trip.
  const Workbook source = Load(ReadPicture("picture_formula_spill"));
  const std::vector<std::uint8_t> saved = SaveXlsb(source);
  ASSERT_FALSE(saved.empty());
  for (const char* cell : {"AD2", "AD4"}) {
    const XlsbCellRecord record = FindXlsbCell(saved, kSheetBin, Row(cell), Col(cell));
    ASSERT_TRUE(record.found) << cell;
    EXPECT_EQ(record.type, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtFmlaError)) << cell;
    EXPECT_EQ(record.before, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtValueMeta)) << cell;
  }
  const Workbook reloaded = LoadXlsb(saved);
  for (const char* cell : {"E2", "E4", "E6", "E8", "G2", "G7", "AD2", "AM2"}) {
    SCOPED_TRACE(cell);
    const Value want = Cached(source, cell);
    const Value got = Cached(reloaded, cell);
    ASSERT_EQ(got.is_picture(), want.is_picture());
    if (want.is_picture()) {
      EXPECT_EQ(got.picture_origin(), want.picture_origin());
      const PictureEntry* a = EntryOf(source, want);
      const PictureEntry* b = EntryOf(reloaded, got);
      ASSERT_TRUE(a != nullptr && b != nullptr);
      EXPECT_EQ(b->media_path, a->media_path);
      EXPECT_EQ(b->alt, a->alt);
    }
  }
  ASSERT_TRUE(Member(reloaded, "AD2", 2).is_picture());
  EXPECT_EQ(Member(reloaded, "AD2", 2).picture_origin(), 4U);
}

// A picture placed in a book whose rich-value parts are kept verbatim is refused: the set cannot take a new
// entry, and a plain error would lose the picture on save. A bad image keeps its own refusal.
void ExpectPictureRefused(Workbook& wb) {
  const std::size_t parts = wb.passthrough_parts().size();
  const std::vector<std::uint8_t> png = Png(8, 8);
  const std::vector<std::uint8_t> junk = {0x01, 0x02, 0x03};
  auto bad = wb.set_cell_picture(0, Row("A1"), Col("A1"), junk.data(), junk.size(), std::nullopt);
  ASSERT_FALSE(static_cast<bool>(bad));
  EXPECT_EQ(bad.error().code, FormulonErrorCode::kIoImageUnsupported);
  auto refused = wb.set_cell_picture(0, Row("A1"), Col("A1"), png.data(), png.size(), std::nullopt);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoFeatureNotWritable);
  EXPECT_EQ(refused.error().context, "context=cell_picture kind=rich_value_set_kept_verbatim");
  EXPECT_FALSE(Cached(wb, "A1").is_picture());
  EXPECT_EQ(wb.passthrough_parts().size(), parts);
}

TEST(CellPictureFormulaIo, PictureIntoAVerbatimRichValueSetIsRefused) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  std::string metadata = Part(excel, kMetadataPart).value_or("");
  const std::string types = "<metadataTypes count=\"1\">";
  const std::size_t at = metadata.find(types);
  ASSERT_NE(at, std::string::npos);
  metadata.replace(at, types.size(),
                   "<metadataTypes count=\"2\"><metadataType name=\"XLFUTURE\" minSupportedVersion=\"120000\"/>");
  const std::vector<std::uint8_t> package = Repack(excel, {{std::string(kMetadataPart), metadata}});

  Workbook wb = Load(package);
  ExpectPictureRefused(wb);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  ASSERT_FALSE(saved.empty());
  for (std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart}) {
    SCOPED_TRACE(std::string(part));
    EXPECT_EQ(Part(saved, part), Part(package, part));
  }
}

TEST(CellPictureFormulaIo, DataTableAnchorHoldingAPictureIsAPlainValueError) {
  // What Excel stores for a picture in a data-table anchor is unmeasured; the anchor is a plain error with no vm.
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> png = Png(8, 8);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0, Row("A1"), Col("A1"), png.data(), png.size(), std::nullopt)));
  DataTableInput input;
  input.present = true;
  input.row = 0U;
  input.col = 1U;
  ASSERT_TRUE(static_cast<bool>(wb.add_data_table(0, MergeRange{1U, 3U, 3U, 3U}, DataTableInput{}, input)));
  wb.sheet(0).set_cell_cached_value(1U, 3U, Cached(wb, "A1"));
  ASSERT_TRUE(Cached(wb, "D2").is_picture());

  const std::vector<std::uint8_t> xlsx = SaveXlsx(wb);
  ASSERT_FALSE(xlsx.empty());
  EXPECT_EQ(CellTypes(xlsx, kSheetPart).at("D2"), "e");
  const std::map<std::string, std::string> vm = VmCells(xlsx, kSheetPart);
  EXPECT_EQ(vm.count("A1"), 1U);
  EXPECT_EQ(vm.count("D2"), 0U);

  const std::vector<std::uint8_t> xlsb_bytes = SaveXlsb(wb);
  ASSERT_FALSE(xlsb_bytes.empty());
  const XlsbCellRecord anchor = FindXlsbCell(xlsb_bytes, kSheetBin, Row("D2"), Col("D2"));
  ASSERT_TRUE(anchor.found);
  EXPECT_EQ(anchor.type, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtCellError));
  EXPECT_NE(anchor.before, static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(anchor.value_byte, kValueErrorWireCode);
}

TEST(CellPictureFormulaIo, DanglingMediaWritesPlainError) {
  Workbook wb = Load(ReadPicture("cell_pictures"));
  ASSERT_TRUE(Cached(wb, "E2").is_picture());
  std::vector<PassthroughPart> parts = wb.passthrough_parts();
  parts.erase(std::remove_if(parts.begin(), parts.end(),
                             [](const PassthroughPart& part) { return part.path == "xl/media/image1.png"; }),
              parts.end());
  wb.set_passthrough_parts(std::move(parts));

  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  ASSERT_FALSE(saved.empty());
  // E2 and E4 named the missing image; E6 and E8 name the other one and keep their rvs.
  const std::map<std::string, std::string> vm = VmCells(saved, kSheetPart);
  EXPECT_EQ(vm.count("E2"), 0U);
  EXPECT_EQ(vm.count("E4"), 0U);
  EXPECT_EQ(vm.count("E6"), 1U);
  EXPECT_EQ(vm.count("E8"), 1U);
  EXPECT_EQ(CellTypes(saved, kSheetPart).at("E2"), "e");
  EXPECT_EQ(MediaParts(saved), (std::vector<std::string>{"xl/media/image2.png"}));

  const Workbook reloaded = Load(saved);
  EXPECT_FALSE(Cached(reloaded, "E2").is_picture());
  EXPECT_TRUE(Cached(reloaded, "E6").is_picture());
}

TEST(CellPictureFormulaIo, ClearedSourceDropsMediaWhileResultStillCached) {
  // G2 is `=E2`; image1 is named by E2, E4 and G2 only.
  Workbook wb = Load(ReadPicture("picture_formula_refs"));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E4"), Col("E4"), Value::blank())));
  ASSERT_TRUE(Cached(wb, "G2").is_picture());

  // Without a recalculation the formula still caches the picture, so its image is kept.
  const std::vector<std::uint8_t> before = SaveXlsx(wb);
  ASSERT_FALSE(before.empty());
  EXPECT_EQ(VmCells(before, kSheetPart).count("G2"), 1U);
  EXPECT_EQ(VmCells(before, kSheetPart).count("E2"), 0U);
  EXPECT_EQ(MediaParts(before), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_FALSE(Cached(wb, "G2").is_picture());
  const std::vector<std::uint8_t> after = SaveXlsx(wb);
  ASSERT_FALSE(after.empty());
  EXPECT_EQ(VmCells(after, kSheetPart).count("G2"), 0U);
  EXPECT_EQ(MediaParts(after), (std::vector<std::string>{"xl/media/image2.png"}));
}

TEST(CellPictureFormulaIo, OutOfRangeCalcOriginStaysOpaque) {
  const std::vector<std::uint8_t> package =
      EditPart(ReadPicture("cell_pictures"), kValuesPart, "<rv s=\"0\"><v>0</v><v>5</v></rv>",
               "<rv s=\"0\"><v>0</v><v>300</v></rv>");
  const Workbook wb = Load(package);
  for (const char* cell : {"E2", "E4"}) {
    SCOPED_TRACE(cell);
    EXPECT_FALSE(Cached(wb, cell).is_picture());
    EXPECT_TRUE(wb.sheet(0).rich_value_at(Row(cell), Col(cell)).has_value());
  }
  EXPECT_TRUE(Cached(wb, "E6").is_picture());
  EXPECT_TRUE(Cached(wb, "E8").is_picture());
  // The value is written back as it was read.
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  EXPECT_NE(Part(saved, kValuesPart).value_or("").find("<v>300</v>"), std::string::npos);
  EXPECT_EQ(VmCells(saved, kSheetPart).count("E2"), 1U);
}

TEST(CellPictureFormulaIo, NonNumericCalcOriginStaysOpaque) {
  const std::vector<std::uint8_t> package =
      EditPart(ReadPicture("cell_pictures"), kValuesPart, "<rv s=\"0\"><v>0</v><v>5</v></rv>",
               "<rv s=\"0\"><v>0</v><v>x</v></rv>");
  const Workbook wb = Load(package);
  for (const char* cell : {"E2", "E4"}) {
    SCOPED_TRACE(cell);
    EXPECT_FALSE(Cached(wb, cell).is_picture());
    EXPECT_TRUE(wb.sheet(0).rich_value_at(Row(cell), Col(cell)).has_value());
  }
  EXPECT_TRUE(Cached(wb, "E6").is_picture());
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  EXPECT_NE(Part(saved, kValuesPart).value_or("").find("<v>x</v>"), std::string::npos);
}

TEST(CellPictureFormulaIo, AltTrailingSpaceAndEmptyTextAreNormalisedOnLoad) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  {
    const Workbook wb =
        Load(EditPart(excel, kValuesPart, "<v>Blue circle on a blue field</v>", "<v>Blue circle on a blue field </v>"));
    const PictureEntry* e8 = EntryOf(wb, Cached(wb, "E8"));
    ASSERT_NE(e8, nullptr);
    EXPECT_EQ(e8->alt, std::optional<std::string>("Blue circle on a blue field"));
    const std::string values = Part(SaveXlsx(wb), kValuesPart).value_or("");
    EXPECT_NE(values.find("<v>Blue circle on a blue field</v>"), std::string::npos) << values;
  }
  {
    const Workbook wb = Load(EditPart(excel, kValuesPart, "<v>Blue circle on a blue field</v>", "<v></v>"));
    const PictureEntry* e8 = EntryOf(wb, Cached(wb, "E8"));
    ASSERT_NE(e8, nullptr);
    EXPECT_FALSE(e8->alt.has_value());
    const std::vector<std::uint8_t> saved = SaveXlsx(wb);
    EXPECT_EQ(Part(saved, kValuesPart).value_or("").find("<k n=\"Text\""), std::string::npos);
    EXPECT_EQ(Part(saved, kStructuresPart).value_or("").find("Text"), std::string::npos);
    EXPECT_EQ(VmCells(saved, kSheetPart).count("E8"), 1U);
  }
}

// There is no Excel-saved binary fixture with formula pictures: the binary tests below are self-consistency and
// round-trip checks over books authored from the .xlsx fixtures or the binary picture fixture.

std::vector<std::uint8_t> ReadXlsbPicture(std::string_view name) {
  std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cell_pictures/" + std::string(name) + ".xlsb");
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

std::uint16_t Type(xlsb::XlsbRecordType t) {
  return static_cast<std::uint16_t>(t);
}

TEST(CellPictureFormulaIo, XlsbDanglingMediaWritesPlainError) {
  Workbook wb = LoadXlsb(ReadXlsbPicture("cell_pictures"));
  ASSERT_TRUE(Cached(wb, "E2").is_picture());
  std::vector<PassthroughPart> parts = wb.passthrough_parts();
  parts.erase(std::remove_if(parts.begin(), parts.end(),
                             [](const PassthroughPart& part) { return part.path == "xl/media/image1.png"; }),
              parts.end());
  wb.set_passthrough_parts(std::move(parts));

  const std::vector<std::uint8_t> saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.empty());
  const XlsbCellRecord e2 = FindXlsbCell(saved, kSheetBin, Row("E2"), Col("E2"));
  ASSERT_TRUE(e2.found);
  EXPECT_EQ(e2.type, Type(xlsb::XlsbRecordType::BrtCellError));
  EXPECT_NE(e2.before, Type(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(FindXlsbCell(saved, kSheetBin, Row("E6"), Col("E6")).before, Type(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(MediaParts(saved), (std::vector<std::string>{"xl/media/image2.png"}));
  // Only the two values naming the remaining image are written.
  EXPECT_NE(Part(saved, kValuesPart).value_or("").find("count=\"2\""), std::string::npos);
  const Workbook reloaded = LoadXlsb(saved);
  EXPECT_FALSE(Cached(reloaded, "E2").is_picture());
  EXPECT_TRUE(Cached(reloaded, "E6").is_picture());
}

TEST(CellPictureFormulaIo, XlsbPictureIntoAVerbatimRichValueSetIsRefused) {
  // An external picture target keeps the rich-value parts verbatim.
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("cell_pictures");
  const std::vector<std::uint8_t> package =
      EditPart(excel, "xl/richData/_rels/richValueRel.xml.rels", "Target=\"../media/image1.png\"",
               "Target=\"https://example.com/a.png\" TargetMode=\"External\"");
  Workbook wb = LoadXlsb(package);
  ExpectPictureRefused(wb);
  const std::vector<std::uint8_t> saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.empty());
  for (std::string_view part : {kValuesPart, kStructuresPart}) {
    SCOPED_TRACE(std::string(part));
    EXPECT_EQ(Part(saved, part), Part(package, part));
  }
}

TEST(CellPictureFormulaIo, XlsbClearedSourceDropsMediaWhileResultStillCached) {
  Workbook wb = Load(ReadPicture("picture_formula_refs"));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E4"), Col("E4"), Value::blank())));
  const std::vector<std::uint8_t> before = SaveXlsb(wb);
  ASSERT_FALSE(before.empty());
  EXPECT_EQ(FindXlsbCell(before, kSheetBin, Row("G2"), Col("G2")).before, Type(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(MediaParts(before), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const std::vector<std::uint8_t> after = SaveXlsb(wb);
  ASSERT_FALSE(after.empty());
  EXPECT_NE(FindXlsbCell(after, kSheetBin, Row("G2"), Col("G2")).before, Type(xlsb::XlsbRecordType::BrtValueMeta));
  EXPECT_EQ(MediaParts(after), (std::vector<std::string>{"xl/media/image2.png"}));
}

TEST(CellPictureFormulaIo, XlsbOriginAndAltNormalisationSurviveAConversion) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  {
    const Workbook wb =
        Load(EditPart(excel, kValuesPart, "<v>Blue circle on a blue field</v>", "<v>Blue circle on a blue field </v>"));
    const Workbook reloaded = LoadXlsb(SaveXlsb(wb));
    const PictureEntry* e8 = EntryOf(reloaded, Cached(reloaded, "E8"));
    ASSERT_NE(e8, nullptr);
    EXPECT_EQ(e8->alt, std::optional<std::string>("Blue circle on a blue field"));
  }
  {
    const Workbook wb = Load(EditPart(excel, kValuesPart, "<v>Blue circle on a blue field</v>", "<v></v>"));
    const std::vector<std::uint8_t> saved = SaveXlsb(wb);
    EXPECT_EQ(Part(saved, kStructuresPart).value_or("").find("Text"), std::string::npos);
    const Workbook reloaded = LoadXlsb(saved);
    const PictureEntry* e8 = EntryOf(reloaded, Cached(reloaded, "E8"));
    ASSERT_NE(e8, nullptr);
    EXPECT_FALSE(e8->alt.has_value());
  }
  {
    const Workbook wb =
        Load(EditPart(excel, kValuesPart, "<rv s=\"0\"><v>0</v><v>5</v></rv>", "<rv s=\"0\"><v>0</v><v>300</v></rv>"));
    const std::vector<std::uint8_t> saved = SaveXlsb(wb);
    EXPECT_NE(Part(saved, kValuesPart).value_or("").find("<v>300</v>"), std::string::npos);
    const Workbook reloaded = LoadXlsb(saved);
    EXPECT_FALSE(Cached(reloaded, "E2").is_picture());
    EXPECT_TRUE(Cached(reloaded, "E6").is_picture());
  }
}

TEST(CellPictureFormulaIo, EqualBytesShareOneMediaPartAndMediaId) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> png = Png(8, 8);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0, Row("A1"), Col("A1"), png.data(), png.size(), "first")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0, Row("B1"), Col("B1"), png.data(), png.size(), "second")));
  const Value a1 = Cached(wb, "A1");
  const Value b1 = Cached(wb, "B1");
  ASSERT_TRUE(a1.is_picture() && b1.is_picture());
  EXPECT_EQ(a1.as_picture().media_id, b1.as_picture().media_id);
  EXPECT_NE(a1.as_picture().picture_id, b1.as_picture().picture_id);

  const std::vector<std::uint8_t> xlsx = SaveXlsx(wb);
  const std::vector<std::uint8_t> xlsb_bytes = SaveXlsb(wb);
  for (const std::vector<std::uint8_t>* saved : {&xlsx, &xlsb_bytes}) {
    ASSERT_FALSE(saved->empty());
    EXPECT_EQ(MediaParts(*saved).size(), 1U);
    EXPECT_NE(Part(*saved, kValuesPart).value_or("").find("count=\"2\""), std::string::npos);
  }
  const Workbook from_xlsx = Load(xlsx);
  const Workbook from_xlsb = LoadXlsb(xlsb_bytes);
  for (const Workbook* loaded : {&from_xlsx, &from_xlsb}) {
    ASSERT_TRUE(Cached(*loaded, "A1").is_picture() && Cached(*loaded, "B1").is_picture());
    EXPECT_EQ(Cached(*loaded, "A1").as_picture().media_id, Cached(*loaded, "B1").as_picture().media_id);
    EXPECT_NE(Cached(*loaded, "A1").as_picture().picture_id, Cached(*loaded, "B1").as_picture().picture_id);
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
