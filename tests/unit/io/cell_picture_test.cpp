// Pictures in cells: placing, reading and clearing one, the media they share, the refusals that leave
// a cell as it was, and the save and reload of what was placed.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/image_header.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr std::string_view kStructuresPart = "xl/richData/rdrichvaluestructure.xml";
constexpr std::string_view kValuesPart = "xl/richData/rdrichvalue.xml";
constexpr std::string_view kRelRelsPart = "xl/richData/_rels/richValueRel.xml.rels";
constexpr std::string_view kPlainStructure =
    "<s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/></s>";
constexpr std::string_view kAltStructure =
    "<s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/>"
    "<k n=\"Text\" t=\"s\"/></s>";
constexpr std::uint8_t kMarkerPng = 0x11;

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

void PutBe32(Bytes& out, std::uint32_t v) {
  for (int shift = 24; shift >= 0; shift -= 8) {
    out.push_back(static_cast<std::uint8_t>(v >> shift));
  }
}

/// A PNG whose header probes as `w` x `h`; `tag` makes otherwise equal images differ.
Bytes Png(std::uint32_t w, std::uint32_t h, std::uint8_t tag = kMarkerPng) {
  Bytes out = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  PutBe32(out, w);
  PutBe32(out, h);
  out.insert(out.end(), {8, 6, 0, 0, 0, tag, tag, tag, tag});
  return out;
}

/// A JPEG that probes as `w` x `h` (SOI, one frame header).
Bytes Jpeg(std::uint16_t w, std::uint16_t h) {
  return {0xFF,
          0xD8,
          0xFF,
          0xC0,
          0x00,
          0x0B,
          0x08,
          static_cast<std::uint8_t>(h >> 8),
          static_cast<std::uint8_t>(h),
          static_cast<std::uint8_t>(w >> 8),
          static_cast<std::uint8_t>(w),
          0x01,
          0x00,
          0xFF,
          0xD9};
}

Bytes Gif() {
  return {'G', 'I', 'F', '8', '9', 'a', 4, 0, 4, 0, 0, 0, 0, ';'};
}

std::string FixturePath(std::string_view name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/cell_pictures/" + std::string(name) + ".xlsx";
}

Bytes ReadFixture(std::string_view name) {
  Bytes bytes = test::read_file_bytes(FixturePath(name));
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

std::optional<Bytes> PartBytes(const Bytes& package, std::string_view part) {
  ZipReader zip;
  if (!zip.open(test::span_of(package)) || !zip.has_entry(part)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(part);
  return bytes ? std::optional<Bytes>(std::move(bytes.value())) : std::nullopt;
}

std::string PartText(const Bytes& package, std::string_view part) {
  const std::optional<Bytes> bytes = PartBytes(package, part);
  return bytes ? std::string(bytes->begin(), bytes->end()) : std::string();
}

std::vector<std::string> MediaParts(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  std::vector<std::string> out;
  for (const std::string& name : zip.list_entries()) {
    if (name.compare(0, 9, "xl/media/") == 0) {
      out.push_back(name);
    }
  }
  return out;
}

Workbook Load(const Bytes& bytes) {
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

Bytes Save(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : Bytes();
}

Expected<void, Error> Place(Workbook& wb, const char* cell, const Bytes& image,
                            std::optional<std::string_view> alt = std::nullopt) {
  return wb.set_cell_picture(0, Row(cell), Col(cell), image.data(), image.size(), alt);
}

std::optional<CellPicture> Get(const Workbook& wb, const char* cell) {
  return wb.cell_picture(0, Row(cell), Col(cell));
}

TEST(CellPicture, PlacesPngAndJpegWithAndWithoutAltText) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(48, 30);
  const Bytes jpeg = Jpeg(64, 40);
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A1", png)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "B1", png, "a blue circle")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "C1", jpeg)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "D1", jpeg, "a photo")));

  const std::optional<CellPicture> a1 = Get(wb, "A1");
  ASSERT_TRUE(a1.has_value());
  EXPECT_EQ(a1->bytes, png);
  EXPECT_FALSE(a1->alt_text.has_value());
  EXPECT_EQ(a1->format, ImageFormat::kPng);
  EXPECT_EQ(a1->width, 48U);
  EXPECT_EQ(a1->height, 30U);
  EXPECT_EQ(Get(wb, "B1")->alt_text, std::optional<std::string>("a blue circle"));
  const std::optional<CellPicture> c1 = Get(wb, "C1");
  ASSERT_TRUE(c1.has_value());
  EXPECT_EQ(c1->bytes, jpeg);
  EXPECT_EQ(c1->format, ImageFormat::kJpeg);
  EXPECT_EQ(c1->width, 64U);
  EXPECT_EQ(c1->height, 40U);
  EXPECT_EQ(Get(wb, "D1")->alt_text, std::optional<std::string>("a photo"));
  EXPECT_FALSE(Get(wb, "E1").has_value());
  EXPECT_FALSE(wb.cell_picture(5, 0, 0).has_value());
  // The cell caches #VALUE! and keeps it after the picture is recorded.
  EXPECT_EQ(wb.sheet(0).cell_at(Row("A1"), Col("A1"))->cached_value.as_error(), ErrorCode::Value);
}

TEST(CellPicture, SavedPicturesReloadWithTheFixtureStructure) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(48, 30);
  const Bytes jpeg = Jpeg(64, 40);
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A1", png)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "B1", png, "a blue circle")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "C1", jpeg)));
  const Bytes saved = Save(wb);
  ASSERT_FALSE(saved.empty());

  const std::string structures = PartText(saved, kStructuresPart);
  EXPECT_NE(structures.find(kPlainStructure), std::string::npos) << structures;
  EXPECT_NE(structures.find(kAltStructure), std::string::npos) << structures;
  const std::string values = PartText(saved, kValuesPart);
  EXPECT_NE(values.find("<rv s=\"0\"><v>0</v><v>5</v></rv>"), std::string::npos) << values;
  EXPECT_NE(values.find("<rv s=\"1\"><v>0</v><v>5</v><v>a blue circle</v></rv>"), std::string::npos) << values;
  EXPECT_NE(values.find("<rv s=\"0\"><v>1</v><v>5</v></rv>"), std::string::npos) << values;
  EXPECT_NE(PartText(saved, kRelRelsPart).find("../media/image"), std::string::npos);

  Workbook reloaded = Load(saved);
  EXPECT_EQ(Get(reloaded, "A1")->bytes, png);
  EXPECT_EQ(Get(reloaded, "B1")->bytes, png);
  EXPECT_EQ(Get(reloaded, "B1")->alt_text, std::optional<std::string>("a blue circle"));
  EXPECT_FALSE(Get(reloaded, "A1")->alt_text.has_value());
  EXPECT_EQ(Get(reloaded, "C1")->bytes, jpeg);
  EXPECT_EQ(Get(reloaded, "C1")->format, ImageFormat::kJpeg);
  EXPECT_EQ(Get(reloaded, "C1")->width, 64U);
  // A value recorded on a loaded book saves the same way again.
  EXPECT_EQ(PartText(Save(reloaded), kValuesPart), values);
}

TEST(CellPicture, IdenticalBytesShareOneMediaPart) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(48, 30);
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A1", png)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A2", png)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A3", Png(48, 30, 0x22))));
  const Bytes saved = Save(wb);
  EXPECT_EQ(MediaParts(saved).size(), 2U);
  Workbook reloaded = Load(saved);
  EXPECT_EQ(Get(reloaded, "A2")->bytes, png);
  EXPECT_NE(Get(reloaded, "A3")->bytes, png);
}

TEST(CellPicture, AnImageAlreadyInTheBookIsReused) {
  const Bytes fixture = ReadFixture("cell_pictures");
  Workbook wb = Load(fixture);
  const std::optional<Bytes> media = PartBytes(fixture, "xl/media/image1.png");
  ASSERT_TRUE(media.has_value());
  ASSERT_TRUE(static_cast<bool>(Place(wb, "G2", *media)));
  EXPECT_EQ(MediaParts(Save(wb)).size(), 2U);
}

TEST(CellPicture, AnOverwrittenOrClearedPictureLeavesNoMedia) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A1", Png(48, 30))));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "B1", Png(20, 20))));
  ASSERT_TRUE(static_cast<bool>(Place(wb, "C1", Png(10, 10), "kept")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("A1"), Col("A1"), Value::number(3.0))));
  EXPECT_FALSE(Get(wb, "A1").has_value());
  ASSERT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, Row("B1"), Col("B1"))));
  EXPECT_FALSE(Get(wb, "B1").has_value());
  const Cell* cleared = wb.sheet(0).cell_at(Row("B1"), Col("B1"));
  EXPECT_TRUE(cleared == nullptr || cleared->cached_value.is_blank());
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("B1"), Col("B1")).has_value());
  EXPECT_TRUE(Get(wb, "C1").has_value());
  EXPECT_EQ(MediaParts(Save(wb)).size(), 1U);

  ASSERT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, Row("C1"), Col("C1"))));
  const Bytes empty = Save(wb);
  EXPECT_TRUE(MediaParts(empty).empty());
  EXPECT_TRUE(PartText(empty, kValuesPart).empty());
}

TEST(CellPicture, ClearingACellWithoutAPictureChangesNothing) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(9.0))));
  ASSERT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, 0, 0)));
  ASSERT_NE(wb.sheet(0).cell_at(0, 0), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 9.0);
  EXPECT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, 40, 40)));
}

TEST(CellPicture, ReadsPicturesLoadedFromExcelFiles) {
  const Bytes fixture = ReadFixture("cell_pictures");
  Workbook wb = Load(fixture);
  const std::optional<CellPicture> e2 = Get(wb, "E2");
  ASSERT_TRUE(e2.has_value());
  EXPECT_EQ(e2->bytes, *PartBytes(fixture, "xl/media/image1.png"));
  EXPECT_EQ(e2->format, ImageFormat::kPng);
  EXPECT_GT(e2->width, 0U);
  EXPECT_FALSE(e2->alt_text.has_value());
  const std::optional<CellPicture> e8 = Get(wb, "E8");
  ASSERT_TRUE(e8.has_value());
  EXPECT_EQ(e8->bytes, *PartBytes(fixture, "xl/media/image2.png"));
  EXPECT_EQ(e8->alt_text, std::optional<std::string>("Blue circle on a blue field"));
  EXPECT_FALSE(Get(wb, "E3").has_value());

  Workbook formulas = Load(ReadFixture("picture_formula_refs"));
  const std::optional<CellPicture> g2 = Get(formulas, "G2");  // =E2, a formula's picture
  ASSERT_TRUE(g2.has_value());
  EXPECT_EQ(g2->bytes, Get(formulas, "E2")->bytes);
}

TEST(CellPicture, RefusalsLeaveTheCellAsItWas) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("A1"), Col("A1"), Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(0, Row("C1"), Col("C1"), 2, 2, "=1")));
  // The block's members are known once it is committed.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const std::size_t parts = wb.passthrough_parts().size();
  const Bytes png = Png(8, 8);
  const auto expect_unchanged = [&](const char* cell) {
    SCOPED_TRACE(cell);
    EXPECT_FALSE(Get(wb, cell).has_value());
    EXPECT_FALSE(wb.sheet(0).rich_value_at(Row(cell), Col(cell)).has_value());
    EXPECT_EQ(wb.passthrough_parts().size(), parts);
  };

  const Bytes gif = Gif();
  auto refused = Place(wb, "A1", gif);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoImageUnsupported);
  const std::uint8_t one = 0;
  refused = wb.set_cell_picture(0, Row("A1"), Col("A1"), &one, 0, std::nullopt);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoImageUnsupported);
  // The size is checked before the bytes are read.
  refused = wb.set_cell_picture(0, Row("A1"), Col("A1"), &one, 32U * 1024U * 1024U + 1U, std::nullopt);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kSecurityResourceLimit);
  refused = Place(wb, "A1", png, std::string_view());
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_FALSE(static_cast<bool>(wb.set_cell_picture(0, Row("A1"), Col("A1"), nullptr, 4, std::nullopt)));
  expect_unchanged("A1");
  ASSERT_NE(wb.sheet(0).cell_at(Row("A1"), Col("A1")), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(Row("A1"), Col("A1"))->cached_value.as_number(), 42.0);

  // A member of an array block is not written.
  EXPECT_FALSE(static_cast<bool>(Place(wb, "D2", png)));
  expect_unchanged("D2");
  EXPECT_FALSE(static_cast<bool>(wb.set_cell_picture(9, 0, 0, png.data(), png.size(), std::nullopt)));
}

TEST(CellPicture, AFormulaReadingAPictureCellSeesValueError) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, "A1", Png(8, 8), "alt")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, Row("B1"), Col("B1"), "=A1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, Row("B2"), Col("B2"), "=ISERROR(A1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* ref = wb.sheet(0).cell_at(Row("B1"), Col("B1"));
  ASSERT_NE(ref, nullptr);
  ASSERT_TRUE(ref->cached_value.is_error());
  EXPECT_EQ(ref->cached_value.as_error(), ErrorCode::Value);
  EXPECT_TRUE(wb.sheet(0).cell_at(Row("B2"), Col("B2"))->cached_value.as_boolean());
  // The recalculation leaves the picture itself in place.
  EXPECT_TRUE(Get(wb, "A1").has_value());
}

}  // namespace
}  // namespace io
}  // namespace formulon
