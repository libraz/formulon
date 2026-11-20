// The workbook's picture table through `intern_picture` and the cell picture API: one media per
// distinct bytes, byte-order placement of a new media, alt normalisation, and pictures read back
// through recalc.

#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "picture_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

void PutBe32(Bytes& out, std::uint32_t v) {
  for (int shift = 24; shift >= 0; shift -= 8) {
    out.push_back(static_cast<std::uint8_t>(v >> shift));
  }
}

// A PNG whose header probes as 4 x 4; `tag` makes otherwise equal images differ.
Bytes Png(std::uint8_t tag) {
  Bytes out = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  PutBe32(out, 4U);
  PutBe32(out, 4U);
  out.insert(out.end(), {8, 6, 0, 0, 0, tag, tag, tag, tag});
  return out;
}

// Media parts by path, standing in for a package or the part store.
struct FakeStore {
  std::map<std::string, Bytes> parts;
};

MediaSpan FakeBytes(const void* ctx, std::string_view path) {
  const FakeStore& store = *static_cast<const FakeStore*>(ctx);
  const auto found = store.parts.find(std::string(path));
  return found == store.parts.end() ? MediaSpan{} : MediaSpan{found->second.data(), found->second.size()};
}

PicturePayload Intern(Workbook& wb, FakeStore& store, const std::string& path, const Bytes& bytes,
                      std::optional<std::string_view> alt = std::nullopt) {
  store.parts[path] = bytes;
  return wb.intern_picture(path, MediaSpan{store.parts[path].data(), bytes.size()}, alt, &FakeBytes, &store);
}

Expected<void, Error> Place(Workbook& wb, std::uint32_t row, std::uint32_t col, const Bytes& image,
                            std::optional<std::string_view> alt = std::nullopt) {
  return wb.set_cell_picture(0, row, col, image.data(), image.size(), alt);
}

PicturePayload PictureAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Value v = wb.sheet(0).resolve_cell_value(row, col);
  EXPECT_TRUE(v.is_picture()) << "kind " << static_cast<int>(v.kind());
  return v.is_picture() ? v.as_picture() : PicturePayload{};
}

TEST(WorkbookPictureTable, InternFindsTheSameEntryAndKeepsEntryPointers) {
  Workbook wb = Workbook::create();
  FakeStore store;
  const PicturePayload first = Intern(wb, store, "xl/media/image1.png", Png(0x11));
  const PictureEntry* entry = wb.picture_entry(first.picture_id);
  ASSERT_NE(nullptr, entry);

  const PicturePayload again = Intern(wb, store, "xl/media/image1.png", Png(0x11));
  EXPECT_EQ(first.picture_id, again.picture_id);
  EXPECT_EQ(first.media_id, again.media_id);

  const PicturePayload titled = Intern(wb, store, "xl/media/image1.png", Png(0x11), "a circle");
  EXPECT_NE(first.picture_id, titled.picture_id);
  EXPECT_EQ(first.media_id, titled.media_id);

  for (std::uint8_t tag = 0x20; tag < 0x60; ++tag) {
    Intern(wb, store, "xl/media/image" + std::to_string(tag) + ".png", Png(tag));
  }
  EXPECT_EQ(2U + 0x40U, wb.picture_count());
  EXPECT_EQ(entry, wb.picture_entry(first.picture_id));
  EXPECT_EQ("xl/media/image1.png", entry->media_path);
  EXPECT_FALSE(entry->alt.has_value());
  EXPECT_EQ(nullptr, wb.picture_entry(static_cast<std::uint32_t>(wb.picture_count())));
}

// Identical bytes in two parts are one media (p77, i1); the entry keeps the first part's path.
TEST(WorkbookPictureTable, IdenticalBytesInTwoPartsShareOneMedia) {
  Workbook wb = Workbook::create();
  FakeStore store;
  const PicturePayload a = Intern(wb, store, "xl/media/image1.png", Png(0x11));
  const PicturePayload b = Intern(wb, store, "xl/media/image2.png", Png(0x11));
  EXPECT_EQ(a.media_id, b.media_id);
  EXPECT_EQ(a.picture_id, b.picture_id);
  EXPECT_EQ(1U, wb.picture_table().media_count());
  EXPECT_EQ("xl/media/image1.png", wb.picture_entry(b.picture_id)->media_path);
}

// Equal pixels in different bytes are different media (i3).
TEST(WorkbookPictureTable, DifferentBytesAreDifferentMedia) {
  Workbook wb = Workbook::create();
  FakeStore store;
  const PicturePayload a = Intern(wb, store, "xl/media/image1.png", Png(0x11));
  const PicturePayload b = Intern(wb, store, "xl/media/image2.png", Png(0x12));
  EXPECT_NE(a.media_id, b.media_id);
  EXPECT_EQ(2U, wb.picture_table().media_count());
}

// A media whose key matches but whose bytes differ is a new media.
TEST(WorkbookPictureTable, DigestCollisionIsResolvedByBytes) {
  Workbook wb = Workbook::create();
  FakeStore store;
  const Bytes interned = Png(0x11);
  const Bytes held = Png(0x22);
  store.parts["xl/media/held.png"] = held;
  const std::uint32_t colliding =
      wb.picture_table().add_media_at(0, media_key_of(interned.data(), interned.size()), "xl/media/held.png");

  const PicturePayload p = Intern(wb, store, "xl/media/image1.png", interned);
  EXPECT_NE(colliding, p.media_id);
  EXPECT_EQ(2U, wb.picture_table().media_count());
  // Placed against the held bytes: 0x22 is greater, so the held media sorts first.
  EXPECT_EQ(0U, wb.picture_table().media_rank(colliding));
  EXPECT_EQ(1U, wb.picture_table().media_rank(p.media_id));
}

// Pinned: trailing U+0020 is removed and an alt left empty is absent (P6, P7); case and leading
// spaces are kept.
TEST(WorkbookPictureTable, AltTrailingSpaceAndEmptyAreNormalised) {
  Workbook wb = Workbook::create();
  const Bytes png = Png(0x11);
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, png, "Picture ")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 1, 0, png, "Picture")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 2, 0, png, "")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 3, 0, png, "   ")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 4, 0, png)));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 5, 0, png, "PICTURE")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 6, 0, png, " Picture")));

  const std::optional<CellPicture> trimmed = wb.cell_picture(0, 0, 0);
  ASSERT_TRUE(trimmed.has_value());
  EXPECT_EQ(std::optional<std::string>("Picture"), trimmed->alt_text);
  EXPECT_EQ(PictureAt(wb, 0, 0).picture_id, PictureAt(wb, 1, 0).picture_id);

  ASSERT_TRUE(wb.cell_picture(0, 2, 0).has_value());
  EXPECT_FALSE(wb.cell_picture(0, 2, 0)->alt_text.has_value());
  EXPECT_EQ(PictureAt(wb, 2, 0).picture_id, PictureAt(wb, 4, 0).picture_id);
  EXPECT_EQ(PictureAt(wb, 3, 0).picture_id, PictureAt(wb, 4, 0).picture_id);

  EXPECT_NE(PictureAt(wb, 1, 0).picture_id, PictureAt(wb, 5, 0).picture_id);
  EXPECT_EQ(std::optional<std::string>(" Picture"), wb.cell_picture(0, 6, 0)->alt_text);
  EXPECT_EQ(1U, wb.picture_table().media_count());
}

// A new media is placed by binary search over the ranked media: the greater bytes first, and of a
// sequence and its prefix the longer first.
TEST(WorkbookPictureTable, NewMediaIsPlacedByByteOrder) {
  Workbook wb = Workbook::create();
  FakeStore store;
  const std::uint32_t low = Intern(wb, store, "xl/media/a.bin", Bytes{0x10}).media_id;
  const std::uint32_t high = Intern(wb, store, "xl/media/b.bin", Bytes{0x30}).media_id;
  const std::uint32_t mid = Intern(wb, store, "xl/media/c.bin", Bytes{0x20}).media_id;
  const std::uint32_t longer = Intern(wb, store, "xl/media/d.bin", Bytes{0x30, 0x00}).media_id;
  const std::uint32_t unsigned_high = Intern(wb, store, "xl/media/e.bin", Bytes{0x89}).media_id;
  const PictureTable& table = wb.picture_table();
  EXPECT_EQ(0U, table.media_rank(unsigned_high));
  EXPECT_EQ(1U, table.media_rank(longer));
  EXPECT_EQ(2U, table.media_rank(high));
  EXPECT_EQ(3U, table.media_rank(mid));
  EXPECT_EQ(4U, table.media_rank(low));
}

// A ranked media whose part was removed compares by its key, and its path is not reused.
TEST(WorkbookPictureTable, RankedMediaWithoutBytesComparesByKey) {
  Workbook wb = Workbook::create();
  const Bytes first = Png(0x11);
  const Bytes second = Png(0x22);
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, first)));
  const PicturePayload a = PictureAt(wb, 0, 0);
  const std::string a_path = wb.picture_entry(a.picture_id)->media_path;

  std::vector<PassthroughPart> kept;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path != a_path) {
      kept.push_back(part);
    }
  }
  wb.set_passthrough_parts(std::move(kept));
  ASSERT_EQ(nullptr, io::ooxml::find_passthrough_part(wb, a_path));

  ASSERT_TRUE(static_cast<bool>(Place(wb, 1, 0, second)));
  const PicturePayload b = PictureAt(wb, 1, 0);
  EXPECT_NE(a.media_id, b.media_id);
  EXPECT_NE(a_path, wb.picture_entry(b.picture_id)->media_path);
  const int by_key =
      compare_media_keys(media_key_of(second.data(), second.size()), media_key_of(first.data(), first.size()));
  ASSERT_NE(0, by_key);
  EXPECT_EQ(by_key < 0 ? 0U : 1U, wb.picture_table().media_rank(b.media_id));
}

TEST(WorkbookPictureTable, PlacedPictureIsACellValueWithOriginFive) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, Png(0x11), "x")));
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(nullptr, cell);
  ASSERT_TRUE(cell->cached_value.is_picture());
  EXPECT_EQ(5U, cell->cached_value.picture_origin());
  EXPECT_FALSE(wb.sheet(0).rich_value_at(0, 0).has_value());
}

TEST(WorkbookPictureTable, ClearedIdStillResolvesAfterSave) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, Png(0x11))));
  const PicturePayload p = PictureAt(wb, 0, 0);
  ASSERT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, 0, 0)));
  EXPECT_TRUE(wb.sheet(0).resolve_cell_value(0, 0).is_blank());

  ASSERT_TRUE(static_cast<bool>(io::write_ooxml(wb)));
  const PictureEntry* entry = wb.picture_entry(p.picture_id);
  ASSERT_NE(nullptr, entry);
  EXPECT_NE(nullptr, io::ooxml::find_passthrough_part(wb, entry->media_path));
}

// A formula producing a picture is reported by the getter and left alone by clear.
TEST(WorkbookPictureTable, FormulaPictureIsReportedAndNotCleared) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, Png(0x11), "x")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.cell_picture(0, 0, 1).has_value());
  EXPECT_EQ(std::optional<std::string>("x"), wb.cell_picture(0, 0, 1)->alt_text);

  ASSERT_TRUE(static_cast<bool>(wb.clear_cell_picture(0, 0, 1)));
  const Cell* cell = wb.sheet(0).cell_at(0, 1);
  ASSERT_NE(nullptr, cell);
  EXPECT_EQ("=A1", cell->formula_text);
  EXPECT_TRUE(cell->cached_value.is_picture());
}

// i4: alts differing in case are two ids that compare equal.
TEST(WorkbookPictureTable, AltCaseFoldsThroughRecalc) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  const Bytes png = Png(0x11);
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, png, "Picture")));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 1, 0, png, "PICTURE")));
  EXPECT_NE(PictureAt(wb, 0, 0).picture_id, PictureAt(wb, 1, 0).picture_id);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=A1=A2")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value eq = wb.sheet(0).resolve_cell_value(0, 2);
  ASSERT_TRUE(eq.is_boolean());
  EXPECT_TRUE(eq.as_boolean());
}

// s1: SORT among pictures puts the greater bytes first.
TEST(WorkbookPictureTable, SortByMediaBytesThroughRecalc) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Place(wb, 0, 0, Png(0x11))));
  ASSERT_TRUE(static_cast<bool>(Place(wb, 1, 0, Png(0x22))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=SORT(A1:A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(PictureAt(wb, 1, 0).picture_id, PictureAt(wb, 0, 2).picture_id);
  EXPECT_EQ(PictureAt(wb, 0, 0).picture_id, PictureAt(wb, 1, 2).picture_id);
}

}  // namespace
}  // namespace formulon
