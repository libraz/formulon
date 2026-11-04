#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kChartRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
constexpr const char* kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
constexpr const char* kChartData =
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart "
    "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{rId}\"/>"
    "</a:graphicData>";
constexpr const char* kChartSpace =
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>";

struct Item {
  std::uint32_t id;
  std::string name;
  bool operator==(const Item& other) const { return id == other.id && name == other.name; }
};

// Inserts charts, each with a part of its own, and reads names and ids back.
class Book {
 public:
  explicit Book(Workbook wb) : wb_(std::move(wb)) {}

  void SetProfile(ExcelProfile profile) { wb_.set_excel_profile(profile); }

  std::uint32_t AddChart(std::size_t sheet, const std::string& name = std::string()) {
    ImageInsertOptions placement;
    placement.anchor_kind = AnchorKind::kTwoCell;
    placement.width_emu = 4572000;
    placement.height_emu = 2743200;
    placement.name = name;
    std::string path;
    do {
      path = "xl/charts/chart" + std::to_string(++charts_) + ".xml";
    } while (io::ooxml::find_passthrough_part(wb_, path) != nullptr);
    auto id = insert_graphic_frame(wb_, sheet, placement, kChartData, path, kChartRel);
    EXPECT_TRUE(static_cast<bool>(id));
    const std::string xml = kChartSpace;
    EXPECT_TRUE(
        static_cast<bool>(wb_.add_passthrough_part(PassthroughPart(path, kCtChart, Bytes(xml.begin(), xml.end())))));
    return id ? id.value() : 0U;
  }

  std::uint32_t AddPicture(std::size_t sheet) {
    const Bytes png = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R', 0,
                       0,    0,   4,   0,   0,    0,    4,    8,    6, 0, 0, 0,  0,   0,   0,   0};
    auto id = insert_image(wb_, sheet, png.data(), png.size(), ImageInsertOptions{});
    EXPECT_TRUE(static_cast<bool>(id));
    return id ? id.value() : 0U;
  }

  std::uint32_t Paste(std::size_t from_sheet, std::uint32_t id, std::size_t to_sheet) {
    auto snap = snapshot_drawing(wb_, from_sheet, id);
    EXPECT_TRUE(static_cast<bool>(snap));
    if (!snap) {
      return 0U;
    }
    auto pasted = restore_drawing(wb_, to_sheet, snap.value().data(), snap.value().size(), kDrawingRestoreNewId);
    EXPECT_TRUE(static_cast<bool>(pasted)) << (pasted ? "" : pasted.error().message);
    return pasted ? pasted.value() : 0U;
  }

  std::vector<Item> Items(std::size_t sheet) const {
    std::vector<Item> items;
    auto objects = list_drawing_objects(wb_, sheet);
    EXPECT_TRUE(static_cast<bool>(objects));
    if (objects) {
      for (const DrawingObject& obj : objects.value()) {
        items.push_back(Item{obj.object_id, obj.name});
      }
    }
    return items;
  }

  void Reload() {
    auto bytes = io::write_ooxml(wb_);
    ASSERT_TRUE(static_cast<bool>(bytes));
    auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
    ASSERT_TRUE(static_cast<bool>(read));
    wb_ = std::move(read.value().workbook);
  }

  Workbook& wb() { return wb_; }

 private:
  Workbook wb_;
  int charts_ = 0;
};

using Items = std::vector<Item>;

TEST(DrawingNames, JapaneseSessionNumbersLikeExcel) {
  Book book(Workbook::create_for_profile(mac_365_ja_jp_profile()));
  book.wb().add_sheet("S2");

  EXPECT_EQ(book.AddChart(0), 2U);
  EXPECT_EQ(book.AddChart(0), 3U);
  ASSERT_TRUE(static_cast<bool>(remove_drawing(book.wb(), 0, 2)));
  EXPECT_EQ(book.AddChart(0), 4U);
  EXPECT_EQ(book.Items(0), (Items{{3, "グラフ 2"}, {4, "グラフ 3"}}));

  EXPECT_EQ(book.Paste(0, 3, 0), 5U);
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(book.wb(), 0, 3, "Sales")));
  EXPECT_EQ(book.Paste(0, 3, 0), 6U);
  EXPECT_EQ(book.Items(0), (Items{{3, "Sales"}, {4, "グラフ 3"}, {5, "グラフ 4"}, {6, "Sales"}}));

  EXPECT_EQ(book.Paste(0, 4, 1), 2U);
  EXPECT_EQ(book.Paste(0, 5, 1), 3U);
  EXPECT_EQ(book.Items(1), (Items{{2, "グラフ 1"}, {3, "グラフ 2"}}));

  book.Reload();
  book.SetProfile(mac_365_ja_jp_profile());
  EXPECT_EQ(book.AddChart(0), 2U);
  EXPECT_EQ(book.AddChart(1), 4U);
  EXPECT_EQ(book.Items(0), (Items{{3, "Sales"}, {4, "グラフ 3"}, {5, "グラフ 4"}, {6, "Sales"}, {2, "グラフ 1"}}));
  EXPECT_EQ(book.Items(1), (Items{{2, "グラフ 1"}, {3, "グラフ 2"}, {4, "グラフ 3"}}));

  auto copy = book.wb().copy_sheet(0, 2, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy)) << copy.error().message;
  EXPECT_EQ(book.Items(copy.value()),
            (Items{{2, "Sales"}, {3, "グラフ 2"}, {4, "グラフ 3"}, {5, "Sales"}, {6, "グラフ 5"}}));
}

TEST(DrawingNames, EnglishNamesFollowTheSameRule) {
  Book book(Workbook::create_for_profile(win_365_en_us_profile()));
  EXPECT_EQ(book.AddChart(0), 2U);
  EXPECT_EQ(book.AddPicture(0), 3U);
  EXPECT_EQ(book.Items(0), (Items{{2, "Chart 1"}, {3, "Picture 2"}}));
}

TEST(DrawingNames, JapanesePictureName) {
  Book book(Workbook::create_for_profile(mac_365_ja_jp_profile()));
  EXPECT_EQ(book.AddPicture(0), 2U);
  EXPECT_EQ(book.Items(0), (Items{{2, "図 1"}}));
}

TEST(DrawingNames, PastedDefaultNamesAreRecognisedByPattern) {
  Book book(Workbook::create_for_profile(mac_365_ja_jp_profile()));
  const std::uint32_t english = book.AddChart(0, "Chart 9");
  const std::uint32_t local = book.AddChart(0, "グラフ 9");
  const std::uint32_t custom = book.AddChart(0, "Chart A");
  const std::uint32_t bare = book.AddChart(0, "Chart");
  const std::uint32_t other_kind = book.AddChart(0, "図 9");
  EXPECT_EQ(book.Paste(0, english, 0), 7U);
  EXPECT_EQ(book.Paste(0, local, 0), 8U);
  EXPECT_EQ(book.Paste(0, custom, 0), 9U);
  EXPECT_EQ(book.Paste(0, bare, 0), 10U);
  EXPECT_EQ(book.Paste(0, other_kind, 0), 11U);
  const Items items = book.Items(0);
  ASSERT_EQ(items.size(), 10U);
  EXPECT_EQ(items[5].name, "グラフ 6");
  EXPECT_EQ(items[6].name, "グラフ 7");
  EXPECT_EQ(items[7].name, "Chart A");
  EXPECT_EQ(items[8].name, "Chart");
  EXPECT_EQ(items[9].name, "図 9");
}

TEST(DrawingNames, EnglishProfileKeepsLocalWordsAsCustom) {
  Book book(Workbook::create_for_profile(win_365_en_us_profile()));
  const std::uint32_t local = book.AddChart(0, "グラフ 9");
  const std::uint32_t english = book.AddChart(0, "Chart 9");
  EXPECT_EQ(book.Paste(0, local, 0), 4U);
  EXPECT_EQ(book.Paste(0, english, 0), 5U);
  const Items items = book.Items(0);
  EXPECT_EQ(items[2].name, "グラフ 9");
  EXPECT_EQ(items[3].name, "Chart 4");
}

TEST(DrawingNames, SetDrawingNameRenamesAndRejectsBadInput) {
  Book book(Workbook::create());
  const std::uint32_t id = book.AddChart(0);
  ASSERT_TRUE(static_cast<bool>(set_drawing_name(book.wb(), 0, id, "Revenue")));
  EXPECT_EQ(book.Items(0), (Items{{id, "Revenue"}}));
  book.Reload();
  EXPECT_EQ(book.Items(0), (Items{{id, "Revenue"}}));

  auto empty = set_drawing_name(book.wb(), 0, id, "");
  ASSERT_FALSE(static_cast<bool>(empty));
  EXPECT_EQ(empty.error().code, FormulonErrorCode::kInvalidArgument);
  auto missing = set_drawing_name(book.wb(), 0, 99, "X");
  ASSERT_FALSE(static_cast<bool>(missing));
  EXPECT_EQ(missing.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(book.Items(0), (Items{{id, "Revenue"}}));
}

}  // namespace
}  // namespace formulon
