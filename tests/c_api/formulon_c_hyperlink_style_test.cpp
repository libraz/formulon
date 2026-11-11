// Cell formats written with a hyperlink edit: setting a link gives its cells the
// built-in Hyperlink cell style, removing or clearing one resets the cells of
// its whole range to cell format 0, and the positional list edits touch no cell.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "column_width.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "style_region.h"
#include "styles.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;
constexpr std::uint32_t kLastCol = formulon::Sheet::kMaxCols - 1U;
constexpr std::int32_t kSourceCell = 0;
constexpr std::int32_t kSourceColumn = 2;
constexpr std::uint32_t kHyperlinkBuiltinId = 8U;
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

class HyperlinkStyle : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_.handle), 0); }

  fm_workbook_t* wb() { return wb_.handle; }
  formulon::StylesTable& styles() { return wb_.handle->workbook().mutable_styles(); }

  static fm_hyperlink_t Link(fm_cell_rect_t rect, const char* target, const char* display = nullptr) {
    fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
    hl.rect = rect;
    hl.target = target;
    hl.display = display;
    return hl;
  }

  fm_status_t Set(fm_cell_rect_t rect, const char* target = "https://example.com", const char* display = nullptr) {
    const fm_hyperlink_t hl = Link(rect, target, display);
    return fm_sheet_set_hyperlink(wb(), 0, &hl);
  }

  fm_status_t Remove(fm_cell_rect_t rect) { return fm_sheet_remove_hyperlink(wb(), 0, &rect); }

  // The cell's resolved xf and where it came from (cell, row or column).
  std::pair<std::int32_t, std::uint32_t> Effective(std::uint32_t row, std::uint32_t col) {
    fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);
    EXPECT_EQ(fm_sheet_get_effective_style(wb(), 0, row, col, &style), 0);
    return {style.source, style.xf_index};
  }

  std::uint32_t XfOf(std::uint32_t row, std::uint32_t col) { return Effective(row, col).second; }

  // The text a cell shows, or "<none>" when it holds no text.
  std::string TextAt(std::uint32_t row, std::uint32_t col) {
    fm_value_t v{};
    EXPECT_EQ(fm_sheet_get_value(wb(), 0, row, col, &v), 0);
    return v.kind == FM_VALUE_KIND_TEXT ? std::string(v.u.text) : std::string("<none>");
  }

  std::size_t LinkCount() {
    std::size_t count = 99U;
    EXPECT_EQ(fm_sheet_get_hyperlink_count(wb(), 0, &count), 0);
    return count;
  }

  std::size_t CellCount() {
    std::size_t count = 99U;
    EXPECT_EQ(fm_sheet_get_cell_count(wb(), 0, &count), 0);
    return count;
  }

  // The builtin-8 entries of the cell-style table.
  std::vector<std::size_t> HyperlinkStyles() {
    std::vector<std::size_t> found;
    const auto& table = styles().cell_styles;
    for (std::size_t i = 0; i < table.size(); ++i) {
      if (table[i].builtin_id == kHyperlinkBuiltinId) {
        found.push_back(i);
      }
    }
    return found;
  }

  // Sizes of every table a hyperlink edit could grow.
  std::vector<std::size_t> TableSizes() {
    const formulon::StylesTable& s = styles();
    return {s.fonts.size(),    s.fills.size(),          s.borders.size(),
            s.cell_xfs.size(), s.cell_style_xfs.size(), s.cell_styles.size()};
  }

  // Gives (row, col) the number `1` and the xf `xf`.
  void SeedCell(std::uint32_t row, std::uint32_t col, std::uint32_t xf) {
    ASSERT_EQ(fm_sheet_set_number(wb(), 0, row, col, 1.0), 0);
    ASSERT_EQ(fm_sheet_set_cell_xf_index(wb(), 0, row, col, xf), 0);
  }

  // An xf that differs from the Hyperlink style in every part a set must keep.
  std::uint32_t AddRichXf() {
    formulon::StylesTable& s = styles();
    formulon::CellXf xf;
    xf.num_fmt_id = 14U;
    xf.fill_index = static_cast<std::uint32_t>(s.fills.size());
    s.fills.emplace_back();
    xf.border_index = static_cast<std::uint32_t>(s.borders.size());
    s.borders.emplace_back();
    xf.has_alignment = true;
    xf.has_horizontal_align = true;
    xf.horizontal_align = 2U;
    xf.has_protection = true;
    xf.locked = false;
    xf.hidden = true;
    xf.quote_prefix = true;
    return formulon::intern_cell_xf(s, xf);
  }

 private:
  WorkbookGuard wb_;
};

TEST_F(HyperlinkStyle, SetAddsTheBuiltinHyperlinkStyleAndGivesTheCellItsFormat) {
  ASSERT_TRUE(HyperlinkStyles().empty());
  const std::size_t style_xfs_before = styles().cell_style_xfs.size();
  const std::size_t fonts_before = styles().fonts.size();
  ASSERT_EQ(Set({0U, 0U, 0U, 0U}), 0);

  const std::vector<std::size_t> found = HyperlinkStyles();
  ASSERT_EQ(found.size(), 1U);
  EXPECT_EQ(styles().cell_styles[found[0]].name, "Hyperlink");
  EXPECT_EQ(styles().cell_style_xfs.size(), style_xfs_before + 1U);
  EXPECT_EQ(styles().fonts.size(), fonts_before + 1U);
  const std::uint32_t style_index = styles().cell_styles[found[0]].xf_id;
  EXPECT_EQ(style_index, style_xfs_before);

  const formulon::CellXf& xf = styles().cell_xfs[XfOf(0, 0)];
  EXPECT_EQ(xf.xf_id, style_index);
  EXPECT_EQ(xf.font_index, styles().cell_style_xfs[style_index].font_index);
  EXPECT_FALSE(xf.apply_font);

  const formulon::FontRecord& normal = *formulon::resolve_normal_font(styles());
  const formulon::FontRecord& font = styles().fonts[xf.font_index];
  EXPECT_EQ(font.underline, 1U);
  EXPECT_EQ(font.color.kind, formulon::ColorSpec::Kind::kTheme);
  EXPECT_EQ(font.color.theme, 10U);
  EXPECT_EQ(font.name, normal.name);
  EXPECT_EQ(font.size, normal.size);
  EXPECT_EQ(font.bold, normal.bold);
  EXPECT_EQ(font.italic, normal.italic);
  EXPECT_EQ(font.has_family, normal.has_family);
  EXPECT_EQ(font.family, normal.family);
  EXPECT_EQ(font.scheme, normal.scheme);
}

TEST_F(HyperlinkStyle, TheStyleNameFollowsTheProfileLocale) {
  WorkbookGuard ja;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &ja.handle), 0);
  const fm_hyperlink_t hl = Link({0U, 0U, 0U, 0U}, "https://example.com");
  ASSERT_EQ(fm_sheet_set_hyperlink(ja.handle, 0, &hl), 0);
  const auto& table = ja.handle->workbook().styles().cell_styles;
  std::vector<std::string> names;
  for (const auto& record : table) {
    if (record.builtin_id == kHyperlinkBuiltinId) {
      names.push_back(record.name);
    }
  }
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0], "ハイパーリンク");
}

TEST_F(HyperlinkStyle, SetReusesAnExistingBuiltinHyperlinkStyleWhateverItsName) {
  formulon::StylesTable& s = styles();
  formulon::CellXf style_xf = s.cell_style_xfs.at(0);
  s.cell_style_xfs.push_back(style_xf);
  formulon::CellStyleRecord record;
  record.name = "MyLink";
  record.xf_id = static_cast<std::uint32_t>(s.cell_style_xfs.size() - 1U);
  record.builtin_id = kHyperlinkBuiltinId;
  s.cell_styles.push_back(record);
  const std::vector<std::size_t> sizes = TableSizes();

  ASSERT_EQ(Set({0U, 0U, 0U, 0U}), 0);
  ASSERT_EQ(Set({2U, 2U, 3U, 3U}), 0);

  ASSERT_EQ(HyperlinkStyles().size(), 1U);
  EXPECT_EQ(styles().cell_styles[HyperlinkStyles()[0]].name, "MyLink");
  EXPECT_EQ(styles().cell_style_xfs.size(), sizes[4]);
  EXPECT_EQ(styles().cell_styles.size(), sizes[5]);
  EXPECT_EQ(styles().cell_xfs[XfOf(0, 0)].xf_id, record.xf_id);
  EXPECT_EQ(styles().cell_xfs[XfOf(3, 3)].xf_id, record.xf_id);
}

TEST_F(HyperlinkStyle, SetKeepsEveryPartOfTheCellFormatButTheFont) {
  const std::uint32_t rich = AddRichXf();
  const formulon::CellXf before = styles().cell_xfs[rich];
  SeedCell(0, 0, rich);
  ASSERT_EQ(Set({0U, 0U, 0U, 0U}), 0);

  const std::uint32_t after_index = XfOf(0, 0);
  ASSERT_NE(after_index, rich);
  const formulon::CellXf& after = styles().cell_xfs[after_index];
  const std::uint32_t style_index = styles().cell_styles[HyperlinkStyles().at(0)].xf_id;
  EXPECT_EQ(after.xf_id, style_index);
  EXPECT_EQ(after.font_index, styles().cell_style_xfs[style_index].font_index);
  EXPECT_NE(after.font_index, before.font_index);
  EXPECT_FALSE(after.apply_font);

  EXPECT_EQ(after.num_fmt_id, 14U);
  EXPECT_TRUE(after.apply_number_format);
  EXPECT_EQ(after.fill_index, before.fill_index);
  EXPECT_TRUE(after.apply_fill);
  EXPECT_EQ(after.border_index, before.border_index);
  EXPECT_TRUE(after.apply_border);
  EXPECT_TRUE(after.has_alignment);
  EXPECT_EQ(after.horizontal_align, 2U);
  EXPECT_TRUE(after.has_protection);
  EXPECT_FALSE(after.locked);
  EXPECT_TRUE(after.hidden);
  EXPECT_TRUE(after.quote_prefix);
}

TEST_F(HyperlinkStyle, SetLeavesAppliedFlagsOffForPartsEqualToTheStyle) {
  SeedCell(0, 0, 0);
  ASSERT_EQ(Set({0U, 0U, 0U, 0U}), 0);
  const formulon::CellXf& xf = styles().cell_xfs[XfOf(0, 0)];
  EXPECT_FALSE(xf.apply_number_format);
  EXPECT_FALSE(xf.apply_fill);
  EXPECT_FALSE(xf.apply_border);
  EXPECT_FALSE(xf.apply_font);
}

TEST_F(HyperlinkStyle, SetFormatsBlankAndMergeCoveredCells) {
  const fm_cell_rect_t merge{1U, 1U, 2U, 2U};
  ASSERT_EQ(fm_sheet_add_merge(wb(), 0, &merge, nullptr), 0);
  ASSERT_EQ(Set({0U, 0U, 2U, 2U}), 0);

  const std::uint32_t linked = XfOf(0, 0);
  EXPECT_NE(linked, 0U);
  for (std::uint32_t row = 0; row <= 2U; ++row) {
    for (std::uint32_t col = 0; col <= 2U; ++col) {
      EXPECT_EQ(Effective(row, col), (std::pair<std::int32_t, std::uint32_t>{kSourceCell, linked}))
          << row << "," << col;
    }
  }
  // Outside the rectangle nothing changes.
  EXPECT_EQ(XfOf(0, 3), 0U);
  EXPECT_EQ(XfOf(3, 0), 0U);
  // A merge-covered cell gets the format but not the text.
  EXPECT_EQ(TextAt(2, 2), "<none>");
}

TEST_F(HyperlinkStyle, SetOverAWholeColumnFormatsTheColumn) {
  ASSERT_EQ(Set({0U, 0U, kLastRow, 0U}), 0);
  const auto [source, xf] = Effective(5000U, 0U);
  EXPECT_EQ(source, kSourceColumn);
  EXPECT_NE(xf, 0U);
  EXPECT_EQ(styles().cell_xfs[xf].xf_id, styles().cell_styles[HyperlinkStyles().at(0)].xf_id);
  EXPECT_EQ(XfOf(5000U, 1U), 0U);
  // Only the empty top-left cell was materialised for its text.
  EXPECT_LE(CellCount(), 1U);
}

TEST_F(HyperlinkStyle, SetOverAWholeRowFormatsTheRow) {
  ASSERT_EQ(Set({4U, 0U, 4U, kLastCol}), 0);
  const auto [source, xf] = Effective(4U, 9000U);
  EXPECT_EQ(source, 1);
  EXPECT_NE(xf, 0U);
  EXPECT_LE(CellCount(), 1U);
}

TEST_F(HyperlinkStyle, SetWritesTheTextOfAnEmptyTopLeftCell) {
  struct Row {
    const char* target;
    const char* location;
    const char* display;
    const char* expected;
  };
  const Row rows[] = {
      {"https://example.com", nullptr, "Shown", "Shown"},
      {"https://example.com", nullptr, nullptr, "https://example.com/"},
      {nullptr, "Sheet1!B2", nullptr, "Sheet1!B2"},
      {"https://example.com", "Sheet1!B2", nullptr, "https://example.com/"},
  };
  std::uint32_t row_index = 0;
  for (const Row& row : rows) {
    fm_hyperlink_t hl = Link({row_index, 0U, row_index, 1U}, row.target, row.display);
    hl.location = row.location;
    ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, &hl), 0);
    EXPECT_EQ(TextAt(row_index, 0), row.expected) << row_index;
    // Only the top-left cell receives text.
    EXPECT_EQ(TextAt(row_index, 1), "<none>") << row_index;
    ++row_index;
  }
}

TEST_F(HyperlinkStyle, SetKeepsTheValueOfANonEmptyTopLeftCell) {
  ASSERT_EQ(fm_sheet_set_text(wb(), 0, 0, 0, "kept"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb(), 0, 1, 0, 42.0), 0);
  ASSERT_EQ(Set({0U, 0U, 0U, 0U}, "https://example.com", "Shown"), 0);
  ASSERT_EQ(Set({1U, 0U, 1U, 0U}, "https://example.com", "Shown"), 0);
  EXPECT_EQ(TextAt(0, 0), "kept");
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 42.0);
  EXPECT_NE(XfOf(0, 0), 0U);
  EXPECT_NE(XfOf(1, 0), 0U);
}

TEST_F(HyperlinkStyle, SetStoresTheTargetTheWayExcelDoes) {
  struct Row {
    const char* in;
    const char* target;
    const char* location;
    const char* text;
  };
  const Row rows[] = {
      {"HTTPS://E.X/Path", "https://e.x/Path", "", "https://e.x/Path"},
      {"https://e.x?q=1", "https://e.x/?q=1", "", "https://e.x/?q=1"},
      {"https://e.x/b#frag", "https://e.x/b", "frag", "https://e.x/b - frag"},
      {"MAILTO:A@B.C", "mailto:A@B.C", "", "mailto:A@B.C"},
      {"/Users/me/f.xlsx", "file:///Users/me/f.xlsx", "", "file:///Users/me/f.xlsx"},
      {"\\\\server\\share\\f.xlsx", "file:////server/share/f.xlsx", "", "file:////server/share/f.xlsx"},
      {"docs/a.xlsx", "docs/a.xlsx", "", "docs/a.xlsx"},
  };
  std::uint32_t row_index = 0;
  for (const Row& row : rows) {
    const fm_hyperlink_t hl = Link({row_index, 0U, row_index, 0U}, row.in);
    ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, &hl), 0) << row.in;
    int32_t present = 0;
    fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
    ASSERT_EQ(fm_sheet_get_hyperlink(wb(), 0, row_index, 0, &present, &out), 0);
    ASSERT_EQ(present, 1);
    EXPECT_STREQ(out.target, row.target) << row.in;
    EXPECT_STREQ(out.location != nullptr ? out.location : "", row.location) << row.in;
    EXPECT_EQ(TextAt(row_index, 0), row.text) << row.in;
    ++row_index;
  }
}

TEST_F(HyperlinkStyle, PositionalInsertStoresTheTargetVerbatimAndWritesNoText) {
  const fm_hyperlink_t hl = Link({0U, 0U, 0U, 0U}, "HTTPS://E.X/b#frag");
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 0, &hl), 0);
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb(), 0, 0, &out), 0);
  EXPECT_STREQ(out.target, "HTTPS://E.X/b#frag");
  EXPECT_EQ(out.location, nullptr);
  EXPECT_EQ(out.display, nullptr);
  EXPECT_EQ(TextAt(0, 0), "<none>");
}

TEST_F(HyperlinkStyle, RemoveAtThenInsertAtOfALoadedUnnormalisedLinkChangesNothing) {
  // A positional insert keeps an address Excel would not store; a save and load carries it into the loaded book.
  const fm_hyperlink_t seed = Link({0U, 0U, 0U, 0U}, "https://E.X", "Shown");
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 0, &seed), 0);
  const auto save = [](fm_workbook_t* book, BufferGuard& bytes) {
    ASSERT_EQ(fm_workbook_save(book, FM_WORKBOOK_FORMAT_XLSX, &bytes.data, &bytes.len), 0);
  };
  BufferGuard first;
  save(wb(), first);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(first.data, first.len, &loaded.handle), 0);

  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(loaded.handle, 0, 0, &got), 0);
  const std::string target = got.target != nullptr ? got.target : "<null>";
  const std::string location = got.location != nullptr ? got.location : "<null>";
  const std::string display = got.display != nullptr ? got.display : "<null>";
  EXPECT_EQ(target, "https://E.X");

  fm_hyperlink_t again = FM_STRUCT_INIT(fm_hyperlink_t);
  again.rect = got.rect;
  again.target = target == "<null>" ? nullptr : target.c_str();
  again.location = location == "<null>" ? nullptr : location.c_str();
  again.display = display == "<null>" ? nullptr : display.c_str();
  ASSERT_EQ(fm_sheet_remove_hyperlink_at(loaded.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(loaded.handle, 0, 0, &again), 0);

  fm_hyperlink_t back = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(loaded.handle, 0, 0, &back), 0);
  EXPECT_STREQ(back.target, target.c_str());
  EXPECT_EQ(back.location != nullptr ? std::string(back.location) : std::string("<null>"), location);
  EXPECT_EQ(back.display != nullptr ? std::string(back.display) : std::string("<null>"), display);

  // The saved relationship Target is the loaded one: reading it again gives the same address.
  BufferGuard second;
  save(loaded.handle, second);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(second.data, second.len, &reloaded.handle), 0);
  fm_hyperlink_t final_link = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(reloaded.handle, 0, 0, &final_link), 0);
  EXPECT_STREQ(final_link.target, target.c_str());
}

TEST_F(HyperlinkStyle, SetKeepsACallerLocationOverAFragment) {
  fm_hyperlink_t hl = Link({0U, 0U, 0U, 0U}, "https://e.x/b#frag");
  hl.location = "Sheet1!A1";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, &hl), 0);
  int32_t present = 0;
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb(), 0, 0, 0, &present, &out), 0);
  EXPECT_STREQ(out.target, "https://e.x/b");
  EXPECT_STREQ(out.location, "Sheet1!A1");
}

TEST_F(HyperlinkStyle, SetOverATextCellRecordsTheTargetAsDisplayUnlessTheyMatch) {
  ASSERT_EQ(fm_sheet_set_text(wb(), 0, 0, 0, "already"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb(), 0, 1, 0, "https://e.x/d5"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb(), 0, 2, 0, "https://e.x/d6/"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb(), 0, 3, 0, "kept"), 0);
  ASSERT_EQ(Set({0U, 0U, 0U, 0U}, "https://e.x/d1"), 0);
  ASSERT_EQ(Set({1U, 0U, 1U, 0U}, "https://e.x/d5"), 0);
  ASSERT_EQ(Set({2U, 0U, 2U, 0U}, "https://e.x/d6"), 0);
  ASSERT_EQ(Set({3U, 0U, 3U, 0U}, "https://e.x/d7", "given"), 0);
  const char* const expected[] = {"https://e.x/d1", "", "https://e.x/d6", "given"};
  for (std::uint32_t r = 0; r < 4U; ++r) {
    int32_t present = 0;
    fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
    ASSERT_EQ(fm_sheet_get_hyperlink(wb(), 0, r, 0, &present, &out), 0);
    EXPECT_STREQ(out.display != nullptr ? out.display : "", expected[r]) << r;
  }
  EXPECT_EQ(TextAt(0, 0), "already");
}

TEST_F(HyperlinkStyle, SetOfALocationOnlyLinkGivesItTheLocationAsDisplay) {
  fm_hyperlink_t hl = Link({0U, 0U, 0U, 0U}, nullptr);
  hl.location = "Sheet1!A1";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, &hl), 0);
  int32_t present = 0;
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb(), 0, 0, 0, &present, &out), 0);
  EXPECT_STREQ(out.display, "Sheet1!A1");
}

TEST_F(HyperlinkStyle, SetRefusesOverlongTextBeforeChangingAnything) {
  const std::vector<std::size_t> sizes = TableSizes();
  const std::string long_text(32768U, 'a');
  EXPECT_EQ(Set({0U, 0U, 0U, 0U}, "https://example.com", long_text.c_str()), kInvalidArgument);
  EXPECT_EQ(LinkCount(), 0U);
  EXPECT_EQ(TableSizes(), sizes);
  EXPECT_EQ(TextAt(0, 0), "<none>");
  const std::string at_cap(32767U, 'a');
  EXPECT_EQ(Set({0U, 0U, 0U, 0U}, "https://example.com", at_cap.c_str()), 0);
}

TEST_F(HyperlinkStyle, SetOverTooManyCellsIsPreconditionFailedAndChangesNothing) {
  const std::vector<std::size_t> sizes = TableSizes();
  // 100,000 x 26 = 2.6 million cells, not a whole row or column.
  EXPECT_EQ(Set({0U, 0U, 99999U, 25U}, "https://example.com", "Shown"), kPreconditionFailed);
  EXPECT_EQ(LinkCount(), 0U);
  EXPECT_EQ(TableSizes(), sizes);
  EXPECT_TRUE(HyperlinkStyles().empty());
  EXPECT_EQ(TextAt(0, 0), "<none>");
  EXPECT_EQ(CellCount(), 0U);
  // One cell more than 2^20.
  EXPECT_EQ(Set({0U, 0U, 1024U, 1023U}), kPreconditionFailed);
  EXPECT_EQ(LinkCount(), 0U);
  EXPECT_EQ(TableSizes(), sizes);
}

TEST_F(HyperlinkStyle, RemoveResetsTheWholeRangeEvenCellsAnotherLinkStillCovers) {
  const std::uint32_t rich = AddRichXf();
  ASSERT_EQ(Set({0U, 0U, 2U, 2U}, "https://wide.example", "Wide"), 0);
  ASSERT_EQ(Set({1U, 1U, 1U, 1U}, "https://narrow.example", "Narrow"), 0);
  // Formatting added after the link is lost with it.
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb(), 0, 2, 2, rich), 0);
  ASSERT_EQ(LinkCount(), 2U);

  // A1 is inside A1:C3 only, so B2 does not intersect the removal.
  ASSERT_EQ(Remove({0U, 0U, 0U, 0U}), 0);
  ASSERT_EQ(LinkCount(), 1U);
  int32_t present = 0;
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb(), 0, 1, 1, &present, &got), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(got.target, "https://narrow.example/");

  // Every cell of A1:C3 is back to xf 0, B2 included although its link survives.
  for (std::uint32_t row = 0; row <= 2U; ++row) {
    for (std::uint32_t col = 0; col <= 2U; ++col) {
      EXPECT_EQ(Effective(row, col), (std::pair<std::int32_t, std::uint32_t>{kSourceCell, 0U})) << row << "," << col;
    }
  }
  // Values stay.
  EXPECT_EQ(TextAt(0, 0), "Wide");
  // B2 was empty when its own link was set, so that link wrote its text there.
  EXPECT_EQ(TextAt(1, 1), "Narrow");
}

TEST_F(HyperlinkStyle, RemoveByARangeResetsEveryRemovedLinkAndOnlyThose) {
  ASSERT_EQ(Set({0U, 0U, 0U, 1U}), 0);
  ASSERT_EQ(Set({5U, 5U, 6U, 6U}), 0);
  ASSERT_EQ(Set({9U, 9U, 9U, 9U}), 0);
  const std::uint32_t kept = XfOf(9, 9);
  ASSERT_NE(kept, 0U);

  // One range touching two links removes both, whole.
  ASSERT_EQ(Remove({0U, 1U, 5U, 5U}), 0);
  EXPECT_EQ(LinkCount(), 1U);
  EXPECT_EQ(XfOf(0, 0), 0U);
  EXPECT_EQ(XfOf(0, 1), 0U);
  EXPECT_EQ(XfOf(5, 5), 0U);
  EXPECT_EQ(XfOf(6, 6), 0U);
  EXPECT_EQ(XfOf(9, 9), kept);
}

TEST_F(HyperlinkStyle, RemoveOverAWholeColumnLinkResetsTheColumn) {
  ASSERT_EQ(Set({0U, 3U, kLastRow, 3U}), 0);
  ASSERT_EQ(Effective(100U, 3U).first, kSourceColumn);
  ASSERT_EQ(Remove({0U, 3U, 0U, 3U}), 0);
  EXPECT_EQ(Effective(100U, 3U).second, 0U);
  EXPECT_EQ(LinkCount(), 0U);
}

TEST_F(HyperlinkStyle, ClearResetsEveryLinkRange) {
  ASSERT_EQ(Set({0U, 0U, 1U, 1U}), 0);
  ASSERT_EQ(Set({4U, 4U, 4U, 6U}), 0);
  ASSERT_NE(XfOf(1, 1), 0U);
  ASSERT_NE(XfOf(4, 6), 0U);
  ASSERT_EQ(fm_sheet_clear_hyperlinks(wb(), 0), 0);
  EXPECT_EQ(LinkCount(), 0U);
  EXPECT_EQ(XfOf(0, 0), 0U);
  EXPECT_EQ(XfOf(1, 1), 0U);
  EXPECT_EQ(XfOf(4, 4), 0U);
  EXPECT_EQ(XfOf(4, 6), 0U);
  // Clearing an empty list stays a no-op.
  EXPECT_EQ(fm_sheet_clear_hyperlinks(wb(), 0), 0);
}

TEST_F(HyperlinkStyle, RemoveAndClearRefuseALinkOverTooManyCellsAndRemoveAtIsTheEscape) {
  // Only a file load or the positional insert can produce such a link.
  const fm_hyperlink_t big = Link({0U, 0U, 99999U, 25U}, "https://big.example");
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 0, &big), 0);
  ASSERT_EQ(Set({200000U, 0U, 200000U, 0U}, "https://small.example"), 0);
  ASSERT_EQ(LinkCount(), 2U);
  const std::vector<std::size_t> sizes = TableSizes();
  const std::uint32_t small_xf = XfOf(200000U, 0U);

  EXPECT_EQ(Remove({0U, 0U, 0U, 0U}), kPreconditionFailed);
  EXPECT_EQ(LinkCount(), 2U);
  EXPECT_EQ(fm_sheet_clear_hyperlinks(wb(), 0), kPreconditionFailed);
  EXPECT_EQ(LinkCount(), 2U);
  EXPECT_EQ(TableSizes(), sizes);
  EXPECT_EQ(XfOf(200000U, 0U), small_xf);

  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb(), 0, 0), 0);
  EXPECT_EQ(LinkCount(), 1U);
  EXPECT_EQ(fm_sheet_clear_hyperlinks(wb(), 0), 0);
  EXPECT_EQ(LinkCount(), 0U);
}

TEST_F(HyperlinkStyle, RemoveOfAWholeColumnLinkInsertedByPositionIsNotRefused) {
  const fm_hyperlink_t column = Link({0U, 2U, kLastRow, 2U}, "https://column.example");
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 0, &column), 0);
  EXPECT_EQ(Remove({10U, 2U, 10U, 2U}), 0);
  EXPECT_EQ(LinkCount(), 0U);
}

TEST_F(HyperlinkStyle, RemoveWithNoIntersectionChangesNoFormat) {
  ASSERT_EQ(Set({0U, 0U, 1U, 1U}), 0);
  const std::uint32_t linked = XfOf(0, 0);
  EXPECT_EQ(Remove({5U, 5U, 5U, 5U}), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(XfOf(0, 0), linked);
  EXPECT_EQ(LinkCount(), 1U);
}

TEST_F(HyperlinkStyle, PositionalInsertAndRemoveWriteNoFormatAndNoText) {
  ASSERT_EQ(Set({0U, 0U, 1U, 1U}, "https://one.example", "One"), 0);
  const formulon::StyleRect rect{0U, 0U, 3U, 3U};
  const std::vector<std::uint8_t> snapshot = formulon::snapshot_range_style(wb()->workbook().sheet(0), rect);
  const std::vector<std::size_t> sizes = TableSizes();
  const std::size_t cells = CellCount();
  const std::uint32_t linked = XfOf(0, 0);

  // A second link over the same cells and another over blank cells.
  const fm_hyperlink_t overlapping = Link({0U, 0U, 1U, 1U}, "https://two.example", "Two");
  const fm_hyperlink_t blank = Link({2U, 2U, 3U, 3U}, "https://three.example", "Three");
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 0, &overlapping), 0);
  ASSERT_EQ(fm_sheet_insert_hyperlink_at(wb(), 0, 2, &blank), 0);
  EXPECT_EQ(LinkCount(), 3U);
  EXPECT_EQ(TableSizes(), sizes);
  EXPECT_EQ(CellCount(), cells);
  EXPECT_EQ(TextAt(2, 2), "<none>");
  EXPECT_EQ(XfOf(2, 2), 0U);

  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb(), 0, 2), 0);
  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb(), 0, 0), 0);
  EXPECT_EQ(LinkCount(), 1U);
  EXPECT_EQ(TableSizes(), sizes);
  EXPECT_EQ(CellCount(), cells);
  EXPECT_EQ(XfOf(0, 0), linked);
  EXPECT_EQ(formulon::snapshot_range_style(wb()->workbook().sheet(0), rect), snapshot);
}

TEST_F(HyperlinkStyle, RemoveAtLeavesTheFormatOfItsCells) {
  ASSERT_EQ(Set({0U, 0U, 1U, 1U}), 0);
  const std::uint32_t linked = XfOf(1, 1);
  ASSERT_NE(linked, 0U);
  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb(), 0, 0), 0);
  EXPECT_EQ(LinkCount(), 0U);
  EXPECT_EQ(XfOf(1, 1), linked);
}

}  // namespace
