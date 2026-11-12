//
// `Workbook::set_location`: the canonical directory forms, the refusals, the
// fixed point of `set_location(location())`, and the anchoring of relative
// local hyperlink targets when a folder first appears.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>

#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

WorkbookLocation At(const char* directory, const char* file_name) {
  return WorkbookLocation{std::string(directory), std::string(file_name)};
}

struct CanonicalRow {
  const char* input;
  const char* canonical;
};

TEST(WorkbookLocation, StoresEveryAcceptedDirectoryInCanonicalForm) {
  const CanonicalRow rows[] = {
      {"/a/b", "/a/b"},
      {"/a/b/", "/a/b"},
      {"/", "/"},
      {"/a\\b", "/a\\b"},  // a backslash is literal in a POSIX directory
      {"/U/x/a b \xC3\x9C", "/U/x/a b \xC3\x9C"},
      {"C:\\a", "C:/a"},
      {"C:/a", "C:/a"},
      {"c:\\a\\", "C:/a"},
      {"C:\\a\\b", "C:/a/b"},
      {"C:/a\\b", "C:/a/b"},
      {"C:\\", "C:/"},
      {"c:/", "C:/"},
      {"\\\\server\\share\\a", "//server/share/a"},
      {"//server/share/a", "//server/share/a"},
      {"\\\\server\\share", "//server/share"},
      {"\\\\server\\share\\", "//server/share"},
      {"//server/share/", "//server/share"},
  };
  for (const CanonicalRow& row : rows) {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_location(At(row.input, "b.xlsx")))) << row.input;
    ASSERT_TRUE(wb.location().has_value()) << row.input;
    ASSERT_TRUE(wb.location()->directory.has_value()) << row.input;
    EXPECT_EQ(*wb.location()->directory, row.canonical) << row.input;
    EXPECT_EQ(wb.location()->file_name, "b.xlsx") << row.input;

    // The getter's form is a fixed point of the setter.
    const WorkbookLocation stored = *wb.location();
    ASSERT_TRUE(static_cast<bool>(wb.set_location(stored))) << row.input;
    EXPECT_EQ(*wb.location(), stored) << row.input;
  }
}

TEST(WorkbookLocation, RefusesEveryMalformedDirectoryAndChangesNothing) {
  const std::string rows[] = {
      "",         "C:",        "C:rel",      "c:a\\b",       "\\\\server", "\\\\server\\",
      "//server", "//server/", "//a",        "relative/dir", "a",          "\\a",
      "/a/../b",  "/a/./b",    "/..",        "/.",           "C:\\a\\..",  "\\\\server\\share\\.",
      "/a//b",    "/a//",      "C:\\a\\\\b", "/a[b",         "/a]b",       std::string("/a\0b", 4),
      "/a\xFF",
  };
  for (const std::string& dir : rows) {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_location(At("/keep", "Keep.xlsx"))));
    const auto result = wb.set_location(WorkbookLocation{dir, "b.xlsx"});
    ASSERT_FALSE(static_cast<bool>(result)) << dir;
    EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument) << dir;
    EXPECT_EQ(*wb.location(), At("/keep", "Keep.xlsx")) << dir;
  }
}

TEST(WorkbookLocation, RefusesEveryMalformedFileName) {
  const std::string rows[] = {
      "",
      "Book",
      "Book.txt",
      "a/b.xlsx",
      "a\\b.xlsx",
      "a[1].xlsx",
      "a]1.xlsx",
      "a:b.xlsx",
      "a?.xlsx",
      "a*.xlsx",
      std::string("a\0.xlsx", 7),
      "\xFF.xlsx",
      std::string(214, 'n') + ".xlsx",
  };
  for (const std::string& name : rows) {
    Workbook wb = Workbook::create();
    for (const bool with_directory : {false, true}) {
      WorkbookLocation location{with_directory ? std::optional<std::string>("/a") : std::nullopt, name};
      const auto result = wb.set_location(location);
      ASSERT_FALSE(static_cast<bool>(result)) << name;
      EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument) << name;
      EXPECT_FALSE(wb.location().has_value()) << name;
    }
  }
  // 218 bytes is the longest name accepted.
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, std::string(213, 'n') + ".xlsx"})));
}

TEST(WorkbookLocation, BoundsTheCanonicalPathAt1024Bytes) {
  // `/` + 1016 bytes + `/` + `b.xlsx` is 1,024 bytes.
  Workbook wb = Workbook::create();
  const std::string fits = "/" + std::string(1016, 'd');
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{fits, "b.xlsx"})));
  EXPECT_EQ(*wb.location()->directory, fits);
  const auto over = wb.set_location(WorkbookLocation{"/" + std::string(1017, 'd'), "b.xlsx"});
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kSecurityResourceLimit);
  EXPECT_EQ(over.error().context, "part=location_path requested=1025 limit=1024");
  // The bound applies after canonicalisation: a stripped trailing separator
  // does not count.
  EXPECT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{fits + "/", "b.xlsx"})));
}

TEST(WorkbookLocation, UnsetAndNameOnlyForms) {
  Workbook wb = Workbook::create();
  EXPECT_FALSE(wb.location().has_value());
  EXPECT_EQ(wb.file_name(), "");

  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "b.xlsx"})));
  ASSERT_TRUE(wb.location().has_value());
  EXPECT_FALSE(wb.location()->directory.has_value());
  EXPECT_EQ(wb.file_name(), "b.xlsx");

  ASSERT_TRUE(static_cast<bool>(wb.set_location(At("/d", "c.xlsm"))));
  EXPECT_EQ(wb.file_name(), "c.xlsm");

  ASSERT_TRUE(static_cast<bool>(wb.set_location(std::nullopt)));
  EXPECT_FALSE(wb.location().has_value());
  EXPECT_EQ(wb.file_name(), "");
}

Hyperlink LinkAt(std::uint32_t row, const char* target) {
  Hyperlink link;
  link.row = row;
  link.last_row = row;
  link.target = target;
  return link;
}

const std::string& TargetAt(const Workbook& wb, std::size_t index) {
  return wb.sheet(0).hyperlinks()[index].target;
}

TEST(WorkbookLocation, AnchorsRelativeTargetsWhenAFolderFirstAppears) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(0, "docs/a.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(1, "../../../../x.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(2, "p%20q/f.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(3, "https://e.x/"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(4, "mailto:a@b.c"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(5, "news:x"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(6, "/Users/libraz/Documents/f.xlsx"))));
  EXPECT_EQ(TargetAt(wb, 0), "docs/a.xlsx");
  EXPECT_EQ(TargetAt(wb, 6), "file:///Users/libraz/Documents/f.xlsx");

  // A name without a folder anchors nothing.
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "shallow.xlsx"})));
  EXPECT_EQ(TargetAt(wb, 0), "docs/a.xlsx");

  ASSERT_TRUE(static_cast<bool>(wb.set_location(At("/U/h/addr_depth", "shallow.xlsx"))));
  EXPECT_EQ(TargetAt(wb, 0), "file:///U/h/addr_depth/docs/a.xlsx");
  EXPECT_EQ(TargetAt(wb, 1), "../../../../x.xlsx");  // `..` above the root stays verbatim
  EXPECT_EQ(TargetAt(wb, 2), "file:///U/h/addr_depth/p q/f.xlsx");
  EXPECT_EQ(TargetAt(wb, 3), "https://e.x/");
  EXPECT_EQ(TargetAt(wb, 4), "mailto:a@b.c");
  EXPECT_EQ(TargetAt(wb, 5), "news:x");
  EXPECT_EQ(TargetAt(wb, 6), "file:///Users/libraz/Documents/f.xlsx");

  // Later moves and an unset change no target.
  ASSERT_TRUE(static_cast<bool>(wb.set_location(At("/U/h/addr_depth/x/y/z", "deep.xlsx"))));
  EXPECT_EQ(TargetAt(wb, 0), "file:///U/h/addr_depth/docs/a.xlsx");
  EXPECT_EQ(TargetAt(wb, 1), "../../../../x.xlsx");
  ASSERT_TRUE(static_cast<bool>(wb.set_location(std::nullopt)));
  EXPECT_EQ(TargetAt(wb, 0), "file:///U/h/addr_depth/docs/a.xlsx");
  EXPECT_EQ(TargetAt(wb, 1), "../../../../x.xlsx");
}

TEST(WorkbookLocation, SetHyperlinkResolvesAgainstTheFolder) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(At("C:\\d\\e", "b.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(0, "docs\\a.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(1, "../f.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(2, "C:\\x\\f.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, LinkAt(3, "../../../f.xlsx"))));
  EXPECT_EQ(TargetAt(wb, 0), "file:///C:/d/e/docs/a.xlsx");
  EXPECT_EQ(TargetAt(wb, 1), "file:///C:/d/f.xlsx");
  EXPECT_EQ(TargetAt(wb, 2), "file:///C:/x/f.xlsx");
  EXPECT_EQ(TargetAt(wb, 3), "../../../f.xlsx");
  // The empty top-left cell takes the stored, absolute form.
  const Value a1 = wb.sheet(0).resolve_cell_value(0, 0);
  ASSERT_TRUE(a1.is_text());
  EXPECT_EQ(std::string(a1.as_text()), "file:///C:/d/e/docs/a.xlsx");
}

}  // namespace
}  // namespace formulon
