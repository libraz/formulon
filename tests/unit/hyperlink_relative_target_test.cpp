//
// The book-relative forms of local hyperlink targets: `book_relative_target`
// (what a save writes), `resolve_relative_target` (what the model holds) and
// the drive-path normalisation. The POSIX rows are Mac Excel measurements of
// books saved in the Excel container; the drive and UNC rows are estimates
// following the same rule.

#include <optional>
#include <string>

#include "gtest/gtest.h"
#include "hyperlink_address.h"
#include "sheet_metadata.h"

namespace formulon {
namespace {

constexpr const char* kData = "/Users/libraz/Library/Containers/com.microsoft.Excel/Data";

struct RelativeRow {
  std::string directory;
  std::string target;
  std::optional<std::string> relative;
};

TEST(HyperlinkRelativeTarget, WritesRelativeOnlyWhenAFolderSegmentIsShared) {
  const std::string probe3 = std::string(kData) + "/loc_probe3/p q";
  const std::string probe4 = std::string(kData) + "/loc_probe4/r";
  const std::string hlmodel = std::string(kData) + "/hlmodel";
  const RelativeRow rows[] = {
      {hlmodel, "file:///Users/libraz/Documents/f.xlsx", "../../../../../Documents/f.xlsx"},
      {probe3, "file:///Users/libraz/Documents/f.xlsx", "../../../../../../Documents/f.xlsx"},
      {probe3, "file:///Users/libraz/Documents/a b/f.xlsx", "../../../../../../Documents/a b/f.xlsx"},
      {probe3, "file:///Users/libraz/Documents/a%20b/f.xlsx", "../../../../../../Documents/a b/f.xlsx"},
      // Only the root is shared: absolute.
      {probe3, "file:///Volumes/music/x.xlsx", std::nullopt},
      {probe3, "file:///private/tmp/x.xlsx", std::nullopt},
      // `data` is not `Data`: the shared prefix ends at com.microsoft.Excel.
      {probe3, "file:///Users/libraz/Library/Containers/com.microsoft.Excel/data/loc_probe3/p q/other.xlsx",
       "../../../data/loc_probe3/p q/other.xlsx"},
      {probe3, "file://" + probe3 + "/sib.xlsx", "sib.xlsx"},
      {probe4, "file://" + probe3 + "/sib.xlsx", "../../loc_probe3/p q/sib.xlsx"},
      {"/a/b", "/a/c/f.xlsx", "../c/f.xlsx"},  // a plain path reads like its `file:` URL
      // A plain path's stray `%` is saved as `%25`; a `%XX` escape is kept.
      {"/a/b", "/a/b/x%y.xlsx", "x%25y.xlsx"},
      {"/a/b", "/a/b/x y.xlsx", "x y.xlsx"},
      {"/a/b", "/a/c/p%20q.xlsx", "../c/p%20q.xlsx"},
      {"/a/b", "/a/b/x%2", "x%252"},
      {"/", "file:///a/f.xlsx", std::nullopt},
      // UNC: another root kind than a POSIX book, so absolute.
      {probe3, "file:////server/share/f.xlsx", std::nullopt},
      // Estimated Windows forms (not measured): a UNC book folder.
      {"//server/share/a", "file:////SERVER/Share/a/f.xlsx", "f.xlsx"},
      {"//server/share/a", "file:////server/other/a/f.xlsx", std::nullopt},
      {"//server/share/a", "file:////server/share/b/f.xlsx", std::nullopt},
      // Estimated Windows forms (not measured). Drive: the letter compares case-insensitively, segments do not.
      {"C:/a/b", "file:///c:/a/x.xlsx", "../x.xlsx"},
      {"C:/a/b", "file:///C:/A/x.xlsx", std::nullopt},
      {"C:/a/b", "file:///D:/a/x.xlsx", std::nullopt},
      {"C:/a/b", "file:///C:/z/x.xlsx", std::nullopt},
      {"C:/a/b", "file:///a/b/x.xlsx", std::nullopt},
      // Not a local absolute target, or no folder.
      {probe3, "https://e.x/a", std::nullopt},
      {probe3, "mailto:a@b.c", std::nullopt},
      {probe3, "docs/a.xlsx", std::nullopt},
      {probe3, "", std::nullopt},
      {"", "file:///Users/libraz/Documents/f.xlsx", std::nullopt},
  };
  for (const RelativeRow& row : rows) {
    EXPECT_EQ(book_relative_target(row.target, row.directory), row.relative) << row.directory << " | " << row.target;
  }
}

struct ResolveRow {
  const char* target;
  const char* directory;
  std::optional<std::string> absolute;
};

TEST(HyperlinkRelativeTarget, ResolvesRelativePathsLexically) {
  const ResolveRow rows[] = {
      {"docs/a.xlsx", "/U/h/addr_depth", "file:///U/h/addr_depth/docs/a.xlsx"},
      {"../docs/a.xlsx", "/U/h/x", "file:///U/h/docs/a.xlsx"},
      {"./a.xlsx", "/U", "file:///U/a.xlsx"},
      {"a//b.xlsx", "/U", "file:///U/a/b.xlsx"},
      {"a.xlsx", "/", "file:///a.xlsx"},
      {"p%20q/f.xlsx", "/a", "file:///a/p q/f.xlsx"},
      {"sub\\f.xlsx", "/a", "file:///a/sub\\f.xlsx"},  // literal under a POSIX folder
      // Estimated Windows forms (not measured): drive and UNC folders.
      {"sub\\f.xlsx", "C:/a", "file:///C:/a/sub/f.xlsx"},
      {"f.xlsx", "C:/", "file:///C:/f.xlsx"},
      {"f.xlsx", "//server/share", "file:////server/share/f.xlsx"},
      {"../g/f.xlsx", "//server/share/a", "file:////server/share/g/f.xlsx"},
      // `..` above the root keeps the target as it is.
      {"../../x.xlsx", "/a", std::nullopt},
      {"../f.xlsx", "//server/share", std::nullopt},
      {"../f.xlsx", "C:/", std::nullopt},
      // Not relative local paths.
      {"https://e.x/", "/a", std::nullopt},
      {"mailto:a@b.c", "/a", std::nullopt},
      {"news:x", "/a", std::nullopt},
      {"onenote:x", "/a", std::nullopt},
      {"file:///a/b.xlsx", "/a", std::nullopt},
      {"/abs/b.xlsx", "/a", std::nullopt},
      {"\\\\server\\share\\b.xlsx", "/a", std::nullopt},
      {"C:\\x\\b.xlsx", "/a", std::nullopt},
      {"C:rel.xlsx", "/a", std::nullopt},
      {"#Sheet1!A1", "/a", std::nullopt},
      {"", "/a", std::nullopt},
      {"docs/a.xlsx", "", std::nullopt},
  };
  for (const ResolveRow& row : rows) {
    EXPECT_EQ(resolve_relative_target(row.target, row.directory), row.absolute) << row.directory << " | " << row.target;
  }
}

// Estimated Windows forms (not measured): drive paths.
TEST(HyperlinkRelativeTarget, NormalisesDrivePathsBeforeSchemes) {
  const char* const rows[][2] = {
      {"C:\\x\\f.xlsx", "file:///C:/x/f.xlsx"},
      {"C:/x/f.xlsx", "file:///C:/x/f.xlsx"},
      {"c:\\x", "file:///c:/x"},
      {"C:rel.xlsx", "C:rel.xlsx"},
      {"onenote:x", "onenote:x"},
  };
  for (const auto& row : rows) {
    EXPECT_EQ(normalize_hyperlink_target(row[0]).target, row[1]) << row[0];
  }
}

TEST(HyperlinkRelativeTarget, SavedFormDerivesDisplayOnlyWhenRelative) {
  Hyperlink link;
  link.target = "file:///a/b/c d/f.xlsx";
  SavedHyperlink saved = saved_hyperlink(link, "/a/x");
  EXPECT_EQ(saved.target, "../b/c d/f.xlsx");
  EXPECT_EQ(saved.display, "../b/c d/f.xlsx");

  link.display = "Mine";
  saved = saved_hyperlink(link, "/a/x");
  EXPECT_EQ(saved.target, "../b/c d/f.xlsx");
  EXPECT_EQ(saved.display, "Mine");

  link.display.clear();
  saved = saved_hyperlink(link, "");
  EXPECT_EQ(saved.target, "file:///a/b/c d/f.xlsx");
  EXPECT_EQ(saved.display, "");

  saved = saved_hyperlink(link, "/z");
  EXPECT_EQ(saved.target, "file:///a/b/c d/f.xlsx");
  EXPECT_EQ(saved.display, "");
}

// Local paths split `#` like URLs; a POSIX path holding `%` stays a path
// with only `%20` decoded (Mac Excel 365, book in loc_probe6/f1).
TEST(HyperlinkRelativeTarget, LocalPathsSplitFragmentsAndKeepPercentPaths) {
  struct Row {
    const char* input;
    const char* target;
    const char* fragment;
  };
  const Row rows[] = {
      {"/x/f1/a#b.xlsx", "file:///x/f1/a", "b.xlsx"},
      // Estimated Windows forms (not measured): UNC and drive paths.
      {"\\\\s\\h\\a#b.xlsx", "file:////s/h/a", "b.xlsx"},
      {"C:\\x\\a#b.xlsx", "file:///C:/x/a", "b.xlsx"},
      {"/x/f1/a%b.xlsx", "/x/f1/a%b.xlsx", ""},
      {"/x/f1/a%20b.xlsx", "/x/f1/a b.xlsx", ""},
      {"/x/f1/a%2Fb%20c.xlsx", "/x/f1/a%2Fb c.xlsx", ""},
      {"/x/f1/\xE6\x97\xA5\xE6\x9C\xAC.xlsx", "file:///x/f1/\xE6\x97\xA5\xE6\x9C\xAC.xlsx", ""},
  };
  for (const Row& row : rows) {
    const NormalizedHyperlinkTarget got = normalize_hyperlink_target(row.input);
    EXPECT_EQ(got.target, row.target) << row.input;
    EXPECT_EQ(got.fragment, row.fragment) << row.input;
  }
}

TEST(HyperlinkRelativeTarget, SavedDisplayAppendsTheLocation) {
  Hyperlink link;
  link.target = "file:///x/f1/a";
  link.location = "b.xlsx";
  const SavedHyperlink saved = saved_hyperlink(link, "/x/f1");
  EXPECT_EQ(saved.target, "a");
  EXPECT_EQ(saved.display, "a - b.xlsx");
}

}  // namespace
}  // namespace formulon
