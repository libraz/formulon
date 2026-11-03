#include "sheet_copy_names.h"

#include <gtest/gtest.h>

#include <algorithm>
#include <cctype>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {
namespace {

// Names a copy of `source` against `existing` and records the result in it.
std::string Copy(std::vector<std::string>& existing, const std::string& source) {
  std::vector<std::string_view> views(existing.begin(), existing.end());
  std::string name = copy_sheet_default_name(source, views);
  existing.push_back(name);
  return name;
}

std::string Lower(std::string_view s) {
  std::string out(s);
  std::transform(out.begin(), out.end(), out.begin(),
                 [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
  return out;
}

std::string Table(std::string_view name, std::size_t tables, const std::vector<std::string>& taken = {}) {
  std::vector<std::string> names = taken;
  return copy_table_name(
      name, tables,
      [](void* context, std::string_view cand) {
        const auto& taken_names = *static_cast<const std::vector<std::string>*>(context);
        const std::string lc = Lower(cand);
        return std::any_of(taken_names.begin(), taken_names.end(),
                           [&](const std::string& t) { return Lower(t) == lc; });
      },
      &names);
}

TEST(SheetCopyNames, SheetMeasured) {
  std::vector<std::string> sheets{"Shx"};
  EXPECT_EQ(Copy(sheets, "Shx"), "Shx (2)");
  EXPECT_EQ(Copy(sheets, "Shx"), "Shx (3)");
  // Copying a copy strips its suffix: (2) and (3) are present.
  EXPECT_EQ(Copy(sheets, "Shx (2)"), "Shx (4)");

  std::vector<std::string> kx{"Kx", "Kx (2)"};
  EXPECT_EQ(Copy(kx, "Kx"), "Kx (3)");

  std::vector<std::string> cs{"Cs", "cs (2)"};
  EXPECT_EQ(Copy(cs, "Cs"), "Cs (3)");

  // 31/30/28/27-unit family sharing a 27-unit prefix.
  const std::string b31 = "abcdefghijklmnopqrstuvwxyzABCDE";
  const std::string b30 = b31.substr(0, 30);
  const std::string b28 = b31.substr(0, 28);
  const std::string b27 = b31.substr(0, 27);
  std::vector<std::string> fam{b31, b30, b28, b27};
  const std::string stem = "abcdefghijklmnopqrstuvwxyzA";
  EXPECT_EQ(Copy(fam, b31), stem + " (2)");
  EXPECT_EQ(Copy(fam, b31), stem + " (3)");
  EXPECT_EQ(Copy(fam, b30), stem + " (4)");
  EXPECT_EQ(Copy(fam, b30), stem + " (5)");
  EXPECT_EQ(Copy(fam, b28), stem + " (6)");
  EXPECT_EQ(Copy(fam, b28), stem + " (7)");
  EXPECT_EQ(Copy(fam, b27), stem + " (8)");
  EXPECT_EQ(Copy(fam, b27), stem + " (9)");
}

TEST(SheetCopyNames, SheetTenAndEdges) {
  const std::string stem = "abcdefghijklmnopqrstuvwxyz";
  const std::string b31 = stem + "ABCDE";

  // N >= 10 leaves 26 units for the base.
  std::vector<std::string> fam{b31};
  for (int n = 2; n <= 9; ++n) {
    fam.push_back(stem + "A (" + std::to_string(n) + ")");
  }
  EXPECT_EQ(Copy(fam, b31), stem + " (10)");

  std::vector<std::string> shx{"Shx"};
  for (int n = 2; n <= 9; ++n) {
    shx.push_back("Shx (" + std::to_string(n) + ")");
  }
  EXPECT_EQ(Copy(shx, "Shx"), "Shx (10)");

  // Japanese 31-unit base keeps 27 units.
  std::string ja;
  std::string ja27;
  for (int i = 0; i < 31; ++i) {
    ja += "\xE3\x81\x82";
    if (i < 27) {
      ja27 += "\xE3\x81\x82";
    }
  }
  std::vector<std::string> jas{ja};
  EXPECT_EQ(Copy(jas, ja), ja27 + " (2)");

  // The emoji (2 units) fits exactly at the 27-unit cut.
  const std::string emoji = "\xF0\x9F\x98\x80";
  std::vector<std::string> fits{stem.substr(0, 25) + emoji + "zzz"};
  EXPECT_EQ(Copy(fits, fits[0]), stem.substr(0, 25) + emoji + " (2)");

  // The emoji straddling the cut is dropped whole.
  std::vector<std::string> straddle{stem + emoji + "zz"};
  EXPECT_EQ(Copy(straddle, straddle[0]), stem + " (2)");

  // Trailing space and apostrophe after the cut are stripped.
  std::vector<std::string> sp{stem + " xyzxyz"};
  EXPECT_EQ(Copy(sp, sp[0]), stem + " (2)");
  std::vector<std::string> ap{stem + "'xyzxyz"};
  EXPECT_EQ(Copy(ap, ap[0]), stem + " (2)");

  // A base that strips to nothing keeps its truncated form.
  const std::string quotes(31, '\'');
  std::vector<std::string> q{quotes};
  EXPECT_EQ(Copy(q, quotes), std::string(27, '\'') + " (2)");

  // A suffix that would leave nothing is part of the name.
  std::vector<std::string> bare{" (2)"};
  EXPECT_EQ(Copy(bare, " (2)"), " (2) (2)");
}

TEST(SheetCopyNames, TableMeasured) {
  EXPECT_EQ(Table("Tbl", 1), "Tbl_3");
  EXPECT_EQ(Table("Tbl", 2), "Tbl_4");

  EXPECT_EQ(Table("My_Tbl", 4), "My_Tbl6");
  EXPECT_EQ(Table("Txn", 5), "Txn_7");
  EXPECT_EQ(Table("Q_12", 6), "Q_128");
  EXPECT_EQ(Table("Plain", 7), "Plain9");

  EXPECT_EQ(Table("My_Tbl", 8), "My_Tbl10");
  EXPECT_EQ(Table("Txn", 9), "Txn_11");
  EXPECT_EQ(Table("Q_12", 10), "Q_1212");
  EXPECT_EQ(Table("Plain", 11), "Plain13");

  EXPECT_EQ(Table("Tbl_3", 3), "Tbl_35");
}

TEST(SheetCopyNames, TableR1C1AndCollision) {
  EXPECT_EQ(Table("R", 4), "R_6");
  EXPECT_EQ(Table("C", 4), "C_6");
  EXPECT_EQ(Table("RC", 4), "RC_6");
  EXPECT_EQ(Table("r1c", 4), "r1c_6");  // R1C1 only, not an A1 cell
}

TEST(SheetCopyNames, TableA1Bounds) {
  EXPECT_EQ(Table("XFD", 0), "XFD_2");
  EXPECT_EQ(Table("XFE", 0), "XFE2");          // column beyond XFD
  EXPECT_EQ(Table("A10485", 75), "A1048577");  // row beyond 1048576
  EXPECT_EQ(Table("A10485", 74), "A10485_76");
  EXPECT_EQ(Table("ABCD", 0), "ABCD2");  // four letters
}

TEST(SheetCopyNames, TableCollisionAdvances) {
  EXPECT_EQ(Table("Tbl", 1, {"Tbl_3"}), "Tbl_4");
  EXPECT_EQ(Table("Plain", 7, {"plain9", "PLAIN10"}), "Plain11");
}

}  // namespace
}  // namespace formulon
