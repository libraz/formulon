// Name identity of every named workbook collection, observed through the
// Workbook API. A name being written collides with an existing one when the
// two compare equal under the mac-365-en_US text `=` collation: case, width,
// hiragana/katakana and composed/decomposed voicing fold, while a different
// base kana, a spacing voiced mark, punctuation and a trailing space do not.
// The rule is the same whatever the workbook's profile. Files are a separate
// boundary: a loaded workbook may hold names that only this fold equates.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

struct NamePair {
  const char* first;
  const char* second;
  bool same;
};

// Rename refusals measured in Excel, judged by the en-US `=` collation. The
// prolonged-mark rows (`ﾃﾞｰﾀ` / `データ`) are equal under that collation even
// though a ja-JP UI accepts them side by side.
constexpr NamePair kSheetPairs[] = {
    {"Sheet2", "Sheet2", true},
    {"Sheet2", "SHEET2", true},
    {"Sheet2", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92", true},  // Ｓｈｅｅｔ２
    {"Sheet2", "\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92", true},  // ｓｈｅｅｔ２
    {"Sheet2", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x99", false},  // Ｓｈｅｅｔ９
    {"ABC", "\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3", true},                                          // ＡＢＣ
    {"\xEF\xBC\xA1", "\xEF\xBD\x81", true},                                                         // Ａ / ａ
    {"\xC3\xA9", "\xC3\x89", true},                                                                 // é / É
    {"\xEF\xBD\xB6\xEF\xBE\x80", "\xE3\x82\xAB\xE3\x82\xBF", true},                                 // ｶﾀ / カタ
    {"\xEF\xBD\xB6\xEF\xBE\x80", "\xE3\x81\x8B\xE3\x81\x9F", true},                                 // ｶﾀ / かた
    {"\xE3\x82\xAB\xE3\x82\xBF", "\xE3\x81\x8B\xE3\x81\x9F", true},                                 // カタ / かた
    {"\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF", "\xE3\x81\xA7\xE3\x83\xBC\xE3\x81\x9F", true},  // データ / でーた
    {"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", true},                                      // ｶﾞ / ガ
    {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x99", true},                                      // ガ / カ + U+3099
    {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x9B", false},                                     // ガ / カ゛
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBE\x80", "\xE3\x83\x87\xE3\x82\xBF", true},              // ﾃﾞﾀ / デタ
    {"\xEF\xBD\xB0", "\xE3\x83\xBC", true},                                                  // ｰ / ー
    {"\xEF\xBD\xB0", "\xEF\xBC\x8D", false},                                                 // ｰ / －
    {"-", "\xEF\xBC\x8D", true},                                                             // - / －
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF",
     true},  // ﾃﾞｰﾀ / データ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xE3\x83\x87\xEF\xBD\xB0\xE3\x82\xBF", true},  // ﾃﾞｰﾀ / デｰタ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xEF\xBE\x83\xEF\xBE\x9E\xE3\x83\xBC\xEF\xBE\x80",
     true},  // ﾃﾞｰﾀ / ﾃﾞーﾀ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80 ",
     false},  // ﾃﾞｰﾀ / ﾃﾞｰﾀ + trailing space
    // Punctuation, spaces and apostrophes are significant; emoji compare by code point.
    {"ab", "a-b", false},
    {"ab", "a b", false},
    {"ab", "a'b", false},
    {"\xF0\x9F\x93\x8A", "\xF0\x9F\x93\x88", false},  // 📊 / 📈
    {"\xC3\x9F", "ss", true},                         // ß / ss
};

// Pairs measured on defined names, table names, table column headers and
// cell styles; every collection followed the same rule.
constexpr NamePair kCollectionPairs[] = {
    {"\xE3\x82\xAB", "\xEF\xBD\xB6", true},               // カ / ｶ
    {"\xE3\x82\xAB", "\xE3\x81\x8B", true},               // カ / か
    {"\xE3\x82\xAB", "\xE3\x82\xAC", false},              // カ / ガ
    {"\xE3\x82\xAB", "\xE3\x82\xAB\xE3\x82\x9B", false},  // カ / カ゛
    {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x99", true},   // ガ / カ + U+3099
    {"Ab", "aB", true},
    {"AB", "\xEF\xBC\xA1\xEF\xBC\xA2", true},  // ＡＢ
    {"\xE3\x82\xAB", "\xE3\x82\xAB", true},    // カ / カ
};

constexpr const char* kPivotName =
    "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC"
    "\xE3\x83\x96\xE3\x83\xAB"
    "1";  // ピボットテーブル1

constexpr NamePair kPivotPairs[] = {
    {kPivotName, kPivotName, true},
    {kPivotName,
     "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB\xEF\xBC\x91",
     true},  // ピボットテーブル１
    {kPivotName,
     "\xE3\x81\xB4\xE3\x81\xBC\xE3\x81\xA3\xE3\x81\xA8\xE3\x81\xA6\xE3\x83\xBC\xE3\x81\xB6\xE3\x82\x8B"
     "1",
     true},  // ぴぼっとてーぶる1
    {kPivotName,
     "\xEF\xBE\x8B\xEF\xBE\x9F\xEF\xBE\x8E\xEF\xBE\x9E\xEF\xBD\xAF\xEF\xBE\x84\xEF\xBE\x83\xEF\xBD\xB0\xEF\xBE\x8C"
     "\xEF\xBE\x9E\xEF\xBE\x99"
     "1",
     true},  // ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1
    {kPivotName,
     "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "9",
     false},  // ピボットテーブル9
};

// Identity does not depend on the profile, including ja-JP profiles whose
// formula `=` keeps the prolonged mark's width and Windows profiles whose `=`
// does not fold width at all.
constexpr const char* kProfiles[] = {"mac-365-en_US", "win-365-en_US", "mac-365-ja_JP", "win-365-ja_JP"};

constexpr const char* kFullWidthSheet1 = "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91";
constexpr const char* kHalfWidthData = "\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80";  // ﾃﾞｰﾀ
constexpr const char* kFullWidthData = "\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF";              // データ
constexpr const char* kHiraganaData = "\xE3\x81\xA7\xE3\x83\xBC\xE3\x81\x9F";               // でーた

Workbook ForProfile(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return Workbook::create_for_profile(profile);
}

FormulonErrorCode Code(const Expected<void, Error>& r) {
  return r ? FormulonErrorCode::kOk : r.error().code;
}

template <typename T>
FormulonErrorCode Code(const Expected<T, Error>& r) {
  return r.has_value() ? FormulonErrorCode::kOk : r.error().code;
}

FormulonErrorCode Expectation(bool same) {
  return same ? FormulonErrorCode::kAlreadyExists : FormulonErrorCode::kOk;
}

TableMetadata Table(std::string name, std::string ref, const std::vector<std::string>& columns) {
  TableMetadata table;
  table.sheet_index = 0;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::move(ref);
  for (const std::string& column_name : columns) {
    TableColumn column;
    column.name = column_name;
    table.columns.push_back(std::move(column));
  }
  return table;
}

void AddPivot(Workbook& wb, std::size_t sheet, std::string name) {
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_name(std::move(name));
  wb.sheet(sheet).add_pivot_table(std::move(pivot));
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// Writes a workbook holding `names` as its sheets (appended without the
/// entry-time identity check) and reads it back through the OOXML reader.
Expected<io::OoxmlReadResult, Error> LoadSheets(const std::vector<std::string>& names) {
  Workbook wb = Workbook::create_empty();
  for (const std::string& name : names) {
    wb.add_sheet(name);
  }
  for (std::uint32_t i = 0; i < names.size(); ++i) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(i, 0U, 0U, Value::number(1.0 + i))));
  }
  auto bytes = wb.save();
  EXPECT_TRUE(static_cast<bool>(bytes));
  if (!bytes) {
    return bytes.error();
  }
  return io::read_ooxml(SpanOf(bytes.value()));
}

// --- Sheets ---

TEST(NameIdentityCollation, AddedSheetCollidesUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kSheetPairs) {
      for (const bool swapped : {false, true}) {
        const char* existing = swapped ? p.second : p.first;
        const char* candidate = swapped ? p.first : p.second;
        SCOPED_TRACE(std::string(id) + ": " + existing + " / " + candidate);
        Workbook wb = ForProfile(id);
        ASSERT_EQ(Code(wb.rename_sheet(0, existing)), FormulonErrorCode::kOk);
        EXPECT_EQ(Code(wb.add_sheet_entered(candidate)), Expectation(p.same));
        EXPECT_EQ(wb.sheet_count(), p.same ? 1U : 2U);
      }
    }
  }
}

TEST(NameIdentityCollation, RenamedSheetCollidesUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kSheetPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.first + " / " + p.second);
      Workbook wb = ForProfile(id);
      ASSERT_EQ(Code(wb.rename_sheet(0, p.first)), FormulonErrorCode::kOk);
      ASSERT_EQ(Code(wb.add_sheet_entered("Other")), FormulonErrorCode::kOk);
      EXPECT_EQ(Code(wb.rename_sheet(1, p.second)), Expectation(p.same));
      EXPECT_EQ(wb.sheet(1).name(), p.same ? std::string("Other") : std::string(p.second));
    }
  }
}

TEST(NameIdentityCollation, CopiedSheetNameCollidesUnderCollationEquality) {
  for (const char* id : kProfiles) {
    SCOPED_TRACE(id);
    Workbook wb = ForProfile(id);
    ASSERT_EQ(Code(wb.rename_sheet(0, kHalfWidthData)), FormulonErrorCode::kOk);
    EXPECT_EQ(Code(wb.copy_sheet(0, 1, std::string(kHiraganaData))), FormulonErrorCode::kAlreadyExists);
    EXPECT_EQ(wb.sheet_count(), 1U);
  }
}

TEST(NameIdentityCollation, RenamingASheetToItsOwnCollationVariantKeepsTheNewSpelling) {
  for (const char* id : kProfiles) {
    SCOPED_TRACE(id);
    Workbook wb = ForProfile(id);
    ASSERT_EQ(Code(wb.add_sheet_entered("Other")), FormulonErrorCode::kOk);
    ASSERT_EQ(Code(wb.rename_sheet(0, kFullWidthSheet1)), FormulonErrorCode::kOk);
    EXPECT_EQ(wb.sheet(0).name(), kFullWidthSheet1);
  }
}

TEST(NameIdentityCollation, DefaultSheetNameSkipsNumbersTakenUnderCollationEquality) {
  for (const char* id : kProfiles) {
    SCOPED_TRACE(id);
    Workbook wb = ForProfile(id);
    ASSERT_EQ(wb.sheet(0).name(), "Sheet1");
    // Ｓｈｅｅｔ２ blocks Sheet2.
    ASSERT_EQ(Code(wb.add_sheet_entered("\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92")),
              FormulonErrorCode::kOk);
    EXPECT_EQ(wb.next_default_sheet_name(), "Sheet3");
    ASSERT_EQ(Code(wb.add_sheet_entered(wb.next_default_sheet_name())), FormulonErrorCode::kOk);
    EXPECT_EQ(wb.next_default_sheet_name(), "Sheet4");

    Workbook lower = ForProfile(id);
    ASSERT_EQ(Code(lower.add_sheet_entered("sheet2")), FormulonErrorCode::kOk);
    EXPECT_EQ(lower.next_default_sheet_name(), "Sheet3");

    Workbook other = ForProfile(id);
    ASSERT_EQ(Code(other.rename_sheet(0, "Data")), FormulonErrorCode::kOk);
    EXPECT_EQ(other.next_default_sheet_name(), "Sheet1");
  }
}

TEST(NameIdentityCollation, SingleSheetResolvesThroughCollationEquality) {
  for (const char* id : kProfiles) {
    SCOPED_TRACE(id);
    Workbook wb = ForProfile(id);
    // ＳＨＥＥＴ１ differs from Sheet1 in width as well as case.
    EXPECT_EQ(wb.sheet_index_by_name("\xEF\xBC\xB3\xEF\xBC\xA8\xEF\xBC\xA5\xEF\xBC\xA5\xEF\xBC\xB4\xEF\xBC\x91"), 0U);
    EXPECT_EQ(wb.sheet_index_by_name("sheet1"), 0U);
    EXPECT_EQ(wb.sheet_index_by_name("Sheet2"), Workbook::kNoSheet);
  }
}

// --- Loaded files ---

TEST(NameIdentityCollation, LoadedWidthVariantSheetsStaySeparateAndResolveExactly) {
  auto loaded = LoadSheets({"Sheet1", kFullWidthSheet1});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet_index_by_name("Sheet1"), 0U);
  EXPECT_EQ(wb.sheet_index_by_name(kFullWidthSheet1), 1U);
  // A case-only match wins over a collation-only one.
  EXPECT_EQ(wb.sheet_index_by_name("SHEET1"), 0U);
  EXPECT_EQ(wb.sheet_index_by_name("\xEF\xBC\xB3\xEF\xBC\xA8\xEF\xBC\xA5\xEF\xBC\xA5\xEF\xBC\xB4\xEF\xBC\x91"), 1U);
  // Ｓheet1 equals both only under collation, so it names neither.
  EXPECT_EQ(wb.sheet_index_by_name("\xEF\xBC\xB3heet1"), Workbook::kNoSheet);
  EXPECT_EQ(wb.sheet_by_name("\xEF\xBC\xB3heet1"), nullptr);

  // A third name that folds onto both is refused; one that only matches case-wise too.
  EXPECT_EQ(Code(wb.add_sheet_entered("\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94"
                                      "1")),
            FormulonErrorCode::kAlreadyExists);  // ｓｈｅｅｔ1
  EXPECT_EQ(Code(wb.add_sheet_entered("\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91")),
            FormulonErrorCode::kAlreadyExists);  // ｓｈｅｅｔ１
  EXPECT_EQ(wb.sheet_count(), 2U);
}

TEST(NameIdentityCollation, LoadedProlongedMarkVariantSheetsStaySeparateAndResolveExactly) {
  auto loaded = LoadSheets({kHalfWidthData, kFullWidthData});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet_index_by_name(kHalfWidthData), 0U);
  EXPECT_EQ(wb.sheet_index_by_name(kFullWidthData), 1U);
  EXPECT_EQ(wb.sheet_index_by_name(kHiraganaData), Workbook::kNoSheet);
  EXPECT_EQ(Code(wb.add_sheet_entered(kHiraganaData)), FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(Code(wb.rename_sheet(1, kHiraganaData)), FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(wb.sheet(1).name(), kFullWidthData);
}

// --- Defined names ---

TEST(NameIdentityCollation, DefinedNamesShareOneEntryUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kCollectionPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.first + " / " + p.second);
      Workbook wb = ForProfile(id);
      ASSERT_EQ(Code(wb.set_defined_name_scoped(p.first, "1", -1)), FormulonErrorCode::kOk);
      // A lookup through the other spelling reaches the entry only when the two are the same name.
      EXPECT_EQ(Code(wb.set_defined_name_hidden(p.second, -1, true)),
                p.same ? FormulonErrorCode::kOk : FormulonErrorCode::kInvalidArgument);
      const FormulonErrorCode second = Code(wb.set_defined_name_scoped(p.second, "2", -1));
      if (p.same) {
        EXPECT_EQ(wb.defined_names().size(), 1U);
      } else {
        EXPECT_EQ(second, FormulonErrorCode::kOk);
        EXPECT_EQ(wb.defined_names().size(), 2U);
      }
    }
  }
}

// --- Tables ---

TEST(NameIdentityCollation, TableNamesCollideUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kCollectionPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.first + " / " + p.second);
      Workbook wb = ForProfile(id);
      ASSERT_EQ(Code(wb.add_table(Table(p.first, "A1:B5", {"C1", "C2"}))), FormulonErrorCode::kOk);
      EXPECT_EQ(Code(wb.add_table(Table(p.second, "D1:E5", {"C1", "C2"}))), Expectation(p.same));

      Workbook renamed = ForProfile(id);
      ASSERT_EQ(Code(renamed.add_table(Table(p.first, "A1:B5", {"C1", "C2"}))), FormulonErrorCode::kOk);
      ASSERT_EQ(Code(renamed.add_table(Table("Other", "D1:E5", {"C1", "C2"}))), FormulonErrorCode::kOk);
      EXPECT_EQ(Code(renamed.rename_table(1, p.second, p.second)), Expectation(p.same));
    }
  }
}

TEST(NameIdentityCollation, DefinedNameAndTableNameCollideUnderCollationEquality) {
  for (const char* id : kProfiles) {
    SCOPED_TRACE(id);
    Workbook wb = ForProfile(id);
    ASSERT_EQ(Code(wb.add_table(Table("\xE3\x82\xAB", "A1:B5", {"C1", "C2"}))), FormulonErrorCode::kOk);      // カ
    EXPECT_EQ(Code(wb.set_defined_name_scoped("\xEF\xBD\xB6", "1", -1)), FormulonErrorCode::kAlreadyExists);  // ｶ
    EXPECT_TRUE(wb.defined_names().empty());
  }
}

TEST(NameIdentityCollation, ColumnHeaderClashGainsSuffixUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kCollectionPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.first + " / " + p.second);
      Workbook wb = ForProfile(id);
      ASSERT_EQ(Code(wb.add_table(Table("Tbl", "A1:B3", {p.first, "x"}))), FormulonErrorCode::kOk);
      Expected<std::string, Error> renamed = wb.rename_table_column(0, 1, p.second);
      ASSERT_TRUE(renamed.has_value());
      const std::string expected = p.same ? std::string(p.second) + "2" : std::string(p.second);
      EXPECT_EQ(renamed.value(), expected);
      EXPECT_EQ(wb.tables()[0].columns[1].name, expected);

      Workbook fresh = ForProfile(id);
      ASSERT_EQ(Code(fresh.add_table(Table("Tbl", "A1:B3", {p.first, p.second}))), FormulonErrorCode::kOk);
      EXPECT_EQ(fresh.tables()[0].columns[1].name, expected);
    }
  }
}

// --- Cell styles ---

TEST(NameIdentityCollation, CellStyleLookupMatchesUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kCollectionPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.first + " / " + p.second);
      Workbook wb = ForProfile(id);
      StylesTable& styles = wb.mutable_styles();
      const std::size_t before = styles.cell_styles.size();
      CellStyleRecord style;
      style.name = p.first;
      styles.cell_styles.push_back(style);
      EXPECT_EQ(Code(remove_cell_style(styles, p.second)),
                p.same ? FormulonErrorCode::kOk : FormulonErrorCode::kNotFound);
      EXPECT_EQ(styles.cell_styles.size(), p.same ? before : before + 1U);
    }
  }
}

// --- Pivot tables ---

TEST(NameIdentityCollation, PivotNamesCollideOnTheirSheetUnderCollationEquality) {
  for (const char* id : kProfiles) {
    for (const NamePair& p : kPivotPairs) {
      SCOPED_TRACE(std::string(id) + ": " + p.second);
      Workbook wb = ForProfile(id);
      ASSERT_EQ(Code(wb.add_sheet_entered("Other")), FormulonErrorCode::kOk);
      AddPivot(wb, 0, p.first);
      const Sheet& own = wb.sheet(0);
      EXPECT_EQ(pivot::pivot_name_taken(own, p.second, own.pivot_tables().size()), p.same);
      // The pivot's own position never collides with itself.
      EXPECT_FALSE(pivot::pivot_name_taken(own, p.second, 0));
      // Pivot names are per sheet.
      const Sheet& other = wb.sheet(1);
      EXPECT_FALSE(pivot::pivot_name_taken(other, p.second, other.pivot_tables().size()));
    }
  }
}

// --- Hash keys ---

TEST(NameIdentityCollation, MalformedNameKeyNeverEqualsAWellFormedKey) {
  const std::vector<std::string> well_formed = {"",
                                                "Sheet1",
                                                "sheet1",
                                                kFullWidthSheet1,
                                                kHalfWidthData,
                                                kFullWidthData,
                                                kHiraganaData,
                                                "\xC3\x9F",
                                                "ss",
                                                "\xC3\x83",
                                                "\xC3\x83\xC2\xA4",
                                                "A",
                                                "a",
                                                std::string(1, '\0'),
                                                "\x01",
                                                "\xEF\xBF\xBD"};
  const std::vector<std::string> malformed = {"\xC3",     "\xC3\xA4\xC3",     "\xFF\x41",
                                              "\xFF\x61", "\xED\xA0\x80",     "\xC0\x80",
                                              "\x80",     "\xF4\x90\x80\x80", "\xE2\x82"};
  for (const std::string& bad : malformed) {
    ASSERT_FALSE(sheet_names::valid_utf8(bad));
    for (const std::string& good : well_formed) {
      EXPECT_NE(collection_names::fold_key(bad), collection_names::fold_key(good))
          << testing::PrintToString(bad) << " vs " << testing::PrintToString(good);
      EXPECT_FALSE(collection_names::equal(bad, good));
    }
    EXPECT_EQ(collection_names::fold_key(bad), collection_names::fold_key(bad));
    EXPECT_TRUE(collection_names::equal(bad, bad));
  }
}

TEST(NameIdentityCollation, FoldKeyAgreesWithCollectionEquality) {
  std::vector<std::string> names = {"Sheet1",       "sheet1",      kFullWidthSheet1, kHalfWidthData,
                                    kFullWidthData, kHiraganaData, "\xC3\x9F",       "ss",
                                    "\xC3\xA9",     "\xC3\x89",    "\xFF\x41",       "\xFF\x61"};
  for (const NamePair& p : kCollectionPairs) {
    names.emplace_back(p.first);
    names.emplace_back(p.second);
  }
  for (const std::string& a : names) {
    for (const std::string& b : names) {
      EXPECT_EQ(collection_names::equal(a, b), collection_names::fold_key(a) == collection_names::fold_key(b))
          << testing::PrintToString(a) << " vs " << testing::PrintToString(b);
    }
  }
}

}  // namespace
}  // namespace formulon
