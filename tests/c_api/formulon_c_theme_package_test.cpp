// C ABI tests for theme packages: fm_workbook_apply_theme / fm_workbook_save_theme.

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kZipCorrupt = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipCorrupt);
constexpr fm_status_t kXmlParse = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoXmlParse);
constexpr fm_status_t kRelationshipBroken =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoRelationshipBroken);

constexpr const char* kThemeEntry = "theme/theme/theme1.xml";
constexpr const char* kRootRels = "_rels/.rels";
constexpr const char* kManagerRels = "theme/theme/_rels/themeManager.xml.rels";
constexpr const char* kEmptyRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";

struct Fonts {
  std::string major_latin;
  std::string major_east_asian;
  std::string minor_latin;
  std::string minor_east_asian;
  bool operator==(const Fonts& other) const {
    return major_latin == other.major_latin && major_east_asian == other.major_east_asian &&
           minor_latin == other.minor_latin && minor_east_asian == other.minor_east_asian;
  }
  bool operator!=(const Fonts& other) const { return !(*this == other); }
};

Fonts ReadFonts(const fm_workbook_t* wb) {
  fm_theme_fonts_t fonts = FM_STRUCT_INIT(fm_theme_fonts_t);
  EXPECT_EQ(fm_workbook_get_theme_fonts(wb, &fonts), 0);
  return Fonts{fonts.major_latin, fonts.major_east_asian, fonts.minor_latin, fonts.minor_east_asian};
}

void SetFonts(fm_workbook_t* wb, const char* major, const char* minor) {
  fm_theme_fonts_t fonts = FM_STRUCT_INIT(fm_theme_fonts_t);
  fonts.major_latin = major;
  fonts.major_east_asian = "Major EA";
  fonts.minor_latin = minor;
  fonts.minor_east_asian = "Minor EA";
  ASSERT_EQ(fm_workbook_set_theme_fonts(wb, &fonts), 0);
}

Bytes SaveTheme(const fm_workbook_t* wb) {
  BufferGuard buffer;
  EXPECT_EQ(fm_workbook_save_theme(wb, &buffer.data, &buffer.len), 0);
  return Bytes(buffer.data, buffer.data + buffer.len);
}

/// Rewrites a package, replacing the entries named in `replacements`; each must already exist.
Bytes ReplaceEntries(const Bytes& package, const std::map<std::string, std::string>& replacements) {
  formulon::io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(formulon::io::ByteSpan{package.data(), package.size()})));
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  std::size_t replaced = 0;
  for (const std::string& name : input.list_entries()) {
    auto body = input.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(body)) << name;
    if (!body) {
      continue;
    }
    Bytes content = body.value();
    if (auto it = replacements.find(name); it != replacements.end()) {
      content.assign(it->second.begin(), it->second.end());
      ++replaced;
    }
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), content.data(), content.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE);
  }
  EXPECT_EQ(replaced, replacements.size());
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  Bytes result(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return result;
}

bool HasEntry(const Bytes& package, const char* name) {
  formulon::io::ZipReader input;
  return static_cast<bool>(input.open(formulon::io::ByteSpan{package.data(), package.size()})) && input.has_entry(name);
}

/// A workbook whose theme fonts were set, so its saved theme differs from the default.
void MakeSource(WorkbookGuard* out) {
  ASSERT_EQ(fm_workbook_create_default(&out->handle), 0);
  SetFonts(out->handle, "Major Source", "Minor Source");
}

TEST(CThemePackage, FreshWorkbookSavesAThemePackage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes package = SaveTheme(wb.handle);
  EXPECT_TRUE(HasEntry(package, kThemeEntry));
  EXPECT_TRUE(HasEntry(package, kRootRels));
}

TEST(CThemePackage, SaveThenApplyRoundTrips) {
  WorkbookGuard source;
  MakeSource(&source);
  const Fonts expected = ReadFonts(source.handle);
  ASSERT_EQ(expected.major_latin, "Major Source");
  ASSERT_EQ(expected.minor_latin, "Minor Source");

  const Bytes package = SaveTheme(source.handle);
  WorkbookGuard target;
  ASSERT_EQ(fm_workbook_create_default(&target.handle), 0);
  ASSERT_NE(ReadFonts(target.handle), expected);
  ASSERT_EQ(fm_workbook_apply_theme(target.handle, package.data(), package.size()), 0);
  EXPECT_EQ(ReadFonts(target.handle), expected);
}

TEST(CThemePackage, NullArgumentsAreRefused) {
  WorkbookGuard wb;
  MakeSource(&wb);
  const Fonts before = ReadFonts(wb.handle);
  const Bytes package = SaveTheme(wb.handle);
  EXPECT_EQ(fm_workbook_apply_theme(nullptr, package.data(), package.size()), kNullPointer);
  EXPECT_EQ(fm_workbook_apply_theme(wb.handle, nullptr, package.size()), kNullPointer);
  EXPECT_EQ(ReadFonts(wb.handle), before);

  BufferGuard buffer;
  EXPECT_EQ(fm_workbook_save_theme(nullptr, &buffer.data, &buffer.len), kNullPointer);
  EXPECT_EQ(fm_workbook_save_theme(wb.handle, nullptr, &buffer.len), kNullPointer);
  EXPECT_EQ(fm_workbook_save_theme(wb.handle, &buffer.data, nullptr), kNullPointer);
  EXPECT_EQ(buffer.data, nullptr);
  EXPECT_EQ(buffer.len, 0U);
}

TEST(CThemePackage, RefusedPackagesLeaveTheThemeUnchanged) {
  WorkbookGuard source;
  MakeSource(&source);
  const Bytes good = SaveTheme(source.handle);

  WorkbookGuard target;
  ASSERT_EQ(fm_workbook_create_default(&target.handle), 0);
  SetFonts(target.handle, "Major Target", "Minor Target");
  const Fonts before = ReadFonts(target.handle);

  const Bytes one_byte = {0x00};
  const Bytes junk = {'n', 'o', 't', ' ', 'a', ' ', 'z', 'i', 'p', ' ', 'a', 'r', 'c', 'h', 'i', 'v', 'e'};
  const Bytes no_office_document = ReplaceEntries(good, {{kRootRels, kEmptyRels}});
  const Bytes no_theme_rel = ReplaceEntries(good, {{kManagerRels, kEmptyRels}});
  const Bytes dangling_target = ReplaceEntries(
      good, {{kRootRels,
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
              "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
              "<Relationship Id=\"rId1\" "
              "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
              "Target=\"theme/theme/missing.xml\"/></Relationships>"}});
  const Bytes bad_theme = ReplaceEntries(good, {{kThemeEntry, "<a:theme"}});

  struct Case {
    const char* name;
    const Bytes* bytes;
    fm_status_t status;
  };
  const Case cases[] = {
      {"not a zip (one byte)", &one_byte, kZipCorrupt},
      {"not a zip (text)", &junk, kZipCorrupt},
      {"no officeDocument relationship", &no_office_document, kRelationshipBroken},
      {"no theme relationship", &no_theme_rel, kRelationshipBroken},
      {"missing manager part", &dangling_target, kRelationshipBroken},
      {"unparseable theme", &bad_theme, kXmlParse},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(fm_workbook_apply_theme(target.handle, c.bytes->data(), c.bytes->size()), c.status) << c.name;
    EXPECT_EQ(ReadFonts(target.handle), before) << c.name;
  }

  static const std::uint8_t kEmpty = 0;
  EXPECT_EQ(fm_workbook_apply_theme(target.handle, &kEmpty, 0U), kInvalidArgument);
  EXPECT_EQ(ReadFonts(target.handle), before);
}

}  // namespace
