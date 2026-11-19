// Conditional-format and data-validation formulas that name another workbook, held to Excel's own saves of the
// fixtures in `excel/cf_dv_external_link/`: an address into a link Excel cannot reach is read as that link's `#REF!`
// site, the rule is written to the x14 extension, and the link survives the save.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr const char* kSheetPart = "xl/worksheets/sheet1.xml";
constexpr const char* kXlsbSheetPart = "xl/worksheets/sheet1.bin";
constexpr std::uint32_t kBeginCfRule14Id = 1048;

std::vector<std::uint8_t> FixtureBytes(const std::string& name, const std::string& extension) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_dv_external_link/" + name + extension;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(file);
  return out;
}

Workbook ReadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

Workbook LoadFixture(const std::string& name, const std::string& suffix = "") {
  return ReadXlsx(FixtureBytes(name + suffix, ".xlsx"));
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

std::string ZipPart(const std::vector<std::uint8_t>& package, const std::string& name) {
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, package.data(), package.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "not a zip package";
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&zip, name.c_str(), &size, 0);
  std::string out = data != nullptr ? std::string(static_cast<const char*>(data), size) : std::string();
  mz_free(data);
  mz_zip_reader_end(&zip);
  return out;
}

/// The inner text of every `<tag>` element in `part`, in order.
std::vector<std::string> ElementTexts(const std::string& part, const std::string& tag) {
  std::vector<std::string> out;
  const std::string open = "<" + tag + ">";
  const std::string close = "</" + tag + ">";
  for (std::size_t at = part.find(open); at != std::string::npos; at = part.find(open, at)) {
    const std::size_t end = part.find(close, at);
    if (end == std::string::npos) {
      break;
    }
    out.push_back(part.substr(at + open.size(), end - at - open.size()));
    at = end;
  }
  return out;
}

/// Every cached sheet name of the package's externalLink parts, one entry per part in part order.
std::vector<std::string> LinkSheetNames(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  for (int n = 1;; ++n) {
    const std::string part = ZipPart(package, "xl/externalLinks/externalLink" + std::to_string(n) + ".xml");
    if (part.empty()) {
      break;
    }
    std::string names;
    for (std::size_t at = part.find("<sheetName val=\""); at != std::string::npos;
         at = part.find("<sheetName val=\"", at)) {
      const std::size_t start = at + 16U;
      const std::size_t end = part.find('"', start);
      names += part.substr(start, end - start) + ";";
      at = end;
    }
    out.push_back(names);
  }
  return out;
}

/// `text` with the directory written before a quoted book cut away (`'/d/[B.xlsx]S'!` becomes `'[B.xlsx]S'!`), then the
/// quotes dropped when the remaining `[Book]Sheet` needs none. The directory Excel writes is unmeasured.
std::string WithoutDirectory(std::string text) {
  const std::size_t bracket = text.find('[');
  if (text.empty() || text[0] != '\'' || bracket == std::string::npos) {
    return text;
  }
  text.erase(1U, bracket - 1U);
  const std::size_t close = text.find("'!");
  if (close == std::string::npos) {
    return text;
  }
  const std::string inner = text.substr(1U, close - 1U);
  const bool plain = inner.find_first_not_of("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.[]") ==
                     std::string::npos;
  if (plain) {
    text.erase(close, 1U);
    text.erase(0U, 1U);
  }
  return text;
}

std::vector<std::string> WithoutDirectories(std::vector<std::string> texts) {
  for (std::string& text : texts) {
    text = WithoutDirectory(std::move(text));
  }
  return texts;
}

std::vector<std::string> LinkTargetNames(const Workbook& wb) {
  std::vector<std::string> out;
  for (const ExternalLinkRecord* rec : wb.external_links_by_index(/*skip_ole_dde=*/false)) {
    out.push_back(rec->target);
  }
  return out;
}

std::size_t LinkPartCount(const std::vector<std::uint8_t>& package) {
  return LinkSheetNames(package).size();
}

/// Every record of an xlsb part as (type, payload), in order.
std::vector<std::pair<std::uint32_t, std::string>> AllRecords(const std::string& part) {
  std::vector<std::pair<std::uint32_t, std::string>> out;
  std::size_t at = 0;
  const auto varint = [&](int max_bytes) {
    std::uint32_t value = 0;
    for (int i = 0; i < max_bytes && at < part.size(); ++i) {
      const auto byte = static_cast<std::uint8_t>(part[at++]);
      value |= static_cast<std::uint32_t>(byte & 0x7FU) << (7 * i);
      if ((byte & 0x80U) == 0U) {
        break;
      }
    }
    return value;
  };
  while (at < part.size()) {
    const std::uint32_t type = varint(2);
    const std::uint32_t size = varint(4);
    if (at + size > part.size()) {
      break;
    }
    out.emplace_back(type, part.substr(at, size));
    at += size;
  }
  return out;
}

/// Every payload of `type` in an xlsb part, in order.
std::vector<std::string> RecordsOfType(const std::string& part, std::uint32_t type) {
  std::vector<std::string> out;
  for (auto& record : AllRecords(part)) {
    if (record.first == type) {
      out.push_back(std::move(record.second));
    }
  }
  return out;
}

std::vector<std::string> CfRule14Records(const std::string& part) {
  return RecordsOfType(part, kBeginCfRule14Id);
}

std::uint32_t LoadU32(const std::string& bytes, std::size_t at) {
  std::uint32_t value = 0;
  for (std::size_t i = 4; i-- > 0;) {
    value = (value << 8U) | static_cast<std::uint8_t>(bytes[at + i]);
  }
  return value;
}

/// The formula Excel's rule record carries: the FRTHeader (4, a formula list), the list count and the first
/// FRTFormula, whole (a leading word, cce, cb, the rgce bytes and the rgcb bytes). Empty when the record has none.
std::string CfRule14Formula(const std::string& payload) {
  const auto u32 = [&payload](std::size_t at) {
    std::uint32_t value = 0;
    for (std::size_t i = 4; i-- > 0;) {
      value = (value << 8U) | static_cast<std::uint8_t>(payload[at + i]);
    }
    return value;
  };
  constexpr std::size_t kListStart = 8;
  if (payload.size() < kListStart + 12U || u32(0) != 4U || u32(4) != 1U) {
    return {};
  }
  const std::size_t size = 12U + u32(kListStart + 4U) + u32(kListStart + 8U);
  return kListStart + size <= payload.size() ? payload.substr(kListStart, size) : std::string();
}

/// The rgce bytes of an FRTFormula `CfRule14Formula` returned.
std::string FormulaRgce(const std::string& formula) {
  return formula.size() < 12U ? std::string() : formula.substr(12U, LoadU32(formula, 4));
}

bool HasLegacyRule(const std::string& sheet_xml) {
  return sheet_xml.find("<conditionalFormatting") != std::string::npos ||
         sheet_xml.find("<cfRule") != std::string::npos;
}

const cf::CFRule& FirstRule(const Workbook& wb) {
  static const cf::CFRule empty;
  const auto& formats = wb.sheet(0).conditional_formats();
  return formats.empty() || formats[0].rules.empty() ? empty : formats[0].rules[0];
}

TEST(CfDvExternalLink, ExpressionRuleOnUnknownSheetBecomesTheLinksRefError) {
  const Workbook wb = LoadFixture("expression_cf_unknown_sheet");
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  EXPECT_EQ(FirstRule(wb).formula1.value_or(""), "[Gone]Gone!#REF!>0");
  EXPECT_EQ(LinkTargetNames(wb), (std::vector<std::string>{"Nope", "Gone"}));

  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("expression_cf_unknown_sheet_excel_saved", ".xlsx");
  const std::string ours = ZipPart(saved, kSheetPart);
  const std::string theirs = ZipPart(excel, kSheetPart);
  EXPECT_EQ(ElementTexts(ours, "xm:f"), ElementTexts(theirs, "xm:f"));
  EXPECT_EQ(ElementTexts(ours, "xm:sqref"), ElementTexts(theirs, "xm:sqref"));
  EXPECT_FALSE(HasLegacyRule(ours));
  EXPECT_EQ(LinkSheetNames(saved), LinkSheetNames(excel));
}

TEST(CfDvExternalLink, ExpressionRuleXlsbCarriesTheRuleFormulaOfExcelsSave) {
  const Workbook wb = LoadFixture("expression_cf_unknown_sheet");
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::vector<std::uint8_t> excel = FixtureBytes("expression_cf_unknown_sheet_excel_saved", ".xlsb");
  const std::vector<std::string> ours = CfRule14Records(ZipPart(saved.value(), kXlsbSheetPart));
  const std::vector<std::string> theirs = CfRule14Records(ZipPart(excel, kXlsbSheetPart));
  ASSERT_EQ(ours.size(), 1U);
  ASSERT_EQ(theirs.size(), 1U);
  const std::string formula = CfRule14Formula(theirs[0]);
  ASSERT_FALSE(formula.empty());
  EXPECT_EQ(CfRule14Formula(ours[0]), formula);
}

TEST(CfDvExternalLink, ListValidationOnUnknownSheetIsWrittenToX14) {
  const Workbook wb = LoadFixture("list_dv_unknown_sheet");
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("list_dv_unknown_sheet_excel_saved", ".xlsx");
  const std::string ours = ZipPart(saved, kSheetPart);
  const std::string theirs = ZipPart(excel, kSheetPart);
  EXPECT_EQ(ElementTexts(ours, "xm:f"), ElementTexts(theirs, "xm:f"));
  EXPECT_EQ(ElementTexts(ours, "xm:sqref"), ElementTexts(theirs, "xm:sqref"));
  EXPECT_EQ(ElementTexts(ours, "xm:sqref"), (std::vector<std::string>{"G1"}));
  EXPECT_EQ(LinkSheetNames(saved), LinkSheetNames(excel));
}

TEST(CfDvExternalLink, ColourScaleThresholdOnUnknownSheetIsWrittenToX14) {
  const Workbook wb = LoadFixture("colour_scale_min_cfvo_unknown_sheet");
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("colour_scale_min_cfvo_unknown_sheet_excel_saved", ".xlsx");
  const std::string ours = ZipPart(saved, kSheetPart);
  const std::string theirs = ZipPart(excel, kSheetPart);
  EXPECT_EQ(ElementTexts(ours, "xm:f"), ElementTexts(theirs, "xm:f"));
  EXPECT_EQ(ours.find("<colorScale>"), std::string::npos);
  EXPECT_FALSE(HasLegacyRule(ours));

  // The threshold is a `PtgRefErr3d` through the XTI of the Gone link. Excel's own self XTI entry is unmeasured.
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  std::string rgce;
  for (const std::string& record : RecordsOfType(ZipPart(xlsb.value(), "xl/worksheets/sheet1.bin"), 1050U)) {
    if (const std::string formula = CfRule14Formula(record); !formula.empty()) {
      rgce = FormulaRgce(formula);
      break;
    }
  }
  ASSERT_GE(rgce.size(), 3U);
  EXPECT_EQ(static_cast<std::uint8_t>(rgce[0]) & 0x1FU, 0x1CU);
  EXPECT_NE(static_cast<std::uint8_t>(rgce[0]) & 0x60U, 0U) << "a bare PtgErr carries no class";
  // Reading it back resolves that XTI to the Gone link's site.
  auto reloaded = io::xlsb::read_xlsb(io::ByteSpan{xlsb.value().data(), xlsb.value().size()});
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  const cf::CFRule& rule = FirstRule(reloaded.value().workbook);
  ASSERT_TRUE(rule.color_scale.has_value());
  ASSERT_FALSE(rule.color_scale->thresholds.empty());
  EXPECT_EQ(rule.color_scale->thresholds[0].value, "[Gone]Gone!#REF!");
}

TEST(CfDvExternalLink, RealPathLinkRuleIsWrittenToX14AsTheRefErrorSite) {
  const Workbook wb = LoadFixture("x14_cf_real_path_link");
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("x14_cf_real_path_link_excel_saved", ".xlsx");
  // The directory written in the path form is unmeasured.
  EXPECT_EQ(WithoutDirectories(ElementTexts(ZipPart(saved, kSheetPart), "xm:f")),
            WithoutDirectories(ElementTexts(ZipPart(excel, kSheetPart), "xm:f")));
  EXPECT_EQ(LinkSheetNames(saved), LinkSheetNames(excel));
}

TEST(CfDvExternalLink, ReloadingExcelsSaveBindsNoExtraLink) {
  const std::vector<std::uint8_t> excel = FixtureBytes("x14_cf_real_path_link_excel_saved", ".xlsx");
  const Workbook wb = ReadXlsx(excel);
  EXPECT_EQ(wb.external_links().size(), LinkPartCount(excel));
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    EXPECT_FALSE(rec.part_path.empty());
  }
}

TEST(CfDvExternalLink, DataBarThresholdOnExternalBookIsWrittenToX14AndKeepsItsLink) {
  const Workbook wb = LoadFixture("x14_data_bar_external_threshold");
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("x14_data_bar_external_threshold_excel_saved", ".xlsx");
  EXPECT_EQ(WithoutDirectories(ElementTexts(ZipPart(saved, kSheetPart), "xm:f")),
            WithoutDirectories(ElementTexts(ZipPart(excel, kSheetPart), "xm:f")));
  EXPECT_EQ(LinkPartCount(saved), 1U);
  EXPECT_EQ(LinkSheetNames(saved), LinkSheetNames(excel));
}

// unmeasured: current behaviour
TEST(CfDvExternalLink, UnmeasuredExternalDefinedNameInRuleStaysAsIs) {
  Workbook wb = LoadFixture("x14_cf_real_path_link");
  const std::string kept =
      wb.ingest_stored_formula("[1]!Nm>0", Workbook::FormulaSource::kEntered, Workbook::UnknownQualifierBinding::kKeep);
  const std::string rule =
      wb.ingest_stored_formula("[1]!Nm>0", Workbook::FormulaSource::kEntered, Workbook::UnknownQualifierBinding::kRule);
  EXPECT_EQ(rule, kept);
  EXPECT_EQ(rule.find("#REF!"), std::string::npos);
}

// unmeasured: current behaviour
TEST(CfDvExternalLink, UnmeasuredPathFormNamingNoLinkStaysUnpointedWhenTwoLinksShareTheFileName) {
  Workbook wb = LoadFixture("x14_cf_real_path_link");
  std::vector<ExternalLinkRecord> links = wb.external_links();
  ASSERT_FALSE(links.empty());
  ExternalLinkRecord twin = links[0];
  twin.index = static_cast<std::uint32_t>(links.size() + 1U);
  twin.part_path.clear();
  twin.absolute_target = "/twin/dir/" + links[0].target.substr(links[0].target.find_last_of('/') + 1U);
  const std::string file = twin.absolute_target.substr(twin.absolute_target.find_last_of('/') + 1U);
  links.push_back(twin);
  wb.set_external_links(std::move(links));
  const std::string text = "'/nowhere/[" + file + "]A'!#REF!>0";
  const std::string rule =
      wb.ingest_stored_formula(text, Workbook::FormulaSource::kEntered, Workbook::UnknownQualifierBinding::kRule);
  EXPECT_NE(rule.find("/nowhere/[" + file + "]A'!#REF!"), std::string::npos) << rule;
}

}  // namespace
}  // namespace formulon
