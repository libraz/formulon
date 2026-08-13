// End-to-end coverage for the cumulative ZIP extraction budget. The package
// bodies come from the real OOXML/XLSB writers; only the additional opaque
// parts are appended with miniz. This keeps the test at the reader boundary
// while ensuring all required workbook parts are valid and parseable.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

// Three valid parts are enough to cross the 256 MiB per-open-session budget,
// while each remains comfortably below the 100 MiB per-entry cap. The single
// buffer is borrowed three times when miniz builds the package; the test never
// keeps independent payload copies around. During a reader failure only the
// first two extracted vectors (~180 MiB) can be retained; the third is refused
// before allocation, and each standalone check releases its one output first.
constexpr std::size_t kBudgetEntryBytes = 90ULL * 1024ULL * 1024ULL;
constexpr std::size_t kBudgetEntryCount = 3U;
constexpr std::size_t kPatternPeriod = 8ULL * 1024ULL;

constexpr std::array<std::string_view, kBudgetEntryCount> kBudgetEntryNames = {
    "xl/formulon-budget-0.budget",
    "xl/formulon-budget-1.budget",
    "xl/formulon-budget-2.budget",
};

io::ByteSpan SpanOf(const Bytes& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

Bytes BuildPatternedPayload() {
  std::array<std::uint8_t, kPatternPeriod> period{};
  std::uint32_t state = 0x9E3779B9U;
  for (std::uint8_t& byte : period) {
    // A deterministic high-entropy period avoids a pathological compression
    // ratio while repeating within deflate's window keeps the archive small.
    state = state * 1664525U + 1013904223U;
    byte = static_cast<std::uint8_t>(state >> 24U);
  }

  Bytes payload(kBudgetEntryBytes);
  for (std::size_t i = 0; i < payload.size(); ++i) {
    payload[i] = period[i % period.size()];
  }
  return payload;
}

Bytes AddBudgetDefault(const Bytes& content_types) {
  std::string xml(reinterpret_cast<const char*>(content_types.data()), content_types.size());
  const std::size_t close = xml.rfind("</Types>");
  EXPECT_NE(close, std::string::npos);
  if (close == std::string::npos) {
    return {};
  }
  xml.insert(close, "  <Default Extension=\"budget\" ContentType=\"application/vnd.formulon.budget\"/>\n");
  return Bytes(xml.begin(), xml.end());
}

Bytes BasePackage(bool xlsb) {
  Workbook workbook = Workbook::create();
  if (xlsb) {
    auto written = io::xlsb::write_xlsb(workbook);
    EXPECT_TRUE(static_cast<bool>(written)) << "write_xlsb failed: " << written.error().message;
    return written ? written.value() : Bytes{};
  }

  auto written = io::write_ooxml(workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << "write_ooxml failed: " << written.error().message;
  return written ? written.value() : Bytes{};
}

// Rebuilds a writer-produced package while borrowing one payload buffer for
// all opaque parts. The content-type Default registration makes the entries
// genuine passthrough parts for both readers (XLSB needs the registry to
// select them; OOXML also walks the residual archive entries).
Bytes BuildBudgetPackage(bool xlsb) {
  const Bytes base = BasePackage(xlsb);
  const Bytes payload = BuildPatternedPayload();

  io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(SpanOf(base)))) << "base package is not a readable ZIP";
  if (input.entry_count() == 0U) {
    return {};
  }

  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const std::string& name : input.list_entries()) {
    auto body = input.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(body)) << "base entry read failed: " << name;
    if (!body) {
      continue;
    }
    Bytes replacement;
    const Bytes* source = &body.value();
    if (name == "[Content_Types].xml") {
      replacement = AddBudgetDefault(*source);
      source = &replacement;
    }
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), source->data(), source->size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "base entry write failed: " << name;
  }
  for (const std::string_view name : kBudgetEntryNames) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, std::string(name).c_str(), payload.data(), payload.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "budget entry write failed: " << name;
  }

  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  if (archive == nullptr) {
    return {};
  }
  Bytes result(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return result;
}

// A fresh open session must accept every individual entry. This proves the
// eventual reader failure is neither the per-entry cap nor the compression
// ratio guard; it is the cumulative session budget.
void ExpectBudgetEntriesFitIndividually(const Bytes& package) {
  for (const std::string_view name : kBudgetEntryNames) {
    io::ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(package)))) << "package open failed";
    auto body = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(body)) << "standalone read failed for " << name << ": " << body.error().message;
    EXPECT_EQ(body.value().size(), kBudgetEntryBytes);
  }
}

TEST(ZipLoadBudget, OoxmlReaderRejectsCumulativePayload) {
  const Bytes package = BuildBudgetPackage(/*xlsb=*/false);
  ASSERT_FALSE(package.empty());
  ExpectBudgetEntriesFitIndividually(package);

  auto loaded = io::read_ooxml(SpanOf(package));
  ASSERT_FALSE(static_cast<bool>(loaded));
  EXPECT_EQ(loaded.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_NE(loaded.error().context.find("limit=total"), std::string::npos);
  EXPECT_NE(loaded.error().message.find("cumulative"), std::string::npos);
}

TEST(ZipLoadBudget, XlsbReaderRejectsCumulativePayload) {
  const Bytes package = BuildBudgetPackage(/*xlsb=*/true);
  ASSERT_FALSE(package.empty());
  ExpectBudgetEntriesFitIndividually(package);

  auto loaded = io::xlsb::read_xlsb(SpanOf(package));
  ASSERT_FALSE(static_cast<bool>(loaded));
  EXPECT_EQ(loaded.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_NE(loaded.error().context.find("limit=total"), std::string::npos);
  EXPECT_NE(loaded.error().message.find("cumulative"), std::string::npos);
}

}  // namespace
}  // namespace formulon
