// C-ABI end-to-end regression for the cumulative ZIP extraction budget. The
// package is a real writer-produced XLSB workbook with three opaque Default-
// typed parts appended through miniz, so fm_workbook_load exercises format
// detection, the XLSB reader, and diagnostic propagation as one call.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

// One 90 MiB source buffer is borrowed for all three ZIP entries. The XLSB
// reader retains only the first two (~180 MiB) before refusing the third before
// allocation, so the package does not require four independent payload copies.
constexpr std::size_t kBudgetEntryBytes = 90ULL * 1024ULL * 1024ULL;
constexpr std::size_t kBudgetEntryCount = 3U;
constexpr std::size_t kPatternPeriod = 8ULL * 1024ULL;

constexpr std::array<std::string_view, kBudgetEntryCount> kBudgetEntryNames = {
    "xl/formulon-budget-0.budget",
    "xl/formulon-budget-1.budget",
    "xl/formulon-budget-2.budget",
};

formulon::io::ByteSpan SpanOf(const Bytes& bytes) {
  return formulon::io::ByteSpan{bytes.data(), bytes.size()};
}

Bytes BuildPatternedPayload() {
  std::array<std::uint8_t, kPatternPeriod> period{};
  std::uint32_t state = 0x9E3779B9U;
  for (std::uint8_t& byte : period) {
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

Bytes BuildXlsbBudgetPackage() {
  formulon::Workbook workbook = formulon::Workbook::create();
  auto base_or = formulon::io::xlsb::write_xlsb(workbook);
  EXPECT_TRUE(static_cast<bool>(base_or)) << "write_xlsb failed: " << base_or.error().message;
  if (!base_or) {
    return {};
  }
  const Bytes& base = base_or.value();
  const Bytes payload = BuildPatternedPayload();

  formulon::io::ZipReader input;
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

TEST(FormulonCApiZipBudget, XlsbLoadReturnsStableFileTooLarge) {
  const Bytes package = BuildXlsbBudgetPackage();
  ASSERT_FALSE(package.empty());

  // Prove that C API format detection sees the binary workbook envelope. The
  // call below must therefore preserve the XLSB reader's I/O error, rather
  // than falling through to an earlier OOXML/parser diagnostic.
  formulon::io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  EXPECT_TRUE(zip.has_entry("xl/workbook.bin"));

  fm_workbook_t* workbook = reinterpret_cast<fm_workbook_t*>(static_cast<std::uintptr_t>(1U));
  const fm_status_t status = fm_workbook_load(package.data(), package.size(), &workbook);
  const std::string context = fm_last_error_context();
  const std::string message = fm_last_error_message();

  EXPECT_EQ(status, static_cast<fm_status_t>(5002));
  EXPECT_EQ(status, static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoFileTooLarge));
  EXPECT_EQ(workbook, nullptr);
  EXPECT_EQ(std::string_view(fm_status_string(status)), "kIoFileTooLarge");
  EXPECT_NE(context.find("limit=total"), std::string::npos);
  EXPECT_NE(context.find("entry=xl/formulon-budget-"), std::string::npos);
  EXPECT_NE(message.find("cumulative"), std::string::npos);
}

}  // namespace
