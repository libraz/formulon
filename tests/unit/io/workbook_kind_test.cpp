// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `io::WorkbookKind`. Spot-checks the content-type and
// default-extension lookups for all four variants (xlsx / xlsm / xltx /
// xltm). The four canonical content-type strings are referenced
// verbatim from [OPC] part 1 §10 / [ECMA-376]; if the engine ever
// changes them the rest of the I/O pipeline (reader detection, writer
// emission, oracle parity) will silently desync, so this test stays
// strict on byte-for-byte equality.

#include "io/workbook_kind.h"

#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace io {
namespace {

TEST(WorkbookKind, ContentTypeXlsx) {
  EXPECT_EQ(std::string_view(workbook_kind_content_type(WorkbookKind::kXlsx)),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
}

TEST(WorkbookKind, ContentTypeXlsm) {
  EXPECT_EQ(std::string_view(workbook_kind_content_type(WorkbookKind::kXlsm)),
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
}

TEST(WorkbookKind, ContentTypeXltx) {
  EXPECT_EQ(std::string_view(workbook_kind_content_type(WorkbookKind::kXltx)),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml");
}

TEST(WorkbookKind, ContentTypeXltm) {
  EXPECT_EQ(std::string_view(workbook_kind_content_type(WorkbookKind::kXltm)),
            "application/vnd.ms-excel.template.macroEnabled.main+xml");
}

TEST(WorkbookKind, DefaultExtensionXlsx) {
  EXPECT_EQ(std::string_view(workbook_kind_default_extension(WorkbookKind::kXlsx)), "xlsx");
}

TEST(WorkbookKind, DefaultExtensionXlsm) {
  EXPECT_EQ(std::string_view(workbook_kind_default_extension(WorkbookKind::kXlsm)), "xlsm");
}

TEST(WorkbookKind, DefaultExtensionXltx) {
  EXPECT_EQ(std::string_view(workbook_kind_default_extension(WorkbookKind::kXltx)), "xltx");
}

TEST(WorkbookKind, DefaultExtensionXltm) {
  EXPECT_EQ(std::string_view(workbook_kind_default_extension(WorkbookKind::kXltm)), "xltm");
}

TEST(WorkbookKind, EnumValuesAreStable) {
  // The numeric values are part of the (private) ABI contract for any
  // future C-API surface that exposes them. Pin them so unintended
  // reordering breaks this test instead of breaking downstream consumers
  // silently.
  EXPECT_EQ(static_cast<int>(WorkbookKind::kXlsx), 0);
  EXPECT_EQ(static_cast<int>(WorkbookKind::kXlsm), 1);
  EXPECT_EQ(static_cast<int>(WorkbookKind::kXltx), 2);
  EXPECT_EQ(static_cast<int>(WorkbookKind::kXltm), 3);
}

}  // namespace
}  // namespace io
}  // namespace formulon
