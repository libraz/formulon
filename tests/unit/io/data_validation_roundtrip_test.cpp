// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Reader tests for `<dataValidations>`.

#include <string>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "pugixml.hpp"
#include "sheet.h"

namespace formulon {
namespace io {
namespace {

pugi::xml_node ParseWorksheet(pugi::xml_document& doc, const std::string& body) {
  const std::string xml = "<worksheet>" + body + "</worksheet>";
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  return doc.child("worksheet");
}

TEST(DataValidationRoundTrip, EmptyYieldsEmptyVector) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<sheetData/>");
  auto out = read_data_validations(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  EXPECT_TRUE(out.value().empty());
}

TEST(DataValidationRoundTrip, ListValidation) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                            "<dataValidations count=\"1\">"
                            "<dataValidation type=\"list\" allowBlank=\"1\" showErrorMessage=\"1\""
                            " errorTitle=\"err\" error=\"pick from list\""
                            " sqref=\"A1:A10\">"
                            "<formula1>\"yes,no,maybe\"</formula1>"
                            "</dataValidation>"
                            "</dataValidations>");
  auto out = read_data_validations(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  const DataValidation& v = out.value()[0];
  EXPECT_EQ(v.type, 3U);  // list
  EXPECT_TRUE(v.allow_blank);
  EXPECT_TRUE(v.show_error_message);
  EXPECT_EQ(v.error_title, "err");
  EXPECT_EQ(v.error_message, "pick from list");
  EXPECT_EQ(v.formula1, "\"yes,no,maybe\"");
  ASSERT_EQ(v.ranges.size(), 1U);
  EXPECT_EQ(v.ranges[0].first_row, 0U);
  EXPECT_EQ(v.ranges[0].last_row, 9U);
}

TEST(DataValidationRoundTrip, BetweenWholeNumberOperator) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                            "<dataValidations>"
                            "<dataValidation type=\"whole\" operator=\"between\" sqref=\"B1\">"
                            "<formula1>1</formula1><formula2>100</formula2>"
                            "</dataValidation>"
                            "</dataValidations>");
  auto out = read_data_validations(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  EXPECT_EQ(out.value()[0].type, 1U);   // whole
  EXPECT_EQ(out.value()[0].op, 0U);     // between
  EXPECT_EQ(out.value()[0].formula1, "1");
  EXPECT_EQ(out.value()[0].formula2, "100");
}

TEST(DataValidationRoundTrip, MultipleRanges) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                            "<dataValidations>"
                            "<dataValidation type=\"custom\" sqref=\"A1:A5 C1:C5 E1\">"
                            "<formula1>A1&gt;0</formula1>"
                            "</dataValidation>"
                            "</dataValidations>");
  auto out = read_data_validations(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  ASSERT_EQ(out.value()[0].ranges.size(), 3U);
  EXPECT_EQ(out.value()[0].formula1, "A1>0");
}

TEST(DataValidationRoundTrip, AllOperatorVariantsParse) {
  pugi::xml_document doc;
  auto ws =
      ParseWorksheet(doc,
                     "<dataValidations>"
                     "<dataValidation type=\"decimal\" operator=\"notBetween\" sqref=\"A1\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"equal\" sqref=\"A2\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"notEqual\" sqref=\"A3\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"greaterThan\" sqref=\"A4\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"lessThan\" sqref=\"A5\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"greaterThanOrEqual\" sqref=\"A6\"><formula1>1</formula1></dataValidation>"
                     "<dataValidation type=\"decimal\" operator=\"lessThanOrEqual\" sqref=\"A7\"><formula1>1</formula1></dataValidation>"
                     "</dataValidations>");
  auto out = read_data_validations(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 7U);
  EXPECT_EQ(out.value()[0].op, 1U);
  EXPECT_EQ(out.value()[1].op, 2U);
  EXPECT_EQ(out.value()[2].op, 3U);
  EXPECT_EQ(out.value()[3].op, 4U);
  EXPECT_EQ(out.value()[4].op, 5U);
  EXPECT_EQ(out.value()[5].op, 6U);
  EXPECT_EQ(out.value()[6].op, 7U);
}

TEST(DataValidationRoundTrip, EmptySqrefRejected) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                            "<dataValidations>"
                            "<dataValidation type=\"list\" sqref=\"\"><formula1>1</formula1></dataValidation>"
                            "</dataValidations>");
  auto out = read_data_validations(ws);
  EXPECT_FALSE(static_cast<bool>(out));
  EXPECT_EQ(out.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
