#include "eval/pivot_formula.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_formula_host.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::eval {
namespace {

using pivot::FormulaRef;
using pivot::FormulaScope;
using pivot::PivotCache;
using pivot::PivotCacheField;

constexpr std::size_t kRegion = 0;
constexpr std::size_t kAmt = 1;
constexpr std::size_t kQty = 2;
constexpr std::size_t kSuryo = 3;
constexpr std::size_t kAmtNet = 4;
constexpr std::size_t kQ1 = 5;
constexpr std::size_t kTrue = 6;
constexpr std::size_t kQuote = 7;
constexpr std::size_t kSalesArea = 9;

Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

PivotCache build_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  std::vector<Value> regions;
  regions.push_back(owned_text(cache, "North"));
  regions.push_back(owned_text(cache, "South"));
  regions.push_back(owned_text(cache, "East"));
  regions.push_back(Value::number(2024));
  cache.mutable_fields().push_back(PivotCacheField{"Region", std::move(regions)});
  for (const char* name : {"Amt", "Qty", "数量", "Amt Net", "Q1", "TRUE", "It's"}) {
    cache.mutable_fields().push_back(PivotCacheField{name, {}});
  }
  std::vector<Value> products;
  products.push_back(owned_text(cache, "Widget"));
  cache.mutable_fields().push_back(PivotCacheField{"Product", std::move(products)});
  std::vector<Value> areas;
  areas.push_back(owned_text(cache, "North East"));
  areas.push_back(owned_text(cache, "West"));
  cache.mutable_fields().push_back(PivotCacheField{"Sales Area", std::move(areas)});
  return cache;
}

class PivotFormula : public ::testing::Test {
 protected:
  FormulaScope field_scope() const { return FormulaScope{&cache_, std::nullopt}; }
  FormulaScope item_scope() const { return FormulaScope{&cache_, kRegion}; }

  std::uint32_t compile_ok(std::string_view formula, const FormulaScope& scope) {
    auto id = host_.compile(formula, scope, &refs_, nullptr);
    EXPECT_TRUE(id.has_value()) << formula << ": " << (id ? "" : id.error().message);
    return id ? id.value() : 0;
  }

  std::string canonical(std::string_view formula, const FormulaScope& scope) {
    std::string out = "unset";
    auto id = host_.compile(formula, scope, &refs_, &out);
    EXPECT_TRUE(id.has_value()) << formula << ": " << (id ? "" : id.error().message);
    return out;
  }

  FormulonErrorCode compile_error(std::string_view formula, const FormulaScope& scope) {
    std::vector<FormulaRef> refs;
    auto id = host_.compile(formula, scope, &refs, nullptr);
    EXPECT_FALSE(id.has_value()) << formula;
    return id ? FormulonErrorCode::kOk : id.error().code;
  }

  const ScopedEvalProfile profile_{mac_365_ja_jp_profile()};
  PivotCache cache_ = build_cache();
  PivotFormulaEvaluator host_;
  std::vector<FormulaRef> refs_;
};

void expect_ref(const FormulaRef& ref, FormulaRef::Kind kind, std::size_t field, std::size_t item = 0,
                std::int32_t position = 0) {
  EXPECT_EQ(ref.kind, kind);
  EXPECT_EQ(ref.field, field);
  EXPECT_EQ(ref.item, item);
  EXPECT_EQ(ref.position, position);
}

TEST_F(PivotFormula, CalculatedFieldDividesFieldSums) {
  const std::uint32_t id = compile_ok("=Amt/Qty", field_scope());
  ASSERT_EQ(refs_.size(), 2U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kAmt);
  expect_ref(refs_[1], FormulaRef::Kind::Field, kQty);
  const Value values[] = {Value::number(630), Value::number(20)};
  const Value v = host_.evaluate(id, values, 2);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 31.5);
}

TEST_F(PivotFormula, NonAsciiAndQuotedFieldNames) {
  const std::uint32_t id = compile_ok("IF(数量>2,'Amt Net',0)", field_scope());
  ASSERT_EQ(refs_.size(), 2U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kSuryo);
  expect_ref(refs_[1], FormulaRef::Kind::Field, kAmtNet);
  const Value above[] = {Value::number(7), Value::number(70)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, above, 2).as_number(), 70.0);
  const Value below[] = {Value::number(1), Value::number(70)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, below, 2).as_number(), 0.0);
}

TEST_F(PivotFormula, FieldNameMatchIgnoresCase) {
  compile_ok("amt*2", field_scope());
  ASSERT_EQ(refs_.size(), 1U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kAmt);
}

TEST_F(PivotFormula, RepeatedNameSharesOneRef) {
  const std::uint32_t id = compile_ok("Amt+Amt*Qty", field_scope());
  ASSERT_EQ(refs_.size(), 2U);
  const Value values[] = {Value::number(3), Value::number(4)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, values, 2).as_number(), 15.0);
}

TEST_F(PivotFormula, CalculatedItemSumsQualifiedItems) {
  const std::uint32_t id = compile_ok("Region[North]+Region[South]", item_scope());
  ASSERT_EQ(refs_.size(), 2U);
  expect_ref(refs_[0], FormulaRef::Kind::Item, kRegion, 0);
  expect_ref(refs_[1], FormulaRef::Kind::Item, kRegion, 1);
  const Value values[] = {Value::number(210), Value::number(350)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, values, 2).as_number(), 560.0);
}

TEST_F(PivotFormula, CalculatedItemReferenceForms) {
  compile_ok("Region[-1]", item_scope());
  ASSERT_EQ(refs_.size(), 1U);
  expect_ref(refs_[0], FormulaRef::Kind::ItemOffset, kRegion, 0, -1);

  compile_ok("Region[+1]+Region[2]+Region[ 'East' ]+North", item_scope());
  ASSERT_EQ(refs_.size(), 4U);
  expect_ref(refs_[0], FormulaRef::Kind::ItemOffset, kRegion, 0, 1);
  expect_ref(refs_[1], FormulaRef::Kind::ItemAt, kRegion, 0, 2);
  expect_ref(refs_[2], FormulaRef::Kind::Item, kRegion, 2);
  expect_ref(refs_[3], FormulaRef::Kind::Item, kRegion, 0);
}

TEST_F(PivotFormula, NamesThatLexAsCellRefOrBoolean) {
  const std::uint32_t id = compile_ok("Q1+TRUE", field_scope());
  ASSERT_EQ(refs_.size(), 2U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kQ1);
  expect_ref(refs_[1], FormulaRef::Kind::Field, kTrue);
  const Value values[] = {Value::number(1), Value::number(2)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, values, 2).as_number(), 3.0);

  // A function call keeps its name even when a field shares it.
  compile_ok("IF(TRUE(),Q1,0)", field_scope());
  ASSERT_EQ(refs_.size(), 1U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kQ1);
}

TEST_F(PivotFormula, DoubledQuoteEscapesInsideQuotedName) {
  const std::uint32_t id = compile_ok("'It''s'*2", field_scope());
  ASSERT_EQ(refs_.size(), 1U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kQuote);
  const Value values[] = {Value::number(4)};
  EXPECT_DOUBLE_EQ(host_.evaluate(id, values, 1).as_number(), 8.0);
}

TEST_F(PivotFormula, StringLiteralIsNotRewritten) {
  const std::uint32_t id = compile_ok("IF(Amt>0,\"Amt 'Qty'\",\"x\")", field_scope());
  ASSERT_EQ(refs_.size(), 1U);
  expect_ref(refs_[0], FormulaRef::Kind::Field, kAmt);
  const Value values[] = {Value::number(1)};
  const Value v = host_.evaluate(id, values, 1);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Amt 'Qty'");
}

TEST_F(PivotFormula, ErrorLiteralIsNotRewritten) {
  const std::uint32_t id = compile_ok("IFERROR(Amt,#N/A)", field_scope());
  ASSERT_EQ(refs_.size(), 1U);
  const Value values[] = {Value::error(ErrorCode::Div0)};
  const Value v = host_.evaluate(id, values, 1);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST_F(PivotFormula, DivisionByZeroPropagates) {
  const std::uint32_t id = compile_ok("Amt/Qty", field_scope());
  const Value zero[] = {Value::number(630), Value::number(0)};
  const Value v = host_.evaluate(id, zero, 2);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);

  const Value err_in[] = {Value::error(ErrorCode::Ref), Value::number(2)};
  const Value e = host_.evaluate(id, err_in, 2);
  ASSERT_TRUE(e.is_error());
  EXPECT_EQ(e.as_error(), ErrorCode::Ref);
}

TEST_F(PivotFormula, RejectsDisallowedFunctions) {
  for (const char* fn : {"GETPIVOTDATA", "DCOUNT",  "DCOUNTA", "DSUM",  "DAVERAGE", "DMIN",  "MAX",  "DPRODUCT",
                         "DSTDEV",       "DSTDEVP", "DVAR",    "DVARP", "INDEX",    "RAND",  "NOW",  "AREAS",
                         "ROWS",         "COLUMNS", "OFFSET",  "CELL",  "INDIRECT", "TODAY", "INFO", "RANDBETWEEN"}) {
    EXPECT_EQ(compile_error(std::string(fn) + "(Amt,1)", field_scope()), FormulonErrorCode::kInvalidArgument) << fn;
    EXPECT_EQ(compile_error(std::string("1+") + fn + "()", field_scope()), FormulonErrorCode::kInvalidArgument) << fn;
  }
  compile_ok("SUM(Amt,1)+MIN(Qty,2)", field_scope());
}

TEST_F(PivotFormula, RejectsReferencesNamesAndArrays) {
  const FormulonErrorCode kInvalid = FormulonErrorCode::kInvalidArgument;
  EXPECT_EQ(compile_error("A1*2", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Amt+B2:C3", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Amt*Rate", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Sheet1!Amt", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Table1[Col]", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("[Book.xlsx]Sheet1!A1+Amt", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("SUM({1,2})+Amt", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("'No Such Field'+1", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("_fmpv0+Amt", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("LET(x,Amt,x*2)", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Amt+", field_scope()), FormulonErrorCode::kParserUnexpectedToken);
  EXPECT_EQ(compile_error("'Amt", field_scope()), kInvalid);
}

TEST_F(PivotFormula, RejectsAllAndBlank) {
  const FormulonErrorCode kInvalid = FormulonErrorCode::kInvalidArgument;
  EXPECT_EQ(compile_error("All+Amt", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("'Blank'*2", field_scope()), kInvalid);
  EXPECT_EQ(compile_error("Region[All]", item_scope()), kInvalid);
  EXPECT_EQ(compile_error("blank", item_scope()), kInvalid);
}

TEST_F(PivotFormula, CalculatedItemScope) {
  const FormulonErrorCode kInvalid = FormulonErrorCode::kInvalidArgument;
  EXPECT_EQ(compile_error("Product[Widget]", item_scope()), kInvalid);
  EXPECT_EQ(compile_error("Amt*2", item_scope()), kInvalid);
  EXPECT_EQ(compile_error("Region[West]", item_scope()), kInvalid);
  EXPECT_EQ(compile_error("Region[99999999999]", item_scope()), kInvalid);
  // A calculated field formula cannot use item references.
  EXPECT_EQ(compile_error("Region[North]", field_scope()), kInvalid);
}

TEST_F(PivotFormula, RejectsOverlongFormula) {
  std::string formula = "Amt";
  while (formula.size() < 8192) {
    formula += "+1";
  }
  formula.resize(8193);
  EXPECT_EQ(compile_error(formula, field_scope()), FormulonErrorCode::kInvalidArgument);
}

TEST_F(PivotFormula, ScratchArenaIsResetPerEvaluation) {
  const std::uint32_t id = compile_ok("Amt&\"-\"&Qty&REPT(\"x\",20)", field_scope());
  const Value values[] = {Value::number(1.5), Value::number(2)};
  (void)host_.evaluate(id, values, 2);
  (void)host_.evaluate(id, values, 2);
  const std::size_t baseline = host_.scratch_bytes_allocated();
  for (int i = 0; i < 1000; ++i) {
    const Value v = host_.evaluate(id, values, 2);
    ASSERT_TRUE(v.is_text());
    EXPECT_EQ(v.as_text(), "1.5-2xxxxxxxxxxxxxxxxxxxx");
  }
  EXPECT_EQ(host_.scratch_bytes_allocated(), baseline);
}

TEST_F(PivotFormula, IdsAreIndependentPerFormula) {
  const std::uint32_t a = compile_ok("Amt*2", field_scope());
  const std::uint32_t b = compile_ok("Qty+Amt", field_scope());
  EXPECT_NE(a, b);
  const Value one[] = {Value::number(5)};
  EXPECT_DOUBLE_EQ(host_.evaluate(a, one, 1).as_number(), 10.0);
  const Value two[] = {Value::number(1), Value::number(5)};
  EXPECT_DOUBLE_EQ(host_.evaluate(b, two, 2).as_number(), 6.0);
}

TEST_F(PivotFormula, CanonicalQualifiesBareItemNames) {
  EXPECT_EQ(canonical("=North+South", item_scope()), "Region[North]+Region[South]");
  EXPECT_EQ(canonical("North*2+SUM(South,1)", item_scope()), "Region[North]*2+SUM(Region[South],1)");
  EXPECT_EQ(canonical("'North'+1", item_scope()), "Region['North']+1");
}

TEST_F(PivotFormula, CanonicalKeepsQualifiedPositionalAndLiterals) {
  EXPECT_EQ(canonical("=Region[-1]", item_scope()), "Region[-1]");
  EXPECT_EQ(canonical("Region[+1]+Region[2]", item_scope()), "Region[+1]+Region[2]");
  EXPECT_EQ(canonical("Region[North]+South", item_scope()), "Region[North]+Region[South]");
  EXPECT_EQ(canonical("IF(North>0,\"North\",#N/A)", item_scope()), "IF(Region[North]>0,\"North\",#N/A)");
}

TEST_F(PivotFormula, CanonicalQuotesFieldNamesThatNeedIt) {
  const FormulaScope scope{&cache_, kSalesArea};
  const std::string text = canonical("'North East'+West", scope);
  EXPECT_EQ(text, "'Sales Area'['North East']+'Sales Area'[West]");
  // The canonical form compiles back to the same references.
  const std::vector<FormulaRef> before = refs_;
  compile_ok(text, scope);
  ASSERT_EQ(refs_.size(), before.size());
  for (std::size_t i = 0; i < refs_.size(); ++i) {
    expect_ref(refs_[i], before[i].kind, before[i].field, before[i].item, before[i].position);
  }
}

TEST_F(PivotFormula, CanonicalFieldFormulaIsAsTyped) {
  EXPECT_EQ(canonical("=Amt/Qty", field_scope()), "Amt/Qty");
  EXPECT_EQ(canonical("IF(数量>2,'Amt Net',0)", field_scope()), "IF(数量>2,'Amt Net',0)");
}

}  // namespace
}  // namespace formulon::eval
