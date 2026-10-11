//
// C ABI collation surface: fm_text_compare and fm_collation_classes.

#include <cstddef>
#include <cstdint>
#include <limits>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPtr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

constexpr int32_t kErrDiv0 = 1;
constexpr int32_t kErrValue = 2;
constexpr int32_t kErrNA = 6;

fm_value_t blank() {
  fm_value_t v{};
  v.kind = FM_VAL_BLANK;
  return v;
}

fm_value_t number(double d) {
  fm_value_t v{};
  v.kind = FM_VAL_NUMBER;
  v.u.number = d;
  return v;
}

fm_value_t boolean(int32_t b) {
  fm_value_t v{};
  v.kind = FM_VAL_BOOL;
  v.u.boolean = b;
  return v;
}

fm_value_t text(const char* s) {
  fm_value_t v{};
  v.kind = FM_VAL_TEXT;
  v.u.text = s;
  return v;
}

fm_value_t error(int32_t code) {
  fm_value_t v{};
  v.kind = FM_VAL_ERROR;
  v.u.error_code = code;
  return v;
}

int32_t compare(const char* a, const char* b, const char* profile, fm_collation_strength_t strength) {
  int32_t order = 99;
  EXPECT_EQ(fm_text_compare(a, b, profile, strength, &order), 0);
  return order;
}

TEST(FormulonCApiCollation, StrengthAndArgumentOrder) {
  const char* halfwidth = "A";
  const char* fullwidth = "\xEF\xBC\xA1";  // U+FF21
  EXPECT_NE(compare(halfwidth, fullwidth, "mac-365-zh_CN", FM_COLLATION_COMPARE), 0);
  EXPECT_EQ(compare(halfwidth, fullwidth, "mac-365-zh_CN", FM_COLLATION_DYNAMIC), 0);

  EXPECT_EQ(compare("a", "b", "mac-365-en_US", FM_COLLATION_COMPARE), -1);
  EXPECT_EQ(compare("b", "a", "mac-365-en_US", FM_COLLATION_COMPARE), 1);
  EXPECT_EQ(compare("a", "b", "mac-365-en_US", FM_COLLATION_DYNAMIC), -1);
  EXPECT_EQ(compare("b", "a", "mac-365-en_US", FM_COLLATION_CLASSIC_EXACT), 1);
  EXPECT_EQ(compare("a", "a", "win-365-en_US", FM_COLLATION_COMPARE), 0);
}

TEST(FormulonCApiCollation, ClassesRoundTrip) {
  const fm_value_t values[] = {
      text("b"),   number(1.0), text("a"),       blank(),    boolean(1),       error(kErrNA), number(-0.0), text("1"),
      number(0.0), text("a"),   error(kErrDiv0), boolean(0), error(kErrValue), blank(),       number(1.0),
  };
  constexpr size_t kCount = sizeof(values) / sizeof(values[0]);
  uint32_t ids[kCount] = {};
  uint32_t classes = 0;
  ASSERT_EQ(fm_collation_classes(values, kCount, "mac-365-en_US", FM_COLLATION_DYNAMIC, ids, &classes), 0);
  // Numbers {0, 1}, text {"1", "a", "b"}, bool {FALSE, TRUE}, errors {#DIV/0!, #VALUE!, #N/A}, blank.
  EXPECT_EQ(classes, 11U);
  EXPECT_EQ(ids[0], 4U);    // "b"
  EXPECT_EQ(ids[1], 1U);    // 1
  EXPECT_EQ(ids[2], 3U);    // "a"
  EXPECT_EQ(ids[3], 10U);   // blank
  EXPECT_EQ(ids[4], 6U);    // TRUE
  EXPECT_EQ(ids[5], 9U);    // #N/A
  EXPECT_EQ(ids[6], 0U);    // -0 shares the class of 0
  EXPECT_EQ(ids[7], 2U);    // "1"
  EXPECT_EQ(ids[8], 0U);    // 0
  EXPECT_EQ(ids[9], 3U);    // "a"
  EXPECT_EQ(ids[10], 7U);   // #DIV/0!
  EXPECT_EQ(ids[11], 5U);   // FALSE
  EXPECT_EQ(ids[12], 8U);   // #VALUE!
  EXPECT_EQ(ids[13], 10U);  // blank
  EXPECT_EQ(ids[14], 1U);   // 1

  uint32_t empty_classes = 7;
  EXPECT_EQ(fm_collation_classes(nullptr, 0, "mac-365-en_US", FM_COLLATION_COMPARE, nullptr, &empty_classes), 0);
  EXPECT_EQ(empty_classes, 0U);
}

TEST(FormulonCApiCollation, RejectsBadInput) {
  const char* profile = "mac-365-en_US";
  int32_t order = 0;
  EXPECT_EQ(fm_text_compare(nullptr, "b", profile, FM_COLLATION_COMPARE, &order), kNullPtr);
  EXPECT_EQ(fm_text_compare("a", nullptr, profile, FM_COLLATION_COMPARE, &order), kNullPtr);
  EXPECT_EQ(fm_text_compare("a", "b", nullptr, FM_COLLATION_COMPARE, &order), kNullPtr);
  EXPECT_EQ(fm_text_compare("a", "b", profile, FM_COLLATION_COMPARE, nullptr), kNullPtr);
  EXPECT_EQ(fm_text_compare("a", "b", "mac-365-xx_XX", FM_COLLATION_COMPARE, &order), kInvalid);
  EXPECT_EQ(fm_text_compare("a", "b", profile, static_cast<fm_collation_strength_t>(3), &order), kInvalid);
  EXPECT_EQ(fm_text_compare("a", "b", profile, static_cast<fm_collation_strength_t>(-1), &order), kInvalid);

  const fm_value_t ok[] = {number(1.0), text("a")};
  uint32_t ids[2] = {};
  uint32_t classes = 0;
  EXPECT_EQ(fm_collation_classes(nullptr, 2, profile, FM_COLLATION_COMPARE, ids, &classes), kNullPtr);
  EXPECT_EQ(fm_collation_classes(ok, 2, profile, FM_COLLATION_COMPARE, nullptr, &classes), kNullPtr);
  EXPECT_EQ(fm_collation_classes(ok, 2, nullptr, FM_COLLATION_COMPARE, ids, &classes), kNullPtr);
  EXPECT_EQ(fm_collation_classes(ok, 2, profile, FM_COLLATION_COMPARE, ids, nullptr), kNullPtr);
  EXPECT_EQ(fm_collation_classes(ok, 2, "nope", FM_COLLATION_COMPARE, ids, &classes), kInvalid);
  EXPECT_EQ(fm_collation_classes(ok, 2, profile, static_cast<fm_collation_strength_t>(3), ids, &classes), kInvalid);

  const fm_value_t null_text[] = {number(1.0), text(nullptr)};
  // value_from_fm reports a NULL text pointer as a null-pointer error.
  EXPECT_EQ(fm_collation_classes(null_text, 2, profile, FM_COLLATION_COMPARE, ids, &classes), kNullPtr);
  const fm_value_t infinite[] = {number(std::numeric_limits<double>::infinity())};
  EXPECT_EQ(fm_collation_classes(infinite, 1, profile, FM_COLLATION_COMPARE, ids, &classes), kInvalid);
  const fm_value_t nan_value[] = {number(std::numeric_limits<double>::quiet_NaN())};
  EXPECT_EQ(fm_collation_classes(nan_value, 1, profile, FM_COLLATION_COMPARE, ids, &classes), kInvalid);
  fm_value_t array_kind{};
  array_kind.kind = FM_VAL_ARRAY;
  EXPECT_EQ(fm_collation_classes(&array_kind, 1, profile, FM_COLLATION_COMPARE, ids, &classes), kInvalid);
  const fm_value_t bad_error[] = {error(9999)};
  EXPECT_EQ(fm_collation_classes(bad_error, 1, profile, FM_COLLATION_COMPARE, ids, &classes), kInvalid);
}

}  // namespace
