// `FM_VERSION_NUMBER` and `fm_get_version_number`: the library version packed
// as `MAJOR * 10000 + MINOR * 100 + PATCH`.

#include <cstdint>
#include <cstdio>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

TEST(FormulonCApiVersionNumber, MacroPacksTheThreeComponents) {
  EXPECT_EQ(FM_VERSION_NUMBER, FM_VERSION_MAJOR * 10000 + FM_VERSION_MINOR * 100 + FM_VERSION_PATCH);
}

TEST(FormulonCApiVersionNumber, FunctionEqualsTheMacro) {
  EXPECT_EQ(fm_get_version_number(), static_cast<uint32_t>(FM_VERSION_NUMBER));
}

TEST(FormulonCApiVersionNumber, AgreesWithTheVersionString) {
  unsigned major = 0;
  unsigned minor = 0;
  unsigned patch = 0;
  ASSERT_EQ(std::sscanf(fm_get_version(), "%u.%u.%u", &major, &minor, &patch), 3) << fm_get_version();
  EXPECT_EQ(fm_get_version_number(), major * 10000U + minor * 100U + patch);
}

TEST(FormulonCApiVersionNumber, LeavesTheLastErrorUntouched) {
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_sheet_count(nullptr, &count), FM_STATUS_BINDING_NULL_POINTER);
  const std::string message = fm_get_last_error_message();
  const std::string context = fm_get_last_error_context();
  ASSERT_FALSE(message.empty());
  static_cast<void>(fm_get_version_number());
  EXPECT_EQ(std::string(fm_get_last_error_message()), message);
  EXPECT_EQ(std::string(fm_get_last_error_context()), context);
}
