//
// Stable C ABI structured-logging configuration tests.
//
// Formulon ships as an embedded library, so the bytes it writes to the
// host's stderr must be under the host's control. These tests pin the
// two properties that give the host that control: the engine is silent
// until the threshold is lowered, and a registered sink receives every
// record instead of stderr.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "utils/structured_log.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

extern "C" void CollectRecord(const char* record, size_t len, void* user_data) {
  static_cast<std::vector<std::string>*>(user_data)->emplace_back(record, len);
}

// The log configuration is process-wide, so every test restores the shipped
// default afterwards instead of leaking a threshold into its neighbours.
class CApiLogging : public ::testing::Test {
 protected:
  void TearDown() override {
    EXPECT_EQ(fm_set_log_sink(nullptr, nullptr), 0);
    EXPECT_EQ(fm_set_log_min_level(FM_LOG_LEVEL_OFF), 0);
  }
};

// Exercises the XLSB writer's per-cell downgrade diagnostics, which are the
// noisiest producer on the save path.
void BuildFormulaWorkbook(fm_workbook_t* handle) {
  for (uint32_t row = 0; row < 64U; ++row) {
    ASSERT_EQ(fm_workbook_set_formula(handle, 0, row, 0, "=ROW()+COLUMN()"), 0);
  }
  ASSERT_EQ(fm_workbook_recalc(handle), 0);
}

}  // namespace

TEST_F(CApiLogging, DefaultConfigurationEmitsNothingForALoadSaveRoundTrip) {
  // Install a sink but leave the threshold at its default. Anything the
  // engine would have written to stderr arrives here instead, so an empty
  // collection is the assertion that the default is silent.
  std::vector<std::string> records;
  ASSERT_EQ(fm_set_log_sink(&CollectRecord, &records), 0);

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  BuildFormulaWorkbook(wb.handle);

  BufferGuard xlsx;
  ASSERT_EQ(fm_workbook_save(wb.handle, &xlsx.data, &xlsx.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(xlsx.data, xlsx.len, &reloaded.handle), 0);

  BufferGuard xlsb;
  ASSERT_EQ(fm_workbook_save_ex(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb.data, &xlsb.len), 0);
  WorkbookGuard from_xlsb;
  ASSERT_EQ(fm_workbook_load(xlsb.data, xlsb.len, &from_xlsb.handle), 0);

  EXPECT_TRUE(records.empty()) << "first unexpected record: " << records.front();
}

TEST_F(CApiLogging, LoweringTheThresholdRoutesRecordsToTheRegisteredSink) {
  std::vector<std::string> records;
  ASSERT_EQ(fm_set_log_sink(&CollectRecord, &records), 0);
  ASSERT_EQ(fm_set_log_min_level(FM_LOG_LEVEL_DEBUG), 0);

  formulon::StructuredLog("binding.log_sink_probe").field("detail", std::string_view("value")).warn();

  ASSERT_EQ(records.size(), 1U);
  EXPECT_NE(records[0].find("\"event\":\"binding.log_sink_probe\""), std::string::npos) << records[0];
  EXPECT_NE(records[0].find("\"level\":\"warn\""), std::string::npos) << records[0];
  // Every record is one complete JSON line, so a host can forward it
  // verbatim into a line-oriented log stream.
  EXPECT_EQ(records[0].back(), '\n');
}

TEST_F(CApiLogging, ThresholdSuppressesRecordsBelowIt) {
  std::vector<std::string> records;
  ASSERT_EQ(fm_set_log_sink(&CollectRecord, &records), 0);
  ASSERT_EQ(fm_set_log_min_level(FM_LOG_LEVEL_ERROR), 0);

  formulon::StructuredLog("binding.below_threshold").warn();
  EXPECT_TRUE(records.empty());

  formulon::StructuredLog("binding.at_threshold").error();
  ASSERT_EQ(records.size(), 1U);
  EXPECT_NE(records[0].find("binding.at_threshold"), std::string::npos) << records[0];
}

TEST_F(CApiLogging, ClearingTheSinkStopsDelivery) {
  std::vector<std::string> records;
  ASSERT_EQ(fm_set_log_sink(&CollectRecord, &records), 0);
  ASSERT_EQ(fm_set_log_min_level(FM_LOG_LEVEL_DEBUG), 0);
  formulon::StructuredLog("binding.before_clear").error();
  ASSERT_EQ(records.size(), 1U);

  // Clearing restores the engine's stderr fallback, so raise the threshold
  // back to `OFF` first and keep this test off the host's stderr.
  ASSERT_EQ(fm_set_log_min_level(FM_LOG_LEVEL_OFF), 0);
  ASSERT_EQ(fm_set_log_sink(nullptr, nullptr), 0);
  formulon::StructuredLog("binding.after_clear").error();
  EXPECT_EQ(records.size(), 1U);
}

TEST_F(CApiLogging, MinLevelRejectsOrdinalsOutsideTheEnum) {
  const int32_t invalid[] = {-1, 5, 1 << 20};
  for (const int32_t level : invalid) {
    EXPECT_EQ(fm_set_log_min_level(level), static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << "level=" << level;
    EXPECT_STREQ(fm_last_error_message(), "fm_set_log_min_level: unknown level");
  }
  // Every documented ordinal is accepted.
  for (int32_t level = FM_LOG_LEVEL_DEBUG; level <= FM_LOG_LEVEL_OFF; ++level) {
    EXPECT_EQ(fm_set_log_min_level(level), 0) << "level=" << level;
  }
}
