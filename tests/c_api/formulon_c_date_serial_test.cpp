//
// C ABI tests for `fm_convert_civil_to_serial` / `fm_convert_serial_to_civil`.
// The cases live in tools/api/date_serial_cases.json, shared with the JS and
// Python drivers; every valid row is checked against the engine's own
// DATE / TIME / YEAR .. SECOND evaluation.

#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <fstream>
#include <limits>
#include <sstream>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// Just enough JSON for the case table: no escapes beyond \" and \\, and numbers via strtod.
struct Json {
  enum class Kind { kNull, kBool, kNumber, kString, kArray, kObject };
  Kind kind = Kind::kNull;
  bool boolean = false;
  double number = 0.0;
  std::string text;
  std::vector<Json> items;
  std::vector<std::string> keys;

  const Json* Find(const std::string& key) const {
    for (std::size_t i = 0; i < keys.size(); ++i) {
      if (keys[i] == key) {
        return &items[i];
      }
    }
    return nullptr;
  }
  const Json& At(const std::string& key) const {
    static const Json kMissing;
    const Json* found = Find(key);
    return found != nullptr ? *found : kMissing;
  }
};

class JsonReader {
 public:
  explicit JsonReader(const std::string& src) : src_(src) {}
  Json Parse() { return Value(); }

 private:
  void Skip() {
    while (pos_ < src_.size() &&
           (src_[pos_] == ' ' || src_[pos_] == '\n' || src_[pos_] == '\r' || src_[pos_] == '\t')) {
      ++pos_;
    }
  }
  std::string String() {
    std::string out;
    ++pos_;  // opening quote
    while (pos_ < src_.size() && src_[pos_] != '"') {
      if (src_[pos_] == '\\') {
        ++pos_;
      }
      out.push_back(src_[pos_++]);
    }
    ++pos_;
    return out;
  }
  Json Value() {
    Skip();
    Json v;
    const char c = src_[pos_];
    if (c == '{' || c == '[') {
      const char close = c == '{' ? '}' : ']';
      v.kind = c == '{' ? Json::Kind::kObject : Json::Kind::kArray;
      ++pos_;
      Skip();
      while (src_[pos_] != close) {
        if (c == '{') {
          Skip();
          v.keys.push_back(String());
          Skip();
          ++pos_;  // colon
        }
        v.items.push_back(Value());
        Skip();
        if (src_[pos_] == ',') {
          ++pos_;
          Skip();
        }
      }
      ++pos_;
    } else if (c == '"') {
      v.kind = Json::Kind::kString;
      v.text = String();
    } else if (src_.compare(pos_, 4, "true") == 0 || src_.compare(pos_, 5, "false") == 0) {
      v.kind = Json::Kind::kBool;
      v.boolean = c == 't';
      pos_ += v.boolean ? 4 : 5;
    } else if (src_.compare(pos_, 4, "null") == 0) {
      pos_ += 4;
    } else {
      v.kind = Json::Kind::kNumber;
      char* end = nullptr;
      v.number = std::strtod(src_.c_str() + pos_, &end);
      pos_ = static_cast<std::size_t>(end - src_.c_str());
    }
    return v;
  }

  const std::string& src_;
  std::size_t pos_ = 0;
};

bool LoadCases(Json* out) {
  // tests/c_api/<this file> -> repository root.
  std::string path = __FILE__;
  path = path.substr(0, path.rfind("tests/c_api/")) + "tools/api/date_serial_cases.json";
  std::ifstream in(path);
  if (!in) {
    return false;
  }
  std::stringstream buffer;
  buffer << in.rdbuf();
  const std::string src = buffer.str();
  *out = JsonReader(src).Parse();
  return true;
}

// A field is a JSON number, or a string standing for a non-finite or signed-zero number.
double Num(const Json& j) {
  if (j.kind == Json::Kind::kNumber) {
    return j.number;
  }
  if (j.text == "NaN") {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (j.text == "Infinity") {
    return std::numeric_limits<double>::infinity();
  }
  if (j.text == "-Infinity") {
    return -std::numeric_limits<double>::infinity();
  }
  return -0.0;
}

fm_civil_time_t ToCivil(const Json& j) {
  return fm_civil_time_t{static_cast<int32_t>(Num(j.At("year"))),   static_cast<int32_t>(Num(j.At("month"))),
                         static_cast<int32_t>(Num(j.At("day"))),    static_cast<int32_t>(Num(j.At("hour"))),
                         static_cast<int32_t>(Num(j.At("minute"))), static_cast<int32_t>(Num(j.At("second")))};
}

void ExpectCivil(const fm_civil_time_t& got, const fm_civil_time_t& want) {
  EXPECT_EQ(got.year, want.year);
  EXPECT_EQ(got.month, want.month);
  EXPECT_EQ(got.day, want.day);
  EXPECT_EQ(got.hour, want.hour);
  EXPECT_EQ(got.minute, want.minute);
  EXPECT_EQ(got.second, want.second);
}

// Evaluates `formula` in a fresh workbook with the given date system; A1 holds `a1`.
double Evaluate(bool date1904, double a1, const std::string& formula) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_set_date1904(wb.handle, date1904 ? 1 : 0), 0);
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, a1), 0);
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_evaluate_formula(wb.handle, 0, 0, 1, formula.c_str(), &v), 0) << formula;
  EXPECT_EQ(v.kind, FM_VAL_NUMBER) << formula;
  return v.u.number;
}

void CheckCivilToSerial(const Json& row, bool date1904, const Json& expect) {
  const fm_civil_time_t civil = ToCivil(row.At("input"));
  double serial = -12345.0;
  const fm_status_t rc = fm_convert_civil_to_serial(&civil, date1904 ? 1 : 0, &serial);
  if (expect.Find("error") != nullptr) {
    EXPECT_EQ(rc, kInvalidArgument);
    EXPECT_EQ(serial, -12345.0) << "out must be untouched on failure";
    return;
  }
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  EXPECT_NEAR(serial, expect.At("value").number, 1e-9);
  const std::string formula = "=DATE(" + std::to_string(civil.year) + "," + std::to_string(civil.month) + "," +
                              std::to_string(civil.day) + ")+TIME(" + std::to_string(civil.hour) + "," +
                              std::to_string(civil.minute) + "," + std::to_string(civil.second) + ")";
  EXPECT_NEAR(serial, Evaluate(date1904, 0.0, formula), 1e-9) << formula;
}

void CheckSerialToCivil(const Json& row, bool date1904, const Json& expect) {
  const double serial = Num(row.At("input").At("serial"));
  fm_civil_time_t civil{-1, -1, -1, -1, -1, -1};
  const fm_status_t rc = fm_convert_serial_to_civil(serial, date1904 ? 1 : 0, &civil);
  if (expect.Find("error") != nullptr) {
    EXPECT_EQ(rc, kInvalidArgument);
    ExpectCivil(civil, fm_civil_time_t{-1, -1, -1, -1, -1, -1});
    return;
  }
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ExpectCivil(civil, ToCivil(expect.At("value")));
  EXPECT_EQ(civil.year, Evaluate(date1904, serial, "=YEAR(A1)"));
  EXPECT_EQ(civil.month, Evaluate(date1904, serial, "=MONTH(A1)"));
  EXPECT_EQ(civil.day, Evaluate(date1904, serial, "=DAY(A1)"));
  EXPECT_EQ(civil.hour, Evaluate(date1904, serial, "=HOUR(A1)"));
  EXPECT_EQ(civil.minute, Evaluate(date1904, serial, "=MINUTE(A1)"));
  EXPECT_EQ(civil.second, Evaluate(date1904, serial, "=SECOND(A1)"));
}

bool RunsInC(const Json& row) {
  const Json* only = row.Find("only");
  if (only != nullptr) {
    bool listed = false;
    for (const Json& name : only->items) {
      listed = listed || name.text == "c";
    }
    if (!listed) {
      return false;
    }
  }
  // Layer 1 rows are rejected inside a binding and never reach C.
  const Json* layer = row.At("expect").At("c").Find("layer");
  return layer == nullptr || layer->number != 1.0;
}

}  // namespace

TEST(FormulonCApiDateSerial, EveryCaseInTheSharedTable) {
  Json table;
  ASSERT_TRUE(LoadCases(&table)) << "tools/api/date_serial_cases.json not found";
  const Json& cases = table.At("cases");
  ASSERT_GE(cases.items.size(), 40U);
  std::size_t ran = 0;
  for (const Json& row : cases.items) {
    if (!RunsInC(row)) {
      continue;
    }
    SCOPED_TRACE(row.At("id").text);
    const bool date1904 = row.At("date1904").boolean;
    const Json& expect = row.At("expect").At("c");
    if (row.At("direction").text == "civil_to_serial") {
      CheckCivilToSerial(row, date1904, expect);
    } else {
      CheckSerialToCivil(row, date1904, expect);
    }
    ++ran;
  }
  EXPECT_GE(ran, 40U);
}

TEST(FormulonCApiDateSerial, NullPointersAreRejected) {
  const fm_civil_time_t civil{2026, 1, 1, 0, 0, 0};
  double serial = 0.0;
  fm_civil_time_t out{};
  EXPECT_EQ(fm_convert_civil_to_serial(nullptr, 0, &serial), kNullPointer);
  EXPECT_EQ(fm_convert_civil_to_serial(&civil, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_convert_serial_to_civil(1.0, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_convert_serial_to_civil(1.0, 0, &out), 0);
}

TEST(FormulonCApiDateSerial, AnyNonZeroFlagSelectsThe1904System) {
  const fm_civil_time_t civil{1904, 1, 1, 0, 0, 0};
  double serial = -1.0;
  ASSERT_EQ(fm_convert_civil_to_serial(&civil, 7, &serial), 0);
  EXPECT_EQ(serial, 0.0);
}
