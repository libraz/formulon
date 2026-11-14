#include "io/xlsb/pivot_cache_bridge.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <optional>
#include <string>
#include <utility>

#include "io/cell_parser.h"
#include "io/ooxml_defs.h"
#include "io/sqref_text.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_cache_definition_records.h"
#include "io/xlsb/pivot_cache_group_records.h"
#include "io/xlsb/pivot_cache_records.h"
#include "io/xlsb/pivot_calc_item_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xsd_double.h"
#include "parser/ast.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/record_access.h"
#include "utils/date_time.h"
#include "utils/number_text.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace pcd = pivot_cache_definition;
using cache_records::CellKind;
using cache_records::RecordCell;

constexpr const char* kContext = "context=xlsb_pivot_cache_bridge";
/// The cache is workbook-scoped; refusals name sheet index 0.
constexpr const char* kNotWritableKind = "pivot_cache";

/// BrtBeginPCDFAtbl bits the definition codec names only in combination:
/// `containsSemiMixedTypes` and `containsString`.
constexpr std::uint16_t kAtblSemiMixed = 0x0001U;
constexpr std::uint16_t kAtblString = 0x0008U;

/// Excel writes 4 or more numbers, strings or dates in a row as one BrtBeginPCDIRun and fewer as single records
/// (3 numbers were 3 records; 4 strings, 4 dates and 12 dates one run each); booleans and errors always single.
constexpr std::size_t kMinRunLength = 4;

/// Attribute values Excel writes on the root of an xlsx cache definition that carries `xr:uid`.
constexpr std::string_view kXrPrefix = "xr";

/// The cache extLst, as the xlsx reader keeps it, for the two extensions the definition records hold.
constexpr std::string_view kExtLstOpen = "<extLst>";
constexpr std::string_view kExtLstClose = "</extLst>";
constexpr std::string_view kPcd14Ext =
    "<ext uri=\"{725AE2AE-9491-48be-B2B4-4EB974FC3084}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotCacheDefinition/></ext>";
constexpr std::string_view kDataSourceCleanExt =
    "<ext uri=\"{F4336242-56EE-4DD3-A5F7-E900642BF6E9}\" "
    "xmlns:xlpdsc=\"http://schemas.microsoft.com/office/spreadsheetml/2026/pivotDataSourceClean\">"
    "<xlpdsc:dataSourceClean>1</xlpdsc:dataSourceClean></ext>";
/// dataSourceClean payload of the `1` text, the only one Excel's parts show.
constexpr std::uint8_t kDataSourceCleanTrue = 1U;

/// BrtBeginPName and BrtBeginPNPair constants of every Excel-saved calculated item.
constexpr std::uint16_t kPNameFlags = 0x0001U;
constexpr std::uint8_t kPNPairFlags = 0U;
/// BrtBeginPRFilter bit set in every Excel-saved reference, `selected` or not.
constexpr std::uint32_t kFilterAlways = 0x000001U;

/// The PivotParsedFormula operators Excel's saved calculated items use; any other has no sample.
constexpr std::uint8_t kPtgAdd = 0x03U;
constexpr std::uint8_t kPtgSub = 0x04U;
constexpr std::uint8_t kPtgMul = 0x05U;
constexpr std::uint8_t kPtgDiv = 0x06U;
constexpr std::uint8_t kPtgPower = 0x07U;
constexpr std::uint8_t kPtgConcat = 0x08U;
constexpr std::uint8_t kPtgGt = 0x0DU;
constexpr std::uint8_t kPtgUminus = 0x13U;
constexpr std::uint8_t kPtgParen = 0x15U;

/// Infix text of a measured binary operator; empty for any other ptg.
std::string_view BinaryOperatorText(std::uint8_t ptg) {
  switch (ptg) {
    case kPtgAdd:
      return "+";
    case kPtgSub:
      return "-";
    case kPtgMul:
      return "*";
    case kPtgDiv:
      return "/";
    case kPtgPower:
      return "^";
    case kPtgConcat:
      return "&";
    case kPtgGt:
      return ">";
    default:
      return {};
  }
}

Error Corrupt(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb pivot cache: " + message, kContext);
}

Error Unmeasured(const std::string& what) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "unmeasured xlsb pivot cache state: " + what, kContext);
}

Error NotWritable() {
  return not_writable(kNotWritableKind, 0);
}

bool FlagAt(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

/// Parses a decimal integer no greater than `max`; nullopt on any other text.
std::optional<std::uint32_t> ParseUnsigned(std::string_view text, std::uint32_t max) {
  if (text.empty() || text.size() > 10U) {
    return std::nullopt;
  }
  std::uint64_t value = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return std::nullopt;
    }
    value = value * 10U + static_cast<std::uint64_t>(c - '0');
  }
  if (value > max) {
    return std::nullopt;
  }
  return static_cast<std::uint32_t>(value);
}

std::string NumberText(double value) {
  std::string out;
  append_xml_number(out, value);
  return out;
}

/// Excel spells `refreshedDate` with 17 significant digits, not the shorter cell-value form.
std::string RefreshedDateText(double value) {
  char text[32];
  const int length = format_general(text, sizeof(text), value, 17);
  return std::string(text, static_cast<std::size_t>(length));
}

std::string_view StoreText(pivot::PivotCache& cache, std::string text) {
  cache.mutable_text_storage().push_back(std::move(text));
  return cache.text_storage().back();
}

// ---------------------------------------------------------------------------
// Dates of the shared-item runs: Excel's calendar, where serial 0 of the 1900
// system is 1900-01-00 and serial 60 the phantom 1900-02-29.

Expected<double, Error> RunDateSerial(const pcd::Date& date, bool date1904) {
  if (date.hour != 0U || date.minute != 0U || date.second != 0U) {
    // A timed item is spelled like a timed record cell, from 1900-03-01 on.
    const cache_records::RecordDate timed{date.year, date.month, date.day, date.hour, date.minute, date.second};
    return record_date_serial(timed, date1904);
  }
  if (date.month < 1U || date.month > 12U) {
    return Corrupt("shared date item is not a calendar date");
  }
  const double serial = date_time::serial_from_ymd(date.year, date.month, date.day, date1904);
  const date_time::YMD back = date1904 ? date_time::ymd_from_serial(serial, true) : date_time::legacy_1900_ymd(serial);
  if (!pivot::is_valid_date_serial(serial, date1904) || back.y != date.year || back.m != date.month ||
      back.d != date.day) {
    return Corrupt("shared date item is not a calendar date");
  }
  return serial;
}

std::optional<pcd::Date> RunDate(double serial, bool date1904) {
  if (!pivot::is_valid_date_serial(serial, date1904)) {
    return std::nullopt;
  }
  if (serial != std::floor(serial)) {
    auto timed = record_date_from_serial(serial, date1904);
    if (!timed) {
      return std::nullopt;
    }
    const cache_records::RecordDate& t = timed.value();
    return pcd::Date{t.year, t.month, t.day, t.hour, t.minute, t.second};
  }
  const date_time::YMD ymd = date1904 ? date_time::ymd_from_serial(serial, true) : date_time::legacy_1900_ymd(serial);
  pcd::Date date;
  date.year = static_cast<std::uint16_t>(ymd.y);
  date.month = static_cast<std::uint16_t>(ymd.m);
  date.day = static_cast<std::uint8_t>(ymd.d);
  return date;
}

/// The ISO text the xlsx reader stores for a date serial; one that does not read back to the same serial (a
/// fraction of a second) has no text form.
Expected<std::string, Error> DateIso(double serial, bool date1904) {
  if (!pivot::is_valid_date_serial(serial, date1904)) {
    return Corrupt("cache date bound outside the date range");
  }
  std::string iso = pivot::group_date_iso(serial, date1904);
  if (pivot::group_date_serial(iso, date1904) != std::optional<double>(serial)) {
    return Unmeasured("cache date bound that is not a whole second");
  }
  return iso;
}

/// The serial of ISO text written back, when it names one exactly.
std::optional<double> IsoSerial(const std::string& iso, bool date1904) {
  const std::optional<double> serial = pivot::group_date_serial(iso, date1904);
  if (!serial || pivot::group_date_iso(*serial, date1904) != iso) {
    return std::nullopt;
  }
  return serial;
}

// ---------------------------------------------------------------------------
// Calculated-item formulas: `Field[Item]` names joined by operators, as the
// PtgSxName tokens and payload-free operators of a PivotParsedFormula spell
// them. Names that would need quoting have no Excel-saved sample.

bool IsNameByte(unsigned char c) {
  return c >= 0x80U || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' ||
         c == '.';
}

bool IsPlainField(std::string_view text) {
  if (text.empty()) {
    return false;
  }
  const auto first = static_cast<unsigned char>(text.front());
  if ((first >= '0' && first <= '9') || first == '.') {
    return false;
  }
  for (const char c : text) {
    if (!IsNameByte(static_cast<unsigned char>(c))) {
      return false;
    }
  }
  return true;
}

bool IsPlainItem(std::string_view text) {
  if (text.empty()) {
    return false;
  }
  for (const char c : text) {
    if (!IsNameByte(static_cast<unsigned char>(c))) {
      return false;
    }
  }
  return true;
}

/// Rendered text a calculated-item formula may reach; far above Excel's 8,192-character formula cap, it only
/// bounds a hostile token stream.
// @limit: engine
constexpr std::uint64_t kMaxCalcItemFormulaBytes = 65536U;

/// Infix text of a token stream; `names` holds the text of each PNames entry.
Expected<std::string, Error> RenderFormula(const std::vector<pivot_calc::FormulaToken>& tokens,
                                           const std::vector<std::string>& names) {
  ResourceBudget bytes(kMaxCalcItemFormulaBytes, "pivot_calc_item_formula_bytes");
  std::vector<std::string> stack;
  for (const pivot_calc::FormulaToken& token : tokens) {
    if (token.ptg == pivot_calc::kPtgSxName) {
      if (token.sx_name >= names.size()) {
        return Corrupt("calculated item formula names an entry past its names table");
      }
      RETURN_IF_ERROR(bytes.consume(names[token.sx_name].size()));
      stack.push_back(names[token.sx_name]);
      continue;
    }
    if (token.ptg == pivot_calc::kPtgInt) {
      std::string text = std::to_string(token.integer);
      RETURN_IF_ERROR(bytes.consume(text.size()));
      stack.push_back(std::move(text));
      continue;
    }
    const std::string_view op = BinaryOperatorText(token.ptg);
    if (!op.empty()) {
      if (stack.size() < 2U) {
        return Corrupt("calculated item formula operator lacks its operands");
      }
      std::string rhs = std::move(stack.back());
      stack.pop_back();
      RETURN_IF_ERROR(bytes.consume(op.size() + rhs.size()));
      stack.back().append(op).append(rhs);
      continue;
    }
    if (token.ptg != kPtgUminus && token.ptg != kPtgParen) {
      return Unmeasured("calculated item formula operator");
    }
    if (stack.empty()) {
      return Corrupt("calculated item formula operator lacks its operand");
    }
    // A prefix needs a fresh string; its full length is charged so nesting cannot copy without bound.
    std::string& top = stack.back();
    std::string out;
    out.reserve(top.size() + 2U);
    out.push_back(token.ptg == kPtgUminus ? '-' : '(');
    out.append(top);
    if (token.ptg == kPtgParen) {
      out.push_back(')');
    }
    RETURN_IF_ERROR(bytes.consume(out.size()));
    top = std::move(out);
  }
  if (stack.size() != 1U) {
    return Corrupt("calculated item formula does not reduce to one expression");
  }
  return std::move(stack.front());
}

/// Compiles calculated-item formula text to rgce and its PNames (cache field, shared item) pairs. The grammar
/// is the operators Excel's saved items use, by precedence: `>` < `&` < `+ -` < `* /` < `^` < unary `-`, over
/// `Field[Item]` names, integers and parentheses. Only parentheses recurse, at most kMaxGroupingParenDepth deep.
class FormulaCompiler {
 public:
  FormulaCompiler(std::string_view text, const pivot::PivotCache& cache) : text_(text), cache_(cache) {}

  bool Compile() { return Comparison() && pos_ == text_.size(); }

  const std::vector<std::uint8_t>& rgce() const { return rgce_; }
  const std::vector<std::pair<std::uint32_t, std::uint32_t>>& names() const { return names_; }

 private:
  bool At(char c) const { return pos_ < text_.size() && text_[pos_] == c; }

  using Rule = bool (FormulaCompiler::*)();

  /// One left-associative level: `next` operands joined by the operators `ops` lists, with their ptgs.
  bool Level(Rule next, std::string_view ops, const std::uint8_t* ptgs) {
    if (!(this->*next)()) {
      return false;
    }
    while (pos_ < text_.size()) {
      const std::size_t which = ops.find(text_[pos_]);
      if (which == std::string_view::npos) {
        return true;
      }
      ++pos_;
      if (!(this->*next)()) {
        return false;
      }
      rgce_.push_back(ptgs[which]);
    }
    return true;
  }

  bool Comparison() {
    static constexpr std::uint8_t kPtgs[] = {kPtgGt};
    return Level(&FormulaCompiler::Concat, ">", kPtgs);
  }
  bool Concat() {
    static constexpr std::uint8_t kPtgs[] = {kPtgConcat};
    return Level(&FormulaCompiler::Additive, "&", kPtgs);
  }
  bool Additive() {
    static constexpr std::uint8_t kPtgs[] = {kPtgAdd, kPtgSub};
    return Level(&FormulaCompiler::Multiplicative, "+-", kPtgs);
  }
  bool Multiplicative() {
    static constexpr std::uint8_t kPtgs[] = {kPtgMul, kPtgDiv};
    return Level(&FormulaCompiler::Power, "*/", kPtgs);
  }
  bool Power() {
    static constexpr std::uint8_t kPtgs[] = {kPtgPower};
    return Level(&FormulaCompiler::Unary, "^", kPtgs);
  }

  /// A run of unary minus signs, read as a loop, then its operand.
  bool Unary() {
    std::size_t signs = 0;
    while (At('-')) {
      ++pos_;
      if (++signs > parser::kMaxOperatorNesting) {
        return false;
      }
    }
    if (!Primary()) {
      return false;
    }
    rgce_.insert(rgce_.end(), signs, kPtgUminus);
    return true;
  }

  bool Primary() {
    if (At('(')) {
      if (++paren_depth_ > parser::kMaxGroupingParenDepth) {
        return false;
      }
      ++pos_;
      if (!Comparison() || !At(')')) {
        return false;
      }
      ++pos_;
      --paren_depth_;
      rgce_.push_back(kPtgParen);
      return true;
    }
    if (pos_ < text_.size() && text_[pos_] >= '0' && text_[pos_] <= '9') {
      return Integer();
    }
    return Name();
  }

  /// A PtgInt literal: decimal digits naming 0..65535 (the text check after compilation rejects leading zeros).
  bool Integer() {
    std::uint32_t value = 0;
    while (pos_ < text_.size() && text_[pos_] >= '0' && text_[pos_] <= '9') {
      value = value * 10U + static_cast<std::uint32_t>(text_[pos_] - '0');
      if (value > 0xFFFFU) {
        return false;
      }
      ++pos_;
    }
    rgce_.push_back(pivot_calc::kPtgInt);
    emit_u16(rgce_, static_cast<std::uint16_t>(value));
    return true;
  }

  bool Name() {
    const std::size_t open = text_.find('[', pos_);
    if (open == std::string_view::npos) {
      return false;
    }
    const std::size_t close = text_.find(']', open);
    if (close == std::string_view::npos) {
      return false;
    }
    const std::string_view field_text = text_.substr(pos_, open - pos_);
    const std::string_view item_text = text_.substr(open + 1U, close - open - 1U);
    if (!IsPlainField(field_text) || !IsPlainItem(item_text)) {
      return false;
    }
    const std::optional<std::uint32_t> field = UniqueField(field_text);
    if (!field) {
      return false;
    }
    const std::optional<std::uint32_t> item = UniqueItem(*field, item_text);
    if (!item) {
      return false;
    }
    pos_ = close + 1U;
    const std::pair<std::uint32_t, std::uint32_t> pair{*field, *item};
    std::uint32_t index = 0;
    while (index < names_.size() && names_[index] != pair) {
      ++index;
    }
    if (index == names_.size()) {
      names_.push_back(pair);
    }
    rgce_.push_back(pivot_calc::kPtgSxName);
    rgce_.push_back(pivot_calc::kEptgSxName);
    emit_u32(rgce_, index);
    return true;
  }

  std::optional<std::uint32_t> UniqueField(std::string_view name) const {
    std::optional<std::uint32_t> found;
    for (std::size_t i = 0; i < cache_.fields().size(); ++i) {
      if (cache_.fields()[i].name == name) {
        if (found) {
          return std::nullopt;
        }
        found = static_cast<std::uint32_t>(i);
      }
    }
    return found;
  }

  std::optional<std::uint32_t> UniqueItem(std::uint32_t field, std::string_view text) const {
    const std::vector<Value>& items = cache_.fields()[field].shared_items;
    std::optional<std::uint32_t> found;
    for (std::size_t i = 0; i < items.size(); ++i) {
      if (items[i].is_text() && items[i].as_text() == text) {
        if (found) {
          return std::nullopt;
        }
        found = static_cast<std::uint32_t>(i);
      }
    }
    return found;
  }

  std::string_view text_;
  const pivot::PivotCache& cache_;
  std::size_t pos_ = 0;
  std::uint32_t paren_depth_ = 0;
  std::vector<std::uint8_t> rgce_;
  std::vector<std::pair<std::uint32_t, std::uint32_t>> names_;
};

// ---------------------------------------------------------------------------
// Read.

void ReadHintFlag(std::uint16_t flags, std::uint16_t bit, bool schema_default, bool* value, bool* has) {
  const bool set = (flags & bit) != 0U;
  if (set != schema_default) {
    *value = set;
    *has = true;
  }
}

/// `<sharedItems>` attributes as the xlsx reader keeps them: only the ones that differ from the schema default
/// are present, which is how Excel writes them.
Expected<pivot::SharedItemsHints, Error> ReadHints(const pcd::ItemTable& table, bool date1904) {
  pivot::SharedItemsHints h;
  const std::uint16_t f = table.flags;
  ReadHintFlag(f, kAtblSemiMixed, true, &h.contains_semi_mixed, &h.has_contains_semi_mixed);
  ReadHintFlag(f, pcd::kAtblNonDate, true, &h.contains_non_date, &h.has_contains_non_date);
  ReadHintFlag(f, pcd::kAtblDate, false, &h.contains_date, &h.has_contains_date);
  ReadHintFlag(f, kAtblString, true, &h.contains_string, &h.has_contains_string);
  ReadHintFlag(f, pcd::kAtblBlank, false, &h.contains_blank, &h.has_contains_blank);
  ReadHintFlag(f, pcd::kAtblMixedTypes, false, &h.contains_mixed_types, &h.has_contains_mixed_types);
  ReadHintFlag(f, pcd::kAtblNumber, false, &h.contains_number, &h.has_contains_number);
  ReadHintFlag(f, pcd::kAtblInteger, false, &h.contains_integer, &h.has_contains_integer);
  if ((f & pcd::kAtblMinMax) != 0U) {
    const std::uint16_t shape = f & (pcd::kAtblDate | pcd::kAtblNumber);
    if (shape == pcd::kAtblDate) {
      ASSIGN_OR_RETURN(h.min_date, DateIso(table.min, date1904));
      ASSIGN_OR_RETURN(h.max_date, DateIso(table.max, date1904));
      h.has_min_date = true;
      h.has_max_date = true;
    } else if (shape == pcd::kAtblNumber) {
      h.min_value = NumberText(table.min);
      h.max_value = NumberText(table.max);
      h.has_min_value = true;
      h.has_max_value = true;
    } else {
      return Unmeasured("shared items range over both dates and numbers");
    }
  }
  h.present = h.has_contains_number || h.has_contains_integer || h.has_contains_date || h.has_contains_string ||
              h.has_contains_blank || h.has_contains_semi_mixed || h.has_contains_non_date ||
              h.has_contains_mixed_types || h.has_min_value || h.has_min_date;
  return h;
}

void PushItem(pivot::PivotCacheField& field, Value value, bool is_date, std::uint8_t attributes) {
  field.shared_items.push_back(value);
  field.shared_item_is_date.push_back(is_date);
  field.shared_item_is_calculated.push_back((attributes & pcd::kItemFormula) != 0U);
  field.shared_item_is_missing.push_back((attributes & pcd::kItemUnused) != 0U);
}

Expected<void, Error> ReadItems(const pcd::ItemTable& table, pivot::PivotCache& cache, pivot::PivotCacheField& field,
                                bool date1904) {
  // A row of plain items of one type the writer would group differently (single records reaching the run length,
  // or a run beside more of its type) has no Excel-saved sample.
  enum class Plain : std::uint8_t { kNone, kString, kNumber, kDate };
  Plain last = Plain::kNone;
  std::size_t last_count = 0;
  for (const pcd::ItemEntry& entry : table.entries) {
    Plain kind = Plain::kNone;
    if (entry.is_run) {
      kind = entry.run.type == pcd::RunType::kStrings   ? Plain::kString
             : entry.run.type == pcd::RunType::kNumbers ? Plain::kNumber
                                                        : Plain::kDate;
    } else if (entry.item.type == pcd::ItemType::kString) {
      kind = Plain::kString;
    } else if (entry.item.type == pcd::ItemType::kNumber) {
      kind = Plain::kNumber;
    } else if (entry.item.type == pcd::ItemType::kDate) {
      kind = Plain::kDate;
    }
    const std::size_t count = entry.is_run ? entry.run.size() : 1U;
    last_count = kind != Plain::kNone && kind == last ? last_count + count : count;
    last = kind;
    if (kind != Plain::kNone && last_count >= kMinRunLength && (!entry.is_run || last_count != count)) {
      return Unmeasured("shared items grouped other than Excel groups them");
    }
    if (!entry.is_run) {
      const pcd::Item& item = entry.item;
      switch (item.type) {
        case pcd::ItemType::kString:
          PushItem(field, Value::text(StoreText(cache, item.text)), false, 0U);
          break;
        case pcd::ItemType::kNumber:
          PushItem(field, Value::number(item.number), false, 0U);
          break;
        case pcd::ItemType::kAttributedString:
          if (item.attributes == 0U) {
            return Unmeasured("attributed string item without an attribute");
          }
          PushItem(field, Value::text(StoreText(cache, item.text)), false, item.attributes);
          break;
        case pcd::ItemType::kAttributedMissing:
          if (item.attributes == 0U) {
            return Unmeasured("attributed missing item without an attribute");
          }
          PushItem(field, Value::blank(), false, item.attributes);
          break;
        case pcd::ItemType::kMissing:
          PushItem(field, Value::blank(), false, 0U);
          break;
        case pcd::ItemType::kBoolean:
          PushItem(field, Value::boolean(item.boolean), false, 0U);
          break;
        case pcd::ItemType::kError: {
          const ErrorCode code = error_from_ooxml_code(item.error);
          if (ooxml_code(code) != item.error) {
            return Unmeasured("shared error item code");
          }
          PushItem(field, Value::error(code), false, 0U);
          break;
        }
        case pcd::ItemType::kDate: {
          ASSIGN_OR_RETURN(const double serial, RunDateSerial(item.date, date1904));
          PushItem(field, Value::number(serial), true, 0U);
          break;
        }
      }
      continue;
    }
    const pcd::ItemRun& run = entry.run;
    if (run.size() < kMinRunLength) {
      return Unmeasured("shared item run shorter than any Excel writes");
    }
    switch (run.type) {
      case pcd::RunType::kNumbers:
        for (const double number : run.numbers) {
          PushItem(field, Value::number(number), false, 0U);
        }
        break;
      case pcd::RunType::kStrings:
        for (const std::string& text : run.strings) {
          PushItem(field, Value::text(StoreText(cache, text)), false, 0U);
        }
        break;
      case pcd::RunType::kDates:
        for (const pcd::Date& date : run.dates) {
          ASSIGN_OR_RETURN(const double serial, RunDateSerial(date, date1904));
          PushItem(field, Value::number(serial), true, 0U);
        }
        break;
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<pivot::FieldGroup, Error> ReadGroup(const std::vector<pivot_records::RawRecord>& records,
                                             pivot::PivotCache& cache, bool date1904) {
  std::size_t pos = 0;
  ASSIGN_OR_RETURN(const pivot_group::FieldGroup group, pivot_group::decode_field_group(records, &pos));
  if (pos != records.size()) {
    return Corrupt("field group is followed by records outside its block");
  }
  pivot::FieldGroup out;
  if (group.par != pivot_group::kNoField) {
    out.par = group.par;
  }
  if (group.base != pivot_group::kNoField) {
    out.base = group.base;
  }
  if (group.range) {
    const pivot_group::RangeGroup& r = *group.range;
    constexpr std::uint8_t kKnown =
        pivot_group::kRangeAutoStart | pivot_group::kRangeAutoEnd | pivot_group::kRangeDates;
    if (r.group_by != pivot_group::kGroupByDays || (r.flags & ~kKnown) != 0U ||
        (r.flags & pivot_group::kRangeDates) == 0U) {
      return Unmeasured("range group other than a days group over dates");
    }
    out.kind = pivot::GroupKind::Range;
    out.range.by = pivot::GroupBy::Days;
    out.range.is_date = true;
    ASSIGN_OR_RETURN(out.range.start_date, DateIso(r.start, date1904));
    ASSIGN_OR_RETURN(out.range.end_date, DateIso(r.end, date1904));
    out.range.auto_start = (r.flags & pivot_group::kRangeAutoStart) != 0U;
    out.range.auto_end = (r.flags & pivot_group::kRangeAutoEnd) != 0U;
    out.range.interval = r.interval;
  }
  if (group.items) {
    for (const std::string& text : *group.items) {
      out.group_items.push_back(Value::text(StoreText(cache, text)));
      out.group_item_is_date.push_back(false);
    }
  }
  return out;
}

Expected<pivot::PivotCacheField, Error> ReadField(const pcd::Field& in, pivot::PivotCache& cache, bool date1904) {
  pivot::PivotCacheField field;
  field.name = in.name;
  field.is_database_field = (in.flags & pcd::kFieldDatabase) != 0U;
  field.num_fmt_id = in.num_fmt_id;
  if (in.items) {
    ASSIGN_OR_RETURN(field.shared_items_hints, ReadHints(*in.items, date1904));
    RETURN_IF_ERROR(ReadItems(*in.items, cache, field, date1904));
  }
  if (!in.grouping.empty()) {
    ASSIGN_OR_RETURN(field.field_group, ReadGroup(in.grouping, cache, date1904));
  }
  return field;
}

Expected<pivot::PivotArea, Error> ReadArea(const pivot_area::PivotArea& in, const pivot::PivotCache& cache) {
  constexpr std::uint32_t kKnown = pivot_area::kAreaDataOnly | pivot_area::kAreaCacheIndex;
  if (in.field != pivot_area::kNoField || in.type != pivot_area::kAreaTypeNormal || (in.flags & ~kKnown) != 0U) {
    return Unmeasured("calculated item pivot area shape");
  }
  pivot::PivotArea area;
  if ((in.flags & pivot_area::kAreaCacheIndex) != 0U) {
    area.cache_index = true;
  }
  // Both Excel-saved pivot rules carry outline="0" fieldPosition="0" with no bit for either.
  area.outline = false;
  area.field_position = 0U;
  if ((in.flags & pivot_area::kAreaDataOnly) == 0U) {
    area.extra_attrs.emplace_back("dataOnly", "0");
  }
  if (in.filters) {
    if (in.filters->empty()) {
      return Unmeasured("pivot area with an empty references list");
    }
    for (const pivot_area::PivotAreaFilter& filter : *in.filters) {
      if ((filter.flags & kFilterAlways) == 0U ||
          (filter.flags & ~(kFilterAlways | pivot_area::kFilterSelected)) != 0U) {
        return Unmeasured("pivot area reference flags");
      }
      if (filter.field >= cache.fields().size()) {
        return Corrupt("calculated item pivot area names a field past the cache");
      }
      for (const std::uint32_t index : filter.items) {
        if (index >= cache.fields()[filter.field].shared_items.size()) {
          return Corrupt("calculated item pivot area names an item past its field");
        }
      }
      pivot::PivotAreaReference ref;
      ref.field = filter.field;
      ref.items = filter.items;
      if ((filter.flags & pivot_area::kFilterSelected) == 0U) {
        ref.extra_attrs.emplace_back("selected", "0");
      }
      area.references.push_back(std::move(ref));
    }
  }
  return area;
}

Expected<pivot::PivotCalculatedItem, Error> ReadCalcItem(const pivot_calc::CalcItem& in,
                                                         const pivot::PivotCache& cache) {
  if (in.field != pivot_calc::kNoField || !in.rgcb.empty()) {
    return Unmeasured("calculated item with a field index or formula extra data");
  }
  ASSIGN_OR_RETURN(const std::vector<pivot_calc::FormulaToken> tokens, pivot_calc::walk_calc_item_formula(in.rgce));
  const std::vector<pivot_calc::PName> empty;
  const std::vector<pivot_calc::PName>& names = in.names ? *in.names : empty;
  std::vector<std::string> texts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    const pivot_calc::PName& name = names[i];
    if (name.field != pivot_calc::kNoField || name.flags != kPNameFlags || name.pairs.size() != 1U ||
        name.pairs.front().flags != kPNPairFlags) {
      return Unmeasured("calculated item name shape");
    }
    const pivot_calc::PNPair& pair = name.pairs.front();
    if (pair.field >= cache.fields().size() || pair.item >= cache.fields()[pair.field].shared_items.size()) {
      return Corrupt("calculated item name points past the cache");
    }
    for (std::size_t j = 0; j < i; ++j) {
      if (names[j].pairs.front().field == pair.field && names[j].pairs.front().item == pair.item) {
        return Unmeasured("calculated item names table with a repeated entry");
      }
    }
    const pivot::PivotCacheField& field = cache.fields()[pair.field];
    const Value& item = field.shared_items[pair.item];
    if (!item.is_text() || !IsPlainField(field.name) || !IsPlainItem(item.as_text())) {
      return Unmeasured("calculated item name that is not a plain Field[Item]");
    }
    std::string text = field.name;
    text.append("[").append(item.as_text()).append("]");
    texts.push_back(std::move(text));
  }
  // The names table lists each name at its first use.
  std::size_t next_new = 0;
  for (const pivot_calc::FormulaToken& token : tokens) {
    if (token.ptg == pivot_calc::kPtgSxName && token.sx_name >= next_new) {
      if (token.sx_name != next_new) {
        return Unmeasured("calculated item names table out of first-use order");
      }
      ++next_new;
    }
  }
  if (next_new != names.size()) {
    return Unmeasured("calculated item names table with an unused entry");
  }
  pivot::PivotCalculatedItem item;
  ASSIGN_OR_RETURN(item.formula, RenderFormula(tokens, texts));
  item.has_area = in.area.has_value();
  if (in.area) {
    ASSIGN_OR_RETURN(item.area, ReadArea(*in.area, cache));
    if (!item.area.references.empty() && item.area.references.front().field.has_value()) {
      item.cache_field_index = *item.area.references.front().field;
    }
  }
  return item;
}

void ReadRoot(const pcd::Document& doc, pivot::PivotCache& cache) {
  std::vector<std::pair<std::string, std::string>>& attrs = cache.mutable_passthrough_attrs();
  if (doc.uid) {
    attrs.emplace_back("xmlns:mc", std::string(kNsMarkupCompat));
    attrs.emplace_back("xmlns:xr", std::string(kNsXr));
    attrs.emplace_back("mc:Ignorable", std::string(kXrPrefix));
  }
  attrs.emplace_back("refreshedBy", doc.refreshed_by);
  attrs.emplace_back("refreshedDate", RefreshedDateText(doc.refreshed_date));
  attrs.emplace_back("createdVersion", std::to_string(doc.created_version));
  attrs.emplace_back("refreshedVersion", std::to_string(doc.refreshed_version));
  attrs.emplace_back("minRefreshableVersion", std::to_string(doc.min_refreshable_version));
  if (doc.uid) {
    attrs.emplace_back("xr:uid", FormatGuid(doc.uid->data()));
  }
  cache.set_save_data(true);
  if (doc.missing_items_limit != 0xFFFFFFFFU) {
    cache.set_missing_items_limit(doc.missing_items_limit);
  }
  if ((doc.flags & pcd::kCacheRefreshOnLoad) != 0U) {
    cache.set_refresh_on_load(true);
  }
  cache.mutable_cache_source_attrs().emplace_back("type", "worksheet");
  if (doc.source.connection_id != 0U) {
    cache.mutable_cache_source_attrs().emplace_back("connectionId", std::to_string(doc.source.connection_id));
  }
  pivot::WorksheetSource& source = cache.mutable_worksheet_source();
  source.present = true;
  source.ref = format_sqref_text({doc.source.ref});
  source.sheet = doc.source.sheet;
}

Expected<void, Error> ReadExtLst(const pcd::Document& doc, pivot::PivotCache& cache) {
  if (!doc.has_pcd14 && !doc.data_source_clean) {
    return Expected<void, Error>::Ok();
  }
  std::string& xml = cache.mutable_trailing_children_after_calculated_xml();
  xml.append(kExtLstOpen);
  if (doc.has_pcd14) {
    xml.append(kPcd14Ext);
  }
  if (doc.data_source_clean) {
    if (*doc.data_source_clean != kDataSourceCleanTrue) {
      return Unmeasured("dataSourceClean other than 1");
    }
    xml.append(kDataSourceCleanExt);
  }
  xml.append(kExtLstClose);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadRecords(const cache_records::PivotCacheRecordsDoc& doc, pivot::PivotCache& cache,
                                  bool date1904) {
  const std::size_t field_count = cache.fields().size();
  std::vector<std::size_t> db_positions;
  for (std::size_t i = 0; i < field_count; ++i) {
    if (cache.fields()[i].is_database_field) {
      db_positions.push_back(i);
    }
  }
  ResourceBudget slots(kMaxPivotCacheRecordCells, "pivot_cache_record_cells");
  std::vector<pivot::PivotCacheRecord>& records = cache.mutable_records();
  for (const std::vector<RecordCell>& row : doc.rows) {
    RETURN_IF_ERROR(slots.consume(1U + field_count));
    if (row.size() != db_positions.size()) {
      return Corrupt("record width does not match the database fields");
    }
    pivot::PivotCacheRecord record;
    record.cells.assign(field_count, Value::blank());
    record.cell_is_index.assign(field_count, false);
    for (std::size_t j = 0; j < row.size(); ++j) {
      const RecordCell& cell = row[j];
      const std::size_t pos = db_positions[j];
      switch (cell.kind) {
        case CellKind::kSharedIndex:
          if (cell.index >= cache.fields()[pos].shared_items.size()) {
            return Corrupt("record indexes a shared item the field does not have");
          }
          record.cells[pos] = Value::number(static_cast<double>(cell.index));
          record.cell_is_index[pos] = true;
          break;
        case CellKind::kNumber:
          record.cells[pos] = Value::number(cell.number);
          break;
        case CellKind::kText:
          record.cells[pos] = Value::text(StoreText(cache, cell.text));
          break;
        case CellKind::kDate: {
          ASSIGN_OR_RETURN(const double serial, record_date_serial(cell.date, date1904));
          record.cells[pos] = Value::number(serial);
          record.cell_is_date.resize(field_count, false);
          record.cell_is_date[pos] = true;
          break;
        }
      }
    }
    records.push_back(std::move(record));
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// Write. Every helper returns false for a model state the records cannot hold.

std::uint16_t HintBit(bool has, bool value, bool schema_default, std::uint16_t bit) {
  return (has ? value : schema_default) ? bit : 0U;
}

bool WriteHints(const pivot::SharedItemsHints& h, bool date1904, pcd::ItemTable* table) {
  if (!h.extra_attrs.empty()) {
    return false;
  }
  std::uint16_t f = pcd::kAtblAlways;
  f |= HintBit(h.has_contains_semi_mixed, h.contains_semi_mixed, true, kAtblSemiMixed);
  f |= HintBit(h.has_contains_non_date, h.contains_non_date, true, pcd::kAtblNonDate);
  f |= HintBit(h.has_contains_date, h.contains_date, false, pcd::kAtblDate);
  f |= HintBit(h.has_contains_string, h.contains_string, true, kAtblString);
  f |= HintBit(h.has_contains_blank, h.contains_blank, false, pcd::kAtblBlank);
  f |= HintBit(h.has_contains_mixed_types, h.contains_mixed_types, false, pcd::kAtblMixedTypes);
  f |= HintBit(h.has_contains_number, h.contains_number, false, pcd::kAtblNumber);
  f |= HintBit(h.has_contains_integer, h.contains_integer, false, pcd::kAtblInteger);
  const bool values = h.has_min_value || h.has_max_value;
  const bool dates = h.has_min_date || h.has_max_date;
  const std::uint16_t shape = f & (pcd::kAtblDate | pcd::kAtblNumber);
  if (values) {
    if (dates || !h.has_min_value || !h.has_max_value || shape != pcd::kAtblNumber ||
        !parse_xsd_double(h.min_value, &table->min) || !parse_xsd_double(h.max_value, &table->max) ||
        NumberText(table->min) != h.min_value || NumberText(table->max) != h.max_value) {
      return false;
    }
    f |= pcd::kAtblMinMax;
  } else if (dates) {
    if (!h.has_min_date || !h.has_max_date || shape != pcd::kAtblDate) {
      return false;
    }
    const std::optional<double> min = IsoSerial(h.min_date, date1904);
    const std::optional<double> max = IsoSerial(h.max_date, date1904);
    if (!min || !max) {
      return false;
    }
    table->min = *min;
    table->max = *max;
    f |= pcd::kAtblMinMax;
  }
  table->flags = f;
  return true;
}

/// One shared item's record shape.
enum class ItemClass : std::uint8_t {
  kString,
  kNumber,
  kDate,
  kAttributedString,
  kAttributedMissing,
  kMissing,
  kBoolean,
  kError
};

bool ClassifyItem(const pivot::PivotCacheField& field, std::size_t i, ItemClass* out, std::uint8_t* attributes) {
  const Value& v = field.shared_items[i];
  const bool is_date = FlagAt(field.shared_item_is_date, i);
  *attributes = static_cast<std::uint8_t>((FlagAt(field.shared_item_is_missing, i) ? pcd::kItemUnused : 0U) |
                                          (FlagAt(field.shared_item_is_calculated, i) ? pcd::kItemFormula : 0U));
  if (v.is_text() && !is_date) {
    *out = *attributes != 0U ? ItemClass::kAttributedString : ItemClass::kString;
    return true;
  }
  if (v.is_blank() && !is_date) {
    *out = *attributes != 0U ? ItemClass::kAttributedMissing : ItemClass::kMissing;
    return true;
  }
  // Only a string or a blank carries the attribute block.
  if (*attributes != 0U) {
    return false;
  }
  if (v.is_number()) {
    *out = is_date ? ItemClass::kDate : ItemClass::kNumber;
    return true;
  }
  if (is_date) {
    return false;
  }
  if (v.is_boolean()) {
    *out = ItemClass::kBoolean;
    return true;
  }
  if (v.is_error()) {
    const std::int32_t code = ooxml_code(v.as_error());
    *out = ItemClass::kError;
    return code >= 0 && code <= 0xFF;
  }
  return false;
}

bool WriteItems(const pivot::PivotCacheField& field, bool date1904, pcd::ItemTable* table) {
  const std::size_t n = field.shared_items.size();
  std::size_t i = 0;
  while (i < n) {
    ItemClass cls = ItemClass::kString;
    std::uint8_t attributes = 0;
    if (!ClassifyItem(field, i, &cls, &attributes)) {
      return false;
    }
    if (cls != ItemClass::kString && cls != ItemClass::kNumber && cls != ItemClass::kDate) {
      const Value& v = field.shared_items[i];
      pcd::ItemEntry entry;
      entry.item.attributes = attributes;
      switch (cls) {
        case ItemClass::kAttributedString:
          entry.item.type = pcd::ItemType::kAttributedString;
          entry.item.text = std::string(v.as_text());
          break;
        case ItemClass::kAttributedMissing:
          entry.item.type = pcd::ItemType::kAttributedMissing;
          break;
        case ItemClass::kMissing:
          entry.item.type = pcd::ItemType::kMissing;
          break;
        case ItemClass::kBoolean:
          entry.item.type = pcd::ItemType::kBoolean;
          entry.item.boolean = v.as_boolean();
          break;
        default:
          entry.item.type = pcd::ItemType::kError;
          entry.item.error = static_cast<std::uint8_t>(ooxml_code(v.as_error()));
          break;
      }
      table->entries.push_back(std::move(entry));
      ++i;
      continue;
    }
    std::size_t end = i + 1U;
    while (end < n) {
      ItemClass next = ItemClass::kString;
      std::uint8_t next_attributes = 0;
      if (!ClassifyItem(field, end, &next, &next_attributes) || next != cls) {
        break;
      }
      ++end;
    }
    if (end - i >= kMinRunLength) {
      pcd::ItemEntry entry;
      entry.is_run = true;
      for (std::size_t k = i; k < end; ++k) {
        const Value& v = field.shared_items[k];
        if (cls == ItemClass::kString) {
          entry.run.type = pcd::RunType::kStrings;
          entry.run.strings.emplace_back(v.as_text());
        } else if (cls == ItemClass::kNumber) {
          entry.run.type = pcd::RunType::kNumbers;
          entry.run.numbers.push_back(v.as_number());
        } else {
          const std::optional<pcd::Date> date = RunDate(v.as_number(), date1904);
          if (!date) {
            return false;
          }
          entry.run.type = pcd::RunType::kDates;
          entry.run.dates.push_back(*date);
        }
      }
      table->entries.push_back(std::move(entry));
    } else {
      for (std::size_t k = i; k < end; ++k) {
        pcd::ItemEntry entry;
        if (cls == ItemClass::kString) {
          entry.item.type = pcd::ItemType::kString;
          entry.item.text = std::string(field.shared_items[k].as_text());
        } else if (cls == ItemClass::kNumber) {
          entry.item.type = pcd::ItemType::kNumber;
          entry.item.number = field.shared_items[k].as_number();
        } else {
          const std::optional<pcd::Date> date = RunDate(field.shared_items[k].as_number(), date1904);
          if (!date) {
            return false;
          }
          entry.item.type = pcd::ItemType::kDate;
          entry.item.date = *date;
        }
        table->entries.push_back(std::move(entry));
      }
    }
    i = end;
  }
  return true;
}

bool WriteGroup(const pivot::FieldGroup& in, bool date1904, std::vector<pivot_records::RawRecord>* out) {
  if (!in.extra_attrs.empty() || in.kind == pivot::GroupKind::Discrete) {
    return false;
  }
  pivot_group::FieldGroup group;
  group.par = in.par.value_or(pivot_group::kNoField);
  group.base = in.base.value_or(pivot_group::kNoField);
  if (in.kind == pivot::GroupKind::Range) {
    const pivot::RangeGroup& r = in.range;
    if (r.by != pivot::GroupBy::Days || !r.is_date) {
      return false;
    }
    const std::optional<double> start = IsoSerial(r.start_date, date1904);
    const std::optional<double> end = IsoSerial(r.end_date, date1904);
    if (!start || !end) {
      return false;
    }
    pivot_group::RangeGroup range;
    range.group_by = pivot_group::kGroupByDays;
    range.flags =
        static_cast<std::uint8_t>(pivot_group::kRangeDates | (r.auto_start ? pivot_group::kRangeAutoStart : 0U) |
                                  (r.auto_end ? pivot_group::kRangeAutoEnd : 0U));
    range.start = *start;
    range.end = *end;
    range.interval = r.interval;
    group.range = range;
  }
  if (!in.group_items.empty()) {
    std::vector<std::string> items;
    for (std::size_t i = 0; i < in.group_items.size(); ++i) {
      if (!in.group_items[i].is_text() || FlagAt(in.group_item_is_date, i)) {
        return false;
      }
      items.emplace_back(in.group_items[i].as_text());
    }
    group.items = std::move(items);
  }
  pivot_group::encode_field_group(group, out);
  return true;
}

bool WriteField(const pivot::PivotCacheField& in, bool date1904, pcd::Field* out) {
  if (in.formula || !in.extra_attrs.empty() || !in.trailing_children_xml.empty() || !in.num_fmt_id ||
      *in.num_fmt_id > 0xFFFFU) {
    return false;
  }
  out->flags = in.is_database_field ? pcd::kFieldDatabase : 0U;
  out->num_fmt_id = static_cast<std::uint16_t>(*in.num_fmt_id);
  out->name = in.name;
  if (in.is_database_field || !in.shared_items.empty() || in.shared_items_hints.present) {
    pcd::ItemTable table;
    if (!WriteHints(in.shared_items_hints, date1904, &table) || !WriteItems(in, date1904, &table)) {
      return false;
    }
    out->items = std::move(table);
  }
  if (in.field_group && !WriteGroup(*in.field_group, date1904, &out->grouping)) {
    return false;
  }
  return true;
}

bool WriteArea(const pivot::PivotArea& in, pivot_area::PivotArea* out) {
  const bool measured_shape = in.outline.has_value() && !*in.outline && in.field_position.has_value() &&
                              *in.field_position == 0U && !(in.cache_index.has_value() && !*in.cache_index);
  if (!in.extra_children_xml.empty() || !measured_shape) {
    return false;
  }
  std::uint32_t flags = pivot_area::kAreaDataOnly;
  for (const auto& [name, value] : in.extra_attrs) {
    if (name != "dataOnly" || value != "0" || (flags & pivot_area::kAreaDataOnly) == 0U) {
      return false;
    }
    flags &= ~pivot_area::kAreaDataOnly;
  }
  if (in.cache_index) {
    flags |= pivot_area::kAreaCacheIndex;
  }
  out->field = pivot_area::kNoField;
  out->type = pivot_area::kAreaTypeNormal;
  out->flags = flags;
  if (in.references.empty()) {
    return true;
  }
  std::vector<pivot_area::PivotAreaFilter> filters;
  for (const pivot::PivotAreaReference& ref : in.references) {
    if (!ref.field || !ref.extra_children_xml.empty()) {
      return false;
    }
    pivot_area::PivotAreaFilter filter;
    filter.field = *ref.field;
    filter.flags = kFilterAlways | pivot_area::kFilterSelected;
    for (const auto& [name, value] : ref.extra_attrs) {
      if (name != "selected" || value != "0" || (filter.flags & pivot_area::kFilterSelected) == 0U) {
        return false;
      }
      filter.flags &= ~pivot_area::kFilterSelected;
    }
    filter.items = ref.items;
    filters.push_back(std::move(filter));
  }
  out->filters = std::move(filters);
  return true;
}

bool WriteCalcItem(const pivot::PivotCalculatedItem& in, const pivot::PivotCache& cache, pivot_calc::CalcItem* out) {
  if (in.has_field_attr || !in.extra_children_xml.empty()) {
    return false;
  }
  FormulaCompiler compiler(in.formula, cache);
  if (!compiler.Compile()) {
    return false;
  }
  std::vector<pivot_calc::PName> names;
  std::vector<std::string> texts;
  for (const auto& [field, item] : compiler.names()) {
    pivot_calc::PName name;
    name.field = pivot_calc::kNoField;
    name.flags = kPNameFlags;
    name.pairs.push_back(pivot_calc::PNPair{field, kPNPairFlags, item});
    names.push_back(std::move(name));
    std::string text = cache.fields()[field].name;
    text.append("[").append(cache.fields()[field].shared_items[item].as_text()).append("]");
    texts.push_back(std::move(text));
  }
  // The read side must give the formula back as stored.
  auto tokens = pivot_calc::walk_calc_item_formula(compiler.rgce());
  if (!tokens) {
    return false;
  }
  auto rendered = RenderFormula(tokens.value(), texts);
  if (!rendered || rendered.value() != in.formula) {
    return false;
  }
  out->field = pivot_calc::kNoField;
  out->rgce = compiler.rgce();
  if (!names.empty()) {
    out->names = std::move(names);
  }
  if (in.has_area) {
    pivot_area::PivotArea area;
    if (!WriteArea(in.area, &area)) {
      return false;
    }
    out->area = std::move(area);
  }
  return true;
}

bool WriteRoot(const pivot::PivotCache& cache, pcd::Document* doc) {
  enum : unsigned {
    kBy = 1U,
    kDate = 2U,
    kCreated = 4U,
    kRefreshed = 8U,
    kMinRefreshable = 16U,
    kRequired = 31U,
    kNsMc = 32U,
    kNsXrDecl = 64U,
    kIgnorable = 128U,
    kUidBookkeeping = 224U,
  };
  unsigned seen = 0;
  const auto take = [&seen](unsigned bit) {
    if ((seen & bit) != 0U) {
      return false;
    }
    seen |= bit;
    return true;
  };
  const auto version = [](std::string_view text, std::uint8_t* out) {
    const std::optional<std::uint32_t> v = ParseUnsigned(text, 0xFFU);
    if (v) {
      *out = static_cast<std::uint8_t>(*v);
    }
    return v.has_value();
  };
  for (const auto& [name, value] : cache.passthrough_attrs()) {
    bool ok = false;
    if (name == "xmlns:mc") {
      ok = take(kNsMc) && value == kNsMarkupCompat;
    } else if (name == "xmlns:xr") {
      ok = take(kNsXrDecl) && value == kNsXr;
    } else if (name == "mc:Ignorable") {
      ok = take(kIgnorable) && value == kXrPrefix;
    } else if (name == "refreshedBy") {
      ok = take(kBy);
      doc->refreshed_by = value;
    } else if (name == "refreshedDate") {
      ok = take(kDate) && parse_xsd_double(value, &doc->refreshed_date) &&
           RefreshedDateText(doc->refreshed_date) == value;
    } else if (name == "createdVersion") {
      ok = take(kCreated) && version(value, &doc->created_version);
    } else if (name == "refreshedVersion") {
      ok = take(kRefreshed) && version(value, &doc->refreshed_version);
    } else if (name == "minRefreshableVersion") {
      ok = take(kMinRefreshable) && version(value, &doc->min_refreshable_version);
    } else if (name == "xr:uid") {
      std::array<std::uint8_t, kGuidBytes> guid{};
      ok = !doc->uid.has_value() && ParseGuid(value, guid);
      doc->uid = guid;
    }
    if (!ok) {
      return false;
    }
  }
  // The namespace bookkeeping is read exactly when the uid is, so any other combination would not read back.
  const unsigned bookkeeping = seen & kUidBookkeeping;
  if ((seen & kRequired) != kRequired || bookkeeping != (doc->uid ? kUidBookkeeping : 0U) || !cache.save_data()) {
    return false;
  }
  doc->flags = static_cast<std::uint8_t>(pcd::kCacheSaveDataEnableRefresh |
                                         (cache.refresh_on_load() ? pcd::kCacheRefreshOnLoad : 0U));
  if (const std::optional<std::uint32_t> limit = cache.missing_items_limit(); limit) {
    if (*limit == 0xFFFFFFFFU) {
      return false;
    }
    doc->missing_items_limit = *limit;
  }
  for (const auto& [name, value] : cache.cache_source_attrs()) {
    if (name == "type") {
      if (value != "worksheet") {
        return false;
      }
    } else if (name == "connectionId") {
      const std::optional<std::uint32_t> id = ParseUnsigned(value, 0xFFFFFFFFU);
      if (!id) {
        return false;
      }
      doc->source.connection_id = *id;
    } else {
      return false;
    }
  }
  const pivot::WorksheetSource& source = cache.worksheet_source();
  if (!cache.cache_source_children_xml().empty() || !source.present || !source.name.empty() || source.sheet.empty() ||
      !parse_a1_range(source.ref, &doc->source.ref) || format_sqref_text({doc->source.ref}) != source.ref) {
    return false;
  }
  doc->source.sheet = source.sheet;
  return true;
}

bool WriteExtLst(const pivot::PivotCache& cache, pcd::Document* doc) {
  if (!cache.trailing_children_xml().empty()) {
    return false;
  }
  std::string_view xml = cache.trailing_children_after_calculated_xml();
  if (xml.empty()) {
    return true;
  }
  const auto consume = [&xml](std::string_view part) {
    if (xml.substr(0, part.size()) != part) {
      return false;
    }
    xml.remove_prefix(part.size());
    return true;
  };
  if (!consume(kExtLstOpen)) {
    return false;
  }
  doc->has_pcd14 = consume(kPcd14Ext);
  if (consume(kDataSourceCleanExt)) {
    doc->data_source_clean = kDataSourceCleanTrue;
  }
  return consume(kExtLstClose) && xml.empty() && (doc->has_pcd14 || doc->data_source_clean);
}

bool WriteRecordCell(const pivot::PivotCache& cache, const pivot::PivotCacheRecord& record, std::size_t pos,
                     CellKind kind, bool date1904, RecordCell* out) {
  const Value& v = record.cells[pos];
  const pivot::PivotCacheField& field = cache.fields()[pos];
  const bool is_index = pivot::cell_is_shared_index(record, field, pos);
  const bool is_date = FlagAt(record.cell_is_date, pos);
  out->kind = kind;
  switch (kind) {
    case CellKind::kSharedIndex: {
      if (!is_index || !v.is_number()) {
        return false;
      }
      const double index = v.as_number();
      if (!(index >= 0.0) || index != std::floor(index) || index >= static_cast<double>(field.shared_items.size())) {
        return false;
      }
      out->index = static_cast<std::uint32_t>(index);
      return true;
    }
    case CellKind::kNumber:
      if (is_index || is_date || !v.is_number()) {
        return false;
      }
      out->number = v.as_number();
      return true;
    case CellKind::kText:
      if (is_index || !v.is_text()) {
        return false;
      }
      out->text = std::string(v.as_text());
      return true;
    case CellKind::kDate: {
      if (is_index || !is_date || !v.is_number()) {
        return false;
      }
      auto date = record_date_from_serial(v.as_number(), date1904);
      if (!date) {
        return false;
      }
      out->date = date.value();
      return true;
    }
  }
  return false;
}

Expected<std::vector<std::uint8_t>, Error> WriteRecords(const pivot::PivotCache& cache, ByteSpan definition,
                                                        bool date1904) {
  auto kinds = cell_kinds_of_cache_definition(definition);
  if (!kinds) {
    return NotWritable();
  }
  std::vector<std::size_t> db_positions;
  for (std::size_t i = 0; i < cache.fields().size(); ++i) {
    if (cache.fields()[i].is_database_field) {
      db_positions.push_back(i);
    }
  }
  if (db_positions.size() != kinds.value().size()) {
    return NotWritable();
  }
  cache_records::PivotCacheRecordsDoc doc;
  doc.rows.reserve(cache.records().size());
  for (const pivot::PivotCacheRecord& record : cache.records()) {
    if (record.cells.size() != cache.fields().size()) {
      return NotWritable();
    }
    std::vector<RecordCell> row(db_positions.size());
    for (std::size_t j = 0; j < db_positions.size(); ++j) {
      if (!WriteRecordCell(cache, record, db_positions[j], kinds.value()[j], date1904, &row[j])) {
        return NotWritable();
      }
    }
    for (std::size_t i = 0; i < record.cells.size(); ++i) {
      if (!cache.fields()[i].is_database_field && !record.cells[i].is_blank()) {
        return NotWritable();
      }
    }
    doc.rows.push_back(std::move(row));
  }
  auto bytes = encode_pivot_cache_records(doc, kinds.value());
  if (!bytes) {
    return NotWritable();
  }
  return std::move(bytes.value());
}

}  // namespace

Expected<pivot::PivotCache, Error> pivot_cache_from_xlsb(ByteSpan definition_part, ByteSpan records_part,
                                                         bool date1904) {
  ASSIGN_OR_RETURN(const pcd::Document doc, decode_pivot_cache_definition(definition_part));
  ASSIGN_OR_RETURN(const std::vector<CellKind> kinds, cell_kinds_of_cache_definition(definition_part));
  ASSIGN_OR_RETURN(const cache_records::PivotCacheRecordsDoc records, decode_pivot_cache_records(records_part, kinds));
  if (records.rows.size() != doc.record_count) {
    return Corrupt("definition record count does not match the records part");
  }
  pivot::PivotCache cache;
  ReadRoot(doc, cache);
  for (const pcd::Field& in : doc.fields) {
    ASSIGN_OR_RETURN(pivot::PivotCacheField field, ReadField(in, cache, date1904));
    cache.mutable_fields().push_back(std::move(field));
  }
  if (!doc.calculated_items.empty()) {
    std::size_t pos = 0;
    ASSIGN_OR_RETURN(const std::vector<pivot_calc::CalcItem> items,
                     pivot_calc::decode_calc_items(doc.calculated_items, &pos));
    if (pos != doc.calculated_items.size()) {
      return Corrupt("calculated items are followed by records outside their block");
    }
    for (const pivot_calc::CalcItem& in : items) {
      ASSIGN_OR_RETURN(pivot::PivotCalculatedItem item, ReadCalcItem(in, cache));
      cache.mutable_calculated_items().push_back(std::move(item));
    }
  }
  RETURN_IF_ERROR(ReadExtLst(doc, cache));
  RETURN_IF_ERROR(ReadRecords(records, cache, date1904));
  return cache;
}

Expected<XlsbPivotCacheParts, Error> pivot_cache_to_xlsb(const pivot::PivotCache& cache, bool date1904,
                                                         std::string_view records_rel_id) {
  pcd::Document doc;
  if (!WriteRoot(cache, &doc) || !WriteExtLst(cache, &doc)) {
    return NotWritable();
  }
  doc.records_rel_id = std::string(records_rel_id);
  if (static_cast<std::uint64_t>(cache.records().size()) > 0xFFFFFFFFULL) {
    return NotWritable();
  }
  doc.record_count = static_cast<std::uint32_t>(cache.records().size());
  for (const pivot::PivotCacheField& in : cache.fields()) {
    pcd::Field field;
    if (!WriteField(in, date1904, &field)) {
      return NotWritable();
    }
    doc.fields.push_back(std::move(field));
  }
  if (!cache.calculated_items().empty()) {
    std::vector<pivot_calc::CalcItem> items;
    for (const pivot::PivotCalculatedItem& in : cache.calculated_items()) {
      pivot_calc::CalcItem item;
      if (!WriteCalcItem(in, cache, &item)) {
        return NotWritable();
      }
      items.push_back(std::move(item));
    }
    pivot_calc::encode_calc_items(items, &doc.calculated_items);
  }
  XlsbPivotCacheParts parts;
  parts.definition = encode_pivot_cache_definition(doc);
  // The codec encodes any flag or record kind; only a layout it decodes has an Excel-saved sample.
  if (!decode_pivot_cache_definition(ByteSpan{parts.definition.data(), parts.definition.size()})) {
    return NotWritable();
  }
  ASSIGN_OR_RETURN(parts.records,
                   WriteRecords(cache, ByteSpan{parts.definition.data(), parts.definition.size()}, date1904));
  return parts;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
