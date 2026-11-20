#include "io/xlsb/pivot_cache_bridge.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>

#include "io/cell_parser.h"
#include "io/ooxml_defs.h"
#include "io/pivot_built_defaults.h"
#include "io/sqref_text.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_cache_definition_records.h"
#include "io/xlsb/pivot_cache_group_records.h"
#include "io/xlsb/pivot_cache_records.h"
#include "io/xlsb/pivot_calc_item_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xsd_double.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_formula_names.h"
#include "pivot/record_access.h"
#include "utils/arena.h"
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

/// Attribute values Excel writes on the root of an xlsx cache definition that carries `xr:uid`.
constexpr std::string_view kXrPrefix = "xr";

/// The cache extLst, as the xlsx reader keeps it, for the x14 extension the definition records hold; the
/// dataSourceClean one is the cache's session state, not kept text.
constexpr std::string_view kExtLstOpen = "<extLst>";
constexpr std::string_view kExtLstClose = "</extLst>";
constexpr std::string_view kPcd14Ext =
    "<ext uri=\"{725AE2AE-9491-48be-B2B4-4EB974FC3084}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotCacheDefinition/></ext>";
/// dataSourceClean payload of the `1` text, the only one Excel's parts show.
constexpr std::uint8_t kDataSourceCleanTrue = 1U;

/// BrtBeginPName and BrtBeginPNPair constants of every Excel-saved calculated item and calculated field.
constexpr std::uint16_t kPNameFlags = 0x0001U;
/// BrtBeginPNPair byte 0: the item is named by its cache index, by its 0-based position among the field's items in
/// their shown order (`Field[n]` stores n - 1), or by its position relative to the calculated item (`Field[-n]`,
/// `Field[+n]`); the value is the pair's item word.
constexpr std::uint8_t kPNPairByIndex = 0U;
constexpr std::uint8_t kPNPairAbsolute = 1U;
constexpr std::uint8_t kPNPairRelative = 3U;
/// BrtBeginPRFilter bit set in every Excel-saved reference, `selected` or not.
constexpr std::uint32_t kFilterAlways = 0x000001U;

Error Corrupt(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb pivot cache: " + message, kContext);
}

Error Unmeasured(const std::string& what) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "unmeasured xlsb pivot cache state: " + what, kContext);
}

Error NotWritable() {
  return not_writable(kNotWritableKind, 0);
}

/// NotWritable() naming the cache shape the writer refuses.
Error NotWritableBecause(const std::string& what) {
  Error error = NotWritable();
  error.message = "xlsb pivot cache: this writer cannot write " + what;
  return error;
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
// Calculated-field and calculated-item formulas: the engine's pivot formula parse lowered to Ptg tokens whose
// names are PtgSxName operands indexing the PNames list.

bool IsNameByte(unsigned char c) {
  return c >= 0x80U || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' ||
         c == '.';
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

/// The PNPair kind and value of a positional item reference (`n`, `+n`, `-n`), or nullopt for any other text.
std::optional<std::pair<std::uint8_t, std::uint32_t>> PositionOf(std::string_view text) {
  if (text.empty()) {
    return std::nullopt;
  }
  const bool signed_text = text.front() == '+' || text.front() == '-';
  const std::optional<std::uint32_t> n =
      ParseUnsigned(signed_text ? text.substr(1) : text, static_cast<std::uint32_t>(INT32_MAX));
  if (!n) {
    return std::nullopt;
  }
  if (!signed_text) {
    return std::make_pair(kPNPairAbsolute, *n - 1U);
  }
  if (*n == 0U) {
    return std::nullopt;
  }
  return std::make_pair(kPNPairRelative, text.front() == '-' ? 0U - *n : *n);
}

/// The text of a positional item reference; the inverse of PositionOf.
std::string PositionText(std::uint8_t kind, std::uint32_t value) {
  const auto v = static_cast<std::int32_t>(value);
  if (kind == kPNPairAbsolute) {
    return std::to_string(static_cast<std::int64_t>(v) + 1);
  }
  return (v < 0 ? "-" : "+") + std::to_string(v < 0 ? -static_cast<std::int64_t>(v) : static_cast<std::int64_t>(v));
}

const parser::AstNode* MoveLiteralSigns(Arena& arena, const parser::AstNode& node, bool fold);

/// Infix text of an rgce whose PtgSxName operands are the entries of `names` (the text of each PNames entry), each
/// listed once, in order of first use.
Expected<std::string, Error> RenderFormula(const std::vector<std::uint8_t>& rgce, const std::vector<std::uint8_t>& rgcb,
                                           const std::vector<std::string>& names, const std::string& what) {
  Arena arena;
  parser::FormulaWhitespace whitespace;
  std::size_t next_new = 0;
  const SxNameOperandHook name_operand = [&](std::uint32_t index, Arena& a) -> Expected<parser::AstNode*, Error> {
    if (index >= names.size()) {
      return Corrupt(what + " names an entry past its names table");
    }
    if (index >= next_new) {
      if (index != next_new) {
        return Unmeasured(what + " names table out of first-use order");
      }
      ++next_new;
    }
    return parser::make_name_ref(a, a.intern(names[index]));
  };
  const std::vector<std::string> no_sheets;
  const std::vector<XlsbName> no_names;
  const std::vector<XlsbSheetRange> no_ranges;
  const XlsbExternalBooks no_books;
  ASSIGN_OR_RETURN(parser::AstNode * root,
                   decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{rgcb.data(), rgcb.size()}, arena, no_sheets,
                               no_names, no_ranges, no_books, -1, std::nullopt, &whitespace, nullptr, name_operand));
  if (next_new != names.size()) {
    return Unmeasured(what + " names table with an unused entry");
  }
  const parser::AstNode* printed = MoveLiteralSigns(arena, *root, /*fold=*/false);
  if (printed == nullptr) {
    return Corrupt(what + " is too large");
  }
  return parser::format_formula(*printed, &whitespace);
}

parser::AstNode* Mutable(const parser::AstNode* node) {
  // Every node lives in the formula's own arena, which the caller owns.
  return const_cast<parser::AstNode*>(node);
}

/// `node` with the sign of number literals moved between the two spellings: `fold` turns a `-` written against a
/// number into a negative literal, the form a calculated formula stores; otherwise a negative literal becomes a `-`
/// over its magnitude, the form the formula text uses. Untouched subtrees are shared. Null when the arena is
/// exhausted.
const parser::AstNode* MoveLiteralSigns(Arena& arena, const parser::AstNode& node, bool fold) {
  const auto rewrite_all = [&arena, fold](std::uint32_t arity, auto operand_at,
                                          std::vector<const parser::AstNode*>* out) {
    bool changed = false;
    for (std::uint32_t i = 0; i < arity; ++i) {
      const parser::AstNode& operand = operand_at(i);
      const parser::AstNode* moved = MoveLiteralSigns(arena, operand, fold);
      if (moved == nullptr) {
        return std::optional<bool>();
      }
      changed = changed || moved != &operand;
      out->push_back(moved);
    }
    return std::optional<bool>(changed);
  };
  parser::AstNode* copy = nullptr;
  switch (node.kind()) {
    case parser::NodeKind::Literal: {
      if (fold || !node.as_literal().is_number() || !(node.as_literal().as_number() < 0.0)) {
        return &node;
      }
      parser::AstNode* magnitude = parser::make_literal(arena, Value::number(-node.as_literal().as_number()));
      copy = magnitude == nullptr ? nullptr : parser::make_unary_op(arena, parser::UnaryOp::Minus, magnitude);
      break;
    }
    case parser::NodeKind::UnaryOp: {
      const parser::AstNode& operand = node.as_unary_operand();
      if (fold && node.as_unary_op() == parser::UnaryOp::Minus && operand.kind() == parser::NodeKind::Literal &&
          operand.as_literal().is_number() && operand.paren_depth() == 0U) {
        copy = parser::make_literal(arena, Value::number(-operand.as_literal().as_number()));
        break;
      }
      const parser::AstNode* moved = MoveLiteralSigns(arena, operand, fold);
      if (moved == &operand) {
        return &node;
      }
      copy = moved == nullptr ? nullptr : parser::make_unary_op(arena, node.as_unary_op(), Mutable(moved));
      break;
    }
    case parser::NodeKind::ImplicitIntersection: {
      const parser::AstNode& operand = node.as_implicit_intersection_operand();
      const parser::AstNode* moved = MoveLiteralSigns(arena, operand, fold);
      if (moved == &operand) {
        return &node;
      }
      copy = moved == nullptr ? nullptr : parser::make_implicit_intersection(arena, Mutable(moved));
      break;
    }
    case parser::NodeKind::BinaryOp: {
      std::vector<const parser::AstNode*> operands;
      const std::optional<bool> changed = rewrite_all(
          node.as_binary_arity(),
          [&node](std::uint32_t i) -> const parser::AstNode& { return node.as_binary_operand(i); }, &operands);
      if (changed && !*changed) {
        return &node;
      }
      copy = changed ? parser::make_chain_with_operands(arena, node, operands.data()) : nullptr;
      break;
    }
    case parser::NodeKind::Call: {
      std::vector<const parser::AstNode*> args;
      const std::optional<bool> changed = rewrite_all(
          node.as_call_arity(), [&node](std::uint32_t i) -> const parser::AstNode& { return node.as_call_arg(i); },
          &args);
      if (changed && !*changed) {
        return &node;
      }
      copy = changed ? parser::make_call(arena, node.as_call_name(), args.data(), node.as_call_arity()) : nullptr;
      break;
    }
    default:
      return &node;
  }
  if (copy != nullptr) {
    copy->set_paren_depth(node.paren_depth());
  }
  return copy;
}

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

/// The error a BErr byte names, shared by error items and inline error record cells.
Expected<ErrorCode, Error> ErrorOfCode(std::uint8_t code, const char* what) {
  const ErrorCode error = error_from_ooxml_code(code);
  if (ooxml_code(error) != code) {
    return Unmeasured(what);
  }
  return error;
}

/// The BErr byte of `error`; false when its code does not fit one.
bool CodeOfError(ErrorCode error, std::uint8_t* out) {
  const std::int32_t code = ooxml_code(error);
  if (code < 0 || code > 0xFF) {
    return false;
  }
  *out = static_cast<std::uint8_t>(code);
  return true;
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
    if (kind != Plain::kNone && last_count >= pivot_records::kMinItemRunLength &&
        (!entry.is_run || last_count != count)) {
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
          ASSIGN_OR_RETURN(const ErrorCode code, ErrorOfCode(item.error, "shared error item code"));
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
    if (run.size() < pivot_records::kMinItemRunLength) {
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

/// The model period of a BrtBeginPCDFGRange `groupBy`, or nullopt past ST_GroupBy.
std::optional<pivot::GroupBy> MeasuredGroupBy(std::uint8_t code) {
  if (code > pivot_group::kGroupByYears) {
    return std::nullopt;
  }
  return static_cast<pivot::GroupBy>(code);
}
static_assert(static_cast<std::uint8_t>(pivot::GroupBy::Range) == pivot_group::kGroupByRange &&
              static_cast<std::uint8_t>(pivot::GroupBy::Seconds) == pivot_group::kGroupBySeconds &&
              static_cast<std::uint8_t>(pivot::GroupBy::Minutes) == pivot_group::kGroupByMinutes &&
              static_cast<std::uint8_t>(pivot::GroupBy::Hours) == pivot_group::kGroupByHours &&
              static_cast<std::uint8_t>(pivot::GroupBy::Days) == pivot_group::kGroupByDays &&
              static_cast<std::uint8_t>(pivot::GroupBy::Months) == pivot_group::kGroupByMonths &&
              static_cast<std::uint8_t>(pivot::GroupBy::Quarters) == pivot_group::kGroupByQuarters &&
              static_cast<std::uint8_t>(pivot::GroupBy::Years) == pivot_group::kGroupByYears);

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
    const std::optional<pivot::GroupBy> by = MeasuredGroupBy(r.group_by);
    // A numeric range groups by `range` over numbers; every period groups dates.
    const bool dates = (r.flags & pivot_group::kRangeDates) != 0U;
    if (!by || (r.flags & ~kKnown) != 0U || dates != (*by != pivot::GroupBy::Range)) {
      return Unmeasured("range group whose period and value type disagree");
    }
    out.kind = pivot::GroupKind::Range;
    out.range.by = *by;
    out.range.is_date = dates;
    if (dates) {
      ASSIGN_OR_RETURN(out.range.start_date, DateIso(r.start, date1904));
      ASSIGN_OR_RETURN(out.range.end_date, DateIso(r.end, date1904));
    } else {
      out.range.start_num = r.start;
      out.range.end_num = r.end;
    }
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
  if (group.discrete) {
    if (out.group_items.empty()) {
      return Unmeasured("discrete group without group items");
    }
    for (const std::uint32_t index : *group.discrete) {
      if (index >= out.group_items.size()) {
        return Corrupt("discrete group maps a base item to a group item it does not list");
      }
    }
    out.kind = pivot::GroupKind::Discrete;
    out.discrete_map = *group.discrete;
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
  const std::vector<pivot_calc::PName> empty;
  const std::vector<pivot_calc::PName>& names = in.names ? *in.names : empty;
  std::vector<std::string> texts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    const pivot_calc::PName& name = names[i];
    if (name.field != pivot_calc::kNoField || name.flags != kPNameFlags || !name.pairs || name.pairs->size() != 1U) {
      return Unmeasured("calculated item name shape");
    }
    const pivot_calc::PNPair& pair = name.pairs->front();
    const bool positional = pair.flags == kPNPairAbsolute || pair.flags == kPNPairRelative;
    if ((pair.flags != kPNPairByIndex && !positional) || (pair.flags == kPNPairRelative && pair.item == 0U)) {
      return Unmeasured("calculated item name kind");
    }
    if (pair.field >= cache.fields().size() ||
        (!positional && pair.item >= cache.fields()[pair.field].shared_items.size())) {
      return Corrupt("calculated item name points past the cache");
    }
    for (std::size_t j = 0; j < i; ++j) {
      const pivot_calc::PNPair& prior = names[j].pairs->front();
      if (prior.field == pair.field && prior.item == pair.item && prior.flags == pair.flags) {
        return Unmeasured("calculated item names table with a repeated entry");
      }
    }
    const pivot::PivotCacheField& field = cache.fields()[pair.field];
    std::string item_text;
    if (positional) {
      item_text = PositionText(pair.flags, pair.item);
    } else {
      const Value& item = field.shared_items[pair.item];
      if (!item.is_text() || !IsPlainItem(item.as_text()) || PositionOf(item.as_text())) {
        return Unmeasured("calculated item name that is not a plain Field[Item]");
      }
      item_text = std::string(item.as_text());
    }
    std::string text;
    pivot::append_formula_name(text, field.name);
    text.append("[").append(item_text).append("]");
    texts.push_back(std::move(text));
  }
  pivot::PivotCalculatedItem item;
  ASSIGN_OR_RETURN(item.formula, RenderFormula(in.rgce, in.rgcb, texts, "calculated item"));
  item.has_area = in.area.has_value();
  if (in.area) {
    ASSIGN_OR_RETURN(item.area, ReadArea(*in.area, cache));
    if (!item.area.references.empty() && item.area.references.front().field.has_value()) {
      item.cache_field_index = *item.area.references.front().field;
    }
  }
  return item;
}

/// A calculated field's formula text: its PNames name cache fields (any field, a calculated one included) by
/// index, each spelled as the formula spells a field name.
Expected<std::string, Error> ReadFieldFormula(const pcd::Field& in, const pivot::PivotCache& cache) {
  if (!in.formula_rgcb.empty()) {
    return Unmeasured("calculated field formula extra data");
  }
  std::vector<pivot_calc::PName> names;
  if (!in.names.empty()) {
    ASSIGN_OR_RETURN(names, pivot_calc::decode_names(in.names));
  }
  std::vector<std::string> texts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    const pivot_calc::PName& name = names[i];
    if (name.flags != kPNameFlags || name.pairs) {
      return Unmeasured("calculated field name shape");
    }
    if (name.field >= cache.fields().size()) {
      return Corrupt("calculated field name points past the cache");
    }
    for (std::size_t j = 0; j < i; ++j) {
      if (names[j].field == name.field) {
        return Unmeasured("calculated field names table with a repeated entry");
      }
    }
    std::string text;
    pivot::append_formula_name(text, cache.fields()[name.field].name);
    texts.push_back(std::move(text));
  }
  return RenderFormula(*in.formula_rgce, in.formula_rgcb, texts, "calculated field");
}

void ReadRoot(const pcd::Document& doc, pivot::PivotCache& cache) {
  std::vector<std::pair<std::string, std::string>>& attrs = cache.mutable_passthrough_attrs();
  if (doc.uid) {
    attrs.emplace_back("xmlns:mc", std::string(kNsMarkupCompat));
    attrs.emplace_back("xmlns:xr", std::string(kNsXr));
    attrs.emplace_back("mc:Ignorable", std::string(kXrPrefix));
  }
  if (doc.refreshed_by) {
    attrs.emplace_back("refreshedBy", *doc.refreshed_by);
  }
  attrs.emplace_back("refreshedDate", RefreshedDateText(doc.refreshed_date));
  attrs.emplace_back("createdVersion", std::to_string(doc.created_version));
  attrs.emplace_back("refreshedVersion", std::to_string(doc.refreshed_version));
  attrs.emplace_back("minRefreshableVersion", std::to_string(doc.min_refreshable_version));
  if (doc.uid) {
    attrs.emplace_back("xr:uid", FormatGuid(doc.uid->data()));
  }
  cache.set_save_data((doc.flags & pcd::kCacheSaveData) != 0U);
  if (!cache.save_data()) {
    cache.set_unsaved_record_count(doc.record_count);
  }
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
  source.sheet = doc.source.sheet;
  source.name = doc.source.name;
  if (doc.source.name.empty()) {
    source.ref = format_sqref_text({doc.source.ref});
  }
}

Expected<void, Error> ReadExtLst(const pcd::Document& doc, pivot::PivotCache& cache) {
  // dataSourceClean is session state: every load starts clean whatever the file says.
  if (doc.data_source_clean && *doc.data_source_clean != kDataSourceCleanTrue) {
    return Unmeasured("dataSourceClean other than 1");
  }
  if (doc.has_pcd14) {
    std::string& xml = cache.mutable_trailing_children_after_calculated_xml();
    xml.append(kExtLstOpen).append(kPcd14Ext).append(kExtLstClose);
  }
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
        case CellKind::kError: {
          ASSIGN_OR_RETURN(const ErrorCode code, ErrorOfCode(cell.error, "record error cell code"));
          record.cells[pos] = Value::error(code);
          break;
        }
        case CellKind::kMissing:
          break;
        case CellKind::kNumberOrError:
        case CellKind::kNumberOrBlank:
        case CellKind::kNumberMixed:
        case CellKind::kNumberMixedBlank:
        case CellKind::kTextOrBlank:
        case CellKind::kDateOrBlank:
          return Corrupt("record cell carries a field-only kind");
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
        !parse_xsd_double(h.min_value, &table->min) || !parse_xsd_double(h.max_value, &table->max)) {
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
    std::uint8_t code = 0;
    *out = ItemClass::kError;
    return CodeOfError(v.as_error(), &code);
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
    if (end - i >= pivot_records::kMinItemRunLength) {
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

/// The measured discrete group: a group field directly over a database field that is not itself grouped, with
/// fewer than pivot_records::kMinItemRunLength base items, each mapped to a listed group item.
Expected<void, Error> WriteDiscrete(const pivot::FieldGroup& in, const pivot::PivotCache& cache,
                                    pivot_group::FieldGroup* group) {
  if (in.par) {
    return NotWritableBecause("a discrete group that a further group field groups");
  }
  if (!in.base || *in.base >= cache.fields().size()) {
    return NotWritableBecause("a discrete group with no base field");
  }
  const pivot::PivotCacheField& base = cache.fields()[*in.base];
  if (!base.is_database_field || (base.field_group && base.field_group->kind != pivot::GroupKind::None)) {
    return NotWritableBecause("a discrete group over a field that is itself a group");
  }
  if (in.discrete_map.size() >= pivot_records::kMinItemRunLength) {
    return NotWritableBecause("a discrete group over four or more base items");
  }
  if (in.discrete_map.size() != base.shared_items.size()) {
    return NotWritableBecause("a discrete group whose map does not cover its base field's items");
  }
  for (const std::uint32_t index : in.discrete_map) {
    if (index >= in.group_items.size()) {
      return NotWritableBecause("a discrete group mapping a base item to a group item it does not list");
    }
  }
  group->discrete = in.discrete_map;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> WriteGroup(const pivot::FieldGroup& in, const pivot::PivotCache& cache, bool date1904,
                                 std::vector<pivot_records::RawRecord>* out) {
  if (!in.extra_attrs.empty()) {
    return NotWritableBecause("a field group attribute it does not model");
  }
  pivot_group::FieldGroup group;
  group.par = in.par.value_or(pivot_group::kNoField);
  group.base = in.base.value_or(pivot_group::kNoField);
  if (in.kind == pivot::GroupKind::Discrete) {
    RETURN_IF_ERROR(WriteDiscrete(in, cache, &group));
  }
  if (in.kind == pivot::GroupKind::Range) {
    const pivot::RangeGroup& r = in.range;
    if (!MeasuredGroupBy(static_cast<std::uint8_t>(r.by)) || r.is_date != (r.by != pivot::GroupBy::Range)) {
      return NotWritableBecause("a range group whose period and value type disagree");
    }
    const std::optional<double> start = r.is_date ? IsoSerial(r.start_date, date1904) : r.start_num;
    const std::optional<double> end = r.is_date ? IsoSerial(r.end_date, date1904) : r.end_num;
    if (!start || !end) {
      return NotWritableBecause("a range group bound that is not a date");
    }
    pivot_group::RangeGroup range;
    range.group_by = static_cast<std::uint8_t>(r.by);
    range.flags = static_cast<std::uint8_t>((r.is_date ? pivot_group::kRangeDates : 0U) |
                                            (r.auto_start ? pivot_group::kRangeAutoStart : 0U) |
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
        return NotWritableBecause("a group item that is not text");
      }
      items.emplace_back(in.group_items[i].as_text());
    }
    group.items = std::move(items);
  }
  pivot_group::encode_field_group(group, out);
  return Expected<void, Error>::Ok();
}

bool HasTextItem(const pivot::PivotCache& cache, std::size_t field, std::string_view text) {
  for (const Value& item : cache.fields()[field].shared_items) {
    if (item.is_text() && item.as_text() == text) {
      return true;
    }
  }
  return false;
}

/// Lowers a calculated field's or calculated item's formula, `text` as `parsed` holds it (null or empty when it does
/// not parse), to rgce and its PNames entries; false when the text is not one Excel stores the same way, so that it
/// does not read back as written.
bool LowerFormula(const std::string& text, std::optional<pivot::ParsedFormula>* parsed_slot,
                  const pivot::PivotCache& cache, std::vector<std::uint8_t>* rgce,
                  std::vector<pivot_calc::PName>* names) {
  if (parsed_slot == nullptr || !parsed_slot->has_value()) {
    return false;
  }
  pivot::ParsedFormula& parsed = **parsed_slot;
  std::vector<std::string> texts;
  for (const pivot::FormulaRef& ref : parsed.refs) {
    const pivot::PivotCacheField& field = cache.fields()[ref.field];
    pivot_calc::PName name;
    name.flags = kPNameFlags;
    std::string name_text;
    pivot::append_formula_name(name_text, field.name);
    if (ref.kind == pivot::FormulaRef::Kind::Field) {
      name.field = static_cast<std::uint32_t>(ref.field);
    } else {
      pivot_calc::PNPair pair{kPNPairByIndex, static_cast<std::uint32_t>(ref.field), 0U};
      std::string item_text;
      if (ref.kind == pivot::FormulaRef::Kind::Item) {
        pair.item = static_cast<std::uint32_t>(ref.item);
        item_text = std::string(field.shared_items[ref.item].as_text());
      } else {
        if (ref.kind == pivot::FormulaRef::Kind::ItemAt) {
          pair.flags = kPNPairAbsolute;
          pair.item = static_cast<std::uint32_t>(ref.position) - 1U;
        } else if (ref.position != 0) {
          pair.flags = kPNPairRelative;
          pair.item = static_cast<std::uint32_t>(ref.position);
        } else {
          return false;
        }
        item_text = PositionText(pair.flags, pair.item);
        // A position spelled the same as an item has no Excel-saved sample.
        if (HasTextItem(cache, ref.field, item_text)) {
          return false;
        }
      }
      name.field = pivot_calc::kNoField;
      name.pairs = std::vector<pivot_calc::PNPair>{pair};
      name_text.append("[").append(item_text).append("]");
    }
    names->push_back(std::move(name));
    texts.push_back(std::move(name_text));
  }
  const parser::AstNode* root = MoveLiteralSigns(parsed.arena, *parsed.root, /*fold=*/true);
  if (root == nullptr) {
    return false;
  }
  const NameOperandHook name_operand = [&parsed](const parser::AstNode& node) -> std::optional<std::uint32_t> {
    const std::optional<std::size_t> slot = pivot::formula_ref_slot(node.as_name(), parsed.refs.size());
    return slot ? std::optional<std::uint32_t>(static_cast<std::uint32_t>(*slot)) : std::nullopt;
  };
  const std::vector<std::string> no_sheets;
  const SheetRangeTable no_ranges;
  const NameTable no_names;
  auto encoded = encode_ptgs(*root, no_sheets, no_ranges, no_names, PtgRootClass::kValue, std::nullopt,
                             PtgEvaluation::kDynamicArray, NameShapes(), {}, name_operand);
  if (!encoded || !encoded.value().rgcb.empty()) {
    return false;
  }
  auto rendered = RenderFormula(encoded.value().rgce, {}, texts, "calculated formula");
  if (!rendered || rendered.value() != text) {
    return false;
  }
  *rgce = std::move(encoded.value().rgce);
  return true;
}

bool WriteFieldFormula(const std::string& text, std::optional<pivot::ParsedFormula>* parsed,
                       const pivot::PivotCache& cache, pcd::Field* out) {
  std::vector<pivot_calc::PName> names;
  std::vector<std::uint8_t> rgce;
  if (!LowerFormula(text, parsed, cache, &rgce, &names) || names.empty()) {
    return false;
  }
  out->formula_rgce = std::move(rgce);
  pivot_calc::encode_names(names, &out->names);
  return true;
}

Expected<void, Error> WriteField(const pivot::PivotCacheField& in, std::optional<pivot::ParsedFormula>* parsed,
                                 const pivot::PivotCache& cache, bool date1904, pcd::Field* out) {
  if (!in.extra_attrs.empty() || !in.trailing_children_xml.empty() || !in.num_fmt_id || *in.num_fmt_id > 0xFFFFU) {
    return NotWritable();
  }
  if (in.formula && (in.is_database_field || !in.shared_items.empty() || in.shared_items_hints.present ||
                     in.field_group || !WriteFieldFormula(*in.formula, parsed, cache, out))) {
    return NotWritable();
  }
  out->flags = in.is_database_field ? pcd::kFieldDatabase : 0U;
  out->num_fmt_id = static_cast<std::uint16_t>(*in.num_fmt_id);
  out->name = in.name;
  if (in.is_database_field || !in.shared_items.empty() || in.shared_items_hints.present) {
    pcd::ItemTable table;
    if (!WriteHints(in.shared_items_hints, date1904, &table) || !WriteItems(in, date1904, &table)) {
      return NotWritable();
    }
    out->items = std::move(table);
  }
  if (in.field_group) {
    RETURN_IF_ERROR(WriteGroup(*in.field_group, cache, date1904, &out->grouping));
  }
  return Expected<void, Error>::Ok();
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

bool WriteCalcItem(const pivot::PivotCalculatedItem& in, std::optional<pivot::ParsedFormula>* parsed,
                   const pivot::PivotCache& cache, pivot_calc::CalcItem* out) {
  if (in.has_field_attr || !in.extra_children_xml.empty()) {
    return false;
  }
  if (in.cache_field_index >= cache.fields().size()) {
    return false;
  }
  std::vector<pivot_calc::PName> names;
  std::vector<std::uint8_t> rgce;
  if (!LowerFormula(in.formula, parsed, cache, &rgce, &names)) {
    return false;
  }
  out->field = pivot_calc::kNoField;
  out->rgce = std::move(rgce);
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
    // A cache no refresh dated keeps refreshedDate 0, as Excel saves one.
    kRequired = kCreated | kRefreshed | kMinRefreshable,
    kNsMc = 32U,
    kNsXrDecl = 64U,
    kIgnorable = 128U,
    kUidBookkeeping = 224U,
  };
  unsigned seen = 0;
  if (cache_built_in_memory(cache)) {
    // What Excel's re-save of the xlsx writer's part holds: refreshedDate 0 and no refreshedBy.
    doc->created_version = kBuiltCacheCreatedVersion;
    doc->refreshed_version = kBuiltCacheRefreshedVersion;
    doc->min_refreshable_version = kBuiltCacheMinRefreshableVersion;
    seen = kRequired;
  }
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
  // Every Excel save of a cache without its data names the user who refreshed it.
  if ((seen & kRequired) != kRequired || bookkeeping != (doc->uid ? kUidBookkeeping : 0U) ||
      (!cache.save_data() && !doc->refreshed_by)) {
    return false;
  }
  const bool refresh_on_load = cache.refresh_on_load() || built_cache_refreshes_on_load(cache);
  doc->flags = static_cast<std::uint8_t>(pcd::kCacheEnableRefresh | (cache.save_data() ? pcd::kCacheSaveData : 0U) |
                                         (refresh_on_load ? pcd::kCacheRefreshOnLoad : 0U));
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
  if (!cache.cache_source_children_xml().empty() || !source.present) {
    return false;
  }
  // A defined name or table is written by its name alone, with the sheet of a sheet-scoped name.
  if (!source.name.empty()) {
    if (!source.ref.empty()) {
      return false;
    }
    doc->source.name = source.name;
    doc->source.sheet = source.sheet;
    return true;
  }
  if (source.sheet.empty() || !parse_a1_range(source.ref, &doc->source.ref) ||
      format_sqref_text({doc->source.ref}) != source.ref) {
    return false;
  }
  doc->source.sheet = source.sheet;
  return true;
}

bool WriteExtLst(const pivot::PivotCache& cache, pcd::Document* doc) {
  if (!cache.trailing_children_xml().empty()) {
    return false;
  }
  if (cache.source_clean()) {
    doc->data_source_clean = kDataSourceCleanTrue;
  }
  std::string_view xml = cache.trailing_children_after_calculated_xml();
  if (xml.empty()) {
    // Excel's re-save of a cache built in memory adds the x14 extension.
    doc->has_pcd14 = cache_built_in_memory(cache);
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
  return consume(kExtLstClose) && xml.empty() && doc->has_pcd14;
}

bool WriteRecordCell(const pivot::PivotCache& cache, const pivot::PivotCacheRecord& record, std::size_t pos,
                     CellKind kind, bool date1904, RecordCell* out) {
  const Value& v = record.cells[pos];
  const pivot::PivotCacheField& field = cache.fields()[pos];
  const bool is_index = pivot::cell_is_shared_index(record, field, pos);
  const bool is_date = FlagAt(record.cell_is_date, pos);
  out->kind = kind;
  switch (kind) {
    case CellKind::kNumberOrBlank:
    case CellKind::kTextOrBlank:
    case CellKind::kDateOrBlank:
      if (!is_index && v.is_blank()) {
        out->kind = CellKind::kMissing;
        return true;
      }
      return WriteRecordCell(cache, record, pos,
                             kind == CellKind::kNumberOrBlank ? CellKind::kNumber
                             : kind == CellKind::kTextOrBlank ? CellKind::kText
                                                              : CellKind::kDate,
                             date1904, out);
    case CellKind::kNumberMixed:
    case CellKind::kNumberMixedBlank:
      if (is_index || is_date) {
        return false;
      }
      if (kind == CellKind::kNumberMixedBlank && v.is_blank()) {
        out->kind = CellKind::kMissing;
        return true;
      }
      if (v.is_number()) {
        out->kind = CellKind::kNumber;
        out->number = v.as_number();
        return true;
      }
      if (v.is_text()) {
        out->kind = CellKind::kText;
        out->text = std::string(v.as_text());
        return true;
      }
      out->kind = CellKind::kError;
      return v.is_error() && CodeOfError(v.as_error(), &out->error);
    case CellKind::kMissing:
      return false;
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
    case CellKind::kNumberOrError:
      if (is_index || is_date) {
        return false;
      }
      if (v.is_number()) {
        out->kind = CellKind::kNumber;
        out->number = v.as_number();
        return true;
      }
      out->kind = CellKind::kError;
      return v.is_error() && CodeOfError(v.as_error(), &out->error);
    case CellKind::kError:
      return false;
    case CellKind::kText:
      if (is_index) {
        return false;
      }
      if (v.is_error()) {
        out->kind = CellKind::kError;
        return CodeOfError(v.as_error(), &out->error);
      }
      if (!v.is_text()) {
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
    return Unmeasured("record cells of a field without shared items holding blanks or more than one value type");
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
    // A calculated field added after the records leaves them without its trailing cell.
    if (record.cells.size() > cache.fields().size() ||
        (!db_positions.empty() && record.cells.size() <= db_positions.back())) {
      return NotWritable();
    }
    std::vector<RecordCell> row(db_positions.size());
    for (std::size_t j = 0; j < db_positions.size(); ++j) {
      if (!WriteRecordCell(cache, record, db_positions[j], kinds.value()[j], date1904, &row[j])) {
        return NotWritable();
      }
    }
    // A group or calculated field's cell (the group number a load materialises) has no record cell.
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
  const bool save_data = (doc.flags & pcd::kCacheSaveData) != 0U;
  cache_records::PivotCacheRecordsDoc records;
  if (save_data) {
    ASSIGN_OR_RETURN(const std::vector<CellKind> kinds, cell_kinds_of_cache_definition(definition_part));
    ASSIGN_OR_RETURN(records, decode_pivot_cache_records(records_part, kinds));
    if (records.rows.size() != doc.record_count) {
      return Corrupt("definition record count does not match the records part");
    }
  }
  pivot::PivotCache cache;
  ReadRoot(doc, cache);
  for (const pcd::Field& in : doc.fields) {
    ASSIGN_OR_RETURN(pivot::PivotCacheField field, ReadField(in, cache, date1904));
    cache.mutable_fields().push_back(std::move(field));
  }
  // A calculated field may name any field, so formulas resolve once every field is read.
  for (std::size_t i = 0; i < doc.fields.size(); ++i) {
    if (doc.fields[i].formula_rgce) {
      ASSIGN_OR_RETURN(std::string formula, ReadFieldFormula(doc.fields[i], cache));
      cache.mutable_fields()[i].formula = std::move(formula);
    }
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

Expected<XlsbPivotCacheParts, Error> pivot_cache_to_xlsb(const pivot::PivotCache& cache,
                                                         pivot::ParsedCacheFormulas formulas, bool date1904,
                                                         std::string_view records_rel_id) {
  const auto slot = [](std::vector<std::optional<pivot::ParsedFormula>>& parsed,
                       std::size_t index) -> std::optional<pivot::ParsedFormula>* {
    return index < parsed.size() ? &parsed[index] : nullptr;
  };
  pcd::Document doc;
  if (!WriteRoot(cache, &doc) || !WriteExtLst(cache, &doc)) {
    return NotWritable();
  }
  if (cache.save_data()) {
    doc.records_rel_id = std::string(records_rel_id);
  }
  if (static_cast<std::uint64_t>(cache.record_count_to_save()) > 0xFFFFFFFFULL) {
    return NotWritable();
  }
  doc.record_count = static_cast<std::uint32_t>(cache.record_count_to_save());
  for (std::size_t i = 0; i < cache.fields().size(); ++i) {
    pcd::Field field;
    RETURN_IF_ERROR(WriteField(cache.fields()[i], slot(formulas.fields, i), cache, date1904, &field));
    doc.fields.push_back(std::move(field));
  }
  if (!cache.calculated_items().empty()) {
    std::vector<pivot_calc::CalcItem> items;
    for (std::size_t i = 0; i < cache.calculated_items().size(); ++i) {
      pivot_calc::CalcItem item;
      if (!WriteCalcItem(cache.calculated_items()[i], slot(formulas.items, i), cache, &item)) {
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
  if (cache.save_data()) {
    ASSIGN_OR_RETURN(parts.records,
                     WriteRecords(cache, ByteSpan{parts.definition.data(), parts.definition.size()}, date1904));
  }
  return parts;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
