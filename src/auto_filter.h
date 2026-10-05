//
// AutoFilter model: the typed filter columns and criteria of a sheet or table
// AutoFilter (ECMA-376 §18.3.2, `CT_AutoFilter`).
//
// The typed model is the source of truth; `io/auto_filter_xml.h` reads and
// writes the `<autoFilter>` element. Content the model does not name is kept
// as raw XML on the owning record, and a fragment that does not parse into
// the model at all is kept verbatim as `opaque_xml`.

#ifndef FORMULON_AUTO_FILTER_H_
#define FORMULON_AUTO_FILTER_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "merge_range.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;
class Workbook;

/// Attributes retained verbatim as `(name, value)` pairs, in document order.
using RawAttributes = std::vector<std::pair<std::string, std::string>>;

/// Which criterion a filter column carries.
enum class FilterKind : std::uint8_t {
  kNone = 0,
  kValues = 1,
  kCustom = 2,
  kTop10 = 3,
  kDynamic = 4,
  kColor = 5,
  kIcon = 6,
};

/// `ST_DateTimeGrouping`: the granularity a date group item matches at.
enum class DateTimeGrouping : std::uint8_t { kYear = 0, kMonth, kDay, kHour, kMinute, kSecond };

/// One `<dateGroupItem>`. Fields finer than `grouping` are unused and zero.
struct DateGroupItem {
  std::uint16_t year = 0;
  std::uint8_t month = 0;
  std::uint8_t day = 0;
  std::uint8_t hour = 0;
  std::uint8_t minute = 0;
  std::uint8_t second = 0;
  DateTimeGrouping grouping = DateTimeGrouping::kYear;
};

/// `<filters>`: a list of displayed values, date groups and/or blanks.
struct ValueFilters {
  bool blank = false;
  /// `calendarType`; empty means the schema default `none`.
  std::string calendar_type;
  std::vector<std::string> values;
  std::vector<DateGroupItem> date_groups;
};

/// `ST_FilterOperator`.
enum class FilterOperator : std::uint8_t {
  kEqual = 0,
  kLessThan,
  kLessThanOrEqual,
  kNotEqual,
  kGreaterThanOrEqual,
  kGreaterThan,
};

/// One `<customFilter>`.
struct CustomFilter {
  FilterOperator op = FilterOperator::kEqual;
  std::string val;
};

/// `<customFilters>`: one or two conditions joined by and/or.
struct CustomFilters {
  bool and_join = false;
  std::vector<CustomFilter> filters;
};

/// `<top10>`.
struct Top10Filter {
  bool top = true;
  bool percent = false;
  double val = 0.0;
  std::optional<double> filter_val;
};

/// `ST_DynamicFilterType`, in schema order.
enum class DynamicFilterType : std::uint8_t {
  kNull = 0,
  kAboveAverage,
  kBelowAverage,
  kTomorrow,
  kToday,
  kYesterday,
  kNextWeek,
  kThisWeek,
  kLastWeek,
  kNextMonth,
  kThisMonth,
  kLastMonth,
  kNextQuarter,
  kThisQuarter,
  kLastQuarter,
  kNextYear,
  kThisYear,
  kLastYear,
  kYearToDate,
  kQ1,
  kQ2,
  kQ3,
  kQ4,
  kM1,
  kM2,
  kM3,
  kM4,
  kM5,
  kM6,
  kM7,
  kM8,
  kM9,
  kM10,
  kM11,
  kM12,
};

/// `<dynamicFilter>`. `val`/`max_val` are the bounds Excel stored when it
/// last applied the filter; ISO strings are kept as written.
struct DynamicFilter {
  DynamicFilterType type = DynamicFilterType::kNull;
  std::optional<double> val;
  std::optional<double> max_val;
  std::string val_iso;
  std::string max_val_iso;
};

/// `<colorFilter>`: the dxf whose fill (`cell_color`) or font colour matches.
struct ColorFilter {
  std::optional<std::uint32_t> dxf_id;
  bool cell_color = true;
};

/// `<iconFilter>`. An absent `icon_id` means "no icon".
struct IconFilter {
  cf::IconSetName icon_set = cf::IconSetName::Three_Arrows;
  std::optional<std::uint32_t> icon_id;
};

/// One `<filterColumn>`. `col_id` is the 0-based offset from the AutoFilter
/// range's first column. Only the member selected by `kind` is meaningful.
struct FilterColumn {
  std::uint32_t col_id = 0;
  bool hidden_button = false;
  bool show_button = true;
  FilterKind kind = FilterKind::kNone;
  ValueFilters values;
  CustomFilters custom;
  Top10Filter top10;
  DynamicFilter dynamic;
  ColorFilter color;
  IconFilter icon;
  RawAttributes extra_attrs;
  /// Unknown child elements, emitted after the criterion.
  std::string extra_xml;
  /// `<ext>` children of the column's `<extLst>` other than the date-group
  /// extension the model lifts into `values`.
  std::string ext_xml;
};

/// `ST_SortBy`.
enum class SortBy : std::uint8_t { kValue = 0, kCellColor, kFontColor, kIcon };

/// `ST_SortMethod`.
enum class SortMethod : std::uint8_t { kNone = 0, kPinYin, kStroke };

/// One `<sortCondition>`.
struct SortCondition {
  MergeRange ref;
  bool descending = false;
  SortBy sort_by = SortBy::kValue;
  std::string custom_list;
  std::optional<std::uint32_t> dxf_id;
  cf::IconSetName icon_set = cf::IconSetName::Three_Arrows;
  std::optional<std::uint32_t> icon_id;
};

/// `<sortState>` nested in an AutoFilter.
struct SortState {
  MergeRange ref;
  bool column_sort = false;
  bool case_sensitive = false;
  SortMethod sort_method = SortMethod::kNone;
  std::vector<SortCondition> conditions;
  RawAttributes extra_attrs;
  /// Raw `<extLst>` element, or empty.
  std::string ext_lst_xml;
};

/// A sheet or table AutoFilter. When `opaque_xml` is non-empty the element
/// did not parse and every other field is meaningless.
struct AutoFilter {
  MergeRange range;
  std::vector<FilterColumn> columns;
  std::optional<SortState> sort;
  /// Raw `<extLst>` element of the AutoFilter itself, or empty.
  std::string ext_lst_xml;
  RawAttributes extra_attrs;
  std::string opaque_xml;

  bool is_opaque() const noexcept { return !opaque_xml.empty(); }

  /// True when at least one column carries a criterion.
  bool has_criteria() const noexcept;
};

/// Parses an A1 rectangle (`B2:F12`, or a single cell `D6`) without `$`
/// anchors or a sheet qualifier.
std::optional<MergeRange> parse_a1_rectangle(std::string_view text) noexcept;

/// Appends `rect` in A1 form, collapsing a single cell to one address.
/// Returns false when a coordinate is outside the grid.
bool append_a1_rectangle(std::string& out, const MergeRange& rect);

/// Checks a caller-built model: a typed, in-grid range, column ids inside
/// the range and strictly ascending, one or two custom conditions, and date
/// groups whose fields fit their grouping. Returns `kAutoFilterInvalid`.
Expected<void, Error> validate_auto_filter(const AutoFilter& filter);

/// True when the sheet AutoFilter, or the AutoFilter of a table on the
/// sheet, carries at least one criterion. Excel then treats every hidden row
/// of the sheet as filtered, inside the filter range or not.
bool sheet_has_filter_criteria(const Workbook* wb, const Sheet& sheet) noexcept;

/// An optional AutoFilter held by a sheet or table.
class AutoFilterSlot {
 public:
  bool empty() const noexcept { return !value_.has_value(); }
  const AutoFilter* get() const noexcept { return value_ ? &*value_ : nullptr; }
  AutoFilter* get() noexcept { return value_ ? &*value_ : nullptr; }
  void set(AutoFilter filter) { value_ = std::move(filter); }
  void reset() noexcept { value_.reset(); }

 private:
  std::optional<AutoFilter> value_;
};

}  // namespace formulon

#endif  // FORMULON_AUTO_FILTER_H_
