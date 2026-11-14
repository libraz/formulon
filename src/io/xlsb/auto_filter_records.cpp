#include "io/xlsb/auto_filter_records.h"

#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "io/auto_filter_xml.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "utils/double_format.h"
#include "utils/double_parse.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// BrtCustomFilter vts values.
constexpr std::uint8_t kVtsNumber = 0x04;
constexpr std::uint8_t kVtsString = 0x06;
// The blank criterion: `equal ""`, kept as a custom filter only as the second of a pair.
constexpr std::uint8_t kVtsBlank = 0x0C;
// The non-blank criterion, which Excel's `.xlsx` spells as notEqual to a single space.
constexpr std::uint8_t kVtsNonBlank = 0x0E;
constexpr std::string_view kNonBlankText = " ";

constexpr std::uint32_t kKpiNil = 0xFFFFFFFFU;
// BrtBeginFilters calendarType: the ST_CalendarType values in schema order; index 0 is `none`.
constexpr std::string_view kCalendarNames[] = {"",
                                               "gregorian",
                                               "gregorianUs",
                                               "japan",
                                               "taiwan",
                                               "korea",
                                               "hijri",
                                               "thai",
                                               "hebrew",
                                               "gregorianMeFrench",
                                               "gregorianArabic",
                                               "gregorianXlitEnglish",
                                               "gregorianXlitFrench"};
constexpr std::uint32_t kCalendarCount = sizeof(kCalendarNames) / sizeof(kCalendarNames[0]);
// BrtBeginSortState fAltMethod: Excel sets it for either alternate method and reads it back as stroke.
constexpr std::uint16_t kSortAltMethod = 0x0004U;
// BrtDynamicFilter cft: CFTNIL..CFTBELOWAVERAGE are 0..2, CFTTOMORROW.. start at 8.
constexpr std::uint32_t kCftDateBase = 8U;
constexpr std::uint32_t kCftLast = 0x27U;
constexpr std::uint32_t kCftFirstPeriod = 0x18U;

Error Unsupported(std::uint16_t type) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "xlsb AutoFilter carries a record outside the model",
                    "context=xlsb_auto_filter record_id=" + std::to_string(type));
}

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_auto_filter");
}

// grbitSgn 1..6 in [MS-XLSB] 2.4.348 order.
std::optional<FilterOperator> OperatorFromSgn(std::uint8_t sgn) {
  switch (sgn) {
    case 1:
      return FilterOperator::kLessThan;
    case 2:
      return FilterOperator::kEqual;
    case 3:
      return FilterOperator::kLessThanOrEqual;
    case 4:
      return FilterOperator::kGreaterThan;
    case 5:
      return FilterOperator::kNotEqual;
    case 6:
      return FilterOperator::kGreaterThanOrEqual;
    default:
      return std::nullopt;
  }
}

/// The calendarType code for its schema name; nullopt for a name outside ST_CalendarType.
std::optional<std::uint32_t> CalendarCode(std::string_view name) {
  if (name == "none") {
    return 0U;
  }
  for (std::uint32_t i = 0; i < kCalendarCount; ++i) {
    if (kCalendarNames[i] == name) {
      return i;
    }
  }
  return std::nullopt;
}

std::uint8_t SgnFromOperator(FilterOperator op) {
  switch (op) {
    case FilterOperator::kLessThan:
      return 1;
    case FilterOperator::kEqual:
      return 2;
    case FilterOperator::kLessThanOrEqual:
      return 3;
    case FilterOperator::kGreaterThan:
      return 4;
    case FilterOperator::kNotEqual:
      return 5;
    case FilterOperator::kGreaterThanOrEqual:
      return 6;
  }
  return 2;
}

Expected<void, Error> ReadCustomFilter(ByteSpan payload, CustomFilter* out) {
  ASSIGN_OR_RETURN(const std::uint8_t vts, read_u8(payload));
  ASSIGN_OR_RETURN(const std::uint8_t sgn, read_u8(payload));
  const std::optional<FilterOperator> op = OperatorFromSgn(sgn);
  if (!op || payload.size < 8U) {
    return Corrupt("xlsb BrtCustomFilter has an unknown operator or a short union");
  }
  ByteSpan value = payload;
  payload.data += 8;
  payload.size -= 8;
  out->op = *op;
  switch (vts) {
    case kVtsNumber: {
      ASSIGN_OR_RETURN(const double number, read_double(value));
      out->val.clear();
      format_formula_number(out->val, number);
      return Expected<void, Error>::Ok();
    }
    case kVtsString: {
      ASSIGN_OR_RETURN(auto text, read_xlwidestring(payload));
      out->val = std::move(text);
      return Expected<void, Error>::Ok();
    }
    case kVtsBlank:
    case kVtsNonBlank: {
      // The union tail after the operator holds no value and is ignored.
      const bool blank = vts == kVtsBlank;
      if (*op != (blank ? FilterOperator::kEqual : FilterOperator::kNotEqual)) {
        return Unsupported(kBrtCustomFilter);
      }
      out->val = blank ? std::string() : std::string(kNonBlankText);
      return Expected<void, Error>::Ok();
    }
    default:
      // Booleans (0x08) have no OOXML spelling the model keeps apart from a string criterion.
      return Unsupported(kBrtCustomFilter);
  }
}

Expected<void, Error> ReadDateGroup(ByteSpan payload, DateGroupItem* out) {
  ASSIGN_OR_RETURN(const std::uint16_t yr, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint16_t mon, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint32_t dom, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint16_t hour, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint16_t min, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint16_t sec, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint16_t unused1, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint32_t unused2, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t dnt, read_u32(payload));
  (void)unused1;
  (void)unused2;
  if (dnt > static_cast<std::uint32_t>(DateTimeGrouping::kSecond) || mon > 12U || dom > 31U || hour > 23U ||
      min > 59U || sec > 59U) {
    return Corrupt("xlsb BrtAFilterDateGroupItem is out of range");
  }
  out->grouping = static_cast<DateTimeGrouping>(dnt);
  // Fields finer than the grouping hold Excel's placeholders; the model keeps them zero.
  const auto finer = [&](DateTimeGrouping field, std::uint32_t value) {
    return static_cast<std::uint8_t>(out->grouping < field ? 0U : value);
  };
  out->year = yr;
  out->month = finer(DateTimeGrouping::kMonth, mon);
  out->day = finer(DateTimeGrouping::kDay, dom);
  out->hour = finer(DateTimeGrouping::kHour, hour);
  out->minute = finer(DateTimeGrouping::kMinute, min);
  out->second = finer(DateTimeGrouping::kSecond, sec);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadTop10(ByteSpan payload, Top10Filter* out) {
  ASSIGN_OR_RETURN(const std::uint8_t flags, read_u8(payload));
  ASSIGN_OR_RETURN(const double value, read_double(payload));
  ASSIGN_OR_RETURN(const double filter, read_double(payload));
  out->top = (flags & 0x01U) != 0U;
  out->percent = (flags & 0x02U) != 0U;
  out->val = value;
  out->filter_val = (flags & 0x04U) != 0U ? std::optional<double>(filter) : std::nullopt;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadDynamic(ByteSpan payload, DynamicFilter* out) {
  ASSIGN_OR_RETURN(const std::uint32_t cft, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint8_t flags, read_u8(payload));
  ASSIGN_OR_RETURN(const double value, read_double(payload));
  ASSIGN_OR_RETURN(const double value_max, read_double(payload));
  std::uint32_t index = 0;
  if (cft <= 2U) {
    index = cft;
  } else if (cft >= kCftDateBase && cft <= kCftLast) {
    index = cft - kCftDateBase + 3U;
  } else {
    return Corrupt("xlsb BrtDynamicFilter has an unknown cft");
  }
  out->type = static_cast<DynamicFilterType>(index);
  if ((flags & 0x01U) != 0U && cft != 0U && cft < kCftFirstPeriod) {
    out->val = value;
    if (cft >= kCftDateBase) {
      out->max_val = value_max;
    }
  }
  return Expected<void, Error>::Ok();
}

/// Reads one FILTERCOLUMN after its BrtBeginFilterColumn.
Expected<FilterColumn, Error> ReadFilterColumn(ByteSpan& cursor, ByteSpan begin) {
  FilterColumn col;
  ASSIGN_OR_RETURN(const std::uint32_t dw_col, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(begin));
  col.col_id = dw_col;
  col.hidden_button = (flags & 0x0001U) != 0U;
  col.show_button = (flags & 0x0002U) == 0U;
  bool in_filters = false;
  bool in_custom = false;
  for (;;) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    const bool criterion_free = col.kind == FilterKind::kNone;
    switch (rec.type) {
      case kBrtEndFilterColumn:
        if (in_filters || in_custom) {
          return Corrupt("xlsb filter column ends inside a criterion block");
        }
        return col;
      case kBrtBeginFilters: {
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t blank, read_u32(payload));
        ASSIGN_OR_RETURN(const std::uint32_t calendar, read_u32(payload));
        if (calendar >= kCalendarCount) {
          return Unsupported(kBrtBeginFilters);
        }
        col.kind = FilterKind::kValues;
        col.values.blank = blank != 0U;
        col.values.calendar_type = std::string(kCalendarNames[calendar]);
        in_filters = true;
        break;
      }
      case kBrtFilter: {
        if (!in_filters) {
          return Corrupt("xlsb BrtFilter outside a filter list");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(auto text, read_xlwidestring(payload));
        col.values.values.push_back(std::move(text));
        break;
      }
      case kBrtAFilterDateGroupItem: {
        if (!in_filters) {
          return Corrupt("xlsb date group item outside a filter list");
        }
        DateGroupItem item;
        RETURN_IF_ERROR(ReadDateGroup(rec.payload, &item));
        col.values.date_groups.push_back(item);
        col.values.plain_date_groups = true;
        break;
      }
      case kBrtEndFilters:
        in_filters = false;
        break;
      case kBrtBeginCustomFilters: {
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t f_and, read_u32(payload));
        col.kind = FilterKind::kCustom;
        // Set for the AND join. The value table of [MS-XLSB] 2.4.40 reads the
        // other way round; Excel writes 0 where its `.xlsx` twin omits `and`.
        col.custom.and_join = f_and != 0U;
        in_custom = true;
        break;
      }
      case kBrtCustomFilter: {
        if (!in_custom || col.custom.filters.size() >= 2U) {
          return Corrupt("xlsb BrtCustomFilter outside its block or past two conditions");
        }
        CustomFilter filter;
        RETURN_IF_ERROR(ReadCustomFilter(rec.payload, &filter));
        col.custom.filters.push_back(std::move(filter));
        break;
      }
      case kBrtEndCustomFilters:
        in_custom = false;
        break;
      case kBrtTop10Filter:
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        col.kind = FilterKind::kTop10;
        RETURN_IF_ERROR(ReadTop10(rec.payload, &col.top10));
        break;
      case kBrtDynamicFilter:
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        col.kind = FilterKind::kDynamic;
        RETURN_IF_ERROR(ReadDynamic(rec.payload, &col.dynamic));
        break;
      case kBrtColorFilter: {
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t dxf, read_u32(payload));
        ASSIGN_OR_RETURN(const std::uint32_t cell_color, read_u32(payload));
        col.kind = FilterKind::kColor;
        col.color.dxf_id = dxf == kNoDxf ? std::nullopt : std::optional<std::uint32_t>(dxf);
        col.color.cell_color = cell_color != 0U;
        break;
      }
      case kBrtIconFilter: {
        if (!criterion_free) {
          return Corrupt("xlsb filter column carries two criteria");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t set, read_u32(payload));
        ASSIGN_OR_RETURN(const std::uint32_t icon, read_u32(payload));
        if (set == kKpiNil || set > static_cast<std::uint32_t>(cf::IconSetName::Five_Quarters)) {
          return Unsupported(kBrtIconFilter);
        }
        col.kind = FilterKind::kIcon;
        col.icon.icon_set = static_cast<cf::IconSetName>(set);
        col.icon.icon_id = icon == kKpiNil ? std::nullopt : std::optional<std::uint32_t>(icon);
        break;
      }
      default:
        return Unsupported(rec.type);
    }
  }
}

void EmitCustomFilter(std::vector<std::uint8_t>& out, const CustomFilter& filter) {
  std::vector<std::uint8_t> payload;
  const bool non_blank = filter.op == FilterOperator::kNotEqual && filter.val == kNonBlankText;
  const bool blank = filter.op == FilterOperator::kEqual && filter.val.empty();
  const std::optional<double> number = parse_full_double(filter.val);
  emit_u8(payload, blank ? kVtsBlank : non_blank ? kVtsNonBlank : number ? kVtsNumber : kVtsString);
  emit_u8(payload, SgnFromOperator(filter.op));
  // Excel's bytes after a blank or non-blank operator look like unset memory; zero is our choice.
  emit_double(payload, number.value_or(0.0));
  if (!number && !non_blank && !blank) {
    emit_xlwidestring(payload, filter.val);
  }
  emit_record(out, kBrtCustomFilter, payload);
}

void EmitDateGroup(std::vector<std::uint8_t>& out, const DateGroupItem& item) {
  // Excel fills an unused month and day with 1 and the unused time fields with 0.
  const bool by_year = item.grouping == DateTimeGrouping::kYear;
  const bool by_month = by_year || item.grouping == DateTimeGrouping::kMonth;
  std::vector<std::uint8_t> payload;
  emit_u16(payload, item.year);
  emit_u16(payload, static_cast<std::uint16_t>(by_year ? 1U : item.month));
  emit_u32(payload, by_month ? 1U : static_cast<std::uint32_t>(item.day));
  emit_u16(payload, item.hour);
  emit_u16(payload, item.minute);
  emit_u16(payload, item.second);
  emit_u16(payload, 0U);
  emit_u32(payload, 0U);
  emit_u32(payload, static_cast<std::uint32_t>(item.grouping));
  emit_record(out, kBrtAFilterDateGroupItem, payload);
}

std::uint32_t DynamicCft(DynamicFilterType type) {
  const auto index = static_cast<std::uint32_t>(type);
  return index <= 2U ? index : index - 3U + kCftDateBase;
}

/// Above / below average and the relative date ranges carry the bounds Excel computed when it applied them.
bool DynamicNeedsBounds(std::uint32_t cft) {
  return cft != 0U && cft < kCftFirstPeriod;
}

enum class ColumnFate : std::uint8_t { kWritten, kDropped, kDeferred };

/// Whether a column is written, left out as Excel leaves it out, or left out and reported.
ColumnFate FateOf(const FilterColumn& col) {
  switch (col.kind) {
    case FilterKind::kColor:
      // Excel saves no column for a colour criterion without a format.
      return col.color.dxf_id ? ColumnFate::kWritten : ColumnFate::kDropped;
    case FilterKind::kValues:
      // The rich date-group form is saved as future records the model does not write.
      return col.values.date_groups.empty() || col.values.plain_date_groups ? ColumnFate::kWritten
                                                                            : ColumnFate::kDeferred;
    case FilterKind::kTop10:
      // Excel writes the threshold it computed from the data; the model has none until a save that applied it.
      return col.top10.filter_val ? ColumnFate::kWritten : ColumnFate::kDeferred;
    case FilterKind::kDynamic: {
      const std::uint32_t cft = DynamicCft(col.dynamic.type);
      const bool bounded = col.dynamic.val && (cft < kCftDateBase || col.dynamic.max_val);
      return !DynamicNeedsBounds(cft) || bounded ? ColumnFate::kWritten : ColumnFate::kDeferred;
    }
    default:
      return ColumnFate::kWritten;
  }
}

/// Emits one column's criterion; returns the settings it could not hold.
std::uint32_t EmitCriterion(std::vector<std::uint8_t>& out, const FilterColumn& col) {
  std::vector<std::uint8_t> payload;
  switch (col.kind) {
    case FilterKind::kNone:
      return 0U;
    case FilterKind::kValues: {
      const std::optional<std::uint32_t> calendar = CalendarCode(col.values.calendar_type);
      emit_u32(payload, col.values.blank ? 1U : 0U);
      emit_u32(payload, calendar.value_or(0U));
      emit_record(out, kBrtBeginFilters, payload);
      for (const std::string& value : col.values.values) {
        payload.clear();
        emit_xlwidestring(payload, value);
        emit_record(out, kBrtFilter, payload);
      }
      for (const DateGroupItem& item : col.values.date_groups) {
        EmitDateGroup(out, item);
      }
      emit_record(out, kBrtEndFilters, ByteSpan{});
      return col.values.calendar_type.empty() || calendar ? 0U : 1U;
    }
    case FilterKind::kCustom:
      if (col.custom.filters.size() == 1U && col.custom.filters[0].op == FilterOperator::kEqual &&
          col.custom.filters[0].val.empty()) {
        // Excel converts a lone `equal ""` into the blank value filter.
        emit_u32(payload, 1U);
        emit_u32(payload, 0U);
        emit_record(out, kBrtBeginFilters, payload);
        emit_record(out, kBrtEndFilters, ByteSpan{});
        return 0U;
      }
      emit_u32(payload, col.custom.and_join ? 1U : 0U);
      emit_record(out, kBrtBeginCustomFilters, payload);
      for (std::size_t i = 0; i < col.custom.filters.size() && i < 2U; ++i) {
        EmitCustomFilter(out, col.custom.filters[i]);
      }
      emit_record(out, kBrtEndCustomFilters, ByteSpan{});
      return col.custom.filters.size() > 2U ? 1U : 0U;
    case FilterKind::kTop10: {
      std::uint32_t flags = 0;
      flags |= col.top10.top ? 0x01U : 0U;
      flags |= col.top10.percent ? 0x02U : 0U;
      flags |= 0x04U;
      emit_u8(payload, static_cast<std::uint8_t>(flags));
      emit_double(payload, col.top10.val);
      emit_double(payload, *col.top10.filter_val);
      emit_record(out, kBrtTop10Filter, payload);
      return 0U;
    }
    case FilterKind::kDynamic: {
      const std::uint32_t cft = DynamicCft(col.dynamic.type);
      const bool applied = DynamicNeedsBounds(cft);
      emit_u32(payload, cft);
      emit_u8(payload, applied ? 0x01U : 0U);
      emit_double(payload, applied ? *col.dynamic.val : 0.0);
      emit_double(payload, applied && cft >= kCftDateBase ? *col.dynamic.max_val : 0.0);
      emit_record(out, kBrtDynamicFilter, payload);
      return 0U;
    }
    case FilterKind::kColor:
      emit_u32(payload, col.color.dxf_id.value_or(0U));
      emit_u32(payload, col.color.cell_color ? 1U : 0U);
      emit_record(out, kBrtColorFilter, payload);
      return 0U;
    case FilterKind::kIcon:
      emit_u32(payload, static_cast<std::uint32_t>(col.icon.icon_set));
      emit_u32(payload, col.icon.icon_id.value_or(kKpiNil));
      emit_record(out, kBrtIconFilter, payload);
      return 0U;
  }
  return 0U;
}

Expected<SortCondition, Error> ReadSortCond(ByteSpan payload) {
  SortCondition cond;
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(payload));
  ASSIGN_OR_RETURN(const MergeRange rfx, read_rfx(payload));
  ASSIGN_OR_RETURN(const std::uint32_t first, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t second, read_u32(payload));
  ASSIGN_OR_RETURN(auto list, read_xlnullablewidestring(payload));
  const std::uint32_t sort_on = (flags >> 1U) & 0x0FU;
  if (sort_on > static_cast<std::uint32_t>(SortBy::kIcon) || !in_grid(rfx)) {
    return Corrupt("xlsb BrtBeginSortCond has an unknown sortOn or range");
  }
  cond.ref = rfx;
  cond.descending = (flags & 0x0001U) != 0U;
  cond.sort_by = static_cast<SortBy>(sort_on);
  if (cond.sort_by == SortBy::kValue) {
    cond.custom_list = std::move(list);
  } else if (cond.sort_by == SortBy::kIcon) {
    if (first == kKpiNil || first > static_cast<std::uint32_t>(cf::IconSetName::Five_Quarters)) {
      return Unsupported(kBrtBeginSortCond);
    }
    cond.icon_set = static_cast<cf::IconSetName>(first);
    cond.icon_id = second == kKpiNil ? std::nullopt : std::optional<std::uint32_t>(second);
  } else {
    cond.dxf_id = first == kNoDxf ? std::nullopt : std::optional<std::uint32_t>(first);
  }
  return cond;
}

}  // namespace

Expected<SortState, Error> read_sort_state_block(ByteSpan& cursor, ByteSpan begin) {
  SortState sort;
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(begin));
  ASSIGN_OR_RETURN(const MergeRange rfx, read_rfx(begin));
  if (!in_grid(rfx)) {
    return Corrupt("xlsb BrtBeginSortState range is inverted or outside the grid");
  }
  sort.ref = rfx;
  sort.column_sort = (flags & 0x0001U) != 0U;
  sort.case_sensitive = (flags & 0x0002U) != 0U;
  sort.sort_method = (flags & kSortAltMethod) != 0U ? SortMethod::kStroke : SortMethod::kNone;
  for (;;) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == kBrtEndSortState) {
      return sort;
    }
    if (rec.type != kBrtBeginSortCond) {
      return Unsupported(rec.type);
    }
    ASSIGN_OR_RETURN(SortCondition cond, ReadSortCond(rec.payload));
    ASSIGN_OR_RETURN(const XlsbRecord end, read_record(cursor));
    if (end.type != kBrtEndSortCond) {
      return Unsupported(end.type);
    }
    sort.conditions.push_back(std::move(cond));
  }
}

std::uint32_t emit_sort_state_block(std::vector<std::uint8_t>& out, const SortState& sort) {
  std::vector<std::uint8_t> payload;
  std::uint16_t flags = 0;
  flags |= sort.column_sort ? 0x0001U : 0U;
  flags |= sort.case_sensitive ? 0x0002U : 0U;
  flags |= sort.sort_method != SortMethod::kNone ? kSortAltMethod : 0U;
  emit_u16(payload, flags);
  emit_rfx(payload, sort.ref);
  emit_record(out, kBrtBeginSortState, payload);
  for (const SortCondition& cond : sort.conditions) {
    payload.clear();
    std::uint16_t cond_flags = cond.descending ? 0x0001U : 0U;
    cond_flags |= static_cast<std::uint16_t>(static_cast<std::uint16_t>(cond.sort_by) << 1U);
    emit_u16(payload, cond_flags);
    emit_rfx(payload, cond.ref);
    if (cond.sort_by == SortBy::kIcon) {
      emit_u32(payload, static_cast<std::uint32_t>(cond.icon_set));
      emit_u32(payload, cond.icon_id.value_or(kKpiNil));
    } else {
      emit_u32(payload, cond.sort_by == SortBy::kValue ? 0U : cond.dxf_id.value_or(kNoDxf));
      emit_u32(payload, 0U);
    }
    emit_xlnullablewidestring(payload, cond.sort_by == SortBy::kValue && !cond.custom_list.empty()
                                           ? std::optional<std::string_view>(cond.custom_list)
                                           : std::nullopt);
    emit_record(out, kBrtBeginSortCond, payload);
    emit_record(out, kBrtEndSortCond, ByteSpan{});
  }
  emit_record(out, kBrtEndSortState, ByteSpan{});
  return sort.ext_lst_xml.empty() ? 0U : 1U;
}

Expected<AutoFilter, Error> read_auto_filter_block(ByteSpan& cursor, ByteSpan begin) {
  AutoFilter filter;
  ASSIGN_OR_RETURN(const MergeRange range, read_rfx(begin));
  if (!in_grid(range)) {
    return Corrupt("xlsb BrtBeginAFilter range is inverted or outside the grid");
  }
  filter.range = range;
  for (;;) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == kBrtEndAFilter) {
      return filter;
    }
    if (rec.type == kBrtBeginSortState && !filter.sort) {
      ASSIGN_OR_RETURN(SortState sort, read_sort_state_block(cursor, rec.payload));
      filter.sort = std::move(sort);
      continue;
    }
    if (rec.type != kBrtBeginFilterColumn || filter.sort) {
      return Unsupported(rec.type);
    }
    ASSIGN_OR_RETURN(FilterColumn col, ReadFilterColumn(cursor, rec.payload));
    if (col.col_id > range.last_col - range.first_col) {
      return Corrupt("xlsb filter column is outside the range");
    }
    filter.columns.push_back(std::move(col));
  }
}

bool auto_filter_block_has_criteria(const AutoFilter& filter) noexcept {
  if (filter.is_opaque()) {
    return false;
  }
  for (const FilterColumn& col : filter.columns) {
    if (col.kind != FilterKind::kNone && FateOf(col) == ColumnFate::kWritten) {
      return true;
    }
  }
  return false;
}

std::uint32_t emit_auto_filter_block(std::vector<std::uint8_t>& out, const AutoFilter& filter) {
  if (filter.is_opaque()) {
    return 1U;
  }
  std::uint32_t lost = 0;
  std::vector<std::uint8_t> payload;
  emit_rfx(payload, filter.range);
  emit_record(out, kBeginAFilter, payload);
  for (const FilterColumn& col : filter.columns) {
    const ColumnFate fate = FateOf(col);
    if (fate != ColumnFate::kWritten) {
      lost += fate == ColumnFate::kDeferred ? 1U : 0U;
      continue;
    }
    payload.clear();
    emit_u32(payload, col.col_id);
    std::uint16_t flags = 0;
    flags |= col.hidden_button ? 0x0001U : 0U;
    flags |= col.show_button ? 0U : 0x0002U;
    emit_u16(payload, flags);
    emit_record(out, kBrtBeginFilterColumn, payload);
    lost += EmitCriterion(out, col);
    lost += (col.extra_xml.empty() ? 0U : 1U) + (col.ext_xml.empty() ? 0U : 1U);
    emit_record(out, kBrtEndFilterColumn, ByteSpan{});
  }
  if (filter.sort) {
    lost += emit_sort_state_block(out, *filter.sort);
  }
  // The filter's own extLst has no record; Excel drops it from its xlsb save too.
  emit_record(out, kBrtEndAFilter, ByteSpan{});
  return lost;
}

namespace {

/// The worksheet-level `<sortState>` kept from an `.xlsx` source, if any.
const std::string* SheetSortStateXml(const Sheet& sheet) {
  const std::string* sort_xml = nullptr;
  for (const WorksheetRawChild& child : sheet.raw_extensions()) {
    if (child.slot == worksheet_child::slot_of("sortState") && child.xml.rfind("<sortState", 0) == 0U) {
      sort_xml = &child.xml;
    }
  }
  return sort_xml;
}

/// Appends the sheet's AutoFilter and sort-state blocks; returns the settings they cannot hold.
std::uint32_t EmitSheetFilterBlocks(std::vector<std::uint8_t>& out, const Sheet& sheet) {
  std::uint32_t lost = 0;
  if (const AutoFilter* filter = sheet.auto_filter(); filter != nullptr) {
    lost += emit_auto_filter_block(out, *filter);
  }
  if (const std::string* sort_xml = SheetSortStateXml(sheet); sort_xml != nullptr) {
    if (auto sort = parse_sort_state_xml(*sort_xml)) {
      lost += emit_sort_state_block(out, sort.value());
    } else {
      ++lost;
    }
  }
  return lost;
}

}  // namespace

std::uint32_t unwritten_auto_filter_settings(const Sheet& sheet) {
  std::vector<std::uint8_t> scratch;
  return EmitSheetFilterBlocks(scratch, sheet);
}

void splice_auto_filter(std::vector<std::uint8_t>& before_merges, const Sheet& sheet) {
  const AutoFilter* filter = sheet.auto_filter();
  if (filter == nullptr && SheetSortStateXml(sheet) == nullptr) {
    return;
  }
  std::vector<std::uint8_t> block;
  EmitSheetFilterBlocks(block, sheet);
  const std::vector<FramedRecord> recs = split_records(before_merges);
  std::vector<std::uint8_t> out;
  std::size_t insert_at = 0;
  bool in_filter = false;
  for (const FramedRecord& rec : recs) {
    const bool retained_filter = filter != nullptr && (in_filter || rec.type == kBeginAFilter);
    in_filter = retained_filter && rec.type != kBrtEndAFilter;
    if (retained_filter) {
      continue;
    }
    out.insert(out.end(), before_merges.begin() + static_cast<std::ptrdiff_t>(rec.begin),
               before_merges.begin() + static_cast<std::ptrdiff_t>(rec.end));
    if (rec.type == kBrtEndScenMan) {
      insert_at = out.size();
    }
  }
  out.insert(out.begin() + static_cast<std::ptrdiff_t>(insert_at), block.begin(), block.end());
  before_merges = std::move(out);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
