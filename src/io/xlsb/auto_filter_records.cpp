#include "io/xlsb/auto_filter_records.h"

#include <cmath>
#include <cstdlib>
#include <optional>
#include <string>
#include <utility>

#include "io/xlsb/record_writer.h"
#include "sheet.h"
#include "utils/double_format.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

enum : std::uint16_t {
  kBrtBeginAFilter = 161,
  kBrtEndAFilter = 162,
  kBrtBeginFilterColumn = 163,
  kBrtEndFilterColumn = 164,
  kBrtBeginFilters = 165,
  kBrtEndFilters = 166,
  kBrtFilter = 167,
  kBrtColorFilter = 168,
  kBrtIconFilter = 169,
  kBrtTop10Filter = 170,
  kBrtDynamicFilter = 171,
  kBrtBeginCustomFilters = 172,
  kBrtEndCustomFilters = 173,
  kBrtCustomFilter = 174,
  kBrtAFilterDateGroupItem = 175,
  kBrtBeginSortState = 530,
  kBrtEndSortState = 531,
  kBrtBeginSortCond = 532,
  kBrtEndSortCond = 533,
};

// BrtCustomFilter vts values.
constexpr std::uint8_t kVtsNumber = 0x04;
constexpr std::uint8_t kVtsString = 0x06;

constexpr std::uint32_t kNoDxf = 0xFFFFFFFFU;
constexpr std::uint32_t kKpiNil = 0xFFFFFFFFU;
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

std::optional<double> ParseNumber(const std::string& text) {
  if (text.empty()) {
    return std::nullopt;
  }
  char* end = nullptr;
  const double value = std::strtod(text.c_str(), &end);
  if (end != text.c_str() + text.size() || !std::isfinite(value)) {
    return std::nullopt;
  }
  return value;
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
    default:
      // Booleans (0x08) and the blank / non-blank forms (0x0C / 0x0E) have no
      // OOXML spelling the model keeps apart from a string criterion.
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
  out->year = yr;
  out->month = static_cast<std::uint8_t>(mon);
  out->day = static_cast<std::uint8_t>(dom);
  out->hour = static_cast<std::uint8_t>(hour);
  out->minute = static_cast<std::uint8_t>(min);
  out->second = static_cast<std::uint8_t>(sec);
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
        col.kind = FilterKind::kValues;
        col.values.blank = blank != 0U;
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
        col.icon.icon_id = icon == 0xFFFFFFFFU ? std::nullopt : std::optional<std::uint32_t>(icon);
        break;
      }
      default:
        return Unsupported(rec.type);
    }
  }
}

void EmitCustomFilter(std::vector<std::uint8_t>& out, const CustomFilter& filter) {
  std::vector<std::uint8_t> payload;
  const std::optional<double> number = ParseNumber(filter.val);
  emit_u8(payload, number ? kVtsNumber : kVtsString);
  emit_u8(payload, SgnFromOperator(filter.op));
  emit_double(payload, number.value_or(0.0));
  if (!number) {
    emit_xlwidestring(payload, filter.val);
  }
  emit_record(out, kBrtCustomFilter, payload);
}

void EmitDateGroup(std::vector<std::uint8_t>& out, const DateGroupItem& item) {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, item.year);
  emit_u16(payload, item.month);
  emit_u32(payload, item.day);
  emit_u16(payload, item.hour);
  emit_u16(payload, item.minute);
  emit_u16(payload, item.second);
  emit_u16(payload, 0);
  emit_u32(payload, 0);
  emit_u32(payload, static_cast<std::uint32_t>(item.grouping));
  emit_record(out, kBrtAFilterDateGroupItem, payload);
}

std::uint32_t DynamicCft(DynamicFilterType type) {
  const auto index = static_cast<std::uint32_t>(type);
  return index <= 2U ? index : index - 3U + kCftDateBase;
}

/// Emits one column's criterion; returns the settings it could not hold.
std::uint32_t EmitCriterion(std::vector<std::uint8_t>& out, const FilterColumn& col) {
  std::vector<std::uint8_t> payload;
  switch (col.kind) {
    case FilterKind::kNone:
      return 0U;
    case FilterKind::kValues: {
      emit_u32(payload, col.values.blank ? 1U : 0U);
      emit_u32(payload, 0U);
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
      return col.values.calendar_type.empty() ? 0U : 1U;
    }
    case FilterKind::kCustom:
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
      flags |= col.top10.filter_val ? 0x04U : 0U;
      emit_u8(payload, static_cast<std::uint8_t>(flags));
      emit_double(payload, col.top10.val);
      emit_double(payload, col.top10.filter_val.value_or(0.0));
      emit_record(out, kBrtTop10Filter, payload);
      return 0U;
    }
    case FilterKind::kDynamic: {
      const std::uint32_t cft = DynamicCft(col.dynamic.type);
      const bool applied = col.dynamic.val.has_value() && cft != 0U && cft < kCftFirstPeriod;
      emit_u32(payload, cft);
      emit_u8(payload, applied ? 0x01U : 0U);
      emit_double(payload, applied ? *col.dynamic.val : 0.0);
      emit_double(payload, applied && cft >= kCftDateBase ? col.dynamic.max_val.value_or(0.0) : 0.0);
      emit_record(out, kBrtDynamicFilter, payload);
      return 0U;
    }
    case FilterKind::kColor:
      if (!col.color.dxf_id) {
        return 1U;
      }
      emit_u32(payload, *col.color.dxf_id);
      emit_u32(payload, col.color.cell_color ? 1U : 0U);
      emit_record(out, kBrtColorFilter, payload);
      return 0U;
    case FilterKind::kIcon:
      emit_u32(payload, static_cast<std::uint32_t>(col.icon.icon_set));
      emit_u32(payload, col.icon.icon_id.value_or(0xFFFFFFFFU));
      emit_record(out, kBrtIconFilter, payload);
      return 0U;
  }
  return 0U;
}

bool InGrid(const MergeRange& r) {
  return r.first_row <= r.last_row && r.first_col <= r.last_col && r.last_row < Sheet::kMaxRows &&
         r.last_col < Sheet::kMaxCols;
}

Expected<SortCondition, Error> ReadSortCond(ByteSpan payload) {
  SortCondition cond;
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(payload));
  ASSIGN_OR_RETURN(const MergeRange rfx, read_rfx(payload));
  ASSIGN_OR_RETURN(const std::uint32_t first, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t second, read_u32(payload));
  ASSIGN_OR_RETURN(auto list, read_xlnullablewidestring(payload));
  const std::uint32_t sort_on = (flags >> 1U) & 0x0FU;
  if (sort_on > static_cast<std::uint32_t>(SortBy::kIcon) || !InGrid(rfx)) {
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
    cond.icon_id = second == 0xFFFFFFFFU ? std::nullopt : std::optional<std::uint32_t>(second);
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
  if ((flags & 0x0004U) != 0U) {
    // fAltMethod does not say which of pinYin / stroke the sort used.
    return Unsupported(kBrtBeginSortState);
  }
  if (!InGrid(rfx)) {
    return Corrupt("xlsb BrtBeginSortState range is inverted or outside the grid");
  }
  sort.ref = rfx;
  sort.column_sort = (flags & 0x0001U) != 0U;
  sort.case_sensitive = (flags & 0x0002U) != 0U;
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
  flags |= sort.sort_method != SortMethod::kNone ? 0x0004U : 0U;
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
      emit_u32(payload, cond.icon_id.value_or(0xFFFFFFFFU));
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
  return (sort.sort_method != SortMethod::kNone ? 1U : 0U) + (sort.ext_lst_xml.empty() ? 0U : 1U);
}

Expected<AutoFilter, Error> read_auto_filter_block(ByteSpan& cursor, ByteSpan begin) {
  AutoFilter filter;
  ASSIGN_OR_RETURN(const MergeRange range, read_rfx(begin));
  if (range.first_row > range.last_row || range.first_col > range.last_col || range.last_row >= Sheet::kMaxRows ||
      range.last_col >= Sheet::kMaxCols) {
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

std::uint32_t emit_auto_filter_block(std::vector<std::uint8_t>& out, const AutoFilter& filter) {
  if (filter.is_opaque()) {
    return 1U;
  }
  std::uint32_t lost = 0;
  std::vector<std::uint8_t> payload;
  emit_rfx(payload, filter.range);
  emit_record(out, kBrtBeginAFilter, payload);
  for (const FilterColumn& col : filter.columns) {
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
  emit_record(out, kBrtEndAFilter, ByteSpan{});
  lost += filter.ext_lst_xml.empty() ? 0U : 1U;
  return lost;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
