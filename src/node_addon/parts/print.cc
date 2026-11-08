// Print pagination and print-settings authoring for the native Node addon.
//
// Method names and returned object shapes are byte-identical to the embind
// surface; `tools/dev/check_binding_drift.py` fails if one side gains a
// method the other lacks, which is what keeps a consumer able to swap
// `@libraz/formulon` for the native package without touching its code.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/addon_common.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

/// Reads an optional string property as the C tri-state: absent stays
/// `nullptr` (leave that section alone), present owns its bytes in
/// `storage` so the borrowed pointer outlives the call.
const char* PullOptionalString(CheckedSpecReader& reader, const Napi::Object& spec, const char* key,
                               std::string& storage) {
  if (!reader.String(spec, key, &storage)) {
    return nullptr;
  }
  return storage.c_str();
}

void PullOptionalDouble(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, double dflt,
                        double& value, int32_t& engaged) {
  bool present = false;
  value = reader.Double(spec, key, dflt, &present);
  engaged = present ? 1 : 0;
}

void PullOptionalU32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, uint32_t dflt,
                     uint32_t& value, int32_t& engaged) {
  bool present = false;
  value = reader.U32(spec, key, dflt, &present);
  engaged = present ? 1 : 0;
}

void PullOptionalBool(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, bool dflt, int32_t& value,
                      int32_t& engaged) {
  bool present = false;
  value = reader.Bool(spec, key, dflt, &present) ? 1 : 0;
  engaged = present ? 1 : 0;
}

// The span and manual flag of a break record `{ first?, last?, manual? }`; an omitted span is the whole sheet
// and an omitted `manual` is true.
struct BreakSpan {
  uint32_t first = 0;
  uint32_t last = 0;
  int32_t manual = 1;
};

BreakSpan ReadBreakSpan(CheckedSpecReader& reader, const Napi::Object& spec) {
  BreakSpan span;
  span.first = reader.U32(spec, "first", 0U);
  span.last = reader.U32(spec, "last", 0U);
  span.manual = reader.Bool(spec, "manual", true) ? 1 : 0;
  return span;
}

}  // namespace

Napi::Value Workbook::SheetStringGetter(const Napi::CallbackInfo& info, SheetStringGetFn getter) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* text = nullptr;
  const fm_status_t rc = getter(handle_, ArgU32(info, 0), &text);
  return ValueOrThrow(env, rc, JsStringOrNull(env, text));
}

Napi::Value Workbook::SheetStringSetter(const Napi::CallbackInfo& info, SheetStringSetFn setter) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string text = ArgString(info, 1);
    return setter(handle, ArgU32(info, 0), text.c_str());
  });
}

Napi::Value Workbook::GetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_page_setup_xml);
}

Napi::Value Workbook::SetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_page_setup_xml);
}

Napi::Value Workbook::GetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_page_margins_xml);
}

Napi::Value Workbook::SetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_page_margins_xml);
}

Napi::Value Workbook::GetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_print_options_xml);
}

Napi::Value Workbook::SetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_print_options_xml);
}

Napi::Value Workbook::GetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_header_footer_xml);
}

Napi::Value Workbook::SetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_header_footer_xml);
}

Napi::Value Workbook::GetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_sheet_pr_xml);
}

Napi::Value Workbook::SetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_sheet_pr_xml);
}

// `null` when the sheet has no `<tabColor>` (the C ABI reports kind None).
Napi::Value Workbook::GetSheetTabColor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_color_spec_t color{};
  const fm_status_t rc = fm_sheet_get_tab_color(handle_, ArgU32(info, 0), &color);
  return ValueOrThrow(env, rc, ColorSpecOrNull(env, color));
}

// `null` clears the tab colour; an object is read as an `fm_color_spec_t`.
Napi::Value Workbook::SetSheetTabColor(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    if (info[1].IsNull()) {
      return fm_sheet_set_tab_color(handle, sheet, nullptr);
    }
    CheckedSpecReader reader(info.Env());
    const fm_color_spec_t color = ReadColorSpec(reader, info[1], "color");
    if (!reader.ok()) {
      return fm_status_t{0};
    }
    return fm_sheet_set_tab_color(handle, sheet, &color);
  });
}

Napi::Value Workbook::GetSheetPrintArea(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const fm_cell_rect_t* found = nullptr;
  size_t count = 0;
  const fm_status_t rc = fm_sheet_get_print_area(handle_, ArgU32(info, 0), &found, &count);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array ranges = Napi::Array::New(env);
  for (size_t i = 0; i < count; ++i) {
    ranges.Set(static_cast<uint32_t>(i), RangeToJs(env, found[i]));
  }
  return ranges;
}

Napi::Value Workbook::SetSheetPrintArea(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  std::vector<fm_cell_rect_t> ranges;
  if (!ReadCellRects(reader, info[1], "ranges", &ranges)) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_print_area(handle_, ArgU32(info, 0), ranges.data(), ranges.size()),
                      env.Undefined());
}

Napi::Value Workbook::GetSheetPrintTitles(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  const fm_status_t rc = fm_sheet_get_print_titles(handle_, ArgU32(info, 0), &titles);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  return PrintTitlesToJs(env, titles);
}

Napi::Value Workbook::SetSheetPrintTitles(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  // `null` removes both axes.
  if (info[1].IsNull()) {
    return ValueOrThrow(env, fm_sheet_set_print_titles(handle_, sheet, nullptr), env.Undefined());
  }
  CheckedSpecReader reader(env);
  Napi::Object titles;
  if (!reader.Object(info[1], "titles", &titles)) {
    return env.Undefined();
  }
  fm_print_titles_t in = FM_STRUCT_INIT(fm_print_titles_t);
  Napi::Object rows;
  if (reader.Object(titles, "rows", &rows)) {
    in.has_rows = 1;
    in.first_row = reader.RequiredU32(rows, "first");
    in.last_row = reader.RequiredU32(rows, "last");
  }
  Napi::Object columns;
  if (reader.Object(titles, "columns", &columns)) {
    in.has_columns = 1;
    in.first_col = reader.RequiredU32(columns, "first");
    in.last_col = reader.RequiredU32(columns, "last");
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_print_titles(handle_, sheet, &in), env.Undefined());
}

// `info[2]` is the break record, or `null` to remove the break at `info[1]`.
Napi::Value Workbook::SetRowBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  if (info[2].IsNull()) {
    return ValueOrThrow(env, fm_sheet_set_row_break(handle_, sheet, row, nullptr), env.Undefined());
  }
  CheckedSpecReader reader(env);
  Napi::Object spec;
  if (!reader.Object(info[2], "break", &spec)) {
    return env.Undefined();
  }
  const BreakSpan span = ReadBreakSpan(reader, spec);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.row = row;
  brk.first = span.first;
  brk.last = span.last;
  brk.manual = span.manual;
  return ValueOrThrow(env, fm_sheet_set_row_break(handle_, sheet, row, &brk), env.Undefined());
}

Napi::Value Workbook::SetColumnBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t col = ArgU32(info, 1);
  if (info[2].IsNull()) {
    return ValueOrThrow(env, fm_sheet_set_column_break(handle_, sheet, col, nullptr), env.Undefined());
  }
  CheckedSpecReader reader(env);
  Napi::Object spec;
  if (!reader.Object(info[2], "break", &spec)) {
    return env.Undefined();
  }
  const BreakSpan span = ReadBreakSpan(reader, spec);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  brk.col = col;
  brk.first = span.first;
  brk.last = span.last;
  brk.manual = span.manual;
  return ValueOrThrow(env, fm_sheet_set_column_break(handle_, sheet, col, &brk), env.Undefined());
}

Napi::Value Workbook::ClearSheetBreaks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env, fm_sheet_clear_breaks(handle_, ArgU32(info, 0)), env.Undefined());
}

Napi::Value Workbook::BreaksArray(const Napi::CallbackInfo& info, bool rows) {
  Napi::Env env = info.Env();
  Napi::Array items = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  std::size_t count = 0;
  const fm_status_t count_rc = rows ? fm_sheet_get_row_break_count(handle_, sheet, &count)
                                    : fm_sheet_get_column_break_count(handle_, sheet, &count);
  if (count_rc != 0) {
    return ThrowStatus(env, count_rc);
  }
  for (std::size_t i = 0; i < count; ++i) {
    Napi::Object item = Napi::Object::New(env);
    if (rows) {
      fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
      if (const fm_status_t rc = fm_sheet_get_row_break_at(handle_, sheet, i, &brk); rc != 0) {
        return ThrowStatus(env, rc);
      }
      item.Set("row", Napi::Number::New(env, brk.row));
      item.Set("first", Napi::Number::New(env, brk.first));
      item.Set("last", Napi::Number::New(env, brk.last));
      item.Set("manual", Napi::Boolean::New(env, brk.manual != 0));
    } else {
      fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
      if (const fm_status_t rc = fm_sheet_get_column_break_at(handle_, sheet, i, &brk); rc != 0) {
        return ThrowStatus(env, rc);
      }
      item.Set("col", Napi::Number::New(env, brk.col));
      item.Set("first", Napi::Number::New(env, brk.first));
      item.Set("last", Napi::Number::New(env, brk.last));
      item.Set("manual", Napi::Boolean::New(env, brk.manual != 0));
    }
    items.Set(i, item);
  }
  return items;
}

Napi::Value Workbook::GetSheetRowBreaks(const Napi::CallbackInfo& info) {
  return BreaksArray(info, /*rows=*/true);
}

Napi::Value Workbook::GetSheetColBreaks(const Napi::CallbackInfo& info) {
  return BreaksArray(info, /*rows=*/false);
}

Napi::Value Workbook::UpdateSheetPageSetup(const Napi::CallbackInfo& info) {
  fm_page_setup_t out = FM_STRUCT_INIT(fm_page_setup_t);
  return InvokeSheetSpec(
      info, "updatePageSetup expects (sheet:number, setup:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        uint32_t orientation = 0U;
        PullOptionalU32(reader, spec, "orientation", 0U, orientation, out.orientation_engaged);
        // Past int32 the ordinal wraps negative, which the C ABI rejects like any unknown one.
        out.orientation = static_cast<int32_t>(orientation);
        PullOptionalU32(reader, spec, "paperSize", 0U, out.paper_size, out.paper_size_engaged);
        PullOptionalU32(reader, spec, "scalePct", 0U, out.scale_pct, out.scale_pct_engaged);
        PullOptionalU32(reader, spec, "fitToWidth", 0U, out.fit_to_width, out.fit_to_width_engaged);
        PullOptionalU32(reader, spec, "fitToHeight", 0U, out.fit_to_height, out.fit_to_height_engaged);
        PullOptionalBool(reader, spec, "fitToPage", false, out.fit_to_page, out.fit_to_page_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_update_page_setup(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::UpdateSheetPageMargins(const Napi::CallbackInfo& info) {
  fm_page_margins_t out = FM_STRUCT_INIT(fm_page_margins_t);
  return InvokeSheetSpec(
      info, "updatePageMargins expects (sheet:number, margins:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        PullOptionalDouble(reader, spec, "leftIn", 0.0, out.left_in, out.left_in_engaged);
        PullOptionalDouble(reader, spec, "rightIn", 0.0, out.right_in, out.right_in_engaged);
        PullOptionalDouble(reader, spec, "topIn", 0.0, out.top_in, out.top_in_engaged);
        PullOptionalDouble(reader, spec, "bottomIn", 0.0, out.bottom_in, out.bottom_in_engaged);
        PullOptionalDouble(reader, spec, "headerIn", 0.0, out.header_in, out.header_in_engaged);
        PullOptionalDouble(reader, spec, "footerIn", 0.0, out.footer_in, out.footer_in_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_update_page_margins(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::UpdateSheetPrintOptions(const Napi::CallbackInfo& info) {
  fm_print_options_t out = FM_STRUCT_INIT(fm_print_options_t);
  return InvokeSheetSpec(
      info, "updatePrintOptions expects (sheet:number, options:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        PullOptionalBool(reader, spec, "gridLines", false, out.grid_lines, out.grid_lines_engaged);
        PullOptionalBool(reader, spec, "headings", false, out.headings, out.headings_engaged);
        PullOptionalBool(reader, spec, "horizontalCentered", false, out.horizontal_centered,
                         out.horizontal_centered_engaged);
        PullOptionalBool(reader, spec, "verticalCentered", false, out.vertical_centered, out.vertical_centered_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_update_print_options(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::UpdateSheetHeaderFooter(const Napi::CallbackInfo& info) {
  // The six section strings are borrowed by the C struct, so their
  // storage has to outlive the call.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  fm_header_footer_t out = FM_STRUCT_INIT(fm_header_footer_t);
  return InvokeSheetSpec(
      info, "updateHeaderFooter expects (sheet:number, headerFooter:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        out.odd_header = PullOptionalString(reader, spec, "oddHeader", odd_header);
        out.odd_footer = PullOptionalString(reader, spec, "oddFooter", odd_footer);
        out.even_header = PullOptionalString(reader, spec, "evenHeader", even_header);
        out.even_footer = PullOptionalString(reader, spec, "evenFooter", even_footer);
        out.first_header = PullOptionalString(reader, spec, "firstHeader", first_header);
        out.first_footer = PullOptionalString(reader, spec, "firstFooter", first_footer);
        PullOptionalBool(reader, spec, "differentOddEven", false, out.different_odd_even,
                         out.different_odd_even_engaged);
        PullOptionalBool(reader, spec, "differentFirst", false, out.different_first, out.different_first_engaged);
        PullOptionalBool(reader, spec, "scaleWithDoc", false, out.scale_with_doc, out.scale_with_doc_engaged);
        PullOptionalBool(reader, spec, "alignWithMargins", false, out.align_with_margins,
                         out.align_with_margins_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_update_header_footer(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::GetSheetPageSetup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  const fm_status_t rc = fm_sheet_get_page_setup(handle_, ArgU32(info, 0), &setup);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object result = Napi::Object::New(env);
  result.Set("orientation", JsNumberOrNull(env, setup.orientation_engaged, setup.orientation));
  result.Set("paperSize", Napi::Number::New(env, setup.paper_size));
  result.Set("scalePct", Napi::Number::New(env, setup.scale_pct));
  result.Set("fitToWidth", Napi::Number::New(env, setup.fit_to_width));
  result.Set("fitToHeight", Napi::Number::New(env, setup.fit_to_height));
  result.Set("fitToPage", Napi::Boolean::New(env, setup.fit_to_page != 0));
  result.Set("orientationStated", Napi::Boolean::New(env, setup.orientation_engaged != 0));
  result.Set("paperSizeStated", Napi::Boolean::New(env, setup.paper_size_engaged != 0));
  result.Set("scalePctStated", Napi::Boolean::New(env, setup.scale_pct_engaged != 0));
  result.Set("fitToWidthStated", Napi::Boolean::New(env, setup.fit_to_width_engaged != 0));
  result.Set("fitToHeightStated", Napi::Boolean::New(env, setup.fit_to_height_engaged != 0));
  result.Set("fitToPageStated", Napi::Boolean::New(env, setup.fit_to_page_engaged != 0));
  return result;
}

Napi::Value Workbook::GetSheetPageMargins(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_page_margins_t margins = FM_STRUCT_INIT(fm_page_margins_t);
  const fm_status_t rc = fm_sheet_get_page_margins(handle_, ArgU32(info, 0), &margins);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object result = Napi::Object::New(env);
  result.Set("leftIn", Napi::Number::New(env, margins.left_in));
  result.Set("rightIn", Napi::Number::New(env, margins.right_in));
  result.Set("topIn", Napi::Number::New(env, margins.top_in));
  result.Set("bottomIn", Napi::Number::New(env, margins.bottom_in));
  result.Set("headerIn", Napi::Number::New(env, margins.header_in));
  result.Set("footerIn", Napi::Number::New(env, margins.footer_in));
  result.Set("leftInStated", Napi::Boolean::New(env, margins.left_in_engaged != 0));
  result.Set("rightInStated", Napi::Boolean::New(env, margins.right_in_engaged != 0));
  result.Set("topInStated", Napi::Boolean::New(env, margins.top_in_engaged != 0));
  result.Set("bottomInStated", Napi::Boolean::New(env, margins.bottom_in_engaged != 0));
  result.Set("headerInStated", Napi::Boolean::New(env, margins.header_in_engaged != 0));
  result.Set("footerInStated", Napi::Boolean::New(env, margins.footer_in_engaged != 0));
  return result;
}

namespace {

// Count reported by a pagination count getter, which cannot fail on a live
// handle with a non-null out-param.
std::size_t PaginationCount(const fm_pagination_t* pagination,
                            fm_status_t (*count_fn)(const fm_pagination_t*, std::size_t*)) {
  std::size_t count = 0;
  (void)count_fn(pagination, &count);
  return count;
}

Napi::Object RectToJs(Napi::Env env, const fm_rect_pt_t& r) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("x", Napi::Number::New(env, r.x));
  out.Set("y", Napi::Number::New(env, r.y));
  out.Set("widthPt", Napi::Number::New(env, r.width_pt));
  out.Set("heightPt", Napi::Number::New(env, r.height_pt));
  return out;
}

// Sets the pagination detail keys on `result`.
void SetPaginationDetail(Napi::Env env, Napi::Object result, const fm_pagination_t* pagination) {
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  fm_margins_pt_t margins = FM_STRUCT_INIT(fm_margins_pt_t);
  fm_rect_pt_t printable = FM_STRUCT_INIT(fm_rect_pt_t);
  double scale_pct = 0.0;
  int32_t page_order = 0;
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  Napi::Array pages = Napi::Array::New(env);
  (void)fm_pagination_get_paper(pagination, &paper);
  (void)fm_pagination_get_margins(pagination, &margins);
  (void)fm_pagination_get_printable(pagination, &printable);
  (void)fm_pagination_get_scale(pagination, &scale_pct);
  (void)fm_pagination_get_page_order(pagination, &page_order);
  (void)fm_pagination_get_print_titles(pagination, &titles);
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_page_count); ++i) {
    fm_page_layout_t page = FM_STRUCT_INIT(fm_page_layout_t);
    if (fm_pagination_get_page_at(pagination, i, &page) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("areaIndex", Napi::Number::New(env, page.area_index));
    item.Set("rect", RangeToJs(env, page.rect));
    item.Set("originXPt", Napi::Number::New(env, page.origin_x_pt));
    item.Set("originYPt", Napi::Number::New(env, page.origin_y_pt));
    item.Set("widthPt", Napi::Number::New(env, page.width_pt));
    item.Set("heightPt", Napi::Number::New(env, page.height_pt));
    pages.Set(i, item);
  }
  Napi::Object paper_js = Napi::Object::New(env);
  paper_js.Set("widthPt", Napi::Number::New(env, paper.width_pt));
  paper_js.Set("heightPt", Napi::Number::New(env, paper.height_pt));
  paper_js.Set("landscape", Napi::Boolean::New(env, paper.landscape != 0));
  paper_js.Set("known", Napi::Boolean::New(env, paper.known != 0));
  Napi::Object margins_js = Napi::Object::New(env);
  margins_js.Set("leftIn", Napi::Number::New(env, margins.left_in));
  margins_js.Set("rightIn", Napi::Number::New(env, margins.right_in));
  margins_js.Set("topIn", Napi::Number::New(env, margins.top_in));
  margins_js.Set("bottomIn", Napi::Number::New(env, margins.bottom_in));
  margins_js.Set("headerIn", Napi::Number::New(env, margins.header_in));
  margins_js.Set("footerIn", Napi::Number::New(env, margins.footer_in));
  result.Set("paper", paper_js);
  result.Set("margins", margins_js);
  result.Set("printable", RectToJs(env, printable));
  result.Set("scalePct", Napi::Number::New(env, scale_pct));
  result.Set("pageOrder", Napi::Number::New(env, page_order));
  result.Set("printTitles", PrintTitlesToJs(env, titles));
  result.Set("pages", pages);
}

}  // namespace

namespace {

// The span a break covers on the perpendicular axis when it is computed rather than stored:
// the whole sheet.
constexpr uint32_t kLastColumnIndex = 16383U;
constexpr uint32_t kLastRowIndex = 1048575U;

}  // namespace

Napi::Value Workbook::Paginate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  if (info.Length() < 1 || !info[0].IsNumber()) {
    Napi::TypeError::New(env, "getPagination expects (sheet:number)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  fm_pagination_t* pagination = nullptr;
  const fm_status_t rc = fm_sheet_get_pagination(handle_, ArgU32(info, 0), &pagination);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object result = Napi::Object::New(env);
  Napi::Array print_area = Napi::Array::New(env);
  Napi::Array row_breaks = Napi::Array::New(env);
  Napi::Array column_breaks = Napi::Array::New(env);
  const uint32_t sheet = ArgU32(info, 0);
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_print_area_count); ++i) {
    fm_cell_rect_t range{};
    if (fm_pagination_get_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    print_area.Set(i, RangeToJs(env, range));
  }
  // A stored break reports its own span; a break the page walk computed covers the whole sheet.
  std::size_t stored_rows = 0;
  std::size_t stored_columns = 0;
  (void)fm_sheet_get_row_break_count(handle_, sheet, &stored_rows);
  (void)fm_sheet_get_column_break_count(handle_, sheet, &stored_columns);
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_row_break_count); ++i) {
    uint32_t row = 0;
    int32_t manual = 0;
    if (fm_pagination_get_row_break_at(pagination, i, &row) != 0 ||
        fm_pagination_is_row_break_manual(pagination, i, &manual) != 0) {
      continue;
    }
    uint32_t first = 0;
    uint32_t last = kLastColumnIndex;
    for (std::size_t k = 0; k < stored_rows; ++k) {
      fm_row_break_t stored = FM_STRUCT_INIT(fm_row_break_t);
      if (fm_sheet_get_row_break_at(handle_, sheet, k, &stored) == 0 && stored.row == row) {
        first = stored.first;
        last = stored.last;
        break;
      }
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("row", Napi::Number::New(env, row));
    item.Set("first", Napi::Number::New(env, first));
    item.Set("last", Napi::Number::New(env, last));
    item.Set("manual", Napi::Boolean::New(env, manual != 0));
    row_breaks.Set(row_breaks.Length(), item);
  }
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_column_break_count); ++i) {
    uint32_t col = 0;
    int32_t manual = 0;
    if (fm_pagination_get_column_break_at(pagination, i, &col) != 0 ||
        fm_pagination_is_column_break_manual(pagination, i, &manual) != 0) {
      continue;
    }
    uint32_t first = 0;
    uint32_t last = kLastRowIndex;
    for (std::size_t k = 0; k < stored_columns; ++k) {
      fm_column_break_t stored = FM_STRUCT_INIT(fm_column_break_t);
      if (fm_sheet_get_column_break_at(handle_, sheet, k, &stored) == 0 && stored.col == col) {
        first = stored.first;
        last = stored.last;
        break;
      }
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("col", Napi::Number::New(env, col));
    item.Set("first", Napi::Number::New(env, first));
    item.Set("last", Napi::Number::New(env, last));
    item.Set("manual", Napi::Boolean::New(env, manual != 0));
    column_breaks.Set(column_breaks.Length(), item);
  }
  result.Set("printArea", print_area);
  result.Set("rowBreaks", row_breaks);
  result.Set("columnBreaks", column_breaks);
  result.Set("pageCount",
             Napi::Number::New(env, static_cast<double>(PaginationCount(pagination, &fm_pagination_get_page_count))));
  SetPaginationDetail(env, result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

Napi::Value Workbook::GetSheetPrintOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_print_options_t options = FM_STRUCT_INIT(fm_print_options_t);
  const fm_status_t rc = fm_sheet_get_print_options(handle_, ArgU32(info, 0), &options);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object result = Napi::Object::New(env);
  result.Set("gridLines", Napi::Boolean::New(env, options.grid_lines != 0));
  result.Set("headings", Napi::Boolean::New(env, options.headings != 0));
  result.Set("horizontalCentered", Napi::Boolean::New(env, options.horizontal_centered != 0));
  result.Set("verticalCentered", Napi::Boolean::New(env, options.vertical_centered != 0));
  result.Set("gridLinesStated", Napi::Boolean::New(env, options.grid_lines_engaged != 0));
  result.Set("headingsStated", Napi::Boolean::New(env, options.headings_engaged != 0));
  result.Set("horizontalCenteredStated", Napi::Boolean::New(env, options.horizontal_centered_engaged != 0));
  result.Set("verticalCenteredStated", Napi::Boolean::New(env, options.vertical_centered_engaged != 0));
  return result;
}

Napi::Value Workbook::GetSheetHeaderFooter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_header_footer_t hf = FM_STRUCT_INIT(fm_header_footer_t);
  const fm_status_t rc = fm_sheet_get_header_footer(handle_, ArgU32(info, 0), &hf);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object result = Napi::Object::New(env);
  result.Set("oddHeader", JsStringOrNull(env, hf.odd_header));
  result.Set("oddFooter", JsStringOrNull(env, hf.odd_footer));
  result.Set("evenHeader", JsStringOrNull(env, hf.even_header));
  result.Set("evenFooter", JsStringOrNull(env, hf.even_footer));
  result.Set("firstHeader", JsStringOrNull(env, hf.first_header));
  result.Set("firstFooter", JsStringOrNull(env, hf.first_footer));
  result.Set("differentOddEven", Napi::Boolean::New(env, hf.different_odd_even != 0));
  result.Set("differentFirst", Napi::Boolean::New(env, hf.different_first != 0));
  result.Set("scaleWithDoc", Napi::Boolean::New(env, hf.scale_with_doc != 0));
  result.Set("alignWithMargins", Napi::Boolean::New(env, hf.align_with_margins != 0));
  result.Set("differentOddEvenStated", Napi::Boolean::New(env, hf.different_odd_even_engaged != 0));
  result.Set("differentFirstStated", Napi::Boolean::New(env, hf.different_first_engaged != 0));
  result.Set("scaleWithDocStated", Napi::Boolean::New(env, hf.scale_with_doc_engaged != 0));
  result.Set("alignWithMarginsStated", Napi::Boolean::New(env, hf.align_with_margins_engaged != 0));
  return result;
}

}  // namespace formulon_node
