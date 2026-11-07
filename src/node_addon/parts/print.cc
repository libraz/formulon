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

// Shared body of the row / column page-break add entries.
using BreakAddFn = fm_status_t (*)(fm_workbook_t*, size_t, uint32_t, int32_t);

Napi::Value InvokeBreakAdd(Workbook& wb, const Napi::CallbackInfo& info, BreakAddFn fn) {
  return wb.StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
    const uint32_t index = Workbook::ArgU32(info, 1);
    return fn(handle, sheet, index, Workbook::ArgFlag(info, 2));
  });
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
    in.has_cols = 1;
    in.first_col = reader.RequiredU32(columns, "first");
    in.last_col = reader.RequiredU32(columns, "last");
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_print_titles(handle_, sheet, &in), env.Undefined());
}

Napi::Value Workbook::AddSheetRowBreak(const Napi::CallbackInfo& info) {
  return InvokeBreakAdd(*this, info, &fm_sheet_add_row_break);
}

Napi::Value Workbook::AddSheetColBreak(const Napi::CallbackInfo& info) {
  return InvokeBreakAdd(*this, info, &fm_sheet_add_column_break);
}

Napi::Value Workbook::RemoveSheetRowBreak(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_row_break);
}

Napi::Value Workbook::RemoveSheetColBreak(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_column_break);
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
    fm_page_break_t brk = FM_STRUCT_INIT(fm_page_break_t);
    const fm_status_t rc = rows ? fm_sheet_get_row_break_at(handle_, sheet, i, &brk)
                                : fm_sheet_get_column_break_at(handle_, sheet, i, &brk);
    if (rc != 0) {
      return ThrowStatus(env, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", Napi::Number::New(env, brk.id));
    item.Set("min", Napi::Number::New(env, brk.min));
    item.Set("max", Napi::Number::New(env, brk.max));
    item.Set("manual", Napi::Boolean::New(env, brk.manual != 0));
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

Napi::Value Workbook::SetSheetPageSetup(const Napi::CallbackInfo& info) {
  fm_page_setup_t out = FM_STRUCT_INIT(fm_page_setup_t);
  return InvokeSheetSpec(
      info, "setPageSetup expects (sheet:number, setup:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        uint32_t orientation = 0U;
        PullOptionalU32(reader, spec, "orientation", 0U, orientation, out.orientation_engaged);
        // Past int32 the ordinal wraps negative, which the C ABI rejects like any unknown one.
        out.orientation = static_cast<int32_t>(orientation);
        PullOptionalU32(reader, spec, "paperSize", 0U, out.paper_size, out.paper_size_engaged);
        PullOptionalU32(reader, spec, "scale", 0U, out.scale, out.scale_engaged);
        PullOptionalU32(reader, spec, "fitToWidth", 0U, out.fit_to_width, out.fit_to_width_engaged);
        PullOptionalU32(reader, spec, "fitToHeight", 0U, out.fit_to_height, out.fit_to_height_engaged);
        PullOptionalBool(reader, spec, "fitToPage", false, out.fit_to_page, out.fit_to_page_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_set_page_setup(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::SetSheetPageMargins(const Napi::CallbackInfo& info) {
  fm_page_margins_t out = FM_STRUCT_INIT(fm_page_margins_t);
  return InvokeSheetSpec(
      info, "setPageMargins expects (sheet:number, margins:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        PullOptionalDouble(reader, spec, "left", 0.0, out.left, out.left_engaged);
        PullOptionalDouble(reader, spec, "right", 0.0, out.right, out.right_engaged);
        PullOptionalDouble(reader, spec, "top", 0.0, out.top, out.top_engaged);
        PullOptionalDouble(reader, spec, "bottom", 0.0, out.bottom, out.bottom_engaged);
        PullOptionalDouble(reader, spec, "header", 0.0, out.header, out.header_engaged);
        PullOptionalDouble(reader, spec, "footer", 0.0, out.footer, out.footer_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_set_page_margins(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::SetSheetPrintOptions(const Napi::CallbackInfo& info) {
  fm_print_options_t out = FM_STRUCT_INIT(fm_print_options_t);
  return InvokeSheetSpec(
      info, "setPrintOptions expects (sheet:number, options:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        PullOptionalBool(reader, spec, "gridLines", false, out.grid_lines, out.grid_lines_engaged);
        PullOptionalBool(reader, spec, "headings", false, out.headings, out.headings_engaged);
        PullOptionalBool(reader, spec, "horizontalCentered", false, out.horizontal_centered,
                         out.horizontal_centered_engaged);
        PullOptionalBool(reader, spec, "verticalCentered", false, out.vertical_centered, out.vertical_centered_engaged);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_set_print_options(handle, ArgU32(info, 0), &out); });
}

Napi::Value Workbook::SetSheetHeaderFooter(const Napi::CallbackInfo& info) {
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
      info, "setHeaderFooter expects (sheet:number, headerFooter:object)",
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
      [&](fm_workbook_t* handle) { return fm_sheet_set_header_footer(handle, ArgU32(info, 0), &out); });
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
  result.Set("scale", Napi::Number::New(env, setup.scale));
  result.Set("fitToWidth", Napi::Number::New(env, setup.fit_to_width));
  result.Set("fitToHeight", Napi::Number::New(env, setup.fit_to_height));
  result.Set("fitToPage", Napi::Boolean::New(env, setup.fit_to_page != 0));
  result.Set("orientationStated", Napi::Boolean::New(env, setup.orientation_engaged != 0));
  result.Set("paperSizeStated", Napi::Boolean::New(env, setup.paper_size_engaged != 0));
  result.Set("scaleStated", Napi::Boolean::New(env, setup.scale_engaged != 0));
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
  result.Set("left", Napi::Number::New(env, margins.left));
  result.Set("right", Napi::Number::New(env, margins.right));
  result.Set("top", Napi::Number::New(env, margins.top));
  result.Set("bottom", Napi::Number::New(env, margins.bottom));
  result.Set("header", Napi::Number::New(env, margins.header));
  result.Set("footer", Napi::Number::New(env, margins.footer));
  result.Set("leftStated", Napi::Boolean::New(env, margins.left_engaged != 0));
  result.Set("rightStated", Napi::Boolean::New(env, margins.right_engaged != 0));
  result.Set("topStated", Napi::Boolean::New(env, margins.top_engaged != 0));
  result.Set("bottomStated", Napi::Boolean::New(env, margins.bottom_engaged != 0));
  result.Set("headerStated", Napi::Boolean::New(env, margins.header_engaged != 0));
  result.Set("footerStated", Napi::Boolean::New(env, margins.footer_engaged != 0));
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
  out.Set("width", Napi::Number::New(env, r.width));
  out.Set("height", Napi::Number::New(env, r.height));
  return out;
}

// Sets the pagination detail keys on `result`.
void SetPaginationDetail(Napi::Env env, Napi::Object result, const fm_pagination_t* pagination) {
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  fm_margins_pt_t margins = FM_STRUCT_INIT(fm_margins_pt_t);
  fm_rect_pt_t printable = FM_STRUCT_INIT(fm_rect_pt_t);
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  Napi::Array pages = Napi::Array::New(env);
  Napi::Array horizontal_manual = Napi::Array::New(env);
  Napi::Array vertical_manual = Napi::Array::New(env);
  (void)fm_pagination_get_paper(pagination, &paper);
  (void)fm_pagination_get_margins(pagination, &margins);
  (void)fm_pagination_get_printable(pagination, &printable);
  (void)fm_pagination_get_scale(pagination, &scale);
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
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_row_break_count); ++i) {
    int32_t manual = 0;
    if (fm_pagination_is_row_break_manual(pagination, i, &manual) == 0) {
      horizontal_manual.Set(i, Napi::Boolean::New(env, manual != 0));
    }
  }
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_column_break_count); ++i) {
    int32_t manual = 0;
    if (fm_pagination_is_column_break_manual(pagination, i, &manual) == 0) {
      vertical_manual.Set(i, Napi::Boolean::New(env, manual != 0));
    }
  }
  Napi::Object paper_js = Napi::Object::New(env);
  paper_js.Set("widthPt", Napi::Number::New(env, paper.width_pt));
  paper_js.Set("heightPt", Napi::Number::New(env, paper.height_pt));
  paper_js.Set("landscape", Napi::Boolean::New(env, paper.landscape != 0));
  paper_js.Set("known", Napi::Boolean::New(env, paper.known != 0));
  Napi::Object margins_js = Napi::Object::New(env);
  margins_js.Set("left", Napi::Number::New(env, margins.left));
  margins_js.Set("right", Napi::Number::New(env, margins.right));
  margins_js.Set("top", Napi::Number::New(env, margins.top));
  margins_js.Set("bottom", Napi::Number::New(env, margins.bottom));
  margins_js.Set("header", Napi::Number::New(env, margins.header));
  margins_js.Set("footer", Napi::Number::New(env, margins.footer));
  result.Set("paper", paper_js);
  result.Set("margins", margins_js);
  result.Set("printable", RectToJs(env, printable));
  result.Set("scale", Napi::Number::New(env, scale));
  result.Set("pageOrder", Napi::Number::New(env, page_order));
  result.Set("printTitles", PrintTitlesToJs(env, titles));
  result.Set("pages", pages);
  result.Set("rowBreakManual", horizontal_manual);
  result.Set("columnBreakManual", vertical_manual);
}

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
  Napi::Array horizontal_breaks = Napi::Array::New(env);
  Napi::Array vertical_breaks = Napi::Array::New(env);
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_print_area_count); ++i) {
    fm_cell_rect_t range{};
    if (fm_pagination_get_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    print_area.Set(i, RangeToJs(env, range));
  }
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_row_break_count); ++i) {
    uint32_t row = 0;
    if (fm_pagination_get_row_break_at(pagination, i, &row) == 0) {
      horizontal_breaks.Set(i, Napi::Number::New(env, row));
    }
  }
  for (std::size_t i = 0; i < PaginationCount(pagination, &fm_pagination_get_column_break_count); ++i) {
    uint32_t col = 0;
    if (fm_pagination_get_column_break_at(pagination, i, &col) == 0) {
      vertical_breaks.Set(i, Napi::Number::New(env, col));
    }
  }
  result.Set("printArea", print_area);
  result.Set("rowBreaks", horizontal_breaks);
  result.Set("columnBreaks", vertical_breaks);
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
