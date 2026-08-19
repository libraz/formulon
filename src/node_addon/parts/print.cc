// Print pagination and print-settings authoring for the native Node addon.
//
// Method names and returned object shapes are byte-identical to the embind
// surface; `tools/dev/check_binding_drift.py` fails if one side gains a
// method the other lacks, which is what keeps a consumer able to swap
// `@libraz/formulon` for the native package without touching its code.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/addon_common.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

/// Reads an optional string property as the C tri-state: absent stays
/// `nullptr` (leave that section alone), present owns its bytes in
/// `storage` so the borrowed pointer outlives the call.
const char* PullOptionalString(const Napi::Object& spec, const char* key, std::string& storage) {
  if (!SpecHas(spec, key)) {
    return nullptr;
  }
  storage = spec.Get(key).ToString().Utf8Value();
  return storage.c_str();
}

}  // namespace

Napi::Value Workbook::XmlFragmentGetter(const Napi::CallbackInfo& info, XmlFragmentGetFn getter) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "xml", nullptr);
  }
  const char* xml = nullptr;
  const fm_status_t rc = getter(handle_, ArgU32(info, 0), &xml);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "xml", nullptr);
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "xml", xml);
}

Napi::Value Workbook::XmlFragmentSetter(const Napi::CallbackInfo& info, XmlFragmentSetFn setter) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string xml = ArgString(info, 1);
  return MakeStatus(env, setter(handle_, ArgU32(info, 0), xml.c_str()));
}

Napi::Value Workbook::GetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return XmlFragmentGetter(info, &fm_sheet_get_page_setup_xml);
}

Napi::Value Workbook::SetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return XmlFragmentSetter(info, &fm_sheet_set_page_setup_xml);
}

Napi::Value Workbook::GetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return XmlFragmentGetter(info, &fm_sheet_get_page_margins_xml);
}

Napi::Value Workbook::SetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return XmlFragmentSetter(info, &fm_sheet_set_page_margins_xml);
}

Napi::Value Workbook::GetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return XmlFragmentGetter(info, &fm_sheet_get_print_options_xml);
}

Napi::Value Workbook::SetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return XmlFragmentSetter(info, &fm_sheet_set_print_options_xml);
}

Napi::Value Workbook::GetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return XmlFragmentGetter(info, &fm_sheet_get_header_footer_xml);
}

Napi::Value Workbook::SetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return XmlFragmentSetter(info, &fm_sheet_set_header_footer_xml);
}

Napi::Value Workbook::GetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return XmlFragmentGetter(info, &fm_sheet_get_sheet_pr_xml);
}

Napi::Value Workbook::SetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return XmlFragmentSetter(info, &fm_sheet_set_sheet_pr_xml);
}

Napi::Value Workbook::SetSheetFitToPage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_set_fit_to_page(handle_, ArgU32(info, 0), ArgBool(info, 1) ? 1 : 0));
}

Napi::Value Workbook::GetSheetPrintArea(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "ranges", nullptr);
  }
  const char* ranges = nullptr;
  const fm_status_t rc = fm_sheet_get_print_area(handle_, ArgU32(info, 0), &ranges);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "ranges", nullptr);
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "ranges", ranges);
}

Napi::Value Workbook::SetSheetPrintArea(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string ranges = ArgString(info, 1);
  return MakeStatus(env, fm_sheet_set_print_area(handle_, ArgU32(info, 0), ranges.c_str()));
}

Napi::Value Workbook::GetSheetPrintTitles(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    result.Set("repeatRows", Napi::String::New(env, ""));
    result.Set("repeatCols", Napi::String::New(env, ""));
    return result;
  }
  const char* rows = nullptr;
  const char* cols = nullptr;
  const fm_status_t rc = fm_sheet_get_print_titles(handle_, ArgU32(info, 0), &rows, &cols);
  if (rc != 0) {
    result.Set("status", MakeErrorStatus(env, rc));
    result.Set("repeatRows", Napi::String::New(env, ""));
    result.Set("repeatCols", Napi::String::New(env, ""));
    return result;
  }
  result.Set("status", MakeOkStatus(env));
  result.Set("repeatRows", Napi::String::New(env, rows != nullptr ? rows : ""));
  result.Set("repeatCols", Napi::String::New(env, cols != nullptr ? cols : ""));
  return result;
}

Napi::Value Workbook::SetSheetPrintTitles(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string rows = ArgString(info, 1);
  const std::string cols = ArgString(info, 2);
  return MakeStatus(env, fm_sheet_set_print_titles(handle_, ArgU32(info, 0), rows.c_str(), cols.c_str()));
}

Napi::Value Workbook::AddSheetRowBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_add_row_break(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgBool(info, 2) ? 1 : 0));
}

Napi::Value Workbook::AddSheetColBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_add_col_break(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgBool(info, 2) ? 1 : 0));
}

Napi::Value Workbook::RemoveSheetRowBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_remove_row_break(handle_, ArgU32(info, 0), ArgU32(info, 1)));
}

Napi::Value Workbook::RemoveSheetColBreak(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_remove_col_break(handle_, ArgU32(info, 0), ArgU32(info, 1)));
}

Napi::Value Workbook::ClearSheetBreaks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_clear_breaks(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::BreaksArray(const Napi::CallbackInfo& info, bool rows) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  Napi::Array items = Napi::Array::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    result.Set("breaks", items);
    return result;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const std::size_t count = rows ? fm_sheet_row_break_count(handle_, sheet) : fm_sheet_col_break_count(handle_, sheet);
  for (std::size_t i = 0; i < count; ++i) {
    fm_page_break_t brk{};
    const fm_status_t rc =
        rows ? fm_sheet_row_break_at(handle_, sheet, i, &brk) : fm_sheet_col_break_at(handle_, sheet, i, &brk);
    if (rc != 0) {
      result.Set("status", MakeErrorStatus(env, rc));
      result.Set("breaks", Napi::Array::New(env));
      return result;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", Napi::Number::New(env, brk.id));
    item.Set("min", Napi::Number::New(env, brk.min));
    item.Set("max", Napi::Number::New(env, brk.max));
    item.Set("manual", Napi::Boolean::New(env, brk.manual != 0));
    items.Set(i, item);
  }
  result.Set("status", MakeOkStatus(env));
  result.Set("breaks", items);
  return result;
}

Napi::Value Workbook::GetSheetRowBreaks(const Napi::CallbackInfo& info) {
  return BreaksArray(info, /*rows=*/true);
}

Napi::Value Workbook::GetSheetColBreaks(const Napi::CallbackInfo& info) {
  return BreaksArray(info, /*rows=*/false);
}

Napi::Value Workbook::SetSheetPageSetup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    Napi::TypeError::New(env, "setSheetPageSetup expects (sheet:number, setup:object)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  fm_page_setup_t out{};
  out.orientation_engaged = SpecHas(spec, "orientation") ? 1 : 0;
  out.orientation = SpecPullU32(spec, "orientation", 0U);
  out.paper_size_engaged = SpecHas(spec, "paperSize") ? 1 : 0;
  out.paper_size = SpecPullU32(spec, "paperSize", 0U);
  out.scale_engaged = SpecHas(spec, "scale") ? 1 : 0;
  out.scale = SpecPullU32(spec, "scale", 0U);
  out.fit_to_width_engaged = SpecHas(spec, "fitToWidth") ? 1 : 0;
  out.fit_to_width = SpecPullU32(spec, "fitToWidth", 0U);
  out.fit_to_height_engaged = SpecHas(spec, "fitToHeight") ? 1 : 0;
  out.fit_to_height = SpecPullU32(spec, "fitToHeight", 0U);
  out.fit_to_page_engaged = SpecHas(spec, "fitToPage") ? 1 : 0;
  out.fit_to_page = SpecPullBool(spec, "fitToPage", false) ? 1 : 0;
  return MakeStatus(env, fm_sheet_set_page_setup(handle_, ArgU32(info, 0), &out));
}

Napi::Value Workbook::SetSheetPageMargins(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    Napi::TypeError::New(env, "setSheetPageMargins expects (sheet:number, margins:object)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  fm_page_margins_t out{};
  out.left_engaged = SpecHas(spec, "left") ? 1 : 0;
  out.left = SpecPullDouble(spec, "left", 0.0);
  out.right_engaged = SpecHas(spec, "right") ? 1 : 0;
  out.right = SpecPullDouble(spec, "right", 0.0);
  out.top_engaged = SpecHas(spec, "top") ? 1 : 0;
  out.top = SpecPullDouble(spec, "top", 0.0);
  out.bottom_engaged = SpecHas(spec, "bottom") ? 1 : 0;
  out.bottom = SpecPullDouble(spec, "bottom", 0.0);
  out.header_engaged = SpecHas(spec, "header") ? 1 : 0;
  out.header = SpecPullDouble(spec, "header", 0.0);
  out.footer_engaged = SpecHas(spec, "footer") ? 1 : 0;
  out.footer = SpecPullDouble(spec, "footer", 0.0);
  return MakeStatus(env, fm_sheet_set_page_margins(handle_, ArgU32(info, 0), &out));
}

Napi::Value Workbook::SetSheetPrintOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    Napi::TypeError::New(env, "setSheetPrintOptions expects (sheet:number, options:object)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  fm_print_options_t out{};
  out.grid_lines_engaged = SpecHas(spec, "gridLines") ? 1 : 0;
  out.grid_lines = SpecPullBool(spec, "gridLines", false) ? 1 : 0;
  out.headings_engaged = SpecHas(spec, "headings") ? 1 : 0;
  out.headings = SpecPullBool(spec, "headings", false) ? 1 : 0;
  out.horizontal_centered_engaged = SpecHas(spec, "horizontalCentered") ? 1 : 0;
  out.horizontal_centered = SpecPullBool(spec, "horizontalCentered", false) ? 1 : 0;
  out.vertical_centered_engaged = SpecHas(spec, "verticalCentered") ? 1 : 0;
  out.vertical_centered = SpecPullBool(spec, "verticalCentered", false) ? 1 : 0;
  return MakeStatus(env, fm_sheet_set_print_options(handle_, ArgU32(info, 0), &out));
}

Napi::Value Workbook::SetSheetHeaderFooter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    Napi::TypeError::New(env, "setSheetHeaderFooter expects (sheet:number, headerFooter:object)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  // The six section strings are borrowed by the C struct, so their
  // storage has to outlive the call.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  fm_header_footer_t out{};
  out.odd_header = PullOptionalString(spec, "oddHeader", odd_header);
  out.odd_footer = PullOptionalString(spec, "oddFooter", odd_footer);
  out.even_header = PullOptionalString(spec, "evenHeader", even_header);
  out.even_footer = PullOptionalString(spec, "evenFooter", even_footer);
  out.first_header = PullOptionalString(spec, "firstHeader", first_header);
  out.first_footer = PullOptionalString(spec, "firstFooter", first_footer);
  out.different_odd_even_engaged = SpecHas(spec, "differentOddEven") ? 1 : 0;
  out.different_odd_even = SpecPullBool(spec, "differentOddEven", false) ? 1 : 0;
  out.different_first_engaged = SpecHas(spec, "differentFirst") ? 1 : 0;
  out.different_first = SpecPullBool(spec, "differentFirst", false) ? 1 : 0;
  out.scale_with_doc_engaged = SpecHas(spec, "scaleWithDoc") ? 1 : 0;
  out.scale_with_doc = SpecPullBool(spec, "scaleWithDoc", false) ? 1 : 0;
  out.align_with_margins_engaged = SpecHas(spec, "alignWithMargins") ? 1 : 0;
  out.align_with_margins = SpecPullBool(spec, "alignWithMargins", false) ? 1 : 0;
  return MakeStatus(env, fm_sheet_set_header_footer(handle_, ArgU32(info, 0), &out));
}

Napi::Value Workbook::GetSheetPageSetup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    return result;
  }
  fm_page_setup_t setup{};
  const fm_status_t rc = fm_sheet_get_page_setup(handle_, ArgU32(info, 0), &setup);
  if (rc != 0) {
    result.Set("status", MakeErrorStatus(env, rc));
    return result;
  }
  result.Set("status", MakeOkStatus(env));
  result.Set("orientation", Napi::Number::New(env, setup.orientation));
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
  Napi::Object result = Napi::Object::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    return result;
  }
  fm_page_margins_t margins{};
  const fm_status_t rc = fm_sheet_get_page_margins(handle_, ArgU32(info, 0), &margins);
  if (rc != 0) {
    result.Set("status", MakeErrorStatus(env, rc));
    return result;
  }
  result.Set("status", MakeOkStatus(env));
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

Napi::Value Workbook::Paginate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  Napi::Array print_area = Napi::Array::New(env);
  Napi::Array horizontal_breaks = Napi::Array::New(env);
  Napi::Array vertical_breaks = Napi::Array::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    result.Set("printArea", print_area);
    result.Set("horizontalBreaks", horizontal_breaks);
    result.Set("verticalBreaks", vertical_breaks);
    result.Set("pageCount", Napi::Number::New(env, 0));
    return result;
  }
  if (info.Length() < 1 || !info[0].IsNumber()) {
    Napi::TypeError::New(env, "paginate expects (sheet:number)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  fm_pagination_t* pagination = nullptr;
  const fm_status_t rc = fm_workbook_paginate(handle_, ArgU32(info, 0), &pagination);
  if (rc != 0) {
    result.Set("status", MakeErrorStatus(env, rc));
    result.Set("printArea", print_area);
    result.Set("horizontalBreaks", horizontal_breaks);
    result.Set("verticalBreaks", vertical_breaks);
    result.Set("pageCount", Napi::Number::New(env, 0));
    return result;
  }
  for (std::size_t i = 0; i < fm_pagination_print_area_count(pagination); ++i) {
    fm_print_range_t range{};
    if (fm_pagination_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("firstRow", Napi::Number::New(env, range.first_row));
    item.Set("firstCol", Napi::Number::New(env, range.first_col));
    item.Set("lastRow", Napi::Number::New(env, range.last_row));
    item.Set("lastCol", Napi::Number::New(env, range.last_col));
    print_area.Set(i, item);
  }
  for (std::size_t i = 0; i < fm_pagination_horizontal_break_count(pagination); ++i) {
    uint32_t row = 0;
    if (fm_pagination_horizontal_break_at(pagination, i, &row) == 0) {
      horizontal_breaks.Set(i, Napi::Number::New(env, row));
    }
  }
  for (std::size_t i = 0; i < fm_pagination_vertical_break_count(pagination); ++i) {
    uint32_t col = 0;
    if (fm_pagination_vertical_break_at(pagination, i, &col) == 0) {
      vertical_breaks.Set(i, Napi::Number::New(env, col));
    }
  }
  result.Set("status", MakeOkStatus(env));
  result.Set("printArea", print_area);
  result.Set("horizontalBreaks", horizontal_breaks);
  result.Set("verticalBreaks", vertical_breaks);
  result.Set("pageCount", Napi::Number::New(env, fm_pagination_page_count(pagination)));
  fm_pagination_destroy(pagination);
  return result;
}

}  // namespace formulon_node
