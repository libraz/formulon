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
const char* PullOptionalString(CheckedSpecReader& reader, const Napi::Object& spec, const char* key,
                               std::string& storage) {
  if (!reader.String(spec, key, &storage)) {
    return nullptr;
  }
  return storage.c_str();
}

}  // namespace

Napi::Value Workbook::SheetStringGetter(const Napi::CallbackInfo& info, SheetStringGetFn getter, const char* field) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), field, nullptr);
  }
  const char* text = nullptr;
  const fm_status_t rc = getter(handle_, ArgU32(info, 0), &text);
  return MakeStringFieldResult(env, rc, field, text);
}

Napi::Value Workbook::SheetStringSetter(const Napi::CallbackInfo& info, SheetStringSetFn setter) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string text = ArgString(info, 1);
  return MakeStatus(env, setter(handle_, ArgU32(info, 0), text.c_str()));
}

Napi::Value Workbook::GetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_page_setup_xml, "xml");
}

Napi::Value Workbook::SetSheetPageSetupXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_page_setup_xml);
}

Napi::Value Workbook::GetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_page_margins_xml, "xml");
}

Napi::Value Workbook::SetSheetPageMarginsXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_page_margins_xml);
}

Napi::Value Workbook::GetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_print_options_xml, "xml");
}

Napi::Value Workbook::SetSheetPrintOptionsXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_print_options_xml);
}

Napi::Value Workbook::GetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_header_footer_xml, "xml");
}

Napi::Value Workbook::SetSheetHeaderFooterXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_header_footer_xml);
}

Napi::Value Workbook::GetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_sheet_pr_xml, "xml");
}

Napi::Value Workbook::SetSheetSheetPrXml(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_sheet_pr_xml);
}

Napi::Value Workbook::SetSheetFitToPage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_set_fit_to_page(handle_, ArgU32(info, 0), ArgBool(info, 1) ? 1 : 0));
}

Napi::Value Workbook::GetSheetPrintArea(const Napi::CallbackInfo& info) {
  return SheetStringGetter(info, &fm_sheet_get_print_area, "ranges");
}

Napi::Value Workbook::SetSheetPrintArea(const Napi::CallbackInfo& info) {
  return SheetStringSetter(info, &fm_sheet_set_print_area);
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
  Napi::Array items = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "breaks", items);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const std::size_t count = rows ? fm_sheet_row_break_count(handle_, sheet) : fm_sheet_col_break_count(handle_, sheet);
  for (std::size_t i = 0; i < count; ++i) {
    fm_page_break_t brk{};
    const fm_status_t rc =
        rows ? fm_sheet_row_break_at(handle_, sheet, i, &brk) : fm_sheet_col_break_at(handle_, sheet, i, &brk);
    if (rc != 0) {
      return MakeFieldResult(env, MakeErrorStatus(env, rc), "breaks", Napi::Array::New(env));
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", Napi::Number::New(env, brk.id));
    item.Set("min", Napi::Number::New(env, brk.min));
    item.Set("max", Napi::Number::New(env, brk.max));
    item.Set("manual", Napi::Boolean::New(env, brk.manual != 0));
    items.Set(i, item);
  }
  return MakeFieldResult(env, MakeOkStatus(env), "breaks", items);
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
    Napi::TypeError::New(env, "setSheetPageSetup expects (sheet:number, setup:object)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  CheckedSpecReader reader(env);
  fm_page_setup_t out{};
  bool present = false;
  out.orientation = reader.U32(spec, "orientation", 0U, &present);
  out.orientation_engaged = present ? 1 : 0;
  out.paper_size = reader.U32(spec, "paperSize", 0U, &present);
  out.paper_size_engaged = present ? 1 : 0;
  out.scale = reader.U32(spec, "scale", 0U, &present);
  out.scale_engaged = present ? 1 : 0;
  out.fit_to_width = reader.U32(spec, "fitToWidth", 0U, &present);
  out.fit_to_width_engaged = present ? 1 : 0;
  out.fit_to_height = reader.U32(spec, "fitToHeight", 0U, &present);
  out.fit_to_height_engaged = present ? 1 : 0;
  out.fit_to_page = reader.Bool(spec, "fitToPage", false, &present) ? 1 : 0;
  out.fit_to_page_engaged = present ? 1 : 0;
  if (!reader.ok()) {
    // A malformed field left a pending JS exception:
    // stop before the C ABI call commits a default value for it.
    return env.Undefined();
  }
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
  CheckedSpecReader reader(env);
  fm_page_margins_t out{};
  bool present = false;
  out.left = reader.Double(spec, "left", 0.0, &present);
  out.left_engaged = present ? 1 : 0;
  out.right = reader.Double(spec, "right", 0.0, &present);
  out.right_engaged = present ? 1 : 0;
  out.top = reader.Double(spec, "top", 0.0, &present);
  out.top_engaged = present ? 1 : 0;
  out.bottom = reader.Double(spec, "bottom", 0.0, &present);
  out.bottom_engaged = present ? 1 : 0;
  out.header = reader.Double(spec, "header", 0.0, &present);
  out.header_engaged = present ? 1 : 0;
  out.footer = reader.Double(spec, "footer", 0.0, &present);
  out.footer_engaged = present ? 1 : 0;
  if (!reader.ok()) {
    // A malformed field left a pending JS exception:
    // stop before the C ABI call commits a default value for it.
    return env.Undefined();
  }
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
  CheckedSpecReader reader(env);
  fm_print_options_t out{};
  bool present = false;
  out.grid_lines = reader.Bool(spec, "gridLines", false, &present) ? 1 : 0;
  out.grid_lines_engaged = present ? 1 : 0;
  out.headings = reader.Bool(spec, "headings", false, &present) ? 1 : 0;
  out.headings_engaged = present ? 1 : 0;
  out.horizontal_centered = reader.Bool(spec, "horizontalCentered", false, &present) ? 1 : 0;
  out.horizontal_centered_engaged = present ? 1 : 0;
  out.vertical_centered = reader.Bool(spec, "verticalCentered", false, &present) ? 1 : 0;
  out.vertical_centered_engaged = present ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
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
  CheckedSpecReader reader(env);
  // The six section strings are borrowed by the C struct, so their
  // storage has to outlive the call.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  fm_header_footer_t out{};
  out.odd_header = PullOptionalString(reader, spec, "oddHeader", odd_header);
  out.odd_footer = PullOptionalString(reader, spec, "oddFooter", odd_footer);
  out.even_header = PullOptionalString(reader, spec, "evenHeader", even_header);
  out.even_footer = PullOptionalString(reader, spec, "evenFooter", even_footer);
  out.first_header = PullOptionalString(reader, spec, "firstHeader", first_header);
  out.first_footer = PullOptionalString(reader, spec, "firstFooter", first_footer);
  bool present = false;
  out.different_odd_even = reader.Bool(spec, "differentOddEven", false, &present) ? 1 : 0;
  out.different_odd_even_engaged = present ? 1 : 0;
  out.different_first = reader.Bool(spec, "differentFirst", false, &present) ? 1 : 0;
  out.different_first_engaged = present ? 1 : 0;
  out.scale_with_doc = reader.Bool(spec, "scaleWithDoc", false, &present) ? 1 : 0;
  out.scale_with_doc_engaged = present ? 1 : 0;
  out.align_with_margins = reader.Bool(spec, "alignWithMargins", false, &present) ? 1 : 0;
  out.align_with_margins_engaged = present ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_header_footer(handle_, ArgU32(info, 0), &out));
}

// Both typed getters below emit their whole declared payload on every exit
// path, as the WASM binding does for the same calls; only `status.ok` and
// the values differ.

Napi::Value Workbook::GetSheetPageSetup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  fm_page_setup_t setup{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_page_setup(handle_, ArgU32(info, 0), &setup) : kBindingInvalidHandle;
  if (rc != 0) {
    setup = fm_page_setup_t{};
  }
  result.Set("status", MakeStatus(env, rc));
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
  fm_page_margins_t margins{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_page_margins(handle_, ArgU32(info, 0), &margins) : kBindingInvalidHandle;
  if (rc != 0) {
    margins = fm_page_margins_t{};
  }
  result.Set("status", MakeStatus(env, rc));
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

Napi::Object RectToJs(Napi::Env env, const fm_rect_pt& r) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("x", Napi::Number::New(env, r.x));
  out.Set("y", Napi::Number::New(env, r.y));
  out.Set("width", Napi::Number::New(env, r.width));
  out.Set("height", Napi::Number::New(env, r.height));
  return out;
}

// Sets the pagination detail keys on `result`. A null `pagination` yields the
// all-zero record every failure path reports, so the keys exist on every exit.
void SetPaginationDetail(Napi::Env env, Napi::Object result, const fm_pagination_t* pagination) {
  fm_paper_info paper{};
  fm_margins_pt margins{};
  fm_rect_pt printable{};
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles titles{};
  Napi::Array pages = Napi::Array::New(env);
  Napi::Array horizontal_manual = Napi::Array::New(env);
  Napi::Array vertical_manual = Napi::Array::New(env);
  if (pagination != nullptr) {
    (void)fm_pagination_paper(pagination, &paper);
    (void)fm_pagination_margins(pagination, &margins);
    (void)fm_pagination_printable(pagination, &printable);
    (void)fm_pagination_scale(pagination, &scale);
    (void)fm_pagination_page_order(pagination, &page_order);
    (void)fm_pagination_print_titles(pagination, &titles);
    for (std::size_t i = 0; i < fm_pagination_page_count(pagination); ++i) {
      fm_page_layout page{};
      if (fm_pagination_page_at(pagination, i, &page) != 0) {
        continue;
      }
      Napi::Object item = Napi::Object::New(env);
      item.Set("areaIndex", Napi::Number::New(env, page.area_index));
      item.Set("firstRow", Napi::Number::New(env, page.first_row));
      item.Set("lastRow", Napi::Number::New(env, page.last_row));
      item.Set("firstCol", Napi::Number::New(env, page.first_col));
      item.Set("lastCol", Napi::Number::New(env, page.last_col));
      item.Set("originXPt", Napi::Number::New(env, page.origin_x_pt));
      item.Set("originYPt", Napi::Number::New(env, page.origin_y_pt));
      item.Set("widthPt", Napi::Number::New(env, page.width_pt));
      item.Set("heightPt", Napi::Number::New(env, page.height_pt));
      pages.Set(i, item);
    }
    for (std::size_t i = 0; i < fm_pagination_horizontal_break_count(pagination); ++i) {
      int32_t manual = 0;
      if (fm_pagination_horizontal_break_is_manual(pagination, i, &manual) == 0) {
        horizontal_manual.Set(i, Napi::Boolean::New(env, manual != 0));
      }
    }
    for (std::size_t i = 0; i < fm_pagination_vertical_break_count(pagination); ++i) {
      int32_t manual = 0;
      if (fm_pagination_vertical_break_is_manual(pagination, i, &manual) == 0) {
        vertical_manual.Set(i, Napi::Boolean::New(env, manual != 0));
      }
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
  Napi::Object titles_js = Napi::Object::New(env);
  titles_js.Set("hasRows", Napi::Boolean::New(env, titles.has_rows != 0));
  titles_js.Set("firstRow", Napi::Number::New(env, titles.first_row));
  titles_js.Set("lastRow", Napi::Number::New(env, titles.last_row));
  titles_js.Set("hasCols", Napi::Boolean::New(env, titles.has_cols != 0));
  titles_js.Set("firstCol", Napi::Number::New(env, titles.first_col));
  titles_js.Set("lastCol", Napi::Number::New(env, titles.last_col));
  result.Set("paper", paper_js);
  result.Set("margins", margins_js);
  result.Set("printable", RectToJs(env, printable));
  result.Set("scale", Napi::Number::New(env, scale));
  result.Set("pageOrder", Napi::Number::New(env, page_order));
  result.Set("printTitles", titles_js);
  result.Set("pages", pages);
  result.Set("horizontalBreakManual", horizontal_manual);
  result.Set("verticalBreakManual", vertical_manual);
}

}  // namespace

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
    SetPaginationDetail(env, result, nullptr);
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
    SetPaginationDetail(env, result, nullptr);
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
  SetPaginationDetail(env, result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

}  // namespace formulon_node
