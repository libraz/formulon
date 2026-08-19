#include "tests/oracle/roundtrip_authoring.h"

#include <algorithm>
#include <cstdint>
#include <cctype>
#include <cstdlib>
#include <map>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/status_macros.h"

namespace formulon {
namespace tests {
namespace oracle {

namespace {

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
}

/// Owns an `fm_workbook_t*` for the duration of one case. The C ABI hands
/// out raw handles; a case that fails halfway through authoring would
/// otherwise leak one per failure.
struct WorkbookDeleter {
  void operator()(fm_workbook_t* wb) const { fm_workbook_destroy(wb); }
};
using WorkbookHandle = std::unique_ptr<fm_workbook_t, WorkbookDeleter>;

/// Same, for the save buffer -- which must go back through
/// `fm_buffer_free`, not `free`.
struct BufferDeleter {
  void operator()(std::uint8_t* bytes) const { fm_buffer_free(bytes); }
};
using BufferHandle = std::unique_ptr<std::uint8_t, BufferDeleter>;

/// Wraps one C ABI call. `what` names the call in the failure message so a
/// broken case says which authoring step rejected it rather than just
/// reporting a status code.
Expected<void, Error> checked(fm_status_t status, const char* what) {
  // The C ABI mirrors `FormulonErrorCode`, whose success value is 0.
  if (status != static_cast<fm_status_t>(FormulonErrorCode::kOk)) {
    return invalid(std::string(what) + " failed with status " + std::to_string(static_cast<int>(status)));
  }
  return {};
}

/// Converts A1 column letters to a 0-based index. Returns false on an
/// empty or non-alphabetic token.
bool column_index(std::string_view letters, std::uint32_t* out) {
  if (letters.empty() || letters.size() > 3) {
    return false;
  }
  std::uint32_t index = 0;
  for (const char ch : letters) {
    const char upper = (ch >= 'a' && ch <= 'z') ? static_cast<char>(ch - 'a' + 'A') : ch;
    if (upper < 'A' || upper > 'Z') {
      return false;
    }
    index = index * 26U + static_cast<std::uint32_t>(upper - 'A' + 1);
  }
  *out = index - 1U;
  return true;
}

/// Splits an A1 cell address into 0-based `(row, col)`.
bool split_a1(std::string_view addr, std::uint32_t* row, std::uint32_t* col) {
  auto is_letter = [](char ch) { return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z'); };
  std::size_t split = 0;
  while (split < addr.size() && is_letter(addr[split])) {
    ++split;
  }
  if (split == 0 || split == addr.size()) {
    return false;
  }
  if (!column_index(addr.substr(0, split), col)) {
    return false;
  }
  std::uint32_t number = 0;
  for (const char ch : addr.substr(split)) {
    if (ch < '0' || ch > '9') {
      return false;
    }
    number = number * 10U + static_cast<std::uint32_t>(ch - '0');
  }
  if (number == 0) {
    return false;
  }
  *row = number - 1U;
  return true;
}

/// Writes one normalised `{kind, value}` cell record.
///
/// Mirrors `_write_cell` in tools/oracle/print_roundtrip.py so a case's
/// `sheets` block means the same thing on both halves of the round trip --
/// including the refusal to author an `error` cell, which the COM side
/// reaches by evaluating a trigger formula and a round-trip case has no
/// reason to need.
Expected<void, Error> write_cell(fm_workbook_t* wb, std::uint32_t sheet, const std::string& addr,
                                 const JsonValue& rec) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!split_a1(addr, &row, &col)) {
    return invalid("not an A1 cell address: " + addr);
  }
  const JsonValue* kind_v = rec.is_object() ? rec.find("kind") : nullptr;
  const std::string kind = kind_v != nullptr && kind_v->is_string() ? kind_v->as_string() : std::string();
  const JsonValue* value_v = rec.is_object() ? rec.find("value") : nullptr;

  if (kind == "blank") {
    return checked(fm_workbook_set_blank(wb, sheet, row, col), "fm_workbook_set_blank");
  }
  if (kind == "number") {
    if (value_v == nullptr || !value_v->is_number()) {
      return invalid("number cell " + addr + " has no numeric value");
    }
    return checked(fm_workbook_set_number(wb, sheet, row, col, value_v->as_number()), "fm_workbook_set_number");
  }
  if (kind == "bool") {
    if (value_v == nullptr || !value_v->is_bool()) {
      return invalid("bool cell " + addr + " has no boolean value");
    }
    return checked(fm_workbook_set_bool(wb, sheet, row, col, value_v->as_bool() ? 1 : 0), "fm_workbook_set_bool");
  }
  if (kind == "text") {
    if (value_v == nullptr || !value_v->is_string()) {
      return invalid("text cell " + addr + " has no string value");
    }
    return checked(fm_workbook_set_text(wb, sheet, row, col, value_v->as_string().c_str()), "fm_workbook_set_text");
  }
  if (kind == "formula") {
    const JsonValue* formula_v = rec.is_object() ? rec.find("formula") : nullptr;
    if (formula_v == nullptr || !formula_v->is_string()) {
      return invalid("formula cell " + addr + " has no formula string");
    }
    return checked(fm_workbook_set_formula(wb, sheet, row, col, formula_v->as_string().c_str()),
                   "fm_workbook_set_formula");
  }
  return invalid("cell kind '" + kind + "' is not supported in a roundtrip case (" + addr + ")");
}

/// Materialises the `sheets` block, returning sheet name -> index.
Expected<std::map<std::string, std::uint32_t>, Error> apply_sheets(fm_workbook_t* wb, const JsonValue& spec) {
  std::map<std::string, std::uint32_t> indices;
  const JsonValue* sheets = spec.find("sheets");
  if (sheets == nullptr || !sheets->is_object()) {
    return indices;
  }
  std::uint32_t position = 0;
  for (const auto& [name, cells] : sheets->as_object()) {
    if (position == 0) {
      RETURN_IF_ERROR(checked(fm_workbook_rename_sheet(wb, 0, name.c_str()), "fm_workbook_rename_sheet"));
    } else {
      RETURN_IF_ERROR(checked(fm_workbook_add_sheet(wb, name.c_str()), "fm_workbook_add_sheet"));
    }
    indices[name] = position;
    if (cells.is_object()) {
      for (const auto& [addr, rec] : cells.as_object()) {
        RETURN_IF_ERROR(write_cell(wb, position, addr, rec));
      }
    }
    ++position;
  }
  return indices;
}

/// Applies case-level `column_widths` / `row_heights` to the first sheet.
/// Scoped to sheet 0 for the same reason the capture side is: the case
/// schema carries one layout map per case, not one per sheet.
Expected<void, Error> apply_layout(fm_workbook_t* wb, const JsonValue& spec) {
  if (const JsonValue* widths = spec.find("column_widths"); widths != nullptr && widths->is_object()) {
    for (const auto& [key, width] : widths->as_object()) {
      std::uint32_t col = 0;
      if (!column_index(key, &col)) {
        return invalid("column_widths key '" + key + "' is not a column letter");
      }
      if (!width.is_number()) {
        return invalid("column_widths['" + key + "'] is not a number");
      }
      RETURN_IF_ERROR(
          checked(fm_sheet_set_column_width(wb, 0, col, col, width.as_number()), "fm_sheet_set_column_width"));
    }
  }
  if (const JsonValue* heights = spec.find("row_heights"); heights != nullptr && heights->is_object()) {
    for (const auto& [key, height] : heights->as_object()) {
      const long row = std::strtol(key.c_str(), nullptr, 10);
      if (row <= 0) {
        return invalid("row_heights key '" + key + "' is not a 1-based row number");
      }
      if (!height.is_number()) {
        return invalid("row_heights['" + key + "'] is not a number");
      }
      RETURN_IF_ERROR(checked(fm_sheet_set_row_height(wb, 0, static_cast<std::uint32_t>(row - 1), height.as_number()),
                              "fm_sheet_set_row_height"));
    }
  }
  return {};
}

/// Copies an optional integer field into a C ABI patch slot, engaging it
/// only when the case states it. "Absent means untouched" is what keeps a
/// case's file silent about everything it did not ask for.
void patch_u32(const JsonValue* block, const char* field, std::int32_t* engaged, std::uint32_t* out) {
  const JsonValue* v = block->find(field);
  if (v != nullptr && v->is_number()) {
    *engaged = 1;
    *out = static_cast<std::uint32_t>(v->as_number());
  }
}

void patch_bool(const JsonValue* block, const char* field, std::int32_t* engaged, std::int32_t* out) {
  const JsonValue* v = block->find(field);
  if (v != nullptr && v->is_bool()) {
    *engaged = 1;
    *out = v->as_bool() ? 1 : 0;
  }
}

void patch_double(const JsonValue* block, const char* field, std::int32_t* engaged, double* out) {
  const JsonValue* v = block->find(field);
  if (v != nullptr && v->is_number()) {
    *engaged = 1;
    *out = v->as_number();
  }
}

/// Returns the string at `field`, or nullptr when absent -- which is the
/// C ABI's "leave this header/footer section alone".
const char* optional_text(const JsonValue* block, const char* field) {
  const JsonValue* v = block->find(field);
  return v != nullptr && v->is_string() ? v->as_string().c_str() : nullptr;
}

/// Drives the print-authoring calls from the `roundtrip` block, in the
/// same order as tools/oracle/print_roundtrip.py.
Expected<void, Error> apply_print_settings(fm_workbook_t* wb, std::uint32_t sheet, const JsonValue& rt) {
  if (const JsonValue* block = rt.find("page_setup"); block != nullptr && block->is_object()) {
    fm_page_setup_t setup{};
    patch_u32(block, "orientation", &setup.orientation_engaged, &setup.orientation);
    patch_u32(block, "paper_size", &setup.paper_size_engaged, &setup.paper_size);
    patch_u32(block, "scale", &setup.scale_engaged, &setup.scale);
    patch_u32(block, "fit_to_width", &setup.fit_to_width_engaged, &setup.fit_to_width);
    patch_u32(block, "fit_to_height", &setup.fit_to_height_engaged, &setup.fit_to_height);
    patch_bool(block, "fit_to_page", &setup.fit_to_page_engaged, &setup.fit_to_page);
    RETURN_IF_ERROR(checked(fm_sheet_set_page_setup(wb, sheet, &setup), "fm_sheet_set_page_setup"));
  }

  if (const JsonValue* block = rt.find("page_margins"); block != nullptr && block->is_object()) {
    fm_page_margins_t margins{};
    patch_double(block, "left", &margins.left_engaged, &margins.left);
    patch_double(block, "right", &margins.right_engaged, &margins.right);
    patch_double(block, "top", &margins.top_engaged, &margins.top);
    patch_double(block, "bottom", &margins.bottom_engaged, &margins.bottom);
    patch_double(block, "header", &margins.header_engaged, &margins.header);
    patch_double(block, "footer", &margins.footer_engaged, &margins.footer);
    RETURN_IF_ERROR(checked(fm_sheet_set_page_margins(wb, sheet, &margins), "fm_sheet_set_page_margins"));
  }

  if (const JsonValue* block = rt.find("print_options"); block != nullptr && block->is_object()) {
    fm_print_options_t options{};
    patch_bool(block, "grid_lines", &options.grid_lines_engaged, &options.grid_lines);
    patch_bool(block, "headings", &options.headings_engaged, &options.headings);
    patch_bool(block, "horizontal_centered", &options.horizontal_centered_engaged, &options.horizontal_centered);
    patch_bool(block, "vertical_centered", &options.vertical_centered_engaged, &options.vertical_centered);
    RETURN_IF_ERROR(checked(fm_sheet_set_print_options(wb, sheet, &options), "fm_sheet_set_print_options"));
  }

  if (const JsonValue* block = rt.find("header_footer"); block != nullptr && block->is_object()) {
    fm_header_footer_t hf{};
    hf.odd_header = optional_text(block, "odd_header");
    hf.odd_footer = optional_text(block, "odd_footer");
    hf.even_header = optional_text(block, "even_header");
    hf.even_footer = optional_text(block, "even_footer");
    hf.first_header = optional_text(block, "first_header");
    hf.first_footer = optional_text(block, "first_footer");
    patch_bool(block, "different_odd_even", &hf.different_odd_even_engaged, &hf.different_odd_even);
    patch_bool(block, "different_first", &hf.different_first_engaged, &hf.different_first);
    patch_bool(block, "scale_with_doc", &hf.scale_with_doc_engaged, &hf.scale_with_doc);
    patch_bool(block, "align_with_margins", &hf.align_with_margins_engaged, &hf.align_with_margins);
    RETURN_IF_ERROR(checked(fm_sheet_set_header_footer(wb, sheet, &hf), "fm_sheet_set_header_footer"));
  }

  const JsonValue* area = rt.find("print_area");
  if (area != nullptr && area->is_string() && !area->as_string().empty()) {
    RETURN_IF_ERROR(checked(fm_sheet_set_print_area(wb, sheet, area->as_string().c_str()), "fm_sheet_set_print_area"));
  }

  if (const JsonValue* titles = rt.find("print_titles"); titles != nullptr && titles->is_object()) {
    const JsonValue* rows = titles->find("repeat_rows");
    const JsonValue* cols = titles->find("repeat_cols");
    const char* repeat_rows = rows != nullptr && rows->is_string() ? rows->as_string().c_str() : "";
    const char* repeat_cols = cols != nullptr && cols->is_string() ? cols->as_string().c_str() : "";
    RETURN_IF_ERROR(
        checked(fm_sheet_set_print_titles(wb, sheet, repeat_rows, repeat_cols), "fm_sheet_set_print_titles"));
  }

  if (const JsonValue* rows = rt.find("row_breaks"); rows != nullptr && rows->is_array()) {
    for (const JsonValue& entry : rows->as_array()) {
      if (!entry.is_number() || entry.as_number() < 1.0) {
        return invalid("row_breaks entries are 1-based Excel row numbers");
      }
      // Case rows are 1-based, matching every other A1-shaped field in the
      // schema; the API takes the 0-based index the break precedes.
      const auto row = static_cast<std::uint32_t>(entry.as_number()) - 1U;
      RETURN_IF_ERROR(checked(fm_sheet_add_row_break(wb, sheet, row, 1), "fm_sheet_add_row_break"));
    }
  }

  if (const JsonValue* cols = rt.find("col_breaks"); cols != nullptr && cols->is_array()) {
    for (const JsonValue& entry : cols->as_array()) {
      std::uint32_t col = 0;
      if (!entry.is_string() || !column_index(entry.as_string(), &col)) {
        return invalid("col_breaks entries are column letters");
      }
      RETURN_IF_ERROR(checked(fm_sheet_add_col_break(wb, sheet, col, 1), "fm_sheet_add_col_break"));
    }
  }

  return {};
}

/// Reads an OOXML boolean attribute. ECMA-376 §22.9.2.7 accepts both the
/// lexical forms, and the two producers differ: Excel writes `1`, our
/// writer writes `true`.
bool xml_bool(const pugi::xml_node& node, const char* name, bool fallback) {
  const pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    return fallback;
  }
  const std::string_view text = attr.value();
  return text == "1" || text == "true" || text == "on";
}

/// Resolves the package path of the worksheet part at `sheet_index`.
///
/// Walks `xl/workbook.xml` for the sheet's relationship id and
/// `xl/_rels/workbook.xml.rels` for its target, rather than assuming
/// `sheet<N+1>.xml`. The naming convention is the writer's own, and this
/// verifier exists precisely to avoid taking the writer's word for
/// anything.
Expected<std::string, Error> worksheet_part(const pugi::xml_document& workbook, const pugi::xml_document& rels,
                                            std::uint32_t sheet_index) {
  std::uint32_t position = 0;
  std::string rel_id;
  for (pugi::xml_node sheet = workbook.child("workbook").child("sheets").child("sheet"); sheet;
       sheet = sheet.next_sibling("sheet")) {
    if (position == sheet_index) {
      rel_id = sheet.attribute("r:id").as_string();
      break;
    }
    ++position;
  }
  if (rel_id.empty()) {
    return invalid("the saved workbook has no sheet at index " + std::to_string(sheet_index));
  }
  for (pugi::xml_node rel = rels.child("Relationships").child("Relationship"); rel;
       rel = rel.next_sibling("Relationship")) {
    if (rel_id == rel.attribute("Id").as_string()) {
      // Workbook-relative target, as the OPC part-name rules define it for
      // a relationship declared in `xl/_rels/`.
      return "xl/" + std::string(rel.attribute("Target").as_string());
    }
  }
  return invalid("relationship '" + rel_id + "' is not declared in xl/_rels/workbook.xml.rels");
}

/// Strips a defined name's sheet qualifier and absolute markers, leaving
/// the shape Excel's `PageSetup.PrintArea` / `PrintTitleRows` report:
/// `Sheet1!$A$1:$H$60` becomes `A1:H60`.
std::string strip_reference_decoration(std::string_view token) {
  const std::size_t bang = token.rfind('!');
  if (bang != std::string_view::npos) {
    token.remove_prefix(bang + 1);
  }
  std::string out;
  out.reserve(token.size());
  for (const char ch : token) {
    if (ch != '$') {
      out.push_back(ch);
    }
  }
  return out;
}

/// True when every character is a digit -- the shape of the row half of a
/// `1:2` whole-row span.
bool all_digits(std::string_view text) {
  return !text.empty() && std::all_of(text.begin(), text.end(), [](char ch) { return ch >= '0' && ch <= '9'; });
}

/// Splits a defined-name formula into its comma-separated areas and files
/// each one as a print area, a repeat-rows span, or a repeat-columns span.
///
/// `_xlnm.Print_Titles` stores rows and columns in one formula
/// (`Sheet1!$1:$2,Sheet1!$A:$A`) while Excel reports them as two
/// properties, so the split happens on shape: `1:2` is a row span,
/// `A:A` a column span, anything else an area.
void classify_defined_name(std::string_view formula, std::string* areas, std::string* rows, std::string* cols) {
  std::size_t begin = 0;
  while (begin <= formula.size()) {
    const std::size_t comma = formula.find(',', begin);
    const std::string_view raw = formula.substr(begin, comma == std::string_view::npos ? comma : comma - begin);
    const std::string token = strip_reference_decoration(raw);
    const std::size_t colon = token.find(':');
    std::string* target = areas;
    if (colon != std::string::npos) {
      const std::string_view first(token.data(), colon);
      const std::string_view second(token.data() + colon + 1, token.size() - colon - 1);
      if (all_digits(first) && all_digits(second)) {
        target = rows;
      } else if (!first.empty() && !second.empty() && !std::isdigit(static_cast<unsigned char>(first.back())) &&
                 !std::isdigit(static_cast<unsigned char>(second.back()))) {
        target = cols;
      }
    }
    if (!target->empty()) {
      target->push_back(',');
    }
    target->append(token);
    if (comma == std::string_view::npos) {
      break;
    }
    begin = comma + 1;
  }
}

/// Reads the print settings straight out of the saved package.
///
/// Deliberately not via `fm_workbook_load`: a reader that undoes a writer
/// mistake reproduces the model the case started from, and the resulting
/// agreement says nothing about the file Excel was handed. What the bytes
/// literally say is the only observation comparable with Excel's.
Expected<RoundtripObservation, Error> observe_saved_bytes(const std::uint8_t* bytes, std::size_t len,
                                                          std::uint32_t sheet_index) {
  RoundtripObservation out;
  out.xlsx_bytes = len;

  io::ZipReader zip;
  RETURN_IF_ERROR(zip.open(io::ByteSpan{bytes, len}));

  std::vector<std::uint8_t> workbook_bytes;
  ASSIGN_OR_RETURN(workbook_bytes, zip.read_entry("xl/workbook.xml"));
  pugi::xml_document workbook;
  if (!workbook.load_buffer(workbook_bytes.data(), workbook_bytes.size())) {
    return invalid("xl/workbook.xml is not well-formed XML");
  }
  std::vector<std::uint8_t> rels_bytes;
  ASSIGN_OR_RETURN(rels_bytes, zip.read_entry("xl/_rels/workbook.xml.rels"));
  pugi::xml_document rels;
  if (!rels.load_buffer(rels_bytes.data(), rels_bytes.size())) {
    return invalid("xl/_rels/workbook.xml.rels is not well-formed XML");
  }

  std::string part;
  ASSIGN_OR_RETURN(part, worksheet_part(workbook, rels, sheet_index));
  std::vector<std::uint8_t> sheet_bytes;
  ASSIGN_OR_RETURN(sheet_bytes, zip.read_entry(part));
  pugi::xml_document sheet_doc;
  if (!sheet_doc.load_buffer(sheet_bytes.data(), sheet_bytes.size())) {
    return invalid(part + " is not well-formed XML");
  }
  const pugi::xml_node worksheet = sheet_doc.child("worksheet");

  const pugi::xml_node page_setup = worksheet.child("pageSetup");
  // Attribute *presence* is the whole point: an attribute the case did not
  // author is absent here, and Excel then answers from the printer's
  // defaults rather than from anything we wrote.
  if (const pugi::xml_attribute attr = page_setup.attribute("orientation"); attr) {
    const std::string_view value = attr.value();
    out.orientation_stated = true;
    out.orientation = value == "landscape" ? 2U : (value == "portrait" ? 1U : 0U);
  }
  if (const pugi::xml_attribute attr = page_setup.attribute("paperSize"); attr) {
    out.paper_size_stated = true;
    out.paper_size = attr.as_uint(0);
  }
  if (const pugi::xml_attribute attr = page_setup.attribute("scale"); attr) {
    out.scale_stated = true;
    out.scale = attr.as_uint(0);
  }
  if (const pugi::xml_attribute attr = page_setup.attribute("fitToWidth"); attr) {
    out.fit_to_width_stated = true;
    out.fit_to_width = attr.as_uint(0);
  }
  if (const pugi::xml_attribute attr = page_setup.attribute("fitToHeight"); attr) {
    out.fit_to_height_stated = true;
    out.fit_to_height = attr.as_uint(0);
  }
  out.fit_to_page = xml_bool(worksheet.child("sheetPr").child("pageSetUpPr"), "fitToPage", false);

  // An absent `<pageMargins>` leaves the OOXML defaults standing, which is
  // also what Excel shows for a file that states none.
  const pugi::xml_node margins = worksheet.child("pageMargins");
  out.margin_left = margins.attribute("left").as_double(ooxml_defaults::kPageMarginSideInches);
  out.margin_right = margins.attribute("right").as_double(ooxml_defaults::kPageMarginSideInches);
  out.margin_top = margins.attribute("top").as_double(ooxml_defaults::kPageMarginTopBottomInches);
  out.margin_bottom = margins.attribute("bottom").as_double(ooxml_defaults::kPageMarginTopBottomInches);
  out.margin_header = margins.attribute("header").as_double(ooxml_defaults::kPageMarginHeaderFooterInches);
  out.margin_footer = margins.attribute("footer").as_double(ooxml_defaults::kPageMarginHeaderFooterInches);

  const pugi::xml_node options = worksheet.child("printOptions");
  out.grid_lines = xml_bool(options, "gridLines", false);
  out.headings = xml_bool(options, "headings", false);
  out.horizontal_centered = xml_bool(options, "horizontalCentered", false);
  out.vertical_centered = xml_bool(options, "verticalCentered", false);

  const pugi::xml_node hf = worksheet.child("headerFooter");
  out.header_footer.different_odd_even = xml_bool(hf, "differentOddEven", false);
  out.header_footer.different_first = xml_bool(hf, "differentFirst", false);
  out.header_footer.scale_with_doc = xml_bool(hf, "scaleWithDoc", true);
  out.header_footer.align_with_margins = xml_bool(hf, "alignWithMargins", true);
  out.header_footer.odd_header = hf.child("oddHeader").text().as_string();
  out.header_footer.odd_footer = hf.child("oddFooter").text().as_string();
  out.header_footer.even_header = hf.child("evenHeader").text().as_string();
  out.header_footer.even_footer = hf.child("evenFooter").text().as_string();
  out.header_footer.first_header = hf.child("firstHeader").text().as_string();
  out.header_footer.first_footer = hf.child("firstFooter").text().as_string();

  // `<brk id>` is read raw. Routing it through the reader's structured
  // view would put the writer's `+1` and the reader's `-1` back on the
  // same path, where they cancel -- which is exactly how the off-by-one
  // this suite now pins survived every Formulon-only test.
  for (pugi::xml_node brk = worksheet.child("rowBreaks").child("brk"); brk; brk = brk.next_sibling("brk")) {
    if (xml_bool(brk, "man", false)) {
      out.manual_row_breaks.push_back(brk.attribute("id").as_uint(0));
    }
  }
  for (pugi::xml_node brk = worksheet.child("colBreaks").child("brk"); brk; brk = brk.next_sibling("brk")) {
    if (xml_bool(brk, "man", false)) {
      out.manual_col_breaks.push_back(brk.attribute("id").as_uint(0));
    }
  }
  std::sort(out.manual_row_breaks.begin(), out.manual_row_breaks.end());
  std::sort(out.manual_col_breaks.begin(), out.manual_col_breaks.end());

  for (pugi::xml_node name = workbook.child("workbook").child("definedNames").child("definedName"); name;
       name = name.next_sibling("definedName")) {
    if (name.attribute("localSheetId").as_uint(0) != sheet_index) {
      continue;
    }
    const std::string_view which = name.attribute("name").as_string();
    std::string discard;
    if (which == "_xlnm.Print_Area") {
      classify_defined_name(name.text().as_string(), &out.print_area, &discard, &discard);
    } else if (which == "_xlnm.Print_Titles") {
      classify_defined_name(name.text().as_string(), &discard, &out.print_title_rows, &out.print_title_cols);
    }
  }

  return out;
}

}  // namespace

Expected<RoundtripObservation, Error> observe_roundtrip_from_spec(const JsonValue& spec) {
  if (!spec.is_object()) {
    return invalid("roundtrip spec is not an object");
  }
  const JsonValue* rt = spec.find("roundtrip");
  if (rt == nullptr || !rt->is_object()) {
    return invalid("spec has no 'roundtrip' block");
  }

  // `fm_workbook_create` rather than `fm_workbook_create_empty`: it is the
  // factory a caller reaches for, so the fixture carries the same seeded
  // style table a real authored report would -- and it is what the capture
  // side's `Workbook.create_default()` calls.
  fm_workbook_t* raw = nullptr;
  RETURN_IF_ERROR(checked(fm_workbook_create(&raw), "fm_workbook_create"));
  WorkbookHandle wb(raw);

  std::map<std::string, std::uint32_t> sheets;
  ASSIGN_OR_RETURN(sheets, apply_sheets(wb.get(), spec));
  RETURN_IF_ERROR(apply_layout(wb.get(), spec));

  std::uint32_t sheet = 0;
  const JsonValue* sheet_v = rt->find("sheet");
  if (sheet_v != nullptr && sheet_v->is_string()) {
    const auto it = sheets.find(sheet_v->as_string());
    if (it == sheets.end()) {
      return invalid("roundtrip sheet '" + sheet_v->as_string() + "' is not one of the case's sheets");
    }
    sheet = it->second;
  } else if (sheet_v != nullptr && sheet_v->is_number()) {
    sheet = static_cast<std::uint32_t>(sheet_v->as_number());
  } else {
    return invalid("roundtrip block missing required 'sheet'");
  }

  RETURN_IF_ERROR(apply_print_settings(wb.get(), sheet, *rt));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  RETURN_IF_ERROR(checked(fm_workbook_save_as(wb.get(), FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), "fm_workbook_save_as"));
  BufferHandle buffer(bytes);

  return observe_saved_bytes(buffer.get(), len, sheet);
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
