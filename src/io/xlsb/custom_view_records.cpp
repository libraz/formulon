// @size-budget: 20 KB raw / 6 KB Brotli (both directions of the custom-view codec, with the plan in
// user_book_views.cpp)
#include "io/xlsb/custom_view_records.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "auto_filter.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml/print_settings_parse.h"
#include "io/sheet_view_xml.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/print_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/view_records.h"
#include "io/xml_utils.h"
#include "sheet_layout.h"
#include "sheet_view.h"
#include "utils/a1_ref.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// A boolean attribute and the flag bit it sets when true.
struct FlagAttr {
  std::string_view name;
  std::uint32_t mask;
  bool fallback;
};

// BrtUserBookView flags (u16 at payload offset 42).
constexpr std::array<FlagAttr, 12> kBookFlags = {{
    {"minimized", 0x0001U, false},
    {"showHorizontalScroll", 0x0002U, true},
    {"showVerticalScroll", 0x0004U, true},
    {"showSheetTabs", 0x0008U, true},
    {"maximized", 0x0010U, false},
    {"showFormulaBar", 0x0020U, true},
    {"showStatusbar", 0x0040U, true},
    {"includePrintSettings", 0x0800U, true},
    {"includeHiddenRowCol", 0x1000U, true},
    {"autoUpdate", 0x2000U, false},
    {"changesSavedWin", 0x4000U, false},
    {"onlySync", 0x8000U, false},
}};
constexpr unsigned kShowCommentsShift = 7U;
constexpr unsigned kShowObjectsShift = 9U;
constexpr std::uint32_t kDefaultTabRatio = 600U;
// @limit: format
constexpr std::uint32_t kMaxTabRatio = 1000U;
constexpr std::uint32_t kNoMergeInterval = 0xFFFFU;

// BrtBeginUserShView flags (u32 at payload offset 28).
constexpr std::array<FlagAttr, 14> kSheetFlags = {{
    {"showPageBreaks", 0x00000001U, false},
    {"showFormulas", 0x00000002U, false},
    {"showGridLines", 0x00000004U, true},
    {"showRowCol", 0x00000008U, true},
    {"outlineSymbols", 0x00000010U, true},
    {"zeroValues", 0x00000020U, true},
    {"fitToPage", 0x00000400U, false},
    {"printArea", 0x00000800U, false},
    {"filter", 0x00002000U, false},
    {"showAutoFilter", 0x00004000U, false},
    {"hiddenRows", 0x00080000U, false},
    {"hiddenColumns", 0x00200000U, false},
    {"filterUnique", 0x02000000U, false},
    {"showRuler", 0x20000000U, true},
}};
// `<printOptions>` repeated in the view's flags.
constexpr std::array<FlagAttr, 4> kSheetPrintFlags = {{
    {"horizontalCentered", 0x00000040U, false},
    {"verticalCentered", 0x00000080U, false},
    {"headings", 0x00000100U, false},
    {"gridLines", 0x00000200U, false},
}};
constexpr std::uint32_t kSheetPaneFlags = 0x00020000U | 0x00040000U;
constexpr std::uint32_t kSheetAlwaysFlag = 0x01000000U;
constexpr std::uint32_t kSheetStateHidden = 0x00400000U;
constexpr std::uint32_t kSheetStateVeryHidden = 0x00800000U;
constexpr std::uint32_t kSheetViewPageBreakPreview = 0x04000000U;
constexpr std::uint32_t kSheetViewPageLayout = 0x08000000U;
// u32 after the GUID, 1 in every view Excel writes.
constexpr std::uint32_t kSheetViewMarker = 1U;
constexpr std::uint32_t kDefaultScale = 100U;
constexpr std::uint32_t kMinScale = 10U;
// @limit: format
constexpr std::uint32_t kMaxScale = 400U;
constexpr std::uint32_t kDefaultColorId = 64U;
// The margins Excel gives a view stating none.
constexpr std::string_view kDefaultViewMargins =
    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";

bool ParseBool(const pugi::xml_attribute& attr, bool fallback, bool* out) {
  if (!attr) {
    *out = fallback;
    return true;
  }
  const std::string_view text = attr.value();
  if (text == "1" || text == "true") {
    *out = true;
    return true;
  }
  if (text == "0" || text == "false") {
    *out = false;
    return true;
  }
  return false;
}

bool ParseUint(const pugi::xml_attribute& attr, std::uint32_t fallback, std::uint32_t* out) {
  if (!attr) {
    *out = fallback;
    return true;
  }
  const std::string_view text = attr.value();
  if (text.empty()) {
    return false;
  }
  std::uint64_t value = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return false;
    }
    value = value * 10U + static_cast<std::uint64_t>(c - '0');
    if (value > std::numeric_limits<std::uint32_t>::max()) {
      return false;
    }
  }
  *out = static_cast<std::uint32_t>(value);
  return true;
}

/// True when every attribute of `node` is one of `known`.
bool OnlyAttributes(const pugi::xml_node& node, std::initializer_list<std::string_view> known) {
  for (const pugi::xml_attribute& attr : node.attributes()) {
    bool found = false;
    for (const std::string_view name : known) {
      found = found || name == attr.name();
    }
    if (!found) {
      return false;
    }
  }
  return true;
}

/// ORs the bits of `flags` the attributes of `node` set; false for a value that is not a boolean.
bool AddFlags(const pugi::xml_node& node, const FlagAttr* flags, std::size_t count, std::uint32_t* bits) {
  for (const FlagAttr* flag_it = flags; flag_it != flags + count; ++flag_it) {
    const FlagAttr& flag = *flag_it;
    bool value = false;
    if (!ParseBool(node.attribute(flag.name.data()), flag.fallback, &value)) {
      return false;
    }
    *bits |= value ? flag.mask : 0U;
  }
  return true;
}

/// The index of `value` in `names`, `fallback` for an absent attribute; false for another value.
bool ParseEnum(const pugi::xml_attribute& attr, std::initializer_list<std::string_view> names, std::uint32_t fallback,
               std::uint32_t* out) {
  if (!attr) {
    *out = fallback;
    return true;
  }
  std::uint32_t index = 0;
  for (const std::string_view name : names) {
    if (name == attr.value()) {
      *out = index;
      return true;
    }
    ++index;
  }
  return false;
}

/// False when `seen` already holds `name`, which is then a repeat; otherwise records it.
bool FirstSeen(std::vector<std::string_view>* seen, std::string_view name) {
  if (std::find(seen->begin(), seen->end(), name) != seen->end()) {
    return false;
  }
  seen->push_back(name);
  return true;
}

bool ParseGuidAttr(const pugi::xml_node& node, std::array<std::uint8_t, kGuidBytes>* guid) {
  return ParseGuid(node.attribute("guid").value(), *guid);
}

/// The pane and selections of a `<customSheetView>`, as Excel holds them: a pane split on both axes and at most one
/// selection per pane, each naming a pane the view has.
bool ReadPaneAndSelections(const pugi::xml_node& view, SheetView* scratch) {
  const pugi::xml_node pane = view.child("pane");
  if (pane && (!OnlyAttributes(pane, {"xSplit", "ySplit", "topLeftCell", "activePane", "state"}) ||
               std::string_view(pane.attribute("state").value()) == "frozenSplit")) {
    return false;
  }
  for (const pugi::xml_node& selection : view.children("selection")) {
    if (!OnlyAttributes(selection, {"pane", "activeCell", "activeCellId", "sqref"})) {
      return false;
    }
  }
  read_sheet_view_element(view, *scratch);
  if (pane && !(scratch->divides_rows() && scratch->divides_cols())) {
    return false;
  }
  std::vector<std::string_view> panes;
  for (const SheetSelection& selection : scratch->selections) {
    const std::string_view key = SheetView::pane_key(selection.pane);
    if (!scratch->has_pane(key) || !FirstSeen(&panes, key)) {
      return false;
    }
  }
  return true;
}

/// The view's print settings from its print children, the default margins when it states none.
bool ReadViewPrint(const pugi::xml_node& view, SheetPrintSettings* print, std::uint32_t* print_flags) {
  const pugi::xml_node margins = view.child("pageMargins");
  if (margins) {
    for (const char* side : {"left", "right", "top", "bottom", "header", "footer"}) {
      if (!margins.attribute(side)) {
        return false;
      }
    }
    if (!OnlyAttributes(margins, {"left", "right", "top", "bottom", "header", "footer"})) {
      return false;
    }
    print->page_margins_xml = raw_xml(margins);
  } else {
    print->page_margins_xml = std::string(kDefaultViewMargins);
  }
  if (const pugi::xml_node options = view.child("printOptions")) {
    if (!OnlyAttributes(options, {"horizontalCentered", "verticalCentered", "headings", "gridLines"}) ||
        !AddFlags(options, kSheetPrintFlags.data(), kSheetPrintFlags.size(), print_flags)) {
      return false;
    }
    print->print_options_xml = raw_xml(options);
  }
  if (const pugi::xml_node setup = view.child("pageSetup")) {
    if (!OnlyAttributes(setup, {"paperSize", "scale", "firstPageNumber", "fitToWidth", "fitToHeight", "pageOrder",
                                "orientation", "blackAndWhite", "draft", "cellComments", "useFirstPageNumber", "errors",
                                "horizontalDpi", "verticalDpi", "copies"})) {
      return false;
    }
    print->page_setup_xml = raw_xml(setup);
  }
  if (const pugi::xml_node header_footer = view.child("headerFooter")) {
    if (!OnlyAttributes(header_footer, {"differentOddEven", "differentFirst", "scaleWithDoc", "alignWithMargins"})) {
      return false;
    }
    for (const pugi::xml_node& part : header_footer.children()) {
      const std::string_view name = part.name();
      if (part.first_attribute() || (name != "oddHeader" && name != "oddFooter" && name != "evenHeader" &&
                                     name != "evenFooter" && name != "firstHeader" && name != "firstFooter")) {
        return false;
      }
    }
    print->header_footer_xml = raw_xml(header_footer);
  }
  for (const char* axis : {"rowBreaks", "colBreaks"}) {
    const pugi::xml_node breaks = view.child(axis);
    if (!breaks) {
      continue;
    }
    if (!OnlyAttributes(breaks, {"count", "manualBreakCount"})) {
      return false;
    }
    for (const pugi::xml_node& brk : breaks.children()) {
      bool manual = false;
      bool page_break = false;
      if (std::string_view(brk.name()) != "brk" || !OnlyAttributes(brk, {"id", "min", "max", "man", "pt"}) ||
          !ParseBool(brk.attribute("man"), false, &manual) || !ParseBool(brk.attribute("pt"), false, &page_break) ||
          !manual || page_break) {
        return false;
      }
    }
    ooxml::read_manual_breaks(
        breaks, std::string_view(axis) == "rowBreaks" ? print->manual_row_breaks : print->manual_col_breaks);
  }
  return true;
}

}  // namespace

bool encode_book_view(const pugi::xml_node& view, std::vector<std::uint8_t>& dst) {
  if (!OnlyAttributes(view, {"name",
                             "guid",
                             "autoUpdate",
                             "mergeInterval",
                             "changesSavedWin",
                             "onlySync",
                             "personalView",
                             "includePrintSettings",
                             "includeHiddenRowCol",
                             "maximized",
                             "minimized",
                             "showHorizontalScroll",
                             "showVerticalScroll",
                             "showSheetTabs",
                             "xWindow",
                             "yWindow",
                             "windowWidth",
                             "windowHeight",
                             "tabRatio",
                             "activeSheetId",
                             "showFormulaBar",
                             "showStatusbar",
                             "showComments",
                             "showObjects"})) {
    return false;
  }
  std::array<std::uint8_t, kGuidBytes> guid{};
  std::uint32_t x = 0;
  std::uint32_t y = 0;
  std::uint32_t width = 0;
  std::uint32_t height = 0;
  std::uint32_t active_sheet = 0;
  std::uint32_t tab_ratio = 0;
  std::uint32_t merge_interval = 0;
  std::uint32_t comments = 0;
  std::uint32_t objects = 0;
  std::uint32_t flags = 0;
  bool personal = false;
  if (!view.attribute("name") || !view.attribute("windowWidth") || !view.attribute("windowHeight") ||
      !view.attribute("activeSheetId") || !ParseGuidAttr(view, &guid) ||
      !ParseUint(view.attribute("xWindow"), 0U, &x) || !ParseUint(view.attribute("yWindow"), 0U, &y) ||
      !ParseUint(view.attribute("windowWidth"), 0U, &width) ||
      !ParseUint(view.attribute("windowHeight"), 0U, &height) ||
      !ParseUint(view.attribute("activeSheetId"), 0U, &active_sheet) ||
      !ParseUint(view.attribute("tabRatio"), kDefaultTabRatio, &tab_ratio) ||
      !ParseUint(view.attribute("mergeInterval"), kNoMergeInterval, &merge_interval) ||
      !ParseEnum(view.attribute("showComments"), {"commNone", "commIndicator", "commIndAndComment"}, 1U, &comments) ||
      !ParseEnum(view.attribute("showObjects"), {"all", "placeholders", "none"}, 0U, &objects) ||
      !ParseBool(view.attribute("personalView"), false, &personal) ||
      !AddFlags(view, kBookFlags.data(), kBookFlags.size(), &flags)) {
    return false;
  }
  constexpr std::uint64_t kU32Max = std::numeric_limits<std::uint32_t>::max();
  if (tab_ratio > kMaxTabRatio || (view.attribute("mergeInterval") && merge_interval >= kNoMergeInterval) ||
      static_cast<std::uint64_t>(x) + width > kU32Max || static_cast<std::uint64_t>(y) + height > kU32Max) {
    return false;
  }
  flags |= (comments << kShowCommentsShift) | (objects << kShowObjectsShift);
  std::vector<std::uint8_t> payload;
  emit_u32(payload, x);
  emit_u32(payload, x + width);
  emit_u32(payload, y);
  emit_u32(payload, y + height);
  emit_u32(payload, active_sheet);
  emit_u32(payload, tab_ratio);
  payload.insert(payload.end(), guid.begin(), guid.end());
  // Both within u16: the interval is below 0xFFFF and the flags fill sixteen bits.
  emit_u16(payload, static_cast<std::uint16_t>(merge_interval));
  emit_u16(payload, static_cast<std::uint16_t>(flags));
  emit_u16(payload, personal ? 1U : 0U);
  emit_xlwidestring(payload, view.attribute("name").value());
  emit_record(dst, kBrtUserBookView, payload);
  return true;
}

bool encode_sheet_view(const pugi::xml_node& view, bool print_settings, std::vector<std::uint8_t>& dst) {
  if (!OnlyAttributes(view, {"guid",          "scale",        "colorId",        "showPageBreaks", "showFormulas",
                             "showGridLines", "showRowCol",   "outlineSymbols", "zeroValues",     "fitToPage",
                             "printArea",     "filter",       "showAutoFilter", "hiddenRows",     "hiddenColumns",
                             "state",         "filterUnique", "view",           "showRuler",      "topLeftCell"})) {
    return false;
  }
  std::vector<std::string_view> children;
  for (const pugi::xml_node& child : view.children()) {
    const std::string_view name = child.name();
    if ((name != "pane" && name != "selection" && name != "rowBreaks" && name != "colBreaks" && name != "pageMargins" &&
         name != "printOptions" && name != "pageSetup" && name != "headerFooter" && name != "autoFilter") ||
        (name != "selection" && !FirstSeen(&children, name))) {
      return false;
    }
  }
  std::array<std::uint8_t, kGuidBytes> guid{};
  std::uint32_t scale = 0;
  std::uint32_t color = 0;
  std::uint32_t state = 0;
  std::uint32_t mode = 0;
  std::uint32_t flags = kSheetAlwaysFlag;
  if (!ParseGuidAttr(view, &guid) || !ParseUint(view.attribute("scale"), kDefaultScale, &scale) ||
      !ParseUint(view.attribute("colorId"), kDefaultColorId, &color) ||
      !ParseEnum(view.attribute("state"), {"visible", "hidden", "veryHidden"}, 0U, &state) ||
      !ParseEnum(view.attribute("view"), {"normal", "pageBreakPreview", "pageLayout"}, 0U, &mode) ||
      !AddFlags(view, kSheetFlags.data(), kSheetFlags.size(), &flags) || scale < kMinScale || scale > kMaxScale ||
      color > kDefaultColorId) {
    return false;
  }
  constexpr std::array<std::uint32_t, 3> kStateBits = {0U, kSheetStateHidden, kSheetStateVeryHidden};
  constexpr std::array<std::uint32_t, 3> kModeBits = {0U, kSheetViewPageBreakPreview, kSheetViewPageLayout};
  flags |= kStateBits[state] | kModeBits[mode];
  std::uint32_t top_row = 0;
  std::uint32_t top_col = 0;
  if (const pugi::xml_attribute top_left = view.attribute("topLeftCell");
      top_left && !a1::parse_a1_ref(top_left.value(), &top_row, &top_col)) {
    return false;
  }
  SheetView scratch;
  SheetPrintSettings print;
  // A view without its print settings keeps no print records; Excel drops its margins.
  const bool print_children = view.child("printOptions") || view.child("pageSetup") || view.child("headerFooter") ||
                              view.child("rowBreaks") || view.child("colBreaks");
  if (!ReadPaneAndSelections(view, &scratch) || (print_settings && !ReadViewPrint(view, &print, &flags)) ||
      (!print_settings && print_children)) {
    return false;
  }
  // The view's AutoFilter is carried whole or the view is refused.
  std::vector<std::uint8_t> filter_block;
  if (const pugi::xml_node filter_node = view.child("autoFilter")) {
    const AutoFilter filter = auto_filter_from_node(filter_node);
    if (filter.is_opaque() || emit_auto_filter_block(filter_block, filter) != 0U) {
      return false;
    }
  }
  if (view.child("pane")) {
    flags |= kSheetPaneFlags;
  }
  std::vector<std::uint8_t> payload;
  payload.insert(payload.end(), guid.begin(), guid.end());
  emit_u32(payload, kSheetViewMarker);
  emit_u32(payload, scale);
  emit_u32(payload, color);
  emit_u32(payload, flags);
  emit_u32(payload, top_row);
  emit_u32(payload, top_row);
  emit_u32(payload, top_col);
  emit_u32(payload, top_col);
  emit_record(dst, kBrtBeginUserShView, payload);
  emit_view_pane(dst, scratch);
  emit_listed_selections(dst, scratch);
  if (print_settings) {
    emit_view_print_records(dst, print);
  }
  dst.insert(dst.end(), filter_block.begin(), filter_block.end());
  emit_record(dst, kBrtEndUserShView, ByteSpan{});
  return true;
}

std::optional<std::string> decode_book_view(ByteSpan payload) {
  constexpr std::size_t kNameAt = 46U;
  if (payload.size < kNameAt + 4U) {
    return std::nullopt;
  }
  const std::uint8_t* p = payload.data;
  const std::uint32_t units = load_u32(p + kNameAt);
  if (payload.size != kNameAt + 4U + 2U * static_cast<std::uint64_t>(units)) {
    return std::nullopt;
  }
  const std::uint32_t x = load_u32(p);
  const std::uint32_t right = load_u32(p + 4U);
  const std::uint32_t y = load_u32(p + 8U);
  const std::uint32_t bottom = load_u32(p + 12U);
  const std::uint32_t tab_ratio = load_u32(p + 20U);
  const std::uint32_t merge_interval = load_u16(p + 40U);
  const std::uint32_t flags = load_u16(p + 42U);
  const std::uint32_t personal = load_u16(p + 44U);
  const std::uint32_t comments = (flags >> kShowCommentsShift) & 3U;
  const std::uint32_t objects = (flags >> kShowObjectsShift) & 3U;
  if (right < x || bottom < y || tab_ratio > kMaxTabRatio || personal > 1U || comments == 3U || objects == 3U) {
    return std::nullopt;
  }
  const auto flag = [flags](std::string_view name) {
    for (const FlagAttr& entry : kBookFlags) {
      if (entry.name == name) {
        return (flags & entry.mask) != 0U;
      }
    }
    return false;
  };
  std::string name;
  append_utf16le_as_utf8(p + kNameAt + 4U, units, &name);
  std::string xml = "<customWorkbookView";
  append_xml_attr(xml, "name", name);
  append_xml_attr(xml, "guid", FormatGuid(p + 24U));
  // Attributes in schema order, each only when it differs from its default.
  const auto bool_attr = [&](const char* attr, bool fallback) {
    if (flag(attr) != fallback) {
      append_xml_attr(xml, attr, fallback ? "0" : "1");
    }
  };
  bool_attr("autoUpdate", false);
  if (merge_interval != kNoMergeInterval) {
    append_xml_attr_uint(xml, "mergeInterval", merge_interval);
  }
  bool_attr("changesSavedWin", false);
  bool_attr("onlySync", false);
  if (personal != 0U) {
    append_xml_attr(xml, "personalView", "1");
  }
  bool_attr("includePrintSettings", true);
  bool_attr("includeHiddenRowCol", true);
  bool_attr("maximized", false);
  bool_attr("minimized", false);
  bool_attr("showHorizontalScroll", true);
  bool_attr("showVerticalScroll", true);
  bool_attr("showSheetTabs", true);
  if (x != 0U) {
    append_xml_attr_uint(xml, "xWindow", x);
  }
  if (y != 0U) {
    append_xml_attr_uint(xml, "yWindow", y);
  }
  append_xml_attr_uint(xml, "windowWidth", right - x);
  append_xml_attr_uint(xml, "windowHeight", bottom - y);
  if (tab_ratio != kDefaultTabRatio) {
    append_xml_attr_uint(xml, "tabRatio", tab_ratio);
  }
  append_xml_attr_uint(xml, "activeSheetId", load_u32(p + 16U));
  bool_attr("showFormulaBar", true);
  bool_attr("showStatusbar", true);
  constexpr std::array<const char*, 3> kComments = {"commNone", "commIndicator", "commIndAndComment"};
  constexpr std::array<const char*, 3> kObjects = {"all", "placeholders", "none"};
  if (comments != 1U) {
    append_xml_attr(xml, "showComments", kComments[comments]);
  }
  if (objects != 0U) {
    append_xml_attr(xml, "showObjects", kObjects[objects]);
  }
  xml.append("/>");
  return xml;
}

std::optional<std::string> decode_sheet_view(const std::vector<FramedRecord>& records,
                                             const std::vector<std::uint8_t>& body) {
  constexpr std::size_t kViewSize = 48U;
  if (records.size() < 2U || records.front().type != kBrtBeginUserShView || records.back().type != kBrtEndUserShView ||
      records.front().payload.size != kViewSize) {
    return std::nullopt;
  }
  const std::uint8_t* p = records.front().payload.data;
  const std::uint32_t scale = load_u32(p + 20U);
  const std::uint32_t color = load_u32(p + 24U);
  const std::uint32_t flags = load_u32(p + 28U);
  std::uint32_t known = kSheetPaneFlags | kSheetAlwaysFlag | kSheetStateHidden | kSheetStateVeryHidden |
                        kSheetViewPageBreakPreview | kSheetViewPageLayout;
  for (const FlagAttr& entry : kSheetFlags) {
    known |= entry.mask;
  }
  for (const FlagAttr& entry : kSheetPrintFlags) {
    known |= entry.mask;
  }
  const std::uint32_t pane_bits = flags & kSheetPaneFlags;
  if (load_u32(p + 16U) != kSheetViewMarker || scale < kMinScale || scale > kMaxScale || color > kDefaultColorId ||
      (flags & ~known) != 0U || (flags & kSheetAlwaysFlag) == 0U || (pane_bits != 0U && pane_bits != kSheetPaneFlags) ||
      (flags & (kSheetStateHidden | kSheetStateVeryHidden)) == (kSheetStateHidden | kSheetStateVeryHidden) ||
      (flags & (kSheetViewPageBreakPreview | kSheetViewPageLayout)) ==
          (kSheetViewPageBreakPreview | kSheetViewPageLayout) ||
      load_u32(p + 32U) != load_u32(p + 36U) || load_u32(p + 40U) != load_u32(p + 44U)) {
    return std::nullopt;
  }
  SheetView scratch;
  std::vector<std::uint8_t> print_records;
  std::optional<std::string> print_options = std::string();
  std::string filter_xml;
  bool has_pane = false;
  for (std::size_t i = 1; i + 1U < records.size(); ++i) {
    const FramedRecord& rec = records[i];
    const XlsbRecord record{rec.type, rec.payload};
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtPane)) {
      auto pane = decode_pane(record, scratch);
      if (has_pane || !pane || !pane.value()) {
        return std::nullopt;
      }
      has_pane = true;
    } else if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtSel)) {
      if (!decode_selection(record, scratch)) {
        return std::nullopt;
      }
    } else if (rec.type == kBrtMargins || rec.type == kBrtPageSetup || rec.type == kBrtBeginHeaderFooter ||
               rec.type == kBrtEndHeaderFooter || rec.type == kBrtBeginRwBrk || rec.type == kBrtEndRwBrk ||
               rec.type == kBrtBeginColBrk || rec.type == kBrtEndColBrk || rec.type == kBrtBrk) {
      print_records.insert(print_records.end(), body.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                           body.begin() + static_cast<std::ptrdiff_t>(rec.end));
    } else if (rec.type == kBrtPrintOptions) {
      print_options = decode_print_options(rec.payload);
      if (!print_options) {
        return std::nullopt;
      }
    } else if (rec.type == kAcBegin || rec.type == kBeginAFilter) {
      // A revision-uid wrapper leads the filter's block, which is the last thing in the view.
      const std::size_t first = rec.type == kAcBegin ? i + 3U : i;
      if (!filter_xml.empty() || first + 1U >= records.size() ||
          (first != i && !(rec.payload.size == kAcUidBeginPayload.size() &&
                           std::equal(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end(), rec.payload.data) &&
                           records[i + 1U].type == kUid && records[i + 1U].payload.size == kGuidBytes &&
                           records[i + 2U].type == kAcEnd && records[i + 2U].payload.size == 0U)) ||
          records[first].type != kBeginAFilter) {
        return std::nullopt;
      }
      // The cursor spans from the record after BrtBeginAFilter to the view's end.
      const std::size_t from = records[first + 1U].begin;
      ByteSpan cursor{body.data() + from, records.back().begin - from};
      auto filter = read_auto_filter_block(cursor, records[first].payload);
      if (!filter || filter.value().is_opaque()) {
        return std::nullopt;
      }
      const std::size_t block_end = from + static_cast<std::size_t>(cursor.data - (body.data() + from));
      std::size_t last = first + 1U;
      while (last + 1U < records.size() && records[last].end < block_end) {
        ++last;
      }
      if (records[last].end != block_end || records[last].type != kBrtEndAFilter || last + 2U != records.size()) {
        return std::nullopt;
      }
      filter_xml = serialize_auto_filter(filter.value());
      i = last;
    } else {
      return std::nullopt;
    }
  }
  if (has_pane != (pane_bits != 0U) || (has_pane && !(scratch.divides_rows() && scratch.divides_cols()))) {
    return std::nullopt;
  }
  SheetPrintSettings print;
  decode_print_records(print_records, print);
  if (unwritten_print_families(print_records, print) != 0U) {
    return std::nullopt;
  }
  // The view's flags repeat its print options.
  std::uint32_t option_bits = 0;
  if (!print_options->empty()) {
    pugi::xml_document doc;
    if (!doc.load_string(print_options->c_str()) ||
        !AddFlags(doc.document_element(), kSheetPrintFlags.data(), kSheetPrintFlags.size(), &option_bits)) {
      return std::nullopt;
    }
  }
  std::uint32_t print_flags = 0;
  for (const FlagAttr& entry : kSheetPrintFlags) {
    print_flags |= entry.mask;
  }
  if ((flags & print_flags) != option_bits) {
    return std::nullopt;
  }
  std::string xml = "<customSheetView";
  append_xml_attr(xml, "guid", FormatGuid(p));
  if (scale != kDefaultScale) {
    append_xml_attr_uint(xml, "scale", scale);
  }
  if (color != kDefaultColorId) {
    append_xml_attr_uint(xml, "colorId", color);
  }
  const auto bool_attr = [&](std::string_view name) {
    for (const FlagAttr& entry : kSheetFlags) {
      if (entry.name == name && ((flags & entry.mask) != 0U) != entry.fallback) {
        append_xml_attr(xml, std::string(name).c_str(), entry.fallback ? "0" : "1");
      }
    }
  };
  for (const char* name :
       {"showPageBreaks", "showFormulas", "showGridLines", "showRowCol", "outlineSymbols", "zeroValues", "fitToPage",
        "printArea", "filter", "showAutoFilter", "hiddenRows", "hiddenColumns"}) {
    bool_attr(name);
  }
  if ((flags & kSheetStateHidden) != 0U) {
    append_xml_attr(xml, "state", "hidden");
  } else if ((flags & kSheetStateVeryHidden) != 0U) {
    append_xml_attr(xml, "state", "veryHidden");
  }
  bool_attr("filterUnique");
  if ((flags & kSheetViewPageBreakPreview) != 0U) {
    append_xml_attr(xml, "view", "pageBreakPreview");
  } else if ((flags & kSheetViewPageLayout) != 0U) {
    append_xml_attr(xml, "view", "pageLayout");
  }
  bool_attr("showRuler");
  const std::uint32_t top_row = load_u32(p + 32U);
  const std::uint32_t top_col = load_u32(p + 40U);
  if (top_row != 0U || top_col != 0U) {
    append_xml_attr(xml, "topLeftCell", a1::encode_a1(top_row, top_col));
  }
  const std::string children = build_pane_and_selections_xml(scratch, true) +
                               ooxml::build_page_breaks_xml("rowBreaks", print.manual_row_breaks) +
                               ooxml::build_page_breaks_xml("colBreaks", print.manual_col_breaks) +
                               print.page_margins_xml + *print_options + print.page_setup_xml +
                               print.header_footer_xml + filter_xml;
  if (children.empty()) {
    xml.append("/>");
  } else {
    xml.append(">").append(children).append("</customSheetView>");
  }
  return xml;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
