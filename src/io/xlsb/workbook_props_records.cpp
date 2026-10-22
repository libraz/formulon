#include "io/xlsb/workbook_props_records.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <string_view>
#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// BrtCalcProp flag word (2.4.318).
constexpr std::uint16_t kCalcFullCalcOnLoad = 1U << 0U;
constexpr std::uint16_t kCalcRefA1 = 1U << 1U;
constexpr std::uint16_t kCalcIter = 1U << 2U;
constexpr std::uint16_t kCalcFullPrec = 1U << 3U;
constexpr std::uint16_t kCalcSomeUncalced = 1U << 4U;
constexpr std::uint16_t kCalcSaveRecalc = 1U << 5U;
constexpr std::uint16_t kCalcMtrEnabled = 1U << 6U;
constexpr std::uint16_t kCalcUserSetThreadCount = 1U << 7U;
constexpr std::uint16_t kCalcNoDeps = 1U << 8U;
// cUserThreadCount must lie in [1, 1024]; 1 is what Excel writes unset.
constexpr std::int32_t kDefaultThreadCount = 1;

// BrtWbProp flag word (2.4.868).
constexpr std::uint32_t kWb1904 = 1U << 0U;
constexpr std::uint32_t kWbUpdateLinksShift = 8U;
constexpr std::uint32_t kWbShowObjectsShift = 13U;
constexpr std::uint32_t kWbTwoBitMask = 3U;
// fShowInkAnnotation | fAutoCompressPictures: the schema defaults.
constexpr std::uint32_t kWbDefaultFlags = 0x00010020U;
constexpr std::uint32_t kWbDefaultThemeVersion = 202300U;

/// A one-bit BrtWbProp field and its `<workbookPr>` attribute. `inverted`
/// fields store the negation of the attribute.
struct WbBool {
  const char* attr;
  std::uint32_t bit;
  bool inverted;
  bool attr_default;
};

constexpr std::array<WbBool, 12> kWbBools = {{
    {"showBorderUnselectedTables", 1U << 2U, true, true},
    {"filterPrivacy", 1U << 3U, false, false},
    {"promptedSolutions", 1U << 4U, false, false},
    {"showInkAnnotation", 1U << 5U, false, true},
    {"backupFile", 1U << 6U, false, false},
    {"saveExternalLinkValues", 1U << 7U, true, true},
    {"hidePivotFieldList", 1U << 10U, false, false},
    {"publishItems", 1U << 11U, false, false},
    {"checkCompatibility", 1U << 12U, false, false},
    {"showPivotChartFilter", 1U << 15U, false, false},
    {"autoCompressPictures", 1U << 16U, false, true},
    {"refreshAllConnections", 1U << 18U, false, false},
}};

constexpr std::array<std::string_view, 3> kUpdateLinks = {"userSet", "never", "always"};
constexpr std::array<std::string_view, 3> kShowObjects = {"all", "placeholders", "none"};

const WbBool* FindWbBool(std::string_view attr) {
  for (const WbBool& field : kWbBools) {
    if (attr == field.attr) {
      return &field;
    }
  }
  return nullptr;
}

bool WbBoolValue(const WbBool& field, std::uint32_t flags) {
  return ((flags & field.bit) != 0U) != field.inverted;
}

void AppendWbBool(std::string& xml, std::string_view attr, std::uint32_t flags) {
  const WbBool& field = *FindWbBool(attr);
  const bool value = WbBoolValue(field, flags);
  if (value != field.attr_default) {
    emit_xsd_bool_attr(xml, field.attr, value);
  }
}

/// Index of `text` in `names`, or `fallback` when absent.
template <std::size_t N>
std::uint32_t EnumIndex(const std::array<std::string_view, N>& names, std::string_view text, std::uint32_t fallback,
                        bool& known) {
  for (std::size_t i = 0; i < N; ++i) {
    if (names[i] == text) {
      return static_cast<std::uint32_t>(i);
    }
  }
  known = false;
  return fallback;
}

// BrtBookView (2.4.313): seven u32 fields, then a flag byte.
constexpr std::uint16_t kBrtBeginBookViews = 135;
constexpr std::uint16_t kBrtEndBookViews = 136;
constexpr std::uint8_t kViewHidden = 1U << 0U;
constexpr std::uint8_t kViewVeryHidden = 1U << 1U;
constexpr std::uint8_t kViewIconic = 1U << 2U;
constexpr std::uint8_t kViewHScroll = 1U << 3U;
constexpr std::uint8_t kViewVScroll = 1U << 4U;
constexpr std::uint8_t kViewSheetTabs = 1U << 5U;
constexpr std::uint8_t kViewDateGroup = 1U << 6U;
constexpr std::uint32_t kDefaultTabRatio = 600U;
// The window Excel 365 (macOS) writes for a new book; used for absent
// position and size attributes.
constexpr std::int32_t kDefaultXWindow = 20240;
constexpr std::int32_t kDefaultYWindow = 5400;
constexpr std::uint32_t kDefaultWindowWidth = 28300;
constexpr std::uint32_t kDefaultWindowHeight = 17360;

/// A one-bit BrtBookView field and its `<workbookView>` attribute.
struct ViewBool {
  const char* attr;
  std::uint8_t bit;
  bool attr_default;
};

constexpr std::array<ViewBool, 5> kViewBools = {{
    {"minimized", kViewIconic, false},
    {"showHorizontalScroll", kViewHScroll, true},
    {"showVerticalScroll", kViewVScroll, true},
    {"showSheetTabs", kViewSheetTabs, true},
    {"autoFilterDateGrouping", kViewDateGroup, true},
}};

void AppendViewBool(std::string& xml, const char* attr, std::uint8_t flags) {
  for (const ViewBool& field : kViewBools) {
    if (std::string_view(attr) == field.attr) {
      const bool value = (flags & field.bit) != 0U;
      if (value != field.attr_default) {
        emit_xsd_bool_attr(xml, field.attr, value);
      }
    }
  }
}

/// Emits one BrtBookView for `node` (null: the default window). False when
/// `node` carries something the record cannot hold.
bool EmitBookView(std::vector<std::uint8_t>& dst, const pugi::xml_node& node, std::uint32_t itab_cur,
                  std::size_t sheet_count) {
  const auto sheet_index = [&](std::uint32_t index) { return index < sheet_count ? index : 0U; };
  std::vector<std::uint8_t> p;
  emit_u32(p, static_cast<std::uint32_t>(node ? attr_i32(node, "xWindow", kDefaultXWindow) : kDefaultXWindow));
  emit_u32(p, static_cast<std::uint32_t>(node ? attr_i32(node, "yWindow", kDefaultYWindow) : kDefaultYWindow));
  emit_u32(p, node ? attr_u32(node, "windowWidth", kDefaultWindowWidth) : kDefaultWindowWidth);
  emit_u32(p, node ? attr_u32(node, "windowHeight", kDefaultWindowHeight) : kDefaultWindowHeight);
  emit_u32(p, node ? attr_u32(node, "tabRatio", kDefaultTabRatio) : kDefaultTabRatio);
  emit_u32(p, sheet_index(node ? attr_u32(node, "firstSheet", 0U) : 0U));
  emit_u32(p, sheet_index(itab_cur));
  std::uint8_t flags = 0U;
  bool carried = true;
  for (const ViewBool& field : kViewBools) {
    if (node ? read_xsd_bool(node, field.attr, field.attr_default) : field.attr_default) {
      flags |= field.bit;
    }
  }
  if (node) {
    const std::string_view visibility = attr_str(node, "visibility", "visible");
    // A very hidden window also has the properties of a hidden one.
    flags |= visibility == "hidden" ? kViewHidden : visibility == "veryHidden" ? kViewHidden | kViewVeryHidden : 0U;
    carried = visibility == "visible" || visibility == "hidden" || visibility == "veryHidden";
    for (const pugi::xml_attribute attr : node.attributes()) {
      const std::string_view name = attr.name();
      const bool known =
          name == "visibility" || name == "xWindow" || name == "yWindow" || name == "windowWidth" ||
          name == "windowHeight" || name == "tabRatio" || name == "firstSheet" || name == "activeTab" ||
          name == "xr2:uid" ||
          std::any_of(kViewBools.begin(), kViewBools.end(), [&](const ViewBool& field) { return name == field.attr; });
      carried = carried && known;
    }
    carried = carried && !node.first_child();
  }
  emit_u8(p, flags);
  emit_record(dst, kBrtBookView, p);
  return carried;
}

}  // namespace

bool decode_calc_prop(ByteSpan payload, XlsbCalcProps& out) {
  ByteSpan p = payload;
  auto recalc_id = read_u32(p);
  auto auto_recalc = read_u32(p);
  auto count = read_u32(p);
  auto delta = read_double(p);
  auto threads = read_u32(p);
  auto flags = read_u16(p);
  if (!recalc_id || !auto_recalc || !count || !delta || !threads || !flags) {
    return false;
  }
  XlsbCalcProps props;
  // An undefined fAutoRecalc falls back to automatic, as `<calcPr>` does.
  props.mode = auto_recalc.value() == 0U   ? CalcMode::kManual
               : auto_recalc.value() == 2U ? CalcMode::kAutoNoTable
                                           : CalcMode::kAuto;
  const std::uint16_t f = flags.value();
  props.iterative.enabled = (f & kCalcIter) != 0U;
  props.iterative.max_iterations = count.value() < 1U                  ? 1U
                                   : count.value() > kMaxIterationsCap ? kMaxIterationsCap
                                                                       : count.value();
  if (std::isfinite(delta.value()) && delta.value() >= 0.0) {
    props.iterative.max_change = delta.value();
  }
  CalcOptions& o = props.options;
  o.has_calc_id = recalc_id.value() != 0U;
  o.calc_id = recalc_id.value();
  o.full_calc_on_load = (f & kCalcFullCalcOnLoad) != 0U;
  o.ref_mode = (f & kCalcRefA1) != 0U ? RefMode::kA1 : RefMode::kR1C1;
  o.full_precision = (f & kCalcFullPrec) != 0U;
  o.calc_completed = (f & kCalcSomeUncalced) == 0U;
  o.calc_on_save = (f & kCalcSaveRecalc) != 0U;
  o.concurrent_calc = (f & kCalcMtrEnabled) != 0U;
  o.has_concurrent_manual_count = (f & kCalcUserSetThreadCount) != 0U;
  o.concurrent_manual_count = o.has_concurrent_manual_count ? threads.value() : 0U;
  o.force_full_calc = (f & kCalcNoDeps) != 0U;
  out = props;
  return true;
}

bool calc_props_are_default(const XlsbCalcProps& props) {
  const IterativeOptions& iter = props.iterative;
  return props.mode == CalcMode::kAuto && !iter.enabled && iter.max_iterations == kDefaultMaxIterations &&
         iter.max_change == kDefaultMaxChange && props.options.is_default();
}

void emit_calc_prop(std::vector<std::uint8_t>& dst, const XlsbCalcProps& props) {
  const CalcOptions& o = props.options;
  std::vector<std::uint8_t> p;
  emit_u32(p, o.has_calc_id ? o.calc_id : 0U);
  emit_u32(p, props.mode == CalcMode::kManual ? 0U : props.mode == CalcMode::kAutoNoTable ? 2U : 1U);
  emit_u32(p, props.iterative.max_iterations);
  emit_double(p, props.iterative.max_change);
  emit_u32(p,
           o.has_concurrent_manual_count ? o.concurrent_manual_count : static_cast<std::uint32_t>(kDefaultThreadCount));
  std::uint16_t f = 0U;
  f |= o.full_calc_on_load ? kCalcFullCalcOnLoad : 0U;
  f |= o.ref_mode == RefMode::kA1 ? kCalcRefA1 : 0U;
  f |= props.iterative.enabled ? kCalcIter : 0U;
  f |= o.full_precision ? kCalcFullPrec : 0U;
  f |= o.calc_completed ? 0U : kCalcSomeUncalced;
  f |= o.calc_on_save ? kCalcSaveRecalc : 0U;
  f |= o.concurrent_calc ? kCalcMtrEnabled : 0U;
  f |= o.has_concurrent_manual_count ? kCalcUserSetThreadCount : 0U;
  f |= o.force_full_calc ? kCalcNoDeps : 0U;
  emit_u16(p, f);
  emit_record(dst, kBrtCalcProp, p);
}

bool decode_wb_prop(ByteSpan payload, std::string& workbook_pr_xml, bool& date1904) {
  ByteSpan p = payload;
  auto flags_or = read_u32(p);
  if (!flags_or) {
    return false;
  }
  const std::uint32_t flags = flags_or.value();
  std::uint32_t theme_version = 0U;
  std::string code_name;
  // A stream that stops early leaves the trailing fields at their defaults.
  if (auto theme = read_u32(p); theme) {
    theme_version = theme.value();
    if (auto name = read_xlwidestring(p); name) {
      code_name = std::move(name).value();
    }
  }
  date1904 = (flags & kWb1904) != 0U;
  // Attributes in CT_WorkbookPr schema order.
  std::string xml = "<workbookPr";
  const std::size_t empty_size = xml.size();
  if (date1904) {
    emit_xsd_bool_attr(xml, "date1904", true);
  }
  const std::uint32_t show_objects = (flags >> kWbShowObjectsShift) & kWbTwoBitMask;
  if (show_objects != 0U && show_objects < kShowObjects.size()) {
    append_xml_attr(xml, "showObjects", kShowObjects[show_objects]);
  }
  for (const char* attr : {"showBorderUnselectedTables", "filterPrivacy", "promptedSolutions", "showInkAnnotation",
                           "backupFile", "saveExternalLinkValues"}) {
    AppendWbBool(xml, attr, flags);
  }
  const std::uint32_t update_links = (flags >> kWbUpdateLinksShift) & kWbTwoBitMask;
  if (update_links != 0U && update_links < kUpdateLinks.size()) {
    append_xml_attr(xml, "updateLinks", kUpdateLinks[update_links]);
  }
  if (!code_name.empty()) {
    append_xml_attr(xml, "codeName", code_name);
  }
  for (const char* attr : {"hidePivotFieldList", "showPivotChartFilter", "publishItems", "checkCompatibility",
                           "autoCompressPictures", "refreshAllConnections"}) {
    AppendWbBool(xml, attr, flags);
  }
  if (theme_version != 0U) {
    append_xml_attr_uint(xml, "defaultThemeVersion", theme_version);
  }
  if (xml.size() == empty_size) {
    workbook_pr_xml.clear();
    return true;
  }
  xml.append("/>");
  workbook_pr_xml = std::move(xml);
  return true;
}

bool emit_wb_prop(std::vector<std::uint8_t>& dst, const std::string& workbook_pr_xml, bool date1904) {
  std::uint32_t flags = kWbDefaultFlags;
  std::uint32_t theme_version = kWbDefaultThemeVersion;
  std::string code_name;
  bool carried = true;
  if (!workbook_pr_xml.empty()) {
    pugi::xml_document doc;
    const pugi::xml_node node =
        doc.load_buffer(workbook_pr_xml.data(), workbook_pr_xml.size()) ? doc.child("workbookPr") : pugi::xml_node();
    if (!node) {
      carried = false;
    } else {
      theme_version = attr_u32(node, "defaultThemeVersion", 0U);
      for (const WbBool& field : kWbBools) {
        const bool value = read_xsd_bool(node, field.attr, field.attr_default);
        flags = (value != field.inverted) ? (flags | field.bit) : (flags & ~field.bit);
      }
      const std::uint32_t update_links = EnumIndex(kUpdateLinks, attr_str(node, "updateLinks", "userSet"), 0U, carried);
      const std::uint32_t show_objects = EnumIndex(kShowObjects, attr_str(node, "showObjects", "all"), 0U, carried);
      flags |= update_links << kWbUpdateLinksShift;
      flags |= show_objects << kWbShowObjectsShift;
      code_name = std::string(attr_str(node, "codeName"));
      for (const pugi::xml_attribute attr : node.attributes()) {
        const std::string_view name = attr.name();
        if (FindWbBool(name) != nullptr || name == "date1904" || name == "1904" || name == "updateLinks" ||
            name == "showObjects" || name == "codeName" || name == "defaultThemeVersion") {
          continue;
        }
        // No field: only a non-default value is a loss.
        if (name == "dateCompatibility" && read_xsd_bool(node, "dateCompatibility", true)) {
          continue;
        }
        if (name == "allowRefreshQuery" && !read_xsd_bool(node, "allowRefreshQuery", false)) {
          continue;
        }
        carried = false;
      }
    }
  }
  flags = date1904 ? (flags | kWb1904) : (flags & ~kWb1904);
  std::vector<std::uint8_t> p;
  emit_u32(p, flags);
  emit_u32(p, theme_version);
  emit_xlwidestring(p, code_name);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp), p);
  return carried;
}

bool decode_book_view(ByteSpan payload, std::string& xml) {
  ByteSpan p = payload;
  std::array<std::uint32_t, 7> fields{};
  for (std::uint32_t& field : fields) {
    auto v = read_u32(p);
    if (!v) {
      return false;
    }
    field = v.value();
  }
  auto flags_or = read_u8(p);
  if (!flags_or) {
    return false;
  }
  const std::uint8_t flags = flags_or.value();
  // Attributes in CT_BookView schema order.
  xml.append("<workbookView");
  if ((flags & kViewVeryHidden) != 0U) {
    append_xml_attr(xml, "visibility", "veryHidden");
  } else if ((flags & kViewHidden) != 0U) {
    append_xml_attr(xml, "visibility", "hidden");
  }
  for (const char* attr : {"minimized", "showHorizontalScroll", "showVerticalScroll", "showSheetTabs"}) {
    AppendViewBool(xml, attr, flags);
  }
  append_xml_attr_int(xml, "xWindow", static_cast<std::int32_t>(fields[0]));
  append_xml_attr_int(xml, "yWindow", static_cast<std::int32_t>(fields[1]));
  append_xml_attr_uint(xml, "windowWidth", fields[2]);
  append_xml_attr_uint(xml, "windowHeight", fields[3]);
  if (fields[4] != kDefaultTabRatio) {
    append_xml_attr_uint(xml, "tabRatio", fields[4]);
  }
  if (fields[5] != 0U) {
    append_xml_attr_uint(xml, "firstSheet", fields[5]);
  }
  if (fields[6] != 0U) {
    append_xml_attr_uint(xml, "activeTab", fields[6]);
  }
  AppendViewBool(xml, "autoFilterDateGrouping", flags);
  xml.append("/>");
  return true;
}

std::uint32_t workbook_view_count(const std::string& book_views_xml) {
  pugi::xml_document doc;
  if (book_views_xml.empty() || !doc.load_buffer(book_views_xml.data(), book_views_xml.size())) {
    return 1U;
  }
  std::uint32_t count = 0U;
  for (pugi::xml_node node = doc.child("bookViews").child("workbookView"); node;
       node = node.next_sibling("workbookView")) {
    ++count;
  }
  return count == 0U ? 1U : count;
}

bool emit_book_views(std::vector<std::uint8_t>& dst, const std::string& book_views_xml, std::uint32_t active_sheet,
                     std::size_t sheet_count) {
  emit_record(dst, kBrtBeginBookViews, ByteSpan{});
  bool carried = true;
  pugi::xml_document doc;
  pugi::xml_node first;
  if (!book_views_xml.empty()) {
    if (doc.load_buffer(book_views_xml.data(), book_views_xml.size()) && doc.child("bookViews")) {
      first = doc.child("bookViews").child("workbookView");
    } else {
      carried = false;
    }
  }
  if (!first) {
    EmitBookView(dst, pugi::xml_node(), active_sheet, sheet_count);
  }
  for (pugi::xml_node node = first; node; node = node.next_sibling("workbookView")) {
    const std::uint32_t itab_cur = node == first ? active_sheet : attr_u32(node, "activeTab", 0U);
    carried = EmitBookView(dst, node, itab_cur, sheet_count) && carried;
  }
  emit_record(dst, kBrtEndBookViews, ByteSpan{});
  return carried;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
