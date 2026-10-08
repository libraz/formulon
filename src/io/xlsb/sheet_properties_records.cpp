#include "io/xlsb/sheet_properties_records.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>
#include <utility>

#include "io/xlsb/brt_color.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/number_text.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// The `BrtWsProp` fields `emit_ws_prop`
/// writes for a sheet with no source record of its own. A source record
/// matching them carries nothing beyond what the writer will re-derive;
/// one that differs carries flags (dialog-sheet, fit-to-page, outline
/// direction, filter mode, sync anchors) the model has no field for.
constexpr std::uint16_t kDefaultWsPropFlags = 0x04C9U;
constexpr std::uint8_t kDefaultWsPropFlags2 = 0x02U;
constexpr std::uint32_t kWsPropSyncUnused = 0xFFFFFFFFU;

/// BrtWsFmtInfo: the dxGCol sentinel for "no default column width" and the
/// row height Excel writes when no default is set.
constexpr std::uint32_t kAbsentDefaultColumnWidth = 0xFFFFFFFFU;
constexpr std::uint16_t kCanonicalDefaultRowHeightTwips = 300U;

/// Appends `<tabColor .../>` for the decoded `BrtColor`, or nothing when
/// the colour is automatic -- Excel writes no `<tabColor>` for the
/// default tab, and an explicit `auto="1"` would make the two containers
/// disagree on an otherwise identical sheet.
void AppendTabColor(std::string& out, const ColorSpec& spec, std::uint32_t argb) {
  char buf[16];
  switch (spec.kind) {
    case ColorSpec::Kind::kRgb:
      format_hex(buf, sizeof(buf), argb, 8, true);
      out.append("<tabColor rgb=\"").append(buf).append("\"/>");
      break;
    case ColorSpec::Kind::kTheme:
      format_unsigned(buf, sizeof(buf), spec.theme);
      out.append("<tabColor theme=\"").append(buf).push_back('"');
      if (spec.tint != 0.0) {
        // Shortest round-trip spelling, matching the styles writer: the
        // same tint must not be spelled two ways depending on whether the
        // sheet came in as XLSB or XLSX.
        out.append(" tint=\"");
        append_xml_number(out, spec.tint);
        out.push_back('"');
      }
      out.append("/>");
      break;
    case ColorSpec::Kind::kIndexed:
      format_unsigned(buf, sizeof(buf), spec.indexed);
      out.append("<tabColor indexed=\"").append(buf).append("\"/>");
      break;
    case ColorSpec::Kind::kAuto:
    case ColorSpec::Kind::kNone:
      break;
  }
}

/// The two `<sheetPr>` members `BrtWsProp` carries: the VBA code name and
/// the tab colour, in the wire form the record wants them.
///
/// `color_type` is the `XColorType` selector; `0` is automatic, which
/// together with the automatic palette index is the "no tab colour" state
/// Excel writes for an untinted tab.
struct WorksheetProperties {
  std::string code_name;
  std::uint8_t color_type = 0U;
  std::uint8_t color_index = 0x40U;
  std::int16_t color_tint = 0;
  std::uint32_t color_argb = 0U;
};

/// Reads the sheet's retained `<sheetPr>` fragment back into the fields
/// `BrtWsProp` can express. The fragment is the same string the OOXML
/// writer emits and the XLSB reader synthesises, so this is the inverse of
/// `decode_ws_prop` and the two containers agree
/// on a sheet's code name and tab colour whichever one it was loaded from.
///
/// Members `<sheetPr>` can carry that `BrtWsProp` has no field for --
/// `<outlinePr>`, `<pageSetUpPr>`, `filterMode` -- stay in the fragment
/// and reach an `.xlsx` save unchanged; they are simply not part of what
/// this record emits.
WorksheetProperties ParseSheetProperties(std::string_view raw) {
  WorksheetProperties out;
  if (raw.empty()) {
    return out;
  }
  pugi::xml_document doc;
  if (!doc.load_buffer(raw.data(), raw.size())) {
    return out;
  }
  const pugi::xml_node sheet_pr = doc.document_element();
  if (!sheet_pr || std::string_view(sheet_pr.name()) != "sheetPr") {
    return out;
  }
  out.code_name = sheet_pr.attribute("codeName").value();
  const pugi::xml_node tab = sheet_pr.child("tabColor");
  if (!tab) {
    return out;
  }
  if (const pugi::xml_attribute rgb = tab.attribute("rgb")) {
    out.color_type = 2U;
    out.color_index = 0U;
    out.color_argb = static_cast<std::uint32_t>(std::strtoul(rgb.value(), nullptr, 16));
    return out;
  }
  if (const pugi::xml_attribute theme = tab.attribute("theme")) {
    out.color_type = 3U;
    out.color_index = static_cast<std::uint8_t>(std::min<unsigned>(theme.as_uint(0U), 0xFFU));
    const double tint = std::clamp(std::round(tab.attribute("tint").as_double(0.0) * 32767.0), -32767.0, 32767.0);
    out.color_tint = static_cast<std::int16_t>(tint);
    return out;
  }
  if (const pugi::xml_attribute indexed = tab.attribute("indexed")) {
    out.color_type = 1U;
    out.color_index = static_cast<std::uint8_t>(std::min<unsigned>(indexed.as_uint(0U), 0xFFU));
    return out;
  }
  return out;
}

}  // namespace

/// Decodes `BrtWsProp` ([MS-XLSB] §2.4.858) into the sheet's raw
/// `<sheetPr>` fragment. Layout, verified against a real
/// Excel-365-produced `xl/worksheets/sheetN.bin` and symmetric with
/// `emit_ws_prop`: three flag bytes, an 8-byte `BrtColor` tab colour,
/// `rwSync` and `colSync`, then the VBA `CodeName` as an XLWideString.
///
/// The record reaches the model as XML rather than as typed fields
/// because `<sheetPr>` is what `SheetPrintSettings` stores and what the
/// introspection API hands back: routing the XLSB form through the same
/// string makes a `.xlsb`-loaded sheet answer exactly as its `.xlsx` twin
/// does. A sheet with neither a code name nor a tab colour produces no
/// fragment at all, which is also what the OOXML reader records for a
/// worksheet with no `<sheetPr>` element.
///
/// Returns `false` when the record carried flag or sync fields the model
/// has no room for, so the caller can report the residue.
Expected<bool, Error> decode_ws_prop(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index) {
  ByteSpan p = rec.payload;
  auto flags_or = read_u16(p);
  auto flags2_or = read_u8(p);
  auto color_flags_or = read_u8(p);
  auto color_index_or = read_u8(p);
  auto color_tint_or = read_u16(p);
  auto red_or = read_u8(p);
  auto green_or = read_u8(p);
  auto blue_or = read_u8(p);
  auto alpha_or = read_u8(p);
  auto rw_sync_or = read_u32(p);
  auto col_sync_or = read_u32(p);
  if (!flags_or || !flags2_or || !color_flags_or || !color_index_or || !color_tint_or || !red_or || !green_or ||
      !blue_or || !alpha_or || !rw_sync_or || !col_sync_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsProp fields truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }
  auto code_name_or = read_xlwidestring(p);
  if (!code_name_or) {
    return std::move(code_name_or.error());
  }

  const std::uint32_t argb =
      (static_cast<std::uint32_t>(alpha_or.value()) << 24U) | (static_cast<std::uint32_t>(red_or.value()) << 16U) |
      (static_cast<std::uint32_t>(green_or.value()) << 8U) | static_cast<std::uint32_t>(blue_or.value());
  ColorSpec tab;
  switch (static_cast<std::uint32_t>(color_flags_or.value() >> 1U)) {
    case 1U:
      tab.kind = ColorSpec::Kind::kIndexed;
      tab.indexed = color_index_or.value();
      break;
    case 2U:
      tab.kind = ColorSpec::Kind::kRgb;
      tab.rgb = argb;
      break;
    case 3U:
      tab.kind = ColorSpec::Kind::kTheme;
      tab.theme = color_index_or.value();
      tab.tint = static_cast<double>(static_cast<std::int16_t>(color_tint_or.value())) / 32767.0;
      break;
    default:
      tab.kind = ColorSpec::Kind::kAuto;
      break;
  }
  // Excel marks the default tab with the automatic palette slot rather
  // than the automatic colour type, so both spellings mean "no tab
  // colour" and neither produces a `<tabColor>` element.
  constexpr std::uint8_t kAutomaticPaletteIndex = 0x40U;
  if (tab.kind == ColorSpec::Kind::kIndexed && color_index_or.value() == kAutomaticPaletteIndex) {
    tab.kind = ColorSpec::Kind::kAuto;
  }

  std::string sheet_pr;
  const std::string& code_name = code_name_or.value();
  std::string body;
  AppendTabColor(body, tab, argb);
  if (!code_name.empty() || !body.empty()) {
    sheet_pr.append("<sheetPr");
    if (!code_name.empty()) {
      sheet_pr.append(" codeName=\"");
      AppendXmlAttrEscaped(sheet_pr, code_name);
      sheet_pr.push_back('"');
    }
    if (body.empty()) {
      sheet_pr.append("/>");
    } else {
      sheet_pr.push_back('>');
      sheet_pr.append(body);
      sheet_pr.append("</sheetPr>");
    }
  }
  if (!sheet_pr.empty()) {
    sheet.mutable_print_settings().sheet_pr_xml = std::move(sheet_pr);
  }

  return flags_or.value() == kDefaultWsPropFlags && flags2_or.value() == kDefaultWsPropFlags2 &&
         rw_sync_or.value() == kWsPropSyncUnused && col_sync_or.value() == kWsPropSyncUnused;
}

/// Emits the worksheet-properties record which starts the mandatory worksheet
/// prefix. `BrtWsDim` must follow this record before the view collection.
void emit_ws_prop(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  const WorksheetProperties properties_model = ParseSheetProperties(sheet.print_settings().sheet_pr_xml);
  std::vector<std::uint8_t> properties;
  emit_u16(properties, kDefaultWsPropFlags);  // page breaks, publish, outline defaults
  emit_u8(properties, kDefaultWsPropFlags2);  // evaluate conditional formatting
  // BrtColor. An automatic type with the automatic palette index is the
  // untinted tab; any other selector sets fValidRGB alongside it, matching
  // how the styles writer spells the same structure.
  emit_u8(properties,
          properties_model.color_type == 0U
              ? std::uint8_t{0}
              : static_cast<std::uint8_t>((static_cast<unsigned>(properties_model.color_type) << 1U) | 0x01U));
  emit_u8(properties, properties_model.color_index);
  emit_u16(properties, static_cast<std::uint16_t>(properties_model.color_tint));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 16U) & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 8U) & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>(properties_model.color_argb & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 24U) & 0xFFU));
  emit_u32(properties, kWsPropSyncUnused);  // rwSync: unused
  emit_u32(properties, kWsPropSyncUnused);  // colSync: unused
  emit_xlwidestring(properties, properties_model.code_name);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp), properties);
}

/// Decodes the fixed-width worksheet default-format record into the sheet's
/// `<sheetFormatPr>` model.
Expected<void, Error> decode_ws_fmt_info(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index) {
  if (rec.payload.size < 12U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsFmtInfo payload truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }

  ByteSpan p = rec.payload;
  auto dx_g_col_or = read_u32(p);
  auto cch_def_col_width_or = read_u16(p);
  auto miy_def_rw_height_or = read_u16(p);
  auto flags_or = read_u32(p);
  if (!dx_g_col_or || !cch_def_col_width_or || !miy_def_rw_height_or || !flags_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsFmtInfo fields truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }
  if (p.size != 0U) {
    return make_error(
        FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo has trailing bytes",
        "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " trailing=" + std::to_string(p.size));
  }

  const std::uint32_t dx_g_col = dx_g_col_or.value();
  const std::uint16_t cch_def_col_width = cch_def_col_width_or.value();
  const std::uint16_t miy_def_rw_height = miy_def_rw_height_or.value();
  const std::uint32_t flags = flags_or.value();
  if (dx_g_col != kAbsentDefaultColumnWidth && dx_g_col > 65535U) {
    return make_error(
        FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo default column width out of range",
        "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " dxGCol=" + std::to_string(dx_g_col));
  }
  if (cch_def_col_width > 255U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo base column width out of range",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) +
                          " cchDefColWidth=" + std::to_string(cch_def_col_width));
  }

  const bool thick_top = (flags & 0x00000004U) != 0U;
  const bool thick_bottom = (flags & 0x00000008U) != 0U;
  const std::uint8_t row_outline_max = static_cast<std::uint8_t>((flags >> 16U) & 0xFFU);
  const std::uint8_t col_outline_max = static_cast<std::uint8_t>((flags >> 24U) & 0xFFU);
  if (row_outline_max > 7U || col_outline_max > 7U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo outline level out of range",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " row_outline_max=" +
                          std::to_string(row_outline_max) + " col_outline_max=" + std::to_string(col_outline_max));
  }

  SheetFormatDefaults& defaults = sheet.mutable_format_defaults();
  defaults.base_col_width = static_cast<double>(cch_def_col_width);
  if (dx_g_col == kAbsentDefaultColumnWidth) {
    defaults.has_default_col_width = false;
    defaults.default_col_width = 0.0;
  } else {
    defaults.has_default_col_width = true;
    defaults.default_col_width = static_cast<double>(dx_g_col) / 256.0;
  }

  const bool f_unsynced = (flags & 0x00000001U) != 0U;
  const bool f_dy_zero = (flags & 0x00000002U) != 0U;
  // Excel always writes miyDefRwHeight, so the record carries the default
  // row height whether or not fUnsynced marks it as a custom override.
  defaults.has_default_row_height = true;
  defaults.default_row_height = static_cast<double>(miy_def_rw_height) / 20.0;
  defaults.custom_height = f_unsynced;
  defaults.zero_height = f_dy_zero;
  defaults.thick_top = thick_top;
  defaults.thick_bottom = thick_bottom;
  defaults.outline_level_row = row_outline_max;
  defaults.outline_level_col = col_outline_max;
  return Expected<void, Error>::Ok();
}

/// Emits BrtWsFmtInfo from the sheet's modelled `<sheetFormatPr>` values.
void emit_ws_fmt_info(std::vector<std::uint8_t>& dst, const SheetFormatDefaults& defaults) {
  // BrtWsFmtInfo stores the default column width as 1/256 character units,
  // while the OOXML model stores it in character units. The sentinel is the
  // only absent marker; zero is a valid explicit width.
  constexpr std::uint16_t kCanonicalDefaultColumnWidth = 8U;
  constexpr double kMaxDefaultColumnWidth = 65535.0 / 256.0;
  constexpr double kMaxDefaultRowHeight = 65535.0 / 20.0;

  const bool valid_base_col_width = std::isfinite(defaults.base_col_width) && defaults.base_col_width >= 0.0 &&
                                    defaults.base_col_width <= 255.0 &&
                                    std::floor(defaults.base_col_width) == defaults.base_col_width;
  const bool valid_default_col_width =
      !defaults.has_default_col_width ||
      (std::isfinite(defaults.default_col_width) && defaults.default_col_width >= 0.0 &&
       defaults.default_col_width <= kMaxDefaultColumnWidth);
  const bool valid_default_row_height =
      !defaults.has_default_row_height ||
      (std::isfinite(defaults.default_row_height) && defaults.default_row_height >= 0.0 &&
       defaults.default_row_height <= kMaxDefaultRowHeight &&
       std::isfinite(std::round(defaults.default_row_height * 20.0)));

  const std::uint16_t base_col_width =
      valid_base_col_width ? static_cast<std::uint16_t>(defaults.base_col_width) : kCanonicalDefaultColumnWidth;
  const bool emit_default_col_width = defaults.has_default_col_width && valid_default_col_width;
  const std::uint32_t dx_g_col = emit_default_col_width
                                     ? static_cast<std::uint32_t>(std::floor(defaults.default_col_width * 256.0))
                                     : kAbsentDefaultColumnWidth;

  std::uint16_t miy_default_row_height = kCanonicalDefaultRowHeightTwips;
  std::uint32_t format_flags = 0U;
  const bool emit_default_row_height = defaults.has_default_row_height && valid_default_row_height;
  if (emit_default_row_height) {
    miy_default_row_height = static_cast<std::uint16_t>(std::round(defaults.default_row_height * 20.0));
    if (miy_default_row_height == 0U) {
      format_flags |= 0x00000002U;  // fDyZero: explicit zero (including quantized-zero).
    }
  }
  if (defaults.custom_height) {
    format_flags |= 0x00000001U;  // fUnsynced
  }
  if (defaults.zero_height) {
    format_flags |= 0x00000002U;  // fDyZero: rows hidden by default.
  }
  if (defaults.thick_top) {
    format_flags |= 0x00000004U;
  }
  if (defaults.thick_bottom) {
    format_flags |= 0x00000008U;
  }
  const std::uint8_t row_outline_max = std::min<std::uint8_t>(defaults.outline_level_row, 7U);
  const std::uint8_t col_outline_max = std::min<std::uint8_t>(defaults.outline_level_col, 7U);
  format_flags |= static_cast<std::uint32_t>(row_outline_max) << 16U;
  format_flags |= static_cast<std::uint32_t>(col_outline_max) << 24U;

  std::vector<std::uint8_t> formatting;
  emit_u32(formatting, dx_g_col);
  emit_u16(formatting, base_col_width);
  emit_u16(formatting, miy_default_row_height);
  emit_u32(formatting, format_flags);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo), formatting);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
