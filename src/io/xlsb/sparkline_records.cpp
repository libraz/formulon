// @size-budget: 12 KB

#include "io/xlsb/sparkline_records.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstring>
#include <optional>
#include <utility>

#include "io/sparkline_groups.h"
#include "io/sqref_text.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kAcBegin = 37;
constexpr std::uint16_t kAcEnd = 38;
constexpr std::uint16_t kBeginSparklineGroup = 1041;
constexpr std::uint16_t kEndSparklineGroup = 1042;
constexpr std::uint16_t kSparkline = 1043;
constexpr std::uint16_t kBeginSparklines = 1056;
constexpr std::uint16_t kEndSparklines = 1057;
constexpr std::uint16_t kBeginSparklineGroups = 1058;
constexpr std::uint16_t kEndSparklineGroups = 1059;
constexpr std::uint16_t kBeginCfs14 = 1135;
constexpr std::uint16_t kBeginDVals14 = 1054;
constexpr std::uint16_t kUid = 3072;
constexpr std::uint16_t kMargins = 476;
constexpr std::uint16_t kPrintOptions = 477;
constexpr std::uint16_t kPageSetup = 478;
// The BrtACBegin Excel writes around a sparkline group's BrtUid: one
// ACProductVersion, version 0x1002, fileExtension set.
constexpr std::array<std::uint8_t, 6> kGroupUidVersion = {0x01, 0x00, 0x02, 0x10, 0x00, 0x80};

// BrtBeginSparklineGroup flag bits ([MS-XLSB] 2.4.228); bits 1-2 are
// fShowEmptyCellAsZero (0 zero, 1 gap, 2 interpolated).
constexpr std::uint16_t kDateAxis = 0x0001;
constexpr std::uint16_t kMarkers = 0x0008;
constexpr std::uint16_t kHigh = 0x0010;
constexpr std::uint16_t kLow = 0x0020;
constexpr std::uint16_t kFirst = 0x0040;
constexpr std::uint16_t kLast = 0x0080;
constexpr std::uint16_t kNegative = 0x0100;
constexpr std::uint16_t kAxis = 0x0200;
constexpr std::uint16_t kHidden = 0x0400;
constexpr std::uint16_t kIndividualAutoMax = 0x0800;
constexpr std::uint16_t kIndividualAutoMin = 0x1000;
constexpr std::uint16_t kGroupAutoMax = 0x2000;
constexpr std::uint16_t kGroupAutoMin = 0x4000;
constexpr std::uint16_t kRightToLeft = 0x8000;
constexpr std::size_t kGroupFixedBytes = 2U + 8U * 8U + 3U * 8U + 4U;
constexpr double kMaxLineWeight = 1584.0;

std::uint32_t LoadU32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

double LoadF64(const std::uint8_t* p) {
  double v = 0;
  std::memcpy(&v, p, sizeof(v));
  return v;
}

std::optional<ColorSpec> DecodeSpec(const std::uint8_t* p) {
  const std::optional<cf::Color> c = DecodeColor(ByteSpan{p, 8U});
  if (!c) {
    return std::nullopt;
  }
  if (c->spec.kind != ColorSpec::Kind::kNone) {
    return c->spec;
  }
  ColorSpec rgb;
  rgb.kind = ColorSpec::Kind::kRgb;
  rgb.rgb = (static_cast<std::uint32_t>(c->a) << 24U) | (static_cast<std::uint32_t>(c->r) << 16U) |
            (static_cast<std::uint32_t>(c->g) << 8U) | static_cast<std::uint32_t>(c->b);
  return rgb;
}

std::array<ColorSpec*, 8> Colors(SparklineGroup& g) {
  return {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
          &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
}

std::array<const ColorSpec*, 8> Colors(const SparklineGroup& g) {
  return {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
          &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
}

/// BrtBeginSparklineGroup into `g`: FRTHeader (no formula, or the date-axis
/// FRTFormula), flags, eight BrtColor, dManualMax, dManualMin, dLineWeight,
/// isltype.
bool DecodeGroupHead(ByteSpan p, const FeatureFormulaReadContext& ctx, SparklineGroup& g) {
  auto flags = read_u32(p);
  if (!flags || (flags.value() & ~kFrtHeaderFormulas) != 0U) {
    return false;
  }
  if (flags.value() == kFrtHeaderFormulas) {
    auto count = read_u32(p);
    if (!count || count.value() != 1U) {
      return false;
    }
    auto ref = read_frt_formula(p, 0, 0, ctx);
    if (!ref || ref.value().empty()) {
      return false;
    }
    g.date_axis_ref = std::move(ref).value();
  }
  if (p.size != kGroupFixedBytes) {
    return false;
  }
  const std::uint8_t* d = p.data;
  const auto bits = static_cast<std::uint16_t>(d[0] | (d[1] << 8U));
  const std::uint32_t empty = (bits >> 1U) & 0x3U;
  if (empty > 2U || ((bits & kIndividualAutoMax) != 0U && (bits & kGroupAutoMax) != 0U) ||
      ((bits & kIndividualAutoMin) != 0U && (bits & kGroupAutoMin) != 0U)) {
    return false;
  }
  constexpr std::array<SparklineEmptyCells, 3> kEmpty = {SparklineEmptyCells::kZero, SparklineEmptyCells::kGap,
                                                         SparklineEmptyCells::kSpan};
  g.empty_cells_as = kEmpty[empty];
  g.date_axis = (bits & kDateAxis) != 0U;
  g.markers = (bits & kMarkers) != 0U;
  g.high = (bits & kHigh) != 0U;
  g.low = (bits & kLow) != 0U;
  g.first = (bits & kFirst) != 0U;
  g.last = (bits & kLast) != 0U;
  g.negative = (bits & kNegative) != 0U;
  g.display_x_axis = (bits & kAxis) != 0U;
  g.display_hidden = (bits & kHidden) != 0U;
  g.right_to_left = (bits & kRightToLeft) != 0U;
  g.max_axis_type = (bits & kIndividualAutoMax) != 0U ? SparklineAxisType::kIndividual
                    : (bits & kGroupAutoMax) != 0U    ? SparklineAxisType::kGroup
                                                      : SparklineAxisType::kCustom;
  g.min_axis_type = (bits & kIndividualAutoMin) != 0U ? SparklineAxisType::kIndividual
                    : (bits & kGroupAutoMin) != 0U    ? SparklineAxisType::kGroup
                                                      : SparklineAxisType::kCustom;
  std::array<ColorSpec*, 8> colors = Colors(g);
  for (std::size_t i = 0; i < colors.size(); ++i) {
    std::optional<ColorSpec> spec = DecodeSpec(d + 2U + 8U * i);
    if (!spec) {
      return false;
    }
    *colors[i] = *spec;
  }
  const std::uint8_t* nums = d + 2U + 64U;
  const double manual_max = LoadF64(nums);
  const double manual_min = LoadF64(nums + 8);
  g.line_weight = LoadF64(nums + 16);
  const std::uint32_t type = LoadU32(nums + 24);
  if (!std::isfinite(manual_max) || !std::isfinite(manual_min) || !std::isfinite(g.line_weight) ||
      g.line_weight < 0.0 || g.line_weight > kMaxLineWeight || type > 2U) {
    return false;
  }
  g.type = static_cast<SparklineType>(type);
  g.has_manual_max = g.max_axis_type == SparklineAxisType::kCustom;
  g.has_manual_min = g.min_axis_type == SparklineAxisType::kCustom;
  g.manual_max = g.has_manual_max ? manual_max : 0.0;
  g.manual_min = g.has_manual_min ? manual_min : 0.0;
  return true;
}

/// BrtSparkline: FRTHeader with one single-cell FRTSqref and one FRTFormula.
bool DecodeSparkline(ByteSpan p, const FeatureFormulaReadContext& ctx, SparklineItem& item) {
  auto flags = read_u32(p);
  auto count = read_u32(p);
  auto sqref_flags = read_u32(p);
  std::vector<MergeRange> cells;
  if (!flags || flags.value() != (kFrtHeaderSqrefs | kFrtHeaderFormulas) || !count || count.value() != 1U ||
      !sqref_flags || !read_sqref(p, cells) || cells.size() != 1U || cells[0].first_row != cells[0].last_row ||
      cells[0].first_col != cells[0].last_col) {
    return false;
  }
  auto formulas = read_u32(p);
  if (!formulas || formulas.value() != 1U) {
    return false;
  }
  auto text = read_frt_formula(p, 0, 0, ctx);
  if (!text || text.value().empty() || p.size != 0U) {
    return false;
  }
  item.formula = std::move(text).value();
  item.location = format_sqref_text(cells);
  return true;
}

/// One group starting at `recs[at]` (its uid block or BrtBeginSparklineGroup);
/// sets `last` to its BrtEndSparklineGroup.
std::optional<SparklineGroup> DecodeGroup(const std::vector<FramedRecord>& recs, std::size_t at, std::size_t until,
                                          const FeatureFormulaReadContext& ctx, std::size_t& last) {
  SparklineGroup g;
  const auto uid_block = [&](std::size_t i) {
    return i + 2U < until && recs[i].type == kAcBegin && recs[i + 1U].type == kUid &&
           recs[i + 1U].payload.size == kGuidBytes && recs[i + 2U].type == kAcEnd;
  };
  std::size_t i = at;
  if (uid_block(i)) {
    g.uid = FormatGuid(recs[i + 1U].payload.data);
    i += 3U;
  }
  if (i >= until || recs[i].type != kBeginSparklineGroup || !DecodeGroupHead(recs[i].payload, ctx, g)) {
    return std::nullopt;
  }
  ++i;
  if (g.uid.empty() && uid_block(i)) {
    g.uid = FormatGuid(recs[i + 1U].payload.data);
    i += 3U;
  }
  if (i >= until || recs[i].type != kBeginSparklines || recs[i].payload.size != 0U) {
    return std::nullopt;
  }
  for (++i; i < until && recs[i].type == kSparkline; ++i) {
    SparklineItem item;
    if (!DecodeSparkline(recs[i].payload, ctx, item)) {
      return std::nullopt;
    }
    g.items.push_back(std::move(item));
  }
  if (g.items.empty() || i + 1U >= until || recs[i].type != kEndSparklines || recs[i + 1U].type != kEndSparklineGroup) {
    return std::nullopt;
  }
  last = i + 1U;
  return g;
}

/// Index of the BrtEndSparklineGroups closing the container opened at `i`
/// (a BrtFRTBegin followed by BrtBeginSparklineGroups), or `recs.size()`.
std::size_t ContainerEnd(const std::vector<FramedRecord>& recs, std::size_t i) {
  if (recs[i].type != kFrtBegin || i + 1U >= recs.size() || recs[i + 1U].type != kBeginSparklineGroups) {
    return recs.size();
  }
  for (std::size_t k = i + 2U; k + 1U < recs.size(); ++k) {
    if (recs[k].type == kEndSparklineGroups) {
      return recs[k + 1U].type == kFrtEnd ? k : recs.size();
    }
  }
  return recs.size();
}

void TakeSlot(std::vector<std::uint8_t>& buf, const FeatureFormulaReadContext& ctx, std::string& ext_lst_xml) {
  const std::vector<FramedRecord> recs = split_records(buf);
  std::vector<std::uint8_t> out;
  bool changed = false;
  for (std::size_t i = 0; i < recs.size();) {
    const std::size_t close = ContainerEnd(recs, i);
    if (close == recs.size()) {
      append_record(out, buf, recs[i]);
      ++i;
      continue;
    }
    std::vector<std::uint8_t> kept;
    for (std::size_t k = i + 2U; k < close;) {
      std::size_t last = k;
      std::optional<SparklineGroup> group = DecodeGroup(recs, k, close, ctx, last);
      if (group && add_sparkline_group(ext_lst_xml, *group)) {
        changed = true;
        k = last + 1U;
        continue;
      }
      // Keep an undecodable group whole: up to its BrtEndSparklineGroup.
      std::size_t end = k;
      while (end + 1U < close && recs[end].type != kEndSparklineGroup) {
        ++end;
      }
      for (std::size_t m = k; m <= end; ++m) {
        append_record(kept, buf, recs[m]);
      }
      k = end + 1U;
    }
    if (!kept.empty()) {
      append_record(out, buf, recs[i]);
      append_record(out, buf, recs[i + 1U]);
      out.insert(out.end(), kept.begin(), kept.end());
      append_record(out, buf, recs[close]);
      append_record(out, buf, recs[close + 1U]);
    }
    i = close + 2U;
  }
  if (changed) {
    buf = std::move(out);
  }
}

Error Refuse(const std::string& what) {
  return make_error(FormulonErrorCode::kInvalidArgument, what, "context=write_xlsb");
}

std::uint32_t f_cce(const std::vector<std::uint8_t>& bytes) {
  return bytes.size() >= 4U ? LoadU32(bytes.data()) : 0U;
}

/// Quotes `name` for a sheet-qualified reference.
std::string QualifiedPrefix(std::string_view name) {
  std::string out = "'";
  for (const char c : name) {
    out.push_back(c);
    if (c == '\'') {
      out.push_back('\'');
    }
  }
  out.append("'!");
  return out;
}

/// The FRTFormula of a sparkline reference: one PtgRef3d / PtgArea3d /
/// PtgName / PtgNameX. A reference without a sheet is read on `sheet_name`.
Expected<EncodedFeatureFormula, Error> EncodeReference(const std::string& text, const FeatureFormulaWriteContext& ctx,
                                                       std::string_view sheet_name) {
  auto f = encode_feature_formula(text, 0, 0, PtgRootClass::kReference, ctx);
  if (!f) {
    return f;
  }
  const std::vector<std::uint8_t>& b = f.value().bytes;
  const std::uint8_t tag = b.size() > 4U ? static_cast<std::uint8_t>(b[4] & 0x1FU) : 0U;
  // Token sizes of PtgRef / PtgArea / PtgName / PtgNameX / PtgRef3d / PtgArea3d.
  const auto single = [&b](std::uint32_t size) { return f_cce(b) == size; };
  if ((tag == 0x04U && single(7U)) || (tag == 0x05U && single(13U))) {
    // A reference without a sheet: qualify it with the host sheet.
    return encode_feature_formula(QualifiedPrefix(sheet_name) + text, 0, 0, PtgRootClass::kReference, ctx);
  }
  if (!(tag == 0x03U && single(5U)) && !(tag == 0x19U && single(7U)) && !(tag == 0x1AU && single(9U)) &&
      !(tag == 0x1BU && single(15U))) {
    return Refuse("xlsb sparkline formula is not a single reference or name: " + text);
  }
  return f;
}

Expected<void, Error> EmitGroup(std::vector<std::uint8_t>& dst, const SparklineGroup& g,
                                const FeatureFormulaWriteContext& ctx, std::string_view sheet_name) {
  std::array<std::uint8_t, kGuidBytes> guid{};
  if (!g.uid.empty() && ParseGuid(g.uid, guid)) {
    emit_record(dst, kAcBegin, ByteSpan{kGroupUidVersion.data(), kGroupUidVersion.size()});
    emit_record(dst, kUid, ByteSpan{guid.data(), guid.size()});
    emit_record(dst, kAcEnd, ByteSpan{});
  }
  std::vector<std::uint8_t> p;
  if (!g.date_axis_ref.empty()) {
    auto ref = EncodeReference(g.date_axis_ref, ctx, sheet_name);
    if (!ref) {
      return std::move(ref.error());
    }
    emit_u32(p, kFrtHeaderFormulas);
    emit_u32(p, 1U);
    emit_frt_formula(p, ref.value());
  } else {
    emit_u32(p, 0U);
  }
  constexpr std::array<std::uint16_t, 3> kEmpty = {1U, 0U, 2U};  // gap, zero, span
  std::uint16_t bits = static_cast<std::uint16_t>(kEmpty[static_cast<std::size_t>(g.empty_cells_as)] << 1U);
  const struct {
    bool on;
    std::uint16_t bit;
  } flags[] = {{g.date_axis, kDateAxis},
               {g.markers, kMarkers},
               {g.high, kHigh},
               {g.low, kLow},
               {g.first, kFirst},
               {g.last, kLast},
               {g.negative, kNegative},
               {g.display_x_axis, kAxis},
               {g.display_hidden, kHidden},
               {g.max_axis_type == SparklineAxisType::kIndividual, kIndividualAutoMax},
               {g.min_axis_type == SparklineAxisType::kIndividual, kIndividualAutoMin},
               {g.max_axis_type == SparklineAxisType::kGroup, kGroupAutoMax},
               {g.min_axis_type == SparklineAxisType::kGroup, kGroupAutoMin},
               {g.right_to_left, kRightToLeft}};
  for (const auto& f : flags) {
    bits = static_cast<std::uint16_t>(bits | (f.on ? f.bit : 0U));
  }
  emit_u16(p, bits);
  for (const ColorSpec* spec : Colors(g)) {
    // An absent colour has no record form; it is written as opaque black.
    const std::uint32_t argb = spec->kind == ColorSpec::Kind::kNone ? 0xFF000000U : spec->rgb;
    const auto bytes = encode_brt_color(argb, *spec);
    p.insert(p.end(), bytes.begin(), bytes.end());
  }
  const bool custom_max = g.max_axis_type == SparklineAxisType::kCustom && g.has_manual_max;
  const bool custom_min = g.min_axis_type == SparklineAxisType::kCustom && g.has_manual_min;
  emit_double(p, custom_max ? g.manual_max : 0.0);
  emit_double(p, custom_min ? g.manual_min : 0.0);
  emit_double(p, g.line_weight);
  emit_u32(p, static_cast<std::uint32_t>(g.type));
  emit_record(dst, kBeginSparklineGroup, p);
  emit_record(dst, kBeginSparklines, ByteSpan{});
  for (const SparklineItem& item : g.items) {
    std::vector<MergeRange> cells;
    if (!parse_sqref_text(item.location, cells) || cells.size() != 1U || cells[0].first_row != cells[0].last_row ||
        cells[0].first_col != cells[0].last_col) {
      return Refuse("xlsb sparkline location is not one cell: " + item.location);
    }
    auto formula = EncodeReference(item.formula, ctx, sheet_name);
    if (!formula) {
      return std::move(formula.error());
    }
    std::vector<std::uint8_t> s;
    emit_u32(s, kFrtHeaderSqrefs | kFrtHeaderFormulas);
    emit_u32(s, 1U);
    emit_u32(s, kFrtItemFlags);
    emit_sqref(s, cells);
    emit_u32(s, 1U);
    emit_frt_formula(s, formula.value());
    emit_record(dst, kSparkline, s);
  }
  emit_record(dst, kEndSparklines, ByteSpan{});
  emit_record(dst, kEndSparklineGroup, ByteSpan{});
  return Expected<void, Error>::Ok();
}

/// Where a new SPARKLINEGROUPS container goes: after the x14 CF and DVals14
/// containers, ahead of any later worksheet future record and the sheet's
/// trailing revision-uid block.
std::size_t InsertionPoint(const std::vector<FramedRecord>& recs, std::size_t size) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kFrtBegin) {
      continue;
    }
    const bool earlier =
        i + 1U < recs.size() && (recs[i + 1U].type == kBeginCfs14 || recs[i + 1U].type == kBeginDVals14);
    if (!earlier) {
      return recs[i].begin;
    }
    while (i + 1U < recs.size() && recs[i].type != kFrtEnd) {
      ++i;
    }
  }
  if (recs.size() >= 3U && recs[recs.size() - 3U].type == kAcBegin && recs.back().type == kAcEnd) {
    return recs[recs.size() - 3U].begin;
  }
  return size;
}

}  // namespace

void take_sparkline_groups(XlsbSheetTail& tail, const FeatureFormulaReadContext& ctx, std::string& ext_lst_xml) {
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    TakeSlot(*slot, ctx, ext_lst_xml);
  }
}

bool has_retained_sparkline_groups(const XlsbSheetTail& tail) {
  for (const std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    for (const FramedRecord& rec : split_records(*slot)) {
      if (rec.type == kBeginSparklineGroups) {
        return true;
      }
    }
  }
  return false;
}

Expected<void, Error> add_sparkline_groups(std::vector<std::uint8_t>& records, std::string_view ext_lst_xml,
                                           const FeatureFormulaWriteContext& ctx, std::string_view sheet_name,
                                           std::uint32_t* skipped) {
  const std::vector<SparklineGroup> groups = read_sparkline_groups(ext_lst_xml);
  std::vector<std::uint8_t> body;
  for (const SparklineGroup& g : groups) {
    std::vector<std::uint8_t> encoded;
    if (EmitGroup(encoded, g, ctx, sheet_name)) {
      body.insert(body.end(), encoded.begin(), encoded.end());
    } else if (skipped != nullptr) {
      ++*skipped;
    }
  }
  if (body.empty()) {
    return Expected<void, Error>::Ok();
  }
  const std::vector<FramedRecord> recs = split_records(records);
  for (std::size_t i = 0; i < recs.size(); ++i) {
    const std::size_t close = ContainerEnd(recs, i);
    if (close != recs.size()) {
      records.insert(records.begin() + static_cast<std::ptrdiff_t>(recs[close].begin), body.begin(), body.end());
      return Expected<void, Error>::Ok();
    }
  }
  std::vector<std::uint8_t> block;
  // As for the x14 CF container: a sheet without print records gets Excel's
  // default BrtPrintOptions ahead of its first future record.
  const bool has_print = std::any_of(recs.begin(), recs.end(), [](const FramedRecord& r) {
    return r.type == kPrintOptions || r.type == kMargins || r.type == kPageSetup;
  });
  if (!has_print) {
    const std::array<std::uint8_t, 2> default_print_options = {0x10, 0x00};
    emit_record(block, kPrintOptions, ByteSpan{default_print_options.data(), default_print_options.size()});
  }
  emit_record(block, kFrtBegin, ByteSpan{kFrtVersion.data(), kFrtVersion.size()});
  emit_record(block, kBeginSparklineGroups, ByteSpan{});
  block.insert(block.end(), body.begin(), body.end());
  emit_record(block, kEndSparklineGroups, ByteSpan{});
  emit_record(block, kFrtEnd, ByteSpan{});
  const std::size_t at = InsertionPoint(recs, records.size());
  records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
