#include "io/ooxml/preset_geometry.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/preset_geometry_table.h"
#include "utils/elementary_math.h"

namespace formulon {
namespace preset_geometry {
namespace {

// @size-budget: 24 KB
// ECMA-376 `ST_ShapeType`, each name NUL-terminated, packed in byte order for the binary search.
// clang-format off
constexpr char kPresetGeometries[] =
    "accentBorderCallout1\0" "accentBorderCallout2\0" "accentBorderCallout3\0" "accentCallout1\0" "accentCallout2\0"
    "accentCallout3\0" "actionButtonBackPrevious\0" "actionButtonBeginning\0" "actionButtonBlank\0"
    "actionButtonDocument\0" "actionButtonEnd\0" "actionButtonForwardNext\0" "actionButtonHelp\0" "actionButtonHome\0"
    "actionButtonInformation\0" "actionButtonMovie\0" "actionButtonReturn\0" "actionButtonSound\0" "arc\0"
    "bentArrow\0" "bentConnector2\0" "bentConnector3\0" "bentConnector4\0" "bentConnector5\0" "bentUpArrow\0"
    "bevel\0" "blockArc\0" "borderCallout1\0" "borderCallout2\0" "borderCallout3\0" "bracePair\0" "bracketPair\0"
    "callout1\0" "callout2\0" "callout3\0" "can\0" "chartPlus\0" "chartStar\0" "chartX\0" "chevron\0" "chord\0"
    "circularArrow\0" "cloud\0" "cloudCallout\0" "corner\0" "cornerTabs\0" "cube\0" "curvedConnector2\0"
    "curvedConnector3\0" "curvedConnector4\0" "curvedConnector5\0" "curvedDownArrow\0" "curvedLeftArrow\0"
    "curvedRightArrow\0" "curvedUpArrow\0" "decagon\0" "diagStripe\0" "diamond\0" "dodecagon\0" "donut\0"
    "doubleWave\0" "downArrow\0" "downArrowCallout\0" "ellipse\0" "ellipseRibbon\0" "ellipseRibbon2\0"
    "flowChartAlternateProcess\0" "flowChartCollate\0" "flowChartConnector\0" "flowChartDecision\0" "flowChartDelay\0"
    "flowChartDisplay\0" "flowChartDocument\0" "flowChartExtract\0" "flowChartInputOutput\0"
    "flowChartInternalStorage\0" "flowChartMagneticDisk\0" "flowChartMagneticDrum\0" "flowChartMagneticTape\0"
    "flowChartManualInput\0" "flowChartManualOperation\0" "flowChartMerge\0" "flowChartMultidocument\0"
    "flowChartOfflineStorage\0" "flowChartOffpageConnector\0" "flowChartOnlineStorage\0" "flowChartOr\0"
    "flowChartPredefinedProcess\0" "flowChartPreparation\0" "flowChartProcess\0" "flowChartPunchedCard\0"
    "flowChartPunchedTape\0" "flowChartSort\0" "flowChartSummingJunction\0" "flowChartTerminator\0" "foldedCorner\0"
    "frame\0" "funnel\0" "gear6\0" "gear9\0" "halfFrame\0" "heart\0" "heptagon\0" "hexagon\0" "homePlate\0"
    "horizontalScroll\0" "irregularSeal1\0" "irregularSeal2\0" "leftArrow\0" "leftArrowCallout\0" "leftBrace\0"
    "leftBracket\0" "leftCircularArrow\0" "leftRightArrow\0" "leftRightArrowCallout\0" "leftRightCircularArrow\0"
    "leftRightRibbon\0" "leftRightUpArrow\0" "leftUpArrow\0" "lightningBolt\0" "line\0" "lineInv\0" "mathDivide\0"
    "mathEqual\0" "mathMinus\0" "mathMultiply\0" "mathNotEqual\0" "mathPlus\0" "moon\0" "noSmoking\0"
    "nonIsoscelesTrapezoid\0" "notchedRightArrow\0" "octagon\0" "parallelogram\0" "pentagon\0" "pie\0" "pieWedge\0"
    "plaque\0" "plaqueTabs\0" "plus\0" "quadArrow\0" "quadArrowCallout\0" "rect\0" "ribbon\0" "ribbon2\0"
    "rightArrow\0" "rightArrowCallout\0" "rightBrace\0" "rightBracket\0" "round1Rect\0" "round2DiagRect\0"
    "round2SameRect\0" "roundRect\0" "rtTriangle\0" "smileyFace\0" "snip1Rect\0" "snip2DiagRect\0" "snip2SameRect\0"
    "snipRoundRect\0" "squareTabs\0" "star10\0" "star12\0" "star16\0" "star24\0" "star32\0" "star4\0" "star5\0"
    "star6\0" "star7\0" "star8\0" "straightConnector1\0" "stripedRightArrow\0" "sun\0" "swooshArrow\0" "teardrop\0"
    "trapezoid\0" "triangle\0" "upArrow\0" "upArrowCallout\0" "upDownArrow\0" "upDownArrowCallout\0" "uturnArrow\0"
    "verticalScroll\0" "wave\0" "wedgeEllipseCallout\0" "wedgeRectCallout\0" "wedgeRoundRectCallout\0";
// clang-format on

constexpr std::size_t preset_count() {
  std::size_t terminators = 0;
  for (const char c : kPresetGeometries) {
    terminators += c == '\0' ? 1U : 0U;
  }
  return terminators - 1U;  // The literal's own terminator ends no name.
}
constexpr std::size_t kPresetCount = preset_count();
static_assert(sizeof(kPresetGeometries) <= 0xFFFFU, "preset offsets are 16-bit");

// Start of each name, plus one past the last name's terminator.
constexpr std::array<std::uint16_t, kPresetCount + 1> preset_offsets() {
  std::array<std::uint16_t, kPresetCount + 1> out{};
  std::size_t n = 1;
  for (std::size_t i = 0; n <= kPresetCount; ++i) {
    if (kPresetGeometries[i] == '\0') {
      out[n++] = static_cast<std::uint16_t>(i + 1U);
    }
  }
  return out;
}
constexpr std::array<std::uint16_t, kPresetCount + 1> kPresetOffsets = preset_offsets();

std::string_view preset_at(std::size_t i) {
  const std::size_t length = static_cast<std::size_t>(kPresetOffsets[i + 1] - kPresetOffsets[i]) - 1U;
  return {kPresetGeometries + kPresetOffsets[i], length};
}

#include "io/ooxml/preset_geometry_table.inc"

static_assert(std::size(kPresetEntries) == kPresetCount, "the geometry table covers every ST_ShapeType name");

constexpr double kPi = 3.14159265358979323846;
// DrawingML angles are in 60000ths of a degree.
constexpr double kAngleUnit = 60000.0 * 180.0 / kPi;

// atan2(y, x) in radians from the portable arctangent.
double angle_of(double y, double x) {
  if (x > 0.0) {
    return elementary::atan(y / x);
  }
  if (x < 0.0) {
    return elementary::atan(y / x) + (y < 0.0 ? -kPi : kPi);
  }
  return y > 0.0 ? kPi / 2.0 : y < 0.0 ? -kPi / 2.0 : 0.0;
}

// The operator byte's operand count, in the generator's order.
constexpr std::array<std::uint8_t, 17> kArity = {3, 3, 3, 3, 1, 2, 3, 2, 2, 2, 3, 3, 3, 2, 1, 2, 1};

// One formula (ECMA-376 Part 1, 20.1.9.11) on its operands.
double apply(std::uint8_t op, const double* a) {
  switch (op) {
    case 0:
      return a[2] == 0.0 ? 0.0 : a[0] * a[1] / a[2];
    case 1:
      return a[0] + a[1] - a[2];
    case 2:
      return a[2] == 0.0 ? 0.0 : (a[0] + a[1]) / a[2];
    case 3:
      return a[0] > 0.0 ? a[1] : a[2];
    case 4:
      return std::fabs(a[0]);
    case 5:
      return angle_of(a[1], a[0]) * kAngleUnit;
    case 6:
      return a[0] * elementary::cos(angle_of(a[2], a[1]));
    case 7:
      return a[0] * elementary::cos(a[1] / kAngleUnit);
    case 8:
      return a[0] > a[1] ? a[0] : a[1];
    case 9:
      return a[0] < a[1] ? a[0] : a[1];
    case 10:
      return std::sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
    case 11:
      return a[1] < a[0] ? a[0] : a[1] > a[2] ? a[2] : a[1];
    case 12:
      return a[0] * elementary::sin(angle_of(a[2], a[1]));
    case 13:
      return a[0] * elementary::sin(a[1] / kAngleUnit);
    case 14:
      return a[0] > 0.0 ? std::sqrt(a[0]) : 0.0;
    case 15:
      return a[0] * elementary::tan(a[1] / kAngleUnit);
    default:
      return a[0];
  }
}

// Reads the operands of one preset's code: shared values, adjust guides and earlier guides.
class Evaluator {
 public:
  Evaluator(double w, double h, std::vector<double> adjusts)
      : w_(w), h_(h), ss_(w < h ? w : h), ls_(w < h ? h : w), adjusts_(std::move(adjusts)) {}

  double operand(const std::uint8_t*& at) const {
    const unsigned tag = *at >> 6U;
    std::size_t index = *at & 0x3FU;
    ++at;
    if (index == 0x3FU) {
      index += *at++;
    }
    if (tag == 1U) {
      return adjusts_[index];
    }
    if (tag == 2U) {
      return guides_[index];
    }
    const PresetValue& v = kPresetValues[index];
    const double base = v.kind == 1 ? w_ : v.kind == 2 ? h_ : v.kind == 3 ? ss_ : ls_;
    return v.kind == 0 ? static_cast<double>(v.arg) : base / v.arg;
  }

  void run(const std::uint8_t*& at, std::size_t count) {
    for (std::size_t i = 0; i < count; ++i) {
      const std::uint8_t op = *at++;
      double args[3] = {0.0, 0.0, 0.0};
      for (std::size_t k = 0; k < kArity[op]; ++k) {
        args[k] = operand(at);
      }
      guides_.push_back(apply(op, args));
    }
  }

 private:
  double w_, h_, ss_, ls_;
  std::vector<double> adjusts_;
  std::vector<double> guides_;
};

}  // namespace

std::optional<std::size_t> index_of(std::string_view prst) {
  std::size_t lo = 0;
  std::size_t hi = kPresetCount;
  while (lo < hi) {
    const std::size_t mid = lo + (hi - lo) / 2U;
    if (preset_at(mid) < prst) {
      lo = mid + 1U;
    } else {
      hi = mid;
    }
  }
  if (lo < kPresetCount && preset_at(lo) == prst) {
    return lo;
  }
  return std::nullopt;
}

std::size_t preset_count() {
  return kPresetCount;
}

std::string_view name_at(std::size_t index) {
  return preset_at(index);
}

std::vector<AdjustGuide> adjust_guides(std::size_t index) {
  const PresetEntry& entry = kPresetEntries[index];
  std::vector<AdjustGuide> out;
  out.reserve(entry.adjust_count);
  for (std::size_t i = 0; i < entry.adjust_count; ++i) {
    const PresetAdjust& a = kPresetAdjusts[entry.adjust_first + i];
    out.push_back(AdjustGuide{kPresetAdjustNames[a.name], a.value});
  }
  return out;
}

std::vector<Point> connection_sites(std::size_t index, double w, double h, const std::vector<AdjustGuide>& stated) {
  const PresetEntry& entry = kPresetEntries[index];
  std::vector<double> adjusts;
  for (const AdjustGuide& guide : adjust_guides(index)) {
    double value = guide.value;
    for (const AdjustGuide& given : stated) {
      if (given.name == guide.name) {
        value = given.value;
      }
    }
    adjusts.push_back(value);
  }
  Evaluator eval(w, h, std::move(adjusts));
  const std::uint8_t* at = kPresetCode + entry.code_offset;
  eval.run(at, entry.guide_count);
  std::vector<Point> sites;
  sites.reserve(entry.site_count);
  for (std::size_t i = 0; i < entry.site_count; ++i) {
    const double x = eval.operand(at);
    const double y = eval.operand(at);
    sites.push_back(Point{x, y});
  }
  return sites;
}

}  // namespace preset_geometry
}  // namespace formulon
