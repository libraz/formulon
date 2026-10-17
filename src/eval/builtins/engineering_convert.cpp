//
// CONVERT(number, from_unit, to_unit) -- Excel's general unit converter.
//
// Design:
//   * Every supported unit is a row of `kUnits` with its category, its
//     dimension (1 linear, 2 squared, 3 cubed), a factor relative to the
//     category's base unit, whether SI / binary prefixes are accepted, and
//     for temperatures the temperature kind. The base units are g, angstrom,
//     s, Pa, dyne, erg, W, gauss, litre, angstrom^2, bit and m/s; the factors
//     are the literal doubles of the table, never recomputed from SI units.
//   * A name resolves in this order, case-sensitively and without trimming:
//     an exact unit name; a one-character SI prefix plus a unit that accepts
//     SI prefixes (`e` is deka); a two-character prefix `Yi`..`ki` plus a
//     unit that accepts binary prefixes, or `da` (deka) plus an SI unit. A
//     prefix is never folded into the unit factor: SI prefixes are kept as a
//     power of ten (times the unit's dimension) and binary prefixes as a
//     multiplier, and combined as one ratio.
//   * result = ((v * from.factor) / to.factor) * ratio. If the first part
//     overflows it is recomputed as (v / to.factor) * from.factor, and an
//     overflow of the final result is #NUM!. Products and quotients flush
//     subnormal results to 0.
//   * Temperatures use an explicit formula for each pair of kinds; only
//     kelvin takes a prefix, and some pairs ignore it.
//   * When the categories differ, one side is Distance and the other Energy,
//     and a side's text is exactly "pc", it reads as pico-calorie.

#include "eval/builtins/engineering_convert.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/double_parse.h"
#include "utils/expected.h"
#include "utils/number_text.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

enum class Category : std::uint8_t {
  kWeight,
  kDistance,
  kTime,
  kPressure,
  kForce,
  kEnergy,
  kPower,
  kMagnetism,
  kTemperature,
  kVolume,
  kArea,
  kInformation,
  kSpeed,
};

// Temperature kinds in `UnitEntry::temperature`; 0 for every other category.
constexpr std::uint8_t kCelsius = 1;
constexpr std::uint8_t kFahrenheit = 2;
constexpr std::uint8_t kKelvin = 3;
constexpr std::uint8_t kRankine = 4;
constexpr std::uint8_t kReaumur = 5;

struct UnitEntry {
  const char* name;
  Category category;
  std::uint8_t dim;  // 1 linear / 2 squared / 3 cubed; scales the SI prefix exponent.
  double factor;     // Relative to the category's base unit.
  bool allows_si;
  bool allows_binary;  // Only bit / byte.
  std::uint8_t temperature;
};

// Unit factors, grouped by category.
constexpr UnitEntry kUnits[] = {
    // Weight (base: g)
    {"g", Category::kWeight, 1, 1.0, true, false, 0},
    {"sg", Category::kWeight, 1, 14593.902937206363, false, false, 0},
    {"lbm", Category::kWeight, 1, 453.59237, false, false, 0},
    {"u", Category::kWeight, 1, 1.660538782e-24, true, false, 0},
    {"ozm", Category::kWeight, 1, 28.349523125, false, false, 0},
    {"grain", Category::kWeight, 1, 0.06479891, false, false, 0},
    {"cwt", Category::kWeight, 1, 45359.237, false, false, 0},
    {"shweight", Category::kWeight, 1, 45359.237, false, false, 0},
    {"uk_cwt", Category::kWeight, 1, 50802.34544, false, false, 0},
    {"lcwt", Category::kWeight, 1, 50802.34544, false, false, 0},
    {"hweight", Category::kWeight, 1, 50802.34544, false, false, 0},
    {"stone", Category::kWeight, 1, 6350.29318, false, false, 0},
    {"ton", Category::kWeight, 1, 907184.74, false, false, 0},
    {"uk_ton", Category::kWeight, 1, 1016046.9088, false, false, 0},
    {"LTON", Category::kWeight, 1, 1016046.9088, false, false, 0},
    {"brton", Category::kWeight, 1, 1016046.9088, false, false, 0},
    // Distance (base: angstrom)
    {"m", Category::kDistance, 1, 10000000000.0, true, false, 0},
    {"mi", Category::kDistance, 1, 16093440000000.0, false, false, 0},
    {"Nmi", Category::kDistance, 1, 18520000000000.0, false, false, 0},
    {"in", Category::kDistance, 1, 254000000.0, false, false, 0},
    {"ft", Category::kDistance, 1, 3048000000.0, false, false, 0},
    {"yd", Category::kDistance, 1, 9144000000.0, false, false, 0},
    {"ang", Category::kDistance, 1, 1.0, true, false, 0},
    {"Pica", Category::kDistance, 1, 3527777.777777778, false, false, 0},
    {"pica", Category::kDistance, 1, 42333333.333333336, false, false, 0},
    {"ell", Category::kDistance, 1, 11430000000.0, false, false, 0},
    {"ly", Category::kDistance, 1, 9.4607304725808e+25, true, false, 0},
    {"parsec", Category::kDistance, 1, 3.0856775812815533e+26, true, false, 0},
    {"pc", Category::kDistance, 1, 3.0856775812815533e+26, true, false, 0},
    {"survey_mi", Category::kDistance, 1, 16093472186944.373, false, false, 0},
    {"Picapt", Category::kDistance, 1, 3527777.777777778, false, false, 0},
    // Time (base: s)
    {"yr", Category::kTime, 1, 31557600.0, false, false, 0},
    {"day", Category::kTime, 1, 86400.0, false, false, 0},
    {"d", Category::kTime, 1, 86400.0, false, false, 0},
    {"hr", Category::kTime, 1, 3600.0, false, false, 0},
    {"mn", Category::kTime, 1, 60.0, false, false, 0},
    {"min", Category::kTime, 1, 60.0, false, false, 0},
    {"sec", Category::kTime, 1, 1.0, true, false, 0},
    {"s", Category::kTime, 1, 1.0, true, false, 0},
    // Pressure (base: Pa)
    {"Pa", Category::kPressure, 1, 1.0, true, false, 0},
    {"p", Category::kPressure, 1, 1.0, true, false, 0},
    {"atm", Category::kPressure, 1, 101325.0, true, false, 0},
    {"at", Category::kPressure, 1, 101325.0, true, false, 0},
    {"mmHg", Category::kPressure, 1, 133.322, true, false, 0},
    {"psi", Category::kPressure, 1, 6894.757293168362, false, false, 0},
    {"Torr", Category::kPressure, 1, 133.32236842105263, false, false, 0},
    {"SI_at", Category::kPressure, 1, 98066.51, true, false, 0},
    // Force (base: dyne)
    {"N", Category::kForce, 1, 100000.0, true, false, 0},
    {"dyn", Category::kForce, 1, 1.0, true, false, 0},
    {"dy", Category::kForce, 1, 1.0, true, false, 0},
    {"lbf", Category::kForce, 1, 444822.16152605, false, false, 0},
    {"pond", Category::kForce, 1, 980.665, true, false, 0},
    // Energy (base: erg)
    {"J", Category::kEnergy, 1, 10000000.0, true, false, 0},
    {"e", Category::kEnergy, 1, 1.0, true, false, 0},
    {"c", Category::kEnergy, 1, 41840000.0, true, false, 0},
    {"cal", Category::kEnergy, 1, 41868000.0, true, false, 0},
    {"eV", Category::kEnergy, 1, 1.602176487e-12, true, false, 0},
    {"ev", Category::kEnergy, 1, 1.602176487e-12, true, false, 0},
    {"HPh", Category::kEnergy, 1, 26845195376961.73, false, false, 0},
    {"hh", Category::kEnergy, 1, 26845195376961.73, false, false, 0},
    {"Wh", Category::kEnergy, 1, 36000000000.0, true, false, 0},
    {"wh", Category::kEnergy, 1, 36000000000.0, true, false, 0},
    {"flb", Category::kEnergy, 1, 13558179.483314004, false, false, 0},
    {"BTU", Category::kEnergy, 1, 10550558526.2, false, false, 0},
    {"btu", Category::kEnergy, 1, 10550558526.2, false, false, 0},
    // Power (base: W)
    {"HP", Category::kPower, 1, 745.6998715822702, false, false, 0},
    {"h", Category::kPower, 1, 745.6998715822702, false, false, 0},
    {"W", Category::kPower, 1, 1.0, true, false, 0},
    {"w", Category::kPower, 1, 1.0, true, false, 0},
    {"PS", Category::kPower, 1, 735.49875, false, false, 0},
    // Magnetism (base: gauss)
    {"T", Category::kMagnetism, 1, 10000.0, true, false, 0},
    {"ga", Category::kMagnetism, 1, 1.0, true, false, 0},
    // Temperature (formulas, not factors)
    {"C", Category::kTemperature, 1, 1.0, false, false, 1},
    {"cel", Category::kTemperature, 1, 1.0, false, false, 1},
    {"F", Category::kTemperature, 1, 1.8, false, false, 2},
    {"fah", Category::kTemperature, 1, 1.8, false, false, 2},
    {"K", Category::kTemperature, 1, 1.0, true, false, 3},
    {"kel", Category::kTemperature, 1, 1.0, true, false, 3},
    {"Rank", Category::kTemperature, 1, 1.0, false, false, 4},
    {"Reau", Category::kTemperature, 1, 1.0, false, false, 5},
    // Volume (base: litre)
    {"tsp", Category::kVolume, 1, 0.00492892159375, false, false, 0},
    {"tbs", Category::kVolume, 1, 0.01478676478125, false, false, 0},
    {"oz", Category::kVolume, 1, 0.0295735295625, false, false, 0},
    {"cup", Category::kVolume, 1, 0.2365882365, false, false, 0},
    {"pt", Category::kVolume, 1, 0.473176473, false, false, 0},
    {"qt", Category::kVolume, 1, 0.946352946, false, false, 0},
    {"gal", Category::kVolume, 1, 3.785411784, false, false, 0},
    {"l", Category::kVolume, 1, 1.0, true, false, 0},
    {"lt", Category::kVolume, 1, 1.0, true, false, 0},
    {"L", Category::kVolume, 1, 1.0, true, false, 0},
    {"uk_pt", Category::kVolume, 1, 0.56826125, true, false, 0},
    {"us_pt", Category::kVolume, 1, 0.473176473, false, false, 0},
    {"ang3", Category::kVolume, 3, 1e-27, true, false, 0},
    {"ang^3", Category::kVolume, 3, 1e-27, true, false, 0},
    {"barrel", Category::kVolume, 1, 158.987294928, false, false, 0},
    {"bushel", Category::kVolume, 1, 35.23907016688, false, false, 0},
    {"ft3", Category::kVolume, 3, 28.316846592, false, false, 0},
    {"ft^3", Category::kVolume, 3, 28.316846592, false, false, 0},
    {"GRT", Category::kVolume, 1, 2831.6846592, false, false, 0},
    {"regton", Category::kVolume, 1, 2831.6846592, false, false, 0},
    {"in3", Category::kVolume, 3, 0.016387064, false, false, 0},
    {"in^3", Category::kVolume, 3, 0.016387064, false, false, 0},
    {"ly3", Category::kVolume, 3, 8.467866646237152e+50, false, false, 0},
    {"ly^3", Category::kVolume, 3, 8.467866646237152e+50, false, false, 0},
    {"m3", Category::kVolume, 3, 1000.0, true, false, 0},
    {"m^3", Category::kVolume, 3, 1000.0, true, false, 0},
    {"mi3", Category::kVolume, 3, 4168181825440.5796, false, false, 0},
    {"mi^3", Category::kVolume, 3, 4168181825440.5796, false, false, 0},
    {"MTON", Category::kVolume, 1, 1132.67386368, false, false, 0},
    {"Nmi3", Category::kVolume, 3, 6352182208000.0, false, false, 0},
    {"Nmi^3", Category::kVolume, 3, 6352182208000.0, false, false, 0},
    {"Pica3", Category::kVolume, 3, 4.39039566186557e-08, false, false, 0},
    {"Pica^3", Category::kVolume, 3, 4.39039566186557e-08, false, false, 0},
    {"tspm", Category::kVolume, 1, 0.005, false, false, 0},
    {"uk_gal", Category::kVolume, 1, 4.54609, false, false, 0},
    {"yd3", Category::kVolume, 3, 764.554857984, false, false, 0},
    {"yd^3", Category::kVolume, 3, 764.554857984, false, false, 0},
    {"uk_qt", Category::kVolume, 1, 1.1365225, false, false, 0},
    {"Picapt3", Category::kVolume, 3, 4.39039566186557e-08, false, false, 0},
    {"Picapt^3", Category::kVolume, 3, 4.39039566186557e-08, false, false, 0},
    // Area (base: angstrom^2)
    {"uk_acre", Category::kArea, 1, 4.0468564224e+23, false, false, 0},
    {"us_acre", Category::kArea, 1, 4.046872609874252e+23, false, false, 0},
    {"ang2", Category::kArea, 2, 1.0, true, false, 0},
    {"ang^2", Category::kArea, 2, 1.0, true, false, 0},
    {"ar", Category::kArea, 1, 1e+22, true, false, 0},
    {"ft2", Category::kArea, 2, 9.290304e+18, false, false, 0},
    {"ft^2", Category::kArea, 2, 9.290304e+18, false, false, 0},
    {"ha", Category::kArea, 1, 1e+24, false, false, 0},
    {"in2", Category::kArea, 2, 6.4516e+16, false, false, 0},
    {"in^2", Category::kArea, 2, 6.4516e+16, false, false, 0},
    {"ly2", Category::kArea, 2, 8.95054210748189e+51, false, false, 0},
    {"ly^2", Category::kArea, 2, 8.95054210748189e+51, false, false, 0},
    {"m2", Category::kArea, 2, 1e+20, true, false, 0},
    {"m^2", Category::kArea, 2, 1e+20, true, false, 0},
    {"Morgen", Category::kArea, 1, 2.5e+23, false, false, 0},
    {"mi2", Category::kArea, 2, 2.589988110336e+26, false, false, 0},
    {"mi^2", Category::kArea, 2, 2.589988110336e+26, false, false, 0},
    {"Nmi2", Category::kArea, 2, 3.429904e+26, false, false, 0},
    {"Nmi^2", Category::kArea, 2, 3.429904e+26, false, false, 0},
    {"Pica2", Category::kArea, 2, 12445216049382.715, false, false, 0},
    {"Pica^2", Category::kArea, 2, 12445216049382.715, false, false, 0},
    {"yd2", Category::kArea, 2, 8.3612736e+19, false, false, 0},
    {"yd^2", Category::kArea, 2, 8.3612736e+19, false, false, 0},
    {"Picapt2", Category::kArea, 2, 12445216049382.715, false, false, 0},
    {"Picapt^2", Category::kArea, 2, 12445216049382.715, false, false, 0},
    // Information (base: bit)
    {"bit", Category::kInformation, 1, 1.0, true, true, 0},
    {"byte", Category::kInformation, 1, 8.0, true, true, 0},
    // Speed (base: m/s)
    {"admkn", Category::kSpeed, 1, 0.5147733333333333, false, false, 0},
    {"kn", Category::kSpeed, 1, 0.5144444444444445, false, false, 0},
    {"m/h", Category::kSpeed, 1, 0.0002777777777777778, true, false, 0},
    {"m/hr", Category::kSpeed, 1, 0.0002777777777777778, true, false, 0},
    {"m/s", Category::kSpeed, 1, 1.0, true, false, 0},
    {"m/sec", Category::kSpeed, 1, 1.0, true, false, 0},
    {"mph", Category::kSpeed, 1, 0.44704, true, false, 0},
};

constexpr std::size_t kUnitCount = sizeof(kUnits) / sizeof(kUnits[0]);

// SI prefix characters and their decimal exponents (`e` is deka).
constexpr char kSiChars[] = "YZEPTGMkhedcmunpfazy";
constexpr int kSiExponents[] = {24, 21, 18, 15, 12, 9, 6, 3, 2, 1, -1, -2, -3, -6, -9, -12, -15, -18, -21, -24};

// Two-character prefixes: the binary ones (with their multipliers) and `da`.
constexpr const char* kPairPrefixes[] = {"Yi", "Zi", "Ei", "Pi", "Ti", "Gi", "Mi", "ki"};
constexpr double kBinaryMultipliers[] = {1208925819614629174706176.0,
                                         1180591620717411303424.0,
                                         1152921504606846976.0,
                                         1125899906842624.0,
                                         1099511627776.0,
                                         1073741824.0,
                                         1048576.0,
                                         1024.0};
constexpr std::size_t kPairPrefixCount = sizeof(kPairPrefixes) / sizeof(kPairPrefixes[0]);

const UnitEntry* find_exact_unit(std::string_view name) noexcept {
  for (std::size_t i = 0; i < kUnitCount; ++i) {
    if (name == kUnits[i].name) {
      return &kUnits[i];
    }
  }
  return nullptr;
}

// A resolved unit token: the unit plus its prefix, either a power of ten
// (`is_power`, `exponent`) or a binary multiplier.
struct Resolved {
  const UnitEntry* unit = nullptr;
  bool is_power = true;
  int exponent = 0;
  double multiplier = 1.0;
};

// Resolves `text`. `forced` skips the exact-name attempt (the pico-calorie reading of "pc").
bool resolve_unit(std::string_view text, bool forced, Resolved* out) noexcept {
  *out = Resolved{};
  if (text.empty()) {
    return false;
  }
  if (text.size() == 1) {
    out->unit = forced ? nullptr : find_exact_unit(text);
    return out->unit != nullptr;
  }
  if (!forced) {
    if (const UnitEntry* exact = find_exact_unit(text); exact != nullptr) {
      out->unit = exact;
      return true;
    }
  }
  for (std::size_t i = 0; i < sizeof(kSiExponents) / sizeof(kSiExponents[0]); ++i) {
    if (text[0] != kSiChars[i]) {
      continue;
    }
    if (const UnitEntry* base = find_exact_unit(text.substr(1)); base != nullptr && base->allows_si) {
      *out = Resolved{base, true, kSiExponents[i], 1.0};
      return true;
    }
    break;
  }
  if (text.size() >= 3) {
    const std::string_view rest = text.substr(2);
    for (std::size_t i = 0; i < kPairPrefixCount; ++i) {
      if (text.substr(0, 2) != kPairPrefixes[i]) {
        continue;
      }
      if (const UnitEntry* base = find_exact_unit(rest); base != nullptr && base->allows_binary) {
        *out = Resolved{base, false, 0, kBinaryMultipliers[i]};
        return true;
      }
    }
    if (text.substr(0, 2) == "da") {
      if (const UnitEntry* base = find_exact_unit(rest); base != nullptr && base->allows_si) {
        *out = Resolved{base, true, 1, 1.0};
        return true;
      }
    }
  }
  return false;
}

// Scalar arithmetic with the converter's error rules: an overflow or a divisor
// below the normal range sets `error`, and a subnormal result flushes to 0.
struct Arith {
  bool error = false;

  double store(double r) noexcept {
    if (std::isnan(r) || std::isinf(r)) {
      error = true;
      return 0.0;
    }
    return std::fabs(r) < std::numeric_limits<double>::min() ? 0.0 : r;
  }
  double mul(double a, double b) noexcept { return store(a * b); }
  double div(double a, double b) noexcept {
    if (std::fabs(b) < std::numeric_limits<double>::min()) {
      error = true;
      return 0.0;
    }
    return store(a / b);
  }
};

// Temperature conversion: an explicit formula per (from, to) kind pair, with the
// kelvin prefix as `from_power` / `to_power` (10^exponent). Returns false on #NUM!.
bool convert_temperature(double v, const Resolved& from, const Resolved& to, double* out) noexcept {
  Arith m;
  const std::uint8_t kf = from.unit->temperature;
  const std::uint8_t kt = to.unit->temperature;
  const double from_power = builtins_detail::power_of_ten_exact(from.exponent);
  const double to_power = builtins_detail::power_of_ten_exact(to.exponent);
  double r = 0.0;
  switch (kf) {
    case kKelvin: {
      const double a = m.mul(m.mul(from_power, from.unit->factor), v);
      switch (kt) {
        case kKelvin:
          r = m.div(a, m.mul(to_power, to.unit->factor));
          break;
        case kCelsius:
          r = a - 273.15;
          break;
        case kFahrenheit:
          r = (m.mul(a - 273.15, 1.8)) + 32.0;
          break;
        case kRankine:
          r = m.mul(v, 1.8);
          break;
        default:
          r = m.mul(v - 273.15, 0.8);
          break;
      }
      break;
    }
    case kCelsius:
      switch (kt) {
        case kCelsius:
          r = v;
          break;
        case kFahrenheit:
          r = m.mul(v, 1.8) + 32.0;
          break;
        case kKelvin:
          r = m.div(273.15 + v, to_power);
          break;
        case kRankine:
          r = m.mul(v + 273.15, 1.8);
          break;
        default:
          r = m.mul(v, 0.8);
          break;
      }
      break;
    case kFahrenheit:
      switch (kt) {
        case kFahrenheit:
          r = v;
          break;
        case kCelsius:
          r = m.div(v - 32.0, 1.8);
          break;
        case kKelvin:
          r = m.div(m.div(v - 32.0, 1.8) + 273.15, to_power);
          break;
        case kRankine:
          r = v + 459.67;
          break;
        default:
          r = m.div(m.mul(v - 32.0, 0.8), 1.8);
          break;
      }
      break;
    case kRankine:
      switch (kt) {
        case kRankine:
          r = v;
          break;
        case kCelsius:
          r = m.div(v, 1.8) - 273.15;
          break;
        case kFahrenheit:
          r = v - 459.67;
          break;
        case kKelvin:
          r = m.div(m.div(v, 1.8), to_power);
          break;
        default:
          r = m.mul(m.div(v, 1.8) - 273.15, 0.8);
          break;
      }
      break;
    default:  // Reaumur
      switch (kt) {
        case kReaumur:
          r = v;
          break;
        case kCelsius:
          r = m.div(v, 0.8);
          break;
        case kFahrenheit:
          r = m.div(m.mul(v, 1.8), 0.8) + 32.0;
          break;
        case kKelvin:
          r = m.div(v, 0.8) + 273.15;
          break;
        default:
          r = m.mul(m.div(v, 0.8) + 273.15, 1.8);
          break;
      }
      break;
  }
  *out = r;
  return !m.error;
}

Value Convert(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  auto from_text = coerce_to_text(args[1]);
  if (!from_text) {
    return Value::error(from_text.error());
  }
  auto to_text = coerce_to_text(args[2]);
  if (!to_text) {
    return Value::error(to_text.error());
  }

  Resolved from;
  Resolved to;
  if (!resolve_unit(from_text.value(), false, &from) || !resolve_unit(to_text.value(), false, &to)) {
    return Value::error(ErrorCode::NA);
  }
  if (from.unit->category != to.unit->category) {
    // "pc" between Distance and Energy is pico-calorie rather than parsec.
    const Category a = from.unit->category;
    const Category b = to.unit->category;
    bool resolved = false;
    if (a == Category::kDistance && b == Category::kEnergy && from_text.value() == "pc") {
      resolved = resolve_unit(from_text.value(), true, &from);
    } else if (b == Category::kDistance && a == Category::kEnergy && to_text.value() == "pc") {
      resolved = resolve_unit(to_text.value(), true, &to);
    }
    if (!resolved) {
      return Value::error(ErrorCode::NA);
    }
  }
  from.exponent *= from.unit->dim;
  to.exponent *= to.unit->dim;

  const double v = number.value();
  if (from.unit->temperature != 0) {
    double r = 0.0;
    if (!convert_temperature(v, from, to, &r)) {
      return Value::error(ErrorCode::Num);
    }
    return Value::number(r);
  }

  Arith m;
  double ratio = 0.0;
  if (from.is_power && to.is_power) {
    ratio = builtins_detail::power_of_ten_exact(from.exponent - to.exponent);
  } else if (from.is_power) {
    ratio = m.div(builtins_detail::power_of_ten_exact(from.exponent), to.multiplier);
  } else if (to.is_power) {
    ratio = m.div(from.multiplier, builtins_detail::power_of_ten_exact(to.exponent));
  } else {
    ratio = m.div(from.multiplier, to.multiplier);
  }
  double r = m.div(m.mul(v, from.unit->factor), to.unit->factor);
  if (m.error) {
    m.error = false;
    r = m.mul(m.div(v, to.unit->factor), from.unit->factor);
  }
  r = m.mul(r, ratio);
  if (m.error) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_engineering_convert_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      builtins_detail::analysis_toolpak({"CONVERT", 3u, 3u, &Convert}),
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
