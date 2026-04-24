// Copyright 2026 libraz. Licensed under the MIT License.
//
// CONVERT(number, from_unit, to_unit) -- Excel's general unit converter.
//
// Design:
//   * Every supported unit is a row in `kUnits` below with (name, category,
//     linear-scale factor relative to the category's canonical unit,
//     optional offset for affine temperature units, per-unit prefix
//     flags). Categories share a canonical unit (m for Distance, g for
//     Weight, J for Energy, etc.), so converting A -> B is a linear
//     `value * A.factor / B.factor` -- except for Temperature which uses
//     the affine form `(value * A.factor + A.offset - B.offset) / B.factor`.
//   * SI / binary prefixes are looked up only when a unit name does not
//     match exactly; this preserves cases like `h` -> horsepower (not
//     hecto-) and `hh` -> horsepower-hour. Prefix compatibility is per-
//     unit (the "allows_si" / "allows_binary" flags). For squared / cubed
//     units (dim = 2 or 3) the prefix multiplier is raised to that power
//     before combining with the base factor, so e.g. `Mm^2` resolves to
//     `(10^6)^2 * m^2 = 10^12 m^2`.
//
// Unit factors and prefix-compatibility flags are taken directly from the
// IronCalc CONVERT fixtures (which themselves mirror Mac Excel 365 output)
// rather than from MS Office documentation, so a handful of apparently
// surprising flags (e.g. `uk_pt` accepts SI prefixes, `mph` accepts SI
// prefixes, `at`/`atm`/`mmHg` accept SI prefixes) intentionally agree with
// the oracle.

#include "eval/builtins/engineering_convert.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Unit taxonomy
// ---------------------------------------------------------------------------

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

struct UnitEntry {
  const char* name;
  Category category;
  std::uint8_t dim;  // 1 (linear) / 2 (squared) / 3 (cubed) -- governs SI-prefix exponentiation.
  double factor;     // `value_in_unit * factor = value_in_canonical_unit`.
  double offset;     // Temperature-only affine offset; 0 for linear categories.
  bool allows_si;
  bool allows_binary;  // Only bit/byte.
};

// Unit table. Values were extracted from the IronCalc CONVERT fixtures.
constexpr UnitEntry kUnits[] = {
    // Weight (canonical: g)
    {"g", Category::kWeight, 1, 1.0, 0.0, true, false},
    {"sg", Category::kWeight, 1, 14593.902937206363, 0.0, false, false},
    {"lbm", Category::kWeight, 1, 453.59237, 0.0, false, false},
    {"u", Category::kWeight, 1, 1.660538782e-24, 0.0, true, false},
    {"ozm", Category::kWeight, 1, 28.349523125, 0.0, false, false},
    {"grain", Category::kWeight, 1, 0.06479891, 0.0, false, false},
    {"cwt", Category::kWeight, 1, 45359.237, 0.0, false, false},
    {"shweight", Category::kWeight, 1, 45359.237, 0.0, false, false},
    {"uk_cwt", Category::kWeight, 1, 50802.34544, 0.0, false, false},
    {"lcwt", Category::kWeight, 1, 50802.34544, 0.0, false, false},
    {"stone", Category::kWeight, 1, 6350.29318, 0.0, false, false},
    {"ton", Category::kWeight, 1, 907184.74, 0.0, false, false},
    {"brton", Category::kWeight, 1, 1016046.9088, 0.0, false, false},
    {"LTON", Category::kWeight, 1, 1016046.9088, 0.0, false, false},
    {"uk_ton", Category::kWeight, 1, 1016046.9088, 0.0, false, false},

    // Distance (canonical: m)
    {"m", Category::kDistance, 1, 1.0, 0.0, true, false},
    {"mi", Category::kDistance, 1, 1609.344, 0.0, false, false},
    {"Nmi", Category::kDistance, 1, 1852.0, 0.0, false, false},
    {"in", Category::kDistance, 1, 0.0254, 0.0, false, false},
    {"ft", Category::kDistance, 1, 0.3048, 0.0, false, false},
    {"yd", Category::kDistance, 1, 0.9144, 0.0, false, false},
    {"ang", Category::kDistance, 1, 1e-10, 0.0, true, false},
    {"ell", Category::kDistance, 1, 1.143, 0.0, false, false},
    {"ly", Category::kDistance, 1, 9460730472580800.0, 0.0, true, false},
    {"parsec", Category::kDistance, 1, 3.0856775812815532e+16, 0.0, true, false},
    {"pc", Category::kDistance, 1, 3.0856775812815532e+16, 0.0, true, false},
    {"Picapt", Category::kDistance, 1, 0.0003527777777777778, 0.0, false, false},
    {"Pica", Category::kDistance, 1, 0.0003527777777777778, 0.0, false, false},
    {"pica", Category::kDistance, 1, 0.004233333333333334, 0.0, false, false},
    {"survey_mi", Category::kDistance, 1, 1609.3472186944373, 0.0, false, false},

    // Time (canonical: s). Prefix flags follow MS CONVERT documentation --
    // the fixture does not exercise them.
    {"yr", Category::kTime, 1, 31557600.0, 0.0, false, false},
    {"day", Category::kTime, 1, 86400.0, 0.0, false, false},
    {"d", Category::kTime, 1, 86400.0, 0.0, false, false},
    {"hr", Category::kTime, 1, 3600.0, 0.0, false, false},
    {"mn", Category::kTime, 1, 60.0, 0.0, false, false},
    {"min", Category::kTime, 1, 60.0, 0.0, false, false},
    {"sec", Category::kTime, 1, 1.0, 0.0, true, false},
    {"s", Category::kTime, 1, 1.0, 0.0, true, false},

    // Pressure (canonical: Pa)
    {"Pa", Category::kPressure, 1, 1.0, 0.0, true, false},
    {"p", Category::kPressure, 1, 1.0, 0.0, true, false},
    {"atm", Category::kPressure, 1, 101325.0, 0.0, true, false},
    {"at", Category::kPressure, 1, 101325.0, 0.0, true, false},
    {"mmHg", Category::kPressure, 1, 133.322, 0.0, true, false},
    {"psi", Category::kPressure, 1, 6894.757293168362, 0.0, false, false},
    {"Torr", Category::kPressure, 1, 133.32236842105263, 0.0, false, false},

    // Force (canonical: N)
    {"N", Category::kForce, 1, 1.0, 0.0, true, false},
    {"dyn", Category::kForce, 1, 1e-05, 0.0, true, false},
    {"dy", Category::kForce, 1, 1e-05, 0.0, true, false},
    {"lbf", Category::kForce, 1, 4.4482216152605005, 0.0, false, false},
    {"pond", Category::kForce, 1, 0.00980665, 0.0, true, false},

    // Energy (canonical: J)
    {"J", Category::kEnergy, 1, 1.0, 0.0, true, false},
    {"e", Category::kEnergy, 1, 1e-07, 0.0, true, false},
    {"c", Category::kEnergy, 1, 4.184, 0.0, true, false},
    {"cal", Category::kEnergy, 1, 4.1868, 0.0, true, false},
    {"eV", Category::kEnergy, 1, 1.602176487e-19, 0.0, true, false},
    {"ev", Category::kEnergy, 1, 1.602176487e-19, 0.0, true, false},
    {"HPh", Category::kEnergy, 1, 2684519.537696173, 0.0, false, false},
    {"hh", Category::kEnergy, 1, 2684519.537696173, 0.0, false, false},
    {"Wh", Category::kEnergy, 1, 3600.0, 0.0, true, false},
    {"wh", Category::kEnergy, 1, 3600.0, 0.0, true, false},
    {"flb", Category::kEnergy, 1, 1.3558179483314003, 0.0, false, false},
    {"BTU", Category::kEnergy, 1, 1055.05585262, 0.0, false, false},
    {"btu", Category::kEnergy, 1, 1055.05585262, 0.0, false, false},

    // Power (canonical: W)
    {"HP", Category::kPower, 1, 745.6998715822702, 0.0, false, false},
    {"h", Category::kPower, 1, 745.6998715822702, 0.0, false, false},
    {"PS", Category::kPower, 1, 735.49875, 0.0, false, false},
    {"W", Category::kPower, 1, 1.0, 0.0, true, false},
    {"w", Category::kPower, 1, 1.0, 0.0, true, false},

    // Magnetism (canonical: T)
    {"T", Category::kMagnetism, 1, 1.0, 0.0, true, false},
    {"ga", Category::kMagnetism, 1, 1e-4, 0.0, true, false},

    // Temperature (canonical: K). `factor` is the per-degree scale to K,
    // `offset` is the zero-offset of the unit's zero in K.
    {"C", Category::kTemperature, 1, 1.0, 273.15, false, false},
    {"cel", Category::kTemperature, 1, 1.0, 273.15, false, false},
    {"F", Category::kTemperature, 1, 5.0 / 9.0, 273.15 - 32.0 * 5.0 / 9.0, false, false},
    {"fah", Category::kTemperature, 1, 5.0 / 9.0, 273.15 - 32.0 * 5.0 / 9.0, false, false},
    {"K", Category::kTemperature, 1, 1.0, 0.0, true, false},
    {"kel", Category::kTemperature, 1, 1.0, 0.0, true, false},
    {"Rank", Category::kTemperature, 1, 5.0 / 9.0, 0.0, false, false},
    {"Reau", Category::kTemperature, 1, 5.0 / 4.0, 273.15, false, false},

    // Volume (canonical: m^3)
    {"tsp", Category::kVolume, 1, 4.92892159375e-06, 0.0, false, false},
    {"tspm", Category::kVolume, 1, 5e-06, 0.0, false, false},
    {"tbs", Category::kVolume, 1, 1.4786764781249999e-05, 0.0, false, false},
    {"oz", Category::kVolume, 1, 2.9573529562499998e-05, 0.0, false, false},
    {"cup", Category::kVolume, 1, 0.00023658823649999998, 0.0, false, false},
    {"pt", Category::kVolume, 1, 0.00047317647299999996, 0.0, false, false},
    {"us_pt", Category::kVolume, 1, 0.00047317647299999996, 0.0, false, false},
    {"uk_pt", Category::kVolume, 1, 0.00056826125, 0.0, true, false},
    {"qt", Category::kVolume, 1, 0.0009463529459999999, 0.0, false, false},
    {"uk_qt", Category::kVolume, 1, 0.0011365225, 0.0, false, false},
    {"gal", Category::kVolume, 1, 0.0037854117839999997, 0.0, false, false},
    {"uk_gal", Category::kVolume, 1, 0.00454609, 0.0, false, false},
    {"l", Category::kVolume, 1, 0.001, 0.0, true, false},
    {"L", Category::kVolume, 1, 0.001, 0.0, true, false},
    {"lt", Category::kVolume, 1, 0.001, 0.0, true, false},
    {"ang3", Category::kVolume, 3, 1e-30, 0.0, true, false},
    {"ang^3", Category::kVolume, 3, 1e-30, 0.0, true, false},
    {"barrel", Category::kVolume, 1, 0.158987294928, 0.0, false, false},
    {"bushel", Category::kVolume, 1, 0.03523907016688, 0.0, false, false},
    {"ft3", Category::kVolume, 3, 0.028316846592, 0.0, false, false},
    {"ft^3", Category::kVolume, 3, 0.028316846592, 0.0, false, false},
    {"in3", Category::kVolume, 3, 1.6387064e-05, 0.0, false, false},
    {"in^3", Category::kVolume, 3, 1.6387064e-05, 0.0, false, false},
    {"ly3", Category::kVolume, 3, 8.467866646237152e+47, 0.0, false, false},
    {"ly^3", Category::kVolume, 3, 8.467866646237152e+47, 0.0, false, false},
    {"m3", Category::kVolume, 3, 1.0, 0.0, true, false},
    {"m^3", Category::kVolume, 3, 1.0, 0.0, true, false},
    {"mi3", Category::kVolume, 3, 4168181825.4405794, 0.0, false, false},
    {"mi^3", Category::kVolume, 3, 4168181825.4405794, 0.0, false, false},
    {"yd3", Category::kVolume, 3, 0.764554857984, 0.0, false, false},
    {"yd^3", Category::kVolume, 3, 0.764554857984, 0.0, false, false},
    {"Nmi3", Category::kVolume, 3, 6352182208.0, 0.0, false, false},
    {"Nmi^3", Category::kVolume, 3, 6352182208.0, 0.0, false, false},
    {"Picapt3", Category::kVolume, 3, 4.3903956618655696e-11, 0.0, false, false},
    {"Picapt^3", Category::kVolume, 3, 4.3903956618655696e-11, 0.0, false, false},
    {"Pica3", Category::kVolume, 3, 4.3903956618655696e-11, 0.0, false, false},
    {"Pica^3", Category::kVolume, 3, 4.3903956618655696e-11, 0.0, false, false},
    {"GRT", Category::kVolume, 1, 2.8316846592, 0.0, false, false},
    {"regton", Category::kVolume, 1, 2.8316846592, 0.0, false, false},
    {"MTON", Category::kVolume, 1, 1.13267386368, 0.0, false, false},

    // Area (canonical: m^2)
    {"uk_acre", Category::kArea, 1, 4046.8564224, 0.0, false, false},
    {"us_acre", Category::kArea, 1, 4046.8726098742522, 0.0, false, false},
    {"ang2", Category::kArea, 2, 1e-20, 0.0, true, false},
    {"ang^2", Category::kArea, 2, 1e-20, 0.0, true, false},
    {"ar", Category::kArea, 1, 100.0, 0.0, true, false},
    {"ft2", Category::kArea, 2, 0.09290304, 0.0, false, false},
    {"ft^2", Category::kArea, 2, 0.09290304, 0.0, false, false},
    {"ha", Category::kArea, 1, 10000.0, 0.0, false, false},
    {"in2", Category::kArea, 2, 0.00064516, 0.0, false, false},
    {"in^2", Category::kArea, 2, 0.00064516, 0.0, false, false},
    {"ly2", Category::kArea, 2, 8.95054210748189e+31, 0.0, false, false},
    {"ly^2", Category::kArea, 2, 8.95054210748189e+31, 0.0, false, false},
    {"m2", Category::kArea, 2, 1.0, 0.0, true, false},
    {"m^2", Category::kArea, 2, 1.0, 0.0, true, false},
    {"Morgen", Category::kArea, 1, 2500.0, 0.0, false, false},
    {"mi2", Category::kArea, 2, 2589988.110336, 0.0, false, false},
    {"mi^2", Category::kArea, 2, 2589988.110336, 0.0, false, false},
    {"Nmi2", Category::kArea, 2, 3429904.0, 0.0, false, false},
    {"Nmi^2", Category::kArea, 2, 3429904.0, 0.0, false, false},
    {"Picapt2", Category::kArea, 2, 1.2445216049382715e-07, 0.0, false, false},
    {"Pica2", Category::kArea, 2, 1.2445216049382715e-07, 0.0, false, false},
    {"Pica^2", Category::kArea, 2, 1.2445216049382715e-07, 0.0, false, false},
    {"Picapt^2", Category::kArea, 2, 1.2445216049382715e-07, 0.0, false, false},
    {"yd2", Category::kArea, 2, 0.83612736, 0.0, false, false},
    {"yd^2", Category::kArea, 2, 0.83612736, 0.0, false, false},

    // Information (canonical: bit). bit and byte additionally accept the
    // binary (Ki / Mi / ...) prefix family.
    {"bit", Category::kInformation, 1, 1.0, 0.0, true, true},
    {"byte", Category::kInformation, 1, 8.0, 0.0, true, true},

    // Speed (canonical: m/s)
    {"admkn", Category::kSpeed, 1, 0.5147733333333333, 0.0, false, false},
    {"kn", Category::kSpeed, 1, 0.5144444444444445, 0.0, false, false},
    {"m/h", Category::kSpeed, 1, 0.0002777777777777778, 0.0, true, false},
    {"m/hr", Category::kSpeed, 1, 0.0002777777777777778, 0.0, true, false},
    {"m/s", Category::kSpeed, 1, 1.0, 0.0, true, false},
    {"m/sec", Category::kSpeed, 1, 1.0, 0.0, true, false},
    {"mph", Category::kSpeed, 1, 0.44704, 0.0, true, false},
};

constexpr std::size_t kUnitCount = sizeof(kUnits) / sizeof(kUnits[0]);

// ---------------------------------------------------------------------------
// Prefix tables
// ---------------------------------------------------------------------------

struct Prefix {
  const char* symbol;
  double multiplier;
};

// SI prefixes. Ordered long-first within each starting letter so that a
// linear prefix-match ("does the candidate begin with this symbol?")
// always picks the longest applicable symbol (e.g. "da" before "d").
constexpr Prefix kSiPrefixes[] = {
    {"Y", 1e24},  {"Z", 1e21}, {"E", 1e18}, {"P", 1e15},  {"T", 1e12},  {"G", 1e9},  {"M", 1e6},
    {"k", 1e3},   {"h", 1e2},  {"da", 1e1}, {"e", 1e1},   {"d", 1e-1},  {"c", 1e-2}, {"m", 1e-3},
    {"u", 1e-6},  {"n", 1e-9}, {"p", 1e-12}, {"f", 1e-15}, {"a", 1e-18}, {"z", 1e-21}, {"y", 1e-24},
};
constexpr std::size_t kSiPrefixCount = sizeof(kSiPrefixes) / sizeof(kSiPrefixes[0]);

// Binary prefixes for bit / byte.
constexpr Prefix kBinaryPrefixes[] = {
    {"Yi", 1208925819614629174706176.0}, {"Zi", 1180591620717411303424.0},
    {"Ei", 1152921504606846976.0},        {"Pi", 1125899906842624.0},
    {"Ti", 1099511627776.0},              {"Gi", 1073741824.0},
    {"Mi", 1048576.0},                    {"ki", 1024.0},
};
constexpr std::size_t kBinaryPrefixCount = sizeof(kBinaryPrefixes) / sizeof(kBinaryPrefixes[0]);

// ---------------------------------------------------------------------------
// Lookup helpers
// ---------------------------------------------------------------------------

const UnitEntry* find_exact_unit(std::string_view name) {
  for (std::size_t i = 0; i < kUnitCount; ++i) {
    if (name == kUnits[i].name) {
      return &kUnits[i];
    }
  }
  return nullptr;
}

struct Resolved {
  Category category;
  std::uint8_t dim;
  double factor;
  double offset;
};

// Resolve a unit token to its canonical-scale factor / offset. Returns
// `false` if the token is not a recognised unit (with or without a valid
// prefix for the base unit in question).
bool resolve_unit(std::string_view name, Resolved* out) {
  if (const UnitEntry* exact = find_exact_unit(name); exact != nullptr) {
    out->category = exact->category;
    out->dim = exact->dim;
    out->factor = exact->factor;
    out->offset = exact->offset;
    return true;
  }
  // Try binary prefixes first -- their symbols ("Yi"/"Zi"/...) all end in
  // `i`, so there is no possibility of a collision with an SI prefix
  // symbol that happens to be a prefix of a SI prefix ("Y" vs. "Yi").
  for (std::size_t i = 0; i < kBinaryPrefixCount; ++i) {
    const std::size_t sym_len = std::strlen(kBinaryPrefixes[i].symbol);
    if (name.size() <= sym_len) continue;
    if (name.substr(0, sym_len) != kBinaryPrefixes[i].symbol) continue;
    const UnitEntry* base = find_exact_unit(name.substr(sym_len));
    if (base == nullptr || !base->allows_binary) continue;
    out->category = base->category;
    out->dim = base->dim;
    out->factor = base->factor * kBinaryPrefixes[i].multiplier;
    out->offset = base->offset;
    return true;
  }
  for (std::size_t i = 0; i < kSiPrefixCount; ++i) {
    const std::size_t sym_len = std::strlen(kSiPrefixes[i].symbol);
    if (name.size() <= sym_len) continue;
    if (name.substr(0, sym_len) != kSiPrefixes[i].symbol) continue;
    const UnitEntry* base = find_exact_unit(name.substr(sym_len));
    if (base == nullptr || !base->allows_si) continue;
    double multiplier = kSiPrefixes[i].multiplier;
    // For squared / cubed units the SI prefix acts on the linear length
    // first and is then raised to the unit's dimensionality. Power-of-3
    // and power-of-2 are the only cases in the catalog.
    if (base->dim == 2) multiplier *= multiplier;
    if (base->dim == 3) multiplier *= multiplier * kSiPrefixes[i].multiplier;
    out->category = base->category;
    out->dim = base->dim;
    out->factor = base->factor * multiplier;
    out->offset = base->offset;
    return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// CONVERT impl
// ---------------------------------------------------------------------------

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

  Resolved from_unit;
  Resolved to_unit;
  if (!resolve_unit(from_text.value(), &from_unit)) {
    return Value::error(ErrorCode::NA);
  }
  if (!resolve_unit(to_text.value(), &to_unit)) {
    return Value::error(ErrorCode::NA);
  }
  if (from_unit.category != to_unit.category) {
    return Value::error(ErrorCode::NA);
  }

  const double v = number.value();
  if (from_unit.category == Category::kTemperature) {
    const double kelvin = v * from_unit.factor + from_unit.offset;
    const double result = (kelvin - to_unit.offset) / to_unit.factor;
    return Value::number(result);
  }
  return Value::number(v * from_unit.factor / to_unit.factor);
}

}  // namespace

void register_engineering_convert_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"CONVERT", 3u, 3u, &Convert});
}

}  // namespace eval
}  // namespace formulon
