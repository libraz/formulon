
#ifndef FORMULON_EXCEL_PROFILE_H_
#define FORMULON_EXCEL_PROFILE_H_

#include <array>
#include <cstdint>
#include <string_view>

namespace formulon {

/// Excel host profile used for observed host-specific formula semantics.
enum class ExcelHost : std::uint8_t {
  kMac365 = 0,
  kWin365 = 1,
};

enum class ExcelLocale : std::uint8_t {
  kJaJP = 0,
  kEnUS = 1,
  kDeDE = 2,
  kFrFR = 3,
  kZhCN = 4,
  kKoKR = 5,
  kThTH = 6,
  kEsES = 7,
  kEsMX = 8,
  kPtBR = 9,
  kRuRU = 10,
  kZhTW = 11,
  kItIT = 12,
  kNlNL = 13,
};

struct ExcelProfile {
  ExcelHost host = ExcelHost::kWin365;
  ExcelLocale locale = ExcelLocale::kJaJP;
};

inline constexpr ExcelProfile mac_365_ja_jp_profile() noexcept {
  return ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP};
}

inline constexpr ExcelProfile win_365_ja_jp_profile() noexcept {
  return ExcelProfile{ExcelHost::kWin365, ExcelLocale::kJaJP};
}

inline constexpr ExcelProfile mac_365_en_us_profile() noexcept {
  return ExcelProfile{ExcelHost::kMac365, ExcelLocale::kEnUS};
}

inline constexpr ExcelProfile win_365_en_us_profile() noexcept {
  return ExcelProfile{ExcelHost::kWin365, ExcelLocale::kEnUS};
}

/// True when the profile was measured against Excel: every Mac profile and
/// `win-365-ja_JP`. The other Windows profiles are estimated from the Mac
/// measurements plus the Windows host rules.
inline constexpr bool is_measured_profile(ExcelProfile p) noexcept {
  return p.host == ExcelHost::kMac365 || p.locale == ExcelLocale::kJaJP;
}

inline constexpr ExcelProfile default_excel_profile() noexcept {
  return win_365_en_us_profile();
}

inline bool same_profile(ExcelProfile a, ExcelProfile b) noexcept {
  return a.host == b.host && a.locale == b.locale;
}

namespace detail {

struct ExcelProfileIdEntry {
  ExcelProfile profile;
  const char* id;
};

inline constexpr std::array<ExcelProfileIdEntry, 28> kExcelProfileIds = {{
    {mac_365_ja_jp_profile(), "mac-365-ja_JP"},
    {win_365_ja_jp_profile(), "win-365-ja_JP"},
    {mac_365_en_us_profile(), "mac-365-en_US"},
    {win_365_en_us_profile(), "win-365-en_US"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kDeDE}, "mac-365-de_DE"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kDeDE}, "win-365-de_DE"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kFrFR}, "mac-365-fr_FR"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kFrFR}, "win-365-fr_FR"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kZhCN}, "mac-365-zh_CN"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kZhCN}, "win-365-zh_CN"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kKoKR}, "mac-365-ko_KR"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kKoKR}, "win-365-ko_KR"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kThTH}, "mac-365-th_TH"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kThTH}, "win-365-th_TH"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kEsES}, "mac-365-es_ES"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kEsES}, "win-365-es_ES"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kEsMX}, "mac-365-es_MX"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kEsMX}, "win-365-es_MX"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kPtBR}, "mac-365-pt_BR"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kPtBR}, "win-365-pt_BR"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kRuRU}, "mac-365-ru_RU"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kRuRU}, "win-365-ru_RU"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kZhTW}, "mac-365-zh_TW"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kZhTW}, "win-365-zh_TW"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kItIT}, "mac-365-it_IT"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kItIT}, "win-365-it_IT"},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kNlNL}, "mac-365-nl_NL"},
    {ExcelProfile{ExcelHost::kWin365, ExcelLocale::kNlNL}, "win-365-nl_NL"},
}};

}  // namespace detail

inline const char* excel_profile_id(ExcelProfile profile) noexcept {
  for (const detail::ExcelProfileIdEntry& entry : detail::kExcelProfileIds) {
    if (same_profile(profile, entry.profile)) {
      return entry.id;
    }
  }
  return "win-365-en_US";
}

inline bool parse_excel_profile_id(std::string_view id, ExcelProfile* out) noexcept {
  if (out == nullptr) {
    return false;
  }
  for (const detail::ExcelProfileIdEntry& entry : detail::kExcelProfileIds) {
    if (id == entry.id) {
      *out = entry.profile;
      return true;
    }
  }
  return false;
}

}  // namespace formulon

#endif  // FORMULON_EXCEL_PROFILE_H_
