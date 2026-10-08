
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

inline constexpr std::array<ExcelProfileIdEntry, 14> kExcelProfileIds = {{
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
