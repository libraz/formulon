
#ifndef FORMULON_EVAL_COMPAT_H_
#define FORMULON_EVAL_COMPAT_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {

/// Excel host profile used for observed host-specific formula semantics.
enum class ExcelHost : std::uint8_t {
  kMac365 = 0,
  kWin365 = 1,
};

enum class ExcelLocale : std::uint8_t { kJaJP = 0 };

struct ExcelProfile {
  ExcelHost host = ExcelHost::kWin365;
  ExcelLocale locale = ExcelLocale::kJaJP;
};

inline constexpr ExcelProfile profile_from_host(ExcelHost host) noexcept {
  return ExcelProfile{host, ExcelLocale::kJaJP};
}

inline constexpr ExcelProfile mac_365_ja_jp_profile() noexcept {
  return ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP};
}

inline constexpr ExcelProfile win_365_ja_jp_profile() noexcept {
  return ExcelProfile{ExcelHost::kWin365, ExcelLocale::kJaJP};
}

inline constexpr ExcelProfile default_excel_profile() noexcept {
  return win_365_ja_jp_profile();
}

inline bool same_profile(ExcelProfile a, ExcelProfile b) noexcept {
  return a.host == b.host && a.locale == b.locale;
}

inline const char* excel_profile_id(ExcelProfile profile) noexcept {
  if (profile.host == ExcelHost::kMac365 && profile.locale == ExcelLocale::kJaJP) {
    return "mac-365-ja_JP";
  }
  if (profile.host == ExcelHost::kWin365 && profile.locale == ExcelLocale::kJaJP) {
    return "win-365-ja_JP";
  }
  return "win-365-ja_JP";
}

inline bool parse_excel_profile_id(std::string_view id, ExcelProfile* out) noexcept {
  if (id == "mac-365-ja_JP") {
    *out = ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP};
    return true;
  }
  if (id == "win-365-ja_JP") {
    *out = ExcelProfile{ExcelHost::kWin365, ExcelLocale::kJaJP};
    return true;
  }
  return false;
}

inline bool uses_mac_jp_text_folding(ExcelProfile profile) noexcept {
  return profile.host == ExcelHost::kMac365 && profile.locale == ExcelLocale::kJaJP;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COMPAT_H_
