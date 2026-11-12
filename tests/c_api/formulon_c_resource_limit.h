#pragma once

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

/// Asserts `status` is `FM_STATUS_SECURITY_RESOURCE_LIMIT` and that the last error context leads with
/// `part=<part> requested=<n> limit=<n>`, both numbers unsigned decimals with `requested > limit`.
/// `requested` may be absent only for `cf_formula_arena`, which reports the available bytes as `limit`.
/// Returns the context so a caller can check the site keys that follow.
inline std::string ExpectResourceLimit(fm_status_t status, const std::string& part) {
  EXPECT_EQ(status, FM_STATUS_SECURITY_RESOURCE_LIMIT) << fm_get_last_error_message();
  const std::string context = fm_get_last_error_context();
  std::vector<std::pair<std::string, std::string>> pairs;
  std::istringstream in(context);
  for (std::string token; in >> token;) {
    const std::size_t eq = token.find('=');
    if (eq == std::string::npos) {
      break;
    }
    pairs.emplace_back(token.substr(0, eq), token.substr(eq + 1));
  }
  const auto is_unsigned = [](const std::string& text) {
    return !text.empty() && text.find_first_not_of("0123456789") == std::string::npos;
  };
  const bool has_requested = part != "cf_formula_arena";
  const std::size_t lead = has_requested ? 3U : 2U;
  EXPECT_GE(pairs.size(), lead) << context;
  if (pairs.size() < lead) {
    return context;
  }
  EXPECT_EQ(pairs[0].first, "part") << context;
  EXPECT_EQ(pairs[0].second, part) << context;
  const std::size_t limit_at = has_requested ? 2U : 1U;
  EXPECT_EQ(pairs[limit_at].first, "limit") << context;
  EXPECT_TRUE(is_unsigned(pairs[limit_at].second)) << context;
  if (has_requested) {
    EXPECT_EQ(pairs[1].first, "requested") << context;
    EXPECT_TRUE(is_unsigned(pairs[1].second)) << context;
    if (is_unsigned(pairs[1].second) && is_unsigned(pairs[limit_at].second)) {
      EXPECT_GT(std::strtoull(pairs[1].second.c_str(), nullptr, 10),
                std::strtoull(pairs[limit_at].second.c_str(), nullptr, 10))
          << context;
    }
  }
  return context;
}
