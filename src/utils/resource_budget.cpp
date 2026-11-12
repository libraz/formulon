#include "utils/resource_budget.h"

#include <string>

namespace formulon {

std::string limit_context(std::string_view part, std::uint64_t requested, std::uint64_t limit) {
  std::string out("part=");
  out.append(part);
  out.append(" requested=");
  out.append(std::to_string(requested));
  out.append(" limit=");
  out.append(std::to_string(limit));
  return out;
}

}  // namespace formulon
