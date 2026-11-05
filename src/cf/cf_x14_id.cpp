//
// Implementation of the x14 rule id builder. See cf_x14_id.h.

#include "cf/cf_x14_id.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_set>

#include "utils/number_text.h"

namespace formulon::cf {
namespace {

void append_le_bytes(std::string& out, std::uint32_t value, int byte_count) {
  for (int i = 0; i < byte_count; ++i) {
    char buf[4];
    format_hex(buf, sizeof(buf), (value >> (8 * i)) & 0xFFU, 2, true);
    out.append(buf, 2);
  }
}

}  // namespace

std::string x14_rule_id(std::size_t sheet_index, std::uint32_t ordinal,
                        const std::unordered_set<std::string>& used_ids) {
  for (;; ++ordinal) {
    std::string id = "{00000000-000E-0000-";
    append_le_bytes(id, static_cast<std::uint32_t>(sheet_index & 0xFFFFU), 2);
    id += "-0000";
    append_le_bytes(id, ordinal, 4);
    id += '}';
    if (used_ids.find(id) == used_ids.end()) {
      return id;
    }
  }
}

}  // namespace formulon::cf
