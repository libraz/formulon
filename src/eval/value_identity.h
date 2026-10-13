#ifndef FORMULON_EVAL_VALUE_IDENTITY_H_
#define FORMULON_EVAL_VALUE_IDENTITY_H_

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/collation.h"
#include "excel_profile.h"
#include "value.h"

namespace formulon::eval {

// Appends a typed, unambiguous identity for a scalar Value. The encoding is
// shared by dynamic-array UNIQUE and callers that need keyed cell counts, so
// both see the same kind-specific identity.
// A false result means that the value has no canonical scalar identity.
inline bool append_value_identity_key(const Value& value, ExcelProfile profile, CollationStrength strength,
                                      std::string* out) {
  const auto append_bytes = [out](const void* data, std::size_t size) {
    out->append(static_cast<const char*>(data), size);
  };
  out->push_back(static_cast<char>(value.kind()));
  switch (value.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number: {
      const double number = value.as_number() == 0.0 ? 0.0 : value.as_number();
      append_bytes(&number, sizeof(number));
      return true;
    }
    case ValueKind::Bool:
      out->push_back(value.as_boolean() ? '\1' : '\0');
      return true;
    case ValueKind::Error: {
      const auto code = static_cast<std::uint32_t>(value.as_error());
      append_bytes(&code, sizeof(code));
      return true;
    }
    case ValueKind::Text: {
      const std::string key = collation_key(value.as_text(), profile, strength);
      const auto size = static_cast<std::uint64_t>(key.size());
      append_bytes(&size, sizeof(size));
      out->append(key);
      return true;
    }
    default:
      // Array / Ref / Lambda values do not have a canonical cell identity.
      return false;
  }
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_VALUE_IDENTITY_H_
