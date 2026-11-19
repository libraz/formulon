#ifndef FORMULON_EVAL_VALUE_IDENTITY_H_
#define FORMULON_EVAL_VALUE_IDENTITY_H_

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/collation.h"
#include "eval/picture_compare.h"
#include "excel_profile.h"
#include "picture_table.h"
#include "value.h"

namespace formulon::eval {

// Byte offset of a Text identity's collation key: the kind tag and the 8-byte
// key length precede it.
inline constexpr std::size_t kValueIdentityTextKeyOffset = 1 + sizeof(std::uint64_t);

// Appends a picture's identity: its media id, an alt-present byte and the alt's
// `kCompare` collation key, read from the published picture table. With no
// table, or an id it does not hold, the identity is the picture id.
inline void append_picture_identity_key(PicturePayload picture, ExcelProfile profile, std::string* out) {
  const auto append_bytes = [out](const void* data, std::size_t size) {
    out->append(static_cast<const char*>(data), size);
  };
  const PictureTable* table = current_picture_table();
  const PictureEntry* entry = table != nullptr ? table->entry(picture.picture_id) : nullptr;
  if (entry == nullptr) {
    out->push_back('\1');
    append_bytes(&picture.picture_id, sizeof(picture.picture_id));
    return;
  }
  out->push_back('\0');
  append_bytes(&entry->media_id, sizeof(entry->media_id));
  out->push_back(entry->alt.has_value() ? '\1' : '\0');
  if (entry->alt.has_value()) {
    const std::string key = collation_key(*entry->alt, profile, CollationStrength::kCompare);
    const auto size = static_cast<std::uint64_t>(key.size());
    append_bytes(&size, sizeof(size));
    out->append(key);
  }
}

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
    case ValueKind::Picture:
      append_picture_identity_key(value.as_picture(), profile, out);
      return true;
    default:
      // Array / Ref / Lambda values do not have a canonical cell identity.
      return false;
  }
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_VALUE_IDENTITY_H_
