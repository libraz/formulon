//
// SDDL text <-> self-relative SECURITY_DESCRIPTOR bytes for the
// security descriptors range protections carry.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/protection_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// SECURITY_DESCRIPTOR control (SE_SELF_RELATIVE, with and without SE_DACL_PRESENT), ACL revision, ACE types.
constexpr std::uint16_t kSdControl = 0x8004U;
constexpr std::uint16_t kSdControlNoDacl = 0x8000U;
constexpr std::uint8_t kAclRevision = 2U;
constexpr std::uint8_t kAclRevisionDs = 4U;        // the empty DACL Excel writes for an object ACE
constexpr std::size_t kObjectAceFixedBytes = 12U;  // header, mask, object flags
constexpr std::size_t kGuidTextChars = 36U;
constexpr std::uint8_t kAccessAllowedAce = 0U;
constexpr std::uint8_t kAccessDeniedAce = 1U;
constexpr std::size_t kSdHeaderBytes = 20U;
constexpr std::size_t kAclHeaderBytes = 8U;
constexpr std::size_t kAceHeaderBytes = 8U;

// A longer SID literal matches no alias value.
// @limit: internal
constexpr std::size_t kMaxSidSubAuthorities = 6;

/// An SDDL SID alias with a fixed value: S-1-`authority`-`sub`... .
struct SidAlias {
  char name[3];
  std::uint8_t authority;
  std::uint8_t count;
  std::uint32_t sub[kMaxSidSubAuthorities];
};

constexpr SidAlias kSidAliases[] = {
    {"WD", 1, 1, {0}},       {"CO", 3, 1, {0}},       {"CG", 3, 1, {1}},       {"OW", 3, 1, {4}},
    {"NU", 5, 1, {2}},       {"IU", 5, 1, {4}},       {"SU", 5, 1, {6}},       {"AN", 5, 1, {7}},
    {"ED", 5, 1, {9}},       {"PS", 5, 1, {10}},      {"AU", 5, 1, {11}},      {"RC", 5, 1, {12}},
    {"SY", 5, 1, {18}},      {"LS", 5, 1, {19}},      {"NS", 5, 1, {20}},      {"WR", 5, 1, {33}},
    {"BA", 5, 2, {32, 544}}, {"BU", 5, 2, {32, 545}}, {"BG", 5, 2, {32, 546}}, {"PU", 5, 2, {32, 547}},
    {"AO", 5, 2, {32, 548}}, {"SO", 5, 2, {32, 549}}, {"PO", 5, 2, {32, 550}}, {"BO", 5, 2, {32, 551}},
    {"RE", 5, 2, {32, 552}}, {"RU", 5, 2, {32, 554}}, {"RD", 5, 2, {32, 555}}, {"NO", 5, 2, {32, 556}},
    {"MU", 5, 2, {32, 558}}, {"LU", 5, 2, {32, 559}}, {"IS", 5, 2, {32, 568}}, {"CY", 5, 2, {32, 569}},
    {"ER", 5, 2, {32, 573}}, {"CD", 5, 2, {32, 574}}, {"RA", 5, 2, {32, 575}}, {"ES", 5, 2, {32, 576}},
    {"MS", 5, 2, {32, 577}}, {"HA", 5, 2, {32, 578}}, {"AA", 5, 2, {32, 579}}, {"RM", 5, 2, {32, 580}},
    {"UD", 5, 6, {84}},      {"AC", 15, 2, {2, 1}},   {"LW", 16, 1, {4096}},   {"ME", 16, 1, {8192}},
    {"MP", 16, 1, {8448}},   {"HI", 16, 1, {12288}},  {"SI", 16, 1, {16384}},  {"AS", 18, 1, {1}},
    {"SS", 18, 1, {2}},
};

struct RightToken {
  char name[3];
  std::uint32_t mask;
};

// One-bit rights in ascending mask order, then the composite ones.
constexpr RightToken kRightTokens[] = {
    {"CC", 0x00000001U}, {"DC", 0x00000002U}, {"LC", 0x00000004U}, {"SW", 0x00000008U}, {"RP", 0x00000010U},
    {"WP", 0x00000020U}, {"DT", 0x00000040U}, {"LO", 0x00000080U}, {"CR", 0x00000100U}, {"SD", 0x00010000U},
    {"RC", 0x00020000U}, {"WD", 0x00040000U}, {"WO", 0x00080000U}, {"GA", 0x10000000U}, {"GX", 0x20000000U},
    {"GW", 0x40000000U}, {"GR", 0x80000000U}, {"FA", 0x001F01FFU}, {"FR", 0x00120089U}, {"FW", 0x00120116U},
    {"FX", 0x001200A0U}, {"KA", 0x000F003FU}, {"KR", 0x00020019U}, {"KW", 0x00020006U}, {"KX", 0x00020019U},
};
constexpr std::size_t kOneBitRightCount = 17;
constexpr std::size_t kRightCount = sizeof(kRightTokens) / sizeof(kRightTokens[0]);

bool IsUpper(char c) {
  return c >= 'A' && c <= 'Z';
}

bool IsLower(char c) {
  return c >= 'a' && c <= 'z';
}

char ToUpper(char c) {
  return IsLower(c) ? static_cast<char>(c - 'a' + 'A') : c;
}

/// Whether `t` holds `s` at `i`.
bool At(std::string_view t, std::size_t i, std::string_view s) {
  return i <= t.size() && t.size() - i >= s.size() && t.substr(i, s.size()) == s;
}

void EmitAliasSid(const SidAlias& alias, std::vector<std::uint8_t>& sid) {
  sid.push_back(1U);
  sid.push_back(alias.count);
  sid.insert(sid.end(), {0U, 0U, 0U, 0U, 0U, alias.authority});
  for (std::uint8_t k = 0; k < alias.count; ++k) {
    emit_u32(sid, alias.sub[k]);
  }
}

/// Reads decimal digits at `t[i]` into `value`; false when there are none or it passes `limit`.
bool ParseDecimal(std::string_view t, std::size_t& i, std::uint64_t limit, std::uint64_t& value) {
  const std::size_t start = i;
  value = 0;
  while (i < t.size() && t[i] >= '0' && t[i] <= '9') {
    value = value * 10U + static_cast<std::uint64_t>(t[i] - '0');
    if (value > limit) {
      return false;
    }
    ++i;
  }
  return i != start;
}

/// Parses a `S-1-...` SID literal at `t[i]`; encoded only when it is the value of an alias, as Excel
/// then writes the alias's bytes.
SddlEncode ParseSidLiteral(std::string_view t, std::size_t& i, std::vector<std::uint8_t>& sid) {
  if (!At(t, i, "S-1-")) {
    return SddlEncode::kUnsupported;
  }
  i += 4U;
  std::uint64_t authority = 0;
  if (!ParseDecimal(t, i, 0xFFU, authority)) {
    return SddlEncode::kUnsupported;
  }
  std::uint32_t sub[kMaxSidSubAuthorities] = {};
  std::size_t count = 0;
  while (At(t, i, "-")) {
    ++i;
    std::uint64_t value = 0;
    if (count == kMaxSidSubAuthorities || !ParseDecimal(t, i, 0xFFFFFFFFU, value)) {
      return SddlEncode::kUnsupported;
    }
    sub[count++] = static_cast<std::uint32_t>(value);
  }
  for (const SidAlias& alias : kSidAliases) {
    if (alias.authority == authority && alias.count == count && std::equal(sub, sub + count, alias.sub)) {
      EmitAliasSid(alias, sid);
      return SddlEncode::kEncoded;
    }
  }
  return SddlEncode::kUnsupported;
}

/// Parses a SID alias at `t[i]`, in either case, into its binary form.
SddlEncode ParseSid(std::string_view t, std::size_t& i, std::vector<std::uint8_t>& sid) {
  if (At(t, i, "S-")) {
    return ParseSidLiteral(t, i, sid);
  }
  if (i + 2U > t.size() || !(IsUpper(t[i]) || IsLower(t[i])) || !(IsUpper(t[i + 1U]) || IsLower(t[i + 1U]))) {
    return SddlEncode::kUnsupported;
  }
  const char upper[2] = {ToUpper(t[i]), ToUpper(t[i + 1U])};
  const std::string_view name(upper, 2U);
  i += 2U;
  for (const SidAlias& alias : kSidAliases) {
    if (name == alias.name) {
      EmitAliasSid(alias, sid);
      return SddlEncode::kEncoded;
    }
  }
  return SddlEncode::kMalformed;
}

/// Parses the right tokens at `t[i]` up to the next `;` into `mask`; false when one is unknown or none is given.
bool ParseRights(std::string_view t, std::size_t& i, std::uint32_t& mask) {
  mask = 0;
  // A hex mask stands alone; Excel encodes it as the mask itself.
  if (At(t, i, "0x")) {
    i += 2U;
    const std::size_t start = i;
    while (i < t.size() && t[i] != ';') {
      const char c = t[i];
      const int digit = strings::hex_digit_value(c);
      if (digit < 0 || i - start == 8U) {
        return false;
      }
      mask = (mask << 4U) | static_cast<std::uint32_t>(digit);
      ++i;
    }
    return mask != 0U;
  }
  while (i + 2U <= t.size() && t[i] != ';') {
    const std::string_view token = t.substr(i, 2U);
    const RightToken* hit = nullptr;
    for (const RightToken& right : kRightTokens) {
      if (token == right.name) {
        hit = &right;
        break;
      }
    }
    if (hit == nullptr) {
      return false;
    }
    mask |= hit->mask;
    i += 2U;
  }
  return mask != 0U;
}

/// Parses one `(A|D;;rights;;;sid)` ACE at `t[i]` and appends it to `acl`.
SddlEncode ParseAce(std::string_view t, std::size_t& i, std::vector<std::uint8_t>& acl) {
  ++i;  // '('
  if (!At(t, i, "A;;") && !At(t, i, "D;;")) {
    return SddlEncode::kUnsupported;
  }
  const std::uint8_t type = t[i] == 'A' ? kAccessAllowedAce : kAccessDeniedAce;
  i += 3U;
  std::uint32_t mask = 0;
  if (!ParseRights(t, i, mask) || !At(t, i, ";;;")) {
    return SddlEncode::kUnsupported;
  }
  i += 3U;
  std::vector<std::uint8_t> sid;
  const SddlEncode read = ParseSid(t, i, sid);
  if (read == SddlEncode::kUnsupported || i >= t.size() || t[i] != ')') {
    return SddlEncode::kUnsupported;
  }
  ++i;
  if (read == SddlEncode::kMalformed) {
    return read;
  }
  acl.push_back(type);
  acl.push_back(0U);
  emit_u16(acl, static_cast<std::uint16_t>(kAceHeaderBytes + sid.size()));
  emit_u32(acl, mask);
  acl.insert(acl.end(), sid.begin(), sid.end());
  return SddlEncode::kEncoded;
}

/// Parses one `(OA;;rights;[object guid];;sid)` ACE at `t[i]`; `size` is the bytes it would take.
SddlEncode ParseObjectAce(std::string_view t, std::size_t& i, std::size_t& size) {
  i += 5U;  // "(OA;;"
  std::uint32_t mask = 0;
  if (!ParseRights(t, i, mask) || !At(t, i, ";")) {
    return SddlEncode::kUnsupported;
  }
  ++i;
  std::size_t guid_bytes = 0;
  if (!At(t, i, ";")) {
    if (t.size() - i < kGuidTextChars) {
      return SddlEncode::kUnsupported;
    }
    for (std::size_t k = 0; k < kGuidTextChars; ++k) {
      const char c = t[i + k];
      const bool dash = k == 8U || k == 13U || k == 18U || k == 23U;
      const bool hex = strings::hex_digit_value(c) >= 0;
      if (dash ? c != '-' : !hex) {
        return SddlEncode::kUnsupported;
      }
    }
    i += kGuidTextChars;
    guid_bytes = 16U;
  }
  if (!At(t, i, ";;")) {
    return SddlEncode::kUnsupported;
  }
  i += 2U;
  std::vector<std::uint8_t> sid;
  const SddlEncode read = ParseSid(t, i, sid);
  if (read == SddlEncode::kUnsupported || i >= t.size() || t[i] != ')') {
    return SddlEncode::kUnsupported;
  }
  ++i;
  size = kObjectAceFixedBytes + guid_bytes + sid.size();
  return read;
}

/// Appends the alias of the SID at `sd[at]` to `out`; false when it has none here.
bool SidAliasAt(ByteSpan sd, std::size_t at, std::string& out, std::size_t& size) {
  if (at > sd.size || sd.size - at < 8U || sd.data[at] != 1U) {
    return false;
  }
  const std::size_t count = sd.data[at + 1U];
  size = 8U + 4U * count;
  if (sd.size - at < size) {
    return false;
  }
  const bool high_authority_zero =
      std::all_of(sd.data + at + 2U, sd.data + at + 7U, [](std::uint8_t b) { return b == 0U; });
  for (const SidAlias& alias : kSidAliases) {
    bool same = high_authority_zero && alias.count == count && sd.data[at + 7U] == alias.authority;
    for (std::size_t k = 0; same && k < count; ++k) {
      same = load_u32(sd.data + at + 8U + 4U * k) == alias.sub[k];
    }
    if (same) {
      out.append(alias.name);
      return true;
    }
  }
  return false;
}

/// Appends the SDDL rights of `mask`: the composite token with exactly that mask, else its one-bit tokens in
/// ascending order. False when only a hex spelling remains, which is not emitted.
bool AppendRights(std::uint32_t mask, std::string& out) {
  for (std::size_t k = kOneBitRightCount; k < kRightCount; ++k) {
    if (kRightTokens[k].mask == mask) {
      out.append(kRightTokens[k].name);
      return true;
    }
  }
  std::string tokens;
  std::uint32_t rest = mask;
  for (std::size_t k = 0; k < kOneBitRightCount; ++k) {
    if ((mask & kRightTokens[k].mask) != 0U) {
      tokens.append(kRightTokens[k].name);
      rest &= ~kRightTokens[k].mask;
    }
  }
  if (rest != 0U) {
    return false;
  }
  out.append(tokens);
  return true;
}

/// Appends the ACEs of the ACL at `sd[at]` to `out`; `size` is the ACL's byte size.
bool AclAt(ByteSpan sd, std::size_t at, std::string& out, std::size_t& size) {
  if (at > sd.size || sd.size - at < kAclHeaderBytes ||
      (sd.data[at] != kAclRevision && sd.data[at] != kAclRevisionDs) || sd.data[at + 1U] != 0U ||
      load_u16(sd.data + at + 6U) != 0U) {
    return false;
  }
  size = load_u16(sd.data + at + 2U);
  if (size < kAclHeaderBytes || sd.size - at < size) {
    return false;
  }
  const std::size_t count = load_u16(sd.data + at + 4U);
  const std::size_t end = at + size;
  // The empty DACL Excel writes in an object ACE's place: no ACE, the rest zero.
  if (sd.data[at] == kAclRevisionDs) {
    return count == 0U &&
           std::all_of(sd.data + at + kAclHeaderBytes, sd.data + end, [](std::uint8_t b) { return b == 0U; });
  }
  std::size_t ace = at + kAclHeaderBytes;
  for (std::size_t k = 0; k < count; ++k) {
    if (end - ace < kAceHeaderBytes || sd.data[ace] > kAccessDeniedAce || sd.data[ace + 1U] != 0U) {
      return false;
    }
    const std::uint32_t mask = load_u32(sd.data + ace + 4U);
    std::string sid;
    std::size_t sid_size = 0;
    if (mask == 0U || !SidAliasAt(sd, ace + kAceHeaderBytes, sid, sid_size) ||
        static_cast<std::size_t>(load_u16(sd.data + ace + 2U)) != kAceHeaderBytes + sid_size ||
        end - ace < kAceHeaderBytes + sid_size) {
      return false;
    }
    out.append(sd.data[ace] == kAccessAllowedAce ? "(A;;" : "(D;;");
    if (!AppendRights(mask, out)) {
      return false;
    }
    out.append(";;;").append(sid).push_back(')');
    ace += kAceHeaderBytes + sid_size;
  }
  return ace == end;
}

}  // namespace

SddlEncode encode_security_descriptor(std::string_view sddl, std::vector<std::uint8_t>& out) {
  const std::string_view t = sddl;
  // Text with no upper-case letter is written without a descriptor when its upper-case
  // spelling encodes.
  if (!t.empty() && std::none_of(t.begin(), t.end(), IsUpper) && std::any_of(t.begin(), t.end(), IsLower)) {
    std::string upper(t);
    std::transform(upper.begin(), upper.end(), upper.begin(), ToUpper);
    std::vector<std::uint8_t> ignored;
    return encode_security_descriptor(upper, ignored) == SddlEncode::kEncoded ? SddlEncode::kMalformed
                                                                              : SddlEncode::kUnsupported;
  }
  std::size_t i = 0;
  std::vector<std::uint8_t> owner;
  std::vector<std::uint8_t> group;
  std::vector<std::uint8_t> aces;
  std::uint16_t ace_count = 0;
  bool has_dacl = false;
  bool unfixed_sid = false;  // a SID with no fixed value: the range is written without a descriptor
  std::size_t object_aces = 0;
  std::size_t object_ace_bytes = 0;
  std::string order;
  while (i < t.size()) {
    if (t.size() < i + 2U || t[i + 1U] != ':' || order.find(t[i]) != std::string::npos) {
      return SddlEncode::kUnsupported;
    }
    const char tag = t[i];
    order.push_back(tag);
    i += 2U;
    if (tag == 'O' || tag == 'G') {
      const SddlEncode read = ParseSid(t, i, tag == 'O' ? owner : group);
      if (read == SddlEncode::kUnsupported) {
        return read;
      }
      unfixed_sid = unfixed_sid || read == SddlEncode::kMalformed;
      continue;
    }
    if (tag != 'D') {
      return SddlEncode::kUnsupported;
    }
    has_dacl = true;
    while (i < t.size() && t[i] == '(') {
      if (At(t, i, "(OA;;")) {
        std::size_t size = 0;
        const SddlEncode read = ParseObjectAce(t, i, size);
        if (read == SddlEncode::kUnsupported) {
          return read;
        }
        unfixed_sid = unfixed_sid || read == SddlEncode::kMalformed;
        ++object_aces;
        object_ace_bytes += size;
        continue;
      }
      const SddlEncode read = ParseAce(t, i, aces);
      if (read == SddlEncode::kUnsupported) {
        return read;
      }
      unfixed_sid = unfixed_sid || read == SddlEncode::kMalformed;
      if (++ace_count == 0U) {
        return SddlEncode::kUnsupported;
      }
    }
  }
  if (unfixed_sid) {
    return SddlEncode::kMalformed;
  }
  // The section orders Excel encodes alike; an object ACE is encoded only as the DACL's one ACE.
  if ((order != "OG" && order != "OGD" && order != "GOD" && order != "DOG") ||
      kAclHeaderBytes + aces.size() > 0xFFFFU || (object_aces != 0U && (object_aces != 1U || ace_count != 0U))) {
    return SddlEncode::kUnsupported;
  }
  // Excel drops an object ACE and keeps its space: an empty DACL of the same size, zero-filled.
  const bool object_dacl = object_aces != 0U;
  if (object_dacl) {
    aces.assign(object_ace_bytes, 0U);
  }
  const std::uint32_t acl_size = has_dacl ? static_cast<std::uint32_t>(kAclHeaderBytes + aces.size()) : 0U;
  const std::uint32_t owner_at = static_cast<std::uint32_t>(kSdHeaderBytes) + acl_size;
  const std::uint32_t group_at = owner_at + static_cast<std::uint32_t>(owner.size());
  // Header, DACL (when present), owner, group.
  out.clear();
  out.push_back(1U);
  out.push_back(0U);
  emit_u16(out, has_dacl ? kSdControl : kSdControlNoDacl);
  emit_u32(out, owner_at);
  emit_u32(out, group_at);
  emit_u32(out, 0U);  // no SACL
  emit_u32(out, has_dacl ? static_cast<std::uint32_t>(kSdHeaderBytes) : 0U);
  if (has_dacl) {
    out.push_back(object_dacl ? kAclRevisionDs : kAclRevision);
    out.push_back(0U);
    emit_u16(out, static_cast<std::uint16_t>(acl_size));
    emit_u16(out, ace_count);
    emit_u16(out, 0U);
    out.insert(out.end(), aces.begin(), aces.end());
  }
  out.insert(out.end(), owner.begin(), owner.end());
  out.insert(out.end(), group.begin(), group.end());
  return SddlEncode::kEncoded;
}

bool decode_security_descriptor(ByteSpan sd, std::string& out) {
  if (sd.size < kSdHeaderBytes || sd.data[0] != 1U || sd.data[1] != 0U) {
    return false;
  }
  const std::uint16_t control = load_u16(sd.data + 2U);
  const bool has_dacl = control == kSdControl;
  const std::size_t owner_at = load_u32(sd.data + 4U);
  const std::size_t group_at = load_u32(sd.data + 8U);
  const std::size_t dacl_at = load_u32(sd.data + 16U);
  if ((!has_dacl && control != kSdControlNoDacl) || load_u32(sd.data + 12U) != 0U || (dacl_at != 0U) != has_dacl) {
    return false;
  }
  std::string text = "O:";
  std::size_t owner_size = 0;
  std::size_t group_size = 0;
  std::size_t dacl_size = 0;
  if (!SidAliasAt(sd, owner_at, text, owner_size)) {
    return false;
  }
  text.append("G:");
  if (!SidAliasAt(sd, group_at, text, group_size)) {
    return false;
  }
  if (has_dacl) {
    text.append("D:");
    if (!AclAt(sd, dacl_at, text, dacl_size)) {
      return false;
    }
  }
  // The parts follow the header back to back, in any order.
  const std::array<std::pair<std::size_t, std::size_t>, 3> parts = {
      {{owner_at, owner_size}, {group_at, group_size}, {dacl_at, dacl_size}}};
  const std::size_t part_count = has_dacl ? 3U : 2U;
  const auto parts_end = parts.begin() + static_cast<std::ptrdiff_t>(part_count);
  std::size_t next = kSdHeaderBytes;
  for (std::size_t round = 0; round < part_count; ++round) {
    const auto hit = std::find_if(parts.begin(), parts_end, [next](const std::pair<std::size_t, std::size_t>& part) {
      return part.first == next;
    });
    if (hit == parts_end) {
      return false;
    }
    next += hit->second;
  }
  if (next != sd.size) {
    return false;
  }
  out = std::move(text);
  return true;
}

std::string canonical_security_descriptor(std::string_view sddl) {
  std::vector<std::uint8_t> sd;
  switch (encode_security_descriptor(sddl, sd)) {
    case SddlEncode::kEncoded: {
      std::string text;
      return decode_security_descriptor(ByteSpan{sd.data(), sd.size()}, text) ? text : std::string(sddl);
    }
    case SddlEncode::kMalformed:
      return std::string();
    case SddlEncode::kUnsupported:
      break;
  }
  return std::string(sddl);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
