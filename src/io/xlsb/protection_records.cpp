// @size-budget: 56 KB

#include "io/xlsb/protection_records.h"

#include <algorithm>
#include <array>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "io/protection_xml.h"
#include "io/sqref_text.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "pugixml.hpp"
#include "utils/number_text.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::size_t kFlagCount = 16;
// @limit: format
constexpr std::uint32_t kMaxBlobBytes = 1024;

/// The fifteen action flags after fLocked, in record order. XLSB stores
/// each as the inverse of the OOXML attribute.
constexpr std::array<bool SheetProtection::*, kFlagCount - 1> kActionFlags = {
    &SheetProtection::objects,        &SheetProtection::scenarios,           &SheetProtection::format_cells,
    &SheetProtection::format_columns, &SheetProtection::format_rows,         &SheetProtection::insert_columns,
    &SheetProtection::insert_rows,    &SheetProtection::insert_hyperlinks,   &SheetProtection::delete_columns,
    &SheetProtection::delete_rows,    &SheetProtection::select_locked_cells, &SheetProtection::sort,
    &SheetProtection::auto_filter,    &SheetProtection::pivot_tables,        &SheetProtection::select_unlocked_cells,
};

/// Reads the 16 u32 booleans; false on a value other than 0 or 1.
bool ReadFlags(ByteSpan& p, std::array<bool, kFlagCount>& flags) {
  for (bool& flag : flags) {
    auto v = read_u32(p);
    if (!v || v.value() > 1U) {
      return false;
    }
    flag = v.value() != 0U;
  }
  return true;
}

void ApplyFlags(const std::array<bool, kFlagCount>& flags, SheetProtection& out) {
  out.enabled = flags[0];
  out.sheet = flags[0];
  for (std::size_t i = 0; i < kActionFlags.size(); ++i) {
    out.*kActionFlags[i] = !flags[i + 1];
  }
}

void EmitFlags(std::vector<std::uint8_t>& dst, const SheetProtection& p) {
  emit_u32(dst, p.sheet ? 1U : 0U);
  for (bool SheetProtection::*member : kActionFlags) {
    emit_u32(dst, p.*member ? 0U : 1U);
  }
}

bool ReadBlob(ByteSpan& p, std::vector<std::uint8_t>& out) {
  auto n = read_u32(p);
  if (!n || n.value() > kMaxBlobBytes || n.value() > p.size) {
    return false;
  }
  out.assign(p.data, p.data + n.value());
  p.data += n.value();
  p.size -= n.value();
  return true;
}

std::string LegacyHex(std::uint16_t pwd) {
  char buf[5];
  format_hex(buf, sizeof(buf), pwd, 4, true);
  return buf;
}

bool ParseLegacyHex(const std::string& text, std::uint16_t& out) {
  if (text.empty() || text.size() > 4U) {
    return false;
  }
  std::uint32_t v = 0;
  for (char c : text) {
    const int d = strings::hex_digit_value(c);
    if (d < 0) {
      return false;
    }
    v = v * 16U + static_cast<std::uint32_t>(d);
  }
  out = static_cast<std::uint16_t>(v);
  return true;
}

constexpr char kBase64Alphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

int Base64Value(char c) {
  for (int i = 0; i < 64; ++i) {
    if (kBase64Alphabet[i] == c) {
      return i;
    }
  }
  return -1;
}

}  // namespace

std::string base64_encode(const std::uint8_t* data, std::size_t size) {
  std::string out;
  out.reserve((size + 2U) / 3U * 4U);
  for (std::size_t i = 0; i < size; i += 3U) {
    const std::size_t n = size - i < 3U ? size - i : 3U;
    std::uint32_t chunk = static_cast<std::uint32_t>(data[i]) << 16U;
    if (n > 1U) {
      chunk |= static_cast<std::uint32_t>(data[i + 1U]) << 8U;
    }
    if (n > 2U) {
      chunk |= data[i + 2U];
    }
    for (std::size_t k = 0; k < 4U; ++k) {
      out.push_back(k <= n ? kBase64Alphabet[(chunk >> (18U - 6U * k)) & 0x3FU] : '=');
    }
  }
  return out;
}

bool base64_decode(const std::string& text, std::vector<std::uint8_t>& out) {
  out.clear();
  if (text.size() % 4U != 0U) {
    return false;
  }
  for (std::size_t i = 0; i < text.size(); i += 4U) {
    std::uint32_t chunk = 0;
    std::size_t pad = 0;
    for (std::size_t k = 0; k < 4U; ++k) {
      const char c = text[i + k];
      if (c == '=' && i + 4U == text.size() && k >= 2U) {
        ++pad;
        chunk <<= 6U;
        continue;
      }
      const int v = Base64Value(c);
      if (v < 0 || pad != 0U) {
        return false;
      }
      chunk = (chunk << 6U) | static_cast<std::uint32_t>(v);
    }
    if ((pad == 2U && (chunk & 0xFFFFU) != 0U) || (pad == 1U && (chunk & 0xFFU) != 0U)) {
      return false;
    }
    out.push_back(static_cast<std::uint8_t>(chunk >> 16U));
    if (pad < 2U) {
      out.push_back(static_cast<std::uint8_t>(chunk >> 8U));
    }
    if (pad < 1U) {
      out.push_back(static_cast<std::uint8_t>(chunk));
    }
  }
  return true;
}

bool decode_sheet_protection(ByteSpan payload, SheetProtection& out) {
  ByteSpan p = payload;
  auto pwd = read_u16(p);
  std::array<bool, kFlagCount> flags{};
  if (!pwd || !ReadFlags(p, flags) || p.size != 0U) {
    return false;
  }
  if (!flags[0]) {
    out = SheetProtection{};
    return true;
  }
  ApplyFlags(flags, out);
  out.legacy_password = pwd.value() != 0U ? LegacyHex(pwd.value()) : std::string();
  return true;
}

bool decode_sheet_protection_iso(ByteSpan payload, SheetProtection& out) {
  ByteSpan p = payload;
  auto spin = read_u32(p);
  std::array<bool, kFlagCount> flags{};
  std::vector<std::uint8_t> hash;
  std::vector<std::uint8_t> salt;
  if (!spin || !ReadFlags(p, flags) || !ReadBlob(p, hash) || !ReadBlob(p, salt)) {
    return false;
  }
  auto algorithm = read_xlwidestring(p);
  if (!algorithm || p.size != 0U) {
    return false;
  }
  if (!flags[0]) {
    return true;
  }
  out.spin_count = spin.value();
  out.hash_value = base64_encode(hash.data(), hash.size());
  out.salt_value = base64_encode(salt.data(), salt.size());
  out.algorithm_name = std::move(algorithm).value();
  return true;
}

Expected<void, Error> emit_sheet_protection(std::vector<std::uint8_t>& dst, const SheetProtection& protection) {
  if (!protection.enabled) {
    return Expected<void, Error>::Ok();
  }
  std::uint16_t pwd = 0;
  if (!protection.legacy_password.empty() && !ParseLegacyHex(protection.legacy_password, pwd)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet protection legacy password is not a 16-bit hex hash",
                      kWriteXlsbContext);
  }
  std::vector<std::uint8_t> payload;
  if (!protection.algorithm_name.empty()) {
    std::vector<std::uint8_t> hash;
    std::vector<std::uint8_t> salt;
    if (!base64_decode(protection.hash_value, hash) || !base64_decode(protection.salt_value, salt) ||
        hash.size() > kMaxBlobBytes || salt.size() > kMaxBlobBytes) {
      return make_error(FormulonErrorCode::kInvalidArgument, "sheet protection hash or salt is not base64",
                        kWriteXlsbContext);
    }
    emit_u32(payload, protection.spin_count);
    EmitFlags(payload, protection);
    emit_u32(payload, static_cast<std::uint32_t>(hash.size()));
    payload.insert(payload.end(), hash.begin(), hash.end());
    emit_u32(payload, static_cast<std::uint32_t>(salt.size()));
    payload.insert(payload.end(), salt.begin(), salt.end());
    emit_xlwidestring(payload, protection.algorithm_name);
    emit_record(dst, kBrtSheetProtectionIso, payload);
    payload.clear();
  }
  emit_u16(payload, pwd);
  EmitFlags(payload, protection);
  emit_record(dst, kBrtSheetProtection, payload);
  return Expected<void, Error>::Ok();
}

namespace {

/// IsoPasswordData ([MS-XLSB] 2.5.80): hash, salt, nullable algorithm name.
bool ReadIsoPassword(ByteSpan& p, std::string& algorithm, std::string& hash, std::string& salt) {
  std::vector<std::uint8_t> hash_bytes;
  std::vector<std::uint8_t> salt_bytes;
  if (!ReadBlob(p, hash_bytes) || !ReadBlob(p, salt_bytes)) {
    return false;
  }
  auto name = read_xlnullablewidestring(p);
  if (!name) {
    return false;
  }
  algorithm = std::move(name).value();
  if (!algorithm.empty()) {
    hash = base64_encode(hash_bytes.data(), hash_bytes.size());
    salt = base64_encode(salt_bytes.data(), salt_bytes.size());
  }
  return true;
}

bool EncodeIsoPassword(std::vector<std::uint8_t>& dst, const std::string& algorithm, const std::string& hash_text,
                       const std::string& salt_text) {
  std::vector<std::uint8_t> hash;
  std::vector<std::uint8_t> salt;
  if (!algorithm.empty() && (!base64_decode(hash_text, hash) || !base64_decode(salt_text, salt) ||
                             hash.size() > kMaxBlobBytes || salt.size() > kMaxBlobBytes)) {
    return false;
  }
  emit_u32(dst, static_cast<std::uint32_t>(hash.size()));
  dst.insert(dst.end(), hash.begin(), hash.end());
  emit_u32(dst, static_cast<std::uint32_t>(salt.size()));
  dst.insert(dst.end(), salt.begin(), salt.end());
  emit_xlnullablewidestring(dst, algorithm.empty() ? std::nullopt : std::make_optional<std::string_view>(algorithm));
  return true;
}

// BookProtectionFlags ([MS-XLSB] 2.5.7); the reserved bits are ignored.
constexpr std::uint16_t kLockStructure = 1U << 0U;
constexpr std::uint16_t kLockWindows = 1U << 1U;
constexpr std::uint16_t kLockRevision = 1U << 2U;

}  // namespace

bool decode_book_protection(ByteSpan payload, ByteSpan iso, std::string& xml) {
  ByteSpan p = payload;
  auto book_pwd = read_u16(p);
  auto rev_pwd = read_u16(p);
  auto flags = read_u16(p);
  if (!book_pwd || !rev_pwd || !flags || p.size != 0U) {
    return false;
  }
  WorkbookProtection out;
  const std::uint16_t f = flags.value();
  out.lock_structure = (f & kLockStructure) != 0U;
  out.lock_windows = (f & kLockWindows) != 0U;
  out.lock_revision = (f & kLockRevision) != 0U;
  if (book_pwd.value() != 0U) {
    out.workbook_password = LegacyHex(book_pwd.value());
  }
  if (rev_pwd.value() != 0U) {
    out.revisions_password = LegacyHex(rev_pwd.value());
  }
  if (iso.size != 0U) {
    ByteSpan q = iso;
    auto book_spin = read_u32(q);
    auto rev_spin = read_u32(q);
    auto iso_flags = read_u16(q);
    if (!book_spin || !rev_spin || !iso_flags || iso_flags.value() != f ||
        !ReadIsoPassword(q, out.workbook_algorithm_name, out.workbook_hash_value, out.workbook_salt_value) ||
        !ReadIsoPassword(q, out.revisions_algorithm_name, out.revisions_hash_value, out.revisions_salt_value) ||
        q.size != 0U) {
      return false;
    }
    if (!out.workbook_algorithm_name.empty()) {
      out.workbook_spin_count = book_spin.value();
    }
    if (!out.revisions_algorithm_name.empty()) {
      out.revisions_spin_count = rev_spin.value();
    }
  }
  xml = apply_workbook_protection({}, out);
  return true;
}

Expected<bool, Error> emit_book_protection(std::vector<std::uint8_t>& dst, const std::string& xml) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  if (!doc.load_string(xml.c_str()) || !doc.child("workbookProtection")) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Workbook protection is not a <workbookProtection> element",
                      kWriteXlsbContext);
  }
  const WorkbookProtection prot = parse_workbook_protection(xml);
  std::uint16_t book_pwd = 0;
  std::uint16_t rev_pwd = 0;
  if ((!prot.workbook_password.empty() && !ParseLegacyHex(prot.workbook_password, book_pwd)) ||
      (!prot.revisions_password.empty() && !ParseLegacyHex(prot.revisions_password, rev_pwd))) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Workbook password is not a 16-bit hex hash",
                      kWriteXlsbContext);
  }
  std::uint16_t flags = 0U;
  flags |= prot.lock_structure ? kLockStructure : 0U;
  flags |= prot.lock_windows ? kLockWindows : 0U;
  flags |= prot.lock_revision ? kLockRevision : 0U;
  const bool has_iso = !prot.workbook_algorithm_name.empty() || !prot.revisions_algorithm_name.empty();
  bool carried = true;
  if (has_iso) {
    std::vector<std::uint8_t> iso;
    emit_u32(iso, prot.workbook_spin_count);
    emit_u32(iso, prot.revisions_spin_count);
    emit_u16(iso, flags);
    if (!EncodeIsoPassword(iso, prot.workbook_algorithm_name, prot.workbook_hash_value, prot.workbook_salt_value) ||
        !EncodeIsoPassword(iso, prot.revisions_algorithm_name, prot.revisions_hash_value, prot.revisions_salt_value)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "Workbook protection hash or salt is not base64",
                        kWriteXlsbContext);
    }
    emit_record(dst, kBrtBookProtectionIso, iso);
    // BrtBookProtection after the Iso record must carry zero verifiers.
    carried = book_pwd == 0U && rev_pwd == 0U;
    book_pwd = 0U;
    rev_pwd = 0U;
  }
  if (book_pwd != 0U || rev_pwd != 0U || flags != 0U || has_iso) {
    std::vector<std::uint8_t> payload;
    emit_u16(payload, book_pwd);
    emit_u16(payload, rev_pwd);
    emit_u16(payload, flags);
    emit_record(dst, kBrtBookProtection, payload);
  }
  return carried;
}

namespace {

// FRTHeader flags (fSqref) and FRTSqref flags (fDoAdjust) of BrtRangeProtection14.
constexpr std::uint32_t kFrtHeaderSqref = 2U;
constexpr std::uint32_t kFrtSqrefDoAdjust = 2U;

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

namespace {

bool ReadSqRfx(ByteSpan& p, std::vector<MergeRange>& out, std::size_t max_count) {
  auto count = read_u32(p);
  if (!count || count.value() == 0U || count.value() > max_count || p.size / 16U < count.value()) {
    return false;
  }
  for (std::uint32_t i = 0; i < count.value(); ++i) {
    auto rfx = read_rfx(p);
    if (!rfx) {
      return false;
    }
    const MergeRange& r = rfx.value();
    if (r.first_row > r.last_row || r.first_col > r.last_col || !Sheet::coord_in_grid(r.last_row, r.last_col)) {
      return false;
    }
    out.push_back(r);
  }
  return true;
}

/// Reads RangeProtectionTitleSDRel; `descriptor` receives the dwsdRel bytes.
bool ReadTitle(ByteSpan& p, std::string& title, ByteSpan& descriptor) {
  auto name = read_xlwidestring(p);
  auto sd_size = read_u32(p);
  if (!name || name.value().empty() || utf16_units_in(name.value()) > kMaxProtectedRangeNameUnits || !sd_size ||
      sd_size.value() > p.size) {
    return false;
  }
  descriptor = ByteSpan{p.data, sd_size.value()};
  p.data += descriptor.size;
  p.size -= descriptor.size;
  title = std::move(name).value();
  return true;
}

/// Reads sqRfX and RangeProtectionTitleSDRel.
bool ReadRangeAndTitle(ByteSpan& p, std::vector<MergeRange>& ranges, std::string& title, ByteSpan& descriptor) {
  return ReadSqRfx(p, ranges, kMaxProtectedRangeRects) && ReadTitle(p, title, descriptor);
}

void EmitRects(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges) {
  emit_u32(dst, static_cast<std::uint32_t>(ranges.size()));
  for (const MergeRange& range : ranges) {
    emit_rfx(dst, range);
  }
}

void EmitTitle(std::vector<std::uint8_t>& dst, const std::string& title, const std::vector<std::uint8_t>& descriptor) {
  emit_xlwidestring(dst, title);
  emit_u32(dst, static_cast<std::uint32_t>(descriptor.size()));
  dst.insert(dst.end(), descriptor.begin(), descriptor.end());
}

void EmitRangeAndTitle(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges, const std::string& title,
                       const std::vector<std::uint8_t>& descriptor) {
  EmitRects(dst, ranges);
  EmitTitle(dst, title, descriptor);
}

/// Sets `out` to the SDDL text of `descriptor`; false when it does not decode.
bool ReadDescriptor(ByteSpan descriptor, std::string& out) {
  return descriptor.size == 0U || decode_security_descriptor(descriptor, out);
}

Error RangeNotWritable(const char* kind, std::size_t sheet_index) {
  return not_writable(kind, sheet_index);
}

RangeProtectionRead DecodeRangeProtection14(ByteSpan payload, ProtectedRange& out) {
  ByteSpan p = payload;
  auto flags = read_u32(p);
  auto sqrefs = read_u32(p);
  auto sqref_flags = read_u32(p);
  if (!flags || flags.value() != kFrtHeaderSqref || !sqrefs || sqrefs.value() != 1U || !sqref_flags ||
      sqref_flags.value() != kFrtSqrefDoAdjust) {
    return RangeProtectionRead::kUndecodable;
  }
  std::vector<MergeRange> ranges;
  std::string title;
  ByteSpan descriptor{};
  if (!ReadSqRfx(p, ranges, p.size / 16U)) {
    return RangeProtectionRead::kUndecodable;
  }
  auto pwd = read_u16(p);
  if (!pwd || !ReadTitle(p, title, descriptor) || p.size != 0U) {
    return RangeProtectionRead::kUndecodable;
  }
  ProtectedRange range;
  if (!ReadDescriptor(descriptor, range.security_descriptor)) {
    return RangeProtectionRead::kSecurityDescriptor;
  }
  if (pwd.value() != 0U) {
    range.password = LegacyHex(pwd.value());
  }
  range.name = std::move(title);
  range.sqref = format_sqref_text(ranges);
  out = std::move(range);
  return RangeProtectionRead::kDecoded;
}

/// A writable range's areas and legacy verifier.
struct RangeFields {
  std::vector<MergeRange> rects;
  std::uint16_t pwd = 0;
  std::vector<std::uint8_t> descriptor;
};

Expected<RangeFields, Error> CheckRange(const ProtectedRange& range, std::size_t sheet_index) {
  RangeFields fields;
  // Excel drops a descriptor naming an unknown SID alias.
  if (!range.security_descriptor.empty() &&
      encode_security_descriptor(range.security_descriptor, fields.descriptor) == SddlEncode::kUnsupported) {
    return RangeNotWritable("protected_range_descriptor", sheet_index);
  }
  const std::size_t units = utf16_units_in(range.name);
  if (units == 0U || units > kMaxProtectedRangeNameUnits) {
    return RangeNotWritable("protected_range_title", sheet_index);
  }
  if (!parse_sqref_text(range.sqref, fields.rects) || fields.rects.empty()) {
    return RangeNotWritable("protected_range_sqref", sheet_index);
  }
  // BrtRangeProtection14 has no Iso companion to carry a modern hash.
  if (fields.rects.size() > kMaxProtectedRangeRects && !range.hash_value.empty()) {
    return RangeNotWritable("protected_range_rects", sheet_index);
  }
  if (!range.password.empty() && !ParseLegacyHex(range.password, fields.pwd)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Protected range password is not a 16-bit hex hash",
                      kWriteXlsbContext);
  }
  return fields;
}

/// Where the BrtRangeProtection14 block goes; see `add_range_protections14`.
std::size_t RangeProtection14InsertionPoint(const std::vector<FramedRecord>& recs, std::size_t size) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kFrtBegin) {
      continue;
    }
    const std::uint16_t first = i + 1U < recs.size() ? recs[i + 1U].type : 0U;
    if (first != kBeginCfs14 && first != kBeginDVals14 && first != kBeginSparklineGroups) {
      return recs[i].begin;
    }
    while (i + 1U < recs.size() && recs[i].type != kFrtEnd) {
      ++i;
    }
  }
  if (recs.size() >= 3U && recs[recs.size() - 3U].type == kAcBegin && recs.back().type == kAcEnd) {
    return recs[recs.size() - 3U].begin;
  }
  return size;
}

}  // namespace

RangeProtectionRead decode_range_protection(ByteSpan payload, ByteSpan iso, ProtectedRange& out) {
  ByteSpan p = payload;
  auto pwd = read_u16(p);
  std::vector<MergeRange> ranges;
  std::string title;
  ByteSpan descriptor{};
  if (!pwd || !ReadRangeAndTitle(p, ranges, title, descriptor) || p.size != 0U) {
    return RangeProtectionRead::kUndecodable;
  }
  ProtectedRange range;
  if (iso.size != 0U) {
    ByteSpan q = iso;
    auto spin = read_u32(q);
    std::vector<MergeRange> iso_ranges;
    std::string iso_title;
    ByteSpan iso_descriptor{};
    std::vector<std::uint8_t> hash;
    std::vector<std::uint8_t> salt;
    if (pwd.value() != 0U || !spin || !ReadRangeAndTitle(q, iso_ranges, iso_title, iso_descriptor) ||
        iso_ranges != ranges || iso_title != title || !ReadBlob(q, hash) || hash.empty() || !ReadBlob(q, salt)) {
      return RangeProtectionRead::kUndecodable;
    }
    auto algorithm = read_xlnullablewidestring(q);
    if (!algorithm || q.size != 0U) {
      return RangeProtectionRead::kUndecodable;
    }
    if (iso_descriptor.size != descriptor.size ||
        !std::equal(descriptor.data, descriptor.data + descriptor.size, iso_descriptor.data)) {
      return RangeProtectionRead::kSecurityDescriptor;
    }
    range.spin_count = spin.value();
    range.hash_value = base64_encode(hash.data(), hash.size());
    range.salt_value = base64_encode(salt.data(), salt.size());
    range.algorithm_name = std::move(algorithm).value();
  } else if (pwd.value() != 0U) {
    range.password = LegacyHex(pwd.value());
  }
  if (!ReadDescriptor(descriptor, range.security_descriptor)) {
    return RangeProtectionRead::kSecurityDescriptor;
  }
  range.name = std::move(title);
  range.sqref = format_sqref_text(ranges);
  out = std::move(range);
  return RangeProtectionRead::kDecoded;
}

Expected<void, Error> emit_range_protections(std::vector<std::uint8_t>& dst, const std::vector<ProtectedRange>& ranges,
                                             std::size_t sheet_index) {
  for (const ProtectedRange& range : ranges) {
    auto fields = CheckRange(range, sheet_index);
    if (!fields) {
      return std::move(fields.error());
    }
    const std::vector<MergeRange>& rfx = fields.value().rects;
    if (rfx.size() > kMaxProtectedRangeRects) {
      continue;
    }
    std::uint16_t pwd = fields.value().pwd;
    std::vector<std::uint8_t> payload;
    if (!range.hash_value.empty()) {
      std::vector<std::uint8_t> hash;
      std::vector<std::uint8_t> salt;
      if (!base64_decode(range.hash_value, hash) || !base64_decode(range.salt_value, salt) ||
          hash.size() > kMaxBlobBytes || salt.size() > kMaxBlobBytes) {
        return make_error(FormulonErrorCode::kInvalidArgument, "Protected range hash or salt is not base64",
                          kWriteXlsbContext);
      }
      emit_u32(payload, range.spin_count);
      EmitRangeAndTitle(payload, rfx, range.name, fields.value().descriptor);
      emit_u32(payload, static_cast<std::uint32_t>(hash.size()));
      payload.insert(payload.end(), hash.begin(), hash.end());
      emit_u32(payload, static_cast<std::uint32_t>(salt.size()));
      payload.insert(payload.end(), salt.begin(), salt.end());
      emit_xlnullablewidestring(
          payload, range.algorithm_name.empty() ? std::nullopt : std::optional<std::string_view>(range.algorithm_name));
      emit_record(dst, kBrtRangeProtectionIso, payload);
      payload.clear();
      pwd = 0U;  // the Iso record's companion carries no legacy verifier
    }
    emit_u16(payload, pwd);
    EmitRangeAndTitle(payload, rfx, range.name, fields.value().descriptor);
    emit_record(dst, kBrtRangeProtection, payload);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> add_range_protections14(std::vector<std::uint8_t>& records,
                                              const std::vector<ProtectedRange>& ranges, std::size_t sheet_index) {
  std::vector<std::uint8_t> body;
  for (const ProtectedRange& range : ranges) {
    auto fields = CheckRange(range, sheet_index);
    if (!fields) {
      return std::move(fields.error());
    }
    if (fields.value().rects.size() <= kMaxProtectedRangeRects) {
      continue;
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, kFrtHeaderSqref);
    emit_u32(payload, 1U);
    emit_u32(payload, kFrtSqrefDoAdjust);
    EmitRects(payload, fields.value().rects);
    emit_u16(payload, fields.value().pwd);
    EmitTitle(payload, range.name, fields.value().descriptor);
    emit_record(body, kBrtRangeProtection14, payload);
  }
  if (body.empty()) {
    return Expected<void, Error>::Ok();
  }
  const std::vector<FramedRecord> recs = split_records(records);
  std::vector<std::uint8_t> block;
  // Excel refuses a future-record block directly after the last legacy one.
  const bool has_print = std::any_of(recs.begin(), recs.end(), [](const FramedRecord& r) {
    return r.type == kBrtPrintOptions || r.type == kBrtMargins || r.type == kBrtPageSetup;
  });
  if (!has_print) {
    const std::array<std::uint8_t, 2> default_print_options = {0x10, 0x00};
    emit_record(block, kBrtPrintOptions, ByteSpan{default_print_options.data(), default_print_options.size()});
  }
  emit_frt_begin(block, kFrtVersionWord);
  block.insert(block.end(), body.begin(), body.end());
  emit_record(block, kFrtEnd, ByteSpan{});
  const std::size_t at = RangeProtection14InsertionPoint(recs, records.size());
  records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
  return Expected<void, Error>::Ok();
}

void take_range_protections14(std::vector<std::uint8_t>& records, std::vector<ProtectedRange>& out) {
  const std::vector<FramedRecord> recs = split_records(records);
  std::vector<std::pair<std::size_t, std::size_t>> cuts;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kFrtBegin || i + 1U >= recs.size() || recs[i + 1U].type != kBrtRangeProtection14) {
      continue;
    }
    std::vector<ProtectedRange> block;
    std::size_t j = i + 1U;
    bool decoded = true;
    for (; j < recs.size() && recs[j].type == kBrtRangeProtection14; ++j) {
      ProtectedRange range;
      decoded = decoded && DecodeRangeProtection14(recs[j].payload, range) == RangeProtectionRead::kDecoded;
      block.push_back(std::move(range));
    }
    if (decoded && j < recs.size() && recs[j].type == kFrtEnd) {
      cuts.emplace_back(recs[i].begin, recs[j].end);
      for (ProtectedRange& range : block) {
        out.push_back(std::move(range));
      }
    }
    i = j;
  }
  for (auto it = cuts.rbegin(); it != cuts.rend(); ++it) {
    records.erase(records.begin() + static_cast<std::ptrdiff_t>(it->first),
                  records.begin() + static_cast<std::ptrdiff_t>(it->second));
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
