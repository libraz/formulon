// @size-budget: 56 KB

#include "io/xlsb/protection_records.h"

#include <array>
#include <optional>
#include <string_view>
#include <utility>

#include "io/protection_xml.h"
#include "io/sqref_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"
#include "utils/number_text.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::size_t kFlagCount = 16;
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
    const int d = (c >= '0' && c <= '9')   ? c - '0'
                  : (c >= 'A' && c <= 'F') ? c - 'A' + 10
                  : (c >= 'a' && c <= 'f') ? c - 'a' + 10
                                           : -1;
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
                      "context=write_xlsb");
  }
  std::vector<std::uint8_t> payload;
  if (!protection.algorithm_name.empty()) {
    std::vector<std::uint8_t> hash;
    std::vector<std::uint8_t> salt;
    if (!base64_decode(protection.hash_value, hash) || !base64_decode(protection.salt_value, salt) ||
        hash.size() > kMaxBlobBytes || salt.size() > kMaxBlobBytes) {
      return make_error(FormulonErrorCode::kInvalidArgument, "sheet protection hash or salt is not base64",
                        "context=write_xlsb");
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
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook protection is not a <workbookProtection> element",
                      "context=write_xlsb");
  }
  const WorkbookProtection prot = parse_workbook_protection(xml);
  std::uint16_t book_pwd = 0;
  std::uint16_t rev_pwd = 0;
  if ((!prot.workbook_password.empty() && !ParseLegacyHex(prot.workbook_password, book_pwd)) ||
      (!prot.revisions_password.empty() && !ParseLegacyHex(prot.revisions_password, rev_pwd))) {
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook password is not a 16-bit hex hash",
                      "context=write_xlsb");
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
      return make_error(FormulonErrorCode::kInvalidArgument, "workbook protection hash or salt is not base64",
                        "context=write_xlsb");
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

// UncheckedSqRfX of BrtRangeProtection: below 8192 ranges, the rest need the
// BrtRangeProtection14 future record.
constexpr std::uint32_t kMaxRangeProtectionRfx = 8191U;
constexpr std::size_t kMaxRangeTitleChars = 255U;

bool ReadSqRfx(ByteSpan& p, std::vector<MergeRange>& out) {
  auto count = read_u32(p);
  if (!count || count.value() == 0U || count.value() > kMaxRangeProtectionRfx || p.size / 16U < count.value()) {
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

/// Reads sqRfX and RangeProtectionTitleSDRel; false when a security
/// descriptor is present.
bool ReadRangeAndTitle(ByteSpan& p, std::vector<MergeRange>& ranges, std::string& title) {
  if (!ReadSqRfx(p, ranges)) {
    return false;
  }
  auto name = read_xlwidestring(p);
  auto sd_size = read_u32(p);
  if (!name || name.value().empty() || !sd_size || sd_size.value() != 0U) {
    return false;
  }
  title = std::move(name).value();
  return true;
}

void EmitRangeAndTitle(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges,
                       const std::string& title) {
  emit_u32(dst, static_cast<std::uint32_t>(ranges.size()));
  for (const MergeRange& range : ranges) {
    emit_rfx(dst, range);
  }
  emit_xlwidestring(dst, title);
  emit_u32(dst, 0U);  // dwsdRel: no security descriptor
}

/// UTF-16 code units of UTF-8 `text`: a four-byte sequence is a surrogate pair.
std::size_t Utf16UnitCount(const std::string& text) {
  std::size_t count = 0;
  for (const char c : text) {
    const auto byte = static_cast<unsigned char>(c);
    if ((byte & 0xC0U) != 0x80U) {
      count += byte >= 0xF0U ? 2U : 1U;
    }
  }
  return count;
}

}  // namespace

bool decode_range_protection(ByteSpan payload, ByteSpan iso, ProtectedRange& out) {
  ByteSpan p = payload;
  auto pwd = read_u16(p);
  std::vector<MergeRange> ranges;
  std::string title;
  if (!pwd || !ReadRangeAndTitle(p, ranges, title) || p.size != 0U) {
    return false;
  }
  ProtectedRange range;
  if (iso.size != 0U) {
    ByteSpan q = iso;
    auto spin = read_u32(q);
    std::vector<MergeRange> iso_ranges;
    std::string iso_title;
    std::vector<std::uint8_t> hash;
    std::vector<std::uint8_t> salt;
    if (pwd.value() != 0U || !spin || !ReadRangeAndTitle(q, iso_ranges, iso_title) || iso_ranges != ranges ||
        iso_title != title || !ReadBlob(q, hash) || hash.empty() || !ReadBlob(q, salt)) {
      return false;
    }
    auto algorithm = read_xlnullablewidestring(q);
    if (!algorithm || q.size != 0U) {
      return false;
    }
    range.spin_count = spin.value();
    range.hash_value = base64_encode(hash.data(), hash.size());
    range.salt_value = base64_encode(salt.data(), salt.size());
    range.algorithm_name = std::move(algorithm).value();
  } else if (pwd.value() != 0U) {
    range.password = LegacyHex(pwd.value());
  }
  range.name = std::move(title);
  range.sqref = format_sqref_text(ranges);
  out = std::move(range);
  return true;
}

bool range_protection_writable(const ProtectedRange& range) {
  std::vector<MergeRange> ranges;
  const std::size_t chars = Utf16UnitCount(range.name);
  return range.security_descriptor.empty() && chars != 0U && chars <= kMaxRangeTitleChars &&
         parse_sqref_text(range.sqref, ranges) && ranges.size() <= kMaxRangeProtectionRfx;
}

Expected<void, Error> emit_range_protections(std::vector<std::uint8_t>& dst,
                                             const std::vector<ProtectedRange>& ranges) {
  for (const ProtectedRange& range : ranges) {
    if (!range_protection_writable(range)) {
      continue;
    }
    std::vector<MergeRange> rfx;
    (void)parse_sqref_text(range.sqref, rfx);
    std::uint16_t pwd = 0;
    if (!range.password.empty() && !ParseLegacyHex(range.password, pwd)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "protected range password is not a 16-bit hex hash",
                        "context=write_xlsb");
    }
    std::vector<std::uint8_t> payload;
    if (!range.hash_value.empty()) {
      std::vector<std::uint8_t> hash;
      std::vector<std::uint8_t> salt;
      if (!base64_decode(range.hash_value, hash) || !base64_decode(range.salt_value, salt) ||
          hash.size() > kMaxBlobBytes || salt.size() > kMaxBlobBytes) {
        return make_error(FormulonErrorCode::kInvalidArgument, "protected range hash or salt is not base64",
                          "context=write_xlsb");
      }
      emit_u32(payload, range.spin_count);
      EmitRangeAndTitle(payload, rfx, range.name);
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
    EmitRangeAndTitle(payload, rfx, range.name);
    emit_record(dst, kBrtRangeProtection, payload);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
