// @size-budget: 56 KB

#include "io/xlsb/protection_records.h"

#include <array>
#include <cstdio>
#include <string_view>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_escape.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"

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
  std::snprintf(buf, sizeof(buf), "%04X", static_cast<unsigned>(pwd));
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

void AppendAttr(std::string& xml, std::string_view name, std::string_view value) {
  xml.push_back(' ');
  xml.append(name);
  xml.append("=\"");
  AppendXmlAttrEscaped(xml, value);
  xml.push_back('"');
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

bool decode_book_protection(ByteSpan payload, ByteSpan iso, std::string& xml) {
  ByteSpan p = payload;
  auto pwd = read_u16(p);
  auto reserved = read_u16(p);
  auto flags = read_u16(p);
  if (!pwd || !reserved || !flags || p.size != 0U || reserved.value() != 0U || (flags.value() & ~1U) != 0U) {
    return false;
  }
  std::vector<std::uint8_t> hash;
  std::vector<std::uint8_t> salt;
  std::string algorithm;
  std::uint32_t spin = 0;
  if (iso.size != 0U) {
    ByteSpan q = iso;
    auto book_spin = read_u32(q);
    auto rev_spin = read_u32(q);
    auto iso_flags = read_u16(q);
    if (!book_spin || !rev_spin || !iso_flags || rev_spin.value() != 0U || iso_flags.value() != flags.value() ||
        !ReadBlob(q, hash) || !ReadBlob(q, salt)) {
      return false;
    }
    auto name = read_xlwidestring(q);
    std::vector<std::uint8_t> rev_hash;
    std::vector<std::uint8_t> rev_salt;
    if (!name || !ReadBlob(q, rev_hash) || !ReadBlob(q, rev_salt) || !rev_hash.empty() || !rev_salt.empty()) {
      return false;
    }
    auto rev_name = read_u32(q);
    if (!rev_name || rev_name.value() != 0xFFFFFFFFU || q.size != 0U) {
      return false;
    }
    spin = book_spin.value();
    algorithm = std::move(name).value();
  }
  xml.clear();
  if (flags.value() == 0U && pwd.value() == 0U && algorithm.empty()) {
    return true;
  }
  xml = "<workbookProtection";
  if (pwd.value() != 0U) {
    AppendAttr(xml, "workbookPassword", LegacyHex(pwd.value()));
  }
  if (!algorithm.empty()) {
    AppendAttr(xml, "workbookAlgorithmName", algorithm);
    AppendAttr(xml, "workbookHashValue", base64_encode(hash.data(), hash.size()));
    AppendAttr(xml, "workbookSaltValue", base64_encode(salt.data(), salt.size()));
    AppendAttr(xml, "workbookSpinCount", std::to_string(spin));
  }
  if ((flags.value() & 1U) != 0U) {
    AppendAttr(xml, "lockStructure", "1");
  }
  xml.append("/>");
  return true;
}

Expected<bool, Error> emit_book_protection(std::vector<std::uint8_t>& dst, const std::string& xml) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node node = doc.load_string(xml.c_str()) ? doc.child("workbookProtection") : pugi::xml_node();
  if (!node) {
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook protection is not a <workbookProtection> element",
                      "context=write_xlsb");
  }
  std::uint16_t pwd = 0;
  const std::string legacy = node.attribute("workbookPassword").value();
  if (!legacy.empty() && !ParseLegacyHex(legacy, pwd)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook password is not a 16-bit hex hash",
                      "context=write_xlsb");
  }
  const std::uint16_t flags = read_xsd_bool(node, "lockStructure", false) ? 1U : 0U;
  const std::string algorithm = node.attribute("workbookAlgorithmName").value();
  if (!algorithm.empty()) {
    std::vector<std::uint8_t> hash;
    std::vector<std::uint8_t> salt;
    if (!base64_decode(node.attribute("workbookHashValue").value(), hash) ||
        !base64_decode(node.attribute("workbookSaltValue").value(), salt) || hash.size() > kMaxBlobBytes ||
        salt.size() > kMaxBlobBytes) {
      return make_error(FormulonErrorCode::kInvalidArgument, "workbook protection hash or salt is not base64",
                        "context=write_xlsb");
    }
    std::vector<std::uint8_t> iso;
    emit_u32(iso, node.attribute("workbookSpinCount").as_uint(0));
    emit_u32(iso, 0U);
    emit_u16(iso, flags);
    emit_u32(iso, static_cast<std::uint32_t>(hash.size()));
    iso.insert(iso.end(), hash.begin(), hash.end());
    emit_u32(iso, static_cast<std::uint32_t>(salt.size()));
    iso.insert(iso.end(), salt.begin(), salt.end());
    emit_xlwidestring(iso, algorithm);
    emit_u32(iso, 0U);
    emit_u32(iso, 0U);
    emit_xlnullablewidestring(iso, std::nullopt);
    emit_record(dst, kBrtBookProtectionIso, iso);
  }
  std::vector<std::uint8_t> payload;
  emit_u16(payload, pwd);
  emit_u16(payload, 0U);
  emit_u16(payload, flags);
  emit_record(dst, kBrtBookProtection, payload);
  return !(read_xsd_bool(node, "lockWindows", false) || read_xsd_bool(node, "lockRevision", false) ||
           node.attribute("revisionsPassword") || node.attribute("revisionsAlgorithmName"));
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
