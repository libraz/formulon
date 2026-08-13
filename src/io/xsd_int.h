//
// Locale-independent parser for the XSD `unsignedInt` /
// `nonNegativeInteger` lexical values OOXML uses for index-shaped
// attributes — `<c s="N">`, `<f si="N">`, `<row r="N">` and friends.
//
// One definition serves the DOM sheet reader, the streaming SAX sheet
// reader, and the cell parser. The OOXML reader picks between the DOM
// and the SAX path purely by sheet size (`kSaxThresholdBytes`), so any
// difference in how the two lex a numeric attribute would make the same
// bytes decode differently depending on how big the sheet happened to
// be. Sharing the lexer is what keeps that dispatch an implementation
// detail.

#ifndef FORMULON_IO_XSD_INT_H_
#define FORMULON_IO_XSD_INT_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace io {

/// Parses `text` as a non-negative decimal integer.
///
/// The accepted lexical space is exactly a non-empty run of ASCII digits:
/// no sign, no surrounding whitespace, no trailing characters, and no
/// value past `UINT32_MAX`. Excel emits none of the tolerated-by-strtoul
/// variants for these attributes, so reading strictly loses nothing real
/// and keeps a malformed attribute from silently becoming a plausible
/// index.
///
/// Returns false and leaves `*out` unchanged on anything outside that
/// space. Each caller decides the disposition of a rejected value — the
/// schema default for a cosmetic attribute, a hard error for an index
/// that binds one cell to another — but both read paths must make the
/// same decision for the same attribute.
inline bool parse_xsd_nonneg_int(std::string_view text, std::uint32_t* out) noexcept {
  if (text.empty()) {
    return false;
  }
  std::uint64_t acc = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return false;
    }
    acc = (acc * 10U) + static_cast<std::uint64_t>(c - '0');
    if (acc > 0xFFFFFFFFULL) {
      return false;
    }
  }
  *out = static_cast<std::uint32_t>(acc);
  return true;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XSD_INT_H_
