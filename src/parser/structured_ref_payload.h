//
// The bracket payload of a structured reference (`Table[...]`): its model,
// the spelling Excel stores, and the spelling the formula bar shows.
//
// Storage is always the long form (`[#This Row],[Qty]`), with `'` escaping
// `' [ ] #` only. The display form depends on the formula's cell: `#This Row`
// is spelled `@`, `@` is escaped too, and inside the referenced table's own
// range a single-column reference with no specifier or `@` alone drops the
// table name. An unbracketed payload is one column name, `,` and `:` included.
//
// @size-budget: 6 KB

#ifndef FORMULON_PARSER_STRUCTURED_REF_PAYLOAD_H_
#define FORMULON_PARSER_STRUCTURED_REF_PAYLOAD_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {
namespace parser {

/// Bitmask of the area specifiers a structured reference carries.
struct StructuredRefSpecifiers {
  static constexpr std::uint8_t kAll = 1u << 0u;
  static constexpr std::uint8_t kData = 1u << 1u;
  static constexpr std::uint8_t kHeaders = 1u << 2u;
  static constexpr std::uint8_t kTotals = 1u << 3u;
  static constexpr std::uint8_t kThisRow = 1u << 4u;

  /// The spellings of the specifiers, as Excel stores them.
  static constexpr std::string_view kAllWord = "#All";
  static constexpr std::string_view kHeadersWord = "#Headers";
  static constexpr std::string_view kDataWord = "#Data";
  static constexpr std::string_view kTotalsWord = "#Totals";
  static constexpr std::string_view kThisRowWord = "#This Row";
};

/// A parsed bracket payload.
struct StructuredRefPayload {
  /// OR of `StructuredRefSpecifiers`; 0 when none is written (Excel reads
  /// that as data rows).
  std::uint8_t specifiers = 0;
  /// Unescaped column names: none (every column), one, or a range's two ends.
  std::vector<std::string> columns;
};

/// Where a formula is displayed, for the table-name drop rule.
struct StructuredRefDisplay {
  /// Display name of the table whose range holds the formula's cell; empty
  /// outside every table.
  std::string_view own_table;
};

/// Parses the text between a structured reference's outer brackets.
///
/// Accepted: empty; `@` alone; `@X`, `@[X]`, `@[A]:[B]`; one `#Spec`; a list
/// of bracketed items `[#Spec],[X]` / `[#Spec],[A]:[B]` with at most one
/// column item or range; any other text as one column name. `'` makes the
/// next character literal. Specifier names match ASCII case-insensitively.
/// Returns nullopt on an unknown specifier or malformed brackets.
std::optional<StructuredRefPayload> parse_structured_ref_payload_text(std::string_view payload);

/// The payload as Excel stores it.
std::string stored_payload(const StructuredRefPayload& payload);

/// `table[...]` as the formula bar shows it, `table` dropped where Excel
/// drops it when `inside_own_table`.
std::string display_structured_ref(std::string_view table, const StructuredRefPayload& payload, bool inside_own_table);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_STRUCTURED_REF_PAYLOAD_H_
