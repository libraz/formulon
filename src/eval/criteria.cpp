// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the criterion parser and matcher declared in
// `criteria.h`. See the header Doxygen for the authoritative contract.

#include "eval/criteria.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Attempts to strip a leading comparator prefix from `text` and writes the
// detected operator into `*out_op`. On success returns the remainder of
// the string; on no-match returns `text` unchanged and leaves `*out_op`
// set to the default (`Eq`). Longer prefixes are matched first so "<=" is
// not misread as "<" followed by "=".
std::string_view strip_comparator(std::string_view text, CriteriaOp* out_op) {
  *out_op = CriteriaOp::Eq;
  if (text.size() >= 2) {
    if (text[0] == '<' && text[1] == '>') {
      *out_op = CriteriaOp::NotEq;
      return text.substr(2);
    }
    if (text[0] == '<' && text[1] == '=') {
      *out_op = CriteriaOp::LtEq;
      return text.substr(2);
    }
    if (text[0] == '>' && text[1] == '=') {
      *out_op = CriteriaOp::GtEq;
      return text.substr(2);
    }
  }
  if (!text.empty()) {
    if (text[0] == '<') {
      *out_op = CriteriaOp::Lt;
      return text.substr(1);
    }
    if (text[0] == '>') {
      *out_op = CriteriaOp::Gt;
      return text.substr(1);
    }
    if (text[0] == '=') {
      // Explicit '=' — same semantics as no prefix, but consume the byte so
      // the RHS probe sees just the value part.
      *out_op = CriteriaOp::Eq;
      return text.substr(1);
    }
  }
  return text;
}

// Attempts to parse `rhs` as a standalone number by wrapping it in a
// temporary `Value::text` and delegating to `coerce_to_number`. Matches
// Excel's behaviour where ">5" and ">5.0" and "> 5 " all parse as numeric.
// Returns true + writes to `*out_number` on success; false otherwise.
bool probe_number(std::string_view rhs, double* out_number) {
  const Value probe = Value::text(rhs);
  const auto result = coerce_to_number(probe);
  if (!result) {
    return false;
  }
  *out_number = result.value();
  return true;
}

// Unescapes a wildcard-free pattern: `~X` becomes `X` for any `X`. Used for
// literal equality comparison when the criterion text contains only escape
// sequences (no real wildcards).
std::string unescape_literal(std::string_view pattern) {
  std::string out;
  out.reserve(pattern.size());
  for (std::size_t i = 0; i < pattern.size(); ++i) {
    if (pattern[i] == '~' && i + 1 < pattern.size()) {
      out.push_back(pattern[i + 1]);
      ++i;
      continue;
    }
    out.push_back(pattern[i]);
  }
  return out;
}

}  // namespace

bool scan_has_wildcard(std::string_view rhs) {
  for (std::size_t i = 0; i < rhs.size(); ++i) {
    const char c = rhs[i];
    if (c == '~' && i + 1 < rhs.size()) {
      ++i;  // Skip the escaped byte.
      continue;
    }
    if (c == '*' || c == '?') {
      return true;
    }
  }
  return false;
}

namespace {

// Core two-pointer wildcard matcher. When `prefix_match_ok` is false the
// pattern must consume `text` exactly (SEARCH whole-string / criteria
// equality semantics). When `prefix_match_ok` is true the match succeeds
// as soon as the pattern is exhausted, even if `text` has unmatched bytes
// remaining — used by `wildcard_find` to answer "does the pattern match a
// prefix of this suffix?". When the match succeeds in prefix mode,
// `*out_consumed` receives the number of `text` bytes the pattern
// consumed.
bool wildcard_match_impl(std::string_view pattern, std::string_view text, bool prefix_match_ok,
                         std::size_t* out_consumed) {
  std::size_t pi = 0;
  std::size_t ti = 0;
  std::size_t star_pi = std::string_view::npos;
  std::size_t star_ti = 0;
  while (ti < text.size()) {
    // Prefix mode: succeed as soon as the pattern is fully consumed, even
    // if there is unmatched text remaining.
    if (prefix_match_ok && pi == pattern.size()) {
      if (out_consumed != nullptr) {
        *out_consumed = ti;
      }
      return true;
    }
    if (pi < pattern.size()) {
      const char pc = pattern[pi];
      if (pc == '~' && pi + 1 < pattern.size()) {
        // Escaped literal: match the next pattern byte exactly.
        if (pattern[pi + 1] == text[ti]) {
          pi += 2;
          ++ti;
          continue;
        }
      } else if (pc == '*') {
        star_pi = pi;
        star_ti = ti;
        ++pi;
        continue;
      } else if (pc == '?' || pc == text[ti]) {
        ++pi;
        ++ti;
        continue;
      }
    }
    if (star_pi != std::string_view::npos) {
      // Backtrack: extend the most-recent `*` by one byte of text.
      pi = star_pi + 1;
      ++star_ti;
      ti = star_ti;
      continue;
    }
    return false;
  }
  // Trailing `*`s in the pattern can consume nothing.
  while (pi < pattern.size() && pattern[pi] == '*') {
    ++pi;
  }
  const bool ok = (pi == pattern.size());
  if (ok && out_consumed != nullptr) {
    *out_consumed = ti;
  }
  return ok;
}

}  // namespace

bool wildcard_match(std::string_view pattern, std::string_view text) {
  return wildcard_match_impl(pattern, text, /*prefix_match_ok=*/false, /*out_consumed=*/nullptr);
}

std::size_t wildcard_find(std::string_view pattern, std::string_view text) {
  // Try each start offset in `text` and attempt a prefix match. Because `*`
  // already permits zero-or-more bytes, a leading `*` in the pattern makes
  // offset 0 always match — but the loop below is still correct in that
  // case (it would find offset 0 on the first iteration).
  for (std::size_t start = 0; start <= text.size(); ++start) {
    if (wildcard_match_impl(pattern, text.substr(start), /*prefix_match_ok=*/true,
                            /*out_consumed=*/nullptr)) {
      return start;
    }
  }
  return std::string_view::npos;
}

ParsedCriterion parse_criterion(const Value& criterion) {
  ParsedCriterion parsed;
  switch (criterion.kind()) {
    case ValueKind::Number: {
      parsed.op = CriteriaOp::Eq;
      parsed.rhs_is_number = true;
      parsed.rhs_number = criterion.as_number();
      return parsed;
    }
    case ValueKind::Bool: {
      parsed.op = CriteriaOp::Eq;
      parsed.rhs_is_number = true;
      parsed.rhs_number = criterion.as_boolean() ? 1.0 : 0.0;
      return parsed;
    }
    case ValueKind::Blank: {
      parsed.op = CriteriaOp::Eq;
      parsed.rhs_is_number = false;
      parsed.rhs_text = std::string_view{};
      return parsed;
    }
    case ValueKind::Text: {
      const std::string_view raw = criterion.as_text();
      CriteriaOp op = CriteriaOp::Eq;
      const std::string_view rhs_view = strip_comparator(raw, &op);
      parsed.op = op;
      const bool prefix_stripped = rhs_view.size() != raw.size();
      // Probe for numeric RHS first — Excel treats ">5" as numeric even
      // though the RHS is textual at this layer.
      double num = 0.0;
      if (probe_number(rhs_view, &num)) {
        // Blank probe ("") coerces to 0.0 under `coerce_to_number`; treat
        // that as "no RHS" when the user wrote bare "<>" or "=" with
        // nothing after, matching Excel's behaviour where "<>" is the
        // non-blank probe rather than "<> 0".
        if (rhs_view.empty()) {
          parsed.rhs_is_number = false;
          parsed.rhs_text = std::string_view{};
          return parsed;
        }
        parsed.rhs_is_number = true;
        parsed.rhs_number = num;
        return parsed;
      }
      // Text RHS. Store a detached copy when we stripped a prefix so the
      // returned view does not alias a smaller substring of the caller's
      // Text payload in subtle ways; when no prefix was stripped we can
      // reuse the original view directly.
      parsed.rhs_is_number = false;
      if (prefix_stripped) {
        parsed.rhs_storage.assign(rhs_view);
        parsed.rhs_text = parsed.rhs_storage;
      } else {
        parsed.rhs_text = rhs_view;
      }
      parsed.has_wildcard =
          (parsed.op == CriteriaOp::Eq || parsed.op == CriteriaOp::NotEq) && scan_has_wildcard(parsed.rhs_text);
      return parsed;
    }
    case ValueKind::Error:
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      // Contract: callers reject errors before calling this helper. For the
      // remaining non-scalar kinds there is no Excel-visible criterion
      // shape today; fall back to `Op::Eq` / text "" which matches nothing
      // interesting and keeps the function total.
      parsed.op = CriteriaOp::Eq;
      parsed.rhs_is_number = false;
      parsed.rhs_text = std::string_view{};
      return parsed;
  }
  return parsed;
}

namespace {

// Applies the ordering comparator `op` given an integer sign (`-1`, `0`,
// `1`) from a compare function. Only the four ordering ops are valid here.
bool apply_ordering(CriteriaOp op, int cmp) {
  switch (op) {
    case CriteriaOp::Lt:
      return cmp < 0;
    case CriteriaOp::LtEq:
      return cmp <= 0;
    case CriteriaOp::Gt:
      return cmp > 0;
    case CriteriaOp::GtEq:
      return cmp >= 0;
    default:
      return false;
  }
}

bool matches_text(const Value& cell, const ParsedCriterion& c) {
  // Project the cell to its text rendering. `coerce_to_text` yields the
  // Excel-visible spelling for numbers / bools and passes Text through; a
  // failure (unsupported kind) falls back to non-match. Blank becomes "".
  const auto cell_text_exp = coerce_to_text(cell);
  if (!cell_text_exp) {
    return false;
  }
  const std::string& cell_text = cell_text_exp.value();
  const std::string_view rhs = c.rhs_text;

  switch (c.op) {
    case CriteriaOp::Eq: {
      if (c.has_wildcard) {
        return wildcard_match(rhs, cell_text);
      }
      const std::string literal = unescape_literal(rhs);
      return strings::case_insensitive_eq(cell_text, literal);
    }
    case CriteriaOp::NotEq: {
      if (c.has_wildcard) {
        return !wildcard_match(rhs, cell_text);
      }
      const std::string literal = unescape_literal(rhs);
      return !strings::case_insensitive_eq(cell_text, literal);
    }
    case CriteriaOp::Lt:
    case CriteriaOp::LtEq:
    case CriteriaOp::Gt:
    case CriteriaOp::GtEq: {
      // Wildcards in ordering ops are literal — Excel does not interpret
      // `*`/`?` for comparison. Pass the raw RHS view through the
      // case-insensitive compare.
      const int cmp = strings::case_insensitive_compare(cell_text, rhs);
      return apply_ordering(c.op, cmp);
    }
  }
  return false;
}

bool matches_numeric(const Value& cell, const ParsedCriterion& c) {
  const auto num_exp = coerce_to_number(cell);
  if (!num_exp) {
    // Non-numeric cells simply do not match a numeric criterion. Errors
    // from non-parseable text here are NOT propagated — the caller has
    // already decided that criterion-driven walks swallow per-cell errors.
    return false;
  }
  const double x = num_exp.value();
  const double y = c.rhs_number;
  switch (c.op) {
    case CriteriaOp::Eq:
      return x == y;
    case CriteriaOp::NotEq:
      return x != y;
    case CriteriaOp::Lt:
      return x < y;
    case CriteriaOp::LtEq:
      return x <= y;
    case CriteriaOp::Gt:
      return x > y;
    case CriteriaOp::GtEq:
      return x >= y;
  }
  return false;
}

}  // namespace

bool matches_criterion(const Value& cell, const ParsedCriterion& c) {
  // Errors in the criteria range are silently skipped. This matches Excel
  // and diverges intentionally from the eager dispatcher's
  // `propagate_errors` default — criteria walks never fail on per-cell
  // errors, they just do not count them.
  if (cell.is_error()) {
    return false;
  }

  // Sentinel "non-blank" filter: bare `"<>"` (NotEq with empty text RHS)
  // matches any cell that is NOT blank, including literal empty-string
  // Text values ("is this cell populated?"). Evaluated before the
  // generic blank-cell branch so the asymmetry with `=""` below is
  // visible.
  if (!c.rhs_is_number && c.op == CriteriaOp::NotEq && c.rhs_text.empty()) {
    return !cell.is_blank();
  }

  // Blank-cell special cases. Excel's rule: `=""` matches blanks (see
  // also the sentinel branch above for the mirror case). For every other
  // op against a blank cell we need to report the Excel-observed outcome
  // of comparing a "missing value" against a concrete RHS:
  //   * NotEq with a concrete RHS (numeric or non-empty text) -> true:
  //     blank is not equal to that concrete value.
  //   * Every other op -> false (blanks have no usable numeric/text
  //     projection for ordering).
  if (cell.is_blank()) {
    if (c.op == CriteriaOp::Eq && !c.rhs_is_number && c.rhs_text.empty()) {
      return true;
    }
    if (c.op == CriteriaOp::NotEq) {
      // Sentinel (NotEq empty text) already handled above; this path
      // therefore has a concrete numeric or non-empty text RHS.
      return true;
    }
    return false;
  }

  if (c.rhs_is_number) {
    return matches_numeric(cell, c);
  }
  return matches_text(cell, c);
}

}  // namespace eval
}  // namespace formulon
