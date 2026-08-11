//
// Per-token scanner for the Excel TEXT() format-string engine. Walks the
// input format byte-by-byte and produces the `Section::tokens` stream
// consumed by `classify()` (in `number_format_section.cpp`) and the
// renderers (`number_format_render.cpp`, `render_numeric.cpp`,
// `render_date.cpp`, `render_fraction.cpp`). The public entry point
// `apply_format` lives in `number_format.cpp`.
//
// Stateless scanning helpers (color/condition/DBNum specifiers, run scan,
// date-letter detection) are factored out into
// `number_format_scanner.{h,cpp}` so this TU only carries the main
// dispatch loop.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

void tokenize_section(std::string_view fmt, Section& out) {
  std::vector<Token>& toks = out.tokens;
  auto push_literal = [&](std::size_t b, std::size_t e) {
    if (b == e) {
      return;
    }
    Token t;
    t.kind = Tok::Literal;
    t.lit_begin = b;
    t.lit_end = e;
    toks.push_back(t);
  };

  bool saw_color = false;
  std::size_t i = 0;
  while (i < fmt.size()) {
    const char c = fmt[i];
    // Quoted literal `"..."`.
    if (c == '"') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != '"') {
        ++j;
      }
      push_literal(i + 1, j);
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    // Escape `\x` or `!x` -> next UTF-8 scalar is a literal.
    if ((c == '\\' || c == '!') && i + 1 < fmt.size()) {
      const std::size_t payload_width = utf8_scalar_width(fmt, i + 1);
      push_literal(i + 1, i + 1 + payload_width);
      i += 1 + payload_width;
      continue;
    }
    // Bracketed specifier. Recognised kinds:
    //   `[h]` / `[m]` / `[s]`   -> elapsed-time tokens (any run length).
    //   `[$...]`                -> locale-currency marker; silently dropped.
    //   `[赤]` / `[青]` / ...    -> named color qualifier; silently dropped
    //                               (Excel ignores color inside TEXT).
    //   `[色N]` (N in 1..56)     -> indexed color qualifier; silently dropped.
    // Anything else (`[>100]`, `[DBNum1]`, unknown qualifiers) still trips
    // the invalid-bracket flag and surfaces as #VALUE!.
    if (c == '[') {
      const std::size_t body_begin = i + 1;
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      const std::string_view body = fmt.substr(body_begin, (j < fmt.size() ? j : fmt.size()) - body_begin);
      i = j < fmt.size() ? j + 1 : j;
      // Elapsed time markers: any run of `h`, `m`, or `s` (case-insensitive).
      bool all_h = !body.empty();
      bool all_m = !body.empty();
      bool all_s = !body.empty();
      for (char ch : body) {
        const char lo = (ch >= 'A' && ch <= 'Z') ? static_cast<char>(ch + 32) : ch;
        if (lo != 'h') {
          all_h = false;
        }
        if (lo != 'm') {
          all_m = false;
        }
        if (lo != 's') {
          all_s = false;
        }
      }
      if (all_h) {
        Token t;
        t.kind = Tok::DateElapsedH;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else if (all_m) {
        Token t;
        t.kind = Tok::DateElapsedM;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else if (!body.empty() && body.front() == '$') {
        // Locale-currency marker. Form: `[$<symbol>-<lcid>]` or `[$<symbol>]`
        // or `[$-<lcid>]`. The `<symbol>` portion (everything after `$` up
        // to the optional `-LCID` suffix) is emitted as a literal prefix;
        // the LCID is metadata only and is discarded. Mac Excel 365 emits
        // e.g. `[$JPY-411]#,##0` -> `JPY1,234,567`.
        //
        // `body_begin` points at the `$` in `fmt`; the symbol bytes start
        // at `body_begin + 1` and run until either `-` or the end of body.
        const std::size_t sym_begin = body_begin + 1;
        std::size_t sym_end = sym_begin + (body.size() - 1);
        for (std::size_t k = 1; k < body.size(); ++k) {
          if (body[k] == '-') {
            sym_end = body_begin + k;
            break;
          }
        }
        push_literal(sym_begin, sym_end);
      } else if (all_s) {
        Token t;
        t.kind = Tok::DateElapsedS;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else if (is_color_specifier(body)) {
        // Named colour (`[赤]`) or indexed colour (`[色12]`). Excel discards
        // the colour inside TEXT, so the rest of the section still formats
        // the value. A section may carry at most one colour, though: Excel
        // rejects `[赤][青]0.00` with #VALUE! even though either bracket
        // alone is inert.
        if (saw_color) {
          out.has_invalid_bracket = true;
        }
        saw_color = true;
      } else if (const int dbnum = parse_dbnum_directive(body); dbnum > 0) {
        // `[DBNum1]` / `[DBNum2]` / `[DBNum3]`: digit-substitution mode for
        // the rest of the section. Multiple directives stack last-write-wins,
        // matching Mac Excel's behaviour. The renderer applies the chosen
        // mapping at digit-emit time.
        switch (dbnum) {
          case 1:
            out.dbnum_mode = DbNumMode::kDBNum1;
            break;
          case 2:
            out.dbnum_mode = DbNumMode::kDBNum2;
            break;
          case 3:
            out.dbnum_mode = DbNumMode::kDBNum3;
            break;
          default:
            break;
        }
      } else {
        // Conditional-section directive `[>1000]`, `[<=0]`, ...
        // Only one predicate per section; if a second one appears,
        // last-write-wins (matches Mac Excel: the rightmost directive
        // shadows earlier ones in the same section).
        CondOp cop = CondOp::kNone;
        double cval = 0.0;
        const int rc = parse_cond_directive(body, &cop, &cval);
        if (rc > 0) {
          out.cond_op = cop;
          out.cond_value = cval;
        } else {
          // Either the body doesn't look like a predicate (rc == 0) or it
          // does but the numeric tail failed to parse (rc < 0). Both surface
          // as #VALUE! through the existing invalid-bracket channel.
          out.has_invalid_bracket = true;
        }
      }
      continue;
    }
    // Underscore-skip `_X`: Excel reserves the width of character `X` and
    // emits a matching amount of whitespace. TEXT's output uses a single
    // space regardless of `X`. If `_` is the last byte of the format with
    // nothing following, fall through to the single-byte literal path.
    if (c == '_' && i + 1 < fmt.size()) {
      Token t;
      t.kind = Tok::Space;
      toks.push_back(t);
      i += 1 + utf8_scalar_width(fmt, i + 1);
      continue;
    }
    // Asterisk-fill `*X`: in cell formats this pads the cell with `X` to
    // fill the column width. TEXT() has no column width, so Mac Excel 365
    // emits this as a no-op (both `*` and the fill char are skipped). If
    // `*` is the last byte of the format with nothing following, fall
    // through to the single-byte literal path.
    if (c == '*' && i + 1 < fmt.size()) {
      i += 1 + utf8_scalar_width(fmt, i + 1);
      continue;
    }
    // `General` keyword (case-insensitive). Must be a standalone "word":
    // we only accept it when the following byte (if any) is not an ASCII
    // letter, so `Generally` and similar words pass through as literals.
    if (c == 'G' || c == 'g') {
      auto match_general = [&](std::size_t start) -> bool {
        static const char kWord[] = "general";
        if (start + 7 > fmt.size()) {
          return false;
        }
        for (std::size_t k = 0; k < 7; ++k) {
          const char fc = fmt[start + k];
          const char fc_lower = (fc >= 'A' && fc <= 'Z') ? static_cast<char>(fc + 32) : fc;
          if (fc_lower != kWord[k]) {
            return false;
          }
        }
        // Boundary check: next byte must not be a letter.
        if (start + 7 < fmt.size()) {
          const char nx = fmt[start + 7];
          if ((nx >= 'A' && nx <= 'Z') || (nx >= 'a' && nx <= 'z')) {
            return false;
          }
        }
        return true;
      };
      if (match_general(i)) {
        Token t;
        t.kind = Tok::GeneralNumber;
        toks.push_back(t);
        i += 7;
        continue;
      }
    }
    // AM/PM (case-insensitive). Match the longest valid prefix. We treat
    // `AM/PM`, `am/pm`, `A/P`, `a/p` as indivisible markers.
    if (c == 'A' || c == 'a' || c == 'P' || c == 'p') {
      auto match_ci = [&](std::size_t start, const char* a) -> bool {
        std::size_t n = 0;
        while (a[n] != '\0') {
          ++n;
        }
        if (start + n > fmt.size()) {
          return false;
        }
        for (std::size_t k = 0; k < n; ++k) {
          const char fc = fmt[start + k];
          const char ac = a[k];
          const char fc_lower = (fc >= 'A' && fc <= 'Z') ? static_cast<char>(fc + 32) : fc;
          const char ac_lower = (ac >= 'A' && ac <= 'Z') ? static_cast<char>(ac + 32) : ac;
          if (fc_lower != ac_lower) {
            return false;
          }
        }
        return true;
      };
      if (match_ci(i, "AM/PM")) {
        Token t;
        t.kind = Tok::AmPm;
        toks.push_back(t);
        i += 5;
        continue;
      }
      if (match_ci(i, "A/P")) {
        Token t;
        t.kind = Tok::AP;
        toks.push_back(t);
        i += 3;
        continue;
      }
    }
    // ja-JP weekday tokens `aaa` / `aaaa` (case-insensitive). Note that
    // `aaa` does NOT collide with `AM/PM` or `A/P` because those are
    // matched first above; a bare run of `a`/`A` characters falls through
    // here. A run shorter than 3 is not a weekday token in Excel and is
    // emitted as a literal.
    if (c == 'a' || c == 'A') {
      const std::size_t run = scan_run(fmt, i, 'a');
      if (run >= 3) {
        Token t;
        t.kind = (run >= 4) ? Tok::DateAaaa : Tok::DateAaa;
        t.width = static_cast<std::uint8_t>(run);
        toks.push_back(t);
        continue;
      }
      // Run of 1 or 2 `a` characters: emit as literal.
      push_literal(i - run, i);
      continue;
    }
    // ja-JP era name tokens: `g` (Roman 1-letter), `gg` (1-char kanji),
    // `ggg` or longer (full kanji name). Case-insensitive. The `General`
    // keyword check above already handled the literal "General" word.
    if (c == 'g' || c == 'G') {
      const std::size_t run = scan_run(fmt, i, 'g');
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      if (run >= 3) {
        t.kind = Tok::EraGGG;
      } else if (run == 2) {
        t.kind = Tok::EraGG;
      } else {
        t.kind = Tok::EraG;
      }
      toks.push_back(t);
      continue;
    }
    // Date letters.
    if (is_date_letter(c)) {
      const char lc = (c >= 'A' && c <= 'Z') ? static_cast<char>(c + 32) : c;
      const std::size_t run = scan_run(fmt, i, lc);
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      switch (lc) {
        case 'y':
          t.kind = (run <= 2) ? Tok::DateY2 : Tok::DateY4;
          break;
        case 'm':
          // Month/minute disambiguation happens in pass 2. A run of 5 or
          // more `m` characters means "first letter of the English month
          // name" (Excel's `mmmmm` convention).
          if (run >= 5) {
            t.kind = Tok::DateMMMMM;
          } else if (run == 4) {
            t.kind = Tok::DateMMMM;
          } else if (run == 3) {
            t.kind = Tok::DateMMM;
          } else {
            t.kind = Tok::DateMOrMin;
          }
          break;
        case 'd':
          if (run >= 4) {
            t.kind = Tok::DateDDDD;
          } else if (run == 3) {
            t.kind = Tok::DateDDD;
          } else if (run == 2) {
            t.kind = Tok::DateDD;
          } else {
            t.kind = Tok::DateD;
          }
          break;
        case 'h':
          t.kind = (run >= 2) ? Tok::DateHH : Tok::DateH;
          break;
        case 's':
          t.kind = (run >= 2) ? Tok::DateSS : Tok::DateS;
          break;
        default:
          t.kind = Tok::Literal;
          t.lit_begin = i - run;
          t.lit_end = i;
          break;
      }
      toks.push_back(t);
      continue;
    }
    // Scientific notation `E+` / `E-` / `e+` / `e-`. Must immediately
    // follow a digit token; we still emit the token and let the classifier
    // require the sign.
    if ((c == 'E' || c == 'e') && i + 1 < fmt.size() && (fmt[i + 1] == '+' || fmt[i + 1] == '-')) {
      Token t;
      t.kind = fmt[i + 1] == '+' ? Tok::SciPlus : Tok::SciMinus;
      toks.push_back(t);
      i += 2;
      continue;
    }
    // ja-JP era year token `e` / `ee`. A bare `e` (or run of `e`) not
    // followed by `+`/`-` is the era-year placeholder when the section is
    // a date section. The renderer falls back to a literal `e` when the
    // section turns out to be numeric.
    if (c == 'e' || c == 'E') {
      const std::size_t run = scan_run(fmt, i, 'e');
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      t.kind = (run >= 2) ? Tok::EraEE : Tok::EraE;
      toks.push_back(t);
      continue;
    }
    // Numeric specifiers.
    switch (c) {
      case '0': {
        Token t;
        t.kind = Tok::DigitZero;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '#': {
        Token t;
        t.kind = Tok::DigitOpt;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '?': {
        Token t;
        t.kind = Tok::DigitPad;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '.': {
        Token t;
        t.kind = Tok::Point;
        toks.push_back(t);
        ++i;
        continue;
      }
      case ',': {
        Token t;
        t.kind = Tok::Comma;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '%': {
        Token t;
        t.kind = Tok::Percent;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '@': {
        Token t;
        t.kind = Tok::At;
        toks.push_back(t);
        ++i;
        continue;
      }
      default:
        break;
    }
    // Fallback: preserve one complete UTF-8 scalar as a literal. Malformed
    // input is consumed one byte at a time by `utf8_scalar_width`, so it is
    // still copied without crashing or swallowing following syntax.
    const std::size_t literal_width = utf8_scalar_width(fmt, i);
    push_literal(i, i + literal_width);
    i += literal_width;
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
