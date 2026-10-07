//
// Implementation of the Ptg-stream -> AST decoder. See
// `io/xlsb/ptg_reader.h` for the contract and the [MS-XLSB] references.
//
// The decoder is an operand-stack machine. Operand Ptgs push a freshly
// built AST node; operator / function Ptgs pop their arity and push a
// combined node. The stack must hold exactly one node at end-of-stream.
// Every multi-byte read goes through the bounds-checked `read_*` helpers
// in `record.h`, so a truncated or malformed stream returns an Error
// instead of reading out of bounds.

#include "io/xlsb/ptg_reader.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "io/xlsb/record.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// MS-XLSB column field: low 14 bits are the 0-based column, bit 14 is
// the "column relative" flag and bit 15 the "row relative" flag. An
// absolute coordinate is the *cleared* relative bit.
constexpr std::uint16_t kColMask = 0x3FFF;
/// Rows in the grid; a relative row offset is stored modulo this.
constexpr std::uint32_t kRowCount = 1U << 20;
constexpr std::uint16_t kColRelBit = 0x4000;
constexpr std::uint16_t kRowRelBit = 0x8000;

// Excel sheet dimensions: the RefErr forms encode the maximum sentinel
// row/col; we never rely on those because the RefErr Ptg kind already
// tells us the reference is `#REF!`.
Error unsupported_ptg(std::uint8_t first_byte, const char* name) {
  std::string ctx("context=xlsb_ptg_reader byte=0x");
  static constexpr char kHex[] = "0123456789ABCDEF";
  ctx.push_back(kHex[(first_byte >> 4) & 0xF]);
  ctx.push_back(kHex[first_byte & 0xF]);
  if (name != nullptr) {
    ctx.append(" ptg=").append(name);
  }
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb formula uses an unsupported Ptg token",
                    std::move(ctx));
}

Error corrupt_stack(const char* detail) {
  return make_error(FormulonErrorCode::kIoXlsbCorrupt, std::string("xlsb formula operand stack imbalance: ") + detail,
                    "context=xlsb_ptg_reader");
}

// `PtgMemArea` has one matching PtgExtraMem in RgbExtra: a u32 count
// followed by `count` 16-byte UncheckedRfX records. The ranges cache the
// result of the following binary-reference expression; the expression's own
// Ptgs remain authoritative for the AST, so this reader only validates and
// consumes the opaque cache payload to keep later RgbExtra entries aligned.
Expected<void, Error> skip_ptg_extra_mem(ByteSpan& extra) {
  auto count_or = read_u32(extra);
  if (!count_or) {
    return count_or.error();
  }
  constexpr std::size_t kUncheckedRfXBytes = 16U;
  const std::size_t count = count_or.value();
  if (count > extra.size / kUncheckedRfXBytes) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgExtraMem range array truncated",
                      "context=xlsb_ptg_reader");
  }
  const std::size_t bytes = count * kUncheckedRfXBytes;
  extra.data += bytes;
  extra.size -= bytes;
  return Expected<void, Error>::Ok();
}

// Mem Ptgs end in a u16 byte length of the following binary-reference
// expression. That expression remains in `cursor` and is decoded normally;
// validating the advertised bound catches a malformed cache marker without
// skipping the actual formula.
Expected<void, Error> read_mem_expression_size(ByteSpan& cursor, const char* ptg_name) {
  auto cce_or = read_u16(cursor);
  if (!cce_or) {
    return cce_or.error();
  }
  if (static_cast<std::size_t>(cce_or.value()) > cursor.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                      std::string(ptg_name) + " binary-reference expression truncated", "context=xlsb_ptg_reader");
  }
  return Expected<void, Error>::Ok();
}

/// Maps an MS-XLSB error wire code to the engine `ErrorCode`. Delegates
/// to the single `kErrorTable`-backed lookup so this path can never drift
/// from the writer's `ooxml_code()` (see `error_from_ooxml_code` in
/// `value.h`) — a wire code that round-trips through the writer always
/// reads back as the same `ErrorCode`, including `#SPILL!` / `#CALC!` /
/// `#FIELD!` / `#BLOCKED!` / `#CONNECT!` / `#EXTERNAL!` / `#BUSY!` /
/// `#PYTHON!`, which a hand-duplicated switch previously missed.
ErrorCode error_from_wire(std::uint8_t code) {
  return error_from_ooxml_code(static_cast<std::int32_t>(code));
}

/// Reads an XLSB rgce string operand (PtgStr): a u16 code-unit count
/// followed by UTF-16LE units. (The engine's writer emits the same
/// shape; this matched pair is what guarantees round-trip.)
Expected<std::string, Error> read_ptg_string(ByteSpan& cursor) {
  auto cch_or = read_u16(cursor);
  if (!cch_or) {
    return cch_or.error();
  }
  const std::uint32_t cch = cch_or.value();
  std::string out;
  out.reserve(cch);
  for (std::uint32_t i = 0; i < cch; ++i) {
    auto unit_or = read_u16(cursor);
    if (!unit_or) {
      return unit_or.error();
    }
    const std::uint16_t cu = unit_or.value();
    // Best-effort UTF-16 -> UTF-8 for the BMP. Surrogate handling mirrors
    // `read_xlwidestring`: lone units pass through as their code point.
    if (cu < 0x80) {
      out.push_back(static_cast<char>(cu));
    } else if (cu < 0x800) {
      out.push_back(static_cast<char>(0xC0 | (cu >> 6)));
      out.push_back(static_cast<char>(0x80 | (cu & 0x3F)));
    } else {
      out.push_back(static_cast<char>(0xE0 | (cu >> 12)));
      out.push_back(static_cast<char>(0x80 | ((cu >> 6) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | (cu & 0x3F)));
    }
  }
  return out;
}

// Builds one corner from its row and its flag-carrying column field.
parser::Reference make_corner(std::uint32_t row, std::uint16_t col, std::string_view sheet) {
  parser::Reference ref;
  ref.sheet = sheet;
  ref.row = row;
  ref.col = static_cast<std::uint32_t>(col & kColMask);
  ref.col_abs = (col & kColRelBit) == 0;
  ref.row_abs = (col & kRowRelBit) == 0;
  return ref;
}

/// Decodes the `RgceLoc` single-cell coordinate (u32 row + u16 col with
/// relative-flag bits) into a `parser::Reference`. `sheet` is applied as
/// the reference's sheet qualifier (empty for the local sheet).
Expected<parser::Reference, Error> read_loc(ByteSpan& cursor, std::string_view sheet) {
  auto row_or = read_u32(cursor);
  if (!row_or) {
    return row_or.error();
  }
  auto col_or = read_u16(cursor);
  if (!col_or) {
    return col_or.error();
  }
  return make_corner(row_or.value(), col_or.value(), sheet);
}

/// Decodes the `RgceArea` two-corner range coordinate: rows first, then
/// columns, i.e. `row1(u32), row2(u32), col1(u16 w/ flags), col2(u16 w/
/// flags)` — NOT two back-to-back `RgceLoc` pairs. Verified against a
/// real Excel-365-produced `xl/worksheets/sheetN.bin`.
Expected<std::pair<parser::Reference, parser::Reference>, Error> read_area(ByteSpan& cursor,
                                                                           std::string_view sheet_first,
                                                                           std::string_view sheet_last) {
  auto row1_or = read_u32(cursor);
  if (!row1_or) {
    return row1_or.error();
  }
  auto row2_or = read_u32(cursor);
  if (!row2_or) {
    return row2_or.error();
  }
  auto col1_or = read_u16(cursor);
  if (!col1_or) {
    return col1_or.error();
  }
  auto col2_or = read_u16(cursor);
  if (!col2_or) {
    return col2_or.error();
  }
  return std::make_pair(make_corner(row1_or.value(), col1_or.value(), sheet_first),
                        make_corner(row2_or.value(), col2_or.value(), sheet_last));
}

// `first`:`last` as a formula spells it. Excel stores whole columns (`A:A`)
// as an area over every row with the rows absolute, and whole rows alike.
parser::AstNode* make_area(Arena& arena, parser::Reference first, parser::Reference last) {
  const bool cols = first.row == 0U && last.row == Sheet::kMaxRows - 1U && first.row_abs && last.row_abs;
  const bool rows = !cols && first.col == 0U && last.col == Sheet::kMaxCols - 1U && first.col_abs && last.col_abs;
  if (cols || rows) {
    first.is_full_col = last.is_full_col = cols;
    first.is_full_row = last.is_full_row = rows;
    const bool one = cols ? first.col == last.col && first.col_abs == last.col_abs
                          : first.row == last.row && first.row_abs == last.row_abs;
    if (one) {
      return parser::make_ref(arena, first);
    }
  }
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  return lhs == nullptr || rhs == nullptr ? nullptr : parser::make_range_op(arena, lhs, rhs);
}

// The external counterpart of `make_area`: whole columns and rows read back
// as such, a single one as the cell form with its flag set.
parser::AstNode* make_external_area(Arena& arena, std::string_view book, std::string_view sheet,
                                    std::string_view sheet_end, parser::Reference first, parser::Reference last) {
  const bool cols = first.row == 0U && last.row == Sheet::kMaxRows - 1U && first.row_abs && last.row_abs;
  const bool rows = !cols && first.col == 0U && last.col == Sheet::kMaxCols - 1U && first.col_abs && last.col_abs;
  bool is_range = true;
  if (cols || rows) {
    first.is_full_col = last.is_full_col = cols;
    first.is_full_row = last.is_full_row = rows;
    is_range = cols ? !(first.col == last.col && first.col_abs == last.col_abs)
                    : !(first.row == last.row && first.row_abs == last.row_abs);
    if (!is_range) {
      last = first;
    }
  }
  return parser::make_external_ref(arena, {}, book, sheet, sheet_end, first, last, is_range);
}

// Resolves a `PtgRefN` / `PtgAreaN` corner read by `read_loc` / `read_area`
// against `base`: a relative axis holds an offset modulo the grid.
void resolve_relative(parser::Reference& ref, PtgBaseCell base) {
  if (!ref.row_abs) {
    ref.row = (base.row + ref.row) & (kRowCount - 1U);
  }
  if (!ref.col_abs) {
    ref.col = (base.col + ref.col) & kColMask;
  }
}

// Validates a decoded single-cell `Reference` against the Excel grid
// bound (`Sheet::kMaxRows` / `Sheet::kMaxCols`) before it is materialized
// into an AST node. `PtgRef`/`PtgRef3d` col fields are already masked to
// 14 bits by `read_loc` (always < `kMaxCols`); `row` is a raw u32 and has
// no such guarantee, so a crafted `row=0xFFFFFFFF` must be rejected here
// rather than silently wrapping in `format_a1`. RefErr/AreaErr payload
// coordinates are never routed through this check -- their sentinel
// max-row/col encoding is a legitimate `#REF!` payload, not a corrupt
// live reference.
Expected<void, Error> check_ref_domain(const parser::Reference& r, const char* ptg_name) {
  if (r.row >= Sheet::kMaxRows || r.col >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb ") + ptg_name + " coordinate out of range", "context=xlsb_ptg_reader");
  }
  return {};
}

// Validates a decoded two-corner `Reference` pair: both corners must be
// in-domain and the range must be normalized (`row_first <= row_last`,
// `col_first <= col_last`), matching `parser::Reference`'s documented
// contract for range endpoints.
Expected<void, Error> check_area_domain(const parser::Reference& first, const parser::Reference& last,
                                        const char* ptg_name) {
  auto first_or = check_ref_domain(first, ptg_name);
  if (!first_or) {
    return first_or;
  }
  auto last_or = check_ref_domain(last, ptg_name);
  if (!last_or) {
    return last_or;
  }
  if (first.row > last.row || first.col > last.col) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb ") + ptg_name + " corners out of order", "context=xlsb_ptg_reader");
  }
  return {};
}

// `read_loc` followed by `check_ref_domain`.
Expected<parser::Reference, Error> read_checked_loc(ByteSpan& cursor, std::string_view sheet, const char* ptg_name) {
  auto ref_or = read_loc(cursor, sheet);
  if (!ref_or) {
    return ref_or.error();
  }
  auto domain_or = check_ref_domain(ref_or.value(), ptg_name);
  if (!domain_or) {
    return domain_or.error();
  }
  return ref_or;
}

// `read_area` followed by `check_area_domain`.
Expected<std::pair<parser::Reference, parser::Reference>, Error> read_checked_area(ByteSpan& cursor,
                                                                                   std::string_view sheet_first,
                                                                                   const char* ptg_name) {
  auto area_or = read_area(cursor, sheet_first, {});
  if (!area_or) {
    return area_or.error();
  }
  auto domain_or = check_area_domain(area_or.value().first, area_or.value().second, ptg_name);
  if (!domain_or) {
    return domain_or.error();
  }
  return area_or;
}

/// Case-insensitive `s` starts-with `prefix` check (ASCII-fold).
bool starts_with_ci(std::string_view s, std::string_view prefix) {
  return s.size() >= prefix.size() && strings::case_insensitive_eq(s.substr(0, prefix.size()), prefix);
}

}  // namespace

Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, ByteSpan rgcb, Arena& arena,
                                              const std::vector<std::string>& sheet_names,
                                              const std::vector<XlsbName>& name_table,
                                              const std::vector<XlsbSheetRange>& sheet_ranges,
                                              const XlsbExternalBooks& external_books, std::int32_t host_itab,
                                              std::optional<PtgBaseCell> base) {
  std::vector<parser::AstNode*> stack;
  // `PtgArray` stores only an 8-byte placeholder inline (see its case
  // below); the real dimensions + elements are consumed from this
  // cursor in encounter order.
  ByteSpan extra = rgcb;

  auto pop = [&stack]() -> parser::AstNode* {
    parser::AstNode* n = stack.back();
    stack.pop_back();
    return n;
  };

  // True when `ixti` names an ExternSheet entry belonging to a
  // supporting book other than this one. Such an entry's `itabFirst` /
  // `itabLast` index the external workbook's own sheet list, so
  // resolving them against `sheet_names` would rebind the reference to
  // an unrelated local sheet (or, out of range, degrade it to an
  // unqualified same-sheet reference). Both outcomes change the value
  // silently, so the token is surfaced as undecodable instead.
  auto ixti_is_external = [&sheet_ranges](std::uint32_t ixti) -> bool {
    return ixti < sheet_ranges.size() && sheet_ranges[ixti].external_book != 0U;
  };

  // Resolves an external `ixti` to the supporting book's `[N]` index and
  // the names of the sheets it spans (`sheet_end_out` empty for one sheet).
  // Returns 0 when the reference cannot be bound: an unknown book, or a
  // sheet index the supporting book's own table does not cover. The caller
  // then reports the token undecodable and Excel's cached value stands.
  auto external_sheet_for_ixti = [&sheet_ranges, &external_books](std::uint32_t ixti, std::string_view& sheet_out,
                                                                  std::string_view& sheet_end_out) -> std::uint32_t {
    if (ixti >= sheet_ranges.size()) {
      return 0;
    }
    const XlsbSheetRange& range = sheet_ranges[ixti];
    if (range.external_book == 0 || range.external_book > external_books.size()) {
      return 0;
    }
    const std::vector<std::string>& sheets = external_books[range.external_book - 1U].sheet_names;
    const auto in_book = [&sheets](std::int32_t itab) {
      return itab >= 0 && static_cast<std::size_t>(itab) < sheets.size();
    };
    if (!in_book(range.itab_first) || !in_book(range.itab_last)) {
      return 0;
    }
    sheet_out = sheets[static_cast<std::size_t>(range.itab_first)];
    sheet_end_out =
        range.itab_first == range.itab_last ? std::string_view() : sheets[static_cast<std::size_t>(range.itab_last)];
    return range.external_book;
  };

  // Single-sheet resolution for `ixti`: prefers the `BrtExternSheet`
  // table's `itabFirst` when present, falling back to treating `ixti`
  // as a direct 0-based `sheet_names` index when the workbook carries
  // no ExternSheet table at all (e.g. no qualified references).
  auto sheet_for_ixti = [&sheet_names, &sheet_ranges](std::uint32_t ixti) -> std::string_view {
    if (!sheet_ranges.empty()) {
      if (ixti >= sheet_ranges.size()) {
        return {};
      }
      const std::int32_t itab = sheet_ranges[ixti].itab_first;
      if (itab < 0 || static_cast<std::size_t>(itab) >= sheet_names.size()) {
        return {};
      }
      return sheet_names[static_cast<std::size_t>(itab)];
    }
    if (ixti < sheet_names.size()) {
      return sheet_names[ixti];
    }
    return {};
  };

  // Multi-sheet (genuine 3-D) resolution for `ixti`: returns `true` and
  // populates `begin_out` / `end_out` only when the ExternSheet entry
  // spans more than one sheet; the caller falls back to
  // `sheet_for_ixti` (a plain qualified reference) otherwise.
  auto sheet_range_for_ixti = [&sheet_names, &sheet_ranges](std::uint32_t ixti, std::string_view& begin_out,
                                                            std::string_view& end_out) -> bool {
    if (ixti >= sheet_ranges.size()) {
      return false;
    }
    const XlsbSheetRange& r = sheet_ranges[ixti];
    if (r.itab_first == r.itab_last) {
      return false;
    }
    if (r.itab_first < 0 || r.itab_last < 0 || static_cast<std::size_t>(r.itab_first) >= sheet_names.size() ||
        static_cast<std::size_t>(r.itab_last) >= sheet_names.size()) {
      return false;
    }
    begin_out = sheet_names[static_cast<std::size_t>(r.itab_first)];
    end_out = sheet_names[static_cast<std::size_t>(r.itab_last)];
    return true;
  };

  // `PtgName`'s `ilbl` is 1-based; out-of-range resolves to an empty
  // name (caller surfaces `kIoXlsbCorrupt`).
  auto resolve_name = [&name_table](std::uint32_t ilbl) -> std::string_view {
    if (ilbl == 0 || ilbl > name_table.size()) {
      return {};
    }
    return name_table[ilbl - 1].name;
  };

  // Builds the `NameRef` for this workbook's `name_table[ilbl - 1]`,
  // qualifying it when the entry is local to a sheet other than the host,
  // or to any sheet when the token was `PtgNameX`: Excel stores `Sheet1!Fn`
  // that way even on Sheet1, and bare `Fn` as `PtgName`.
  auto make_local_name_ref = [&arena, &name_table, &sheet_names, host_itab](std::uint32_t ilbl,
                                                                            bool qualified) -> parser::AstNode* {
    const XlsbName& entry = name_table[ilbl - 1];
    // `_xlpm.`-prefixed names are LET/LAMBDA-local parameter references;
    // strip the storage prefix at every use site (not just the LET
    // binding-name slot decoded_future_function handles) so a bare `x*3`
    // reference inside a LET body matches the plain-identifier `NameRef`
    // the text parser would have produced for the same formula.
    const std::string_view name = entry.name;
    const std::string_view display_name = starts_with_ci(name, "_xlpm.") ? name.substr(6) : name;
    if (entry.itab >= 0 && (qualified || entry.itab != host_itab) &&
        static_cast<std::size_t>(entry.itab) < sheet_names.size()) {
      const std::string& sheet = sheet_names[static_cast<std::size_t>(entry.itab)];
      return parser::make_sheet_name_ref(arena, sheet, display_name, /*sheet_quoted=*/false);
    }
    return parser::make_name_ref(arena, arena.intern(display_name));
  };

  // Resolves a `PtgFuncVar` with the `id == 255` future-function
  // sentinel. `cparams` operands are popped; the first (in original
  // push order) must be a `NameRef` naming the real callee. `_xlfn.LET`
  // gets its own AST shape (`LetBinding`) because its remaining
  // operands alternate `_xlpm.*`-prefixed parameter-name references
  // with value expressions, terminated by the body — not a flat
  // argument list a generic `Call` node can represent. Every other
  // future function (XLOOKUP, TEXTJOIN, CONCAT, IFS, SEQUENCE, ...)
  // becomes a plain `Call` keeping the `_xlfn.` prefix intact, matching
  // the OOXML storage convention (`eval/tree_walker/dispatch.cpp`'s
  // `strip_future_prefix` removes it at evaluation time).
  auto decode_future_function = [&arena, &pop, &stack](std::uint32_t cparams) -> Expected<parser::AstNode*, Error> {
    if (cparams == 0 || stack.size() < cparams) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb PtgFuncVar(255): operand stack underflow",
                        "context=xlsb_ptg_reader cparams=" + std::to_string(cparams));
    }
    auto** ops = arena.create_array<const parser::AstNode*>(cparams);
    if (ops == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgFuncVar(255) operands)",
                        "context=xlsb_ptg_reader");
    }
    for (std::uint32_t i = 0; i < cparams; ++i) {
      ops[cparams - 1 - i] = pop();
    }
    const std::uint32_t real_count = cparams - 1;
    // `LAMBDA(z,z*z)(4)` and a curried call store the callee expression
    // itself as the first operand; `[0]!Fn(3)` stores the self-book name,
    // and `A1(1)` / `(A1:A2)(1)` / `(A1,B1)(1)` the reference.
    const parser::NodeKind callee_kind = ops[0]->kind();
    if (callee_kind == parser::NodeKind::Lambda || callee_kind == parser::NodeKind::LambdaCall ||
        callee_kind == parser::NodeKind::Ref || callee_kind == parser::NodeKind::RangeOp ||
        callee_kind == parser::NodeKind::UnionOp || callee_kind == parser::NodeKind::IntersectOp ||
        parser::is_self_book_name_ref(*ops[0])) {
      parser::AstNode* n = parser::make_lambda_call(arena, const_cast<parser::AstNode*>(ops[0]), ops + 1, real_count);
      if (n == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LAMBDA call)", "context=xlsb_ptg_reader");
      }
      return n;
    }
    if (ops[0]->kind() != parser::NodeKind::NameRef) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt,
                        "xlsb PtgFuncVar(255): callee operand is not a name reference", "context=xlsb_ptg_reader");
    }
    const std::string_view callee = ops[0]->as_name();
    if (strings::case_insensitive_eq(callee, "_xlfn.LAMBDA")) {
      // LAMBDA(param1, ..., body): `_xlpm.*` parameter-name references
      // followed by the body.
      if (real_count < 1) {
        return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb LAMBDA: missing body", "context=xlsb_ptg_reader");
      }
      const std::uint32_t param_count = real_count - 1;
      auto* params = param_count == 0 ? nullptr : arena.create_array<std::string_view>(param_count);
      if (param_count != 0 && params == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LAMBDA parameters)",
                          "context=xlsb_ptg_reader");
      }
      for (std::uint32_t i = 0; i < param_count; ++i) {
        const parser::AstNode* param = ops[1 + i];
        if (param->kind() != parser::NodeKind::NameRef || !param->as_name_sheet().empty()) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb LAMBDA: parameter operand is not a name reference",
                            "context=xlsb_ptg_reader");
        }
        const std::string_view raw = param->as_name();
        params[i] = starts_with_ci(raw, "_xlpm.") ? raw.substr(6) : raw;
      }
      auto* body = const_cast<parser::AstNode*>(ops[cparams - 1]);
      parser::AstNode* n = parser::make_lambda(arena, params, param_count, /*optional_count=*/0, body);
      if (n == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LAMBDA node)", "context=xlsb_ptg_reader");
      }
      return n;
    }
    if (starts_with_ci(callee, "_xlfn.LET")) {
      // LET(name1, value1, [name2, value2, ...], body): an odd count of
      // >= 3 real operands (name/value pairs plus a trailing body).
      if (real_count < 3 || (real_count % 2) == 0) {
        return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb LET: malformed operand count",
                          "context=xlsb_ptg_reader real_count=" + std::to_string(real_count));
      }
      const std::uint32_t binding_count = (real_count - 1) / 2;
      auto* names = arena.create_array<std::string_view>(binding_count);
      auto** exprs = arena.create_array<const parser::AstNode*>(binding_count);
      if (names == nullptr || exprs == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LET bindings)", "context=xlsb_ptg_reader");
      }
      for (std::uint32_t b = 0; b < binding_count; ++b) {
        const parser::AstNode* name_node = ops[1 + (2 * b)];
        if (name_node->kind() != parser::NodeKind::NameRef) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb LET: binding name operand is not a name reference",
                            "context=xlsb_ptg_reader");
        }
        std::string_view raw = name_node->as_name();
        names[b] = starts_with_ci(raw, "_xlpm.") ? raw.substr(6) : raw;
        exprs[b] = ops[2 + (2 * b)];
      }
      auto* body = const_cast<parser::AstNode*>(ops[cparams - 1]);
      parser::AstNode* n = parser::make_let_binding(arena, names, exprs, binding_count, body);
      if (n == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LET node)", "context=xlsb_ptg_reader");
      }
      return n;
    }
    auto** args = real_count == 0 ? nullptr : arena.create_array<const parser::AstNode*>(real_count);
    if (real_count != 0 && args == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (future-function args)",
                        "context=xlsb_ptg_reader");
    }
    for (std::uint32_t i = 0; i < real_count; ++i) {
      args[i] = ops[1 + i];
    }
    // Another sheet's local name keeps its qualifier: `Sheet2!Fn(3)` is a
    // call through that sheet's scope, which only a `LambdaCall` can carry.
    if (!ops[0]->as_name_sheet().empty()) {
      parser::AstNode* n = parser::make_lambda_call(arena, const_cast<parser::AstNode*>(ops[0]), args, real_count);
      if (n == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (qualified name call)",
                          "context=xlsb_ptg_reader");
      }
      return n;
    }
    parser::AstNode* n = parser::make_call(arena, arena.intern(callee), args, real_count);
    if (n == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (future-function call)",
                        "context=xlsb_ptg_reader");
    }
    return n;
  };

  // Pops `arity` operands into a `Call` of the classic function `name`;
  // `token` names the Ptg in the OOM messages, `detail` the stack imbalance.
  auto pop_call = [&arena, &pop, &stack](const char* name, std::uint32_t arity, const char* token,
                                         const char* detail) -> Expected<parser::AstNode*, Error> {
    if (stack.size() < arity) {
      return corrupt_stack(detail);
    }
    auto** args = arity == 0 ? nullptr : arena.create_array<const parser::AstNode*>(arity);
    if (arity != 0 && args == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, std::string("arena exhausted (") + token + " args)",
                        "context=xlsb_ptg_reader");
    }
    for (std::uint32_t i = 0; i < arity; ++i) {
      args[arity - 1 - i] = pop();
    }
    parser::AstNode* n = parser::make_call(arena, arena.intern(name), args, arity);
    if (n == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, std::string("arena exhausted (") + token + " call)",
                        "context=xlsb_ptg_reader");
    }
    return n;
  };

  ByteSpan cursor = ptgs;
  while (cursor.size > 0) {
    const std::uint8_t first_byte = cursor.data[0];
    const PtgInfo* info = lookup_ptg_from_wire(first_byte);
    if (info == nullptr) {
      return unsupported_ptg(first_byte, nullptr);
    }
    if (info->status == PtgStatus::Unsupported) {
      return unsupported_ptg(first_byte, info->name);
    }
    // Consume the dispatch byte.
    cursor.data += 1;
    cursor.size -= 1;

    switch (info->kind) {
      // ---- Memory/cache markers ------------------------------------------
      // These markers do not push an operand. Their following expression is
      // still encoded in the normal Ptg stream, so preserve it by consuming
      // only the marker payload (and the matching PtgMemArea extra cache).
      case PtgKind::MemArea:
      case PtgKind::MemNoMem: {
        auto unused_or = read_u32(cursor);
        if (!unused_or) {
          return unused_or.error();
        }
        auto size_check = read_mem_expression_size(cursor, info->name);
        if (!size_check) {
          return size_check.error();
        }
        if (info->kind == PtgKind::MemArea) {
          auto extra_check = skip_ptg_extra_mem(extra);
          if (!extra_check) {
            return extra_check.error();
          }
        }
        break;
      }
      case PtgKind::MemErr: {
        auto error_or = read_u8(cursor);
        if (!error_or) {
          return error_or.error();
        }
        auto unused_or = read_u8(cursor);
        if (!unused_or) {
          return unused_or.error();
        }
        auto unused2_or = read_u16(cursor);
        if (!unused2_or) {
          return unused2_or.error();
        }
        auto size_check = read_mem_expression_size(cursor, info->name);
        if (!size_check) {
          return size_check.error();
        }
        break;
      }
      case PtgKind::MemFunc: {
        auto size_check = read_mem_expression_size(cursor, info->name);
        if (!size_check) {
          return size_check.error();
        }
        break;
      }

      // ---- Operands -------------------------------------------------------
      case PtgKind::Int: {
        auto v_or = read_u16(cursor);
        if (!v_or) {
          return v_or.error();
        }
        parser::AstNode* n = parser::make_literal(arena, Value::number(static_cast<double>(v_or.value())));
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgInt)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Num: {
        if (cursor.size < 8) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgNum payload truncated",
                            "context=xlsb_ptg_reader");
        }
        double v;
        std::memcpy(&v, cursor.data, sizeof(v));
        cursor.data += 8;
        cursor.size -= 8;
        // Excel stores a typed `-1` as the negative number; read it back as
        // the formula spells it, a minus over the number.
        parser::AstNode* n = parser::make_literal(arena, Value::number(v < 0.0 ? -v : v));
        if (n != nullptr && v < 0.0) {
          n = parser::make_unary_op(arena, parser::UnaryOp::Minus, n);
        }
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgNum)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Str: {
        auto s_or = read_ptg_string(cursor);
        if (!s_or) {
          return s_or.error();
        }
        const std::string_view interned = arena.intern(s_or.value());
        parser::AstNode* n = parser::make_literal(arena, Value::text(interned));
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgStr)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Bool: {
        auto b_or = read_u8(cursor);
        if (!b_or) {
          return b_or.error();
        }
        parser::AstNode* n = parser::make_literal(arena, Value::boolean(b_or.value() != 0));
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgBool)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Err: {
        auto code_or = read_u8(cursor);
        if (!code_or) {
          return code_or.error();
        }
        parser::AstNode* n = parser::make_error_literal(arena, error_from_wire(code_or.value()));
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgErr)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::MissArg: {
        // An omitted argument (e.g. `IF(,x,y)`) maps to a blank literal;
        // the formatter renders it as an empty slot.
        parser::AstNode* n = parser::make_literal(arena, Value::blank());
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgMissArg)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Array: {
        // The main token stream carries only a 15-byte placeholder (the
        // class-marked opcode, already consumed by the caller, + 14
        // reserved bytes here -- verified against a real Excel-produced
        // `xl/worksheets/sheetN.bin`). The real dimensions and elements
        // live in `extra` (the `CellParsedFormula`'s `rgcb`), consumed
        // here in encounter order.
        //
        // The first u32 is the row count and the second the column count,
        // elements row-major, each a tag byte and its payload (measured on
        // non-square Excel constants, `xlsb_phantom_cells.xlsb`).
        if (cursor.size < 14) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgArray placeholder truncated",
                            "context=xlsb_ptg_reader");
        }
        cursor.data += 14;
        cursor.size -= 14;
        auto rows_or = read_u32(extra);
        if (!rows_or) {
          return rows_or.error();
        }
        auto cols_or = read_u32(extra);
        if (!cols_or) {
          return cols_or.error();
        }
        const std::uint32_t rows = rows_or.value();
        const std::uint32_t cols = cols_or.value();
        if (rows == 0 || cols == 0) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "PtgArray zero dimension", "context=xlsb_ptg_reader");
        }
        if (static_cast<std::uint64_t>(rows) * cols > 0x10000U) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "PtgArray dimension overflow",
                            "context=xlsb_ptg_reader");
        }
        const std::uint32_t count = rows * cols;
        auto** elems = arena.create_array<const parser::AstNode*>(count);
        if (elems == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArray elems)",
                            "context=xlsb_ptg_reader");
        }
        for (std::uint32_t i = 0; i < count; ++i) {
          auto tag_or = read_u8(extra);
          if (!tag_or) {
            return tag_or.error();
          }
          parser::AstNode* elem = nullptr;
          switch (tag_or.value()) {
            case 0: {  // number (verified: tag byte 0x00 precedes the double)
              if (extra.size < 8) {
                return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgArray number truncated",
                                  "context=xlsb_ptg_reader");
              }
              double v;
              std::memcpy(&v, extra.data, sizeof(v));
              extra.data += 8;
              extra.size -= 8;
              elem = parser::make_literal(arena, Value::number(v));
              break;
            }
            case 1: {  // string: u16 count + UTF-16LE, as PtgStr carries it
              auto s_or = read_ptg_string(extra);
              if (!s_or) {
                return s_or.error();
              }
              elem = parser::make_literal(arena, Value::text(arena.intern(s_or.value())));
              break;
            }
            case 2: {  // boolean: one byte
              auto b_or = read_u8(extra);
              if (!b_or) {
                return b_or.error();
              }
              elem = parser::make_literal(arena, Value::boolean(b_or.value() != 0));
              break;
            }
            case 4: {  // error: the code, then three unused bytes
              auto e_or = read_u32(extra);
              if (!e_or) {
                return e_or.error();
              }
              elem =
                  parser::make_error_literal(arena, error_from_wire(static_cast<std::uint8_t>(e_or.value() & 0xFFU)));
              break;
            }
            default:
              return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "PtgArray element tag not decoded",
                                "context=xlsb_ptg_reader tag=" + std::to_string(tag_or.value()));
          }
          if (elem == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArray element)",
                              "context=xlsb_ptg_reader");
          }
          elems[i] = elem;
        }
        parser::AstNode* n = parser::make_array_literal(arena, rows, cols, elems);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArray)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- Names ------------------------------------------------------------
      case PtgKind::Name: {
        // `ilbl` (1-based) indexes the workbook's `BrtName` table
        // (`name_table`). Ordinary defined names ("Rate") and the
        // hidden `_xlfn.*` / `_xlpm.*` future-function / LET-parameter
        // placeholders share this same token; the future-function
        // dispatch above (`decode_future_function`) is what tells them
        // apart, by inspecting the resolved name's prefix.
        auto ilbl_or = read_u32(cursor);
        if (!ilbl_or) {
          return ilbl_or.error();
        }
        const std::string_view name = resolve_name(ilbl_or.value());
        if (name.empty()) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb PtgName: ilbl out of range",
                            "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl_or.value()));
        }
        parser::AstNode* n = make_local_name_ref(ilbl_or.value(), /*qualified=*/false);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgName)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      case PtgKind::NameX: {
        // `ixti` names an ExternSheet entry whose supporting book is the
        // one holding the name, and `ilbl` (1-based) indexes that book's
        // own name table. The entry's sheet indices are `-2` here (the
        // name is book-scope, so it qualifies no sheet), which is why
        // this resolves the book through `sheet_ranges` but never asks
        // `external_sheet_for_ixti` for a sheet.
        auto ixti_or = read_u16(cursor);
        if (!ixti_or) {
          return ixti_or.error();
        }
        auto ilbl_or = read_u32(cursor);
        if (!ilbl_or) {
          return ilbl_or.error();
        }
        const std::uint32_t ixti = ixti_or.value();
        // An ExternSheet entry of this workbook names one of its own
        // `BrtName` records; that record's own scope decides the qualifier.
        // A workbook-scoped record is the self-book `[0]!Name`, which is
        // what Excel saves both that spelling and `Sheet1!Name` as when
        // Sheet1 has no local `Name`.
        if (ixti < sheet_ranges.size() && sheet_ranges[ixti].external_book == 0U) {
          const std::string_view name = resolve_name(ilbl_or.value());
          if (name.empty()) {
            return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb PtgNameX: ilbl out of range",
                              "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl_or.value()));
          }
          parser::AstNode* n = name_table[ilbl_or.value() - 1U].itab < 0
                                   ? parser::make_external_name_ref(arena, {}, "0", {}, name)
                                   : make_local_name_ref(ilbl_or.value(), /*qualified=*/true);
          if (n == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgNameX)", "context=xlsb_ptg_reader");
          }
          stack.push_back(n);
          break;
        }
        const std::uint32_t book = ixti < sheet_ranges.size() ? sheet_ranges[ixti].external_book : 0U;
        if (book == 0 || book > external_books.size()) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                            "PtgNameX names a supporting workbook this reader cannot bind", "context=xlsb_ptg_reader");
        }
        const std::vector<std::string>& names = external_books[book - 1U].names;
        if (ilbl_or.value() == 0 || ilbl_or.value() > names.size()) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                            "PtgNameX ilbl is outside the supporting workbook's name table",
                            "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl_or.value()));
        }
        // The book is spelled as its link index; ingestion rewrites it to the
        // file name the formula bar shows.
        parser::AstNode* n =
            parser::make_external_name_ref(arena, {}, std::to_string(book), {}, names[ilbl_or.value() - 1U]);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgNameX)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- References -----------------------------------------------------
      case PtgKind::Ref: {
        auto ref_or = read_checked_loc(cursor, {}, "PtgRef");
        if (!ref_or) {
          return ref_or.error();
        }
        parser::AstNode* n = parser::make_ref(arena, ref_or.value());
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRef)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::RefN:
      case PtgKind::AreaN: {
        if (!base) {
          return unsupported_ptg(first_byte, info->name);
        }
        parser::AstNode* n = nullptr;
        if (info->kind == PtgKind::RefN) {
          auto ref_or = read_loc(cursor, {});
          if (!ref_or) {
            return ref_or.error();
          }
          resolve_relative(ref_or.value(), *base);
          auto domain_or = check_ref_domain(ref_or.value(), info->name);
          if (!domain_or) {
            return domain_or.error();
          }
          n = parser::make_ref(arena, ref_or.value());
        } else {
          auto area_or = read_area(cursor, {}, {});
          if (!area_or) {
            return area_or.error();
          }
          resolve_relative(area_or.value().first, *base);
          resolve_relative(area_or.value().second, *base);
          auto domain_or = check_area_domain(area_or.value().first, area_or.value().second, info->name);
          if (!domain_or) {
            return domain_or.error();
          }
          n = make_area(arena, area_or.value().first, area_or.value().second);
        }
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRefN/PtgAreaN)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Area: {
        auto area_or = read_checked_area(cursor, {}, "PtgArea");
        if (!area_or) {
          return area_or.error();
        }
        parser::AstNode* n = make_area(arena, area_or.value().first, area_or.value().second);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea range)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Ref3d: {
        auto ixti_or = read_u16(cursor);
        if (!ixti_or) {
          return ixti_or.error();
        }
        // A single-cell 3-D reference can span more than one sheet
        // (e.g. `Data:S2!B1`) entirely through the ExternSheet entry's
        // `(itabFirst, itabLast)` — the Ptg token itself is identical to
        // the single-sheet form. Build a `Ref3D` node when the range is
        // genuinely multi-sheet; otherwise the plain qualified `Ref`
        // this token already produced.
        if (ixti_is_external(ixti_or.value())) {
          std::string_view external_sheet;
          std::string_view external_sheet_end;
          const std::uint32_t book = external_sheet_for_ixti(ixti_or.value(), external_sheet, external_sheet_end);
          if (book == 0) {
            return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                              "PtgRef3d qualifies a sheet of an external workbook this reader cannot bind",
                              "context=xlsb_ptg_reader");
          }
          auto loc_or = read_checked_loc(cursor, {}, "PtgRef3d");
          if (!loc_or) {
            return loc_or.error();
          }
          parser::AstNode* n = parser::make_external_ref(arena, {}, std::to_string(book), external_sheet,
                                                         external_sheet_end, loc_or.value(), loc_or.value(),
                                                         /*is_range=*/false);
          if (n == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRef3d external)",
                              "context=xlsb_ptg_reader");
          }
          stack.push_back(n);
          break;
        }
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti_or.value(), begin_sheet, end_sheet)) {
          auto loc_or = read_checked_loc(cursor, {}, "PtgRef3d");
          if (!loc_or) {
            return loc_or.error();
          }
          parser::AstNode* n =
              parser::make_ref3d(arena, arena.intern(begin_sheet), arena.intern(end_sheet), loc_or.value());
          if (n == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRef3d range)",
                              "context=xlsb_ptg_reader");
          }
          stack.push_back(n);
          break;
        }
        auto ref_or = read_checked_loc(cursor, sheet_for_ixti(ixti_or.value()), "PtgRef3d");
        if (!ref_or) {
          return ref_or.error();
        }
        parser::Reference ref = ref_or.value();
        ref.sheet = arena.intern(ref.sheet);
        ref.sheet_quoted = parser::local_sheet_needs_quoting_a1(ref.sheet);
        parser::AstNode* n = parser::make_ref(arena, ref);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRef3d)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Area3d: {
        auto ixti_or = read_u16(cursor);
        if (!ixti_or) {
          return ixti_or.error();
        }
        // A genuine 3-D range (multiple sheets AND a cell rectangle, e.g.
        // `Sheet1:Sheet2!A1:B2`) decodes into a range-tail `Ref3D`. A
        // single-sheet qualified area (`Sheet2!A1:B2`) keeps the plain
        // `RangeOp` of two qualified refs.
        if (ixti_is_external(ixti_or.value())) {
          std::string_view external_sheet;
          std::string_view external_sheet_end;
          const std::uint32_t book = external_sheet_for_ixti(ixti_or.value(), external_sheet, external_sheet_end);
          if (book == 0) {
            return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                              "PtgArea3d qualifies a sheet of an external workbook this reader cannot bind",
                              "context=xlsb_ptg_reader");
          }
          auto area_or = read_checked_area(cursor, {}, "PtgArea3d");
          if (!area_or) {
            return area_or.error();
          }
          parser::AstNode* n = make_external_area(arena, std::to_string(book), external_sheet, external_sheet_end,
                                                  area_or.value().first, area_or.value().second);
          if (n == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea3d external)",
                              "context=xlsb_ptg_reader");
          }
          stack.push_back(n);
          break;
        }
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti_or.value(), begin_sheet, end_sheet)) {
          auto area_or = read_checked_area(cursor, {}, "PtgArea3d");
          if (!area_or) {
            return area_or.error();
          }
          parser::AstNode* n = parser::make_ref3d_range(arena, arena.intern(begin_sheet), arena.intern(end_sheet),
                                                        area_or.value().first, area_or.value().second);
          if (n == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea3d range)",
                              "context=xlsb_ptg_reader");
          }
          stack.push_back(n);
          break;
        }
        const std::string_view sheet = sheet_for_ixti(ixti_or.value());
        auto area_or = read_checked_area(cursor, sheet, "PtgArea3d");
        if (!area_or) {
          return area_or.error();
        }
        parser::Reference first = area_or.value().first;
        first.sheet = arena.intern(first.sheet);
        first.sheet_quoted = parser::local_sheet_needs_quoting_a1(first.sheet);
        parser::AstNode* n = make_area(arena, first, area_or.value().second);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea3d range)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::RefErr:
      case PtgKind::RefErr3d:
      case PtgKind::AreaErr:
      case PtgKind::AreaErr3d: {
        // Consume the payload (ixti for the 3d forms, then one loc, or two
        // for the area forms) and emit a `#REF!` literal.
        const bool area = info->kind == PtgKind::AreaErr || info->kind == PtgKind::AreaErr3d;
        if (info->kind == PtgKind::RefErr3d || info->kind == PtgKind::AreaErr3d) {
          auto ixti_or = read_u16(cursor);
          if (!ixti_or) {
            return ixti_or.error();
          }
        }
        for (int i = area ? 2 : 1; i > 0; --i) {
          auto skip = read_loc(cursor, {});
          if (!skip) {
            return skip.error();
          }
        }
        parser::AstNode* n = parser::make_error_literal(arena, ErrorCode::Ref);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory,
                            area ? "arena exhausted (PtgAreaErr)" : "arena exhausted (PtgRefErr)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- Binary operators ----------------------------------------------
      case PtgKind::Add:
      case PtgKind::Sub:
      case PtgKind::Mul:
      case PtgKind::Div:
      case PtgKind::Power:
      case PtgKind::Concat:
      case PtgKind::Lt:
      case PtgKind::Le:
      case PtgKind::Eq:
      case PtgKind::Ge:
      case PtgKind::Gt:
      case PtgKind::Ne: {
        if (stack.size() < 2) {
          return corrupt_stack("binary operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        parser::BinOp op = parser::BinOp::Add;
        switch (info->kind) {
          case PtgKind::Add:
            op = parser::BinOp::Add;
            break;
          case PtgKind::Sub:
            op = parser::BinOp::Sub;
            break;
          case PtgKind::Mul:
            op = parser::BinOp::Mul;
            break;
          case PtgKind::Div:
            op = parser::BinOp::Div;
            break;
          case PtgKind::Power:
            op = parser::BinOp::Pow;
            break;
          case PtgKind::Concat:
            op = parser::BinOp::Concat;
            break;
          case PtgKind::Lt:
            op = parser::BinOp::Lt;
            break;
          case PtgKind::Le:
            op = parser::BinOp::LtEq;
            break;
          case PtgKind::Eq:
            op = parser::BinOp::Eq;
            break;
          case PtgKind::Ge:
            op = parser::BinOp::GtEq;
            break;
          case PtgKind::Gt:
            op = parser::BinOp::Gt;
            break;
          case PtgKind::Ne:
            op = parser::BinOp::NotEq;
            break;
          default:
            break;
        }
        parser::AstNode* n = parser::make_binary_op(arena, op, lhs, rhs);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (binary op)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- Range / set operators -----------------------------------------
      case PtgKind::Range: {
        if (stack.size() < 2) {
          return corrupt_stack("range operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        parser::AstNode* n = parser::make_range_op(arena, lhs, rhs);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (range op)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Union: {
        if (stack.size() < 2) {
          return corrupt_stack("union operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        const parser::AstNode* children[2] = {lhs, rhs};
        parser::AstNode* n = parser::make_union_op(arena, children, 2);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (union op)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Isect: {
        if (stack.size() < 2) {
          return corrupt_stack("intersect operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        parser::AstNode* n = parser::make_intersect_op(arena, lhs, rhs);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (intersect op)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- Unary operators -----------------------------------------------
      case PtgKind::Uplus:
      case PtgKind::Uminus:
      case PtgKind::Percent: {
        if (stack.empty()) {
          return corrupt_stack("unary operator");
        }
        parser::AstNode* operand = pop();
        parser::UnaryOp op = parser::UnaryOp::Plus;
        if (info->kind == PtgKind::Uminus) {
          op = parser::UnaryOp::Minus;
        } else if (info->kind == PtgKind::Percent) {
          op = parser::UnaryOp::Percent;
        }
        parser::AstNode* n = parser::make_unary_op(arena, op, operand);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (unary op)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Paren: {
        // Evaluation ignores parentheses; the operand records them so the
        // formula text keeps the ones Excel shows.
        if (stack.empty()) {
          return corrupt_stack("paren");
        }
        stack.back()->add_paren();
        break;
      }

      // ---- Functions ------------------------------------------------------
      case PtgKind::Func: {
        auto id_or = read_u16(cursor);
        if (!id_or) {
          return id_or.error();
        }
        const XlsbFuncEntry* entry = lookup_func_by_id(id_or.value());
        if (entry == nullptr) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb PtgFunc unknown function id",
                            "context=xlsb_ptg_reader id=" + std::to_string(id_or.value()));
        }
        // Fixed arity.
        auto call_or = pop_call(entry->name, entry->arg_min, "PtgFunc", "function (fixed)");
        if (!call_or) {
          return call_or.error();
        }
        stack.push_back(call_or.value());
        break;
      }
      case PtgKind::FuncVar: {
        auto cparams_or = read_u8(cursor);
        if (!cparams_or) {
          return cparams_or.error();
        }
        auto id_or = read_u16(cursor);
        if (!id_or) {
          return id_or.error();
        }
        const std::uint32_t cparams = cparams_or.value();
        // id == 255 is the "future function" sentinel: the real callee
        // is not in the classic function-id table at all (XLOOKUP, LET,
        // TEXTJOIN, CONCAT, IFS, SEQUENCE, ...). Its name was pushed as
        // the FIRST operand via a preceding `PtgName`, so `cparams`
        // counts that name-ref plus the real arguments.
        if (id_or.value() == 255) {
          auto call_or = decode_future_function(cparams);
          if (!call_or) {
            return call_or.error();
          }
          stack.push_back(call_or.value());
          break;
        }
        const XlsbFuncEntry* entry = lookup_func_by_id(id_or.value());
        if (entry == nullptr) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb PtgFuncVar unknown function id",
                            "context=xlsb_ptg_reader id=" + std::to_string(id_or.value()));
        }
        auto call_or = pop_call(entry->name, cparams, "PtgFuncVar", "function (var)");
        if (!call_or) {
          return call_or.error();
        }
        stack.push_back(call_or.value());
        break;
      }

      // ---- Attributes -----------------------------------------------------
      case PtgKind::Attr: {
        auto sub_or = read_u8(cursor);
        if (!sub_or) {
          return sub_or.error();
        }
        const auto sub = static_cast<PtgAttrKind>(sub_or.value());
        switch (sub) {
          case PtgAttrKind::Sum: {
            // Optimised single-argument SUM. The attr carries a u16 of
            // unused data; collapse the top operand into `SUM(x)`.
            auto unused_or = read_u16(cursor);
            if (!unused_or) {
              return unused_or.error();
            }
            if (stack.empty()) {
              return corrupt_stack("attr-sum");
            }
            parser::AstNode* operand = pop();
            const parser::AstNode* args[1] = {operand};
            parser::AstNode* n = parser::make_call(arena, arena.intern("SUM"), args, 1);
            if (n == nullptr) {
              return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (attr-sum)",
                                "context=xlsb_ptg_reader");
            }
            stack.push_back(n);
            break;
          }
          case PtgAttrKind::Space:
          case PtgAttrKind::SpaceSemi: {
            // Whitespace attr: two bytes of (type, count) to skip. The
            // operand stack is untouched.
            if (cursor.size < 2) {
              return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgAttrSpace payload truncated",
                                "context=xlsb_ptg_reader");
            }
            cursor.data += 2;
            cursor.size -= 2;
            break;
          }
          case PtgAttrKind::If:
          case PtgAttrKind::Choose:
          case PtgAttrKind::Goto:
          case PtgAttrKind::Semi:
          case PtgAttrKind::Baxcel:
          default: {
            // Control / volatile attrs carry a u16 (If/Goto/Semi) or a
            // jump table (Choose: u16 count + (count+1) u16 offsets).
            // [MS-XLSB] 2.5.98.25 defines rgOffset as an array of 2-byte
            // unsigned integers, not 4-byte -- reading them as u32 desyncs
            // the rest of the Ptg stream for any Excel-authored CHOOSE().
            if (sub == PtgAttrKind::Choose) {
              auto count_or = read_u16(cursor);
              if (!count_or) {
                return count_or.error();
              }
              const std::uint32_t entries = static_cast<std::uint32_t>(count_or.value()) + 1U;
              for (std::uint32_t i = 0; i < entries; ++i) {
                auto off_or = read_u16(cursor);
                if (!off_or) {
                  return off_or.error();
                }
              }
            } else {
              auto unused_or = read_u16(cursor);
              if (!unused_or) {
                return unused_or.error();
              }
            }
            // These attrs are control-flow only; they do not consume or
            // produce operands.
            break;
          }
        }
        break;
      }

      // ---- IFERROR optimisation marker (transparent) ----------------------
      case PtgKind::IfError: {
        // Treated as a no-op marker; the surrounding IFERROR call is
        // reconstructed from its PtgFuncVar. Nothing to read.
        break;
      }

      default:
        return unsupported_ptg(first_byte, info->name);
    }
  }

  if (stack.size() != 1) {
    return corrupt_stack(stack.empty() ? "empty stack at end" : "multiple values at end");
  }
  if (!parser::ast_depth_within_limit(*stack.front(), parser::kMaxFormulaAstDepth)) {
    return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb formula exceeds maximum AST depth",
                      "context=xlsb_ptg_reader max_depth=" + std::to_string(parser::kMaxFormulaAstDepth));
  }
  return stack.front();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
