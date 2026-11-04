// @size-budget: 64 KB

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
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "io/xlsb/ptg_list.h"
#include "io/xlsb/ptg_space.h"
#include "io/xlsb/record.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Rows in the grid; a relative row offset is stored modulo this.
constexpr std::uint32_t kRowCount = 1U << 20;

// Excel sheet dimensions: the RefErr forms encode the maximum sentinel
// row/col; we never rely on those because the RefErr Ptg kind already
// tells us the reference is `#REF!`.
Error unsupported_ptg(std::uint8_t first_byte, const char* name) {
  std::string ctx("context=xlsb_ptg_reader byte=0x");
  ctx.push_back(strings::hex_digit_upper(first_byte >> 4U));
  ctx.push_back(strings::hex_digit_upper(first_byte));
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
  ASSIGN_OR_RETURN(const std::size_t count, read_u32(extra));
  constexpr std::size_t kUncheckedRfXBytes = 16U;
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
  ASSIGN_OR_RETURN(auto cce, read_u16(cursor));
  if (static_cast<std::size_t>(cce) > cursor.size) {
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
  ASSIGN_OR_RETURN(auto cch, read_u16(cursor));
  std::string out;
  if (!read_utf16_units(cursor, cch, &out)) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb PtgStr body truncated", "context=xlsb.ptg");
  }
  return out;
}

/// Pushes `node` on the operand stack, or fails when the arena could not
/// allocate it; `what` names the token for the message.
Expected<void, Error> push_node(std::vector<parser::AstNode*>& stack, parser::AstNode* node, const char* what) {
  if (node == nullptr) {
    return make_error(FormulonErrorCode::kOutOfMemory, std::string("arena exhausted (") + what + ")",
                      "context=xlsb_ptg_reader");
  }
  stack.push_back(node);
  return Expected<void, Error>::Ok();
}

// Builds one corner from its row and its flag-carrying column field.
parser::Reference make_corner(std::uint32_t row, std::uint16_t col, std::string_view sheet) {
  parser::Reference ref;
  ref.sheet = sheet;
  ref.row = row;
  ref.col = static_cast<std::uint32_t>(col & kPtgColumnMask);
  ref.col_abs = (col & kPtgColumnRelativeBit) == 0;
  ref.row_abs = (col & kPtgRowRelativeBit) == 0;
  return ref;
}

/// Decodes the `RgceLoc` single-cell coordinate (u32 row + u16 col with
/// relative-flag bits) into a `parser::Reference`. `sheet` is applied as
/// the reference's sheet qualifier (empty for the local sheet).
Expected<parser::Reference, Error> read_loc(ByteSpan& cursor, std::string_view sheet) {
  ASSIGN_OR_RETURN(auto row, read_u32(cursor));
  ASSIGN_OR_RETURN(auto col, read_u16(cursor));
  return make_corner(row, col, sheet);
}

/// Decodes the `RgceArea` two-corner range coordinate: rows first, then
/// columns, i.e. `row1(u32), row2(u32), col1(u16 w/ flags), col2(u16 w/
/// flags)` — NOT two back-to-back `RgceLoc` pairs. Verified against a
/// real Excel-365-produced `xl/worksheets/sheetN.bin`.
Expected<std::pair<parser::Reference, parser::Reference>, Error> read_area(ByteSpan& cursor,
                                                                           std::string_view sheet_first,
                                                                           std::string_view sheet_last) {
  ASSIGN_OR_RETURN(auto row1, read_u32(cursor));
  ASSIGN_OR_RETURN(auto row2, read_u32(cursor));
  ASSIGN_OR_RETURN(auto col1, read_u16(cursor));
  ASSIGN_OR_RETURN(auto col2, read_u16(cursor));
  return std::make_pair(make_corner(row1, col1, sheet_first), make_corner(row2, col2, sheet_last));
}

// How an area spells in a formula: Excel stores whole columns (`A:A`) as an
// area over every row with the rows absolute, and whole rows alike. `one` is
// set when the whole-column or whole-row area covers a single column or row.
struct FullAxis {
  bool cols = false;
  bool rows = false;
  bool one = false;
};

FullAxis classify_full_axis(const parser::Reference& first, const parser::Reference& last) {
  FullAxis axis;
  axis.cols = first.row == 0U && last.row == Sheet::kMaxRows - 1U && first.row_abs && last.row_abs;
  axis.rows = !axis.cols && first.col == 0U && last.col == Sheet::kMaxCols - 1U && first.col_abs && last.col_abs;
  axis.one = axis.cols ? first.col == last.col && first.col_abs == last.col_abs
                       : axis.rows && first.row == last.row && first.row_abs == last.row_abs;
  return axis;
}

// `first`:`last` as a formula spells it.
parser::AstNode* make_area(Arena& arena, parser::Reference first, parser::Reference last) {
  const FullAxis axis = classify_full_axis(first, last);
  const bool full = axis.cols || axis.rows;
  if (full) {
    first.is_full_col = last.is_full_col = axis.cols;
    first.is_full_row = last.is_full_row = axis.rows;
    if (axis.one) {
      return parser::make_ref(arena, first);
    }
  }
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  return lhs == nullptr || rhs == nullptr ? nullptr : parser::make_range_op(arena, lhs, rhs, /*spliced=*/full);
}

// The deleted-sheet `#REF!` of the area `first`:`last` (`#REF!A1:B2`,
// `#REF!A:A`), spelled as `make_area` spells the area.
parser::AstNode* make_dead_sheet_area(Arena& arena, parser::Reference first, parser::Reference last) {
  const FullAxis axis = classify_full_axis(first, last);
  if (axis.cols || axis.rows) {
    first.is_full_col = last.is_full_col = axis.cols;
    first.is_full_row = last.is_full_row = axis.rows;
    if (axis.one) {
      return parser::make_dead_sheet_ref_error(arena, &first, nullptr);
    }
  }
  return parser::make_dead_sheet_ref_error(arena, &first, &last);
}

// The external counterpart of `make_area`: whole columns and rows read back
// as such, a single one as the cell form with its flag set.
parser::AstNode* make_external_area(Arena& arena, std::string_view book, std::string_view sheet,
                                    std::string_view sheet_end, parser::Reference first, parser::Reference last) {
  const FullAxis axis = classify_full_axis(first, last);
  if (axis.cols || axis.rows) {
    first.is_full_col = last.is_full_col = axis.cols;
    first.is_full_row = last.is_full_row = axis.rows;
    if (axis.one) {
      last = first;
    }
  }
  return parser::make_external_ref(arena, {}, book, sheet, sheet_end, first, last, /*is_range=*/!axis.one);
}

// Resolves a `PtgRefN` / `PtgAreaN` corner read by `read_loc` / `read_area`
// against `base`: a relative axis holds an offset modulo the grid.
void resolve_relative(parser::Reference& ref, PtgBaseCell base) {
  if (!ref.row_abs) {
    ref.row = (base.row + ref.row) & (kRowCount - 1U);
  }
  if (!ref.col_abs) {
    ref.col = (base.col + ref.col) & kPtgColumnMask;
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
  RETURN_IF_ERROR(check_ref_domain(first, ptg_name));
  RETURN_IF_ERROR(check_ref_domain(last, ptg_name));
  if (first.row > last.row || first.col > last.col) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb ") + ptg_name + " corners out of order", "context=xlsb_ptg_reader");
  }
  return {};
}

// `read_loc` followed by `check_ref_domain`.
Expected<parser::Reference, Error> read_checked_loc(ByteSpan& cursor, std::string_view sheet, const char* ptg_name) {
  ASSIGN_OR_RETURN(auto ref, read_loc(cursor, sheet));
  RETURN_IF_ERROR(check_ref_domain(ref, ptg_name));
  return ref;
}

// `read_area` followed by `check_area_domain`.
Expected<std::pair<parser::Reference, parser::Reference>, Error> read_checked_area(ByteSpan& cursor,
                                                                                   std::string_view sheet_first,
                                                                                   const char* ptg_name) {
  ASSIGN_OR_RETURN(auto area, read_area(cursor, sheet_first, {}));
  RETURN_IF_ERROR(check_area_domain(area.first, area.second, ptg_name));
  return area;
}

}  // namespace

Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, ByteSpan rgcb, Arena& arena,
                                              const std::vector<std::string>& sheet_names,
                                              const std::vector<XlsbName>& name_table,
                                              const std::vector<XlsbSheetRange>& sheet_ranges,
                                              const XlsbExternalBooks& external_books, std::int32_t host_itab,
                                              std::optional<PtgBaseCell> base, parser::FormulaWhitespace* whitespace,
                                              const std::vector<TableMetadata>* tables) {
  std::vector<parser::AstNode*> stack;
  // `PtgArray` stores only an 8-byte placeholder inline (see its case
  // below); the real dimensions + elements are consumed from this
  // cursor in encounter order.
  ByteSpan extra = rgcb;

  // A binary or `:` chain still open on the stack, held there by a stand-in
  // node: `1 2 + 3 4 * +` keeps `1+2` open under `3*4` so the second `+`
  // extends it, and `A1:B2 C3:D4 : E5 :` grows one range. Anything other
  // than a further operator of the same kind that consumes the stand-in
  // closes the chain into one `BinaryOp` or `RangeOp` node.
  // An open chain's stand-in stays on the stack until the chain closes, so
  // the chains close in reverse order of opening: `stand_ins` and
  // `chain_marks` hold the open ones, innermost last.
  parser::ChainBuilder chains;
  std::vector<parser::AstNode*> stand_ins;
  std::vector<std::uint32_t> chain_marks;
  bool chain_out_of_memory = false;
  // The mark of the chain `node` stands in for, or nullptr.
  auto find_open_chain = [&](const parser::AstNode* node) -> const std::uint32_t* {
    return !stand_ins.empty() && stand_ins.back() == node ? &chain_marks.back() : nullptr;
  };
  auto close_chain = [&](parser::AstNode* node) -> parser::AstNode* {
    if (find_open_chain(node) == nullptr) {
      return node;
    }
    parser::AstNode* closed = chains.close(arena, chain_marks.back());
    stand_ins.pop_back();
    chain_marks.pop_back();
    if (closed == nullptr) {
      chain_out_of_memory = true;
      return node;
    }
    if (whitespace != nullptr) {
      for (parser::WhitespaceRun& run : *whitespace) {
        if (run.node == node) {
          run.node = closed;
        }
      }
    }
    return closed;
  };

  auto pop = [&stack, &close_chain]() -> parser::AstNode* {
    parser::AstNode* n = stack.back();
    stack.pop_back();
    return close_chain(n);
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

  // True when `ixti` names this workbook's deleted-sheet entry (itab -1).
  auto ixti_is_dead_sheet = [&sheet_ranges](std::uint32_t ixti) -> bool {
    return ixti < sheet_ranges.size() && sheet_ranges[ixti].external_book == 0U && sheet_ranges[ixti].itab_first == -1;
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
    const std::string_view display_name = strings::case_insensitive_starts_with(name, "_xlpm.") ? name.substr(6) : name;
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
    // `LAMBDA(z,z*z)(4)`, a curried call and `CHOOSE(1,SUM,ABS)(5)` store the
    // callee expression itself as the first operand; `[0]!Fn(3)` stores the
    // self-book name, and `A1(1)` / `(A1:A2)(1)` / `(A1,B1)(1)` the reference.
    if (ops[0]->kind() != parser::NodeKind::NameRef || parser::is_self_book_name_ref(*ops[0])) {
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
        params[i] = strings::case_insensitive_starts_with(raw, "_xlpm.") ? raw.substr(6) : raw;
      }
      auto* body = const_cast<parser::AstNode*>(ops[cparams - 1]);
      parser::AstNode* n = parser::make_lambda(arena, params, param_count, /*optional_count=*/0, body);
      if (n == nullptr) {
        return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (LAMBDA node)", "context=xlsb_ptg_reader");
      }
      return n;
    }
    if (strings::case_insensitive_starts_with(callee, "_xlfn.LET")) {
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
        names[b] = strings::case_insensitive_starts_with(raw, "_xlpm.") ? raw.substr(6) : raw;
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

  // `PtgAttrSpace` attrs waiting for the token they precede: how many, and
  // the (type, count) of the first `kMaxPendingSpaces`.
  struct PendingSpace {
    std::uint8_t type;
    std::uint8_t count;
  };
  constexpr std::uint32_t kMaxPendingSpaces = 3;
  std::uint32_t pending_spaces = 0;
  PendingSpace pending[kMaxPendingSpaces] = {};

  // Records the pending attrs against `node`, the result of a `kind` token,
  // when they form one of the sequences `decode_ptgs` documents.
  auto record_space = [&](PtgKind kind, const parser::AstNode* node) {
    if (whitespace == nullptr || node == nullptr || pending_spaces == 0 || pending_spaces > kMaxPendingSpaces) {
      return;
    }
    parser::WhitespaceSlot slot = parser::WhitespaceSlot::kBeforeText;
    std::uint16_t index = 0;
    bool paren = false;
    switch (kind) {
      case PtgKind::Int:
      case PtgKind::Num:
      case PtgKind::Str:
      case PtgKind::Bool:
      case PtgKind::Err:
      case PtgKind::Array:
      case PtgKind::Name:
      case PtgKind::NameX:
      case PtgKind::Ref:
      case PtgKind::RefN:
      case PtgKind::Area:
      case PtgKind::AreaN:
      case PtgKind::Ref3d:
      case PtgKind::Area3d:
      case PtgKind::RefErr:
      case PtgKind::RefErr3d:
      case PtgKind::AreaErr:
      case PtgKind::AreaErr3d:
        break;
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
      case PtgKind::Ne:
      case PtgKind::Isect:
      case PtgKind::Percent:
        slot = parser::WhitespaceSlot::kBeforeOperator;
        // A binary operator's run is keyed by its position in the chain.
        if (const std::uint32_t* mark = find_open_chain(node); mark != nullptr) {
          index = static_cast<std::uint16_t>(chains.op_count(*mark) - 1U);
        }
        break;
      case PtgKind::Paren:
        slot = parser::WhitespaceSlot::kBeforeClose;
        index = node->paren_depth();
        paren = true;
        break;
      case PtgKind::Func:
      case PtgKind::FuncVar:
      case PtgKind::Attr:
        if (node->kind() != parser::NodeKind::Call) {
          return;
        }
        slot = parser::WhitespaceSlot::kBeforeClose;
        break;
      default:
        return;
    }
    const std::uint8_t spaces_type = ptg_space_type(slot, false);
    const std::uint8_t feeds_type = ptg_space_type(slot, true);
    // Spaces before a pair's `(`, then line feeds, then spaces, each at most once.
    std::uint32_t at = 0;
    if (paren && pending[at].type == kPtgSpaceBeforeOpen) {
      ++at;
    }
    for (const std::uint8_t type : {feeds_type, spaces_type}) {
      if (at < pending_spaces && pending[at].type == type) {
        ++at;
      }
    }
    if (at != pending_spaces) {
      return;
    }
    for (std::uint32_t i = 0; i < pending_spaces; ++i) {
      const PendingSpace& p = pending[i];
      if (p.count == 0) {
        continue;
      }
      if (paren && p.type == kPtgSpaceBeforeOpen) {
        whitespace->push_back({node, parser::WhitespaceSlot::kBeforeOpen, index, ' ', p.count});
      } else {
        whitespace->push_back({node, slot, index, p.type == feeds_type ? '\n' : ' ', p.count});
      }
    }
  };

  ByteSpan cursor = ptgs;
  while (cursor.size > 0) {
    // Set by the tokens a pending `PtgAttrSpace` passes over.
    bool keeps_pending_space = false;
    const std::uint8_t first_byte = cursor.data[0];
    if (first_byte == kPtgList && cursor.size >= 2U && cursor.data[1] == kEptgList) {
      ASSIGN_OR_RETURN(auto list, decode_ptg_list(cursor, arena, tables));
      stack.push_back(list);
      if (pending_spaces != 0) {
        record_space(PtgKind::Area, stack.back());
        pending_spaces = 0;
      }
      continue;
    }
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
        RETURN_IF_ERROR(read_u32(cursor));
        RETURN_IF_ERROR(read_mem_expression_size(cursor, info->name));
        if (info->kind == PtgKind::MemArea) {
          RETURN_IF_ERROR(skip_ptg_extra_mem(extra));
        }
        keeps_pending_space = true;
        break;
      }
      case PtgKind::MemErr: {
        RETURN_IF_ERROR(read_u8(cursor));
        RETURN_IF_ERROR(read_u8(cursor));
        RETURN_IF_ERROR(read_u16(cursor));
        RETURN_IF_ERROR(read_mem_expression_size(cursor, info->name));
        keeps_pending_space = true;
        break;
      }
      case PtgKind::MemFunc: {
        RETURN_IF_ERROR(read_mem_expression_size(cursor, info->name));
        keeps_pending_space = true;
        break;
      }

      // ---- Operands -------------------------------------------------------
      case PtgKind::Int: {
        ASSIGN_OR_RETURN(auto v, read_u16(cursor));
        parser::AstNode* n = parser::make_literal(arena, Value::number(static_cast<double>(v)));
        RETURN_IF_ERROR(push_node(stack, n, "PtgInt"));
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
        RETURN_IF_ERROR(push_node(stack, n, "PtgNum"));
        break;
      }
      case PtgKind::Str: {
        ASSIGN_OR_RETURN(auto s, read_ptg_string(cursor));
        const std::string_view interned = arena.intern(s);
        parser::AstNode* n = parser::make_literal(arena, Value::text(interned));
        RETURN_IF_ERROR(push_node(stack, n, "PtgStr"));
        break;
      }
      case PtgKind::Bool: {
        ASSIGN_OR_RETURN(auto b, read_u8(cursor));
        parser::AstNode* n = parser::make_literal(arena, Value::boolean(b != 0));
        RETURN_IF_ERROR(push_node(stack, n, "PtgBool"));
        break;
      }
      case PtgKind::Err: {
        ASSIGN_OR_RETURN(auto code, read_u8(cursor));
        parser::AstNode* n = parser::make_error_literal(arena, error_from_wire(code));
        RETURN_IF_ERROR(push_node(stack, n, "PtgErr"));
        break;
      }
      case PtgKind::MissArg: {
        // An omitted argument (e.g. `IF(,x,y)`) maps to a blank literal;
        // the formatter renders it as an empty slot.
        parser::AstNode* n = parser::make_literal(arena, Value::blank());
        RETURN_IF_ERROR(push_node(stack, n, "PtgMissArg"));
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
        // elements row-major, each a tag byte and its payload (see
        // `xlsb_phantom_cells.xlsb`).
        if (cursor.size < 14) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgArray placeholder truncated",
                            "context=xlsb_ptg_reader");
        }
        cursor.data += 14;
        cursor.size -= 14;
        ASSIGN_OR_RETURN(const std::uint32_t rows, read_u32(extra));
        ASSIGN_OR_RETURN(const std::uint32_t cols, read_u32(extra));
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
          ASSIGN_OR_RETURN(auto tag, read_u8(extra));
          parser::AstNode* elem = nullptr;
          switch (tag) {
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
              ASSIGN_OR_RETURN(auto s, read_ptg_string(extra));
              elem = parser::make_literal(arena, Value::text(arena.intern(s)));
              break;
            }
            case 2: {  // boolean: one byte
              ASSIGN_OR_RETURN(auto b, read_u8(extra));
              elem = parser::make_literal(arena, Value::boolean(b != 0));
              break;
            }
            case 4: {  // error: the code, then three unused bytes
              ASSIGN_OR_RETURN(auto e, read_u32(extra));
              elem = parser::make_error_literal(arena, error_from_wire(static_cast<std::uint8_t>(e & 0xFFU)));
              break;
            }
            default:
              return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "PtgArray element tag not decoded",
                                "context=xlsb_ptg_reader tag=" + std::to_string(tag));
          }
          if (elem == nullptr) {
            return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArray element)",
                              "context=xlsb_ptg_reader");
          }
          elems[i] = elem;
        }
        parser::AstNode* n = parser::make_array_literal(arena, rows, cols, elems);
        RETURN_IF_ERROR(push_node(stack, n, "PtgArray"));
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
        ASSIGN_OR_RETURN(auto ilbl, read_u32(cursor));
        const std::string_view name = resolve_name(ilbl);
        if (name.empty()) {
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb PtgName: ilbl out of range",
                            "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl));
        }
        parser::AstNode* n = make_local_name_ref(ilbl, /*qualified=*/false);
        RETURN_IF_ERROR(push_node(stack, n, "PtgName"));
        break;
      }

      case PtgKind::NameX: {
        // `ixti` names an ExternSheet entry whose supporting book is the
        // one holding the name, and `ilbl` (1-based) indexes that book's
        // own name table. The entry's sheet indices are `-2` here (the
        // name is book-scope, so it qualifies no sheet), which is why
        // this resolves the book through `sheet_ranges` but never asks
        // `external_sheet_for_ixti` for a sheet.
        ASSIGN_OR_RETURN(const std::uint32_t ixti, read_u16(cursor));
        ASSIGN_OR_RETURN(auto ilbl, read_u32(cursor));
        // An ExternSheet entry of this workbook names one of its own
        // `BrtName` records; that record's own scope decides the qualifier.
        // A workbook-scoped record is the self-book `[0]!Name`, which is
        // what Excel saves both that spelling and `Sheet1!Name` as when
        // Sheet1 has no local `Name`.
        if (ixti < sheet_ranges.size() && sheet_ranges[ixti].external_book == 0U) {
          const std::string_view name = resolve_name(ilbl);
          if (name.empty()) {
            return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb PtgNameX: ilbl out of range",
                              "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl));
          }
          parser::AstNode* n = name_table[ilbl - 1U].itab < 0 ? parser::make_external_name_ref(arena, {}, "0", {}, name)
                                                              : make_local_name_ref(ilbl, /*qualified=*/true);
          RETURN_IF_ERROR(push_node(stack, n, "PtgNameX"));
          break;
        }
        const std::uint32_t book = ixti < sheet_ranges.size() ? sheet_ranges[ixti].external_book : 0U;
        if (book == 0 || book > external_books.size()) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                            "PtgNameX names a supporting workbook this reader cannot bind", "context=xlsb_ptg_reader");
        }
        const std::vector<std::string>& names = external_books[book - 1U].names;
        if (ilbl == 0 || ilbl > names.size()) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                            "PtgNameX ilbl is outside the supporting workbook's name table",
                            "context=xlsb_ptg_reader ilbl=" + std::to_string(ilbl));
        }
        // The book is spelled as its link index; ingestion rewrites it to the
        // file name the formula bar shows.
        parser::AstNode* n = parser::make_external_name_ref(arena, {}, std::to_string(book), {}, names[ilbl - 1U]);
        RETURN_IF_ERROR(push_node(stack, n, "PtgNameX"));
        break;
      }

      // ---- References -----------------------------------------------------
      case PtgKind::Ref: {
        ASSIGN_OR_RETURN(auto ref, read_checked_loc(cursor, {}, "PtgRef"));
        parser::AstNode* n = parser::make_ref(arena, ref);
        RETURN_IF_ERROR(push_node(stack, n, "PtgRef"));
        break;
      }
      case PtgKind::RefN:
      case PtgKind::AreaN: {
        if (!base) {
          return unsupported_ptg(first_byte, info->name);
        }
        parser::AstNode* n = nullptr;
        if (info->kind == PtgKind::RefN) {
          ASSIGN_OR_RETURN(auto ref, read_loc(cursor, {}));
          resolve_relative(ref, *base);
          RETURN_IF_ERROR(check_ref_domain(ref, info->name));
          n = parser::make_ref(arena, ref);
        } else {
          ASSIGN_OR_RETURN(auto area, read_area(cursor, {}, {}));
          resolve_relative(area.first, *base);
          resolve_relative(area.second, *base);
          RETURN_IF_ERROR(check_area_domain(area.first, area.second, info->name));
          n = make_area(arena, area.first, area.second);
        }
        RETURN_IF_ERROR(push_node(stack, n, "PtgRefN/PtgAreaN"));
        break;
      }
      case PtgKind::Area: {
        ASSIGN_OR_RETURN(auto area, read_checked_area(cursor, {}, "PtgArea"));
        parser::AstNode* n = make_area(arena, area.first, area.second);
        RETURN_IF_ERROR(push_node(stack, n, "PtgArea range"));
        break;
      }
      case PtgKind::Ref3d: {
        ASSIGN_OR_RETURN(auto ixti, read_u16(cursor));
        // A single-cell 3-D reference can span more than one sheet
        // (e.g. `Data:S2!B1`) entirely through the ExternSheet entry's
        // `(itabFirst, itabLast)` — the Ptg token itself is identical to
        // the single-sheet form. Build a `Ref3D` node when the range is
        // genuinely multi-sheet; otherwise the plain qualified `Ref`
        // this token already produced.
        if (ixti_is_external(ixti)) {
          std::string_view external_sheet;
          std::string_view external_sheet_end;
          const std::uint32_t book = external_sheet_for_ixti(ixti, external_sheet, external_sheet_end);
          if (book == 0) {
            return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                              "PtgRef3d qualifies a sheet of an external workbook this reader cannot bind",
                              "context=xlsb_ptg_reader");
          }
          ASSIGN_OR_RETURN(auto loc, read_checked_loc(cursor, {}, "PtgRef3d"));
          parser::AstNode* n =
              parser::make_external_ref(arena, {}, std::to_string(book), external_sheet, external_sheet_end, loc, loc,
                                        /*is_range=*/false);
          RETURN_IF_ERROR(push_node(stack, n, "PtgRef3d external"));
          break;
        }
        if (ixti_is_dead_sheet(ixti)) {
          ASSIGN_OR_RETURN(auto loc, read_checked_loc(cursor, {}, "PtgRef3d"));
          parser::AstNode* n = parser::make_dead_sheet_ref_error(arena, &loc, nullptr);
          RETURN_IF_ERROR(push_node(stack, n, "PtgRef3d deleted sheet"));
          break;
        }
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti, begin_sheet, end_sheet)) {
          ASSIGN_OR_RETURN(auto loc, read_checked_loc(cursor, {}, "PtgRef3d"));
          parser::AstNode* n = parser::make_ref3d(arena, arena.intern(begin_sheet), arena.intern(end_sheet), loc);
          RETURN_IF_ERROR(push_node(stack, n, "PtgRef3d range"));
          break;
        }
        ASSIGN_OR_RETURN(parser::Reference ref, read_checked_loc(cursor, sheet_for_ixti(ixti), "PtgRef3d"));
        ref.sheet = arena.intern(ref.sheet);
        ref.sheet_quoted = parser::local_sheet_needs_quoting_a1(ref.sheet);
        parser::AstNode* n = parser::make_ref(arena, ref);
        RETURN_IF_ERROR(push_node(stack, n, "PtgRef3d"));
        break;
      }
      case PtgKind::Area3d: {
        ASSIGN_OR_RETURN(auto ixti, read_u16(cursor));
        // A genuine 3-D range (multiple sheets AND a cell rectangle, e.g.
        // `Sheet1:Sheet2!A1:B2`) decodes into a range-tail `Ref3D`. A
        // single-sheet qualified area (`Sheet2!A1:B2`) keeps the plain
        // `RangeOp` of two qualified refs.
        if (ixti_is_external(ixti)) {
          std::string_view external_sheet;
          std::string_view external_sheet_end;
          const std::uint32_t book = external_sheet_for_ixti(ixti, external_sheet, external_sheet_end);
          if (book == 0) {
            return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                              "PtgArea3d qualifies a sheet of an external workbook this reader cannot bind",
                              "context=xlsb_ptg_reader");
          }
          ASSIGN_OR_RETURN(auto area, read_checked_area(cursor, {}, "PtgArea3d"));
          parser::AstNode* n = make_external_area(arena, std::to_string(book), external_sheet, external_sheet_end,
                                                  area.first, area.second);
          RETURN_IF_ERROR(push_node(stack, n, "PtgArea3d external"));
          break;
        }
        if (ixti_is_dead_sheet(ixti)) {
          ASSIGN_OR_RETURN(auto area, read_checked_area(cursor, {}, "PtgArea3d"));
          parser::AstNode* n = make_dead_sheet_area(arena, area.first, area.second);
          RETURN_IF_ERROR(push_node(stack, n, "PtgArea3d deleted sheet"));
          break;
        }
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti, begin_sheet, end_sheet)) {
          ASSIGN_OR_RETURN(auto area, read_checked_area(cursor, {}, "PtgArea3d"));
          parser::AstNode* n = parser::make_ref3d_range(arena, arena.intern(begin_sheet), arena.intern(end_sheet),
                                                        area.first, area.second);
          RETURN_IF_ERROR(push_node(stack, n, "PtgArea3d range"));
          break;
        }
        const std::string_view sheet = sheet_for_ixti(ixti);
        ASSIGN_OR_RETURN(auto area, read_checked_area(cursor, sheet, "PtgArea3d"));
        parser::Reference first = area.first;
        first.sheet = arena.intern(first.sheet);
        first.sheet_quoted = parser::local_sheet_needs_quoting_a1(first.sheet);
        parser::AstNode* n = make_area(arena, first, area.second);
        RETURN_IF_ERROR(push_node(stack, n, "PtgArea3d range"));
        break;
      }
      case PtgKind::RefErr:
      case PtgKind::RefErr3d:
      case PtgKind::AreaErr:
      case PtgKind::AreaErr3d: {
        // Consume the payload (ixti for the 3d forms, then one loc, or two
        // for the area forms; the locations are unused) and emit a `#REF!`
        // literal. A 3d form keeps the sheet its XTI names: `Sheet3!#REF!`,
        // or `#REF!#REF!` for a deleted sheet; an external one stays bare.
        const bool area = info->kind == PtgKind::AreaErr || info->kind == PtgKind::AreaErr3d;
        std::optional<std::uint32_t> ixti;
        if (info->kind == PtgKind::RefErr3d || info->kind == PtgKind::AreaErr3d) {
          ASSIGN_OR_RETURN(ixti, read_u16(cursor));
        }
        for (int i = area ? 2 : 1; i > 0; --i) {
          RETURN_IF_ERROR(read_loc(cursor, {}));
        }
        parser::AstNode* n = nullptr;
        if (!ixti || ixti_is_external(*ixti)) {
          n = parser::make_error_literal(arena, ErrorCode::Ref);
        } else if (ixti_is_dead_sheet(*ixti)) {
          n = parser::make_dead_sheet_ref_error(arena, nullptr, nullptr);
        } else if (const std::string_view sheet = sheet_for_ixti(*ixti); !sheet.empty()) {
          n = parser::make_qualified_ref_error(arena, sheet, parser::local_sheet_needs_quoting_a1(sheet));
        } else {
          n = parser::make_error_literal(arena, ErrorCode::Ref);
        }
        RETURN_IF_ERROR(push_node(stack, n, area ? "PtgAreaErr" : "PtgRefErr"));
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
        parser::AstNode* lhs = stack.back();
        stack.pop_back();
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
        const std::uint32_t* mark = find_open_chain(lhs);
        if (mark != nullptr && !chains.is_range(*mark) &&
            parser::binop_binding_power(chains.first_op(*mark)) == parser::binop_binding_power(op)) {
          chains.push_binary(op, rhs);
          stack.push_back(lhs);
          break;
        }
        lhs = close_chain(lhs);
        parser::AstNode* stand_in = parser::make_error_placeholder(arena);
        if (stand_in == nullptr || chain_out_of_memory) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (binary op)", "context=xlsb_ptg_reader");
        }
        stand_ins.push_back(stand_in);
        chain_marks.push_back(chains.mark());
        chains.start_binary(lhs);
        chains.push_binary(op, rhs);
        stack.push_back(stand_in);
        break;
      }

      // ---- Range / set operators -----------------------------------------
      case PtgKind::Range: {
        if (stack.size() < 2) {
          return corrupt_stack("range operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = stack.back();
        stack.pop_back();
        if (const std::uint32_t* mark = find_open_chain(lhs); mark != nullptr && chains.is_range(*mark)) {
          chains.push_range(*rhs);
          stack.push_back(lhs);
          break;
        }
        lhs = close_chain(lhs);
        parser::AstNode* stand_in = parser::make_error_placeholder(arena);
        if (stand_in == nullptr || chain_out_of_memory) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (range op)", "context=xlsb_ptg_reader");
        }
        stand_ins.push_back(stand_in);
        chain_marks.push_back(chains.mark());
        chains.start_range(*lhs);
        chains.push_range(*rhs);
        stack.push_back(stand_in);
        break;
      }
      case PtgKind::Union: {
        if (stack.size() < 2) {
          return corrupt_stack("union operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        // Excel chains a list `(a,b,c)` as `a b UNION c UNION`; an unparenthesised
        // union on the left is the start of the same list.
        const bool extends = lhs->kind() == parser::NodeKind::UnionOp && lhs->paren_depth() == 0U;
        std::vector<const parser::AstNode*> children;
        if (extends) {
          for (std::uint32_t i = 0; i < lhs->as_union_arity(); ++i) {
            children.push_back(&lhs->as_union_child(i));
          }
        } else {
          children.push_back(lhs);
        }
        children.push_back(rhs);
        parser::AstNode* n = parser::make_union_op(arena, children.data(), static_cast<std::uint32_t>(children.size()));
        RETURN_IF_ERROR(push_node(stack, n, "union op"));
        break;
      }
      case PtgKind::Isect: {
        if (stack.size() < 2) {
          return corrupt_stack("intersect operator");
        }
        parser::AstNode* rhs = pop();
        parser::AstNode* lhs = pop();
        parser::AstNode* n = parser::make_intersect_op(arena, lhs, rhs);
        RETURN_IF_ERROR(push_node(stack, n, "intersect op"));
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
        RETURN_IF_ERROR(push_node(stack, n, "unary op"));
        break;
      }
      case PtgKind::Paren: {
        // Evaluation ignores parentheses; the operand records them so the
        // formula text keeps the ones Excel shows.
        if (stack.empty()) {
          return corrupt_stack("paren");
        }
        stack.back() = close_chain(stack.back());
        stack.back()->add_paren();
        break;
      }

      // ---- Functions ------------------------------------------------------
      case PtgKind::Func: {
        ASSIGN_OR_RETURN(auto id, read_u16(cursor));
        const XlsbFuncEntry* entry = lookup_func_by_id(id);
        if (entry == nullptr) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb PtgFunc unknown function id",
                            "context=xlsb_ptg_reader id=" + std::to_string(id));
        }
        // Fixed arity.
        ASSIGN_OR_RETURN(auto call, pop_call(entry->name, entry->arg_min, "PtgFunc", "function (fixed)"));
        stack.push_back(call);
        break;
      }
      case PtgKind::FuncVar: {
        ASSIGN_OR_RETURN(const std::uint32_t cparams, read_u8(cursor));
        ASSIGN_OR_RETURN(auto id, read_u16(cursor));
        // id == 255 is the "future function" sentinel: the real callee
        // is not in the classic function-id table at all (XLOOKUP, LET,
        // TEXTJOIN, CONCAT, IFS, SEQUENCE, ...). Its name was pushed as
        // the FIRST operand via a preceding `PtgName`, so `cparams`
        // counts that name-ref plus the real arguments.
        if (id == 255) {
          ASSIGN_OR_RETURN(auto call, decode_future_function(cparams));
          stack.push_back(call);
          break;
        }
        const XlsbFuncEntry* entry = lookup_func_by_id(id);
        if (entry == nullptr) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb PtgFuncVar unknown function id",
                            "context=xlsb_ptg_reader id=" + std::to_string(id));
        }
        ASSIGN_OR_RETURN(auto call, pop_call(entry->name, cparams, "PtgFuncVar", "function (var)"));
        stack.push_back(call);
        break;
      }

      // ---- Attributes -----------------------------------------------------
      case PtgKind::Attr: {
        ASSIGN_OR_RETURN(auto sub_byte, read_u8(cursor));
        const auto sub = static_cast<PtgAttrKind>(sub_byte);
        switch (sub) {
          case PtgAttrKind::Sum: {
            // Optimised single-argument SUM. The attr carries a u16 of
            // unused data; collapse the top operand into `SUM(x)`.
            RETURN_IF_ERROR(read_u16(cursor));
            if (stack.empty()) {
              return corrupt_stack("attr-sum");
            }
            parser::AstNode* operand = pop();
            const parser::AstNode* args[1] = {operand};
            parser::AstNode* n = parser::make_call(arena, arena.intern("SUM"), args, 1);
            RETURN_IF_ERROR(push_node(stack, n, "attr-sum"));
            break;
          }
          case PtgAttrKind::Space:
          case PtgAttrKind::SpaceSemi: {
            // Whitespace attr: (type, count), applied to the next token. The
            // operand stack is untouched.
            if (cursor.size < 2) {
              return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgAttrSpace payload truncated",
                                "context=xlsb_ptg_reader");
            }
            if (pending_spaces < kMaxPendingSpaces) {
              pending[pending_spaces] = PendingSpace{cursor.data[0], cursor.data[1]};
            }
            ++pending_spaces;
            cursor.data += 2;
            cursor.size -= 2;
            keeps_pending_space = true;
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
              ASSIGN_OR_RETURN(auto count, read_u16(cursor));
              const std::uint32_t entries = static_cast<std::uint32_t>(count) + 1U;
              for (std::uint32_t i = 0; i < entries; ++i) {
                RETURN_IF_ERROR(read_u16(cursor));
              }
            } else {
              RETURN_IF_ERROR(read_u16(cursor));
            }
            // These attrs are control-flow only; they do not consume or
            // produce operands.
            keeps_pending_space = true;
            break;
          }
        }
        break;
      }

      // ---- IFERROR optimisation marker (transparent) ----------------------
      case PtgKind::IfError: {
        // Treated as a no-op marker; the surrounding IFERROR call is
        // reconstructed from its PtgFuncVar. Nothing to read.
        keeps_pending_space = true;
        break;
      }

      default:
        return unsupported_ptg(first_byte, info->name);
    }
    if (pending_spaces != 0 && !keeps_pending_space) {
      if (!stack.empty()) {
        record_space(info->kind, stack.back());
      }
      pending_spaces = 0;
    }
  }

  if (stack.size() != 1) {
    return corrupt_stack(stack.empty() ? "empty stack at end" : "multiple values at end");
  }
  stack.front() = close_chain(stack.front());
  if (chain_out_of_memory) {
    return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (binary op)", "context=xlsb_ptg_reader");
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
