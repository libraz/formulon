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

// Mem Ptgs encode the byte length of the following binary-reference
// expression. That expression remains in `cursor` and is decoded normally;
// validating the advertised bound catches a malformed cache marker without
// skipping the actual formula.
Expected<void, Error> validate_mem_expression_size(std::uint16_t cce, const ByteSpan& cursor, const char* ptg_name) {
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
  parser::Reference ref;
  ref.sheet = sheet;
  ref.row = row_or.value();
  ref.col = static_cast<std::uint32_t>(col_or.value() & kColMask);
  ref.col_abs = (col_or.value() & kColRelBit) == 0;
  ref.row_abs = (col_or.value() & kRowRelBit) == 0;
  return ref;
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
  parser::Reference first;
  first.sheet = sheet_first;
  first.row = row1_or.value();
  first.col = static_cast<std::uint32_t>(col1_or.value() & kColMask);
  first.col_abs = (col1_or.value() & kColRelBit) == 0;
  first.row_abs = (col1_or.value() & kRowRelBit) == 0;
  parser::Reference last;
  last.sheet = sheet_last;
  last.row = row2_or.value();
  last.col = static_cast<std::uint32_t>(col2_or.value() & kColMask);
  last.col_abs = (col2_or.value() & kColRelBit) == 0;
  last.row_abs = (col2_or.value() & kRowRelBit) == 0;
  return std::make_pair(first, last);
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

/// Case-insensitive `s` starts-with `prefix` check (ASCII-fold).
bool starts_with_ci(std::string_view s, std::string_view prefix) {
  return s.size() >= prefix.size() && strings::case_insensitive_eq(s.substr(0, prefix.size()), prefix);
}

/// Returns true when `sheet` must be single-quoted to round-trip as a
/// qualified-reference prefix (`'sheet'!ref` rather than `sheet!ref`).
///
/// XLSB's `PtgRef3d` / `PtgArea3d` carry the sheet only as an `ixti` table
/// index -- there is no "was this quoted in the formula bar" bit to
/// preserve, so the decoder must re-derive Excel's own quoting rule from
/// the sheet name text. Two cases require quoting:
///
///   * The name contains a character outside the bare-identifier set
///     (matches the heuristic `parser::ast_format.cpp`'s
///     `AppendSheetNameQuoted` already applies to genuine `Ref3D` nodes).
///   * The name has the exact shape of an A1 cell reference
///     (`[A-Za-z]{1,3}[1-9][0-9]*`, e.g. `S2`), which is otherwise
///     indistinguishable from a cell reference when unquoted -- the
///     tokenizer resolves `S2` to a `CellRef` token before a qualifying
///     `!` can reclassify it. Verified against a real Excel-365-produced
///     package: a sheet literally named `S2` serialises its cross-sheet
///     formula as `'S2'!A1*2` in `xl/worksheets/sheet1.xml`, never the
///     bare form.
bool NeedsSheetQuote(std::string_view sheet) {
  if (sheet.empty()) {
    return true;
  }
  for (const char c : sheet) {
    const bool ok = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.';
    if (!ok) {
      return true;
    }
  }
  std::size_t i = 0;
  while (i < sheet.size() && ((sheet[i] >= 'A' && sheet[i] <= 'Z') || (sheet[i] >= 'a' && sheet[i] <= 'z'))) {
    ++i;
  }
  const std::size_t letters = i;
  if (letters == 0 || letters > 3 || i >= sheet.size()) {
    return false;
  }
  if (sheet[i] < '1' || sheet[i] > '9') {
    return false;
  }
  ++i;
  while (i < sheet.size() && sheet[i] >= '0' && sheet[i] <= '9') {
    ++i;
  }
  return i == sheet.size();
}

}  // namespace

Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, ByteSpan rgcb, Arena& arena,
                                              const std::vector<std::string>& sheet_names,
                                              const std::vector<XlsbName>& name_table,
                                              const std::vector<XlsbSheetRange>& sheet_ranges) {
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
    if (ops[0]->kind() != parser::NodeKind::NameRef) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt,
                        "xlsb PtgFuncVar(255): callee operand is not a name reference", "context=xlsb_ptg_reader");
    }
    const std::string_view callee = ops[0]->as_name();
    const std::uint32_t real_count = cparams - 1;
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
    parser::AstNode* n = parser::make_call(arena, arena.intern(callee), args, real_count);
    if (n == nullptr) {
      return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (future-function call)",
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
        auto cce_or = read_u16(cursor);
        if (!cce_or) {
          return cce_or.error();
        }
        auto size_check = validate_mem_expression_size(cce_or.value(), cursor, info->name);
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
        auto cce_or = read_u16(cursor);
        if (!cce_or) {
          return cce_or.error();
        }
        auto size_check = validate_mem_expression_size(cce_or.value(), cursor, info->name);
        if (!size_check) {
          return size_check.error();
        }
        break;
      }
      case PtgKind::MemFunc: {
        auto cce_or = read_u16(cursor);
        if (!cce_or) {
          return cce_or.error();
        }
        auto size_check = validate_mem_expression_size(cce_or.value(), cursor, info->name);
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
        parser::AstNode* n = parser::make_literal(arena, Value::number(v));
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
        // Field order (first u32 = rows, second u32 = cols) and element
        // consumption order (row-major: row 0 left-to-right, then row 1,
        // ...) cannot be distinguished from the square 2x2 real-Excel
        // fixture alone (`xlsb_fidelity_base.xlsb`'s `=SUM({1,2;3,4})`
        // pins element *order* but not which dimension word is which for
        // a square array). The layout below is what [MS-XLSB] 2.5.98.26
        // specifies and what independent third-party decoders of the same
        // record agree on; a non-square array constant produced by Excel
        // would pin it directly, and the fixture corpus has none.
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
            default:
              // Only the numeric element tag has been verified against
              // real Excel output; string / bool / error array-constant
              // elements are not decoded speculatively. Surfacing as
              // unsupported preserves the cell's cached value instead
              // of risking a silently wrong array.
              return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                                "PtgArray element tag not decoded (only numeric elements are verified)",
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
        // `_xlpm.`-prefixed names are LET/LAMBDA-local parameter
        // references; strip the storage prefix at every use site (not
        // just the LET binding-name slot decoded_future_function
        // handles) so a bare `x*3` reference inside a LET body matches
        // the plain-identifier `NameRef` the text parser would have
        // produced for the same formula.
        const std::string_view display_name = starts_with_ci(name, "_xlpm.") ? name.substr(6) : name;
        parser::AstNode* n = parser::make_name_ref(arena, arena.intern(display_name));
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgName)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }

      // ---- References -----------------------------------------------------
      case PtgKind::Ref: {
        auto ref_or = read_loc(cursor, {});
        if (!ref_or) {
          return ref_or.error();
        }
        auto domain_or = check_ref_domain(ref_or.value(), "PtgRef");
        if (!domain_or) {
          return domain_or.error();
        }
        parser::AstNode* n = parser::make_ref(arena, ref_or.value());
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRef)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::Area: {
        auto area_or = read_area(cursor, {}, {});
        if (!area_or) {
          return area_or.error();
        }
        auto domain_or = check_area_domain(area_or.value().first, area_or.value().second, "PtgArea");
        if (!domain_or) {
          return domain_or.error();
        }
        parser::AstNode* lhs = parser::make_ref(arena, area_or.value().first);
        parser::AstNode* rhs = parser::make_ref(arena, area_or.value().second);
        if (lhs == nullptr || rhs == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea)", "context=xlsb_ptg_reader");
        }
        parser::AstNode* n = parser::make_range_op(arena, lhs, rhs);
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
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti_or.value(), begin_sheet, end_sheet)) {
          auto loc_or = read_loc(cursor, {});
          if (!loc_or) {
            return loc_or.error();
          }
          auto domain_or = check_ref_domain(loc_or.value(), "PtgRef3d");
          if (!domain_or) {
            return domain_or.error();
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
        auto ref_or = read_loc(cursor, sheet_for_ixti(ixti_or.value()));
        if (!ref_or) {
          return ref_or.error();
        }
        auto domain_or = check_ref_domain(ref_or.value(), "PtgRef3d");
        if (!domain_or) {
          return domain_or.error();
        }
        parser::Reference ref = ref_or.value();
        ref.sheet = arena.intern(ref.sheet);
        ref.sheet_quoted = NeedsSheetQuote(ref.sheet);
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
        std::string_view begin_sheet;
        std::string_view end_sheet;
        if (sheet_range_for_ixti(ixti_or.value(), begin_sheet, end_sheet)) {
          auto area_or = read_area(cursor, {}, {});
          if (!area_or) {
            return area_or.error();
          }
          auto domain_or = check_area_domain(area_or.value().first, area_or.value().second, "PtgArea3d");
          if (!domain_or) {
            return domain_or.error();
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
        auto area_or = read_area(cursor, sheet, {});
        if (!area_or) {
          return area_or.error();
        }
        auto domain_or = check_area_domain(area_or.value().first, area_or.value().second, "PtgArea3d");
        if (!domain_or) {
          return domain_or.error();
        }
        parser::Reference first = area_or.value().first;
        first.sheet = arena.intern(first.sheet);
        first.sheet_quoted = NeedsSheetQuote(first.sheet);
        parser::AstNode* lhs = parser::make_ref(arena, first);
        parser::AstNode* rhs = parser::make_ref(arena, area_or.value().second);
        if (lhs == nullptr || rhs == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea3d)", "context=xlsb_ptg_reader");
        }
        parser::AstNode* n = parser::make_range_op(arena, lhs, rhs);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgArea3d range)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::RefErr:
      case PtgKind::RefErr3d: {
        // Consume the payload (ixti for the 3d form, then the loc) and
        // emit a `#REF!` literal.
        if (info->kind == PtgKind::RefErr3d) {
          auto ixti_or = read_u16(cursor);
          if (!ixti_or) {
            return ixti_or.error();
          }
        }
        auto skip = read_loc(cursor, {});
        if (!skip) {
          return skip.error();
        }
        parser::AstNode* n = parser::make_error_literal(arena, ErrorCode::Ref);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgRefErr)", "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
        break;
      }
      case PtgKind::AreaErr:
      case PtgKind::AreaErr3d: {
        if (info->kind == PtgKind::AreaErr3d) {
          auto ixti_or = read_u16(cursor);
          if (!ixti_or) {
            return ixti_or.error();
          }
        }
        auto skip1 = read_loc(cursor, {});
        if (!skip1) {
          return skip1.error();
        }
        auto skip2 = read_loc(cursor, {});
        if (!skip2) {
          return skip2.error();
        }
        parser::AstNode* n = parser::make_error_literal(arena, ErrorCode::Ref);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgAreaErr)", "context=xlsb_ptg_reader");
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
        // Parentheses are structurally transparent: the AST formatter
        // re-inserts whatever parens precedence requires. Keep the
        // operand as-is.
        if (stack.empty()) {
          return corrupt_stack("paren");
        }
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
        const std::uint32_t arity = entry->arg_min;  // fixed arity
        if (stack.size() < arity) {
          return corrupt_stack("function (fixed)");
        }
        auto** args = arity == 0 ? nullptr : arena.create_array<const parser::AstNode*>(arity);
        if (arity != 0 && args == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgFunc args)",
                            "context=xlsb_ptg_reader");
        }
        for (std::uint32_t i = 0; i < arity; ++i) {
          args[arity - 1 - i] = pop();
        }
        parser::AstNode* n = parser::make_call(arena, arena.intern(entry->name), args, arity);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgFunc call)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
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
        const std::uint32_t arity = cparams;
        if (stack.size() < arity) {
          return corrupt_stack("function (var)");
        }
        auto** args = arity == 0 ? nullptr : arena.create_array<const parser::AstNode*>(arity);
        if (arity != 0 && args == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgFuncVar args)",
                            "context=xlsb_ptg_reader");
        }
        for (std::uint32_t i = 0; i < arity; ++i) {
          args[arity - 1 - i] = pop();
        }
        parser::AstNode* n = parser::make_call(arena, arena.intern(entry->name), args, arity);
        if (n == nullptr) {
          return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgFuncVar call)",
                            "context=xlsb_ptg_reader");
        }
        stack.push_back(n);
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
