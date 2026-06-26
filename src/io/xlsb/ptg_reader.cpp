// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

/// Maps an MS-XLSB error wire code to the engine `ErrorCode`. Mirrors the
/// inverse table the reader uses for `BrtCellError`.
ErrorCode error_from_wire(std::uint8_t code) {
  switch (code) {
    case 0:
      return ErrorCode::Null;
    case 7:
      return ErrorCode::Div0;
    case 15:
      return ErrorCode::Value;
    case 23:
      return ErrorCode::Ref;
    case 29:
      return ErrorCode::Name;
    case 36:
      return ErrorCode::Num;
    case 42:
      return ErrorCode::NA;
    case 43:
      return ErrorCode::GettingData;
    default:
      return ErrorCode::Unknown;
  }
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

}  // namespace

Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, Arena& arena,
                                              const std::vector<std::string>& sheet_names) {
  std::vector<parser::AstNode*> stack;

  auto pop = [&stack]() -> parser::AstNode* {
    parser::AstNode* n = stack.back();
    stack.pop_back();
    return n;
  };

  auto sheet_for_ixti = [&sheet_names](std::uint32_t ixti) -> std::string_view {
    if (ixti < sheet_names.size()) {
      return sheet_names[ixti];
    }
    return {};
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
        // The class byte already determines this is a constant-array
        // operand. The XLSB rgce form stores the dimensions inline:
        //   u32 rows (1-based count - 1 in some encodings); we use the
        //   engine's matched writer form: u32 rows, u32 cols, then
        //   row-major elements each tagged with a 1-byte kind.
        auto rows_or = read_u32(cursor);
        if (!rows_or) {
          return rows_or.error();
        }
        auto cols_or = read_u32(cursor);
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
          auto tag_or = read_u8(cursor);
          if (!tag_or) {
            return tag_or.error();
          }
          parser::AstNode* elem = nullptr;
          switch (tag_or.value()) {
            case 1: {  // number
              if (cursor.size < 8) {
                return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgArray number truncated",
                                  "context=xlsb_ptg_reader");
              }
              double v;
              std::memcpy(&v, cursor.data, sizeof(v));
              cursor.data += 8;
              cursor.size -= 8;
              elem = parser::make_literal(arena, Value::number(v));
              break;
            }
            case 2: {  // string
              auto s_or = read_ptg_string(cursor);
              if (!s_or) {
                return s_or.error();
              }
              elem = parser::make_literal(arena, Value::text(arena.intern(s_or.value())));
              break;
            }
            case 4: {  // bool
              auto b_or = read_u8(cursor);
              if (!b_or) {
                return b_or.error();
              }
              elem = parser::make_literal(arena, Value::boolean(b_or.value() != 0));
              break;
            }
            case 16: {  // error
              auto e_or = read_u8(cursor);
              if (!e_or) {
                return e_or.error();
              }
              elem = parser::make_error_literal(arena, error_from_wire(e_or.value()));
              break;
            }
            default:
              return make_error(FormulonErrorCode::kIoXlsbCorrupt, "PtgArray unknown element tag",
                                "context=xlsb_ptg_reader");
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

      // ---- References -----------------------------------------------------
      case PtgKind::Ref: {
        auto ref_or = read_loc(cursor, {});
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
      case PtgKind::Area: {
        auto first_or = read_loc(cursor, {});
        if (!first_or) {
          return first_or.error();
        }
        auto last_or = read_loc(cursor, {});
        if (!last_or) {
          return last_or.error();
        }
        parser::AstNode* lhs = parser::make_ref(arena, first_or.value());
        parser::AstNode* rhs = parser::make_ref(arena, last_or.value());
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
        auto ref_or = read_loc(cursor, sheet_for_ixti(ixti_or.value()));
        if (!ref_or) {
          return ref_or.error();
        }
        parser::Reference ref = ref_or.value();
        ref.sheet = arena.intern(ref.sheet);
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
        const std::string_view sheet = sheet_for_ixti(ixti_or.value());
        auto first_or = read_loc(cursor, sheet);
        if (!first_or) {
          return first_or.error();
        }
        auto last_or = read_loc(cursor, {});
        if (!last_or) {
          return last_or.error();
        }
        parser::Reference first = first_or.value();
        first.sheet = arena.intern(first.sheet);
        parser::AstNode* lhs = parser::make_ref(arena, first);
        parser::AstNode* rhs = parser::make_ref(arena, last_or.value());
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
        const XlsbFuncEntry* entry = lookup_func_by_id(id_or.value());
        if (entry == nullptr) {
          return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb PtgFuncVar unknown function id",
                            "context=xlsb_ptg_reader id=" + std::to_string(id_or.value()));
        }
        const std::uint32_t arity = cparams_or.value();
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
            // jump table (Choose: u16 count + (count+1) u32 offsets).
            if (sub == PtgAttrKind::Choose) {
              auto count_or = read_u16(cursor);
              if (!count_or) {
                return count_or.error();
              }
              const std::uint32_t entries = static_cast<std::uint32_t>(count_or.value()) + 1U;
              for (std::uint32_t i = 0; i < entries; ++i) {
                auto off_or = read_u32(cursor);
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
  return stack.front();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
