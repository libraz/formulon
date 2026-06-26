// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the AST -> Ptg encoder. See `io/xlsb/ptg_writer.h`.
//
// The encoder recurses post-order. Helper `emit_*` functions append the
// little-endian wire form for each token; the shapes are byte-matched to
// the decoder in `ptg_reader.cpp`.

#include "io/xlsb/ptg_writer.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "parser/reference.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kColRelBit = 0x4000;
constexpr std::uint16_t kRowRelBit = 0x8000;

// Reference-class base bytes for the class-marked operand Ptgs we emit.
// Cell / range references are emitted as value-class (`| 0x40`) so they
// dereference to a value in expression context, which is what Excel does
// for the common scalar formulas in scope here.
constexpr std::uint8_t kPtgValueClass = 0x40;

void emit_u8(std::vector<std::uint8_t>& dst, std::uint8_t v) {
  dst.push_back(v);
}

void emit_u16(std::vector<std::uint8_t>& dst, std::uint16_t v) {
  dst.push_back(static_cast<std::uint8_t>(v & 0xFF));
  dst.push_back(static_cast<std::uint8_t>((v >> 8) & 0xFF));
}

void emit_u32(std::vector<std::uint8_t>& dst, std::uint32_t v) {
  dst.push_back(static_cast<std::uint8_t>(v & 0xFF));
  dst.push_back(static_cast<std::uint8_t>((v >> 8) & 0xFF));
  dst.push_back(static_cast<std::uint8_t>((v >> 16) & 0xFF));
  dst.push_back(static_cast<std::uint8_t>((v >> 24) & 0xFF));
}

void emit_double(std::vector<std::uint8_t>& dst, double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  emit_u32(dst, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  emit_u32(dst, static_cast<std::uint32_t>((bits >> 32) & 0xFFFFFFFFU));
}

/// Emits a u16 code-unit count + UTF-16LE units for `text` (UTF-8 in).
/// Mirrors `read_ptg_string` in the decoder. Code points above the BMP
/// are emitted as a surrogate pair (so the count is in UTF-16 units).
void emit_ptg_string_body(std::vector<std::uint8_t>& dst, std::string_view text) {
  // First pass: decode UTF-8 into UTF-16 code units.
  std::vector<std::uint16_t> units;
  units.reserve(text.size());
  std::size_t i = 0;
  auto cont = [&text](std::size_t idx) -> std::uint32_t {
    return static_cast<std::uint32_t>(static_cast<std::uint8_t>(text[idx]) & 0x3F);
  };
  while (i < text.size()) {
    const auto b0 = static_cast<std::uint8_t>(text[i]);
    std::uint32_t cp = 0;
    std::size_t len = 1;
    if (b0 < 0x80) {
      cp = b0;
      len = 1;
    } else if ((b0 & 0xE0) == 0xC0 && i + 1 < text.size()) {
      cp = (static_cast<std::uint32_t>(b0 & 0x1F) << 6) | cont(i + 1);
      len = 2;
    } else if ((b0 & 0xF0) == 0xE0 && i + 2 < text.size()) {
      cp = (static_cast<std::uint32_t>(b0 & 0x0F) << 12) | (cont(i + 1) << 6) | cont(i + 2);
      len = 3;
    } else if ((b0 & 0xF8) == 0xF0 && i + 3 < text.size()) {
      cp = (static_cast<std::uint32_t>(b0 & 0x07) << 18) | (cont(i + 1) << 12) | (cont(i + 2) << 6) | cont(i + 3);
      len = 4;
    } else {
      cp = 0xFFFD;
      len = 1;
    }
    if (cp <= 0xFFFF) {
      units.push_back(static_cast<std::uint16_t>(cp));
    } else {
      cp -= 0x10000;
      units.push_back(static_cast<std::uint16_t>(0xD800 | (cp >> 10)));
      units.push_back(static_cast<std::uint16_t>(0xDC00 | (cp & 0x3FF)));
    }
    i += len;
  }
  emit_u16(dst, static_cast<std::uint16_t>(units.size()));
  for (std::uint16_t u : units) {
    emit_u16(dst, u);
  }
}

std::uint8_t error_wire_code(ErrorCode e) {
  const std::int32_t code = ooxml_code(e);
  if (code < 0 || code > 0xFF) {
    return 0x09;  // #UNKNOWN!
  }
  return static_cast<std::uint8_t>(code);
}

/// Emits the RgceLoc (u32 row + u16 col with relative-flag bits) for a
/// single-cell reference. The relative bit is *set* when the coordinate
/// is relative (i.e. not `$`-anchored), matching the decoder.
void emit_loc(std::vector<std::uint8_t>& dst, const parser::Reference& ref) {
  emit_u32(dst, ref.row);
  std::uint16_t col = static_cast<std::uint16_t>(ref.col & 0x3FFF);
  if (!ref.col_abs) {
    col |= kColRelBit;
  }
  if (!ref.row_abs) {
    col |= kRowRelBit;
  }
  emit_u16(dst, col);
}

Error unsupported_node(const char* kind) {
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                    std::string("xlsb encoder cannot lower AST node kind: ") + kind, "context=xlsb_ptg_writer");
}

/// Resolves a sheet name to its 0-based ixti. Returns -1 when absent.
int resolve_ixti(const std::vector<std::string>& sheet_names, std::string_view sheet) {
  for (std::size_t i = 0; i < sheet_names.size(); ++i) {
    if (sheet_names[i] == sheet) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

class Encoder {
 public:
  explicit Encoder(const std::vector<std::string>& sheet_names) : sheet_names_(sheet_names) {}

  Expected<void, Error> emit(const parser::AstNode& node) {
    switch (node.kind()) {
      case parser::NodeKind::Literal:
        return emit_literal(node.as_literal());
      case parser::NodeKind::ErrorLiteral:
        emit_u8(out_, 0x1C);  // PtgErr
        emit_u8(out_, error_wire_code(node.as_error_literal()));
        return Expected<void, Error>::Ok();
      case parser::NodeKind::Ref:
        return emit_ref(node.as_ref());
      case parser::NodeKind::Ref3D:
        // A node-level 3-D ref spans a run of sheets; only single-sheet
        // qualified refs round-trip through the scalar Ptg forms here.
        return unsupported_node("Ref3D");
      case parser::NodeKind::UnaryOp:
        return emit_unary(node);
      case parser::NodeKind::BinaryOp:
        return emit_binary(node);
      case parser::NodeKind::RangeOp:
        return emit_range(node);
      case parser::NodeKind::UnionOp:
        return emit_union(node);
      case parser::NodeKind::IntersectOp:
        return emit_intersect(node);
      case parser::NodeKind::Call:
        return emit_call(node);
      case parser::NodeKind::ArrayLiteral:
        return emit_array(node);
      case parser::NodeKind::NameRef:
        return unsupported_node("NameRef");
      case parser::NodeKind::ExternalRef:
        return unsupported_node("ExternalRef");
      case parser::NodeKind::StructuredRef:
        return unsupported_node("StructuredRef");
      case parser::NodeKind::SpillRef:
        return unsupported_node("SpillRef");
      case parser::NodeKind::ImplicitIntersection:
        return unsupported_node("ImplicitIntersection");
      case parser::NodeKind::Lambda:
        return unsupported_node("Lambda");
      case parser::NodeKind::LetBinding:
        return unsupported_node("LetBinding");
      case parser::NodeKind::LambdaCall:
        return unsupported_node("LambdaCall");
      case parser::NodeKind::ErrorPlaceholder:
        return unsupported_node("ErrorPlaceholder");
    }
    return unsupported_node("unknown");
  }

  std::vector<std::uint8_t> take() { return std::move(out_); }

 private:
  Expected<void, Error> emit_literal(const Value& v) {
    switch (v.kind()) {
      case ValueKind::Blank:
        emit_u8(out_, 0x16);  // PtgMissArg
        return Expected<void, Error>::Ok();
      case ValueKind::Number: {
        const double d = v.as_number();
        // Prefer PtgInt for small non-negative integers.
        if (d >= 0.0 && d <= 65535.0 && d == static_cast<double>(static_cast<std::uint16_t>(d))) {
          emit_u8(out_, 0x1E);  // PtgInt
          emit_u16(out_, static_cast<std::uint16_t>(d));
        } else {
          emit_u8(out_, 0x1F);  // PtgNum
          emit_double(out_, d);
        }
        return Expected<void, Error>::Ok();
      }
      case ValueKind::Bool:
        emit_u8(out_, 0x1D);  // PtgBool
        emit_u8(out_, v.as_boolean() ? 1U : 0U);
        return Expected<void, Error>::Ok();
      case ValueKind::Text:
        emit_u8(out_, 0x17);  // PtgStr
        emit_ptg_string_body(out_, v.as_text());
        return Expected<void, Error>::Ok();
      case ValueKind::Error:
        emit_u8(out_, 0x1C);  // PtgErr
        emit_u8(out_, error_wire_code(v.as_error()));
        return Expected<void, Error>::Ok();
      case ValueKind::Array:
      case ValueKind::Lambda:
      case ValueKind::Ref:
        return unsupported_node("Literal(non-scalar)");
    }
    return unsupported_node("Literal(unknown)");
  }

  Expected<void, Error> emit_ref(const parser::Reference& ref) {
    if (ref.is_full_col || ref.is_full_row) {
      // Whole-column / whole-row refs need the Area form with sentinel
      // bounds; out of scope for the common-token codec.
      return unsupported_node("Ref(full-col/row)");
    }
    if (ref.sheet.empty()) {
      emit_u8(out_, 0x24 | kPtgValueClass);  // PtgRef
      emit_loc(out_, ref);
    } else {
      const int ixti = resolve_ixti(sheet_names_, ref.sheet);
      if (ixti < 0) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet not found for 3-D ref",
                          std::string("context=xlsb_ptg_writer sheet=") + std::string(ref.sheet));
      }
      emit_u8(out_, 0x3A | kPtgValueClass);  // PtgRef3d
      emit_u16(out_, static_cast<std::uint16_t>(ixti));
      emit_loc(out_, ref);
    }
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_unary(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit(node.as_unary_operand()));
    switch (node.as_unary_op()) {
      case parser::UnaryOp::Plus:
        emit_u8(out_, 0x12);  // PtgUplus
        break;
      case parser::UnaryOp::Minus:
        emit_u8(out_, 0x13);  // PtgUminus
        break;
      case parser::UnaryOp::Percent:
        emit_u8(out_, 0x14);  // PtgPercent
        break;
    }
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_binary(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit(node.as_binary_lhs()));
    RETURN_IF_ERROR(emit(node.as_binary_rhs()));
    std::uint8_t byte = 0x03;
    switch (node.as_binary_op()) {
      case parser::BinOp::Add:
        byte = 0x03;
        break;
      case parser::BinOp::Sub:
        byte = 0x04;
        break;
      case parser::BinOp::Mul:
        byte = 0x05;
        break;
      case parser::BinOp::Div:
        byte = 0x06;
        break;
      case parser::BinOp::Pow:
        byte = 0x07;
        break;
      case parser::BinOp::Concat:
        byte = 0x08;
        break;
      case parser::BinOp::Lt:
        byte = 0x09;
        break;
      case parser::BinOp::LtEq:
        byte = 0x0A;
        break;
      case parser::BinOp::Eq:
        byte = 0x0B;
        break;
      case parser::BinOp::GtEq:
        byte = 0x0C;
        break;
      case parser::BinOp::Gt:
        byte = 0x0D;
        break;
      case parser::BinOp::NotEq:
        byte = 0x0E;
        break;
    }
    emit_u8(out_, byte);
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_range(const parser::AstNode& node) {
    // Fast path: a range whose endpoints are both plain cell refs maps
    // to PtgArea / PtgArea3d (a single operand) rather than two refs +
    // the `:` operator. The decoder produces a RangeOp of two refs, so
    // either form round-trips; we emit the compact Area form.
    const parser::AstNode& lhs = node.as_range_lhs();
    const parser::AstNode& rhs = node.as_range_rhs();
    if (lhs.kind() == parser::NodeKind::Ref && rhs.kind() == parser::NodeKind::Ref) {
      const parser::Reference& a = lhs.as_ref();
      const parser::Reference& b = rhs.as_ref();
      if (!a.is_full_col && !a.is_full_row && !b.is_full_col && !b.is_full_row && b.sheet.empty()) {
        if (a.sheet.empty()) {
          emit_u8(out_, 0x25 | kPtgValueClass);  // PtgArea
          emit_loc(out_, a);
          emit_loc(out_, b);
          return Expected<void, Error>::Ok();
        }
        const int ixti = resolve_ixti(sheet_names_, a.sheet);
        if (ixti >= 0) {
          emit_u8(out_, 0x3B | kPtgValueClass);  // PtgArea3d
          emit_u16(out_, static_cast<std::uint16_t>(ixti));
          emit_loc(out_, a);
          emit_loc(out_, b);
          return Expected<void, Error>::Ok();
        }
      }
    }
    // General form: emit both operands then the `:` operator.
    RETURN_IF_ERROR(emit(lhs));
    RETURN_IF_ERROR(emit(rhs));
    emit_u8(out_, 0x11);  // PtgRange
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_union(const parser::AstNode& node) {
    const std::uint32_t arity = node.as_union_arity();
    if (arity < 2) {
      return unsupported_node("UnionOp(arity<2)");
    }
    // Emit left-associated: (((a,b),c),d) so each `,` pops exactly two.
    RETURN_IF_ERROR(emit(node.as_union_child(0)));
    for (std::uint32_t i = 1; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_union_child(i)));
      emit_u8(out_, 0x10);  // PtgUnion
    }
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_intersect(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit(node.as_intersect_lhs()));
    RETURN_IF_ERROR(emit(node.as_intersect_rhs()));
    emit_u8(out_, 0x0F);  // PtgIsect
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_call(const parser::AstNode& node) {
    const std::string_view name = node.as_call_name();
    const XlsbFuncEntry* entry = lookup_func_by_name(name);
    if (entry == nullptr) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: function has no XLSB id",
                        std::string("context=xlsb_ptg_writer fn=") + std::string(name));
    }
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_call_arg(i)));
    }
    const bool use_var = entry->variadic || entry->arg_min != entry->arg_max;
    if (use_var) {
      if (arity > 0xFF) {
        return unsupported_node("Call(arity>255)");
      }
      emit_u8(out_, 0x22);  // PtgFuncVar
      emit_u8(out_, static_cast<std::uint8_t>(arity));
      emit_u16(out_, entry->id);
    } else {
      if (arity != entry->arg_min) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: fixed-arity function arity mismatch",
                          std::string("context=xlsb_ptg_writer fn=") + std::string(name));
      }
      emit_u8(out_, 0x21);  // PtgFunc
      emit_u16(out_, entry->id);
    }
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_array(const parser::AstNode& node) {
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    emit_u8(out_, 0x20);  // PtgArray (reference-class base)
    emit_u32(out_, rows);
    emit_u32(out_, cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const parser::AstNode& elem = node.as_array_element(r, c);
        if (elem.kind() == parser::NodeKind::ErrorLiteral) {
          emit_u8(out_, 16);
          emit_u8(out_, error_wire_code(elem.as_error_literal()));
          continue;
        }
        if (elem.kind() != parser::NodeKind::Literal) {
          return unsupported_node("ArrayLiteral(non-constant element)");
        }
        const Value& v = elem.as_literal();
        switch (v.kind()) {
          case ValueKind::Number:
            emit_u8(out_, 1);
            emit_double(out_, v.as_number());
            break;
          case ValueKind::Text:
            emit_u8(out_, 2);
            emit_ptg_string_body(out_, v.as_text());
            break;
          case ValueKind::Bool:
            emit_u8(out_, 4);
            emit_u8(out_, v.as_boolean() ? 1U : 0U);
            break;
          case ValueKind::Error:
            emit_u8(out_, 16);
            emit_u8(out_, error_wire_code(v.as_error()));
            break;
          default:
            return unsupported_node("ArrayLiteral(non-scalar element)");
        }
      }
    }
    return Expected<void, Error>::Ok();
  }

  const std::vector<std::string>& sheet_names_;
  std::vector<std::uint8_t> out_;
};

}  // namespace

Expected<std::vector<std::uint8_t>, Error> encode_ptgs(const parser::AstNode& node,
                                                       const std::vector<std::string>& sheet_names) {
  Encoder enc(sheet_names);
  auto status = enc.emit(node);
  if (!status) {
    return status.error();
  }
  return enc.take();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
