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

// Storage name Excel writes into the BrtName table for a "future"
// (post-2007) function absent from the classic func-id table. Worksheet-
// only dynamic-array functions (the original 2018 set) carry the
// `_xlfn._xlws.` prefix; every other future function carries `_xlfn.`.
// Mirrors ClassifyStoragePrefix in the OOXML writer so both persistence
// formats name the callee identically. The name is upper-cased to match
// Excel's own BrtName spelling byte-for-byte.
std::string future_function_storage_name(std::string_view name) {
  std::string upper;
  upper.reserve(name.size());
  for (char c : name) {
    upper.push_back((c >= 'a' && c <= 'z') ? static_cast<char>(c - 'a' + 'A') : c);
  }
  static constexpr std::string_view kXlwsFunctions[] = {"FILTER", "SORT", "SORTBY", "UNIQUE"};
  for (const std::string_view fn : kXlwsFunctions) {
    if (upper == fn) {
      return std::string("_xlfn._xlws.") + upper;
    }
  }
  return std::string("_xlfn.") + upper;
}

namespace {

constexpr std::uint16_t kColRelBit = 0x4000;
constexpr std::uint16_t kRowRelBit = 0x8000;

// Formula arguments which denote cells and ranges use the reference-class
// base bytes. The result of a function call, in contrast, is a value-class
// Ptg (its low 5-bit type plus class bits `0x40`). Using `| 0x40` on the
// already class-marked base byte had emitted array-class
// references (for example 0x65 instead of PtgArea 0x25), while leaving
// function results in the reference class. Excel repairs those streams.
constexpr std::uint8_t kPtgValueClass = 0x40;
constexpr std::uint8_t kPtgTypeMask = 0x1FU;

constexpr std::uint8_t ValueClassPtg(std::uint8_t reference_class_ptg) {
  return static_cast<std::uint8_t>((reference_class_ptg & kPtgTypeMask) | kPtgValueClass);
}

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

/// Packs a single `RgceArea` corner's column field (14-bit column plus
/// the two relative-flag bits), matching `emit_loc`'s bit layout.
std::uint16_t pack_area_col(const parser::Reference& ref) {
  std::uint16_t col = static_cast<std::uint16_t>(ref.col & 0x3FFF);
  if (!ref.col_abs) {
    col |= kColRelBit;
  }
  if (!ref.row_abs) {
    col |= kRowRelBit;
  }
  return col;
}

/// Emits the `RgceArea` two-corner range coordinate: rows first, then
/// columns — `row1(u32), row2(u32), col1(u16 w/ flags), col2(u16 w/
/// flags)` — NOT two back-to-back `RgceLoc` (`emit_loc`) pairs. Verified
/// against a real Excel-365-produced `xl/worksheets/sheetN.bin` (see
/// `ptg_reader.cpp`'s `read_area`, the decoder counterpart).
void emit_area(std::vector<std::uint8_t>& dst, const parser::Reference& a, const parser::Reference& b) {
  emit_u32(dst, a.row);
  emit_u32(dst, b.row);
  emit_u16(dst, pack_area_col(a));
  emit_u16(dst, pack_area_col(b));
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
  Encoder(const std::vector<std::string>& sheet_names, const SheetRangeTable& sheet_ranges, const NameTable& name_table)
      : sheet_names_(sheet_names), sheet_ranges_(sheet_ranges), name_table_(name_table) {}

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
        return emit_ref3d(node);
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
        return emit_name_ref(node.as_name());
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
        return emit_let(node);
      case parser::NodeKind::LambdaCall:
        return unsupported_node("LambdaCall");
      case parser::NodeKind::ErrorPlaceholder:
        return unsupported_node("ErrorPlaceholder");
    }
    return unsupported_node("unknown");
  }

  EncodedFormula take() { return EncodedFormula{std::move(out_), std::move(extra_)}; }

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

  /// Emits `PtgName` (reference-class) for a defined-name reference,
  /// OR — when `name` matches a LET/LAMBDA parameter currently in scope
  /// (`let_scope_`, innermost binding first) — for that parameter's
  /// hidden `_xlpm.<name>` placeholder. `name_table_` must already
  /// carry every name this encoder is asked to reference — the caller
  /// (`write_xlsb`) builds it from `collect_ptg_names` before encoding
  /// any cell, so a live `NameRef` always resolves.
  Expected<void, Error> emit_name_ref(std::string_view name) {
    for (auto it = let_scope_.rbegin(); it != let_scope_.rend(); ++it) {
      if (it->first == name) {
        emit_u8(out_, 0x23);  // PtgName (reference-class base)
        emit_u32(out_, it->second);
        return Expected<void, Error>::Ok();
      }
    }
    const auto it = name_table_.find(std::string(name));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: name reference not in name table",
                        std::string("context=xlsb_ptg_writer name=") + std::string(name));
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class base)
    emit_u32(out_, it->second);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a `LetBinding` the same way a real Excel-365-produced
  /// `xl/worksheets/sheetN.bin` does: `PtgName(ilbl for "_xlfn.LET")`,
  /// then for each binding `PtgName(ilbl for "_xlpm.<name>")` + the
  /// value expression, then the body, then `PtgFuncVar` with
  /// `id == 255` and `cparams == 1 + 2*n + 1` (LET name-ref + `n`
  /// name/value pairs + body). Verified against real bytes — see
  /// `ptg_reader.cpp`'s LET handling in `decode_future_function`.
  Expected<void, Error> emit_let(const parser::AstNode& node) {
    const auto it_let = name_table_.find("_xlfn.LET");
    if (it_let == name_table_.end()) {
      return unsupported_node("LetBinding(_xlfn.LET not registered)");
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the LET name-ref
    emit_u32(out_, it_let->second);
    const std::uint32_t n = node.as_let_binding_count();
    for (std::uint32_t i = 0; i < n; ++i) {
      const std::string_view raw_name = node.as_let_binding_name(i);
      const std::string param_name = std::string("_xlpm.") + std::string(raw_name);
      const auto it_param = name_table_.find(param_name);
      if (it_param == name_table_.end()) {
        return unsupported_node("LetBinding(param name not registered)");
      }
      emit_u8(out_, 0x23);  // PtgName (reference-class): the binding-name-ref
      emit_u32(out_, it_param->second);
      RETURN_IF_ERROR(emit(node.as_let_binding_expr(i)));
      // Subsequent binding expressions and the body can reference this
      // binding; push it onto scope only after its own value expression
      // has been emitted (a binding cannot reference itself).
      let_scope_.emplace_back(raw_name, it_param->second);
    }
    auto status = emit(node.as_let_body());
    for (std::uint32_t i = 0; i < n; ++i) {
      let_scope_.pop_back();
    }
    RETURN_IF_ERROR(status);
    const std::uint32_t cparams = 1U + (2U * n) + 1U;
    if (cparams > 0xFF) {
      return unsupported_node("LetBinding(too many bindings)");
    }
    emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, static_cast<std::uint8_t>(cparams));
    emit_u16(out_, 255);
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_ref(const parser::Reference& ref) {
    if (ref.is_full_col || ref.is_full_row) {
      // XLSB has no standalone whole-column / whole-row token. Encode the
      // logical extent as an Area using Excel's grid sentinels; this is
      // semantically identical to A:A / 1:1 and, crucially, keeps one
      // unsupported formula from aborting the entire workbook save.
      parser::Reference first = ref;
      parser::Reference last = ref;
      first.is_full_col = false;
      first.is_full_row = false;
      last.is_full_col = false;
      last.is_full_row = false;
      if (ref.is_full_col) {
        first.row = 0;
        last.row = 1048575U;
      } else {
        first.col = 0;
        last.col = 16383U;
      }
      if (ref.sheet.empty()) {
        emit_u8(out_, 0x25);  // PtgArea (reference-class argument)
        emit_area(out_, first, last);
        return Expected<void, Error>::Ok();
      }
      ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_single_sheet_ixti(ref.sheet));
      emit_u8(out_, 0x3B);  // PtgArea3d (reference-class argument)
      emit_u16(out_, ixti);
      emit_area(out_, first, last);
      return Expected<void, Error>::Ok();
    }
    if (ref.sheet.empty()) {
      emit_u8(out_, 0x24);  // PtgRef (reference-class argument)
      emit_loc(out_, ref);
      return Expected<void, Error>::Ok();
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_single_sheet_ixti(ref.sheet));
    emit_u8(out_, 0x3A);  // PtgRef3d (reference-class argument)
    emit_u16(out_, ixti);
    emit_loc(out_, ref);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a `Ref3D` node over a sheet span resolved to `ixti` through
  /// `sheet_ranges_`. A single-cell tail (`Sheet1:Sheet3!A1`) emits
  /// `PtgRef3d(ixti) + RgceLoc`; a range tail (`Sheet1:Sheet3!A1:B2`) emits
  /// `PtgArea3d(ixti) + RgceArea` so the rectangle survives the round-trip.
  Expected<void, Error> emit_ref3d(const parser::AstNode& node) {
    const std::string_view begin = node.as_ref3d_sheet_begin();
    const std::string_view end = node.as_ref3d_sheet_end();
    const int itab_begin = resolve_ixti(sheet_names_, begin);
    const int itab_end = resolve_ixti(sheet_names_, end);
    if (itab_begin < 0 || itab_end < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet not found for 3-D range",
                        std::string("context=xlsb_ptg_writer sheets=") + std::string(begin) + ":" + std::string(end));
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_range_ixti(itab_begin, itab_end));
    // 3-D references only appear as arguments to reference-taking functions
    // (SUM, COUNT, ...), so Excel writes them in the reference class (the
    // base ptg with no value/array class bit). A value/array-class 3-D
    // reference reads back as #REF! in real Excel.
    if (node.as_ref3d_is_range()) {
      emit_u8(out_, 0x3B);  // PtgArea3d (reference class)
      emit_u16(out_, ixti);
      emit_area(out_, node.as_ref3d_cell(), node.as_ref3d_cell_end());
      return Expected<void, Error>::Ok();
    }
    emit_u8(out_, 0x3A);  // PtgRef3d (reference class)
    emit_u16(out_, ixti);
    emit_loc(out_, node.as_ref3d_cell());
    return Expected<void, Error>::Ok();
  }

  /// Resolves `sheet`'s single-sheet-qualified `ixti` (stored in
  /// `sheet_ranges_` as `(itab, itab)`), so single- and multi-sheet
  /// qualified references share the same `ixti` numbering space. See
  /// the `SheetRangeTable` doc comment for why: once the workbook emits
  /// any `BrtExternSheet` entry, the reader interprets every `ixti` as a
  /// table index rather than a direct sheet index, so a single-sheet ref
  /// cannot fall back to a bare index once a genuine 3-D range exists
  /// anywhere in the same workbook.
  Expected<std::uint16_t, Error> resolve_single_sheet_ixti(std::string_view sheet) {
    const int itab = resolve_ixti(sheet_names_, sheet);
    if (itab < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet not found for 3-D ref",
                        std::string("context=xlsb_ptg_writer sheet=") + std::string(sheet));
    }
    return resolve_range_ixti(itab, itab);
  }

  Expected<std::uint16_t, Error> resolve_range_ixti(int itab_first, int itab_last) {
    const int ixti = try_resolve_range_ixti(itab_first, itab_last);
    if (ixti < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet range not pre-registered",
                        "context=xlsb_ptg_writer");
    }
    return static_cast<std::uint16_t>(ixti);
  }

  /// Non-`Expected` variant for callers (the `PtgArea` fast path) that
  /// fall back to a different encoding on a lookup miss rather than
  /// failing outright. Returns -1 when `(itab_first, itab_last)` is not
  /// in `sheet_ranges_`.
  int try_resolve_range_ixti(int itab_first, int itab_last) const {
    for (std::size_t i = 0; i < sheet_ranges_.size(); ++i) {
      if (sheet_ranges_[i].first == itab_first && sheet_ranges_[i].second == itab_last) {
        return static_cast<int>(i);
      }
    }
    return -1;
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
          emit_u8(out_, 0x25);  // PtgArea (reference-class argument)
          emit_area(out_, a, b);
          return Expected<void, Error>::Ok();
        }
        const int itab = resolve_ixti(sheet_names_, a.sheet);
        const int ixti = itab >= 0 ? try_resolve_range_ixti(itab, itab) : -1;
        if (ixti >= 0) {
          emit_u8(out_, 0x3B);  // PtgArea3d (reference-class argument)
          emit_u16(out_, static_cast<std::uint16_t>(ixti));
          emit_area(out_, a, b);
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
      return emit_future_function_call(node, name);
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
      emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
      emit_u8(out_, static_cast<std::uint8_t>(arity));
      emit_u16(out_, entry->id);
    } else {
      if (arity != entry->arg_min) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: fixed-arity function arity mismatch",
                          std::string("context=xlsb_ptg_writer fn=") + std::string(name));
      }
      emit_u8(out_, ValueClassPtg(0x21));  // PtgFunc result
      emit_u16(out_, entry->id);
    }
    return Expected<void, Error>::Ok();
  }

  /// Encodes a call to a function absent from the classic
  /// `func_id_table` as a "future function": `PtgName(ilbl)` naming the
  /// callee, then the real arguments, then `PtgFuncVar` with the
  /// `id == 255` sentinel and `cparams == arity + 1` (the name-ref
  /// counts as an operand). Verified against a real Excel-365-produced
  /// `xl/worksheets/sheetN.bin` for XLOOKUP / TEXTJOIN / CONCAT / IFS /
  /// SEQUENCE — see `ptg_reader.cpp`'s `decode_future_function`, the
  /// decoder counterpart.
  Expected<void, Error> emit_future_function_call(const parser::AstNode& node, std::string_view name) {
    const auto it = name_table_.find(future_function_storage_name(name));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: function has no XLSB id and no future-function name registered",
                        std::string("context=xlsb_ptg_writer fn=") + std::string(name));
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the callee name-ref
    emit_u32(out_, it->second);
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_call_arg(i)));
    }
    const std::uint32_t cparams = arity + 1;  // +1 for the name-ref operand
    if (cparams > 0xFF) {
      return unsupported_node("Call(arity>254, future function)");
    }
    emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, static_cast<std::uint8_t>(cparams));
    emit_u16(out_, 255);
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_array(const parser::AstNode& node) {
    // The main token stream carries only a 15-byte placeholder (opcode
    // + 14 reserved bytes, contents unconstrained); the real dimensions
    // and elements go into `extra_` (this formula's `rgcb`), consumed
    // by the decoder in encounter order. Verified against a real
    // Excel-365-produced `xl/worksheets/sheetN.bin` (see
    // `ptg_reader.cpp`'s `PtgKind::Array` case, the decoder
    // counterpart).
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    emit_u8(out_, 0x60);  // PtgArray (array-class base, matches the decoder)
    for (int i = 0; i < 14; ++i) {
      emit_u8(out_, 0);
    }
    emit_u32(extra_, rows);
    emit_u32(extra_, cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const parser::AstNode& elem = node.as_array_element(r, c);
        if (elem.kind() != parser::NodeKind::Literal || elem.as_literal().kind() != ValueKind::Number) {
          // Only the numeric element tag (`0x00`) has been verified
          // against real Excel output; string / bool / error
          // array-constant elements are not encoded speculatively (see
          // the matching decoder-side note).
          return unsupported_node("ArrayLiteral(non-numeric element)");
        }
        emit_u8(extra_, 0);  // number (verified: tag byte 0x00 precedes the double)
        emit_double(extra_, elem.as_literal().as_number());
      }
    }
    return Expected<void, Error>::Ok();
  }

  const std::vector<std::string>& sheet_names_;
  const SheetRangeTable& sheet_ranges_;
  const NameTable& name_table_;
  std::vector<std::uint8_t> out_;
  /// `rgcb`: the array-constant extra-data area `emit_array` appends to.
  std::vector<std::uint8_t> extra_;
  /// Stack of `(parameter name, ilbl)` pairs currently in scope from an
  /// enclosing `LetBinding` (innermost last). Consulted by
  /// `emit_name_ref` before falling back to `name_table_`.
  std::vector<std::pair<std::string_view, std::uint32_t>> let_scope_;
};

/// Recursion helper for `collect_ptg_names`: adds `name` to `names` (and
/// marks it in `seen`) unless already present.
void AddName(std::string_view name, std::vector<std::string>& names, std::unordered_set<std::string>& seen) {
  std::string owned(name);
  if (seen.insert(owned).second) {
    names.push_back(std::move(owned));
  }
}

/// Packs an `(itabFirst, itabLast)` pair into a single dedupe key.
std::uint64_t PackRangeKey(std::int32_t itab_first, std::int32_t itab_last) {
  return (static_cast<std::uint64_t>(static_cast<std::uint32_t>(itab_first)) << 32) |
         static_cast<std::uint64_t>(static_cast<std::uint32_t>(itab_last));
}

/// Recursion helper for `collect_ptg_sheet_ranges`: resolves `itab_first`
/// / `itab_last` and, when both are valid, appends the pair to `ranges`
/// unless already present in `seen`.
void AddSheetRange(std::int32_t itab_first, std::int32_t itab_last, SheetRangeTable& ranges,
                   std::unordered_set<std::uint64_t>& seen) {
  if (itab_first < 0 || itab_last < 0) {
    return;  // Unresolvable sheet name; the encode fails later with a precise error.
  }
  if (seen.insert(PackRangeKey(itab_first, itab_last)).second) {
    ranges.emplace_back(itab_first, itab_last);
  }
}

/// Recursive worker for `collect_ptg_names` carrying the LET / LAMBDA
/// parameter names currently in scope (innermost last). A `NameRef`
/// matching an in-scope parameter resolves at encode time to that
/// parameter's hidden `_xlpm.<name>` placeholder (see
/// `Encoder::emit_name_ref`), so it must not be registered as an ordinary
/// workbook defined name.
void CollectNamesScoped(const parser::AstNode& node, std::vector<std::string>& names,
                        std::unordered_set<std::string>& seen, std::vector<std::string_view>& scope) {
  switch (node.kind()) {
    case parser::NodeKind::NameRef: {
      const std::string_view name = node.as_name();
      for (const std::string_view param : scope) {
        if (param == name) {
          return;  // LET / LAMBDA parameter: encoded via its _xlpm. placeholder.
        }
      }
      AddName(name, names, seen);
      return;
    }
    case parser::NodeKind::Call: {
      const std::string_view name = node.as_call_name();
      if (lookup_func_by_name(name) == nullptr) {
        AddName(future_function_storage_name(name), names, seen);
      }
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_call_arg(i), names, seen, scope);
      }
      return;
    }
    case parser::NodeKind::UnaryOp:
      CollectNamesScoped(node.as_unary_operand(), names, seen, scope);
      return;
    case parser::NodeKind::BinaryOp:
      CollectNamesScoped(node.as_binary_lhs(), names, seen, scope);
      CollectNamesScoped(node.as_binary_rhs(), names, seen, scope);
      return;
    case parser::NodeKind::RangeOp:
      CollectNamesScoped(node.as_range_lhs(), names, seen, scope);
      CollectNamesScoped(node.as_range_rhs(), names, seen, scope);
      return;
    case parser::NodeKind::IntersectOp:
      CollectNamesScoped(node.as_intersect_lhs(), names, seen, scope);
      CollectNamesScoped(node.as_intersect_rhs(), names, seen, scope);
      return;
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_union_child(i), names, seen, scope);
      }
      return;
    }
    case parser::NodeKind::ImplicitIntersection:
      CollectNamesScoped(node.as_implicit_intersection_operand(), names, seen, scope);
      return;
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          CollectNamesScoped(node.as_array_element(r, c), names, seen, scope);
        }
      }
      return;
    }
    case parser::NodeKind::LetBinding: {
      AddName("_xlfn.LET", names, seen);
      const std::uint32_t n = node.as_let_binding_count();
      const std::size_t scope_base = scope.size();
      for (std::uint32_t i = 0; i < n; ++i) {
        AddName(std::string("_xlpm.") + std::string(node.as_let_binding_name(i)), names, seen);
        // Excel LET binds sequentially: a value expression sees only the
        // earlier bindings, so collect it before pushing this parameter.
        CollectNamesScoped(node.as_let_binding_expr(i), names, seen, scope);
        scope.push_back(node.as_let_binding_name(i));
      }
      CollectNamesScoped(node.as_let_body(), names, seen, scope);
      scope.resize(scope_base);
      return;
    }
    // Leaves, and forms the encoder does not lower (Lambda / LambdaCall /
    // StructuredRef / ExternalRef / SpillRef): nothing to collect. A
    // future writer bundle that lowers these would extend this switch
    // alongside the corresponding `emit_*` case.
    default:
      return;
  }
}

}  // namespace

void collect_ptg_names(const parser::AstNode& node, std::vector<std::string>& names,
                       std::unordered_set<std::string>& seen) {
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, names, seen, scope);
}

void collect_ptg_sheet_ranges(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                              SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen) {
  switch (node.kind()) {
    case parser::NodeKind::Ref: {
      const parser::Reference& r = node.as_ref();
      if (!r.sheet.empty()) {
        const int itab = resolve_ixti(sheet_names, r.sheet);
        AddSheetRange(itab, itab, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::Ref3D: {
      const int itab_begin = resolve_ixti(sheet_names, node.as_ref3d_sheet_begin());
      const int itab_end = resolve_ixti(sheet_names, node.as_ref3d_sheet_end());
      AddSheetRange(itab_begin, itab_end, ranges, seen);
      return;
    }
    case parser::NodeKind::Call: {
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_call_arg(i), sheet_names, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::UnaryOp:
      collect_ptg_sheet_ranges(node.as_unary_operand(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::BinaryOp:
      collect_ptg_sheet_ranges(node.as_binary_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_binary_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::RangeOp:
      collect_ptg_sheet_ranges(node.as_range_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_range_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::IntersectOp:
      collect_ptg_sheet_ranges(node.as_intersect_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_intersect_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_union_child(i), sheet_names, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::ImplicitIntersection:
      collect_ptg_sheet_ranges(node.as_implicit_intersection_operand(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          collect_ptg_sheet_ranges(node.as_array_element(r, c), sheet_names, ranges, seen);
        }
      }
      return;
    }
    case parser::NodeKind::LetBinding: {
      const std::uint32_t n = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < n; ++i) {
        collect_ptg_sheet_ranges(node.as_let_binding_expr(i), sheet_names, ranges, seen);
      }
      collect_ptg_sheet_ranges(node.as_let_body(), sheet_names, ranges, seen);
      return;
    }
    default:
      return;
  }
}

Expected<EncodedFormula, Error> encode_ptgs(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                                            const SheetRangeTable& sheet_ranges, const NameTable& name_table) {
  Encoder enc(sheet_names, sheet_ranges, name_table);
  auto status = enc.emit(node);
  if (!status) {
    return status.error();
  }
  return enc.take();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
