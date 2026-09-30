//
// Implementation of the AST -> Ptg encoder. See `io/xlsb/ptg_writer.h`.
//
// The encoder recurses post-order. Helper `emit_*` functions append the
// little-endian wire form for each token; the shapes are byte-matched to
// the decoder in `ptg_reader.cpp`.

#include "io/xlsb/ptg_writer.h"

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "parser/ast_format.h"
#include "parser/reference.h"
#include "sheet_name.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// True when Excel stores a call to `name` under a hidden `_xlfn.*`
/// BrtName rather than a function id. Classified by
/// `io::classify_storage_prefix`, the same enumeration the OOXML writer
/// consults, so the two formats name the callee identically.
bool IsFutureFunction(std::string_view name) {
  return classify_storage_prefix(name) != parser::StoragePrefixKind::None;
}

/// True when a call to `name` is encoded through the hidden-name route
/// (`PtgName` + `PtgFuncVar(255)`) rather than a function id: either a
/// `_xlfn.*` future function, or one of the enumerated names Excel
/// spells bare in OOXML yet has no function id for. Both sets are
/// enumerated from observed Excel output — "absent from
/// `func_id_table`" deliberately does not qualify, because that table
/// grows incrementally and a hidden `_xlfn.<NAME>` Excel does not know
/// resolves to `#NAME?`.
bool UsesHiddenNameRoute(std::string_view name) {
  return IsFutureFunction(name) || xlsb_uses_hidden_name(name);
}

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

constexpr std::uint8_t kPtgReferenceClass = 0x20;
constexpr std::uint8_t kPtgArrayClass = 0x60;

constexpr std::uint8_t ValueClassPtg(std::uint8_t reference_class_ptg) {
  return static_cast<std::uint8_t>((reference_class_ptg & kPtgTypeMask) | kPtgValueClass);
}

/// `reference_class_ptg` carrying the class bits `cls` instead.
constexpr std::uint8_t ClassedPtg(std::uint8_t reference_class_ptg, std::uint8_t cls) {
  return static_cast<std::uint8_t>((reference_class_ptg & kPtgTypeMask) | cls);
}

/// Where a node sits in a formula: the parameter-class letter
/// (`xlsb_parameter_class`) of the built-in argument it is, or is inside,
/// 0 elsewhere; `operand` when it is an operator's operand there.
struct Slot {
  char letter = 0;
  bool operand = false;
};

/// Class bits Excel gives a cell (`area == false`) or area reference at
/// `slot` (measured per parameter; see `xlsb_parameter_class`). Outside
/// every built-in an operand is value class and a direct argument keeps
/// reference class.
std::uint8_t SlotClass(Slot slot, bool area, PtgEvaluation evaluation = PtgEvaluation::kDynamicArray) {
  if (slot.operand && (evaluation == PtgEvaluation::kLegacy || evaluation == PtgEvaluation::kLegacyArray)) {
    // A legacy formula intersects an operand unless the parameter forces arrays.
    return slot.letter == 'F' ? kPtgArrayClass : kPtgValueClass;
  }
  if (slot.operand && evaluation == PtgEvaluation::kConditionalFormat && slot.letter != 0 && slot.letter != 'V') {
    return kPtgArrayClass;
  }
  if (slot.operand) {
    switch (slot.letter) {
      case 'S':
      case 'I':
      case 'Y':
        return area ? kPtgArrayClass : kPtgValueClass;
      case 'F':
      case 'X':
        return kPtgArrayClass;
      default:
        return kPtgValueClass;
    }
  }
  switch (slot.letter) {
    case 'V':
      return kPtgValueClass;
    case 'F':
      return kPtgArrayClass;
    case 'I':
      return area ? kPtgArrayClass : kPtgValueClass;
    case 'Y':
      return area ? kPtgValueClass : kPtgReferenceClass;
    default:
      return kPtgReferenceClass;
  }
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

/// True when either axis of `ref` is relative.
bool IsRelative(const parser::Reference& ref) {
  return !ref.row_abs || !ref.col_abs;
}

/// `ref` with each relative axis replaced by its offset from `base` modulo
/// the grid (2^20 rows, 2^14 columns), the `PtgRefN` / `PtgAreaN` field.
parser::Reference OffsetFrom(parser::Reference ref, PtgBaseCell base) {
  if (!ref.row_abs) {
    ref.row = (ref.row - base.row) & ((1U << 20) - 1U);
  }
  if (!ref.col_abs) {
    ref.col = (ref.col - base.col) & 0x3FFFU;
  }
  return ref;
}

/// `itabFirst` / `itabLast` of the `BrtExternSheet` entry a `PtgNameX`
/// naming one of this workbook's own defined names resolves through: the
/// entry names no sheet, the record's own scope does.
constexpr std::int32_t kXtiNoSheet = -2;

Error unsupported_node(const char* kind) {
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                    std::string("xlsb encoder cannot lower AST node kind: ") + kind, "context=xlsb_ptg_writer");
}

/// Resolves a sheet name to its 0-based ixti. Returns -1 when absent.
int resolve_ixti(const std::vector<std::string>& sheet_names, std::string_view sheet) {
  for (std::size_t i = 0; i < sheet_names.size(); ++i) {
    if (formulon::sheet_names::equal(sheet_names[i], sheet)) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

/// One rectangle of a precomputed reference result (`PtgMemArea`'s cache).
struct Rect {
  std::uint32_t row_first;
  std::uint32_t row_last;
  std::uint32_t col_first;
  std::uint32_t col_last;
};

/// Fills `out` with the rectangles `node` denotes when it is built from
/// plain same-sheet cell and area references by `,` and single-area ` `
/// alone, so they can be computed without evaluation; an empty
/// intersection leaves no rectangle. False for anything else.
bool StaticRects(const parser::AstNode& node, std::vector<Rect>& out) {
  auto plain = [](const parser::AstNode& n) {
    return n.kind() == parser::NodeKind::Ref && n.as_ref().sheet.empty() && !n.as_ref().is_full_col &&
           !n.as_ref().is_full_row;
  };
  switch (node.kind()) {
    case parser::NodeKind::Ref:
      if (!plain(node)) {
        return false;
      }
      out.push_back(Rect{node.as_ref().row, node.as_ref().row, node.as_ref().col, node.as_ref().col});
      return true;
    case parser::NodeKind::RangeOp: {
      const parser::AstNode& lhs = node.as_range_lhs();
      const parser::AstNode& rhs = node.as_range_rhs();
      if (!plain(lhs) || !plain(rhs)) {
        return false;
      }
      const parser::Reference& a = lhs.as_ref();
      const parser::Reference& b = rhs.as_ref();
      out.push_back(
          Rect{std::min(a.row, b.row), std::max(a.row, b.row), std::min(a.col, b.col), std::max(a.col, b.col)});
      return true;
    }
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (!StaticRects(node.as_union_child(i), out)) {
          return false;
        }
      }
      return true;
    case parser::NodeKind::IntersectOp: {
      std::vector<Rect> lhs;
      std::vector<Rect> rhs;
      if (!StaticRects(node.as_intersect_lhs(), lhs) || !StaticRects(node.as_intersect_rhs(), rhs) ||
          lhs.size() != 1U || rhs.size() != 1U) {
        return false;
      }
      const Rect r{std::max(lhs[0].row_first, rhs[0].row_first), std::min(lhs[0].row_last, rhs[0].row_last),
                   std::max(lhs[0].col_first, rhs[0].col_first), std::min(lhs[0].col_last, rhs[0].col_last)};
      if (r.row_first <= r.row_last && r.col_first <= r.col_last) {
        out.push_back(r);
      }
      return true;
    }
    default:
      return false;
  }
}

/// True when `node` calls one of Excel's volatile functions anywhere.
bool ContainsVolatileCall(const parser::AstNode& node) {
  using parser::NodeKind;
  auto any = [](std::uint32_t n, auto child) {
    for (std::uint32_t i = 0; i < n; ++i) {
      if (ContainsVolatileCall(child(i))) {
        return true;
      }
    }
    return false;
  };
  switch (node.kind()) {
    case NodeKind::Call:
      return parser::is_volatile_function_name(node.as_call_name()) ||
             any(node.as_call_arity(), [&](std::uint32_t i) -> const parser::AstNode& { return node.as_call_arg(i); });
    case NodeKind::UnaryOp:
      return ContainsVolatileCall(node.as_unary_operand());
    case NodeKind::BinaryOp:
      return ContainsVolatileCall(node.as_binary_lhs()) || ContainsVolatileCall(node.as_binary_rhs());
    case NodeKind::RangeOp:
      return ContainsVolatileCall(node.as_range_lhs()) || ContainsVolatileCall(node.as_range_rhs());
    case NodeKind::IntersectOp:
      return ContainsVolatileCall(node.as_intersect_lhs()) || ContainsVolatileCall(node.as_intersect_rhs());
    case NodeKind::UnionOp:
      return any(node.as_union_arity(),
                 [&](std::uint32_t i) -> const parser::AstNode& { return node.as_union_child(i); });
    case NodeKind::ImplicitIntersection:
      return ContainsVolatileCall(node.as_implicit_intersection_operand());
    case NodeKind::SpillRef: {
      const parser::AstNode* anchor = node.as_spill_ref_anchor_expr();
      return anchor != nullptr && ContainsVolatileCall(*anchor);
    }
    case NodeKind::Lambda:
      return ContainsVolatileCall(node.as_lambda_body());
    case NodeKind::LetBinding:
      return ContainsVolatileCall(node.as_let_body()) ||
             any(node.as_let_binding_count(),
                 [&](std::uint32_t i) -> const parser::AstNode& { return node.as_let_binding_expr(i); });
    case NodeKind::LambdaCall:
      return ContainsVolatileCall(node.as_lambda_call_callee()) ||
             any(node.as_lambda_call_arity(),
                 [&](std::uint32_t i) -> const parser::AstNode& { return node.as_lambda_call_arg(i); });
    case NodeKind::Literal:
    case NodeKind::Ref:
    case NodeKind::Ref3D:
    case NodeKind::ExternalRef:
    case NodeKind::StructuredRef:
    case NodeKind::NameRef:
    case NodeKind::ArrayLiteral:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
      return false;
  }
  return false;
}

/// Walker behind `formula_uses_array_evaluation`: true when `node`, at
/// `slot` (`root` for the formula's own node), holds an area or a name a
/// pre-dynamic-array Excel would have intersected: value class as an
/// argument, value or array class as an operator's operand. An argument an
/// array parameter takes whole (`SUMPRODUCT(A1:A2)`) is not one.
bool UsesArrayEvaluation(const parser::AstNode& node, Slot slot, bool root) {
  using parser::NodeKind;
  auto array_class = [&](bool area) {
    const std::uint8_t cls = SlotClass(slot, area);
    return root || cls == kPtgValueClass || (slot.operand && cls == kPtgArrayClass);
  };
  switch (node.kind()) {
    case NodeKind::Ref:
      return (node.as_ref().is_full_col || node.as_ref().is_full_row) && array_class(true);
    case NodeKind::RangeOp:
      return array_class(true);
    case NodeKind::NameRef:
      return array_class(false);
    case NodeKind::ExternalRef:
      return parser::is_self_book_name_ref(node) && array_class(false);
    case NodeKind::IntersectOp:
      // A root intersection of two cells (`A1 B1`) is one cell either way.
      return root &&
             (node.as_intersect_lhs().kind() != NodeKind::Ref || node.as_intersect_rhs().kind() != NodeKind::Ref);
    case NodeKind::UnaryOp:
      return UsesArrayEvaluation(node.as_unary_operand(), Slot{slot.letter, true}, false);
    case NodeKind::BinaryOp:
      return UsesArrayEvaluation(node.as_binary_lhs(), Slot{slot.letter, true}, false) ||
             UsesArrayEvaluation(node.as_binary_rhs(), Slot{slot.letter, true}, false);
    case NodeKind::Call: {
      const std::string_view name = canonical_function_name(node.as_call_name());
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (UsesArrayEvaluation(node.as_call_arg(i), Slot{xlsb_parameter_class(name, i), false}, false)) {
          return true;
        }
      }
      return false;
    }
    case NodeKind::LambdaCall:
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        if (UsesArrayEvaluation(node.as_lambda_call_arg(i), Slot{}, false)) {
          return true;
        }
      }
      return false;
    case NodeKind::LetBinding:
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        if (UsesArrayEvaluation(node.as_let_binding_expr(i), Slot{}, false)) {
          return true;
        }
      }
      return UsesArrayEvaluation(node.as_let_body(), Slot{}, root);
    case NodeKind::Lambda:
      return UsesArrayEvaluation(node.as_lambda_body(), Slot{}, false);
    case NodeKind::Literal:
    case NodeKind::Ref3D:
    case NodeKind::StructuredRef:
    case NodeKind::UnionOp:
    case NodeKind::ImplicitIntersection:
    case NodeKind::ArrayLiteral:
    case NodeKind::SpillRef:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
      return false;
  }
  return false;
}

/// Walker behind `legacy_intersections`: where Excel 365 shows an `@` in a
/// formula saved without the dynamic-array mark (measured cell by cell
/// against its formula2 text). An area, a multi-cell name or an
/// array-returning call takes one wherever an area would be value class.
class LegacyIntersections {
 public:
  LegacyIntersections(const NameIsScalar& name_is_scalar, std::vector<const parser::AstNode*>& out)
      : name_is_scalar_(name_is_scalar), out_(out) {}

  void walk(const parser::AstNode& node, Slot slot, bool root) {
    const bool value = root || SlotClass(slot, /*area=*/true, PtgEvaluation::kLegacy) == kPtgValueClass;
    if (node.kind() == parser::NodeKind::ImplicitIntersection) {
      // A written `@` is listed when it is the one Excel would show; its
      // operand never takes a second one.
      const parser::AstNode& operand = node.as_implicit_intersection_operand();
      if (value && array_valued(operand)) {
        out_.push_back(&node);
      }
      walk_children(operand, slot, root);
      return;
    }
    if (value && array_valued(node)) {
      out_.push_back(&node);
    }
    walk_children(node, slot, root);
  }

 private:
  void walk_children(const parser::AstNode& node, Slot slot, bool root) {
    using parser::NodeKind;
    switch (node.kind()) {
      case NodeKind::UnaryOp:
        walk(node.as_unary_operand(), Slot{slot.letter, true}, false);
        return;
      case NodeKind::BinaryOp:
        walk(node.as_binary_lhs(), Slot{slot.letter, true}, false);
        walk(node.as_binary_rhs(), Slot{slot.letter, true}, false);
        return;
      case NodeKind::RangeOp:
        walk(node.as_range_lhs(), Slot{'R', false}, false);
        walk(node.as_range_rhs(), Slot{'R', false}, false);
        return;
      case NodeKind::IntersectOp:
        walk(node.as_intersect_lhs(), Slot{'R', false}, false);
        walk(node.as_intersect_rhs(), Slot{'R', false}, false);
        return;
      case NodeKind::UnionOp:
        for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
          walk(node.as_union_child(i), Slot{'R', false}, false);
        }
        return;
      case NodeKind::Call: {
        const std::string_view name = canonical_function_name(node.as_call_name());
        const bool builtin = known_builtin(name);
        for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
          walk(node.as_call_arg(i), builtin ? Slot{xlsb_parameter_class(name, i), false} : Slot{}, false);
        }
        return;
      }
      case NodeKind::LambdaCall:
        for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
          walk(node.as_lambda_call_arg(i), Slot{}, false);
        }
        return;
      case NodeKind::LetBinding: {
        const std::size_t scope = let_scope_.size();
        for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
          walk(node.as_let_binding_expr(i), Slot{}, false);
          let_scope_.emplace_back(node.as_let_binding_name(i), array_valued(node.as_let_binding_expr(i)));
        }
        walk(node.as_let_body(), slot, root);
        let_scope_.resize(scope);
        return;
      }
      default:
        // Leaves; a LAMBDA body is evaluated where it is called, not here.
        return;
    }
  }

  bool known_builtin(std::string_view name) const {
    return !let_bound(name) && (lookup_func_by_name(name) != nullptr || UsesHiddenNameRoute(name));
  }

  bool let_bound(std::string_view name) const {
    return std::any_of(let_scope_.begin(), let_scope_.end(),
                       [&](const auto& bound) { return strings::case_insensitive_eq(bound.first, name); });
  }

  bool name_array_valued(const parser::AstNode& node) const {
    if (node.kind() == parser::NodeKind::NameRef && node.as_name_sheet().empty()) {
      for (auto it = let_scope_.rbegin(); it != let_scope_.rend(); ++it) {
        if (strings::case_insensitive_eq(it->first, node.as_name())) {
          return it->second;
        }
      }
    }
    return !name_is_scalar_ || !name_is_scalar_(node);
  }

  bool array_valued(const parser::AstNode& node) const {
    using parser::NodeKind;
    switch (node.kind()) {
      case NodeKind::Ref:
        return node.as_ref().is_full_col || node.as_ref().is_full_row;
      case NodeKind::RangeOp:
        return true;
      case NodeKind::Ref3D:
        return node.as_ref3d_is_range() || node.as_ref3d_sheet_begin() != node.as_ref3d_sheet_end();
      case NodeKind::NameRef:
        return name_array_valued(node);
      case NodeKind::ExternalRef:
        return parser::is_self_book_name_ref(node) ? name_array_valued(node) : node.as_external_ref_is_range();
      case NodeKind::IntersectOp:
        // `A1 B1` is one cell either way.
        return !single_cell(node.as_intersect_lhs()) || !single_cell(node.as_intersect_rhs());
      case NodeKind::StructuredRef:
        return node.as_structured_ref_modifier() != parser::StructuredRefModifier::At;
      case NodeKind::Call:
        return call_returns_array(node);
      case NodeKind::LambdaCall:
      case NodeKind::SpillRef:
        return true;
      default:
        return false;
    }
  }

  static bool single_cell(const parser::AstNode& node) {
    return node.kind() == parser::NodeKind::Ref && !node.as_ref().is_full_col && !node.as_ref().is_full_row;
  }

  /// A number literal equal to `want`, or any nonzero one when `want` is 0.
  static bool number_literal(const parser::AstNode& node, double want) {
    if (node.kind() != parser::NodeKind::Literal || !node.as_literal().is_number()) {
      return false;
    }
    const double n = node.as_literal().as_number();
    return want == 0.0 ? n != 0.0 : n == want;
  }

  bool call_returns_array(const parser::AstNode& node) const {
    const std::string_view name = canonical_function_name(node.as_call_name());
    if (!known_builtin(name)) {
      return true;
    }
    const std::uint32_t arity = node.as_call_arity();
    auto is = [&](std::string_view fn) { return strings::case_insensitive_eq(name, fn); };
    auto any_arg_array_valued = [&](std::uint32_t first) {
      for (std::uint32_t i = first; i < arity; ++i) {
        if (array_valued(node.as_call_arg(i))) {
          return true;
        }
      }
      return false;
    };
    if (is("INDEX")) {
      // A literal nonzero row and column pick one cell; 0, omitted or
      // computed ones may pick a whole row or column.
      return arity < 2 || !number_literal(node.as_call_arg(1), 0.0) ||
             (arity >= 3 && !number_literal(node.as_call_arg(2), 0.0)) ||
             (arity >= 4 && node.as_call_arg(3).kind() != parser::NodeKind::Literal);
    }
    if (is("OFFSET")) {
      // One cell only from a cell base with height and width literal 1 or absent.
      return arity < 3 || !single_cell(node.as_call_arg(0)) ||
             (arity >= 4 && !number_literal(node.as_call_arg(3), 1.0)) ||
             (arity >= 5 && !number_literal(node.as_call_arg(4), 1.0));
    }
    if (is("ROW") || is("COLUMN")) {
      return any_arg_array_valued(0);
    }
    if (is("CHOOSE")) {
      return any_arg_array_valued(1);
    }
    if (is("IF") || is("IFERROR") || is("IFNA") || is("IFS") || is("SWITCH")) {
      return any_arg_array_valued(0);
    }
    if (is("XLOOKUP")) {
      return arity < 3 || !single_line(node.as_call_arg(2));
    }
    static constexpr std::string_view kArrayFunctions[] = {
        "SEQUENCE",   "RANDARRAY",  "FILTER",    "SORT",     "SORTBY",    "UNIQUE",       "TAKE",        "DROP",
        "EXPAND",     "HSTACK",     "VSTACK",    "TOCOL",    "TOROW",     "WRAPCOLS",     "WRAPROWS",    "TRANSPOSE",
        "CHOOSECOLS", "CHOOSEROWS", "MAKEARRAY", "MAP",      "SCAN",      "REDUCE",       "BYROW",       "BYCOL",
        "FREQUENCY",  "MUNIT",      "MMULT",     "MINVERSE", "TEXTSPLIT", "GROUPBY",      "PIVOTBY",     "LINEST",
        "LOGEST",     "TREND",      "GROWTH",    "INDIRECT", "TRIMRANGE", "REGEXEXTRACT", "STOCKHISTORY"};
    return std::any_of(std::begin(kArrayFunctions), std::end(kArrayFunctions), is);
  }

  /// A bounded area one row or one column across.
  static bool single_line(const parser::AstNode& node) {
    if (node.kind() != parser::NodeKind::RangeOp || !single_cell(node.as_range_lhs()) ||
        !single_cell(node.as_range_rhs())) {
      return false;
    }
    const parser::Reference& a = node.as_range_lhs().as_ref();
    const parser::Reference& b = node.as_range_rhs().as_ref();
    return a.row == b.row || a.col == b.col;
  }

  const NameIsScalar& name_is_scalar_;
  std::vector<const parser::AstNode*>& out_;
  std::vector<std::pair<std::string_view, bool>> let_scope_;
};

class Encoder {
 public:
  Encoder(const parser::AstNode& root, const std::vector<std::string>& sheet_names, const SheetRangeTable& sheet_ranges,
          const NameTable& name_table, PtgRootClass root_class, std::optional<PtgBaseCell> base,
          PtgEvaluation evaluation, const NameIsScalar& name_is_scalar)
      : sheet_names_(sheet_names),
        sheet_ranges_(sheet_ranges),
        name_table_(name_table),
        base_(base),
        evaluation_(evaluation),
        promote_root_(root_class == PtgRootClass::kValue) {
    parser::collect_parenthesized_nodes(root, parens_);
    if (evaluation == PtgEvaluation::kLegacy) {
      implied_at_ = legacy_intersections(root, name_is_scalar);
    }
  }

  /// Emits `node`, then `PtgParen` wherever the formula text parenthesises
  /// it: Excel renders a formula from its tokens, and without `PtgParen`
  /// shows `(1+2)*3` as `1+2*3`.
  Expected<void, Error> emit(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit_node(node));
    if (parens_.count(&node) != 0) {
      emit_u8(out_, 0x15);  // PtgParen
    }
    return Expected<void, Error>::Ok();
  }

  EncodedFormula take() { return EncodedFormula{std::move(out_), std::move(extra_)}; }

  void emit_attr_semi() {
    emit_u8(out_, 0x19);  // PtgAttr
    emit_u8(out_, 0x01);  // bitSemi
    emit_u16(out_, 0);
  }

 private:
  Expected<void, Error> emit_node(const parser::AstNode& node) {
    const bool root = std::exchange(is_root_, false);
    // Root promotion applies only to the single outermost node, and only when the caller wants it.
    const bool promote = root && promote_root_;
    const Slot slot = std::exchange(next_slot_, Slot{});
    // A conditional-format operand under a parameter other than a value one
    // is evaluated as an array, and so is everything inside it.
    const bool enters_array = evaluation_ == PtgEvaluation::kConditionalFormat && slot.operand && slot.letter != 0 &&
                              slot.letter != 'V' && !in_array_operand_;
    if (enters_array) {
      in_array_operand_ = true;
    }
    struct Restore {
      bool& flag;
      bool entered;
      ~Restore() {
        if (entered) {
          flag = false;
        }
      }
    } restore{in_array_operand_, enters_array};
    // A reference's class follows its slot in the formulas a value root is
    // measured for; a defined-name body keeps reference class (unmeasured).
    auto cls = [&](bool area) -> std::uint8_t {
      if (promote) {
        return kPtgValueClass;
      }
      const std::uint8_t c = promote_root_ ? SlotClass(slot, area, evaluation_) : kPtgReferenceClass;
      return in_array_operand_ && c == kPtgValueClass ? kPtgArrayClass : c;
    };
    switch (node.kind()) {
      case parser::NodeKind::Literal:
        return emit_literal(node.as_literal());
      case parser::NodeKind::ErrorLiteral:
        emit_u8(out_, 0x1C);  // PtgErr
        emit_u8(out_, error_wire_code(node.as_error_literal()));
        return Expected<void, Error>::Ok();
      case parser::NodeKind::Ref:
        return emit_ref(node.as_ref(), cls(node.as_ref().is_full_col || node.as_ref().is_full_row));
      case parser::NodeKind::Ref3D:
        return emit_ref3d(node, promote);
      case parser::NodeKind::UnaryOp:
        return emit_unary(node, slot);
      case parser::NodeKind::BinaryOp:
        return emit_binary(node, slot);
      case parser::NodeKind::RangeOp:
        return emit_range(node, promote, cls(true));
      case parser::NodeKind::UnionOp:
      case parser::NodeKind::IntersectOp:
        return emit_reference_operation(node, root);
      case parser::NodeKind::Call:
        return emit_call(node);
      case parser::NodeKind::ArrayLiteral:
        return emit_array(node);
      case parser::NodeKind::NameRef:
        if (!node.as_name_sheet().empty()) {
          return emit_sheet_name_ref(node.as_name_sheet(), node.as_name(), cls(false));
        }
        return emit_name_ref(node.as_name(), cls(false));
      case parser::NodeKind::ExternalRef:
        if (parser::is_self_book_name_ref(node)) {
          return emit_self_book_name_ref(node.as_external_ref_name(), cls(false));
        }
        return unsupported_node("ExternalRef");
      case parser::NodeKind::StructuredRef:
        return unsupported_node("StructuredRef");
      case parser::NodeKind::SpillRef:
        return emit_spill_ref(node);
      case parser::NodeKind::ImplicitIntersection:
        if (std::find(implied_at_.begin(), implied_at_.end(), &node) != implied_at_.end()) {
          is_root_ = root;
          next_slot_ = slot;
          return emit(node.as_implicit_intersection_operand());
        }
        return emit_implicit_intersection(node);
      case parser::NodeKind::Lambda:
        return emit_lambda(node);
      case parser::NodeKind::LetBinding:
        return emit_let(node);
      case parser::NodeKind::LambdaCall:
        return emit_lambda_call(node);
      case parser::NodeKind::ErrorPlaceholder:
        return unsupported_node("ErrorPlaceholder");
    }
    return unsupported_node("unknown");
  }

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
  bool in_let_scope(std::string_view name) const {
    return std::any_of(let_scope_.begin(), let_scope_.end(),
                       [name](const auto& binding) { return strings::case_insensitive_eq(binding.first, name); });
  }

  /// `cls`: the class a defined name takes where it sits (see `SlotClass`);
  /// a LET / LAMBDA parameter stays 0x23.
  Expected<void, Error> emit_name_ref(std::string_view name, std::uint8_t cls = kPtgReferenceClass) {
    // Case-insensitive: LET/LAMBDA parameter names resolve case-insensitively
    // (Excel folds ASCII case on name resolution), so a NameRef spelled in a
    // different case than its binding is still that parameter.
    for (auto it = let_scope_.rbegin(); it != let_scope_.rend(); ++it) {
      if (strings::case_insensitive_eq(it->first, name)) {
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
    emit_u8(out_, ClassedPtg(0x23, cls));  // PtgName
    emit_u32(out_, it->second);
    return Expected<void, Error>::Ok();
  }

  /// Emits `PtgNameX` for `sheet!name`, naming the record local to `sheet`
  /// when one exists, else the workbook-scoped one, which is what the
  /// reference resolves to. A name defined in neither names the empty stub
  /// record scoped to `sheet` that Excel 365 saves for it (never another
  /// sheet's definition). Excel writes a sheet-qualified name as `PtgNameX`
  /// through a book-scope `BrtExternSheet` entry rather than as `PtgName`.
  Expected<void, Error> emit_sheet_name_ref(std::string_view sheet, std::string_view name, std::uint8_t cls) {
    const int itab = resolve_ixti(sheet_names_, sheet);
    if (itab < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet-qualified name names no sheet",
                        std::string("context=xlsb_ptg_writer sheet=") + std::string(sheet));
    }
    auto it = name_table_.find(sheet_scoped_name_key(itab, name));
    if (it == name_table_.end()) {
      it = name_table_.find(sheet_scoped_name_key(-1, name));
    }
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: name reference not in name table",
                        std::string("context=xlsb_ptg_writer name=") + std::string(name));
    }
    return emit_name_x(it->second, cls);
  }

  /// Emits `PtgNameX` for the self-book `[0]!name`, naming the record it
  /// resolves to: the workbook-scoped one, else the lowest sheet's local
  /// one, as Excel 365 saves it. An undefined name falls back to its
  /// placeholder record.
  Expected<void, Error> emit_self_book_name_ref(std::string_view name, std::uint8_t cls) {
    auto it = name_table_.find(sheet_scoped_name_key(-1, name));
    for (std::size_t itab = 0; it == name_table_.end() && itab < sheet_names_.size(); ++itab) {
      it = name_table_.find(sheet_scoped_name_key(static_cast<std::int32_t>(itab), name));
    }
    if (it == name_table_.end()) {
      it = name_table_.find(std::string(name));
    }
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: name reference not in name table",
                        std::string("context=xlsb_ptg_writer name=") + std::string(name));
    }
    return emit_name_x(it->second, cls);
  }

  /// Emits `PtgNameX` for record `ilbl` through the sheetless
  /// `BrtExternSheet` entry this workbook's own names resolve through.
  Expected<void, Error> emit_name_x(std::uint32_t ilbl, std::uint8_t cls) {
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_range_ixti(kXtiNoSheet, kXtiNoSheet));
    emit_u8(out_, ClassedPtg(0x39, cls));  // PtgNameX
    emit_u16(out_, ixti);
    emit_u32(out_, ilbl);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a `LambdaCall` the way Excel 365 stores any LAMBDA
  /// invocation: the callee operand (`Sheet1!Fn` / `[0]!Fn` as `PtgNameX`, or an
  /// inline `LAMBDA(...)` / curried call), then the arguments, then
  /// `PtgFuncVar` with the `id == 255` sentinel and `cparams == arity + 1`.
  Expected<void, Error> emit_lambda_call(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit(node.as_lambda_call_callee()));
    const std::uint32_t arity = node.as_lambda_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_lambda_call_arg(i)));
    }
    return emit_hidden_call_tail(arity + 1, "LambdaCall(arity>254)");
  }

  /// Encodes `Fn(args)` whose callee is a defined name or an in-scope LET /
  /// LAMBDA parameter: the callee `PtgName`, the arguments, then
  /// `PtgFuncVar(255)`, as Excel 365 saves a call to a named LAMBDA.
  Expected<void, Error> emit_name_call(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit_name_ref(node.as_call_name()));
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_call_arg(i)));
    }
    return emit_hidden_call_tail(arity + 1, "Call(arity>254, named LAMBDA)");
  }

  /// Encodes a `Lambda` as Excel 365 saves one: `PtgName(_xlfn.LAMBDA)`, a
  /// `PtgName(_xlpm.<param>)` per parameter, the body with the parameters
  /// in scope, then `PtgFuncVar(255)` with `cparams == params + 2`.
  /// Optional `[param]`s have no measured encoding and are refused.
  Expected<void, Error> emit_lambda(const parser::AstNode& node) {
    if (node.as_lambda_optional_count() != 0) {
      return unsupported_node("Lambda(optional parameter)");
    }
    const auto it_lambda = name_table_.find("_xlfn.LAMBDA");
    if (it_lambda == name_table_.end()) {
      return unsupported_node("Lambda(_xlfn.LAMBDA not registered)");
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the LAMBDA name-ref
    emit_u32(out_, it_lambda->second);
    const std::uint32_t n = node.as_lambda_param_count();
    for (std::uint32_t i = 0; i < n; ++i) {
      const std::string_view raw_name = node.as_lambda_param(i);
      const auto it_param = name_table_.find(std::string("_xlpm.") + std::string(raw_name));
      if (it_param == name_table_.end()) {
        return unsupported_node("Lambda(param name not registered)");
      }
      emit_u8(out_, 0x23);  // PtgName (reference-class): the parameter name-ref
      emit_u32(out_, it_param->second);
      let_scope_.emplace_back(raw_name, it_param->second);
    }
    auto status = emit(node.as_lambda_body());
    let_scope_.resize(let_scope_.size() - n);
    RETURN_IF_ERROR(status);
    return emit_hidden_call_tail(n + 2, "Lambda(too many parameters)");
  }

  /// Emits the `PtgFuncVar(255)` that closes a hidden-name-route call whose
  /// `cparams` operands (callee name-ref included) are already on the stack.
  Expected<void, Error> emit_hidden_call_tail(std::uint32_t cparams, const char* too_many) {
    if (cparams > 0xFF) {
      return unsupported_node(too_many);
    }
    emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, static_cast<std::uint8_t>(cparams));
    emit_u16(out_, 255);
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

  /// `cls`: the class bits the reference takes where it sits (see
  /// `SlotClass`); a cell formula's entire body is value class.
  Expected<void, Error> emit_ref(const parser::Reference& ref, std::uint8_t cls) {
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
        emit_u8(out_, ClassedPtg(0x25, cls));  // PtgArea
        emit_area(out_, first, last);
        return Expected<void, Error>::Ok();
      }
      ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_single_sheet_ixti(ref.sheet));
      emit_u8(out_, ClassedPtg(0x3B, cls));  // PtgArea3d
      emit_u16(out_, ixti);
      emit_area(out_, first, last);
      return Expected<void, Error>::Ok();
    }
    if (ref.sheet.empty() && base_ && IsRelative(ref)) {
      emit_u8(out_, ClassedPtg(0x2C, cls));  // PtgRefN
      emit_loc(out_, OffsetFrom(ref, *base_));
      return Expected<void, Error>::Ok();
    }
    if (ref.sheet.empty()) {
      emit_u8(out_, ClassedPtg(0x24, cls));  // PtgRef
      emit_loc(out_, ref);
      return Expected<void, Error>::Ok();
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_single_sheet_ixti(ref.sheet));
    emit_u8(out_, ClassedPtg(0x3A, cls));  // PtgRef3d
    emit_u16(out_, ixti);
    emit_loc(out_, ref);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a `Ref3D` node over a sheet span resolved to `ixti` through
  /// `sheet_ranges_`. A single-cell tail (`Sheet1:Sheet3!A1`) emits
  /// `PtgRef3d(ixti) + RgceLoc`; a range tail (`Sheet1:Sheet3!A1:B2`) emits
  /// `PtgArea3d(ixti) + RgceArea`. `promote`: see `emit_ref`.
  Expected<void, Error> emit_ref3d(const parser::AstNode& node, bool promote) {
    const std::string_view begin = node.as_ref3d_sheet_begin();
    const std::string_view end = node.as_ref3d_sheet_end();
    const int itab_begin = resolve_ixti(sheet_names_, begin);
    const int itab_end = resolve_ixti(sheet_names_, end);
    if (itab_begin < 0 || itab_end < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet not found for 3-D range",
                        std::string("context=xlsb_ptg_writer sheets=") + std::string(begin) + ":" + std::string(end));
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_range_ixti(itab_begin, itab_end));
    // Measured: a value/array-class 3-D reference used as a function
    // argument reads back as #REF! in real Excel -- do not promote here.
    if (node.as_ref3d_is_range()) {
      emit_u8(out_, promote ? ValueClassPtg(0x3B) : 0x3B);  // PtgArea3d
      emit_u16(out_, ixti);
      emit_area(out_, node.as_ref3d_cell(), node.as_ref3d_cell_end());
      return Expected<void, Error>::Ok();
    }
    emit_u8(out_, promote ? ValueClassPtg(0x3A) : 0x3A);  // PtgRef3d
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

  /// Emits an operator's operand; `slot` is where the operator sits.
  Expected<void, Error> emit_operand(const parser::AstNode& node, Slot slot) {
    next_slot_ = Slot{slot.letter, true};
    return emit(node);
  }

  Expected<void, Error> emit_unary(const parser::AstNode& node, Slot slot) {
    RETURN_IF_ERROR(emit_operand(node.as_unary_operand(), slot));
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

  Expected<void, Error> emit_binary(const parser::AstNode& node, Slot slot) {
    RETURN_IF_ERROR(emit_operand(node.as_binary_lhs(), slot));
    RETURN_IF_ERROR(emit_operand(node.as_binary_rhs(), slot));
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

  /// `cls`: the class of the fast-path Area/Area3d collapse (see
  /// `emit_ref`). A root the general form wraps, `operand+operand+PtgRange`,
  /// in a value-class `PtgMemFunc` instead (`promote`; both measured, real
  /// Excel 365).
  Expected<void, Error> emit_range(const parser::AstNode& node, bool promote, std::uint8_t cls) {
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
        if (a.sheet.empty() && base_ && (IsRelative(a) || IsRelative(b))) {
          emit_u8(out_, ClassedPtg(0x2D, cls));  // PtgAreaN
          emit_area(out_, OffsetFrom(a, *base_), OffsetFrom(b, *base_));
          return Expected<void, Error>::Ok();
        }
        if (a.sheet.empty()) {
          emit_u8(out_, ClassedPtg(0x25, cls));  // PtgArea
          emit_area(out_, a, b);
          return Expected<void, Error>::Ok();
        }
        const int itab = resolve_ixti(sheet_names_, a.sheet);
        const int ixti = itab >= 0 ? try_resolve_range_ixti(itab, itab) : -1;
        if (ixti >= 0) {
          emit_u8(out_, ClassedPtg(0x3B, cls));  // PtgArea3d
          emit_u16(out_, static_cast<std::uint16_t>(ixti));
          emit_area(out_, a, b);
          return Expected<void, Error>::Ok();
        }
      }
    }
    // General form: emit both operands then the `:` operator.
    if (!promote) {
      RETURN_IF_ERROR(emit(lhs));
      RETURN_IF_ERROR(emit(rhs));
      emit_u8(out_, 0x11);  // PtgRange
      return Expected<void, Error>::Ok();
    }
    // Wrap the operand+operand+PtgRange run in a value-class PtgMemFunc.
    const std::size_t mark = out_.size();
    RETURN_IF_ERROR(emit(lhs));
    RETURN_IF_ERROR(emit(rhs));
    emit_u8(out_, 0x11);  // PtgRange
    const std::vector<std::uint8_t> wrapped(out_.begin() + static_cast<std::ptrdiff_t>(mark), out_.end());
    out_.resize(mark);
    emit_u8(out_, ValueClassPtg(0x29));  // PtgMemFunc
    emit_u16(out_, static_cast<std::uint16_t>(wrapped.size()));
    out_.insert(out_.end(), wrapped.begin(), wrapped.end());
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_union_or_intersect(const parser::AstNode& node) {
    if (node.kind() == parser::NodeKind::IntersectOp) {
      RETURN_IF_ERROR(emit(node.as_intersect_lhs()));
      RETURN_IF_ERROR(emit(node.as_intersect_rhs()));
      emit_u8(out_, 0x0F);  // PtgIsect
      return Expected<void, Error>::Ok();
    }
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

  /// A union or intersection, inside the memory token Excel 365 puts in
  /// front of it (measured):
  ///   * a cell formula over plain same-sheet references stores the result
  ///     it precomputes: `PtgMemArea` with the rectangles in `rgcb`, or at
  ///     the root, where the value is taken, `PtgMemErr` when that is an
  ///     error (#VALUE! for several areas, #NULL! for an empty intersection);
  ///   * a defined-name body's root takes `PtgMemFunc`.
  /// Other positions and operands stay unwrapped, which Excel also reads.
  Expected<void, Error> emit_reference_operation(const parser::AstNode& node, bool root) {
    if (in_memory_token_) {
      return emit_union_or_intersect(node);  // Only the outermost operation carries one.
    }
    std::vector<Rect> rects;
    // A formula with a base cell (CF / DV) is evaluated per cell, so it has no single result to cache.
    if (promote_root_ && !base_ && StaticRects(node, rects)) {
      if (root && rects.size() != 1U) {
        emit_u8(out_, ValueClassPtg(0x27));  // PtgMemErr
        emit_u8(out_, error_wire_code(rects.empty() ? ErrorCode::Null : ErrorCode::Value));
        emit_u8(out_, 0);
        emit_u16(out_, 0);
        return emit_wrapped(node);
      }
      if (!rects.empty()) {
        emit_u8(out_, root ? ValueClassPtg(0x26) : 0x26);  // PtgMemArea
        emit_u32(out_, 0);                                 // unused
        emit_u32(extra_, static_cast<std::uint32_t>(rects.size()));
        for (const Rect& r : rects) {
          emit_u32(extra_, r.row_first);
          emit_u32(extra_, r.row_last);
          emit_u32(extra_, r.col_first);
          emit_u32(extra_, r.col_last);
        }
        return emit_wrapped(node);
      }
    }
    if (!promote_root_ && root) {
      emit_u8(out_, 0x29);  // PtgMemFunc
      return emit_wrapped(node);
    }
    return emit_union_or_intersect(node);
  }

  /// Emits `node` after the memory token just written, then back-fills the
  /// token's trailing `cce` with the byte length of what it covers.
  Expected<void, Error> emit_wrapped(const parser::AstNode& node) {
    const std::size_t cce_at = out_.size();
    emit_u16(out_, 0);
    in_memory_token_ = true;
    const auto status = emit_union_or_intersect(node);
    in_memory_token_ = false;
    RETURN_IF_ERROR(status);
    const std::size_t cce = out_.size() - cce_at - 2U;
    if (cce > 0xFFFFU) {
      return unsupported_node("memory token (cce>65535)");
    }
    out_[cce_at] = static_cast<std::uint8_t>(cce & 0xFFU);
    out_[cce_at + 1U] = static_cast<std::uint8_t>(cce >> 8);
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_call(const parser::AstNode& node) {
    // A localised formula-bar spelling resolves to the name Excel stores
    // before anything is looked up, so both containers agree on the
    // callee and `func_id_table` needs no alias of its own.
    const std::string_view name = canonical_function_name(node.as_call_name());
    if (in_let_scope(node.as_call_name())) {
      return emit_name_call(node);
    }
    if (UsesHiddenNameRoute(name)) {
      return emit_future_function_call(node, name);
    }
    const XlsbFuncEntry* entry = lookup_func_by_name(name);
    if (entry == nullptr && name_table_.count(std::string(node.as_call_name())) != 0) {
      return emit_name_call(node);
    }
    if (entry == nullptr) {
      // A callee with no id whose name the caller's table lacks. Encoding it
      // through the hidden-name route would make real Excel resolve a
      // hidden `_xlfn.<NAME>` that does not exist (#NAME?), and guessing an
      // id would silently substitute a different function — so the encode
      // fails instead. A callee Excel really has no id for takes that route
      // only by being enumerated in `io::xlsb_uses_hidden_name`.
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: no XLSB function id known for callee",
                        std::string("context=xlsb_ptg_writer fn=") + std::string(name));
    }
    const std::uint32_t arity = node.as_call_arity();
    std::vector<std::size_t> gotos;
    if (strings::case_insensitive_eq(name, "IF") && (arity == 2U || arity == 3U)) {
      RETURN_IF_ERROR(emit_call_arg(node, name, 0U));
      const std::size_t attr_if = emit_attr(0x02);
      RETURN_IF_ERROR(emit_call_arg(node, name, 1U));
      gotos.push_back(emit_attr(0x08));
      patch_u16(attr_if + 2U, out_.size() - (attr_if + 4U));
      if (arity == 3U) {
        RETURN_IF_ERROR(emit_call_arg(node, name, 2U));
        gotos.push_back(emit_attr(0x08));
      }
    } else if (strings::case_insensitive_eq(name, "CHOOSE") && arity >= 2U) {
      RETURN_IF_ERROR(emit_call_arg(node, name, 0U));
      // PtgAttrChoose: the choice count, then one offset per choice and one
      // past the last, each from the start of that offset table.
      const std::size_t table = emit_attr(0x04, static_cast<std::uint16_t>(arity - 1U)) + 4U;
      out_.resize(table + 2U * arity);
      for (std::uint32_t i = 1; i < arity; ++i) {
        patch_u16(table + 2U * (i - 1U), out_.size() - table);
        RETURN_IF_ERROR(emit_call_arg(node, name, i));
        gotos.push_back(emit_attr(0x08));
      }
      patch_u16(table + 2U * (arity - 1U), out_.size() - table);
    } else if (strings::case_insensitive_eq(name, "IFERROR") && arity == 2U) {
      RETURN_IF_ERROR(emit_call_arg(node, name, 0U));
      const std::size_t attr_if_error = emit_attr(0x80);
      RETURN_IF_ERROR(emit_call_arg(node, name, 1U));
      patch_u16(attr_if_error + 2U, out_.size() - (attr_if_error + 4U));
      gotos.push_back(emit_attr(0x08));
    } else {
      for (std::uint32_t i = 0; i < arity; ++i) {
        RETURN_IF_ERROR(emit_call_arg(node, name, i));
      }
    }
    // Each PtgAttrGoto lands on the last byte of the call token that follows.
    const auto patch_gotos = [&]() {
      for (const std::size_t at : gotos) {
        patch_u16(at + 2U, out_.size() - (at + 4U) - 1U);
      }
    };
    // Excel 365 stores a one-argument SUM as `PtgAttrSum` rather than a call.
    if (arity == 1U && strings::case_insensitive_eq(name, "SUM")) {
      emit_u8(out_, 0x19);  // PtgAttr
      emit_u8(out_, 0x10);  // bitSum
      emit_u16(out_, 0);
      return Expected<void, Error>::Ok();
    }
    const bool use_var = entry->variadic || entry->arg_min != entry->arg_max;
    if (use_var) {
      if (arity > 0xFF) {
        return unsupported_node("Call(arity>255)");
      }
      emit_u8(out_, ResultPtg(0x22));  // PtgFuncVar result
      emit_u8(out_, static_cast<std::uint8_t>(arity));
      emit_u16(out_, entry->id);
    } else {
      if (arity != entry->arg_min) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: fixed-arity function arity mismatch",
                          std::string("context=xlsb_ptg_writer fn=") + std::string(name));
      }
      emit_u8(out_, ResultPtg(0x21));  // PtgFunc result
      emit_u16(out_, entry->id);
    }
    patch_gotos();
    return Expected<void, Error>::Ok();
  }

  /// A built-in's result token: value class, array class inside an array
  /// operand of a conditional format (`AND(MONTH(H3)=1)` -> 0x61).
  std::uint8_t ResultPtg(std::uint8_t reference_class_ptg) const {
    return in_array_operand_ ? ClassedPtg(reference_class_ptg, kPtgArrayClass) : ValueClassPtg(reference_class_ptg);
  }

  Expected<void, Error> emit_call_arg(const parser::AstNode& node, std::string_view name, std::uint32_t i) {
    next_slot_ = Slot{xlsb_parameter_class(name, i), false};
    return emit(node.as_call_arg(i));
  }

  /// Emits `PtgAttr` of kind `bits` with `data` and returns its offset.
  std::size_t emit_attr(std::uint8_t bits, std::uint16_t data = 0U) {
    const std::size_t at = out_.size();
    emit_u8(out_, 0x19);  // PtgAttr
    emit_u8(out_, bits);
    emit_u16(out_, data);
    return at;
  }

  void patch_u16(std::size_t at, std::size_t value) {
    out_[at] = static_cast<std::uint8_t>(value & 0xFFU);
    out_[at + 1U] = static_cast<std::uint8_t>((value >> 8) & 0xFFU);
  }

  /// Encodes a call taking the hidden-name route (see
  /// `io/future_functions.h`): `PtgName(ilbl)` naming the callee, then
  /// the real arguments, then `PtgFuncVar` with the `id == 255` sentinel
  /// and `cparams == arity + 1` (the name-ref counts as an operand).
  /// Verified against a real Excel-365-produced
  /// `xl/worksheets/sheetN.bin` for XLOOKUP / TEXTJOIN / CONCAT / IFS /
  /// SEQUENCE and for the bare-in-OOXML `ISO.CEILING` — see
  /// `ptg_reader.cpp`'s `decode_future_function`, the decoder
  /// counterpart. The hidden name is always `_xlfn.`-prefixed even when
  /// OOXML spells the callee bare, so it is spelled by
  /// `xlsb_hidden_function_name` rather than by the OOXML-facing
  /// `storage_function_name`.
  Expected<void, Error> emit_future_function_call(const parser::AstNode& node, std::string_view name) {
    const auto it = name_table_.find(xlsb_hidden_function_name(name));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: hidden-name callee has no BrtName registered",
                        std::string("context=xlsb_ptg_writer fn=") + std::string(name));
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the callee name-ref
    emit_u32(out_, it->second);
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      next_slot_ = Slot{xlsb_parameter_class(name, i), false};
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

  /// Encodes the postfix `#` spill operator as a call to the hidden
  /// `_xlfn.ANCHORARRAY` name -- Excel's own file-format spelling (see
  /// `eval/dynamic_array/anchor.h` and the matching OOXML storage form in
  /// `ast_format.cpp`'s `StorageEmitter`). Same shape as
  /// `emit_future_function_call`, inlined because the anchor operand comes
  /// from `as_spill_ref_anchor_expr`/`as_spill_ref`, not a `Call` node's
  /// argument list.
  Expected<void, Error> emit_spill_ref(const parser::AstNode& node) {
    const auto it = name_table_.find(xlsb_hidden_function_name("ANCHORARRAY"));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: hidden-name callee has no BrtName registered",
                        "context=xlsb_ptg_writer fn=ANCHORARRAY");
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the callee name-ref
    emit_u32(out_, it->second);
    if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
      RETURN_IF_ERROR(emit(*anchor));
    } else {
      // Always a PtgFuncVar argument below, so it stays reference-class.
      RETURN_IF_ERROR(emit_ref(node.as_spill_ref(), kPtgReferenceClass));
    }
    emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, 2);                    // cparams: name-ref + one anchor operand
    emit_u16(out_, 255);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a written `@` as a call to the hidden `_xlfn.SINGLE` name,
  /// as Excel 365 saves `=@A1` (`23 <ilbl> 24 .. 42 02 ff 00`).
  Expected<void, Error> emit_implicit_intersection(const parser::AstNode& node) {
    const auto it = name_table_.find(xlsb_hidden_function_name("SINGLE"));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: hidden-name callee has no BrtName registered",
                        "context=xlsb_ptg_writer fn=SINGLE");
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the callee name-ref
    emit_u32(out_, it->second);
    next_slot_ = Slot{xlsb_parameter_class("SINGLE", 0), false};
    RETURN_IF_ERROR(emit(node.as_implicit_intersection_operand()));
    emit_u8(out_, ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, 2);                    // cparams: name-ref + the operand
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
    // counterpart). Field order (rows before cols) and row-major
    // element order follow [MS-XLSB] 2.5.98.26 -- see the longer note on
    // the decoder side for why the fixture corpus cannot pin them.
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
  /// Nodes the formula text parenthesises (`parser::collect_parenthesized_nodes`).
  std::unordered_set<const parser::AstNode*> parens_;
  /// True while emitting the operation a memory token covers.
  bool in_memory_token_ = false;
  /// Where the next node `emit_node` starts sits; set by its parent.
  Slot next_slot_;
  /// Base cell of `PtgRefN` / `PtgAreaN` offsets, when the formula has one.
  const std::optional<PtgBaseCell> base_;
  /// See `PtgEvaluation`.
  const PtgEvaluation evaluation_;
  /// Inside a conditional format's array operand (see `emit_node`).
  bool in_array_operand_ = false;
  /// Written `@` nodes a legacy formula stores as nothing.
  std::vector<const parser::AstNode*> implied_at_;
  /// See `emit()`'s `promote` local. Cleared on the first call.
  bool is_root_ = true;
  const bool promote_root_;
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

enum class NameCollectMode : std::uint8_t { kPtg, kScopeResolved, kSheetQualified };

/// True when `name` matches an in-scope LET / LAMBDA parameter.
bool InParamScope(const std::vector<std::string_view>& scope, std::string_view name) {
  // Case-insensitive: see `Encoder::emit_name_ref`.
  return std::any_of(scope.begin(), scope.end(),
                     [name](std::string_view param) { return strings::case_insensitive_eq(param, name); });
}

/// Recursive worker for `collect_ptg_names` carrying the LET / LAMBDA
/// parameter names currently in scope (innermost last). A `NameRef`
/// matching an in-scope parameter resolves at encode time to that
/// parameter's hidden `_xlpm.<name>` placeholder (see
/// `Encoder::emit_name_ref`), so it must not be registered as an ordinary
/// workbook defined name. A callee with no function id (`Fn(3)`: a named
/// LAMBDA, or a name no one defined) is a name too. `mode` selects the view:
/// every unqualified name a `PtgName` or self-book `PtgNameX` needs
/// (`kPtg`); only names resolved from the formula's own scope
/// (`kScopeResolved`); or only sheet-qualified names, each added as `sheet`
/// NUL `name` (`kSheetQualified`).
void CollectNamesScoped(const parser::AstNode& node, std::vector<std::string>& names,
                        std::unordered_set<std::string>& seen, std::vector<std::string_view>& scope,
                        NameCollectMode mode) {
  // Hidden `_xlfn.*` / `_xlpm.*` records and unqualified names.
  auto add = [&](std::string_view name) {
    if (mode != NameCollectMode::kSheetQualified) {
      AddName(name, names, seen);
    }
  };
  switch (node.kind()) {
    case parser::NodeKind::NameRef: {
      const std::string_view name = node.as_name();
      if (!node.as_name_sheet().empty()) {
        // Resolves through a scoped key (a definition or its sheet's stub),
        // never a LET / LAMBDA parameter or a workbook placeholder.
        if (mode == NameCollectMode::kSheetQualified) {
          AddName(std::string(node.as_name_sheet()) + '\0' + std::string(name), names, seen);
        }
        return;
      }
      if (InParamScope(scope, name)) {
        return;  // LET / LAMBDA parameter: encoded via its _xlpm. placeholder.
      }
      add(name);
      return;
    }
    case parser::NodeKind::Call: {
      const std::string_view name = canonical_function_name(node.as_call_name());
      if (UsesHiddenNameRoute(name)) {
        add(xlsb_hidden_function_name(name));
      } else if (lookup_func_by_name(name) == nullptr && !InParamScope(scope, node.as_call_name())) {
        add(node.as_call_name());
      }
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_call_arg(i), names, seen, scope, mode);
      }
      return;
    }
    case parser::NodeKind::UnaryOp:
      CollectNamesScoped(node.as_unary_operand(), names, seen, scope, mode);
      return;
    case parser::NodeKind::BinaryOp:
      CollectNamesScoped(node.as_binary_lhs(), names, seen, scope, mode);
      CollectNamesScoped(node.as_binary_rhs(), names, seen, scope, mode);
      return;
    case parser::NodeKind::RangeOp:
      CollectNamesScoped(node.as_range_lhs(), names, seen, scope, mode);
      CollectNamesScoped(node.as_range_rhs(), names, seen, scope, mode);
      return;
    case parser::NodeKind::IntersectOp:
      CollectNamesScoped(node.as_intersect_lhs(), names, seen, scope, mode);
      CollectNamesScoped(node.as_intersect_rhs(), names, seen, scope, mode);
      return;
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_union_child(i), names, seen, scope, mode);
      }
      return;
    }
    case parser::NodeKind::ImplicitIntersection:
      add(xlsb_hidden_function_name("SINGLE"));
      CollectNamesScoped(node.as_implicit_intersection_operand(), names, seen, scope, mode);
      return;
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          CollectNamesScoped(node.as_array_element(r, c), names, seen, scope, mode);
        }
      }
      return;
    }
    case parser::NodeKind::SpillRef: {
      // Stored as a call to the hidden `_xlfn.ANCHORARRAY` name (see
      // `Encoder::emit_spill_ref`), so it needs the same BrtName
      // registration any other future-function callee gets.
      add(xlsb_hidden_function_name("ANCHORARRAY"));
      if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        CollectNamesScoped(*anchor, names, seen, scope, mode);
      }
      return;
    }
    case parser::NodeKind::ExternalRef:
      // `[0]!Rate` is never a LET / LAMBDA parameter and, like `Sheet2!Rate`,
      // not resolved from the formula's own scope.
      if (parser::is_self_book_name_ref(node) && mode == NameCollectMode::kPtg) {
        AddName(node.as_external_ref_name(), names, seen);
      }
      return;
    case parser::NodeKind::LambdaCall: {
      CollectNamesScoped(node.as_lambda_call_callee(), names, seen, scope, mode);
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_lambda_call_arg(i), names, seen, scope, mode);
      }
      return;
    }
    case parser::NodeKind::LetBinding: {
      add("_xlfn.LET");
      const std::uint32_t n = node.as_let_binding_count();
      const std::size_t scope_base = scope.size();
      for (std::uint32_t i = 0; i < n; ++i) {
        add(std::string("_xlpm.") + std::string(node.as_let_binding_name(i)));
        // Excel LET binds sequentially: a value expression sees only the
        // earlier bindings, so collect it before pushing this parameter.
        CollectNamesScoped(node.as_let_binding_expr(i), names, seen, scope, mode);
        scope.push_back(node.as_let_binding_name(i));
      }
      CollectNamesScoped(node.as_let_body(), names, seen, scope, mode);
      scope.resize(scope_base);
      return;
    }
    case parser::NodeKind::Lambda: {
      add("_xlfn.LAMBDA");
      const std::uint32_t n = node.as_lambda_param_count();
      const std::size_t scope_base = scope.size();
      for (std::uint32_t i = 0; i < n; ++i) {
        add(std::string("_xlpm.") + std::string(node.as_lambda_param(i)));
        scope.push_back(node.as_lambda_param(i));
      }
      CollectNamesScoped(node.as_lambda_body(), names, seen, scope, mode);
      scope.resize(scope_base);
      return;
    }
    // Leaves, and forms the encoder does not lower (StructuredRef): nothing
    // to collect. A future writer bundle that lowers these would extend
    // this switch alongside the corresponding `emit_*` case.
    default:
      return;
  }
}

}  // namespace

std::string sheet_scoped_name_key(std::int32_t itab, std::string_view name) {
  std::string key = std::to_string(itab);
  key.push_back('!');
  key.append(strings::to_ascii_lower(name));
  return key;
}

void collect_ptg_names(const parser::AstNode& node, std::vector<std::string>& names,
                       std::unordered_set<std::string>& seen) {
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, names, seen, scope, NameCollectMode::kPtg);
}

void collect_scope_resolved_names(const parser::AstNode& node, std::vector<std::string>& names,
                                  std::unordered_set<std::string>& seen) {
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, names, seen, scope, NameCollectMode::kScopeResolved);
}

void collect_sheet_qualified_names(const parser::AstNode& node,
                                   std::vector<std::pair<std::string, std::string>>& qualified) {
  std::vector<std::string> keys;
  std::unordered_set<std::string> seen;
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, keys, seen, scope, NameCollectMode::kSheetQualified);
  for (const std::string& key : keys) {
    const std::size_t split = key.find('\0');
    qualified.emplace_back(key.substr(0, split), key.substr(split + 1));
  }
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
    case parser::NodeKind::NameRef:
    case parser::NodeKind::ExternalRef: {
      // `Sheet1!Rate` and `[0]!Rate` encode as `PtgNameX` through the
      // book-scope entry.
      const bool name_x = node.kind() == parser::NodeKind::NameRef ? !node.as_name_sheet().empty()
                                                                   : parser::is_self_book_name_ref(node);
      if (name_x && seen.insert(PackRangeKey(kXtiNoSheet, kXtiNoSheet)).second) {
        ranges.emplace_back(kXtiNoSheet, kXtiNoSheet);
      }
      return;
    }
    case parser::NodeKind::LambdaCall: {
      collect_ptg_sheet_ranges(node.as_lambda_call_callee(), sheet_names, ranges, seen);
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_lambda_call_arg(i), sheet_names, ranges, seen);
      }
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
    case parser::NodeKind::Lambda:
      collect_ptg_sheet_ranges(node.as_lambda_body(), sheet_names, ranges, seen);
      return;
    default:
      return;
  }
}

bool formula_uses_array_evaluation(const parser::AstNode& root) {
  return UsesArrayEvaluation(root, Slot{}, /*root=*/true);
}

std::vector<const parser::AstNode*> legacy_intersections(const parser::AstNode& root,
                                                         const NameIsScalar& name_is_scalar) {
  std::vector<const parser::AstNode*> out;
  LegacyIntersections(name_is_scalar, out).walk(root, Slot{}, /*root=*/true);
  return out;
}

Expected<EncodedFormula, Error> encode_ptgs(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                                            const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                            PtgRootClass root_class, std::optional<PtgBaseCell> base,
                                            PtgEvaluation evaluation, const NameIsScalar& name_is_scalar) {
  Encoder enc(node, sheet_names, sheet_ranges, name_table, root_class, base, evaluation, name_is_scalar);
  // A formula calling a volatile function itself (not through a name) opens
  // with `PtgAttrSemi`, without which Excel does not recalculate it
  // (measured for cells and name bodies; the u16 is unused).
  if (ContainsVolatileCall(node)) {
    enc.emit_attr_semi();
  }
  auto status = enc.emit(node);
  if (!status) {
    return status.error();
  }
  return enc.take();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
