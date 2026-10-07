//
// Implementation of the AST -> Ptg encoder. See `io/xlsb/ptg_writer.h`.
//
// The encoder recurses post-order. Helper `emit_*` functions append the
// little-endian wire form for each token; the shapes are byte-matched to
// the decoder in `ptg_reader.cpp`.

#include "io/xlsb/ptg_writer.h"

#include <algorithm>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/record_writer.h"
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

/// Built-ins whose result can be a reference, as measured for the class of
/// their call token.
bool ReturnsReference(std::string_view name) {
  for (const std::string_view fn : {"INDEX", "INDIRECT", "CHOOSE", "IF", "OFFSET", "IFS", "SWITCH", "XLOOKUP"}) {
    if (strings::case_insensitive_eq(name, fn)) {
      return true;
    }
  }
  return false;
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

/// Packs a reference's column field (14-bit column plus the two
/// relative-flag bits), shared by `RgceLoc` and each `RgceArea` corner.
/// The relative bit is *set* when the coordinate is relative (i.e. not
/// `$`-anchored), matching the decoder.
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

/// Emits the RgceLoc (u32 row + u16 col with relative-flag bits) for a
/// single-cell reference.
void emit_loc(std::vector<std::uint8_t>& dst, const parser::Reference& ref) {
  emit_u32(dst, ref.row);
  emit_u16(dst, pack_area_col(ref));
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

/// Turns `first`:`last` over whole columns or rows (per `first`'s flags)
/// into the grid-spanning area XLSB stores, its spanned axis absolute
/// (measured: `A:A` -> 0x4000, `1:1` -> 0x8000..0xBFFF).
void SpanGrid(parser::Reference& first, parser::Reference& last) {
  if (first.is_full_col) {
    first.row = 0;
    last.row = 1048575U;
    first.row_abs = last.row_abs = true;
  } else if (first.is_full_row) {
    first.col = 0;
    last.col = 16383U;
    first.col_abs = last.col_abs = true;
  }
  first.is_full_col = first.is_full_row = last.is_full_col = last.is_full_row = false;
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

Error unsupported_node(const char* kind) {
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                    std::string("xlsb encoder cannot lower AST node kind: ") + kind, "context=xlsb_ptg_writer");
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
  if (node.kind() == parser::NodeKind::Call && parser::is_volatile_function_name(node.as_call_name())) {
    return true;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (ContainsVolatileCall(*child)) {
      return true;
    }
  }
  return false;
}

/// A node's result as Excel 365 judges it statically (measured against the
/// dynamic-array marks it gives formulas and the `@`s it shows in legacy
/// ones): one value, a reference that may cover several cells, or an array.
enum class Shape : std::uint8_t { kScalar, kReference, kArray };

/// Built-ins that return one of their `S` arguments, or an error in its place.
bool PassesArgumentThrough(std::string_view name) {
  for (const std::string_view fn : {"IF", "CHOOSE", "IFERROR", "IFNA"}) {
    if (strings::case_insensitive_eq(name, fn)) {
      return true;
    }
  }
  return false;
}

/// True when a call to `name` is a built-in Excel knows, by id or by hidden name.
bool IsBuiltin(std::string_view name) {
  return lookup_func_by_name(name) != nullptr || UsesHiddenNameRoute(name);
}

/// `Shape` of formula nodes, tracking the LET / LAMBDA bindings in scope
/// while a walker descends. With `lifting` off, a node is multi-valued only
/// by what it is or returns, not through operators or element-wise
/// evaluation of its arguments: what a legacy formula intersects.
class Shapes {
 public:
  explicit Shapes(const NameShapes& names, bool lifting = true) : names_(names), lifting_(lifting) {}

  /// A multi-valued argument at a direct `letter` parameter, which the
  /// built-in then evaluates element by element: any at a value (`V`) or
  /// `I` parameter, and an array at a reference (`R`) one.
  static bool lifts(char letter, Shape shape) {
    return shape != Shape::kScalar && (letter == 'V' || letter == 'I' || (letter == 'R' && shape == Shape::kArray));
  }

  /// `lifts` for argument `index` of the built-in `name`: VLOOKUP's and
  /// HLOOKUP's index, and FORMULATEXT's and ISFORMULA's reference, also take
  /// a multi-valued argument element by element.
  static bool lifts_argument(std::string_view name, std::uint32_t index, char letter, Shape shape) {
    auto is = [&](std::string_view fn) { return strings::case_insensitive_eq(name, fn); };
    return lifts(letter, shape) ||
           (shape != Shape::kScalar && (((is("VLOOKUP") || is("HLOOKUP")) && index == 2U) ||
                                        ((is("FORMULATEXT") || is("ISFORMULA")) && letter == 'R')));
  }

  void bind(std::string_view name, Shape shape) { scope_.emplace_back(name, shape); }
  void unbind(std::size_t count) { scope_.resize(scope_.size() - count); }

  bool bound(std::string_view name) const {
    return std::any_of(scope_.begin(), scope_.end(),
                       [&](const auto& b) { return strings::case_insensitive_eq(b.first, name); });
  }

  Shape of(const parser::AstNode& node) {
    using parser::NodeKind;
    switch (node.kind()) {
      case NodeKind::Ref:
        return node.as_ref().is_full_col || node.as_ref().is_full_row ? Shape::kReference : Shape::kScalar;
      case NodeKind::RangeOp:
      case NodeKind::SpillRef:
        return Shape::kReference;
      case NodeKind::Ref3D:
        return node.as_ref3d_is_range() || node.as_ref3d_sheet_begin() != node.as_ref3d_sheet_end() ? Shape::kReference
                                                                                                    : Shape::kScalar;
      case NodeKind::IntersectOp:
        // One cell on either side bounds the intersection to it.
        return of(node.as_intersect_lhs()) != Shape::kScalar && of(node.as_intersect_rhs()) != Shape::kScalar
                   ? Shape::kReference
                   : Shape::kScalar;
      case NodeKind::NameRef:
        return name_shape(node);
      case NodeKind::ExternalRef:
        if (parser::is_self_book_name_ref(node)) {
          return name_shape(node);
        }
        return node.as_external_ref_is_range() || !node.as_external_ref_name().empty() ||
                       !node.as_external_ref_sheet_end().empty() || node.as_external_ref_cell().is_full_col ||
                       node.as_external_ref_cell().is_full_row
                   ? Shape::kReference
                   : Shape::kScalar;
      case NodeKind::StructuredRef:
        return node.as_structured_ref_modifier() == parser::StructuredRefModifier::At ? Shape::kScalar
                                                                                      : Shape::kReference;
      case NodeKind::UnaryOp:
        return !lifting_ || of(node.as_unary_operand()) == Shape::kScalar ? Shape::kScalar : Shape::kArray;
      case NodeKind::BinaryOp:
        return !lifting_ || (of(node.as_binary_lhs()) == Shape::kScalar && of(node.as_binary_rhs()) == Shape::kScalar)
                   ? Shape::kScalar
                   : Shape::kArray;
      case NodeKind::ArrayLiteral:
        return lifting_ ? Shape::kArray : Shape::kScalar;
      case NodeKind::LambdaCall:
        return Shape::kArray;
      case NodeKind::LetBinding: {
        const std::uint32_t n = node.as_let_binding_count();
        for (std::uint32_t i = 0; i < n; ++i) {
          bind(node.as_let_binding_name(i), of(node.as_let_binding_expr(i)));
        }
        const Shape body = of(node.as_let_body());
        unbind(n);
        return body;
      }
      case NodeKind::Call:
        return call_shape(node);
      case NodeKind::Literal:
      case NodeKind::ErrorLiteral:
      case NodeKind::ErrorPlaceholder:
      case NodeKind::UnionOp:
      case NodeKind::ImplicitIntersection:
      case NodeKind::Lambda:
        return Shape::kScalar;
    }
    return Shape::kReference;
  }

 private:
  const std::pair<std::string_view, Shape>* binding(const parser::AstNode& node) const {
    if (node.kind() != parser::NodeKind::NameRef || !node.as_name_sheet().empty()) {
      return nullptr;
    }
    for (auto it = scope_.rbegin(); it != scope_.rend(); ++it) {
      if (strings::case_insensitive_eq(it->first, node.as_name())) {
        return &*it;
      }
    }
    return nullptr;
  }

  NameShape defined(const parser::AstNode& node) const { return names_ ? names_(node) : NameShape{}; }

  Shape name_shape(const parser::AstNode& node) const {
    if (const auto* b = binding(node)) {
      return b->second;
    }
    return defined(node).scalar ? Shape::kScalar : Shape::kReference;
  }

  /// `node` is a positive integer constant Excel stores as `PtgInt` (equal to
  /// `want` unless that is 0), literally or through a name.
  bool positive_int(const parser::AstNode& node, std::uint32_t want = 0U) const {
    std::uint32_t n = 0;
    if (node.kind() == parser::NodeKind::Literal && node.as_literal().is_number()) {
      const double d = node.as_literal().as_number();
      n = d >= 1.0 && d <= 65535.0 && d == static_cast<double>(static_cast<std::uint32_t>(d))
              ? static_cast<std::uint32_t>(d)
              : 0U;
    } else if (node.kind() == parser::NodeKind::NameRef && binding(node) == nullptr) {
      n = defined(node).positive_int;
    }
    return n != 0U && (want == 0U || n == want);
  }

  /// Rows and columns of the one plain rectangle `node` denotes, if it does.
  std::optional<std::pair<std::uint32_t, std::uint32_t>> extent(const parser::AstNode& node) const {
    constexpr std::uint32_t kRows = 1048576U;
    constexpr std::uint32_t kCols = 16384U;
    auto one = [&](const parser::Reference& r) {
      return std::make_pair(r.is_full_col ? kRows : 1U, r.is_full_row ? kCols : 1U);
    };
    switch (node.kind()) {
      case parser::NodeKind::Ref:
        return one(node.as_ref());
      case parser::NodeKind::RangeOp: {
        const parser::AstNode& lhs = node.as_range_lhs();
        const parser::AstNode& rhs = node.as_range_rhs();
        if (lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
          return std::nullopt;
        }
        const parser::Reference& a = lhs.as_ref();
        const parser::Reference& b = rhs.as_ref();
        const std::uint32_t rows = a.is_full_col ? kRows : (a.row > b.row ? a.row - b.row : b.row - a.row) + 1U;
        const std::uint32_t cols = a.is_full_row ? kCols : (a.col > b.col ? a.col - b.col : b.col - a.col) + 1U;
        return std::make_pair(rows, cols);
      }
      case parser::NodeKind::NameRef: {
        if (binding(node) != nullptr) {
          return std::nullopt;
        }
        const NameShape s = defined(node);
        if (s.rows == 0U) {
          return std::nullopt;
        }
        return std::make_pair(s.rows, s.cols);
      }
      default:
        return std::nullopt;
    }
  }

  Shape call_shape(const parser::AstNode& node) {
    const std::string_view name = canonical_function_name(node.as_call_name());
    if (bound(node.as_call_name()) || !IsBuiltin(name)) {
      return Shape::kArray;  // a LAMBDA through a name, or a function Excel does not know
    }
    const std::uint32_t arity = node.as_call_arity();
    auto is = [&](std::string_view fn) { return strings::case_insensitive_eq(name, fn); };
    Shape passed = Shape::kScalar;
    for (std::uint32_t i = 0; i < arity; ++i) {
      const char letter = xlsb_parameter_class(name, i);
      const Shape arg = of(node.as_call_arg(i));
      if (lifting_ && lifts_argument(name, i, letter, arg)) {
        return Shape::kArray;
      }
      // Without lifting, a multi-valued `I` argument still reaches the result
      // of a call that passes its arguments through (`IF(A1:A2,1,0)`).
      if ((letter == 'S' || (!lifting_ && letter == 'I')) && arg != Shape::kScalar && passed != Shape::kArray) {
        passed = arg;
      }
    }
    if (PassesArgumentThrough(name)) {
      return passed;
    }
    if (is("INDEX")) {
      // A reference indexed by positive integer constants is one cell.
      const Shape source = arity == 0U ? Shape::kReference : of(node.as_call_arg(0));
      if (source != Shape::kReference) {
        return source;
      }
      for (std::uint32_t i = 1; i < arity; ++i) {
        if (!positive_int(node.as_call_arg(i))) {
          return Shape::kReference;
        }
      }
      return arity < 2U ? Shape::kReference : Shape::kScalar;
    }
    if (is("OFFSET")) {
      // One cell from a one-cell base whose height and width are 1 or absent.
      if (arity < 3U || of(node.as_call_arg(0)) != Shape::kScalar ||
          (arity >= 4U && !positive_int(node.as_call_arg(3), 1U)) ||
          (arity >= 5U && !positive_int(node.as_call_arg(4), 1U))) {
        return Shape::kReference;
      }
      return Shape::kScalar;
    }
    if (is("INDIRECT")) {
      return Shape::kReference;
    }
    if (is("ROW") || is("COLUMN")) {
      return arity != 0U && of(node.as_call_arg(0)) != Shape::kScalar ? Shape::kArray : Shape::kScalar;
    }
    if (is("XLOOKUP")) {
      // One cell when the return array runs along the lookup array: one
      // column for a column (or a single cell), one row for a row.
      if (arity < 3U || passed != Shape::kScalar) {
        return Shape::kReference;
      }
      const auto lookup = extent(node.as_call_arg(1));
      const auto result = extent(node.as_call_arg(2));
      if (!lookup || !result) {
        return Shape::kReference;
      }
      if (lookup->second == 1U) {
        return result->second == 1U ? Shape::kScalar : Shape::kReference;
      }
      if (lookup->first == 1U) {
        return result->first == 1U ? Shape::kScalar : Shape::kReference;
      }
      return Shape::kReference;
    }
    return xlsb_returns_array(name) ? Shape::kArray : Shape::kScalar;
  }

  const NameShapes& names_;
  const bool lifting_;
  std::vector<std::pair<std::string_view, Shape>> scope_;
};

/// Walker behind `formula_is_dynamic_array`: true when `node`, at `slot`
/// (`root` for the formula's own node), is multi-valued where Excel would
/// otherwise take one value: the root, an operator's operand, or an argument
/// its parameter evaluates element by element (`Shapes::lifts`).
bool MarksDynamic(Shapes& shapes, const parser::AstNode& node, Slot slot, bool root) {
  using parser::NodeKind;
  const Shape shape = shapes.of(node);
  if (shape != Shape::kScalar && (root || slot.operand || Shapes::lifts(slot.letter, shape))) {
    return true;
  }
  switch (node.kind()) {
    case NodeKind::UnaryOp:
      return MarksDynamic(shapes, node.as_unary_operand(), Slot{slot.letter, true}, false);
    case NodeKind::BinaryOp:
      return MarksDynamic(shapes, node.as_binary_lhs(), Slot{slot.letter, true}, false) ||
             MarksDynamic(shapes, node.as_binary_rhs(), Slot{slot.letter, true}, false);
    case NodeKind::RangeOp:
      return MarksDynamic(shapes, node.as_range_lhs(), Slot{'R', false}, false) ||
             MarksDynamic(shapes, node.as_range_rhs(), Slot{'R', false}, false);
    case NodeKind::IntersectOp:
      return MarksDynamic(shapes, node.as_intersect_lhs(), Slot{'R', false}, false) ||
             MarksDynamic(shapes, node.as_intersect_rhs(), Slot{'R', false}, false);
    case NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (MarksDynamic(shapes, node.as_union_child(i), Slot{'R', false}, false)) {
          return true;
        }
      }
      return false;
    case NodeKind::Call: {
      const std::string_view name = canonical_function_name(node.as_call_name());
      const bool builtin = !shapes.bound(node.as_call_name()) && IsBuiltin(name);
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        // A LAMBDA's arguments are bound whole.
        const Slot arg = builtin ? Slot{xlsb_parameter_class(name, i), false} : Slot{};
        if ((builtin && Shapes::lifts_argument(name, i, arg.letter, shapes.of(node.as_call_arg(i)))) ||
            MarksDynamic(shapes, node.as_call_arg(i), arg, false)) {
          return true;
        }
      }
      return false;
    }
    case NodeKind::LambdaCall:
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        if (MarksDynamic(shapes, node.as_lambda_call_arg(i), Slot{}, false)) {
          return true;
        }
      }
      return false;
    case NodeKind::LetBinding: {
      const std::uint32_t n = node.as_let_binding_count();
      bool marks = false;
      for (std::uint32_t i = 0; i < n; ++i) {
        marks = marks || MarksDynamic(shapes, node.as_let_binding_expr(i), Slot{'S', false}, false);
        shapes.bind(node.as_let_binding_name(i), shapes.of(node.as_let_binding_expr(i)));
      }
      marks = marks || MarksDynamic(shapes, node.as_let_body(), Slot{'S', false}, root);
      shapes.unbind(n);
      return marks;
    }
    case NodeKind::ImplicitIntersection:
      // The operand is intersected; what it is computed from is not.
      return MarksDynamic(shapes, node.as_implicit_intersection_operand(), Slot{'Y', false}, false);
    default:
      // Leaves, `#`, and a LAMBDA, whose body is evaluated where it is called.
      return false;
  }
}

/// Sets `flag` while it lives, when `enter` and the flag is not already set.
class ArrayScope {
 public:
  ArrayScope(bool& flag, bool enter) : flag_(flag), entered_(enter && !flag) {
    if (entered_) {
      flag_ = true;
    }
  }
  ~ArrayScope() {
    if (entered_) {
      flag_ = false;
    }
  }
  ArrayScope(const ArrayScope&) = delete;
  ArrayScope& operator=(const ArrayScope&) = delete;

 private:
  bool& flag_;
  const bool entered_;
};

/// Walker behind `legacy_intersections`: where Excel 365 shows an `@` in a
/// formula saved without the dynamic-array mark (measured cell by cell
/// against its formula2 text). A multi-valued node takes one wherever an
/// area would be value class.
class LegacyIntersections {
 public:
  LegacyIntersections(const NameShapes& names, std::vector<const parser::AstNode*>& out)
      : shapes_(names, /*lifting=*/false), out_(out) {}

  void walk(const parser::AstNode& node, Slot slot, bool root) {
    // Everything under a forced-array parameter is evaluated as an array,
    // nested calls included (`SUMPRODUCT(ABS(A1:A2))` shows no `@`).
    const bool value =
        !in_forced_array_ && (root || SlotClass(slot, /*area=*/true, PtgEvaluation::kLegacy) == kPtgValueClass);
    if (node.kind() == parser::NodeKind::ImplicitIntersection) {
      // A written `@` is listed when it is the one Excel would show; its
      // operand never takes a second one.
      const parser::AstNode& operand = node.as_implicit_intersection_operand();
      if (value && shapes_.of(operand) != Shape::kScalar) {
        out_.push_back(&node);
      }
      walk_children(operand, slot, root);
      return;
    }
    if (value && shapes_.of(node) != Shape::kScalar) {
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
        const bool builtin = !shapes_.bound(node.as_call_name()) && IsBuiltin(name);
        for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
          const Slot arg = builtin ? Slot{xlsb_parameter_class(name, i), false} : Slot{};
          const ArrayScope forced(in_forced_array_, arg.letter == 'F');
          walk(node.as_call_arg(i), arg, false);
        }
        return;
      }
      case NodeKind::LambdaCall:
        for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
          walk(node.as_lambda_call_arg(i), Slot{}, false);
        }
        return;
      case NodeKind::LetBinding: {
        const std::uint32_t n = node.as_let_binding_count();
        for (std::uint32_t i = 0; i < n; ++i) {
          walk(node.as_let_binding_expr(i), Slot{}, false);
          shapes_.bind(node.as_let_binding_name(i), shapes_.of(node.as_let_binding_expr(i)));
        }
        walk(node.as_let_body(), slot, root);
        shapes_.unbind(n);
        return;
      }
      default:
        // Leaves; a LAMBDA body is evaluated where it is called, not here.
        return;
    }
  }

  Shapes shapes_;
  std::vector<const parser::AstNode*>& out_;
  bool in_forced_array_ = false;
};

class Encoder {
 public:
  Encoder(const parser::AstNode& root, const std::vector<std::string>& sheet_names, const SheetRangeTable& sheet_ranges,
          const NameTable& name_table, PtgRootClass root_class, std::optional<PtgBaseCell> base,
          PtgEvaluation evaluation, const NameShapes& names)
      : sheet_names_(sheet_names),
        sheet_ranges_(sheet_ranges),
        name_table_(name_table),
        base_(base),
        evaluation_(evaluation),
        shapes_(names),
        promote_root_(root_class == PtgRootClass::kValue),
        name_body_(root_class == PtgRootClass::kReference && !base) {
    parser::collect_parenthesized_nodes(root, parens_);
    if (evaluation == PtgEvaluation::kLegacy) {
      implied_at_ = legacy_intersections(root, names);
      // With nothing to intersect, a legacy formula means what it means typed
      // into Excel 365, and is written the same way.
      if (implied_at_.empty()) {
        evaluation_ = PtgEvaluation::kDynamicArray;
      }
    }
    measured_calls_ = !base && evaluation_ == PtgEvaluation::kDynamicArray;
    marked_ = promote_root_ && !base && evaluation == PtgEvaluation::kDynamicArray;
    // A defined name's body is evaluated as an array formula throughout.
    in_array_operand_ = name_body_;
  }

  /// Emits `node`, then `PtgParen` wherever the formula text parenthesises
  /// it: Excel renders a formula from its tokens, and without `PtgParen`
  /// shows `(1+2)*3` as `1+2*3`.
  Expected<void, Error> emit(const parser::AstNode& node) {
    RETURN_IF_ERROR(emit_node(node));
    if (const auto it = parens_.find(&node); it != parens_.end()) {
      out_.insert(out_.end(), it->second, 0x15);  // PtgParen, one per pair
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
    const ArrayScope scope(in_array_operand_, evaluation_ == PtgEvaluation::kConditionalFormat && slot.operand &&
                                                  slot.letter != 0 && slot.letter != 'V');
    // A reference's class follows its slot in a cell formula and a defined
    // name's body; other formulas without a value root keep reference class.
    auto cls = [&](bool area) -> std::uint8_t {
      if (promote) {
        return kPtgValueClass;
      }
      const std::uint8_t c =
          promote_root_ || name_body_ ? SlotClass(slot, area || MarkedAsArea(slot), evaluation_) : kPtgReferenceClass;
      return in_array_operand_ && c == kPtgValueClass ? kPtgArrayClass : c;
    };
    // A call's token: reference class where its slot takes a reference, else
    // the class its result takes there (`CallClass`).
    const bool reference_slot =
        !promote &&
        (root ? name_body_ : (promote_root_ || name_body_) && SlotClass(slot, true, evaluation_) == kPtgReferenceClass);
    switch (node.kind()) {
      case parser::NodeKind::Literal:
        return emit_literal(node.as_literal());
      case parser::NodeKind::ErrorLiteral:
        if (measured_calls_ && node.as_error_literal() == ErrorCode::Ref) {
          // Excel stores a written #REF! as a reference error with its slot's class.
          emit_u8(out_, ClassedPtg(0x2A, cls(false)));  // PtgRefErr
          emit_u32(out_, 0);
          emit_u16(out_, 0);
          return Expected<void, Error>::Ok();
        }
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
        call_in_reference_slot_ = reference_slot;
        call_class_ = CallClass(node, slot, promote);
        return emit_call(node);
      case parser::NodeKind::ArrayLiteral: {
        // An array constant takes its slot's class, array where that would be reference.
        const std::uint8_t c = measured_calls_ ? cls(true) : kPtgArrayClass;
        return emit_array(node, c == kPtgReferenceClass ? kPtgArrayClass : c);
      }
      case parser::NodeKind::NameRef: {
        // A name's class follows what it stands for, as an area's or a cell's.
        const std::uint8_t name_cls = cls(measured_calls_ && shapes_.of(node) != Shape::kScalar);
        if (!node.as_name_sheet().empty()) {
          if (external_link_position(sheet_ranges_, node) != 0U) {
            return emit_external_ref(node, cls(measured_calls_), promote);
          }
          return emit_sheet_name_ref(node.as_name_sheet(), node.as_name(), name_cls);
        }
        return emit_name_ref(node.as_name(), name_cls);
      }
      case parser::NodeKind::ExternalRef:
        if (parser::is_self_book_name_ref(node)) {
          return emit_self_book_name_ref(node.as_external_ref_name(),
                                         cls(measured_calls_ && shapes_.of(node) != Shape::kScalar));
        }
        // A name's class follows a local name's; a cell's or an area's a local reference's.
        return emit_external_ref(
            node, cls(node.as_external_ref_name().empty() ? shapes_.of(node) != Shape::kScalar : measured_calls_),
            promote);
      case parser::NodeKind::StructuredRef:
        return unsupported_node("StructuredRef");
      case parser::NodeKind::SpillRef:
        return emit_spill_ref(node, HiddenCallClass(node, slot, promote, reference_slot));
      case parser::NodeKind::ImplicitIntersection:
        if (std::find(implied_at_.begin(), implied_at_.end(), &node) != implied_at_.end()) {
          is_root_ = root;
          next_slot_ = slot;
          return emit(node.as_implicit_intersection_operand());
        }
        return emit_implicit_intersection(node, name_body_ && reference_slot);
      case parser::NodeKind::Lambda:
        return emit_lambda(node);
      case parser::NodeKind::LetBinding:
        return emit_let(node, HiddenCallClass(node, slot, promote, reference_slot));
      case parser::NodeKind::LambdaCall:
        return emit_lambda_call(node, HiddenCallClass(node, slot, promote, reference_slot));
      case parser::NodeKind::ErrorPlaceholder:
        return unsupported_node("ErrorPlaceholder");
    }
    return unsupported_node("unknown");
  }

  /// Class of the token of a call at `slot` whose result is not taken as a
  /// reference: value at a cell formula's root, array inside an array scope,
  /// else what an operand of its shape takes at `slot` (measured in cell
  /// formulas: `SUM(SEQUENCE(2))` -> 0x62, `SUM(ABS(A1))` -> 0x41,
  /// `ROWS(ABS(A1))` -> 0x61).
  std::uint8_t CallClass(const parser::AstNode& node, Slot slot, bool promote) {
    if (promote) {
      return kPtgValueClass;
    }
    if (in_array_operand_) {
      return kPtgArrayClass;
    }
    if (!measured_calls_) {
      return kPtgValueClass;
    }
    return SlotClass(Slot{slot.letter, true}, MarkedAsArea(slot) || shapes_.of(node) != Shape::kScalar, evaluation_);
  }

  /// In a dynamic-array formula a single value at an `S` operand or an `I`
  /// argument takes the class an area takes there (measured:
  /// `AND(A1>0,A1:A2>0)` and `IFERROR(A1,A1:A2)` store A1 as 0x64).
  bool MarkedAsArea(Slot slot) const { return marked_ && (slot.letter == 'I' || (slot.operand && slot.letter == 'S')); }

  /// Class of a `LET`, `#` or LAMBDA call token: as a built-in's in a cell
  /// formula or a defined name (reference class where the slot takes one
  /// and `reference_slot` says the token may return one), else value.
  std::uint8_t HiddenCallClass(const parser::AstNode& node, Slot slot, bool promote, bool reference_slot) {
    if (!measured_calls_) {
      return kPtgValueClass;
    }
    return reference_slot ? kPtgReferenceClass : CallClass(node, slot, promote);
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
        emit_u8(out_, measured_calls_ ? ClassedPtg(0x23, cls) : std::uint8_t{0x23});  // PtgName
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
  /// `BrtExternSheet` entry of `book` (0: this workbook's own names).
  Expected<void, Error> emit_name_x(std::uint32_t ilbl, std::uint8_t cls, std::uint32_t book = 0U) {
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_range_ixti(kXtiNoSheet, kXtiNoSheet, book));
    emit_u8(out_, ClassedPtg(0x39, cls));  // PtgNameX
    emit_u16(out_, ixti);
    emit_u32(out_, ilbl);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a `LambdaCall` the way Excel 365 stores any LAMBDA
  /// invocation: the callee operand (`Sheet1!Fn` / `[0]!Fn` as `PtgNameX`, or an
  /// inline `LAMBDA(...)` / curried call), then the arguments, then
  /// `PtgFuncVar` with the `id == 255` sentinel and `cparams == arity + 1`.
  Expected<void, Error> emit_lambda_call(const parser::AstNode& node, std::uint8_t cls) {
    const ArrayScope scope(in_array_operand_, measured_calls_ && cls == kPtgArrayClass);
    RETURN_IF_ERROR(emit(node.as_lambda_call_callee()));
    const std::uint32_t arity = node.as_lambda_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_lambda_call_arg(i)));
    }
    return emit_hidden_call_tail(arity + 1, "LambdaCall(arity>254)", cls);
  }

  /// Encodes `Fn(args)` whose callee is a defined name or an in-scope LET /
  /// LAMBDA parameter: the callee `PtgName`, the arguments, then
  /// `PtgFuncVar(255)`, as Excel 365 saves a call to a named LAMBDA.
  Expected<void, Error> emit_name_call(const parser::AstNode& node, std::uint8_t cls) {
    RETURN_IF_ERROR(emit_name_ref(node.as_call_name()));
    const ArrayScope scope(in_array_operand_, measured_calls_ && cls == kPtgArrayClass);
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit(node.as_call_arg(i)));
    }
    return emit_hidden_call_tail(arity + 1, "Call(arity>254, named LAMBDA)", cls);
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
      shapes_.bind(raw_name, Shape::kArray);  // bound to whatever a call passes
    }
    if (measured_calls_) {
      next_slot_ = Slot{'S', false};
    }
    auto status = emit(node.as_lambda_body());
    let_scope_.resize(let_scope_.size() - n);
    shapes_.unbind(n);
    RETURN_IF_ERROR(status);
    return emit_hidden_call_tail(n + 2, "Lambda(too many parameters)", kPtgValueClass);
  }

  /// Emits the `PtgFuncVar(255)` that closes a hidden-name-route call whose
  /// `cparams` operands (callee name-ref included) are already on the stack.
  Expected<void, Error> emit_hidden_call_tail(std::uint32_t cparams, const char* too_many, std::uint8_t cls) {
    if (cparams > 0xFF) {
      return unsupported_node(too_many);
    }
    emit_u8(out_, ClassedPtg(0x22, cls));  // PtgFuncVar result
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
  Expected<void, Error> emit_let(const parser::AstNode& node, std::uint8_t cls) {
    const auto it_let = name_table_.find("_xlfn.LET");
    if (it_let == name_table_.end()) {
      return unsupported_node("LetBinding(_xlfn.LET not registered)");
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the LET name-ref
    emit_u32(out_, it_let->second);
    const ArrayScope scope(in_array_operand_, measured_calls_ && cls == kPtgArrayClass);
    // Bindings and the body take their argument the way an `S` parameter does.
    const Slot slot = measured_calls_ ? Slot{'S', false} : Slot{};
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
      next_slot_ = slot;
      RETURN_IF_ERROR(emit(node.as_let_binding_expr(i)));
      // Subsequent binding expressions and the body can reference this
      // binding; push it onto scope only after its own value expression
      // has been emitted (a binding cannot reference itself).
      let_scope_.emplace_back(raw_name, it_param->second);
      shapes_.bind(raw_name, shapes_.of(node.as_let_binding_expr(i)));
    }
    next_slot_ = slot;
    auto status = emit(node.as_let_body());
    let_scope_.resize(let_scope_.size() - n);
    shapes_.unbind(n);
    RETURN_IF_ERROR(status);
    return emit_hidden_call_tail(1U + (2U * n) + 1U, "LetBinding(too many bindings)", cls);
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
      SpanGrid(first, last);
      return emit_area_ref(first, last, ref.sheet, cls);
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

  /// `PtgArea` for the rectangle `first`:`last`, or `PtgArea3d` on `sheet`.
  Expected<void, Error> emit_area_ref(const parser::Reference& first, const parser::Reference& last,
                                      std::string_view sheet, std::uint8_t cls) {
    if (sheet.empty()) {
      emit_u8(out_, ClassedPtg(0x25, cls));  // PtgArea
      emit_area(out_, first, last);
      return Expected<void, Error>::Ok();
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_single_sheet_ixti(sheet));
    emit_u8(out_, ClassedPtg(0x3B, cls));  // PtgArea3d
    emit_u16(out_, ixti);
    emit_area(out_, first, last);
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

  /// Encodes a cross-workbook reference through the XTI naming its link: a
  /// defined name as `PtgNameX` into the link part's name list, a cell or
  /// rectangle as `PtgRef3d` / `PtgArea3d` over the link's sheets. A span
  /// across sheets keeps reference class, as a local 3-D one does.
  Expected<void, Error> emit_external_ref(const parser::AstNode& node, std::uint8_t cls, bool promote) {
    const std::uint32_t position = external_link_position(sheet_ranges_, node);
    if (position == 0U) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: cross-workbook reference names no saved link",
                        std::string("context=xlsb_ptg_writer book=") + std::string(node.as_external_ref_book()));
    }
    const XlsbLinkTables& link = sheet_ranges_.links[position - 1U];
    const std::string_view name =
        node.kind() == parser::NodeKind::NameRef ? node.as_name() : node.as_external_ref_name();
    if (!name.empty()) {
      const std::uint32_t ilbl = external_name_ilbl(link, name);
      if (ilbl == 0U) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                          "xlsb encoder: external name not in its link's name list",
                          std::string("context=xlsb_ptg_writer name=") + std::string(name));
      }
      return emit_name_x(ilbl, cls, position);
    }
    const std::string_view sheet_end = node.as_external_ref_sheet_end();
    const int first = external_sheet_index(link, node.as_external_ref_sheet());
    const int last = sheet_end.empty() ? first : external_sheet_index(link, sheet_end);
    if (first < 0 || last < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: external sheet not in its link's sheet list",
                        std::string("context=xlsb_ptg_writer sheet=") + std::string(node.as_external_ref_sheet()));
    }
    ASSIGN_OR_RETURN(const std::uint16_t ixti, resolve_range_ixti(first, last, position));
    if (!sheet_end.empty()) {
      cls = promote ? kPtgValueClass : kPtgReferenceClass;
    }
    parser::Reference a = node.as_external_ref_cell();
    parser::Reference b = node.as_external_ref_is_range() ? node.as_external_ref_cell_end() : a;
    if (!node.as_external_ref_is_range() && !a.is_full_col && !a.is_full_row) {
      emit_u8(out_, ClassedPtg(0x3A, cls));  // PtgRef3d
      emit_u16(out_, ixti);
      emit_loc(out_, a);
      return Expected<void, Error>::Ok();
    }
    // A whole column or row is the grid-spanning area.
    SpanGrid(a, b);
    emit_u8(out_, ClassedPtg(0x3B, cls));  // PtgArea3d
    emit_u16(out_, ixti);
    emit_area(out_, a, b);
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

  Expected<std::uint16_t, Error> resolve_range_ixti(int itab_first, int itab_last, std::uint32_t book = 0U) {
    const int ixti = try_resolve_range_ixti(itab_first, itab_last, book);
    if (ixti < 0) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: sheet range not pre-registered",
                        "context=xlsb_ptg_writer");
    }
    return static_cast<std::uint16_t>(ixti);
  }

  /// Non-`Expected` variant for callers (the `PtgArea` fast path) that
  /// fall back to a different encoding on a lookup miss rather than
  /// failing outright. Returns -1 when `(itab_first, itab_last)` is not
  /// in `sheet_ranges_` for `book` (0: this workbook).
  int try_resolve_range_ixti(int itab_first, int itab_last, std::uint32_t book = 0U) const {
    const XtiEntry want{book, itab_first, itab_last};
    for (std::size_t i = 0; i < sheet_ranges_.xti.size(); ++i) {
      if (sheet_ranges_.xti[i] == want) {
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
    const parser::AstNode& operand = node.as_unary_operand();
    // Excel stores a negated number literal as the number it is (measured:
    // `-1` -> PtgNum -1.0, `-0` -> PtgInt 0).
    if (measured_calls_ && node.as_unary_op() == parser::UnaryOp::Minus &&
        operand.kind() == parser::NodeKind::Literal && operand.as_literal().is_number() &&
        parens_.count(&operand) == 0U) {
      return emit_literal(Value::number(-operand.as_literal().as_number()));
    }
    RETURN_IF_ERROR(emit_operand(operand, slot));
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
      // Whole columns (`A:B`) or whole rows (`1:2`) are one area (measured).
      if (b.sheet.empty() && ((a.is_full_col && b.is_full_col) || (a.is_full_row && b.is_full_row))) {
        parser::Reference first = a;
        parser::Reference last = b;
        SpanGrid(first, last);
        return emit_area_ref(first, last, a.sheet, cls);
      }
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
    // A call that can return a reference keeps reference class where its
    // parameter takes one (`ROWS(OFFSET(A1,0,0))` -> 0x22), as measured.
    const bool reference_slot = std::exchange(call_in_reference_slot_, false);
    const std::uint8_t result_class = std::exchange(call_class_, kPtgValueClass);
    // A named LAMBDA's token outside the formulas measured for it stays value class.
    const std::uint8_t name_call_class = reference_slot    ? kPtgReferenceClass
                                         : measured_calls_ ? result_class
                                                           : kPtgValueClass;
    // A localised formula-bar spelling resolves to the name Excel stores
    // before anything is looked up, so both containers agree on the
    // callee and `func_id_table` needs no alias of its own.
    const std::string_view name = canonical_function_name(node.as_call_name());
    const bool reference_result = reference_slot && ReturnsReference(name);
    if (in_let_scope(node.as_call_name())) {
      return emit_name_call(node, name_call_class);
    }
    if (UsesHiddenNameRoute(name)) {
      return emit_future_function_call(node, name,
                                       reference_result  ? kPtgReferenceClass
                                       : measured_calls_ ? result_class
                                                         : kPtgValueClass);
    }
    const XlsbFuncEntry* entry = lookup_func_by_name(name);
    if (entry == nullptr && name_table_.count(std::string(node.as_call_name())) != 0) {
      return emit_name_call(node, name_call_class);
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
    const std::uint8_t token_class = reference_result ? kPtgReferenceClass : result_class;
    const ArrayScope scope(in_array_operand_, measured_calls_ && token_class == kPtgArrayClass);
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
      emit_u8(out_, ClassedPtg(0x22, token_class));  // PtgFuncVar result
      emit_u8(out_, static_cast<std::uint8_t>(arity));
      emit_u16(out_, entry->id);
    } else {
      if (arity != entry->arg_min) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb encoder: fixed-arity function arity mismatch",
                          std::string("context=xlsb_ptg_writer fn=") + std::string(name));
      }
      emit_u8(out_, ClassedPtg(0x21, token_class));  // PtgFunc result
      emit_u16(out_, entry->id);
    }
    patch_gotos();
    return Expected<void, Error>::Ok();
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
  Expected<void, Error> emit_future_function_call(const parser::AstNode& node, std::string_view name,
                                                  std::uint8_t cls) {
    RETURN_IF_ERROR(emit_hidden_callee(name));
    const ArrayScope scope(in_array_operand_, measured_calls_ && cls == kPtgArrayClass);
    const std::uint32_t arity = node.as_call_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      RETURN_IF_ERROR(emit_call_arg(node, name, i));
    }
    // +1 for the name-ref operand.
    return emit_hidden_call_tail(arity + 1, "Call(arity>254, future function)", cls);
  }

  /// Emits `PtgName` for the hidden `_xlfn.` name of the callee `name`.
  Expected<void, Error> emit_hidden_callee(std::string_view name) {
    const auto it = name_table_.find(xlsb_hidden_function_name(name));
    if (it == name_table_.end()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb encoder: hidden-name callee has no BrtName registered",
                        std::string("context=xlsb_ptg_writer fn=") + std::string(name));
    }
    emit_u8(out_, 0x23);  // PtgName (reference-class): the callee name-ref
    emit_u32(out_, it->second);
    return Expected<void, Error>::Ok();
  }

  /// Encodes the postfix `#` spill operator as a call to the hidden
  /// `_xlfn.ANCHORARRAY` name -- Excel's own file-format spelling (see
  /// `eval/dynamic_array/anchor.h` and the matching OOXML storage form in
  /// `ast_format.cpp`'s `StorageEmitter`). Same shape as
  /// `emit_future_function_call`, inlined because the anchor operand comes
  /// from `as_spill_ref_anchor_expr`/`as_spill_ref`, not a `Call` node's
  /// argument list.
  Expected<void, Error> emit_spill_ref(const parser::AstNode& node, std::uint8_t cls) {
    RETURN_IF_ERROR(emit_hidden_callee("ANCHORARRAY"));
    if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
      RETURN_IF_ERROR(emit(*anchor));
    } else {
      // Always a PtgFuncVar argument below, so it stays reference-class.
      RETURN_IF_ERROR(emit_ref(node.as_spill_ref(), kPtgReferenceClass));
    }
    emit_u8(out_, ClassedPtg(0x22, cls));  // PtgFuncVar result
    emit_u8(out_, 2);                      // cparams: name-ref + one anchor operand
    emit_u16(out_, 255);
    return Expected<void, Error>::Ok();
  }

  /// Encodes a written `@` as a call to the hidden `_xlfn.SINGLE` name,
  /// as Excel 365 saves `=@A1` (`23 <ilbl> 24 .. 42 02 ff 00`).
  /// In a defined name's body, which is evaluated as an array, the operand
  /// is a reference and the token one where its slot takes one (measured at
  /// the root: `@Sheet1!$A$1:$A$2` -> 0x3B, 0x22).
  Expected<void, Error> emit_implicit_intersection(const parser::AstNode& node, bool reference_result) {
    RETURN_IF_ERROR(emit_hidden_callee("SINGLE"));
    next_slot_ = Slot{name_body_ ? 'R' : xlsb_parameter_class("SINGLE", 0), false};
    RETURN_IF_ERROR(emit(node.as_implicit_intersection_operand()));
    emit_u8(out_, reference_result ? std::uint8_t{0x22} : ValueClassPtg(0x22));  // PtgFuncVar result
    emit_u8(out_, 2);                                                            // cparams: name-ref + the operand
    emit_u16(out_, 255);
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> emit_array(const parser::AstNode& node, std::uint8_t cls) {
    // The token carries the column and row counts less one as u16s, then
    // bytes Excel leaves unused; the dimensions and elements go into
    // `extra_` (this formula's `rgcb`): rows and columns as u32s, then each
    // element row-major as a tag byte and its payload (measured).
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    emit_u8(out_, ClassedPtg(0x20, cls));  // PtgArray
    emit_u16(out_, static_cast<std::uint16_t>(cols - 1U));
    emit_u16(out_, static_cast<std::uint16_t>(rows - 1U));
    for (int i = 0; i < 10; ++i) {
      emit_u8(out_, 0);
    }
    emit_u32(extra_, rows);
    emit_u32(extra_, cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const parser::AstNode& elem = node.as_array_element(r, c);
        if (elem.kind() == parser::NodeKind::ErrorLiteral) {
          emit_u8(extra_, 4);  // error: the code, then three unused bytes
          emit_u32(extra_, error_wire_code(elem.as_error_literal()));
          continue;
        }
        if (elem.kind() != parser::NodeKind::Literal) {
          return unsupported_node("ArrayLiteral(element)");
        }
        const Value& v = elem.as_literal();
        switch (v.kind()) {
          case ValueKind::Number:
            emit_u8(extra_, 0);
            emit_double(extra_, v.as_number());
            break;
          case ValueKind::Text:
            emit_u8(extra_, 1);  // u16 count + UTF-16LE
            emit_ptg_string_body(extra_, v.as_text());
            break;
          case ValueKind::Bool:
            emit_u8(extra_, 2);
            emit_u8(extra_, v.as_boolean() ? 1U : 0U);
            break;
          case ValueKind::Error:
            emit_u8(extra_, 4);
            emit_u32(extra_, error_wire_code(v.as_error()));
            break;
          default:
            return unsupported_node("ArrayLiteral(element)");
        }
      }
    }
    return Expected<void, Error>::Ok();
  }

  const std::vector<std::string>& sheet_names_;
  const SheetRangeTable& sheet_ranges_;
  const NameTable& name_table_;
  /// Parenthesis pairs per node, as the formula text prints them
  /// (`parser::collect_parenthesized_nodes`).
  std::unordered_map<const parser::AstNode*, std::uint8_t> parens_;
  /// True while emitting the operation a memory token covers.
  bool in_memory_token_ = false;
  /// Where the next node `emit_node` starts sits; set by its parent.
  Slot next_slot_;
  /// Base cell of `PtgRefN` / `PtgAreaN` offsets, when the formula has one.
  const std::optional<PtgBaseCell> base_;
  /// See `PtgEvaluation`; a legacy formula with nothing to intersect is
  /// encoded as a dynamic-array one.
  PtgEvaluation evaluation_;
  Shapes shapes_;
  /// A cell formula or defined name encoded as Excel 365 types it: calls take
  /// the classes measured for them there (`CallClass`), LET / LAMBDA
  /// parameters the class of their slot.
  bool measured_calls_ = false;
  /// A cell formula stored with the dynamic-array mark (`MarkedAsArea`).
  bool marked_ = false;
  /// The next `emit_call`'s slot takes a reference (see `emit_node`).
  bool call_in_reference_slot_ = false;
  /// The next `emit_call`'s token class when it is not a reference (`CallClass`).
  std::uint8_t call_class_ = kPtgValueClass;
  /// Inside a conditional format's array operand (see `emit_node`).
  bool in_array_operand_ = false;
  /// Written `@` nodes a legacy formula stores as nothing.
  std::vector<const parser::AstNode*> implied_at_;
  /// See `emit()`'s `promote` local. Cleared on the first call.
  bool is_root_ = true;
  const bool promote_root_;
  /// A defined name's body: reference-class root, array scope throughout.
  const bool name_body_;
  std::vector<std::uint8_t> out_;
  /// `rgcb`: the array-constant extra-data area `emit_array` appends to.
  std::vector<std::uint8_t> extra_;
  /// Stack of `(parameter name, ilbl)` pairs currently in scope from an
  /// enclosing `LetBinding` (innermost last). Consulted by
  /// `emit_name_ref` before falling back to `name_table_`.
  std::vector<std::pair<std::string_view, std::uint32_t>> let_scope_;
};

}  // namespace

bool formula_is_dynamic_array(const parser::AstNode& root, const NameShapes& names) {
  Shapes shapes(names);
  return MarksDynamic(shapes, root, Slot{}, /*root=*/true);
}

namespace {

/// `node` is the value `name_sets_calc_exp` looks for, or reaches one through operators.
bool ValueSetsCalcExp(const parser::AstNode& node,
                      const std::unordered_map<const parser::AstNode*, std::uint8_t>& parens) {
  using parser::NodeKind;
  if (parens.count(&node) != 0U) {
    return false;
  }
  switch (node.kind()) {
    case NodeKind::Call: {
      // A call through a name (a LAMBDA, or one Excel does not know) counts as a LAMBDA call.
      const std::string_view name = canonical_function_name(node.as_call_name());
      return !IsBuiltin(name) || xlsb_sets_calc_exp(name);
    }
    case NodeKind::LetBinding:
    case NodeKind::Lambda:
    case NodeKind::LambdaCall:
    case NodeKind::SpillRef:
      return true;
    case NodeKind::UnaryOp:
      return ValueSetsCalcExp(node.as_unary_operand(), parens);
    case NodeKind::BinaryOp:
      return ValueSetsCalcExp(node.as_binary_lhs(), parens) || ValueSetsCalcExp(node.as_binary_rhs(), parens);
    case NodeKind::RangeOp:
      return ValueSetsCalcExp(node.as_range_lhs(), parens) || ValueSetsCalcExp(node.as_range_rhs(), parens);
    default:
      return false;
  }
}

/// `node` refers, anywhere within it, to a defined name `names` flags.
bool RefersToCalcExpName(const parser::AstNode& node, const NameShapes& names) {
  const bool name_ref =
      node.kind() == parser::NodeKind::NameRef ||
      (node.kind() == parser::NodeKind::Call && !IsBuiltin(canonical_function_name(node.as_call_name())));
  if (name_ref && names && names(node).calc_exp) {
    return true;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (RefersToCalcExpName(*child, names)) {
      return true;
    }
  }
  return false;
}

/// Walker behind `formula_always_calculates`; `bound` holds the LET / LAMBDA
/// names in scope, which a call may go through without being unknown.
bool AlwaysCalculates(const parser::AstNode& node, const NameShapes& names, std::vector<std::string_view>& bound) {
  using parser::NodeKind;
  auto is_bound = [&](std::string_view name) {
    return std::any_of(bound.begin(), bound.end(),
                       [&](std::string_view b) { return strings::case_insensitive_eq(b, name); });
  };
  switch (node.kind()) {
    case NodeKind::Call: {
      const std::string_view name = node.as_call_name();
      if (parser::is_volatile_function_name(name)) {
        return true;
      }
      if (!IsBuiltin(canonical_function_name(name)) && !is_bound(name)) {
        const NameShape shape = names ? names(node) : NameShape{};
        if (!shape.defined || shape.always_calculates) {
          return true;
        }
      }
      break;
    }
    case NodeKind::NameRef:
    case NodeKind::ExternalRef:
      if ((node.kind() == NodeKind::ExternalRef && !parser::is_self_book_name_ref(node)) ||
          (node.kind() == NodeKind::NameRef && is_bound(node.as_name()))) {
        break;
      }
      if (names && names(node).always_calculates) {
        return true;
      }
      break;
    case NodeKind::LambdaCall: {
      // A call through a reference (`A1(1)`) calls nothing Excel knows.
      const NodeKind callee = node.as_lambda_call_callee().kind();
      if (callee != NodeKind::Lambda && callee != NodeKind::LambdaCall) {
        return true;
      }
      break;
    }
    case NodeKind::LetBinding: {
      const std::size_t depth = bound.size();
      bool found = false;
      for (std::uint32_t i = 0; i < node.as_let_binding_count() && !found; ++i) {
        found = AlwaysCalculates(node.as_let_binding_expr(i), names, bound);
        bound.push_back(node.as_let_binding_name(i));
      }
      found = found || AlwaysCalculates(node.as_let_body(), names, bound);
      bound.resize(depth);
      return found;
    }
    case NodeKind::Lambda: {
      const std::size_t depth = bound.size();
      for (std::uint32_t i = 0; i < node.as_lambda_param_count(); ++i) {
        bound.push_back(node.as_lambda_param(i));
      }
      const bool found = AlwaysCalculates(node.as_lambda_body(), names, bound);
      bound.resize(depth);
      return found;
    }
    default:
      break;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (AlwaysCalculates(*child, names, bound)) {
      return true;
    }
  }
  return false;
}

}  // namespace

bool name_sets_calc_exp(const parser::AstNode& root, const NameShapes& names) {
  std::unordered_map<const parser::AstNode*, std::uint8_t> parens;
  parser::collect_parenthesized_nodes(root, parens);
  return ValueSetsCalcExp(root, parens) || RefersToCalcExpName(root, names);
}

bool formula_calls_volatile(const parser::AstNode& root) {
  return ContainsVolatileCall(root);
}

bool formula_always_calculates(const parser::AstNode& root, const NameShapes& names) {
  std::vector<std::string_view> bound;
  return AlwaysCalculates(root, names, bound);
}

std::vector<const parser::AstNode*> legacy_intersections(const parser::AstNode& root, const NameShapes& names) {
  std::vector<const parser::AstNode*> out;
  LegacyIntersections(names, out).walk(root, Slot{}, /*root=*/true);
  return out;
}

Expected<EncodedFormula, Error> encode_ptgs(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                                            const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                            PtgRootClass root_class, std::optional<PtgBaseCell> base,
                                            PtgEvaluation evaluation, const NameShapes& names) {
  Encoder enc(node, sheet_names, sheet_ranges, name_table, root_class, base, evaluation, names);
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
