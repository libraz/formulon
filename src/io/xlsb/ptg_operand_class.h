//
// Internal to the Ptg encoder (`ptg_writer.cpp`) and the formula-flag
// analyses: where a node sits (`Slot`), the class bits a reference takes
// there (`SlotClass`), and the result `Shape` Excel judges a node to have.

#ifndef FORMULON_IO_XLSB_PTG_OPERAND_CLASS_H_
#define FORMULON_IO_XLSB_PTG_OPERAND_CLASS_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/ptg_writer.h"
#include "parser/ast.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

// Formula arguments which denote cells and ranges use the reference-class
// base bytes. The result of a function call, in contrast, is a value-class
// Ptg (its low 5-bit type plus class bits `0x40`). Using `| 0x40` on the
// already class-marked base byte had emitted array-class
// references (for example 0x65 instead of PtgArea 0x25), while leaving
// function results in the reference class. Excel repairs those streams.
constexpr std::uint8_t kPtgValueClass = 0x40;

constexpr std::uint8_t kPtgReferenceClass = 0x20;
constexpr std::uint8_t kPtgArrayClass = 0x60;

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
std::uint8_t SlotClass(Slot slot, bool area, PtgEvaluation evaluation = PtgEvaluation::kDynamicArray);

/// A node's result as Excel 365 judges it statically (measured against the
/// dynamic-array marks it gives formulas and the `@`s it shows in legacy
/// ones): one value, a reference that may cover several cells, or an array.
enum class Shape : std::uint8_t { kScalar, kReference, kArray };

/// True when a call to `name` is a built-in Excel knows, by id or by hidden name.
bool IsBuiltin(std::string_view name);

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

  Shape of(const parser::AstNode& node);

 private:
  const std::pair<std::string_view, Shape>* binding(const parser::AstNode& node) const;

  NameShape defined(const parser::AstNode& node) const { return names_ ? names_(node) : NameShape{}; }

  Shape name_shape(const parser::AstNode& node) const;

  /// `node` is a positive integer constant Excel stores as `PtgInt` (equal to
  /// `want` unless that is 0), literally or through a name.
  bool positive_int(const parser::AstNode& node, std::uint32_t want = 0U) const;

  /// Rows and columns of the one plain rectangle `node` denotes, if it does.
  std::optional<std::pair<std::uint32_t, std::uint32_t>> extent(const parser::AstNode& node) const;

  Shape call_shape(const parser::AstNode& node);

  const NameShapes& names_;
  const bool lifting_;
  std::vector<std::pair<std::string_view, Shape>> scope_;
};

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

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_OPERAND_CLASS_H_
