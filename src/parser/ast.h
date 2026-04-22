// Copyright 2026 libraz. Licensed under the MIT License.
//
// Parser abstract syntax tree.
//
// `AstNode` is the parser's tagged-union output. The 17 `NodeKind` variants
// cover every Excel 365 surface construct: literals, references (including
// external workbook and structured table refs), unary/binary/range/union/
// intersect operators, the `@` implicit-intersection wrapper, function calls,
// inline array literals, `LAMBDA` and `LET` forms, immediately-invoked lambda
// calls, and source-level error literals such as `#DIV/0!`.
//
// Storage model: every node and every variable-arity child array lives in an
// `Arena`. `AstNode` is trivially destructible (enforced by `static_assert`)
// so the arena can release its memory wholesale without invoking destructors.
// Strings (sheet names, function names, identifier names, lambda parameters)
// are interned into the same arena via `Arena::intern`, so the AST owns its
// strings independently of the original source buffer.
//
// Nodes are constructed exclusively via the `make_*` free functions; the
// `AstNode` constructor is private. Each factory returns `nullptr` if any
// arena allocation fails. The `range_` field is left default-initialised by
// the factories; the parser sets it via `set_range()` once it has scanned the
// node's source span.
//
// See `backup/plans/02-calc-engine.md` §2.3 for the authoritative AST shape.

#ifndef FORMULON_PARSER_AST_H_
#define FORMULON_PARSER_AST_H_

#include <cstdint>
#include <string_view>
#include <type_traits>

#include "parser/reference.h"
#include "parser/token.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {

/// Discriminator tag for every AST variant.
///
/// The numeric assignments are stable: switch tables in the dumper and (later)
/// the compiler dispatch on these values, so do not reorder.
enum class NodeKind : std::uint8_t {
  Literal = 0,
  Ref = 1,
  ExternalRef = 2,
  StructuredRef = 3,
  NameRef = 4,
  UnaryOp = 5,
  BinaryOp = 6,
  RangeOp = 7,
  UnionOp = 8,
  IntersectOp = 9,
  ImplicitIntersection = 10,
  Call = 11,
  ArrayLiteral = 12,
  Lambda = 13,
  LetBinding = 14,
  LambdaCall = 15,
  ErrorLiteral = 16,
};

/// Binary operator catalog covering arithmetic, concat, and comparisons.
enum class BinOp : std::uint8_t {
  Add,
  Sub,
  Mul,
  Div,
  Pow,
  Concat,
  Eq,
  NotEq,
  Lt,
  LtEq,
  Gt,
  GtEq,
};

/// Unary operator catalog. `Plus` and `Minus` are prefix; `Percent` is postfix.
enum class UnaryOp : std::uint8_t {
  Plus,
  Minus,
  Percent,
};

/// Modifier applied to a structured (table) reference.
///
/// `None` is the default for `Table[col]`. The `#`-prefixed variants are the
/// special area selectors (`#Headers`, `#Data`, `#Totals`, `#All`); `At`
/// represents the single-row implicit-intersection form `Table[@col]`.
enum class StructuredRefModifier : std::uint8_t {
  None = 0,
  At = 1,
  Headers = 2,
  Data = 3,
  Totals = 4,
  All = 5,
};

/// Tagged-union AST node.
///
/// Construction is restricted to the `make_*` factories declared below. Each
/// per-kind accessor (`as_*`) precondition-checks the discriminator via
/// `FM_CHECK`: calling, e.g., `as_binary_op()` on a `Literal` aborts the
/// process. Callers must branch on `kind()` first.
class AstNode final {
 public:
  /// Returns the discriminator tag.
  NodeKind kind() const noexcept { return kind_; }

  /// Returns the source span set by the parser. Default-initialised (all
  /// zeros) until `set_range()` is invoked.
  TextRange range() const noexcept { return range_; }

  /// Records the source span for this node. The parser typically calls this
  /// immediately after a successful factory invocation.
  void set_range(TextRange r) noexcept { range_ = r; }

  // --- Literal -------------------------------------------------------------
  const Value& as_literal() const;

  // --- Ref -----------------------------------------------------------------
  const Reference& as_ref() const;

  // --- ExternalRef ---------------------------------------------------------
  std::uint32_t as_external_ref_book_id() const;
  std::string_view as_external_ref_sheet() const;
  const Reference& as_external_ref_cell() const;

  // --- StructuredRef -------------------------------------------------------
  std::string_view as_structured_ref_table() const;
  std::string_view as_structured_ref_column() const;
  StructuredRefModifier as_structured_ref_modifier() const;

  // --- NameRef -------------------------------------------------------------
  std::string_view as_name() const;

  // --- UnaryOp -------------------------------------------------------------
  UnaryOp as_unary_op() const;
  const AstNode& as_unary_operand() const;

  // --- BinaryOp ------------------------------------------------------------
  BinOp as_binary_op() const;
  const AstNode& as_binary_lhs() const;
  const AstNode& as_binary_rhs() const;

  // --- RangeOp -------------------------------------------------------------
  const AstNode& as_range_lhs() const;
  const AstNode& as_range_rhs() const;

  // --- UnionOp -------------------------------------------------------------
  std::uint32_t as_union_arity() const;
  const AstNode& as_union_child(std::uint32_t i) const;

  // --- IntersectOp ---------------------------------------------------------
  const AstNode& as_intersect_lhs() const;
  const AstNode& as_intersect_rhs() const;

  // --- ImplicitIntersection -----------------------------------------------
  const AstNode& as_implicit_intersection_operand() const;

  // --- Call ----------------------------------------------------------------
  std::string_view as_call_name() const;
  std::uint32_t as_call_arity() const;
  const AstNode& as_call_arg(std::uint32_t i) const;

  // --- ArrayLiteral --------------------------------------------------------
  std::uint32_t as_array_rows() const;
  std::uint32_t as_array_cols() const;
  const AstNode& as_array_element(std::uint32_t row, std::uint32_t col) const;

  // --- Lambda --------------------------------------------------------------
  std::uint32_t as_lambda_param_count() const;
  std::string_view as_lambda_param(std::uint32_t i) const;
  const AstNode& as_lambda_body() const;

  // --- LetBinding ----------------------------------------------------------
  std::uint32_t as_let_binding_count() const;
  std::string_view as_let_binding_name(std::uint32_t i) const;
  const AstNode& as_let_binding_expr(std::uint32_t i) const;
  const AstNode& as_let_body() const;

  // --- LambdaCall ----------------------------------------------------------
  const AstNode& as_lambda_call_callee() const;
  std::uint32_t as_lambda_call_arity() const;
  const AstNode& as_lambda_call_arg(std::uint32_t i) const;

  // --- ErrorLiteral --------------------------------------------------------
  ErrorCode as_error_literal() const;

 private:
  AstNode() = default;

  // Arena::create<AstNode>() invokes the private default constructor; mark
  // Arena as a friend so the placement-new in `create<T>()` can reach it.
  friend class formulon::Arena;

  // Each factory writes the discriminator, payload, and (for variable-arity
  // nodes) installs arena-owned child / name arrays.
  friend AstNode* make_literal(Arena&, Value);
  friend AstNode* make_ref(Arena&, const Reference&);
  friend AstNode* make_external_ref(Arena&, std::uint32_t, std::string_view, const Reference&);
  friend AstNode* make_structured_ref(Arena&, std::string_view, std::string_view, StructuredRefModifier);
  friend AstNode* make_name_ref(Arena&, std::string_view);
  friend AstNode* make_unary_op(Arena&, UnaryOp, AstNode*);
  friend AstNode* make_binary_op(Arena&, BinOp, AstNode*, AstNode*);
  friend AstNode* make_range_op(Arena&, AstNode*, AstNode*);
  friend AstNode* make_union_op(Arena&, const AstNode* const*, std::uint32_t);
  friend AstNode* make_intersect_op(Arena&, AstNode*, AstNode*);
  friend AstNode* make_implicit_intersection(Arena&, AstNode*);
  friend AstNode* make_call(Arena&, std::string_view, const AstNode* const*, std::uint32_t);
  friend AstNode* make_array_literal(Arena&, std::uint32_t, std::uint32_t, const AstNode* const*);
  friend AstNode* make_lambda(Arena&, const std::string_view*, std::uint32_t, AstNode*);
  friend AstNode* make_let_binding(Arena&, const std::string_view*, const AstNode* const*, std::uint32_t, AstNode*);
  friend AstNode* make_lambda_call(Arena&, AstNode*, const AstNode* const*, std::uint32_t);
  friend AstNode* make_error_literal(Arena&, ErrorCode);

  // --- Per-kind payload structs --------------------------------------------
  // Each is trivially destructible; pointer arrays are arena-owned. The
  // ExternalRef payload (book id + sheet view + Reference) is large enough
  // that we store it through an arena-allocated pointer to keep the union
  // small; every other variant fits inline.
  struct ExternalRefPayload {
    std::uint32_t book_id;
    std::string_view sheet;
    Reference cell;
  };
  struct StructuredRefPayload {
    std::string_view table;
    std::string_view column;
    StructuredRefModifier modifier;
  };
  struct UnaryPayload {
    UnaryOp op;
    const AstNode* operand;
  };
  struct BinaryPayload {
    BinOp op;
    const AstNode* lhs;
    const AstNode* rhs;
  };
  struct RangePayload {
    const AstNode* lhs;
    const AstNode* rhs;
  };
  struct VariadicPayload {
    const AstNode* const* children;
    std::uint32_t count;
  };
  struct CallPayload {
    std::string_view name;
    const AstNode* const* args;
    std::uint32_t arity;
  };
  struct ArrayPayload {
    const AstNode* const* elements;  // row-major, rows*cols entries
    std::uint32_t rows;
    std::uint32_t cols;
  };
  struct LambdaPayload {
    const std::string_view* params;
    std::uint32_t param_count;
    const AstNode* body;
  };
  struct LetPayload {
    const std::string_view* names;
    const AstNode* const* exprs;
    std::uint32_t binding_count;
    const AstNode* body;
  };
  struct LambdaCallPayload {
    const AstNode* callee;
    const AstNode* const* args;
    std::uint32_t arity;
  };

  // Tagged union of every payload variant. We deliberately store the larger
  // payloads (Lambda, Let) as pointer + count tuples so the union itself stays
  // small. The Literal payload is inlined because Value is only 16 bytes and
  // is trivially copyable.
  union Payload {
    Value literal;
    Reference ref;
    const ExternalRefPayload* external_ref;
    StructuredRefPayload structured_ref;
    std::string_view name;
    UnaryPayload unary;
    BinaryPayload binary;
    RangePayload range;
    VariadicPayload variadic;  // Used by UnionOp.
    CallPayload call;
    ArrayPayload array;
    LambdaPayload lambda;
    LetPayload let;
    LambdaCallPayload lambda_call;
    ErrorCode error_literal;

    // The factories always overwrite the active member before any accessor
    // observes it, so we leave the union in an indeterminate state on
    // construction. The user-provided ctor is necessary because some members
    // (e.g. Value) have non-trivial default constructors; the destructor is
    // defaulted because every member is trivially destructible.
    Payload() noexcept {}
    ~Payload() = default;
  };

  NodeKind kind_ = NodeKind::Literal;
  TextRange range_{};
  Payload data_;
};

static_assert(std::is_trivially_destructible_v<AstNode>,
              "AstNode must be trivially destructible to live in an Arena");

// Size budget: this asserts our payload layout stays compact enough that the
// AST does not balloon working-set memory. 64 bytes is generous; the actual
// sizeof is reported in unit tests for visibility.
static_assert(sizeof(AstNode) <= 64, "AstNode exceeds 64-byte size budget");

// -- Factory free functions --------------------------------------------------
//
// Every factory takes the arena as its first argument. Strings passed via
// `string_view` are interned into the arena; pointer arrays are deep-copied
// into arena-owned storage, so the caller may free or mutate the original
// argv after the call returns. Returns `nullptr` only on arena allocation
// failure.

/// Builds a `Literal` node wrapping the scalar `v`.
AstNode* make_literal(Arena& arena, Value v);

/// Builds a `Ref` node from the structural reference `r`. The sheet view is
/// re-interned into `arena` so the caller's storage need not outlive the AST.
AstNode* make_ref(Arena& arena, const Reference& r);

/// Builds an `ExternalRef` node referencing `[book_id]sheet!cell`.
AstNode* make_external_ref(Arena& arena, std::uint32_t book_id, std::string_view sheet, const Reference& cell);

/// Builds a `StructuredRef` node.  `column` may be empty when the reference
/// targets the whole table.  `modifier` is `None` for plain `Table[col]`.
AstNode* make_structured_ref(Arena& arena, std::string_view table, std::string_view column,
                             StructuredRefModifier modifier);

/// Builds a `NameRef` node referencing the defined name `name`.
AstNode* make_name_ref(Arena& arena, std::string_view name);

/// Builds a `UnaryOp` node. `operand` must be non-null.
AstNode* make_unary_op(Arena& arena, UnaryOp op, AstNode* operand);

/// Builds a `BinaryOp` node. `lhs` and `rhs` must both be non-null.
AstNode* make_binary_op(Arena& arena, BinOp op, AstNode* lhs, AstNode* rhs);

/// Builds a `RangeOp` node. `lhs` and `rhs` must both be non-null.
AstNode* make_range_op(Arena& arena, AstNode* lhs, AstNode* rhs);

/// Builds a `UnionOp` node from the `count` children pointed at by
/// `children`. The pointer array is deep-copied; `count` must be `>= 2`.
AstNode* make_union_op(Arena& arena, const AstNode* const* children, std::uint32_t count);

/// Builds an `IntersectOp` node. `lhs` and `rhs` must both be non-null.
AstNode* make_intersect_op(Arena& arena, AstNode* lhs, AstNode* rhs);

/// Builds an `ImplicitIntersection` node wrapping `operand`.
AstNode* make_implicit_intersection(Arena& arena, AstNode* operand);

/// Builds a `Call` node. `args` may be null when `arity == 0`.
AstNode* make_call(Arena& arena, std::string_view name, const AstNode* const* args, std::uint32_t arity);

/// Builds an `ArrayLiteral`. `elems_row_major` must contain `rows * cols`
/// non-null entries; both dimensions must be `>= 1`.
AstNode* make_array_literal(Arena& arena, std::uint32_t rows, std::uint32_t cols,
                            const AstNode* const* elems_row_major);

/// Builds a `Lambda` node. `params` may be null when `param_count == 0`.
AstNode* make_lambda(Arena& arena, const std::string_view* params, std::uint32_t param_count, AstNode* body);

/// Builds a `LetBinding`. Each `i < binding_count` pairs `names[i]` with
/// `exprs[i]`. `binding_count` must be `>= 1`.
AstNode* make_let_binding(Arena& arena, const std::string_view* names, const AstNode* const* exprs,
                          std::uint32_t binding_count, AstNode* body);

/// Builds a `LambdaCall` node invoking `callee` with `arity` arguments.
AstNode* make_lambda_call(Arena& arena, AstNode* callee, const AstNode* const* args, std::uint32_t arity);

/// Builds an `ErrorLiteral` node.
AstNode* make_error_literal(Arena& arena, ErrorCode code);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_H_
