//
// Pivot formula host backed by the standard parser and tree walker.
//
// @size-budget: 20 KB

#include "eval/pivot_formula.h"

#include <algorithm>
#include <cassert>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/collation.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/name_env.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"

namespace formulon::eval {

namespace {

constexpr std::uint32_t kMaxFormulaUnits = 8192;
constexpr std::string_view kSyntheticPrefix = "_fmpv";

// Functions a pivot formula may not call (MS-OI29500 pivot-function-invalid-name).
constexpr std::string_view kDisallowedFunctions[] = {
    "GETPIVOTDATA", "DCOUNT",  "DCOUNTA", "DSUM",  "DAVERAGE", "DMIN",  "MAX",  "DPRODUCT",
    "DSTDEV",       "DSTDEVP", "DVAR",    "DVARP", "INDEX",    "RAND",  "NOW",  "AREAS",
    "ROWS",         "COLUMNS", "OFFSET",  "CELL",  "INDIRECT", "TODAY", "INFO", "RANDBETWEEN",
};

Error invalid(const char* reason, std::string_view context = {}) {
  return make_error(FormulonErrorCode::kInvalidArgument, reason, std::string(context));
}

// Bytes that extend an identifier run; a superset of the tokenizer's name bytes
// so a run never stops inside a cell reference or number.
bool is_run_byte(unsigned char c) {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.' ||
         c == '\\' || c == '?' || c == '$' || c >= 0x80;
}

bool is_error_literal_byte(unsigned char c) {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '/' ||
         c == '!' || c == '?';
}

bool is_reserved_name(std::string_view name) {
  return strings::case_insensitive_eq(name, "All") || strings::case_insensitive_eq(name, "Blank");
}

bool names_equal(std::string_view a, std::string_view b) {
  return collate(a, b, current_eval_profile(), CollationStrength::kDynamic) == 0;
}

void append_synthetic(std::string& out, std::size_t index) {
  out.append(kSyntheticPrefix);
  char digits[20];
  std::size_t n = 0;
  do {
    digits[n++] = static_cast<char>('0' + index % 10);
    index /= 10;
  } while (index != 0);
  while (n > 0) {
    out.push_back(digits[--n]);
  }
}

std::string_view trim_spaces(std::string_view s) {
  while (!s.empty() && s.front() == ' ') {
    s.remove_prefix(1);
  }
  while (!s.empty() && s.back() == ' ') {
    s.remove_suffix(1);
  }
  return s;
}

// Rewrites pivot names in a formula body to synthetic identifiers and records
// the reference each one stands for.
class Rewriter {
 public:
  Rewriter(const pivot::FormulaScope& scope, std::vector<pivot::FormulaRef>* refs) : scope_(scope), refs_(refs) {}

  Expected<std::string, Error> run(std::string_view src) {
    src_ = src;
    std::size_t i = 0;
    while (i < src_.size()) {
      const std::size_t start = i;
      const auto c = static_cast<unsigned char>(src_[i]);
      if (c == '"') {
        i = copy_string_literal(i);
      } else if (c == '#') {
        out_.push_back('#');
        ++i;
        while (i < src_.size() && is_error_literal_byte(static_cast<unsigned char>(src_[i]))) {
          out_.push_back(src_[i++]);
        }
      } else if (c == '\'') {
        auto next = quoted_name(i);
        if (!next) {
          return next.error();
        }
        i = next.value();
      } else if (is_run_byte(c)) {
        auto next = bare_run(i);
        if (!next) {
          return next.error();
        }
        i = next.value();
      } else {
        out_.push_back(src_[i++]);
      }
      append_canonical(src_.substr(start, i - start));
    }
    return std::move(out_);
  }

  /// The source with every bare item name qualified as `Field[Item]`.
  std::string& canonical() { return canonical_; }

 private:
  void append_canonical(std::string_view token) {
    if (!qualify_) {
      canonical_.append(token);
      return;
    }
    qualify_ = false;
    const std::string& field = scope_.cache->fields()[*scope_.item_field].name;
    if (needs_quotes(field)) {
      canonical_.push_back('\'');
      for (const char ch : field) {
        canonical_.append(ch == '\'' ? 2U : 1U, ch);
      }
      canonical_.push_back('\'');
    } else {
      canonical_.append(field);
    }
    canonical_.push_back('[');
    canonical_.append(token);
    canonical_.push_back(']');
  }

  // A field name stays bare in `Field[...]` only when it reads back as one run.
  static bool needs_quotes(std::string_view name) {
    if (name.empty() || strings::case_insensitive_starts_with(name, kSyntheticPrefix)) {
      return true;
    }
    for (const char ch : name) {
      if (!is_run_byte(static_cast<unsigned char>(ch))) {
        return true;
      }
    }
    return false;
  }

  std::size_t copy_string_literal(std::size_t i) {
    out_.push_back(src_[i++]);
    while (i < src_.size()) {
      out_.push_back(src_[i]);
      if (src_[i++] == '"') {
        if (i < src_.size() && src_[i] == '"') {
          out_.push_back(src_[i++]);
          continue;
        }
        break;
      }
    }
    return i;
  }

  // Reads `'...'` at `i` (`''` is a literal quote); returns the index past it.
  std::optional<std::size_t> read_quoted(std::size_t i, std::string* name) const {
    ++i;
    while (i < src_.size()) {
      if (src_[i] == '\'') {
        if (i + 1 < src_.size() && src_[i + 1] == '\'') {
          name->push_back('\'');
          i += 2;
          continue;
        }
        return i + 1;
      }
      name->push_back(src_[i++]);
    }
    return std::nullopt;
  }

  Expected<std::size_t, Error> quoted_name(std::size_t i) {
    std::string name;
    const std::optional<std::size_t> end = read_quoted(i, &name);
    if (!end) {
      return invalid("formula has an unterminated quoted name");
    }
    if (*end < src_.size() && src_[*end] == '[') {
      return bracket_ref(name, *end);
    }
    if (auto st = bind_name(name, /*quoted=*/true); !st) {
      return st.error();
    }
    return *end;
  }

  Expected<std::size_t, Error> bare_run(std::size_t i) {
    std::size_t end = i;
    while (end < src_.size() && is_run_byte(static_cast<unsigned char>(src_[end]))) {
      ++end;
    }
    const std::string_view run = src_.substr(i, end - i);
    if (end < src_.size() && src_[end] == '(') {
      out_.append(run);
      return end;
    }
    if (strings::case_insensitive_starts_with(run, kSyntheticPrefix)) {
      return invalid("formula uses a reserved identifier", run);
    }
    if (end < src_.size() && src_[end] == '[') {
      return bracket_ref(run, end);
    }
    if (auto st = bind_name(run, /*quoted=*/false); !st) {
      return st.error();
    }
    return end;
  }

  // Binds a bare or quoted name; an unmatched bare run is copied verbatim.
  Expected<void, Error> bind_name(std::string_view name, bool quoted) {
    if (is_reserved_name(name)) {
      return invalid("All and Blank cannot be used as names in a pivot formula", name);
    }
    if (scope_.item_field) {
      if (const std::optional<std::size_t> item = find_item(name)) {
        emit(pivot::FormulaRef{pivot::FormulaRef::Kind::Item, *scope_.item_field, *item, 0});
        qualify_ = true;
        return {};
      }
      if (find_field(name)) {
        return invalid("a calculated item formula cannot name a field", name);
      }
    } else if (const std::optional<std::size_t> field = find_field(name)) {
      emit(pivot::FormulaRef{pivot::FormulaRef::Kind::Field, *field, 0, 0});
      return {};
    }
    if (quoted) {
      return invalid("formula uses an unknown name", name);
    }
    out_.append(name);
    return {};
  }

  // `Field[Item]`, `Field['Item']`, `Field[n]`, `Field[+n]` / `Field[-n]`; `open` is the `[`.
  Expected<std::size_t, Error> bracket_ref(std::string_view field_name, std::size_t open) {
    if (!scope_.item_field) {
      return invalid("Field[Item] references are allowed only in calculated item formulas", field_name);
    }
    if (is_reserved_name(field_name)) {
      return invalid("All and Blank cannot be used as names in a pivot formula", field_name);
    }
    const std::optional<std::size_t> field = find_field(field_name);
    if (!field) {
      return invalid("formula uses an unknown field", field_name);
    }
    if (*field != *scope_.item_field) {
      return invalid("a calculated item formula can reference only items of its own field", field_name);
    }

    std::size_t i = open + 1;
    while (i < src_.size() && src_[i] == ' ') {
      ++i;
    }
    std::string quoted;
    std::string_view content;
    bool is_quoted = false;
    if (i < src_.size() && src_[i] == '\'') {
      const std::optional<std::size_t> end = read_quoted(i, &quoted);
      if (!end) {
        return invalid("formula has an unterminated quoted name");
      }
      i = *end;
      while (i < src_.size() && src_[i] == ' ') {
        ++i;
      }
      if (i >= src_.size() || src_[i] != ']') {
        return invalid("formula has a malformed Field[Item] reference", field_name);
      }
      content = quoted;
      is_quoted = true;
    } else {
      const std::size_t close = src_.find(']', i);
      if (close == std::string_view::npos) {
        return invalid("formula has an unterminated Field[Item] reference", field_name);
      }
      content = trim_spaces(src_.substr(i, close - i));
      i = close;
    }
    const std::size_t after = i + 1;

    std::int32_t position = 0;
    const Position parsed = is_quoted ? Position::kNotPosition : parse_position(content, &position);
    if (parsed == Position::kOverflow) {
      return invalid("formula has an item position out of range", content);
    }
    if (parsed == Position::kOk) {
      const bool signed_offset = content.front() == '+' || content.front() == '-';
      emit(pivot::FormulaRef{signed_offset ? pivot::FormulaRef::Kind::ItemOffset : pivot::FormulaRef::Kind::ItemAt,
                             *field, 0, position});
      return after;
    }
    if (is_reserved_name(content)) {
      return invalid("All and Blank cannot be used as names in a pivot formula", content);
    }
    const std::optional<std::size_t> item = find_item(content);
    if (!item) {
      return invalid("formula uses an unknown item", content);
    }
    emit(pivot::FormulaRef{pivot::FormulaRef::Kind::Item, *field, *item, 0});
    return after;
  }

  enum class Position : std::uint8_t { kNotPosition, kOverflow, kOk };

  // Parses `[+-]digits`; anything else is kNotPosition.
  static Position parse_position(std::string_view s, std::int32_t* out) {
    bool negative = false;
    if (!s.empty() && (s.front() == '+' || s.front() == '-')) {
      negative = s.front() == '-';
      s.remove_prefix(1);
    }
    if (s.empty()) {
      return Position::kNotPosition;
    }
    std::int64_t v = 0;
    bool overflow = false;
    for (const char ch : s) {
      if (ch < '0' || ch > '9') {
        return Position::kNotPosition;
      }
      if (!overflow) {
        v = v * 10 + (ch - '0');
        overflow = v > INT32_MAX;
      }
    }
    if (overflow) {
      return Position::kOverflow;
    }
    *out = static_cast<std::int32_t>(negative ? -v : v);
    return Position::kOk;
  }

  std::optional<std::size_t> find_field(std::string_view name) const {
    const auto& fields = scope_.cache->fields();
    for (std::size_t f = 0; f < fields.size(); ++f) {
      if (names_equal(fields[f].name, name)) {
        return f;
      }
    }
    return std::nullopt;
  }

  std::optional<std::size_t> find_item(std::string_view name) const {
    const auto& items = scope_.cache->fields()[*scope_.item_field].shared_items;
    for (std::size_t k = 0; k < items.size(); ++k) {
      if (items[k].is_text() && names_equal(items[k].as_text(), name)) {
        return k;
      }
    }
    return std::nullopt;
  }

  // Appends the synthetic identifier for `ref`, reusing the slot of an equal ref.
  void emit(const pivot::FormulaRef& ref) {
    std::size_t index = 0;
    while (index < refs_->size()) {
      const pivot::FormulaRef& r = (*refs_)[index];
      if (r.kind == ref.kind && r.field == ref.field && r.item == ref.item && r.position == ref.position) {
        break;
      }
      ++index;
    }
    if (index == refs_->size()) {
      refs_->push_back(ref);
    }
    append_synthetic(out_, index);
  }

  const pivot::FormulaScope& scope_;
  std::vector<pivot::FormulaRef>* refs_;
  std::string_view src_;
  std::string out_;
  std::string canonical_;
  /// Set when the token being scanned is a bare item name to qualify.
  bool qualify_ = false;
};

// True when `name` is a synthetic identifier below `ref_count`.
bool is_bound_synthetic(std::string_view name, std::size_t ref_count) {
  if (!strings::case_insensitive_starts_with(name, kSyntheticPrefix) || name.size() == kSyntheticPrefix.size()) {
    return false;
  }
  std::size_t index = 0;
  for (const char ch : name.substr(kSyntheticPrefix.size())) {
    if (ch < '0' || ch > '9' || index > ref_count) {
      return false;
    }
    index = index * 10 + static_cast<std::size_t>(ch - '0');
  }
  return index < ref_count;
}

Expected<void, Error> check_node(const parser::AstNode& node, std::size_t ref_count) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
      return {};
    case parser::NodeKind::NameRef:
      if (!node.as_name_sheet().empty() || !is_bound_synthetic(node.as_name(), ref_count)) {
        return invalid("formula must not use defined names or unknown names", node.as_name());
      }
      return {};
    case parser::NodeKind::UnaryOp:
      return check_node(node.as_unary_operand(), ref_count);
    case parser::NodeKind::BinaryOp:
      for (std::uint32_t i = 0; i < node.as_binary_arity(); ++i) {
        if (auto st = check_node(node.as_binary_operand(i), ref_count); !st) {
          return st;
        }
      }
      return {};
    case parser::NodeKind::ImplicitIntersection:
      return check_node(node.as_implicit_intersection_operand(), ref_count);
    case parser::NodeKind::Call: {
      for (const std::string_view banned : kDisallowedFunctions) {
        if (strings::case_insensitive_eq(node.as_call_name(), banned)) {
          return invalid("function not allowed in a pivot formula", node.as_call_name());
        }
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (auto st = check_node(node.as_call_arg(i), ref_count); !st) {
          return st;
        }
      }
      return {};
    }
    case parser::NodeKind::ArrayLiteral:
      return invalid("array constants are not allowed in a pivot formula");
    case parser::NodeKind::Lambda:
    case parser::NodeKind::LetBinding:
    case parser::NodeKind::LambdaCall:
      return invalid("LET and LAMBDA are not allowed in a pivot formula");
    case parser::NodeKind::Ref:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::ErrorPlaceholder:
    case parser::NodeKind::SpillRef:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::ExternalRef:
      break;
  }
  return invalid("formula must not use cell, range, structured or external references");
}

}  // namespace

Expected<std::uint32_t, Error> PivotFormulaEvaluator::compile(std::string_view formula,
                                                              const pivot::FormulaScope& scope,
                                                              std::vector<pivot::FormulaRef>* refs,
                                                              std::string* canonical) {
  assert(eval_profile_scope_active());
  if (scope.cache == nullptr || refs == nullptr) {
    return invalid("pivot formula scope has no cache");
  }
  if (scope.item_field && *scope.item_field >= scope.cache->fields().size()) {
    return invalid("field_index is out of range for the calculated item");
  }
  formula = strip_formula_prefix(formula);
  if (utf16_units_in(formula) > kMaxFormulaUnits) {
    return invalid("pivot formula longer than 8192 characters");
  }

  refs->clear();
  Rewriter rewriter(scope, refs);
  Expected<std::string, Error> text = rewriter.run(formula);
  if (!text) {
    return text.error();
  }

  Compiled compiled;
  // Placeholders can lengthen the text past the limit the input already met.
  parser::ParserOptions parse_opts;
  parse_opts.max_formula_length_utf16 = std::max(parse_opts.max_formula_length_utf16, utf16_units_in(text.value()));
  compiled.root = parser::parse_strict(text.value(), compiled.arena, parse_opts);
  if (compiled.root == nullptr) {
    return make_error(FormulonErrorCode::kParserUnexpectedToken, "formula has a syntax error", std::string(formula));
  }
  if (auto st = check_node(*compiled.root, refs->size()); !st) {
    return st.error();
  }
  compiled.ref_count = refs->size();

  while (names_.size() < refs->size()) {
    std::string name;
    append_synthetic(name, names_.size());
    names_.push_back(std::move(name));
  }
  formulas_.push_back(std::move(compiled));
  if (canonical != nullptr) {
    *canonical = std::move(rewriter.canonical());
  }
  return static_cast<std::uint32_t>(formulas_.size() - 1);
}

Value PivotFormulaEvaluator::evaluate(std::uint32_t formula_id, const Value* ref_values, std::size_t ref_count) {
  if (formula_id >= formulas_.size() || formulas_[formula_id].ref_count != ref_count) {
    return Value::error(ErrorCode::Value);
  }
  scratch_.reset();
  NameEnv env;
  for (std::size_t i = 0; i < ref_count; ++i) {
    env = env.extend(names_[i], ref_values[i], scratch_);
  }
  const EvalContext ctx = EvalContext().with_name_env(&env);
  return evaluate_first_element(*formulas_[formula_id].root, scratch_, default_registry(), ctx);
}

}  // namespace formulon::eval
