
#include "eval/textjoin_lazy.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "eval/text_ops.h"
#include "eval/utf8_length.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Excel caps the result of TEXTJOIN (and REPT / SUBSTITUTE and other text
// builtins) at 32,767 UTF-16 units.
constexpr std::uint64_t kExcelTextCapUnits = 32767u;

// Longest UTF-8 sequence, and therefore the most trailing bytes of a buffer
// whose decoding a later append can still change.
constexpr std::size_t kMaxUtf8SequenceBytes = 4u;

// Running UTF-16 unit count for a buffer that is appended to repeatedly.
// Moved here verbatim from the retired eager `TextJoin` impl in
// `eval/builtins/text.cpp` — this is its only remaining caller.
//
// `utf16_units_in` walks the whole buffer, so calling it after every append
// costs O(n^2) in the accumulated length. Summing the pieces' own counts
// instead would be O(1) but not the same number: a truncated sequence at
// the end of one piece counts as one malformed unit per byte until the
// continuation bytes arrive, at which point the same bytes decode as a
// single codepoint. Since the cap decides whether TEXTJOIN returns text or
// `#VALUE!`, that difference is Excel-observable.
//
// So the count is committed only for the prefix a later append can no
// longer affect — everything up to the last `kMaxUtf8SequenceBytes` — and
// the short tail is re-decoded each round. `units_of` returns exactly what
// `utf16_units_in` would, at O(1) amortised cost per append.
class RunningUtf16Units {
 public:
  std::uint64_t units_of(std::string_view buffer) {
    while (committed_bytes_ + kMaxUtf8SequenceBytes <= buffer.size()) {
      std::size_t step = 0;
      const std::uint32_t cp = decode_utf8_step(buffer, committed_bytes_, &step);
      if (step == 0) {
        break;  // Defensive: the decoder only reports 0 past the end.
      }
      committed_units_ += cp > 0xFFFFu ? 2u : 1u;
      committed_bytes_ += step;
    }
    return committed_units_ + utf16_units_in(buffer.substr(committed_bytes_));
  }

 private:
  std::size_t committed_bytes_ = 0;
  std::uint64_t committed_units_ = 0;
};

// Flattens `node` (scalar, range, or array) into `out`, appending each
// cell's text projection in row-major order. Mirrors the range-flattening
// the generic `accepts_ranges` dispatcher performs for an ordinary
// aggregator argument; TEXTJOIN's own `text1, [text2], ...` positions need
// the same treatment, just resolved directly from the AST instead of
// through the dispatcher's per-position loop (which cannot special-case
// `delimiter` — see the header comment).
bool flatten_text_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::vector<std::string>* out, Value* out_err) {
  const Value v = eval_node_as_array(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  const ArrayValue* arr = v.as_array();
  const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
  out->reserve(out->size() + n);
  for (std::size_t i = 0; i < n; ++i) {
    auto t = coerce_to_text(arr->cells[i]);
    if (!t) {
      *out_err = Value::error(t.error());
      return false;
    }
    out->push_back(std::move(t.value()));
  }
  return true;
}

}  // namespace

Value eval_textjoin_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U) {
    return Value::error(ErrorCode::Value);
  }

  // delimiter (arg 0): scalar or range/array text. `eval_node_as_array`
  // always wraps a scalar into a 1x1 array, so `delimiters` is never
  // empty. An array delimiter is applied cyclically below, over the
  // flattened text pieces.
  const Value delim_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (delim_v.is_error()) {
    return delim_v;
  }
  const ArrayValue* delim_arr = delim_v.as_array();
  const std::size_t delim_n = static_cast<std::size_t>(delim_arr->rows) * static_cast<std::size_t>(delim_arr->cols);
  std::vector<std::string> delimiters;
  delimiters.reserve(delim_n);
  for (std::size_t i = 0; i < delim_n; ++i) {
    auto t = coerce_to_text(delim_arr->cells[i]);
    if (!t) {
      return Value::error(t.error());
    }
    delimiters.push_back(std::move(t.value()));
  }

  // ignore_empty (arg 1): scalar only.
  const Value ignore_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (ignore_v.is_error()) {
    return ignore_v;
  }
  auto ignore_c = coerce_to_bool(ignore_v);
  if (!ignore_c) {
    return Value::error(ignore_c.error());
  }
  const bool ignore_empty = ignore_c.value();

  // text1, [text2], ... (args 2..N): each may be scalar, range, or array;
  // flatten every argument's cells, in call order, into one text-piece
  // list.
  std::vector<std::string> pieces;
  for (std::uint32_t i = 2; i < arity; ++i) {
    Value err = Value::blank();
    if (!flatten_text_arg(call.as_call_arg(i), arena, registry, ctx, &pieces, &err)) {
      return err;
    }
  }

  std::string out;
  bool first = true;
  RunningUtf16Units counter;
  std::size_t delim_index = 0;
  for (const std::string& piece : pieces) {
    if (ignore_empty && piece.empty()) {
      continue;
    }
    if (!first) {
      out.append(delimiters[delim_index % delimiters.size()]);
      ++delim_index;
    }
    out.append(piece);
    first = false;
    // Cap check after each appended piece, on the joined result rather than
    // on the piece: the cap is a property of the whole string, and the
    // first piece that carries it over is the one that fails the call.
    if (counter.units_of(out) > kExcelTextCapUnits) {
      return Value::error(ErrorCode::Value);
    }
  }
  return Value::text(arena.intern(out));
}

}  // namespace eval
}  // namespace formulon
