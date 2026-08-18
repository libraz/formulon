//
// Internal header -- do not include outside `src/eval/builtins/text*`.
//
// Shared helpers for the text builtin family. `read_int_arg` is referenced
// from multiple TUs (the core text family, the DBCS family, and the modern
// TEXTBEFORE/TEXTAFTER family) so its definition lives in a dedicated TU
// (`text_detail.cpp`) to avoid ODR violations. `DbcsCharRec` and
// `build_dbcs_char_map` are declared here because the register-side in
// `text.cpp` takes the address of `Lenb` / `Leftb` / ... to populate the
// `FunctionRegistry`; those implementations live in `text_dbcs.cpp` and
// `text_modern.cpp`.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_DETAIL_H_
#define FORMULON_EVAL_BUILTINS_TEXT_DETAIL_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace text_detail {

// Helper: read a numeric arg as an `int` via `std::trunc`. Returns `#VALUE!`
// on coercion failure, `#NUM!` on non-finite input. Used by LEFT/RIGHT/MID/
// REPT/FIND/SEARCH/SUBSTITUTE for their integer-typed parameters.
Expected<int, ErrorCode> read_int_arg(const Value& v);

// Reads optional integer argument `args[index]`, returning `default_value`
// when the caller omitted it. Used by LEFT/RIGHT/FIND/SEARCH and their DBCS
// variants so the defaulting rule stays in one place.
Expected<int, ErrorCode> read_optional_int_arg(const Value* args, std::uint32_t arity, std::uint32_t index,
                                               int default_value);

// Unit a search family counts positions in: UTF-16 units for FIND and
// SEARCH, ja-JP DBCS bytes for FINDB and SEARCHB.
enum class SearchUnit : std::uint8_t { Utf16, DbcsByte };

// The coerced `find_text` / `within_text` / `start_num` arguments the
// four search builtins share.
struct SearchArgs {
  std::string needle;
  std::string haystack;
  int start;
};

// Reads the argument prologue of FIND / SEARCH / FINDB / SEARCHB,
// including the two answers Excel gives before it ever looks at the
// haystack: a `start_num` outside `[1, length + 1]` is `#VALUE!`, and an
// empty `find_text` returns `start_num` itself. `unit` selects the unit
// the length bound is measured in.
//
// Returns `true` with `*out` filled when the caller should go on to
// search; `false` when the caller must return `*out_result` verbatim.
bool read_search_args(const Value* args, std::uint32_t arity, SearchUnit unit, SearchArgs* out, Value* out_result);

// Per-character record: UTF-8 byte offset, byte length, 1-based DBCS
// position (byte position under the ja-JP DBCS rule), and DBCS cost.
struct DbcsCharRec {
  std::size_t byte_offset;
  std::size_t byte_len;
  std::uint64_t dbcs_position;  // 1-based
  int dbcs_bytes;
};

// Walks `src` once and builds the per-character map. O(n) time, one pass.
std::vector<DbcsCharRec> build_dbcs_char_map(std::string_view src);

// Byte-oriented text family (ja-JP DBCS) implementations. Defined in
// `text_dbcs.cpp`. Exposed so `register_text_builtins()` in `text.cpp` can
// take their addresses when populating the FunctionRegistry.
Value Lenb(const Value* args, std::uint32_t arity, Arena& arena);
Value Leftb(const Value* args, std::uint32_t arity, Arena& arena);
Value Rightb(const Value* args, std::uint32_t arity, Arena& arena);
Value Midb(const Value* args, std::uint32_t arity, Arena& arena);
Value ReplaceB_(const Value* args, std::uint32_t arity, Arena& arena);
Value FindB_(const Value* args, std::uint32_t arity, Arena& arena);
Value SearchB_(const Value* args, std::uint32_t arity, Arena& arena);

// Modern text accessor family (TEXTBEFORE / TEXTAFTER). Defined in
// `text_modern.cpp`. Same rationale as the DBCS block above.
Value TextBefore_(const Value* args, std::uint32_t arity, Arena& arena);
Value TextAfter_(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_DETAIL_H_
