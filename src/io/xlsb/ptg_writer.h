//
// Parser AST -> MS-XLSB Ptg stream encoder. The exact inverse of
// `ptg_reader.h`: a post-order walk of the AST emits operand tokens
// first, then the operator / function token that consumes them, so the
// resulting `rgce` byte stream decodes back to a structurally
// equivalent AST.
//
// The encoder is matched byte-for-byte with the decoder in this module
// (the two are the engine's own private round-trip pair); the on-wire
// shapes follow [MS-XLSB] §2.5.97 for the common token set but the
// decoder is the authoritative consumer, so the encoder need only stay
// consistent with it.
//
// Tokens the AST can carry but the encoder cannot lower (structured refs,
// implicit-intersection, optional LAMBDA parameters)
// return `kIoXlsbUnsupportedPtg` rather
// than silently dropping data; the cell writer surfaces that as a hard
// failure through `write_xlsb`'s `Expected` return.

#ifndef FORMULON_IO_XLSB_PTG_WRITER_H_
#define FORMULON_IO_XLSB_PTG_WRITER_H_

#include <cstdint>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/xlsb/ptg.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Maps a name (a genuine defined name, or a hidden `_xlfn.*` /
/// `_xlpm.*` future-function / LET-parameter placeholder) to its
/// 1-based `BrtName` declaration index. Shared between the encoder
/// (`encode_ptgs`, which emits `PtgName(ilbl)` for any lookup hit) and
/// the top-level writer (which emits the matching `BrtName` records
/// into `xl/workbook.bin`'s globals).
using NameTable = std::unordered_map<std::string, std::uint32_t>;

/// `NameTable` key of the defined-name record scoped to 0-based sheet
/// `itab` (`-1`: workbook scope) spelling `name`, ASCII case-folded. A
/// sheet-qualified reference (`Sheet2!Local`) encodes through these keys;
/// `!` never occurs in a defined name, so they cannot collide with one.
std::string sheet_scoped_name_key(std::int32_t itab, std::string_view name);

/// One `BrtExternSheet` XTI: the book it qualifies and the span of that
/// book's sheets. `book` is 0 for this workbook, else the 1-based position
/// of an external link in `SheetRangeTable::links`; the supporting-book
/// index the file stores follows from the table (`xti_sup_book`). A
/// single-sheet qualified reference (`Sheet1!A1`) stores `first == last`, a
/// genuine 3-D range (`Sheet1:Sheet3!A1`) the full span, and a book-scope
/// `PtgNameX` `kXtiNoSheet` for both.
struct XtiEntry {
  std::uint32_t book = 0;
  std::int32_t first = 0;
  std::int32_t last = 0;

  bool operator==(const XtiEntry& other) const noexcept {
    return book == other.book && first == other.first && last == other.last;
  }
};

/// What a formula can name inside one saved external link.
struct XlsbLinkTables {
  /// `ExternalLinkRecord::index` of the link.
  std::uint32_t index = 0;
  /// The supporting workbook's sheets; an XTI's span indexes these.
  std::vector<std::string> sheet_names;
  /// The defined names its part lists: the cached ones in cache order, then
  /// any a formula names that the cache lacks. `PtgNameX` names entry `i`
  /// as `ilbl == i + 1`.
  std::vector<std::string> names;
};

/// Mirrors `xl/workbook.bin`'s `BrtExternSheet` table: `xti[i]` is the
/// entry a `PtgRef3d` / `PtgArea3d` / `PtgNameX` token resolves via
/// `ixti == i`. Built once per workbook by `collect_ptg_sheet_ranges`
/// (mirroring `NameTable` / `collect_ptg_names`) so every sheet's cell
/// encoder and the `BrtExternSheet` record the top-level writer emits agree
/// on `ixti` assignments.
struct SheetRangeTable {
  std::vector<XtiEntry> xti;
  /// The external links the package saves, in the order their parts and
  /// `BrtSupBookSrc` records are written.
  std::vector<XlsbLinkTables> links;
  /// Maps a cross-workbook reference's directory and book to the
  /// `ExternalLinkRecord::index` it names; unset when there are no links.
  parser::ExternalBookIndexer indexer{nullptr, nullptr};
};

/// The `iSupBook` the file stores for an XTI naming `book`: this workbook's
/// `BrtSupSelf` comes first when any XTI names it, then one `BrtSupBookSrc`
/// per link. Index 0 is therefore an external link in a table without a
/// self entry.
std::uint32_t xti_sup_book(const SheetRangeTable& table, std::uint32_t book);

/// True when an XTI names this workbook, which then needs `BrtSupSelf`.
bool xti_names_self(const SheetRangeTable& table);

/// 1-based position in `table.links` of the link the cross-workbook
/// reference `node` names; 0 when no saved link matches.
std::uint32_t external_link_position(const SheetRangeTable& table, const parser::AstNode& node);

/// 0-based index of `sheet` among `link.sheet_names` under ASCII case
/// folding, as Excel matches sheet names; -1 when absent.
int external_sheet_index(const XlsbLinkTables& link, std::string_view sheet);

/// `ilbl` of `name` in `link.names` under ASCII case folding; 0 when absent.
std::uint32_t external_name_ilbl(const XlsbLinkTables& link, std::string_view name);

/// Ptg class of a formula's own root token (the whole formula, or a bare
/// reference an enclosing function/operator does not consume). Measured
/// against real Excel 365 output: a cell formula's root promotes to value
/// class (`kValue`); a `BrtName` defined-name body's root stays reference
/// class (`kReference`). See `encode_ptgs`'s doc comment.
enum class PtgRootClass : std::uint8_t {
  kValue,
  kReference,
};

/// How the formula evaluates arrays, which decides an operator's operand
/// class inside a function (measured): a dynamic-array formula gives an area
/// operand array class where the parameter takes arrays
/// (`SUM(A1:A2*2)` -> 0x65); a legacy formula (`Cell::dynamic_array` clear)
/// intersects it, value class (0x45), except under a forced-array parameter
/// such as SUMPRODUCT's. A legacy formula also stores no `@` where Excel's
/// own implied one stands (`legacy_intersections`); a legacy CSE block
/// (`kLegacyArray`) intersects nothing, so every written `@` is kept. A
/// conditional-format formula (`kConditionalFormat`) gives any operand under a
/// parameter other than a value one array class, a cell or a name included
/// (`AND(H2>Lim)` -> 0x6C / 0x63).
enum class PtgEvaluation : std::uint8_t {
  kDynamicArray,
  kLegacy,
  kLegacyArray,
  kConditionalFormat,
};

/// Result of `encode_ptgs`: the main token stream plus the array-
/// constant extra-data area a `CellParsedFormula` appends after it.
struct EncodedFormula {
  /// The `rgce` Ptg token stream.
  std::vector<std::uint8_t> rgce;
  /// The `rgcb` extra-data area `PtgArray` tokens in `rgce` reference
  /// (rows/cols + tagged elements, consumed by the decoder in
  /// encounter order). Empty when the formula carries no array
  /// constants. The caller emits this as the `CellParsedFormula`'s
  /// `cb` (byte length) + `rgcb` (payload) trailer, after `cce` + `rgce`.
  std::vector<std::uint8_t> rgcb;
};

/// Walks `node`'s AST appending, in encounter order, every name a
/// `PtgName` reference will be needed for while encoding it:
///
///   * A `Call` node naming one of the enumerated future functions
///     (`io/future_functions.h`: XLOOKUP, TEXTJOIN, CONCAT, IFS,
///     SEQUENCE, ...), which Excel stores as a hidden `_xlfn.*` name
///     rather than a function id.
///   * An unqualified `NameRef` (an ordinary defined-name reference, e.g.
///     `Rate`) or a self-book `[0]!Rate`. A sheet-qualified one resolves
///     through a scoped key instead (`collect_sheet_qualified_names`).
///
///   * A callee with no function id (`Fn(3)`): a named LAMBDA, or a name no
///     one defined, which Excel stores against an empty BrtName stub. Every
///     built-in has an id or a hidden-name route, so this is never one.
///
///   * The hidden `_xlfn.LET` / `_xlfn.LAMBDA` callee and one
///     `_xlpm.<param>` placeholder per LET binding or LAMBDA parameter.
///
/// Names already present in `seen` are skipped (both to dedupe and so
/// callers can pre-seed `seen` with names that already have an assigned
/// `ilbl`, e.g. from the workbook's existing defined-name table).
void collect_ptg_names(const parser::AstNode& node, std::vector<std::string>& names,
                       std::unordered_set<std::string>& seen);

/// `collect_ptg_names` without the self-book `[0]!Rate`: every name the
/// formula resolves from its own scope.
void collect_scope_resolved_names(const parser::AstNode& node, std::vector<std::string>& names,
                                  std::unordered_set<std::string>& seen);

/// Appends every distinct sheet-qualified defined-name reference in `node`
/// (`Sheet2!Rate`, `Sheet2!Fn(3)`) as `(sheet, name)`, in encounter order.
/// Excel stores an undefined one against an empty `BrtName` stub scoped to
/// that sheet, which the writer has to emit.
void collect_sheet_qualified_names(const parser::AstNode& node,
                                   std::vector<std::pair<std::string, std::string>>& qualified);

/// Walks `node`'s AST appending, in encounter order, every distinct
/// `(itabFirst, itabLast)` sheet-range pair a qualified reference will
/// need an `ixti` for while encoding it: a single-sheet `Ref` whose
/// `sheet` is non-empty contributes `(itab, itab)`; a `Ref3D` node
/// contributes its full `(begin, end)` span; a sheet-qualified defined
/// or self-book name (`Sheet2!Rate`, `[0]!Rate`) contributes the sheetless
/// `(-2, -2)` entry its `PtgNameX` resolves through. A cross-workbook
/// reference contributes the same shapes against its link's sheets, and a
/// name the link's cache lacks is appended to that link's `names`.
/// `sheet_names` resolves a sheet display name to its 0-based index; a name
/// absent from `sheet_names` is skipped here (the encode fails later with a
/// precise error instead of silently fabricating an entry). `seen` dedupes
/// (both across one call and across callers pre-seeding it), and
/// `ranges.xti`'s index order becomes the `ixti` assignment `encode_ptgs`
/// consults via `sheet_ranges`.
void collect_ptg_sheet_ranges(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                              SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen);

/// What a formula can know of a defined name from the name's own formula,
/// without evaluating it.
struct NameShape {
  /// The name stands for one value: a one-cell reference or a constant.
  bool scalar = false;
  /// Rows and columns of the one plain rectangle it refers to; 0 otherwise.
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  /// The integer its formula is, when that is a literal Excel stores as
  /// `PtgInt` (1..65535); 0 otherwise.
  std::uint32_t positive_int = 0;
  /// Excel stores the name with `fCalcExp` (`name_sets_calc_exp`).
  bool calc_exp = false;
  /// The name is defined.
  bool defined = false;
  /// Its formula is recalculated every time (`formula_always_calculates`).
  bool always_calculates = false;
};

/// The `NameShape` of the defined name a `NameRef`, `[0]!Name` or name-call
/// node refers to. An empty one takes every name as an unknown, multi-cell
/// reference.
using NameShapes = std::function<NameShape(const parser::AstNode& name)>;

/// True when Excel 365 enters the cell formula `root` as a dynamic-array
/// formula (`cm` / `BrtCellMeta`): its result may be more than one value, or
/// it evaluates a multi-valued operand where a formula without the mark
/// would intersect it (`=SUM(A1:A2*2)`). Measured against Excel's own marks.
bool formula_is_dynamic_array(const parser::AstNode& root, const NameShapes& names);

/// True when Excel 365 sets `fCalcExp` on a defined name whose formula is
/// `root` (measured name by name): its value is a call to a built-in
/// `xlsb_sets_calc_exp` lists or through a name, a LET, a LAMBDA or a call
/// to one, or a spill
/// reference, reached from the root through operators but not parentheses;
/// or it refers anywhere to a defined name `names` says carries it.
bool name_sets_calc_exp(const parser::AstNode& root, const NameShapes& names);

/// True when a formula calls one of Excel's volatile functions
/// (`parser::is_volatile_function_name`); XLSB opens it with `PtgAttrSemi`.
bool formula_calls_volatile(const parser::AstNode& root);

/// True when Excel 365 recalculates the formula `root` every time, and so
/// stores it with `ca="1"` in .xlsx (measured): it calls a volatile function,
/// one that is neither built in nor defined, or a reference (`A1(1)`), or
/// refers to a defined name whose formula `names` says is recalculated every
/// time.
bool formula_always_calculates(const parser::AstNode& root, const NameShapes& names);

/// The nodes of `root`, a cell formula without the dynamic-array mark, that
/// Excel 365 shows behind an implicit-intersection `@` (its formula2 text):
/// an area, a multi-cell name or an array-returning call where an area would
/// be value class. A written `@` node is listed itself when it stands where
/// Excel's own would, so a legacy writer can drop it; in document order.
std::vector<const parser::AstNode*> legacy_intersections(const parser::AstNode& root, const NameShapes& names);

/// Encodes the AST rooted at `node` into an `rgce` Ptg byte stream (plus
/// its `rgcb` array-constant extra data, see `EncodedFormula`).
/// `sheet_names` maps a sheet display name to its 0-based index.
/// `sheet_ranges` maps a resolved `(itabFirst, itabLast)` pair to its
/// `ixti` (its index in the table, built by `collect_ptg_sheet_ranges`);
/// a qualified reference whose sheet(s) are not present in
/// `sheet_ranges` fails the encode rather than fabricating an entry. (In
/// practice the cell writer passes a table built from every formula in
/// the workbook, so any sheet referenced by a live formula resolves.)
/// `name_table` resolves `NameRef` nodes and future-function `Call`
/// callees to a `PtgName` `ilbl` (see `collect_ptg_names`); a name
/// absent from the table fails the encode. `root_class` selects whether
/// a bare reference/range that is `node` itself (not consumed by any
/// enclosing function or operator) promotes to value class -- see
/// `PtgRootClass`.
///
/// With a `base` cell (a conditional-format or data-validation formula,
/// see `PtgBaseCell`), a same-sheet cell or area reference with any
/// relative axis becomes `PtgRefN` / `PtgAreaN`: a relative axis stores its
/// offset from `base` modulo the grid, an absolute one the index itself. A
/// fully absolute reference, and every reference without a base, keeps
/// `PtgRef` / `PtgArea`.
///
/// Returns `kIoXlsbUnsupportedPtg` for any node kind outside the
/// supported set (see header banner). The error context names the
/// offending node kind. `names` feeds the shapes that decide a call's class
/// and a legacy formula's `legacy_intersections`.
Expected<EncodedFormula, Error> encode_ptgs(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                                            const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                            PtgRootClass root_class, std::optional<PtgBaseCell> base = std::nullopt,
                                            PtgEvaluation evaluation = PtgEvaluation::kDynamicArray,
                                            const NameShapes& names = NameShapes());

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_WRITER_H_
