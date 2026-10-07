//
// Implementation of the `xl/workbook.bin` stream writer. See
// `io/xlsb/workbook_bin_writer.h`.

#include "io/xlsb/workbook_bin_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "external_book.h"
#include "external_link.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/xlsb/external_link_writer.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Workbook-globals record ids the reader does not consume (so they are absent
// from `XlsbRecordType`) but the writer must emit for a well-formed stream.
// Workbook-globals structural records Excel expects before the sheet bundle.
constexpr std::uint16_t kBrtFileVersion = 128;
constexpr std::uint16_t kBrtBeginBookViews = 135;
constexpr std::uint16_t kBrtWbView = 158;
constexpr std::uint16_t kBrtEndBookViews = 136;
// Opens a future-record block in a sheet tail (x14 extensions).
constexpr std::uint16_t kBrtFrtBegin = 35;

}  // namespace

// ---------------------------------------------------------------------------
// PtgName table (BrtName): defined names + future-function callees
// ---------------------------------------------------------------------------

namespace {

/// Parses `formula` (with or without a leading `=`) and folds every name
/// `collect_ptg_names` finds into `names` / `seen`. Parse failures are
/// silently skipped here — `EncodeCellFormula` / the defined-name
/// encode pass below surface the same failure as a proper `Expected`
/// error when the formula is actually encoded.
///
/// `scope_sheet_id` is the scope the formula resolves names from (see
/// `BuildNameTableForScope`). Every referenced text defined only as the
/// local name of sheets other than that scope is added to `invisible`:
/// Excel resolves such a reference to `#NAME?`, so it needs a
/// workbook-scope record of its own rather than an ordinal borrowed from
/// another sheet's definition, which would read back as that sheet's
/// qualified name. Every sheet-qualified reference is appended to
/// `qualified` as `(sheet, name)`.
void CollectNamesFromFormula(std::string_view formula, std::int32_t scope_sheet_id,
                             const std::unordered_map<std::string, std::vector<std::int32_t>>& defined_scopes,
                             std::vector<std::string>& names, std::unordered_set<std::string>& seen,
                             std::vector<std::string>& invisible, std::unordered_set<std::string>& invisible_seen,
                             std::vector<std::pair<std::string, std::string>>& qualified) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr || !p.errors().empty()) {
    return;
  }
  collect_ptg_names(*root, names, seen);
  collect_sheet_qualified_names(*root, qualified);
  std::vector<std::string> referenced;
  std::unordered_set<std::string> referenced_seen;
  collect_scope_resolved_names(*root, referenced, referenced_seen);
  for (std::string& text : referenced) {
    const auto it = defined_scopes.find(text);
    if (it == defined_scopes.end()) {
      continue;
    }
    const std::vector<std::int32_t>& scopes = it->second;
    const bool visible = std::any_of(scopes.begin(), scopes.end(),
                                     [scope_sheet_id](std::int32_t s) { return s < 0 || s == scope_sheet_id; });
    if (!visible && invisible_seen.insert(text).second) {
      invisible.push_back(std::move(text));
    }
  }
}

/// Calls `visit` with the text (no leading `=`) of every formula a sheet
/// part carries: cell formulas in row-major order, as Excel numbers their
/// entries, then conditional-format rule and threshold formulas, and
/// data-validation formulas. The name and ExternSheet tables are both built
/// over this one walk, so every formula the sheet writer encodes finds its
/// entries.
template <typename Visit>
void ForEachSheetFormula(const Sheet& sheet, Visit&& visit) {
  std::vector<std::uint32_t> rows;
  rows.reserve(sheet.rows().size());
  for (const auto& kv : sheet.rows()) {
    rows.push_back(kv.first);
  }
  sort_ascending(rows);
  for (const std::uint32_t row : rows) {
    for (const Cell& cell : sheet.rows().at(row)) {
      std::string_view body(cell.formula_text);
      if (!body.empty() && body.front() == '=') {
        body.remove_prefix(1);
      }
      if (!body.empty()) {
        visit(body);
      }
    }
  }
  const auto visit_text = [&visit](const std::string& text) {
    if (!text.empty()) {
      visit(std::string_view(text));
    }
  };
  const auto visit_cfvo = [&visit_text](const cf::CfValueObject& v) {
    if (v.type == cf::CfvoType::Formula) {
      visit_text(v.value);
    }
  };
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    for (const cf::CFRule& rule : format.rules) {
      visit_text(rule.formula1.value_or(std::string()));
      visit_text(rule.formula2.value_or(std::string()));
      if (rule.color_scale) {
        std::for_each(rule.color_scale->thresholds.begin(), rule.color_scale->thresholds.end(), visit_cfvo);
      }
      if (rule.data_bar) {
        visit_cfvo(rule.data_bar->min);
        visit_cfvo(rule.data_bar->max);
      }
      if (rule.icon_set) {
        std::for_each(rule.icon_set->thresholds.begin(), rule.icon_set->thresholds.end(), visit_cfvo);
      }
    }
  }
  for (const DataValidation& dv : sheet.validations()) {
    visit_text(dv.formula1);
    visit_text(dv.formula2);
  }
}

}  // namespace

/// Builds the workbook's `BrtName` record order: every genuine defined
/// name (`Workbook::defined_names()`, in declaration order) occupies the
/// leading slots, followed by every future-function callee / `NameRef`
/// `collect_ptg_names` discovers across defined-name and sheet formulas
/// (in first-encounter order) that is not already a defined name, then a
/// workbook-scope placeholder for each name referenced where none of its
/// definitions is visible, then an empty stub scoped to `S` for each
/// `S!Name` that neither `S` nor the workbook defines, as Excel 365 saves
/// it. Slot `i` is the record a `PtgName` reaches with `ilbl == i + 1`.
///
/// Defined names take one slot each, duplicate name text included:
/// `Workbook::set_defined_name_scoped` deliberately admits a
/// workbook-scoped and a sheet-local name spelling the same text, and
/// each needs its own `BrtName` record for the emission pass (which
/// walks `wb.defined_names()` by the same index) to stay aligned with
/// the `ilbl` values encoded into formulas.
void BuildOrderedNames(const Workbook& wb, std::vector<OrderedName>& ordered_names) {
  std::vector<std::string> names;
  std::unordered_set<std::string> seen;
  std::vector<std::string> invisible;
  std::unordered_set<std::string> invisible_seen;
  std::vector<std::pair<std::string, std::string>> qualified;
  std::unordered_map<std::string, std::vector<std::int32_t>> defined_scopes;
  std::unordered_set<std::string> scoped_keys;
  for (const DefinedName& dn : wb.defined_names()) {
    names.push_back(dn.name);
    seen.insert(dn.name);
    defined_scopes[dn.name].push_back(dn.local_sheet_id);
    scoped_keys.insert(sheet_scoped_name_key(dn.local_sheet_id, dn.name));
  }
  for (const DefinedName& dn : wb.defined_names()) {
    CollectNamesFromFormula(dn.formula, dn.local_sheet_id, defined_scopes, names, seen, invisible, invisible_seen,
                            qualified);
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    ForEachSheetFormula(wb.sheet(i), [&](std::string_view formula) {
      CollectNamesFromFormula(formula, static_cast<std::int32_t>(i), defined_scopes, names, seen, invisible,
                              invisible_seen, qualified);
    });
  }
  for (std::string& text : names) {
    ordered_names.push_back(OrderedName{std::move(text), -1});
  }
  for (std::string& text : invisible) {
    ordered_names.push_back(OrderedName{std::move(text), -1});
  }
  for (auto& [sheet, name] : qualified) {
    const std::size_t itab = wb.sheet_index_by_name(sheet);
    if (itab >= wb.sheet_count()) {
      continue;  // The encoder reports the unknown sheet.
    }
    const auto scope = static_cast<std::int32_t>(itab);
    if (scoped_keys.count(sheet_scoped_name_key(-1, name)) == 0 &&
        scoped_keys.insert(sheet_scoped_name_key(scope, name)).second) {
      ordered_names.push_back(OrderedName{std::move(name), scope});
    }
  }
}

/// Builds the `name -> ilbl` map `encode_ptgs` consults for a formula
/// that lives in `scope_sheet_id` (a 0-based sheet index for a cell
/// formula or a sheet-local defined name, `-1` for a workbook-scoped
/// defined name).
///
/// A `PtgName` token carries no scope of its own — it is a bare ordinal
/// into the `BrtName` table — so the scope has to be resolved here, at
/// encode time, exactly as Excel resolves it when reading the file back:
/// a sheet-local name shadows a workbook-scoped one of the same text for
/// formulas on that sheet, and the workbook-scoped name is reached only
/// where no local one exists. Encoding a single workbook-wide ordinal
/// per name text would silently re-point one of the two.
///
/// The three passes are layered by precedence and rely on `emplace`
/// leaving an entry that is already present alone:
///   1. names local to `scope_sheet_id`, stubs included — these shadow
///      everything;
///   2. workbook-scoped names, filling any text no local one claimed;
///   3. every placeholder slot, then every remaining defined slot. The
///      placeholders are the hidden `_xlfn.*` / `_xlpm.*` names (which
///      have no scope), undefined names, and the workbook-scope record
///      `BuildOrderedNames` adds for a text that exists only as another
///      sheet's local name: that reference is `#NAME?` in Excel, and
///      reaching the other sheet's record instead would read back as
///      that sheet's qualified name.
NameTable BuildNameTableForScope(const Workbook& wb, const std::vector<OrderedName>& ordered_names,
                                 std::int32_t scope_sheet_id) {
  const std::vector<DefinedName>& defined = wb.defined_names();
  const std::size_t defined_count = std::min(defined.size(), ordered_names.size());
  NameTable name_table;
  name_table.reserve(ordered_names.size());
  if (scope_sheet_id >= 0) {
    for (std::size_t i = 0; i < defined_count; ++i) {
      if (defined[i].local_sheet_id == scope_sheet_id) {
        name_table.emplace(defined[i].name, static_cast<std::uint32_t>(i + 1));
      }
    }
    // A stub scoped to this sheet is its local name too: Excel resolves a
    // bare `NOSUCH(1)` on Sheet1 to the stub `Sheet1!NOSUCH` created there.
    for (std::size_t i = defined_count; i < ordered_names.size(); ++i) {
      if (ordered_names[i].itab == scope_sheet_id) {
        name_table.emplace(ordered_names[i].name, static_cast<std::uint32_t>(i + 1));
      }
    }
  }
  for (std::size_t i = 0; i < defined_count; ++i) {
    if (defined[i].local_sheet_id < 0) {
      name_table.emplace(defined[i].name, static_cast<std::uint32_t>(i + 1));
    }
  }
  for (std::size_t i = defined_count; i < ordered_names.size(); ++i) {
    const OrderedName& slot = ordered_names[i];
    // A sheet-scoped stub is reachable only through its qualified key.
    name_table.emplace(slot.itab < 0 ? slot.name : sheet_scoped_name_key(slot.itab, slot.name),
                       static_cast<std::uint32_t>(i + 1));
  }
  for (std::size_t i = 0; i < defined_count; ++i) {
    name_table.emplace(ordered_names[i].name, static_cast<std::uint32_t>(i + 1));
  }
  // Scope-independent keys for sheet-qualified references (`Sheet2!Local`).
  for (std::size_t i = 0; i < defined_count; ++i) {
    name_table.emplace(sheet_scoped_name_key(defined[i].local_sheet_id, defined[i].name),
                       static_cast<std::uint32_t>(i + 1));
  }
  return name_table;
}

// ---------------------------------------------------------------------------
// ExternSheet table (BrtExternSheet): every sheet-qualified reference's
// (itabFirst, itabLast) span, single- and multi-sheet alike.
// ---------------------------------------------------------------------------

namespace {

/// Parses `formula` (no leading `=`) and folds every distinct
/// `(itabFirst, itabLast)` span `collect_ptg_sheet_ranges` finds into
/// `ranges` / `seen`. Parse failures are silently skipped here --
/// `EncodeCellFormula` surfaces the same failure as a proper `Expected`
/// error when the formula is actually encoded.
void CollectSheetRangesFromFormula(std::string_view formula, const std::vector<std::string>& sheet_names,
                                   SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen) {
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr || !p.errors().empty()) {
    return;
  }
  collect_ptg_sheet_ranges(*root, sheet_names, ranges, seen);
}

/// True when `records` holds a future-record block (BrtFRTBegin). Such
/// blocks -- x14 conditional formats and validations -- are the retained
/// content that carries sheet-qualified formulas.
bool HoldsFrtBlock(const std::vector<std::uint8_t>& records) {
  ByteSpan cursor{records.data(), records.size()};
  while (cursor.size != 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      return false;
    }
    if (rec.value().type == kBrtFrtBegin) {
      return true;
    }
  }
  return false;
}

/// The XTI that names, among the saved `links`, what a retained entry of
/// another workbook named in the source; none when no link carries it.
std::optional<XtiEntry> ExternalSeedEntry(const XlsbExternSheetEntry& entry, const std::vector<XlsbLinkTables>& links) {
  for (std::size_t k = 0; k < links.size(); ++k) {
    if (links[k].index != entry.external_book) {
      continue;
    }
    const auto position = static_cast<std::uint32_t>(k + 1U);
    if (entry.first.empty() && entry.last.empty()) {
      return XtiEntry{position, kXtiNoSheet, kXtiNoSheet};
    }
    const int first = external_sheet_index(links[k], entry.first);
    const int last = external_sheet_index(links[k], entry.last);
    if (first < 0 || last < 0) {
      return std::nullopt;
    }
    return XtiEntry{position, first, last};
  }
  return std::nullopt;
}

/// The leading entries of the `BrtExternSheet` table: the source
/// workbook's table, in its `ixti` order, when a retained sheet tail holds
/// a block that may reference it, so those verbatim indices keep naming
/// the same sheets. Another workbook's entry is re-pointed at its link's
/// position in `links`. Fails when an entry no longer names a sheet it can
/// be bound to (renamed or removed since load, or an unbound link), since
/// the retained bytes would then reference the wrong sheet.
Expected<SheetRangeTable, Error> SeedFromRetainedTails(const Workbook& wb, const std::vector<XlsbLinkTables>& links) {
  SheetRangeTable seed;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const XlsbSheetTail& tail = wb.sheet(i).xlsb_tail();
    if (tail.extern_sheets.empty() ||
        !(HoldsFrtBlock(tail.before_merges) || HoldsFrtBlock(tail.after_merges_before_hyperlinks) ||
          HoldsFrtBlock(tail.after_hyperlinks))) {
      continue;
    }
    for (const XlsbExternSheetEntry& entry : tail.extern_sheets) {
      if (entry.external_book != 0U) {
        const std::optional<XtiEntry> xti = ExternalSeedEntry(entry, links);
        if (!xti) {
          return make_error(FormulonErrorCode::kIoXlsbRetainedPartStale,
                            "retained XLSB sheet records reference an external sheet no link carries",
                            "context=write_xlsb sheet=" + wb.sheet(i).name() + " ref=" + entry.first);
        }
        seed.xti.push_back(*xti);
        continue;
      }
      if (entry.first.empty() && entry.last.empty() && !entry.unresolved) {
        seed.xti.push_back(XtiEntry{0U, kXtiNoSheet, kXtiNoSheet});
        continue;
      }
      const std::size_t first = wb.sheet_index_by_name(entry.first);
      const std::size_t last = wb.sheet_index_by_name(entry.last);
      if (entry.unresolved || first >= wb.sheet_count() || last >= wb.sheet_count()) {
        return make_error(FormulonErrorCode::kIoXlsbRetainedPartStale,
                          "retained XLSB sheet records reference a sheet this workbook no longer has",
                          "context=write_xlsb sheet=" + wb.sheet(i).name() + " ref=" + entry.first);
      }
      seed.xti.push_back(XtiEntry{0U, static_cast<std::int32_t>(first), static_cast<std::int32_t>(last)});
    }
    break;
  }
  return seed;
}

}  // namespace

/// Builds the `BrtExternSheet` table for the whole workbook: every
/// distinct sheet-qualified reference span (single-sheet `(itab, itab)`
/// or a genuine 3-D range `(itabFirst, itabLast)`) any sheet formula (see
/// `ForEachSheetFormula`) or defined-name formula needs, in first-encounter order. Every
/// `PtgRef3d` / `PtgArea3d` token this writer emits -- single- or
/// multi-sheet alike -- resolves its `ixti` through this one table:
/// once the workbook emits any `BrtExternSheet` entry, the reader
/// interprets *every* `ixti` as an index into it rather than a bare
/// sheet index (see `ptg_reader.cpp`'s `sheet_for_ixti` /
/// `sheet_range_for_ixti`), so single- and multi-sheet references
/// cannot use two different numbering schemes in the same file.
Expected<SheetRangeTable, Error> BuildSheetRangeTable(const Workbook& wb, const std::vector<std::string>& sheet_names) {
  SheetRangeTable collected;
  for (const ExternalLinkRecord* link : written_external_links(wb)) {
    XlsbLinkTables tables;
    tables.index = link->index;
    tables.sheet_names = link->book.sheet_names;
    for (const ExternalBookName& name : link->book.names) {
      tables.names.push_back(name.name);
    }
    collected.links.push_back(std::move(tables));
  }
  if (!collected.links.empty()) {
    collected.indexer = wb.external_book_indexer();
  }
  std::unordered_set<std::uint64_t> seen;
  for (const DefinedName& dn : wb.defined_names()) {
    CollectSheetRangesFromFormula(dn.formula, sheet_names, collected, seen);
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    ForEachSheetFormula(wb.sheet(i), [&](std::string_view formula) {
      CollectSheetRangesFromFormula(formula, sheet_names, collected, seen);
    });
  }
  auto ranges = SeedFromRetainedTails(wb, collected.links);
  if (!ranges) {
    return ranges.error();
  }
  std::vector<XtiEntry>& xti = ranges.value().xti;
  for (const XtiEntry& entry : collected.xti) {
    if (std::find(xti.begin(), xti.end(), entry) == xti.end()) {
      xti.push_back(entry);
    }
  }
  ranges.value().links = std::move(collected.links);
  ranges.value().indexer = collected.indexer;
  return ranges;
}

namespace {

/// Emits one `BrtExternSheet` record. Byte layout verified against a
/// real Excel-365-produced `xl/workbook.bin` (see `workbook_bin_reader.cpp`'s
/// `DecodeExternSheet`, the decoder counterpart): `count(u32)` followed
/// by `count` entries of `(iSupBook, itabFirst, itabLast)` as three
/// `i32`s each, `iSupBook` indexing the supporting books `BuildWorkbookBin`
/// lists before it.
void EmitExternSheet(std::vector<std::uint8_t>& body, const SheetRangeTable& ranges) {
  std::vector<std::uint8_t> p;
  emit_u32(p, static_cast<std::uint32_t>(ranges.xti.size()));
  for (const XtiEntry& entry : ranges.xti) {
    emit_u32(p, xti_sup_book(ranges, entry.book));
    emit_u32(p, static_cast<std::uint32_t>(entry.first));
    emit_u32(p, static_cast<std::uint32_t>(entry.last));
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet), p);
}

/// Emits one `BrtName` record. Byte layout verified against a real
/// Excel-365-produced `xl/workbook.bin` (see `workbook_bin_reader.cpp`'s
/// `DecodeWorkbookNames`, the decoder counterpart):
///   flags (u16: bit 0 = fHidden) + 3 reserved bytes + itab (i32,
///   `-1` = workbook scope) + cch (u32) + cch x UTF-16LE name +
///   cce (u32) + cce bytes rgce + cb (u32) + cb bytes rgcb.
/// `formula` is the name's own bare expression text (no leading `=`,
/// matching `<definedName>`'s XML text content); an empty string emits
/// a zero-length `rgce` (the hidden future-function / LET-parameter
/// placeholders this writer registers carry no real formula body — real
/// Excel stores a `#NAME?` placeholder there, which is not required for
/// this writer's own reader to round-trip the name table).
Expected<void, Error> EmitName(std::vector<std::uint8_t>& body, const std::string& name, std::string_view formula,
                               std::int32_t itab, bool hidden, bool calc_exp, std::string_view comment,
                               const std::vector<std::string>& sheet_names, const SheetRangeTable& sheet_ranges,
                               const NameTable& name_table) {
  // Names carrying Excel's hidden storage prefixes are not ordinary
  // defined names: `_xlfn.<FN>` registers a post-2007 "future function"
  // and `_xlpm.<param>` a LET / LAMBDA parameter. Real Excel stores each
  // with a specific flag word (fHidden | fFunc | fFutureFunction for the
  // former, the proc-parameter flags for the latter), a PtgErr(#NAME?)
  // placeholder body, and five trailing null strings. A cell's future-
  // function call resolves through the matching BrtName's ilbl, so these
  // flags and the placeholder body must match Excel byte-for-byte or the
  // callee shows up as #NAME? on load. A prefixed name Excel does not know
  // (`_xlfn.FOOBAR`) is an ordinary undefined-name stub instead (measured).
  const bool is_param = name.rfind("_xlpm.", 0) == 0;
  std::string_view callee(name);
  const bool xlfn = callee.rfind("_xlfn.", 0) == 0;
  callee.remove_prefix(xlfn ? 6U : 0U);
  if (callee.rfind("_xlws.", 0) == 0) {
    callee.remove_prefix(6U);
  }
  const bool is_future_fn = xlfn && has_storage_prefix(callee);
  const bool is_placeholder = is_param || is_future_fn;
  std::uint32_t flags;
  if (is_param) {
    flags = 0x00020019U;
  } else if (is_future_fn) {
    flags = 0x0002000bU;
  } else {
    flags = hidden ? 0x00000001U : 0x00000000U;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, flags);  // grbit (flags word)
  emit_u8(p, 0);       // chKey
  emit_u32(p, static_cast<std::uint32_t>(itab));
  emit_xlwidestring(p, name);
  if (is_placeholder) {
    emit_u32(p, 2);    // cce
    emit_u8(p, 0x1C);  // PtgErr
    emit_u8(p, 0x1D);  // #NAME? error code
    emit_u32(p, 0);    // cb
  } else if (formula.empty()) {
    emit_u32(p, 0);  // cce
    emit_u32(p, 0);  // cb
  } else {
    Arena arena;
    parser::Parser parser(formula, arena);
    parser::AstNode* root = parser.parse();
    if (root == nullptr || !parser.errors().empty()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb writer: defined-name formula failed to parse for Ptg encoding",
                        std::string("context=xlsb_writer name=") + name);
    }
    if (calc_exp) {
      p[0] |= 0x10U;  // fCalcExp (grbit bit 4)
    }
    // Measured: a defined-name body's root reference stays reference class.
    auto encoded_or = encode_ptgs(*root, sheet_names, sheet_ranges, name_table, PtgRootClass::kReference);
    if (!encoded_or) {
      return encoded_or.error();
    }
    const EncodedFormula& encoded = encoded_or.value();
    emit_u32(p, static_cast<std::uint32_t>(encoded.rgce.size()));
    p.insert(p.end(), encoded.rgce.begin(), encoded.rgce.end());
    emit_u32(p, static_cast<std::uint32_t>(encoded.rgcb.size()));
    p.insert(p.end(), encoded.rgcb.begin(), encoded.rgcb.end());
  }
  // Trailing BrtName strings ([MS-XLSB] §2.4.649). Excel rejects a
  // BrtName that stops after the formula. A plain defined name carries
  // just the comment (a null `XLNullableWideString` when absent, matching
  // Excel's own encoding for a name whose Name Manager "Comment" field
  // was never set -- not a zero-length string); a future-function /
  // proc-parameter placeholder carries five strings (comment plus four
  // further unused strings, always null -- these are internal storage
  // artifacts with no Name Manager entry a comment could attach to),
  // matching real Excel output.
  if (is_placeholder) {
    for (int i = 0; i < 5; ++i) {
      emit_u32(p, 0xFFFFFFFFU);  // null XLNullableWideString
    }
  } else {
    emit_xlnullablewidestring(p, comment.empty() ? std::nullopt : std::make_optional(comment));
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtName), p);
  return Expected<void, Error>::Ok();
}

}  // namespace

// ---------------------------------------------------------------------------
// Workbook stream (xl/workbook.bin)
// ---------------------------------------------------------------------------

Expected<std::vector<std::uint8_t>, Error> BuildWorkbookBin(const Workbook& wb,
                                                            const std::vector<OrderedName>& ordered_names,
                                                            const SheetRangeTable& sheet_ranges,
                                                            const std::vector<std::string>& sheet_names,
                                                            const std::vector<std::string>& link_rel_ids) {
  std::vector<std::uint8_t> body;
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBook), ByteSpan{});

  // Workbook globals Excel expects before the sheet bundle. These are
  // fixed-shape records with no dependency on workbook content, so we emit
  // known-valid default payloads (byte layout captured from a real Excel 365
  // `xl/workbook.bin`):
  //   * BrtFileVersion : appName "xl", lastEdited/lowestEdited "7", build.
  //   * BrtWbProp      : default flags + defaultThemeVersion 202300.
  //   * BrtWbView      : a single window view; itabCur (offset 24, u32) is
  //                      patched below to the tab-selected sheet.
  // Omitting these produced a workbook stream Excel rejected.
  static const std::vector<std::uint8_t> kFileVersionPayload = {
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02,
      0x00, 0x00, 0x00, 0x78, 0x00, 0x6c, 0x00, 0x01, 0x00, 0x00, 0x00, 0x37, 0x00, 0x01, 0x00, 0x00, 0x00,
      0x37, 0x00, 0x05, 0x00, 0x00, 0x00, 0x31, 0x00, 0x30, 0x00, 0x36, 0x00, 0x32, 0x00, 0x38, 0x00};
  static const std::vector<std::uint8_t> kDefaultWbPropPayload = {0x20, 0x00, 0x01, 0x00, 0x3c, 0x16,
                                                                  0x03, 0x00, 0x00, 0x00, 0x00, 0x00};
  static const std::vector<std::uint8_t> kWbViewPayload = {0x10, 0x4f, 0x00, 0x00, 0x18, 0x15, 0x00, 0x00, 0x8c, 0x6e,
                                                           0x00, 0x00, 0xd0, 0x43, 0x00, 0x00, 0x58, 0x02, 0x00, 0x00,
                                                           0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x78};
  emit_record(body, kBrtFileVersion, kFileVersionPayload);
  std::vector<std::uint8_t> wb_prop_payload = kDefaultWbPropPayload;
  // BrtWbProp ([MS-XLSB] §2.4.866): f1904 is bit 0 of the leading
  // little-endian u32 grbit. Preserve Excel's captured default flags.
  if (wb.date1904()) {
    wb_prop_payload[0] |= 0x01U;
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp), wb_prop_payload);
  // Excel places the protection records between BrtWbProp and the views.
  if (auto protection = emit_book_protection(body, wb.workbook_protection_xml()); !protection) {
    return protection.error();
  }
  emit_record(body, kBrtBeginBookViews, ByteSpan{});
  // itabCur (u32 at offset 24): measured to equal the tab-selected sheet's
  // index in every fixture; a mismatch made Excel show every tab selected.
  std::vector<std::uint8_t> wb_view_payload = kWbViewPayload;
  std::uint32_t itab_cur = 0U;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).view().tab_selected) {
      itab_cur = static_cast<std::uint32_t>(i);
      break;
    }
  }
  wb_view_payload[24] = static_cast<std::uint8_t>(itab_cur & 0xFFU);
  wb_view_payload[25] = static_cast<std::uint8_t>((itab_cur >> 8) & 0xFFU);
  wb_view_payload[26] = static_cast<std::uint8_t>((itab_cur >> 16) & 0xFFU);
  wb_view_payload[27] = static_cast<std::uint8_t>((itab_cur >> 24) & 0xFFU);
  emit_record(body, kBrtWbView, wb_view_payload);
  emit_record(body, kBrtEndBookViews, ByteSpan{});

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBundleShs), ByteSpan{});
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    // BrtBundleSh ([MS-XLSB] §2.4.304):
    //   hsState    : u32 (0 = visible, 1 = hidden, 2 = very hidden)
    //   iTabID     : u32 (sheet id; 1-based)
    //   strRelID   : XLNullableWideString
    //   strName    : XLWideString
    std::vector<std::uint8_t> p;
    // `SheetVisibility` is numbered to match `hsState`, so the model state
    // is the field value.
    const auto hs_state = static_cast<std::uint32_t>(wb.sheet(i).view().visibility());
    emit_u32(p, hs_state);                            // hsState
    emit_u32(p, static_cast<std::uint32_t>(i + 1U));  // iTabID
    const std::string rid = std::string("rId") + std::to_string(i + 1U);
    emit_xlnullablewidestring(p, std::optional<std::string_view>{rid});
    emit_xlwidestring(p, wb.sheet(i).name());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh), p);
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBundleShs), ByteSpan{});

  // BrtExternSheet: every distinct sheet-qualified reference span this
  // workbook's formulas need an `ixti` for (single- and multi-sheet
  // alike; see `BuildSheetRangeTable`'s doc comment for why both share
  // this one table). Omitted entirely when no formula uses a qualified
  // reference and there is no external link, matching the fallback the
  // reader's `sheet_for_ixti` / `sheet_range_for_ixti` already implement
  // for that case.
  //
  // The record MUST be wrapped in the externals block: a bare
  // `BrtExternSheet` outside `BrtBeginExternals ... BrtEndExternals` is an
  // out-of-place record that makes Excel reject the package. The block lists
  // the supporting books in Excel's order: `BrtSupSelf` when an XTI names
  // this workbook, then one `BrtSupBookSrc` per external link, named by its
  // workbook relationship. Every link is listed, referenced or not, so a
  // reload finds each part.
  if (!sheet_ranges.xti.empty() || !link_rel_ids.empty()) {
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternals), ByteSpan{});
    if (xti_names_self(sheet_ranges)) {
      emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSupSelf), ByteSpan{});
    }
    for (const std::string& rel_id : link_rel_ids) {
      std::vector<std::uint8_t> p;
      emit_xlwidestring(p, rel_id);
      emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSupBookSrc), p);
    }
    EmitExternSheet(body, sheet_ranges);
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndExternals), ByteSpan{});
  }

  // BrtName table: every genuine defined name (hidden per its own OOXML
  // `hidden` flag) followed by every future-function callee / defined
  // name `collect_ptg_names` needed a `PtgName` reference for that
  // wasn't already a defined name (the `_xlfn.` / `_xlpm.` records
  // hidden, a stub for an undefined name not). `ordered_names` is built once by `BuildOrderedNames`
  // and shared with every sheet's cell encoder so `ilbl` assignments
  // stay consistent workbook-wide. The leading `defined_count` slots are
  // `wb.defined_names()` element-for-element, so every defined name
  // emits exactly one record at the `ilbl` the encoder assigned it.
  //
  // A name's own formula resolves other names from the scope the name
  // itself lives in, the same rule a cell formula follows, so each entry
  // encodes against a table built for its `local_sheet_id`. The tables
  // are memoized per scope: workbooks routinely carry many names in the
  // same scope and each table is a full pass over the name list.
  const std::size_t defined_count = wb.defined_names().size();
  std::unordered_map<std::int32_t, NameTable> scoped_tables;
  auto table_for_scope = [&](std::int32_t scope_sheet_id) -> const NameTable& {
    auto it = scoped_tables.find(scope_sheet_id);
    if (it == scoped_tables.end()) {
      it = scoped_tables.emplace(scope_sheet_id, BuildNameTableForScope(wb, ordered_names, scope_sheet_id)).first;
    }
    return it->second;
  };
  const std::vector<NameShape> shapes = defined_name_shapes(wb);
  for (std::size_t i = 0; i < ordered_names.size(); ++i) {
    if (i < defined_count) {
      const DefinedName& dn = wb.defined_names()[i];
      if (auto r = EmitName(body, dn.name, dn.formula, dn.local_sheet_id, dn.hidden, shapes[i].calc_exp, dn.comment,
                            sheet_names, sheet_ranges, table_for_scope(dn.local_sheet_id));
          !r) {
        return r.error();
      }
    } else {
      // Placeholders carry no formula body, so the table they are handed is
      // never consulted; the workbook-scope one keeps the call uniform. A
      // stub for an undefined name is not hidden, as Excel saves it.
      const OrderedName& slot = ordered_names[i];
      if (auto r = EmitName(body, slot.name, /*formula=*/{}, slot.itab, /*hidden=*/false, /*calc_exp=*/false,
                            /*comment=*/{}, sheet_names, sheet_ranges, table_for_scope(-1));
          !r) {
        return r.error();
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBook), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
