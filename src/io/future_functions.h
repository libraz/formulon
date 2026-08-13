//
// Storage-prefix classification for Excel function names — the single
// source both persistence writers consult.
//
// Excel stores a call to a function introduced after the Excel 2007 file
// format was frozen under a hidden name: `_xlfn.<NAME>`, or
// `_xlfn._xlws.<NAME>` for the worksheet-only dynamic-array forms. An
// older Excel then resolves the hidden name through the defined-name
// table and shows `#NAME?` instead of mis-evaluating an unknown callee.
// The formula bar never shows the prefix, so Formulon strips it on
// ingestion (`io::strip_storage_prefixes`) and re-applies it on save
// through this module.
//
// The prefixed set is enumerated by name. It is deliberately NOT derived
// from the absence of an id in `io/xlsb/func_id_table.h`: that table is
// grown incrementally, so a missing id says nothing about when the
// function was introduced. Inferring "future" from it prefixes the
// Analysis ToolPak set that Excel 2007 made native (MROUND, YEARFRAC,
// WEEKNUM, ISEVEN / ISODD, and the financial and engineering families),
// which real Excel then renders as `#NAME?`.
//
// The OOXML `<f>` writer and the XLSB Ptg encoder both classify here, so
// the two formats cannot disagree about how a callee is spelled.

#ifndef FORMULON_IO_FUTURE_FUNCTIONS_H_
#define FORMULON_IO_FUTURE_FUNCTIONS_H_

#include <string>
#include <string_view>

#include "parser/ast_format.h"

namespace formulon {
namespace io {

/// Classifies `canonical_name` (the unprefixed formula-bar spelling, any
/// case) into the storage prefix Excel writes for a call to it.
/// A name outside the enumerated future-function set classifies as
/// `None`, i.e. it is stored exactly as spelled.
parser::StoragePrefixKind classify_storage_prefix(std::string_view canonical_name);

/// Resolves a localised formula-bar spelling to the name Excel actually
/// stores, and returns `name` unchanged for everything else.
///
/// Excel localises the formula bar but not the file: the ja-JP UI spells
/// `DBCS` as `JIS`, and typing `JIS(...)` there yields a cell whose
/// stored text — `<f>` and the XLSB function id alike — says `DBCS`.
/// A model can reach either writer carrying either spelling, because the
/// engine registers both against one implementation, so both persistence
/// paths resolve here. This is the only place aliases are enumerated:
/// a second table would drift, and a container that skipped it would
/// store a name Excel's invariant grammar does not have (`#NAME?` on
/// open).
std::string_view canonical_function_name(std::string_view name);

/// Returns the exact text the OOXML `<f>` element carries for a call to
/// `name`: the storage prefix, then the `canonical_function_name`.
///
/// **Case is preserved**, which is the one difference from its sibling
/// `storage_function_name` — the two are otherwise the same computation
/// and the names do not say so. `<f>` text is round-tripped, so a model
/// holding `sum` must save as `sum`; upper-casing here would rewrite the
/// `<f>` text of every classic formula on save. Excel upper-cases a name
/// when the *user* commits a formula, which is a different event from
/// serialising a model.
std::string storage_call_name(std::string_view name);

/// Returns the name Excel stores for a call to `canonical_name`: the
/// upper-cased `canonical_function_name` with its
/// `classify_storage_prefix` prefix applied.
///
/// **Upper-cases**, unlike its sibling `storage_call_name` — this one
/// spells a hidden `BrtName`, an identifier the XLSB `PtgName`
/// future-function route resolves through rather than text a user ever
/// sees, and Excel writes those upper-cased. Use `storage_call_name`
/// for anything that becomes visible formula text.
std::string storage_function_name(std::string_view canonical_name);

/// True when the XLSB Ptg encoder must reach a call to `canonical_name`
/// through the hidden-name route (`PtgName` naming the callee plus
/// `PtgFuncVar` with the `id == 255` sentinel) even though OOXML spells
/// the call bare.
///
/// This is a second, independent axis from `classify_storage_prefix`.
/// That one answers "how is the callee spelled in the `<f>` text"; this
/// one answers "does Excel's classic function table have an id for the
/// callee at all". Excel-365 stores `ISO.CEILING` bare in `<f>` yet has
/// no function id for it, so the two answers genuinely disagree for it
/// and neither can be derived from the other.
///
/// Membership is enumerated and requires a decoded observation of a
/// real Excel-produced `.xlsb` — the same bar `io/xlsb/func_id_table.h`
/// sets for an id. Absence of an id from `func_id_table` is NOT
/// evidence of membership: that table is grown incrementally, so
/// treating "no id yet" as "Excel has no id" would write a hidden
/// `_xlfn.<NAME>` that real Excel resolves to `#NAME?`.
bool xlsb_uses_hidden_name(std::string_view canonical_name);

/// Returns the name the hidden `BrtName` carries for a call to
/// `canonical_name` encoded through the XLSB hidden-name route.
///
/// For a `_xlfn.*` future function this is `storage_function_name`. For
/// a member of the `xlsb_uses_hidden_name` set it is `_xlfn.<NAME>`
/// regardless of the OOXML spelling, because `_xlfn.` there is the
/// naming convention of the hidden-name route rather than a statement
/// about how the callee is spelled in `<f>`. Callers must have
/// established that the call takes that route.
std::string xlsb_hidden_function_name(std::string_view canonical_name);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_FUTURE_FUNCTIONS_H_
