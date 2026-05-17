// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Locale-driven labelling for the pivot-grid layout layer.
//
// The pivot projection in `src/pivot/pivot_layout.{h,cpp}` is locale-
// agnostic; it leaves the choice of placeholders ("Row Labels" /
// "Column Labels"), subtotal suffix, grand-total label, and data-field
// name template up to the caller. This translation unit centralises
// those strings per `eval::ExcelProfile`, and exposes the helper that
// turns a profile into a populated `pivot::PivotLayoutOptions`.

#ifndef FORMULON_EVAL_PIVOT_LOCALE_H_
#define FORMULON_EVAL_PIVOT_LOCALE_H_

#include <string>
#include <string_view>

#include "eval/compat.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_types.h"

namespace formulon::eval {

/// Returns the layout-option overrides appropriate for the locale of
/// the given Excel profile.
///
/// The host (Mac vs Windows) is intentionally ignored: pivot labels
/// follow the workbook's display locale, not the running Excel binary.
/// For non-ja locales the function returns the default-constructed
/// `PivotLayoutOptions{}` so callers keep the legacy English layout.
pivot::PivotLayoutOptions pivot_layout_options_for(ExcelProfile profile);

/// Returns the localized aggregation label used in pivot data-field
/// headers (e.g. "Sum of" / "合計"). The returned view points at
/// static storage and is safe to copy or compare.
std::string_view aggregation_label(pivot::Aggregation agg, ExcelProfile profile);

/// Returns the separator placed between the localized aggregation
/// label and the source field name in a data-field display name
/// ("合計 / Amount" uses " / "; "Sum of Amount" uses " of ").
std::string_view data_field_separator(ExcelProfile profile);

/// Formats the full data-field display name for the locale of the
/// given profile, e.g. "合計 / Amount" (ja-JP) or "Sum of Amount"
/// (default). The English form preserves the historical
/// `<Agg> of <field>` shape used by the workbook-oracle harness and
/// the OOXML round-trip layer.
std::string data_field_display_name(pivot::Aggregation agg, std::string_view field_name, ExcelProfile profile);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_PIVOT_LOCALE_H_
