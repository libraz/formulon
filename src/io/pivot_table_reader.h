// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Reader for one OOXML `xl/pivotTables/pivotTable*.xml` part. Parses the
// pivot-table definition into a `formulon::pivot::PivotTable`. Does NOT
// resolve the cache link (cacheId integer is captured verbatim; the
// caller looks it up against the workbook's pivot caches).
//
// Design references:
//   * backup/plans/15-pivot-and-advanced.md §15.1.1 / §15.1.4
//   * src/io/pivot_cache_reader.h (sister reader)

#ifndef FORMULON_IO_PIVOT_TABLE_READER_H_
#define FORMULON_IO_PIVOT_TABLE_READER_H_

#include <cstdint>
#include <vector>

#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Parses one `xl/pivotTables/pivotTable*.xml` part.
///
/// Behaviour:
///   * Root must be `<pivotTableDefinition>`. Otherwise
///     `kIoContentTypeInvalid`.
///   * `name` and `cacheId` come from the root attributes;
///     `cacheId` defaults to 0 if missing/non-numeric (Excel always
///     emits it).
///   * `<location>` `ref` attribute is required and must parse as
///     a valid A1:A1 range. Missing/empty/unparseable -> `kIoSheetCorrupt`.
///     The parsed range becomes `anchor_row` / `anchor_col` /
///     `span_rows` / `span_cols`.
///   * `<pivotFields>` walked in document order. Each `<pivotField>`:
///       - axis attribute mapped:
///           "axisRow"  -> PivotAxis::Row
///           "axisCol"  -> PivotAxis::Col
///           "axisPage" -> PivotAxis::Page
///           "axisValues" / unset -> PivotAxis::Value (Excel writes
///             `axisValues` for fields exclusively used as data, but
///             unset is also common — both fold to Value here).
///       - `name` attribute -> custom_name (Excel uses this when
///         the user has renamed the field; falls back to the cache
///         field name when absent — the caller resolves that).
///       - `subtotalTop` "1" / "0" -> subtotal_top.
///       - `<items>` walked in document order. `<item t="default">` /
///         `<item t="grand">` are skipped (subtotal/grand-total
///         entries; not real items). `<item x="N">` references the
///         shared-item index in the cache; we capture the index but
///         leave the `name` field empty (resolved at evaluator time).
///         `<item h="1">` flips visibility off.
///   * `<rowFields>` -> row_field_order (each `<field x="N">` -> N).
///     `<colFields>` -> col_field_order similarly.
///   * `<dataFields>` walked in document order. Each `<dataField>`:
///       - `name` -> PivotDataField::name (required; missing/empty
///         is `kIoSheetCorrupt`).
///       - `fld` -> field_index (defaults to 0).
///       - `subtotal` mapped:
///           "sum"     -> Aggregation::Sum  (default)
///           "count"   -> Aggregation::Count
///           "average" -> Aggregation::Average
///           "max"     -> Aggregation::Max
///           "min"     -> Aggregation::Min
///           "product" -> Aggregation::Product
///           "countNums" -> Aggregation::CountNumbers
///           "stdDev"  -> Aggregation::StdDev
///           "stdDevp" -> Aggregation::StdDevP
///           "var"     -> Aggregation::Var
///           "varp"    -> Aggregation::VarP
///         Unrecognized -> Aggregation::Sum (with no error; future-
///         compat with Excel additions).
///       - `numFmtId` -> number_format (we capture it as a stringified
///         int; a follow-up converts it to a real format code via
///         the styles part).
///   * `<pageFields>` is parsed but only for round-trip preservation
///     in the future; for this reader it is silently skipped.
///   * `<formats>`, `<conditionalFormats>`, `<chartFormats>`, etc. are
///     silently skipped (preserved as bytes by a future writer).
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the bytes.
///   * `kIoContentTypeInvalid` — root is not `<pivotTableDefinition>`.
///   * `kIoSheetCorrupt` — `<location ref>` missing/unparseable, or a
///     `<dataField>` has no `name`.
Expected<pivot::PivotTable, Error> read_pivot_table_definition(const std::vector<std::uint8_t>& definition_bytes);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_TABLE_READER_H_
