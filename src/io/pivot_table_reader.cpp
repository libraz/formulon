//
// Implementation of the pivot-table-definition reader. See
// pivot_table_reader.h for the public contract.

#include "io/pivot_table_reader.h"

#include <cstdint>
#include <cstring>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/xml_utils.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon::io {
namespace {

/// Maps an OOXML `axis="..."` attribute body to a `PivotAxis`. Unknown /
/// missing values fold to `Value` — both `axisValues` and the absent-
/// axis case mean "this field is exclusively used as a data field" in
/// the OOXML spec.
pivot::PivotAxis ParseAxis(std::string_view text) {
  if (text == "axisRow") {
    return pivot::PivotAxis::Row;
  }
  if (text == "axisCol") {
    return pivot::PivotAxis::Col;
  }
  if (text == "axisPage") {
    return pivot::PivotAxis::Page;
  }
  return pivot::PivotAxis::Value;
}

/// Maps an OOXML `subtotal="..."` attribute body on a `<dataField>` to an
/// `Aggregation`. Unknown spellings fold to `Sum` so the reader stays
/// forward-compatible with Excel additions; the writer round-trips the
/// original attribute through a separate passthrough path (future PR).
pivot::Aggregation ParseAggregation(std::string_view text) {
  if (text == "count") {
    return pivot::Aggregation::Count;
  }
  if (text == "average") {
    return pivot::Aggregation::Average;
  }
  if (text == "max") {
    return pivot::Aggregation::Max;
  }
  if (text == "min") {
    return pivot::Aggregation::Min;
  }
  if (text == "product") {
    return pivot::Aggregation::Product;
  }
  if (text == "countNums") {
    return pivot::Aggregation::CountNumbers;
  }
  if (text == "stdDev") {
    return pivot::Aggregation::StdDev;
  }
  if (text == "stdDevp") {
    return pivot::Aggregation::StdDevP;
  }
  if (text == "var") {
    return pivot::Aggregation::Var;
  }
  if (text == "varp") {
    return pivot::Aggregation::VarP;
  }
  // "sum" and anything unrecognised.
  return pivot::Aggregation::Sum;
}

/// Maps an OOXML `showDataAs="..."` attribute body to a `ShowValuesAs`.
/// Unknown spellings fold to `Normal` so the reader stays
/// forward-compatible with Excel additions; the writer's inverse helper
/// produces the same set of attribute names. The mapping for
/// `RunningTotalInCol` uses the Excel-private `runTotalInCol` spelling
/// since the standard OOXML schema only specifies a row-direction
/// `runTotal`; this is documented as a known scoping in the round-trip
/// table on the `ShowValuesAs` enum.
pivot::ShowValuesAs ParseShowDataAs(std::string_view text) {
  if (text == "percentOfRow") {
    return pivot::ShowValuesAs::PercentOfRow;
  }
  if (text == "percentOfCol") {
    return pivot::ShowValuesAs::PercentOfCol;
  }
  if (text == "percentOfTotal") {
    return pivot::ShowValuesAs::PercentOfTotal;
  }
  if (text == "runTotal") {
    return pivot::ShowValuesAs::RunningTotalInRow;
  }
  if (text == "runTotalInCol") {
    return pivot::ShowValuesAs::RunningTotalInCol;
  }
  if (text == "index") {
    return pivot::ShowValuesAs::Index;
  }
  if (text == "difference") {
    return pivot::ShowValuesAs::DifferenceFrom;
  }
  if (text == "percentDiff") {
    return pivot::ShowValuesAs::PercentDifferenceFrom;
  }
  if (text == "percentOfParentRow") {
    return pivot::ShowValuesAs::PercentOfParentRow;
  }
  if (text == "percentOfParentCol") {
    return pivot::ShowValuesAs::PercentOfParentCol;
  }
  if (text == "percentOfParent") {
    return pivot::ShowValuesAs::PercentOfParent;
  }
  return pivot::ShowValuesAs::Normal;
}

/// Decodes one `<location ref="A3:D10"/>` (or single-cell `ref="A3"`)
/// into 0-based anchor + spans. Defers cell-coordinate decoding to the
/// shared `parse_a1` helper so the conventions match cell parsing
/// elsewhere.
Expected<void, Error> DecodeLocationRef(std::string_view ref, pivot::PivotTable* out) {
  if (ref.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt,
                      "pivotTableDefinition: <location> missing required ref attribute", "context=pivot_table_reader");
  }
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    // Single-cell anchor (Excel writes this for an empty pivot).
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=pivot_table_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <location> ref unparseable",
                        std::move(ctx));
    }
    out->set_anchor(rc.value().first, rc.value().second, 1U, 1U);
    return Expected<void, Error>::Ok();
  }
  const std::string_view a = ref.substr(0, colon);
  const std::string_view b = ref.substr(colon + 1);
  auto a_rc = parse_a1(a);
  auto b_rc = parse_a1(b);
  if (!a_rc || !b_rc) {
    std::string ctx("context=pivot_table_reader ref=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <location> ref unparseable",
                      std::move(ctx));
  }
  const std::uint32_t r0 = a_rc.value().first;
  const std::uint32_t c0 = a_rc.value().second;
  const std::uint32_t r1 = b_rc.value().first;
  const std::uint32_t c1 = b_rc.value().second;
  const std::uint32_t row_top = (r0 < r1) ? r0 : r1;
  const std::uint32_t row_bot = (r0 < r1) ? r1 : r0;
  const std::uint32_t col_left = (c0 < c1) ? c0 : c1;
  const std::uint32_t col_right = (c0 < c1) ? c1 : c0;
  out->set_anchor(row_top, col_left, row_bot - row_top + 1U, col_right - col_left + 1U);
  return Expected<void, Error>::Ok();
}

/// Walks the `<items>` child of a `<pivotField>` and populates
/// `field.items`. `<item t="default">` and `<item t="grand">` are
/// subtotal / grand-total markers in the OOXML schema, NOT real items;
/// they are skipped here. For real items we capture visibility from
/// `h="1"` (hidden -> `visible = false`) and the `x` attribute, which
/// indexes the bound cache's `shared_items`. `name` is left empty here
/// and resolved against the cache after both parts are loaded (see
/// `resolve_pivot_names`). When `x` is absent we fall back to the item's
/// document-order position among real items so name resolution still has
/// an index to look up; `has_cache_index` stays false so the writer can
/// re-emit the original (attribute-absent) form.
void ParseItems(const pugi::xml_node& items_node, pivot::PivotField* field) {
  std::uint32_t ordinal = 0;
  for (pugi::xml_node it = items_node.child("item"); it; it = it.next_sibling("item")) {
    const std::string_view t = it.attribute("t").as_string();
    if (t == "default" || t == "grand" || t == "blank" || t == "sum" || t == "count" || t == "avg" || t == "max" ||
        t == "min" || t == "product" || t == "countA" || t == "stdDev" || t == "stdDevP" || t == "var" || t == "varP") {
      // Subtotal / grand-total marker; not a real item.
      continue;
    }
    pivot::PivotItem entry;
    entry.visible = !parse_xml_bool_attr(it.attribute("h"));
    if (pugi::xml_attribute x = it.attribute("x"); x) {
      entry.has_cache_index = true;
      entry.cache_index = parse_xml_u32_attr(x, 0U);
    } else {
      entry.cache_index = ordinal;
    }
    field->items.push_back(std::move(entry));
    ++ordinal;
  }
}

/// Pairs an OOXML `<pivotField>` `*Subtotal` boolean attribute name with
/// the `SubtotalFn` it selects. `defaultSubtotal` is handled separately
/// (it gates the implicit default rather than a custom function) and is
/// not part of this table. The ordering is the canonical ECMA-376
/// attribute order so the writer can re-emit deterministically.
struct SubtotalAttrEntry {
  std::string_view attr;
  pivot::SubtotalFn fn;
};
constexpr SubtotalAttrEntry kSubtotalAttrs[] = {
    {"sumSubtotal", pivot::SubtotalFn::Sum},
    {"countASubtotal", pivot::SubtotalFn::Count},
    {"avgSubtotal", pivot::SubtotalFn::Average},
    {"maxSubtotal", pivot::SubtotalFn::Max},
    {"minSubtotal", pivot::SubtotalFn::Min},
    {"productSubtotal", pivot::SubtotalFn::Product},
    {"countSubtotal", pivot::SubtotalFn::CountNumbers},
    {"stdDevSubtotal", pivot::SubtotalFn::StdDev},
    {"stdDevPSubtotal", pivot::SubtotalFn::StdDevP},
    {"varSubtotal", pivot::SubtotalFn::Var},
    {"varPSubtotal", pivot::SubtotalFn::VarP},
};

/// Walks `<pivotFields>` in document order, materialising each
/// `<pivotField>` into the table.
void ParsePivotFields(const pugi::xml_node& fields_node, pivot::PivotTable* out) {
  for (pugi::xml_node f = fields_node.child("pivotField"); f; f = f.next_sibling("pivotField")) {
    pivot::PivotField field;
    // Axis resolution: an explicit `axis` attribute wins; otherwise a
    // `dataField="1"` marks a Value field; a field with neither is an
    // unused ("available") field, which must round-trip as None rather
    // than being stamped `dataField="1"` on write.
    if (pugi::xml_attribute axis_attr = f.attribute("axis"); axis_attr) {
      field.axis = ParseAxis(axis_attr.as_string());
    } else if (attr_bool(f, "dataField", false)) {
      field.axis = pivot::PivotAxis::Value;
    } else {
      field.axis = pivot::PivotAxis::None;
    }
    if (pugi::xml_attribute name_attr = f.attribute("name"); name_attr) {
      field.custom_name = name_attr.value();
    }
    // `subtotalTop` defaults to true in OOXML (subtotals render above the
    // group). Only an explicit "0" moves them to the bottom.
    field.subtotal_top = attr_bool(f, "subtotalTop", true);
    // `defaultSubtotal` defaults to true in OOXML; only an explicit "0"
    // turns the implicit default subtotal off. The `*Subtotal` boolean
    // family selects custom subtotal functions; capture each present /
    // true attribute so a custom selection survives the round trip
    // instead of silently reverting to the default after save.
    if (pugi::xml_attribute ds = f.attribute("defaultSubtotal"); ds) {
      field.default_subtotal = parse_xml_bool_attr(ds);
    }
    for (const SubtotalAttrEntry& e : kSubtotalAttrs) {
      if (parse_xml_bool_attr(f.attribute(e.attr.data()))) {
        field.subtotal_fns.push_back(e.fn);
      }
    }
    if (pugi::xml_node items_node = f.child("items"); items_node) {
      ParseItems(items_node, &field);
    }
    // Preserve unmodelled `<pivotField>` attributes (`compact`, `outline`,
    // `showAll`, `includeNewItemsInFilter`, ...) for verbatim round-trip.
    capture_unknown_attrs(f,
                          {"axis", "dataField", "name", "subtotalTop", "defaultSubtotal", "sumSubtotal",
                           "countASubtotal", "avgSubtotal", "maxSubtotal", "minSubtotal", "productSubtotal",
                           "countSubtotal", "stdDevSubtotal", "stdDevPSubtotal", "varSubtotal", "varPSubtotal"},
                          field.passthrough_attrs);
    out->mutable_fields().push_back(std::move(field));
  }
}

/// Walks `<rowFields>` / `<colFields>` and pushes each `<field x="N">`
/// onto `out`. Missing `x` defaults to 0 (matches OOXML schema).
void ParseFieldOrder(const pugi::xml_node& parent, std::vector<std::uint32_t>* out) {
  for (pugi::xml_node f = parent.child("field"); f; f = f.next_sibling("field")) {
    out->push_back(parse_xml_u32_attr(f.attribute("x"), 0U));
  }
}

/// Walks `<dataFields>` in document order. Returns `kIoSheetCorrupt` on
/// the first `<dataField>` that lacks a `name` attribute, since that
/// name is the GETPIVOTDATA lookup key.
Expected<void, Error> ParseDataFields(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node df = parent.child("dataField"); df; df = df.next_sibling("dataField")) {
    pivot::PivotDataField entry;
    pugi::xml_attribute name_attr = df.attribute("name");
    if (!name_attr || name_attr.value() == nullptr || *name_attr.value() == '\0') {
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "pivotTableDefinition: <dataField> missing required name attribute",
                        "context=pivot_table_reader");
    }
    entry.name = name_attr.value();
    entry.field_index = parse_xml_u32_attr(df.attribute("fld"), 0U);
    entry.aggregation = ParseAggregation(df.attribute("subtotal").as_string());
    if (pugi::xml_attribute nf = df.attribute("numFmtId"); nf) {
      entry.number_format = nf.value();
    }
    if (pugi::xml_attribute sa = df.attribute("showDataAs"); sa) {
      entry.show_as = ParseShowDataAs(sa.as_string());
    }
    if (pugi::xml_attribute bf = df.attribute("baseField"); bf) {
      entry.show_as_base_field = parse_xml_u32_attr(bf, 0U);
    }
    if (pugi::xml_attribute bi = df.attribute("baseItem"); bi) {
      entry.show_as_base_item = parse_xml_u32_attr(bi, 0U);
    }
    out->mutable_data_fields().push_back(std::move(entry));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<pivot::PivotTable, Error> read_pivot_table_definition(const std::vector<std::uint8_t>& definition_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, definition_bytes, "pivot_table_reader", "pivotTable*.xml"));
  pugi::xml_node root = doc.child("pivotTableDefinition");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "pivotTable*.xml: missing <pivotTableDefinition> root",
                      "context=pivot_table_reader");
  }

  pivot::PivotTable table;
  if (pugi::xml_attribute name_attr = root.attribute("name"); name_attr) {
    table.set_name(name_attr.value());
  }
  // `dataCaption` is a required attribute; capture the authored value so
  // the writer re-emits it verbatim (default stays "Values" when absent).
  if (pugi::xml_attribute cap = root.attribute("dataCaption"); cap) {
    table.set_data_caption(cap.value());
  }
  table.set_pivot_cache_id(parse_xml_u32_attr(root.attribute("cacheId"), 0U));

  // Grand-total layout flags. OOXML defaults both to true when absent, so a
  // file that turned grand totals OFF carries `rowGrandTotals="0"` /
  // `colGrandTotals="0"` explicitly. The model also defaults to true; we
  // read whatever is present so an OFF state survives the round trip
  // instead of silently flipping back to ON.
  {
    const bool row_grand = attr_bool(root, "rowGrandTotals", true);
    const bool col_grand = attr_bool(root, "colGrandTotals", true);
    table.set_grand_totals(row_grand, col_grand);
  }

  // Report layout mode. OOXML expresses it on `<pivotTableDefinition>` via
  // `compact` (default true) and `outline` (default false): `outline="1"`
  // is Outline form, an explicit `compact="0"` (with outline off) is
  // Tabular, and the default is Compact. The writer re-derives the same
  // attributes so the layout round-trips and drives the renderer.
  {
    const bool outline = attr_bool(root, "outline", false);
    const bool compact = attr_bool(root, "compact", true);
    if (outline) {
      table.set_layout(pivot::PivotLayout::Outline);
    } else if (compact) {
      table.set_layout(pivot::PivotLayout::Compact);
    } else {
      table.set_layout(pivot::PivotLayout::Tabular);
    }
  }

  // Capture every root attribute the model does not represent structurally
  // (`updatedVersion`, `createdVersion`, `itemPrintTitles`, `indent`, ...)
  // so the writer re-emits them verbatim.
  capture_unknown_attrs(root,
                        {"name", "cacheId", "dataCaption", "rowGrandTotals", "colGrandTotals", "compact", "outline"},
                        table.mutable_passthrough_attrs());

  pugi::xml_node loc = root.child("location");
  if (!loc) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTable*.xml: missing <location> element",
                      "context=pivot_table_reader");
  }
  pugi::xml_attribute ref_attr = loc.attribute("ref");
  if (!ref_attr) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTable*.xml: <location> missing required ref attribute",
                      "context=pivot_table_reader");
  }
  if (auto status = DecodeLocationRef(ref_attr.value(), &table); !status) {
    return status.error();
  }

  // Capture the `<location>` offset attributes. ECMA-376 requires
  // `firstHeaderRow`, `firstDataRow`, and `firstDataCol`; `rowPageCount`
  // and `colPageCount` are optional but commonly present. We preserve each
  // exactly as present (absent stays absent) so the writer re-emits a
  // schema-valid `<location>` on round trip. Dropping the required ones
  // makes Excel flag the file for repair and rebuild the pivot, losing
  // layout.
  auto opt_u32_attr = [](const pugi::xml_node& node, const char* name) -> std::optional<std::uint32_t> {
    const pugi::xml_attribute a = node.attribute(name);
    if (!a) {
      return std::nullopt;
    }
    return parse_xml_u32_attr(a, 0U);
  };
  table.set_location_attributes(opt_u32_attr(loc, "firstHeaderRow"), opt_u32_attr(loc, "firstDataRow"),
                                opt_u32_attr(loc, "firstDataCol"), opt_u32_attr(loc, "rowPageCount"),
                                opt_u32_attr(loc, "colPageCount"));

  if (pugi::xml_node fields = root.child("pivotFields"); fields) {
    ParsePivotFields(fields, &table);
  }
  if (pugi::xml_node rows = root.child("rowFields"); rows) {
    ParseFieldOrder(rows, &table.mutable_row_field_order());
  }
  if (pugi::xml_node cols = root.child("colFields"); cols) {
    ParseFieldOrder(cols, &table.mutable_col_field_order());
  }
  if (pugi::xml_node data = root.child("dataFields"); data) {
    if (auto status = ParseDataFields(data, &table); !status) {
      return status.error();
    }
  }
  // Capture any remaining direct children of <pivotTableDefinition> that
  // we do not model structurally into position-keyed raw-XML buffers, so
  // the writer can re-emit them at schema-valid slots.
  //
  // `CT_pivotTableDefinition` mandates a strict child order. `<rowItems>`,
  // `<colItems>`, and `<pageFields>` are the only unmodelled elements the
  // schema places *before* `<dataFields>`; they are binned by name into
  // the two pre-dataFields slots (rowItems after `<rowFields>`; colItems
  // and pageFields after `<colFields>`). Every other unmodelled element
  // (`<formats>`, `<conditionalFormats>`, `<chartFormats>`,
  // `<calculatedFields>`, `<calculatedItems>`, `<pivotTableStyleInfo>`,
  // `<extLst>`, ...) belongs after `<dataFields>` and goes to the tail
  // buffer. Relative order within each bin follows document order.
  struct StringXmlWriter : pugi::xml_writer {
    std::string* dst;
    void write(const void* data, std::size_t size) override { dst->append(static_cast<const char*>(data), size); }
  };
  static const std::string_view kRecognized[] = {"location", "pivotFields", "rowFields", "colFields", "dataFields"};
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = child.name();
    bool recognised = false;
    for (std::string_view r : kRecognized) {
      if (r == name) {
        recognised = true;
        break;
      }
    }
    if (recognised) {
      continue;
    }
    std::string* bucket = nullptr;
    if (name == "rowItems") {
      bucket = &table.mutable_raw_passthrough_after_row_fields();
    } else if (name == "colItems" || name == "pageFields") {
      bucket = &table.mutable_raw_passthrough_after_col_fields();
    } else {
      bucket = &table.mutable_raw_passthrough_xml();
    }
    StringXmlWriter sink{};
    sink.dst = bucket;
    child.print(sink, /*indent=*/"", pugi::format_raw);
  }

  return table;
}

}  // namespace formulon::io
