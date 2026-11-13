#include "io/dxf_uses_xlsx.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cf/cf_types.h"
#include "io/cf_writer.h"
#include "io/dxf_xml_ids.h"
#include "io/ooxml/emission_plan.h"
#include "io/passthrough_part_index.h"
#include "io/writer_common.h"
#include "passthrough_part.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "styles.h"
#include "table.h"
#include "workbook.h"

namespace formulon::io {
namespace {

/// The holder kind of an attribute found in raw XML, by its element's local name.
DxfUseKind KindOf(std::string_view element, DxfUseKind fallback) {
  struct ElementKind {
    std::string_view name;
    DxfUseKind kind;
  };
  static constexpr ElementKind kKinds[] = {
      {"cfRule", DxfUseKind::kConditionalFormat}, {"table", DxfUseKind::kTable},
      {"tableColumn", DxfUseKind::kTableColumn},  {"tableStyleElement", DxfUseKind::kTableStyleElement},
      {"colorFilter", DxfUseKind::kColorFilter},  {"sortCondition", DxfUseKind::kSortCondition},
      {"format", DxfUseKind::kPivotFormat},
  };
  const std::size_t colon = element.rfind(':');
  const std::string_view local = colon == std::string_view::npos ? element : element.substr(colon + 1U);
  for (const ElementKind& entry : kKinds) {
    if (entry.name == local) {
      return entry.kind;
    }
  }
  return fallback;
}

/// Records uses in document order within one part.
class Collector {
 public:
  void at(std::int32_t sheet, std::uint32_t part) {
    sheet_ = sheet;
    part_ = part;
    position_ = 0;
  }

  void add(std::optional<std::uint32_t> id, DxfUseKind kind) {
    if (id.has_value()) {
      out_.uses.push_back(DxfUse{*id, kind, sheet_, part_, position_++});
    }
  }

  void xml(std::string_view fragment, DxfUseKind fallback) {
    if (!fragment.empty()) {
      record(visit_dxf_xml_ids(fragment), fallback);
    }
  }

  void attrs(std::string_view list, DxfUseKind fallback) {
    if (!list.empty()) {
      record(visit_dxf_attr_ids(list), fallback);
    }
  }

  void block() { out_.blocked = true; }

  DxfUseSet take() { return std::move(out_); }

 private:
  void record(const Expected<std::vector<DxfXmlId>, Error>& ids, DxfUseKind fallback) {
    if (!ids) {
      block();
      return;
    }
    for (const DxfXmlId& id : ids.value()) {
      add(id.value, KindOf(id.element, fallback));
    }
  }

  DxfUseSet out_;
  std::int32_t sheet_ = -1;
  std::uint32_t part_ = 0;
  std::uint32_t position_ = 0;
};

/// The uses `serialize_auto_filter` writes for `filter`.
void CollectAutoFilter(Collector& c, const AutoFilter& filter) {
  if (filter.is_opaque()) {
    c.xml(filter.opaque_xml, DxfUseKind::kColorFilter);
    return;
  }
  for (const FilterColumn& col : filter.columns) {
    if (col.kind == FilterKind::kColor) {
      c.add(col.color.dxf_id, DxfUseKind::kColorFilter);
    }
    c.xml(col.extra_xml, DxfUseKind::kColorFilter);
    c.xml(col.ext_xml, DxfUseKind::kColorFilter);
  }
  if (filter.sort.has_value()) {
    for (const SortCondition& cond : filter.sort->conditions) {
      c.add(cond.dxf_id, DxfUseKind::kSortCondition);
    }
    c.xml(filter.sort->ext_lst_xml, DxfUseKind::kSortCondition);
  }
  c.xml(filter.ext_lst_xml, DxfUseKind::kColorFilter);
}

/// The uses of every CF rule. An x14-only rule's dxf is written inline, yet
/// Excel keeps it in `<dxfs>` too (structural_ref_m6b.xlsx), so it is a user.
void CollectConditionalFormats(Collector& c, const Sheet& sheet, const std::vector<std::string>& sheet_names) {
  for (const cf::ConditionalFormat& block : sheet.conditional_formats()) {
    bool any_legacy = block.rules.empty();
    for (const cf::CFRule& rule : block.rules) {
      c.add(rule.dxf_id, DxfUseKind::kConditionalFormat);
      if (cf_rule_is_x14_only(rule, sheet.name(), &sheet_names)) {
        continue;
      }
      any_legacy = true;
      if (rule.ext_lst_raw.has_value()) {
        c.xml(*rule.ext_lst_raw, DxfUseKind::kConditionalFormat);
      }
    }
    if (any_legacy && block.ext_lst_raw.has_value()) {
      c.xml(*block.ext_lst_raw, DxfUseKind::kConditionalFormat);
    }
  }
}

/// The uses `BuildWorksheetXml` writes, in its document order.
void CollectWorksheet(Collector& c, const Sheet& sheet, const std::vector<std::string>& sheet_names) {
  const WorksheetRawExtensions& raw = sheet.raw_extensions();
  std::size_t cursor = 0;
  const auto raw_before = [&](std::string_view element) {
    const std::size_t slot = worksheet_child::slot_of(element);
    for (; cursor < raw.size() && raw[cursor].slot < slot; ++cursor) {
      c.xml(raw[cursor].xml, DxfUseKind::kSortCondition);
    }
  };
  raw_before("autoFilter");
  if (sheet.has_auto_filter()) {
    CollectAutoFilter(c, *sheet.auto_filter());
  }
  raw_before("conditionalFormatting");
  CollectConditionalFormats(c, sheet, sheet_names);
  raw_before("extLst");
  c.xml(sheet.ext_lst_xml(), DxfUseKind::kConditionalFormat);
  for (; cursor < raw.size(); ++cursor) {
    c.xml(raw[cursor].xml, DxfUseKind::kSortCondition);
  }
}

/// The uses `BuildTableXml` writes for `table`.
void CollectTable(Collector& c, const TableMetadata& table) {
  c.attrs(table.root_extra_attrs, DxfUseKind::kTable);
  if (const AutoFilter* filter = table.auto_filter_xml.get(); filter != nullptr) {
    CollectAutoFilter(c, *filter);
  }
  c.xml(table.sort_state_xml, DxfUseKind::kSortCondition);
  for (const TableColumn& col : table.columns) {
    c.attrs(col.extra_attrs, DxfUseKind::kTableColumn);
    c.xml(col.trailing_children_xml, DxfUseKind::kTableColumn);
  }
  c.xml(table.ext_lst_xml, DxfUseKind::kTable);
}

using PivotRawField = const std::string& (pivot::PivotTable::*)() const;

/// The raw fields `write_pivot_table_definition` re-emits, in document order.
constexpr PivotRawField kPivotRawFields[] = {
    &pivot::PivotTable::raw_passthrough_after_row_fields,
    &pivot::PivotTable::raw_col_items,
    &pivot::PivotTable::raw_passthrough_before_style_info,
    &pivot::PivotTable::raw_passthrough_xml,
};

/// True when a slicer or timeline style element in the styles `<extLst>`
/// names a `dxfId` while no extension dxf collection exists for it to index.
bool ExtensionStylesMayIndexMainTable(std::string_view ext) {
  if (ext.find("<x14:dxfs") != std::string_view::npos || ext.find("<x15:dxfs") != std::string_view::npos) {
    return false;
  }
  for (const std::string_view element :
       {std::string_view("slicerStyleElement"), std::string_view("timelineStyleElement")}) {
    for (std::size_t at = ext.find(element); at != std::string_view::npos; at = ext.find(element, at + 1U)) {
      if (at == 0U || (ext[at - 1U] != '<' && ext[at - 1U] != ':')) {
        continue;
      }
      const std::size_t tag_end = ext.find('>', at);
      const std::string_view tag = ext.substr(at, tag_end == std::string_view::npos ? tag_end : tag_end - at);
      for (std::size_t attr = tag.find("dxfId"); attr != std::string_view::npos; attr = tag.find("dxfId", attr + 1U)) {
        const char before = tag[attr - 1U];
        if (before == ' ' || before == '\t' || before == '\r' || before == '\n') {
          return true;
        }
      }
    }
  }
  return false;
}

}  // namespace

Expected<DxfUseSet, Error> collect_xlsx_dxf_uses(const Workbook& wb, const EmissionPlan& plan) {
  Collector c;
  const std::vector<std::string> sheet_names = sheet_names_of(wb);
  const PassthroughPartIndex parts(wb.passthrough_parts());
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    const Sheet& sheet = wb.sheet(si);
    const auto sheet_index = static_cast<std::int32_t>(si);
    // An opaque sheet is written verbatim and cannot be renumbered.
    if (sheet.is_opaque_ooxml_sheet()) {
      if (const PassthroughPart* part = parts.find(sheet.opaque_ooxml_part_path()); part != nullptr) {
        const auto ids =
            visit_dxf_xml_ids(std::string_view(reinterpret_cast<const char*>(part->bytes.data()), part->bytes.size()));
        if (!ids || !ids.value().empty()) {
          c.block();
        }
      }
    } else {
      c.at(sheet_index, 0U);
      CollectWorksheet(c, sheet, sheet_names);
    }
    for (std::size_t ti = 0; ti < plan.tables_by_sheet[si].size(); ++ti) {
      c.at(sheet_index, static_cast<std::uint32_t>(ti));
      CollectTable(c, *plan.tables_by_sheet[si][ti].table);
    }
    for (std::size_t ti = 0; ti < plan.pivot_tables_by_sheet[si].size(); ++ti) {
      c.at(sheet_index, static_cast<std::uint32_t>(ti));
      const pivot::PivotTable& table = *plan.pivot_tables_by_sheet[si][ti].table;
      for (const PivotRawField field : kPivotRawFields) {
        c.xml((table.*field)(), DxfUseKind::kPivotFormat);
      }
    }
  }
  const StylesTable& styles = wb.styles();
  c.at(-1, 0U);
  for (const CustomTableStyle& style : styles.table_styles.styles) {
    for (const TableStyleElement& element : style.elements) {
      c.add(element.dxf_id, DxfUseKind::kTableStyleElement);
    }
  }
  for (const std::string& raw : styles.unknown_top_level_xml) {
    c.xml(raw, DxfUseKind::kTableStyleElement);
  }
  if (ExtensionStylesMayIndexMainTable(styles.ext_lst_xml)) {
    c.block();
  }
  return c.take();
}

std::optional<std::uint32_t> saved_dxf_id(std::optional<std::uint32_t> model_id, const DxfSavePlan* plan) {
  if (plan == nullptr || !model_id.has_value()) {
    return model_id;
  }
  return plan->map(*model_id);
}

std::string saved_dxf_xml(std::string_view xml, const DxfSavePlan* plan) {
  if (plan == nullptr || xml.empty()) {
    return std::string(xml);
  }
  auto rewritten = rewrite_dxf_xml_ids(xml, *plan);
  return rewritten ? std::move(rewritten.value()) : std::string(xml);
}

std::string saved_dxf_attrs(std::string_view attrs, const DxfSavePlan* plan) {
  if (plan == nullptr || attrs.empty()) {
    return std::string(attrs);
  }
  auto rewritten = rewrite_dxf_attr_ids(attrs, *plan);
  return rewritten ? std::move(rewritten.value()) : std::string(attrs);
}

}  // namespace formulon::io
