#include "io/xlsb/pivot_part_plan.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/pivot_projection.h"
#include "io/xlsb/pivot_cache_bridge.h"
#include "io/xlsb/pivot_cache_ids.h"
#include "io/xlsb/pivot_table_bridge.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/index_sort.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// A table of the group: its sheet, its index on the sheet, and the table.
struct Member {
  std::size_t sheet = 0;
  std::size_t index = 0;
  const pivot::PivotTable* table = nullptr;
};

std::vector<Member> GroupOf(const Workbook& wb, std::uint32_t cache_id) {
  std::vector<Member> out;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const auto& tables = wb.sheet(s).pivot_tables();
    for (std::size_t i = 0; i < tables.size(); ++i) {
      if (tables[i] != nullptr && tables[i]->pivot_cache_id() == cache_id) {
        out.push_back(Member{s, i, tables[i].get()});
      }
    }
  }
  return out;
}

bool Fresh(const Workbook& wb, const PassthroughPart& part) {
  return static_cast<bool>(CheckRetainedPartFreshness(wb, part));
}

/// The group's retained parts when every one is still as loaded and no table
/// was edited: the definition, the records, then one table part per member.
std::vector<const PassthroughPart*> KeptParts(const Workbook& wb, std::uint32_t cache_id,
                                              const std::vector<Member>& group) {
  const PassthroughPart* definition = nullptr;
  const PassthroughPart* records = nullptr;
  std::vector<const PassthroughPart*> tables(group.size(), nullptr);
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    switch (part.retained_origin) {
      case PassthroughPart::RetainedOrigin::kPivotCacheDefinition:
        definition = part.origin_cache_id == cache_id ? &part : definition;
        break;
      case PassthroughPart::RetainedOrigin::kPivotCacheRecords:
        records = part.origin_cache_id == cache_id ? &part : records;
        break;
      case PassthroughPart::RetainedOrigin::kPivotTable:
        for (std::size_t m = 0; m < group.size(); ++m) {
          if (part.origin_sheet_index == group[m].sheet && part.origin_pivot_index == group[m].index) {
            tables[m] = &part;
          }
        }
        break;
      default:
        break;
    }
  }
  if (definition == nullptr || records == nullptr || !Fresh(wb, *definition) || !Fresh(wb, *records)) {
    return {};
  }
  std::vector<const PassthroughPart*> out = {definition, records};
  for (std::size_t m = 0; m < group.size(); ++m) {
    if (tables[m] == nullptr || group[m].table->layout_dirty() || !Fresh(wb, *tables[m])) {
      return {};
    }
    out.push_back(tables[m]);
  }
  return out;
}

/// True when the group of `cache_id` has retained table parts but no retained
/// cache part (a part set built without the reader), every one still fresh:
/// its parts stay plain passthrough, as before the model could write pivots.
bool PassthroughOnlyGroup(const Workbook& wb, std::uint32_t cache_id, const std::vector<Member>& group) {
  std::size_t tables = 0;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if ((part.retained_origin == PassthroughPart::RetainedOrigin::kPivotCacheDefinition ||
         part.retained_origin == PassthroughPart::RetainedOrigin::kPivotCacheRecords) &&
        part.origin_cache_id == cache_id) {
      return false;
    }
    if (part.retained_origin != PassthroughPart::RetainedOrigin::kPivotTable) {
      continue;
    }
    for (const Member& member : group) {
      if (part.origin_sheet_index == member.sheet && part.origin_pivot_index == member.index) {
        if (member.table->layout_dirty() || !Fresh(wb, part)) {
          return false;
        }
        ++tables;
      }
    }
  }
  return tables != 0U && tables == group.size();
}

/// The ids of the retained PIVOTCACHEIDS entries whose cache is still a passthrough part.
std::unordered_set<std::uint32_t> PassthroughCacheIds(const Workbook& wb) {
  std::unordered_set<std::uint32_t> out;
  for (const BookPivotCacheId& entry : list_book_pivot_cache_ids(wb.xlsb_pivot_cache_ids()).entries) {
    for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
      if (rel.id == entry.rel_id && rel.type == kRelPivotCacheDefinition) {
        out.insert(entry.id);
      }
    }
  }
  return out;
}

/// The smallest `k` from `first` whose every path `paths(k)` names is free.
template <typename Paths>
std::size_t FreePartNumber(std::size_t first, const std::unordered_set<std::string>& taken, const Paths& paths) {
  for (std::size_t k = first;; ++k) {
    bool free = true;
    for (const std::string& path : paths(k)) {
      free = free && taken.count(path) == 0U;
    }
    if (free) {
      return k;
    }
  }
}

}  // namespace

bool is_retained_pivot_part(const Workbook& wb, const PassthroughPart& part) {
  if (part.retained_origin == PassthroughPart::RetainedOrigin::kPivotCacheDefinition ||
      part.retained_origin == PassthroughPart::RetainedOrigin::kPivotCacheRecords) {
    return true;
  }
  if (part.retained_origin != PassthroughPart::RetainedOrigin::kPivotTable) {
    return false;
  }
  if (part.origin_sheet_index >= wb.sheet_count()) {
    return true;
  }
  const auto& tables = wb.sheet(part.origin_sheet_index).pivot_tables();
  if (part.origin_pivot_index >= tables.size() || tables[part.origin_pivot_index] == nullptr) {
    return true;
  }
  const std::uint32_t cache_id = tables[part.origin_pivot_index]->pivot_cache_id();
  return !PassthroughOnlyGroup(wb, cache_id, GroupOf(wb, cache_id));
}

bool pivot_table_kept_verbatim(const Workbook& wb, std::size_t sheet, std::size_t pivot) {
  const auto& tables = wb.sheet(sheet).pivot_tables();
  if (pivot >= tables.size() || tables[pivot] == nullptr) {
    return false;
  }
  const std::uint32_t cache_id = tables[pivot]->pivot_cache_id();
  const std::vector<Member> group = GroupOf(wb, cache_id);
  return PassthroughOnlyGroup(wb, cache_id, group) ||
         (wb.find_pivot_cache(cache_id) != nullptr && !KeptParts(wb, cache_id, group).empty());
}

PivotPartLoss plan_pivot_parts(const Workbook& wb, const std::unordered_set<std::string>& taken, EmissionPlan& plan) {
  plan.sheet_pivots.assign(wb.sheet_count(), {});
  PivotPartLoss loss;
  loss.tables.assign(wb.sheet_count(), 0U);
  std::vector<std::size_t>& lost = loss.tables;
  // Tables per sheet, with the creation order Excel numbers generated parts by.
  struct Planned {
    std::uint64_t creation_order = 0;
    EmissionPlan::PivotTablePart part;
  };
  std::vector<std::vector<Planned>> planned(wb.sheet_count());
  const std::unordered_set<std::uint32_t> passthrough_ids = PassthroughCacheIds(wb);
  // A retained original no group reuses is not written, so its path is free for the group regenerated in its place.
  std::unordered_set<std::string> reused;
  for (const auto& cache : wb.pivot_caches()) {
    if (cache != nullptr && passthrough_ids.count(cache->cache_id()) == 0U) {
      for (const PassthroughPart* part : KeptParts(wb, cache->cache_id(), GroupOf(wb, cache->cache_id()))) {
        reused.insert(part->path);
      }
    }
  }
  std::unordered_set<std::string> occupied = taken;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (is_retained_pivot_part(wb, part) && reused.count(part.path) == 0U) {
      occupied.erase(part.path);
    }
  }
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& table : wb.sheet(s).pivot_tables()) {
      lost[s] += table != nullptr && wb.find_pivot_cache(table->pivot_cache_id()) == nullptr ? 1U : 0U;
    }
  }
  std::size_t next_cache = 1;
  for (const auto& cache : wb.pivot_caches()) {
    if (cache == nullptr) {
      continue;
    }
    const std::vector<Member> group = GroupOf(wb, cache->cache_id());
    if (PassthroughOnlyGroup(wb, cache->cache_id(), group)) {
      continue;
    }
    EmissionPlan::PivotCachePart part;
    part.cache_id = cache->cache_id();
    std::vector<std::vector<std::uint8_t>> tables;
    std::vector<std::string> table_paths;
    // A cache id an unmodelled cache of the package already holds would name two caches.
    const bool id_free = passthrough_ids.count(cache->cache_id()) == 0U;
    const std::vector<const PassthroughPart*> kept =
        id_free ? KeptParts(wb, cache->cache_id(), group) : std::vector<const PassthroughPart*>{};
    if (!kept.empty()) {
      part.definition_path = kept[0]->path;
      part.records_path = kept[1]->path;
      part.definition = kept[0]->bytes;
      part.records = kept[1]->bytes;
      for (std::size_t m = 0; m < group.size(); ++m) {
        tables.push_back(kept[2U + m]->bytes);
        table_paths.push_back(kept[2U + m]->path);
      }
    } else {
      auto parts = id_free ? pivot_cache_to_xlsb(*cache, wb.date1904(), kPivotCacheRecordsRelId)
                           : Expected<XlsbPivotCacheParts, Error>(not_writable("pivot_cache", 0));
      for (std::size_t m = 0; parts && m < group.size(); ++m) {
        // Each field's items are saved as modelled, so the lines address them without a write plan.
        ProjectedPivot projection;
        {
          const eval::ScopedEvalProfile profile_scope(wb.excel_profile(), wb.text_index_unit());
          projection = project_pivot(wb, *group[m].table, nullptr);
        }
        auto bytes = pivot_table_to_xlsb(*group[m].table, wb.excel_profile().host, group[m].sheet, projection.span,
                                         projection.items ? &*projection.items : nullptr);
        if (!bytes) {
          StructuredLog("xlsb.writer.pivot_deferred").field("reason", bytes.error().message).warn();
          break;
        }
        tables.push_back(std::move(bytes.value()));
      }
      if (!parts || tables.size() != group.size()) {
        if (!parts) {
          StructuredLog("xlsb.writer.pivot_deferred").field("reason", parts.error().message).warn();
        }
        for (const Member& member : group) {
          ++lost[member.sheet];
        }
        loss.caches += group.empty() ? 1U : 0U;
        continue;
      }
      const std::size_t k = FreePartNumber(next_cache, occupied, [](std::size_t n) {
        const std::string id = std::to_string(n);
        return std::array<std::string, 3>{"xl/pivotCache/pivotCacheDefinition" + id + ".bin",
                                          "xl/pivotCache/_rels/pivotCacheDefinition" + id + ".bin.rels",
                                          "xl/pivotCache/pivotCacheRecords" + id + ".bin"};
      });
      next_cache = k + 1U;
      part.definition_path = "xl/pivotCache/pivotCacheDefinition" + std::to_string(k) + ".bin";
      // A cache saved without its data has neither a records part nor the relationship naming it.
      if (cache->save_data()) {
        part.records_path = "xl/pivotCache/pivotCacheRecords" + std::to_string(k) + ".bin";
      }
      part.definition = std::move(parts.value().definition);
      part.records = std::move(parts.value().records);
    }
    for (std::size_t m = 0; m < group.size(); ++m) {
      Planned entry;
      entry.creation_order = group[m].table->creation_order();
      entry.part.cache = plan.pivot_caches.size();
      entry.part.index = group[m].index;
      entry.part.verbatim = !kept.empty();
      entry.part.path = kept.empty() ? std::string() : table_paths[m];
      entry.part.bytes = std::move(tables[m]);
      planned[group[m].sheet].push_back(std::move(entry));
    }
    plan.pivot_caches.push_back(std::move(part));
  }
  // Excel numbers a sheet's pivots newest first, with one counter across the package.
  std::size_t next_table = 1;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    sort_by_index(planned[s], [](const Planned& a, const Planned& b) { return a.creation_order > b.creation_order; });
    for (Planned& entry : planned[s]) {
      if (entry.part.path.empty()) {
        const std::size_t k = FreePartNumber(next_table, occupied, [](std::size_t n) {
          const std::string id = std::to_string(n);
          return std::array<std::string, 2>{"xl/pivotTables/pivotTable" + id + ".bin",
                                            "xl/pivotTables/_rels/pivotTable" + id + ".bin.rels"};
        });
        next_table = k + 1U;
        entry.part.path = "xl/pivotTables/pivotTable" + std::to_string(k) + ".bin";
      }
      plan.sheet_pivots[s].push_back(std::move(entry.part));
    }
  }
  return loss;
}

std::vector<std::uint8_t> book_pivot_cache_ids(const std::vector<std::uint8_t>& retained, const EmissionPlan& plan) {
  if (plan.pivot_caches.empty()) {
    return retained;
  }
  std::vector<std::uint8_t> out;
  ByteSpan cursor{retained.data(), retained.size()};
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      out.clear();
      break;
    }
    if (rec.value().type != kBrtEndPivotCacheIds) {
      emit_record(out, rec.value().type, rec.value().payload);
    }
  }
  if (out.empty()) {
    emit_record(out, kBrtBeginPivotCacheIds, ByteSpan{});
  }
  for (std::size_t k = 0; k < plan.pivot_caches.size(); ++k) {
    std::vector<std::uint8_t> payload;
    emit_u32(payload, plan.pivot_caches[k].cache_id);
    emit_xlwidestring(payload, pivot_cache_rel_key(k));
    emit_record(out, kBrtBeginPivotCacheId, payload);
    emit_record(out, kBrtEndPivotCacheId, ByteSpan{});
  }
  emit_record(out, kBrtEndPivotCacheIds, ByteSpan{});
  return out;
}

std::string relative_part_target(std::string_view from, std::string_view to) {
  const std::size_t from_dir = from.rfind('/') == std::string_view::npos ? 0U : from.rfind('/') + 1U;
  // The longest common directory prefix, ending at a '/'.
  std::size_t common = 0;
  for (std::size_t i = 0; i < from_dir && i < to.size() && from[i] == to[i]; ++i) {
    if (from[i] == '/') {
      common = i + 1U;
    }
  }
  std::string out;
  for (std::size_t i = common; i < from_dir; ++i) {
    if (from[i] == '/') {
      out.append("../");
    }
  }
  out.append(to.substr(common));
  return out;
}

std::string single_relationship_part(std::string_view type, std::string_view target) {
  std::string out(kXmlDecl);
  BeginRelationships(out);
  AppendRelationship(out, 1U, type, target);
  EndRelationships(out);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
