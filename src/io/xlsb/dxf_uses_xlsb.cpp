#include "io/xlsb/dxf_uses_xlsb.h"

#include <cstddef>
#include <optional>
#include <utility>
#include <vector>

#include "io/xlsb/dxf_props.h"
#include "io/xlsb/dxf_record_slots.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/styles_writer.h"

namespace formulon::io::xlsb {
namespace {

/// The holder kind of a dxf slot by its record type.
DxfUseKind KindOf(std::uint16_t type) {
  switch (type) {
    case kBrtBeginList:
      return DxfUseKind::kTable;
    case kBrtBeginListCol:
      return DxfUseKind::kTableColumn;
    case kBrtColorFilter:
      return DxfUseKind::kColorFilter;
    case kBrtBeginSortCond:
      return DxfUseKind::kSortCondition;
    case kBrtTableStyleElement:
      return DxfUseKind::kTableStyleElement;
    default:
      return DxfUseKind::kConditionalFormat;
  }
}

/// The records of `stream` that can carry a dxf slot.
std::vector<std::uint8_t> SlotRange(const XlsbDxfStream& stream) {
  return std::vector<std::uint8_t>(stream.bytes->begin() + static_cast<std::ptrdiff_t>(stream.begin),
                                   stream.bytes->end());
}

/// Adds the uses of `bytes`; false when its slots cannot be located.
bool AddUses(const std::vector<std::uint8_t>& bytes, const DxfPartKey& part, DxfUseSet& out) {
  auto scan = scan_dxf_slots(bytes);
  if (!scan) {
    return false;
  }
  if (!scan.value().undecodable.empty()) {
    out.blocked = true;
  }
  std::uint32_t position = 0;
  for (const DxfRecordSlot& slot : scan.value().slots) {
    DxfUse use;
    use.model_index = slot.value;
    use.kind = KindOf(slot.record_type);
    use.part = part;
    use.position = position++;
    if (slot.priority.has_value()) {
      use.priority = *slot.priority;
    }
    use.block = slot.block;
    out.uses.push_back(use);
  }
  return true;
}

/// Maps the reference slots of `stream` through `plan` in place.
Expected<void, Error> MapSlots(const XlsbDxfStream& stream, const DxfSavePlan& plan) {
  const std::vector<std::uint8_t> range = SlotRange(stream);
  auto scan = scan_dxf_slots(range);
  if (!scan) {
    return std::move(scan.error());
  }
  std::uint8_t* base = stream.bytes->data() + stream.begin;
  DxfPartCursor cursor(plan, stream.part);
  for (const DxfRecordSlot& slot : scan.value().slots) {
    store_u32(base + slot.offset, cursor.next(slot.value).value_or(kNoDxf));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

DxfUseSet collect_xlsb_dxf_uses(const Workbook& wb, const XlsbEmission& emission) {
  DxfUseSet out;
  // No reader decodes a pivot table part's dxf records.
  out.blocked = emission.retained_pivot;
  for (const XlsbDxfStream& stream : emission.streams) {
    if (!AddUses(SlotRange(stream), stream.part, out)) {
      out.blocked = true;
    }
  }
  if (emission.raw_styles == nullptr) {
    // A generated styles part names a dxf from each written table style element, registered in model order.
    std::uint32_t position = 0;
    for (const CustomTableStyle& style : wb.styles().table_styles.styles) {
      if (!table_style_is_written(style)) {
        continue;
      }
      for (const TableStyleElement& element : style.elements) {
        DxfUse use;
        use.model_index = *element.dxf_id;
        use.kind = DxfUseKind::kTableStyleElement;
        use.part = DxfPartKey{-1, DxfPartType::kStyles, 0U};
        use.position = position++;
        out.uses.push_back(use);
      }
    }
  }
  if (emission.raw_styles != nullptr) {
    // The raw list is rewritten by model index, so it must be the model's.
    const std::optional<std::size_t> raw_count = styles_bin_dxf_count(*emission.raw_styles);
    if (!raw_count.has_value() || *raw_count != wb.styles().dxfs.size() ||
        !AddUses(*emission.raw_styles, DxfPartKey{-1, DxfPartType::kStyles, 0U}, out)) {
      out.blocked = true;
    }
  }
  return out;
}

Expected<void, Error> apply_xlsb_dxf_plan(XlsbEmission& emission, const DxfSavePlan& plan) {
  for (const XlsbDxfStream& stream : emission.streams) {
    if (auto mapped = MapSlots(stream, plan); !mapped) {
      return mapped;
    }
  }
  if (emission.raw_styles != nullptr) {
    auto compacted = compact_styles_bin(*emission.raw_styles, plan);
    if (!compacted) {
      return std::move(compacted.error());
    }
    *emission.raw_styles = std::move(compacted.value());
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon::io::xlsb
