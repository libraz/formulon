#include "io/xlsb/metadata_bin.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Record ids of the metadata part ([MS-XLSB] 2.4.x). They live in their own
// dispatch space rather than in `XlsbRecordType`, which enumerates the
// worksheet / workbook / styles records the reader consumes.
constexpr std::uint16_t kBrtBeginMetadata = 332;
constexpr std::uint16_t kBrtEndMetadata = 333;
constexpr std::uint16_t kBrtBeginEsmdtinfo = 334;  // Opens the metadata-type table.
constexpr std::uint16_t kBrtMdtinfo = 335;         // One metadata type.
constexpr std::uint16_t kBrtEndEsmdtinfo = 336;
constexpr std::uint16_t kBrtBeginEsfmd = 337;  // Opens the cell-metadata table.
constexpr std::uint16_t kBrtEndEsfmd = 338;
constexpr std::uint16_t kBrtBeginEsfmdInfo = 339;
constexpr std::uint16_t kBrtEndEsfmdInfo = 340;
constexpr std::uint16_t kBrtMdb = 51;  // One cell-metadata entry.
constexpr std::uint16_t kBrtBeginExt = 52;
constexpr std::uint16_t kBrtEndExt = 53;
constexpr std::uint16_t kBrtBeginDynamicArrayExt = 4096;
constexpr std::uint16_t kBrtDynamicArrayProperties = 4097;
constexpr std::uint16_t kBrtBeginRichValueExt = 5002;
constexpr std::uint16_t kBrtRichValueBlock = 5003;

// Flag words of the two type records, and the second word both share.
constexpr std::uint32_t kDynamicArrayTypeFlags = 0xD86AC0B0U;
constexpr std::uint32_t kRichValueTypeFlags = 0x986AC0B0U;
constexpr std::uint32_t kTypeFlagsHigh = 0x0001D4C0U;

// The first word of a future-metadata block's `BrtFrtBegin` payload names the
// feature version; it differs between the two types.
constexpr std::uint32_t kDynamicArrayFrtVersion = 0x00020002U;
constexpr std::uint32_t kRichValueFrtVersion = 0x00020001U;

// The `BrtBeginEsfmd` second word: 1 opens the cell-metadata list, 0 the
// value-metadata list.
constexpr std::uint32_t kCellMetadataList = 1U;
constexpr std::uint32_t kValueMetadataList = 0U;

/// What one validating walk of a metadata part learned.
struct MetadataScan {
  bool framing_complete = false;   ///< Both tables and the part itself closed.
  std::uint32_t type_ordinal = 0;  ///< 1-based ordinal of the XLDAPR type.
  std::uint32_t entry_index = 0;   ///< 1-based index of the entry naming it.
};

/// Walks the whole part once, collecting the XLDAPR type ordinal and the
/// index of the first cell-metadata entry that names it.
///
/// The walk runs to the end of the buffer even after both are known, because
/// the index is only meaningful if the part it indexes is intact: a stream cut
/// short after the matching entry still has an unterminated table, and naming
/// an entry inside it would put a `BrtCellMeta` index into a worksheet whose
/// metadata part Excel rejects. `framing_complete` is therefore the gate, and
/// it requires the part to open, both tables to open and close in order, and
/// the part to close, with every record header and payload inside the buffer.
MetadataScan ScanMetadataPart(ByteSpan bytes) {
  MetadataScan scan;
  ByteSpan cursor = bytes;
  bool part_open = false;
  bool part_closed = false;
  bool in_type_table = false;
  bool type_table_closed = false;
  bool in_cell_table = false;
  bool cell_table_closed = false;
  std::uint32_t type_ordinal = 0;
  std::uint32_t entry_index = 0;

  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return MetadataScan{};  // Truncated header or payload: trust nothing.
    }
    const XlsbRecord& record = record_or.value();
    switch (record.type) {
      case kBrtBeginMetadata:
        part_open = true;
        continue;
      case kBrtEndMetadata:
        part_closed = true;
        continue;
      case kBrtBeginEsmdtinfo:
        in_type_table = true;
        continue;
      case kBrtEndEsmdtinfo:
        in_type_table = false;
        type_table_closed = true;
        continue;
      case kBrtBeginEsfmd:
        // The cell table indexes ordinals the type table assigns, so a part
        // that opens it first is not one this can resolve against.
        if (!type_table_closed) {
          return MetadataScan{};
        }
        in_cell_table = true;
        continue;
      case kBrtEndEsfmd:
        in_cell_table = false;
        cell_table_closed = true;
        continue;
      default:
        break;
    }

    if (in_type_table && record.type == kBrtMdtinfo) {
      // Ordinals count every declared type, including ones whose payload is
      // shaped differently from what Excel writes: a later entry's index only
      // means anything if the earlier ones were counted.
      ++type_ordinal;
      // MDTINFO: two flag words, then the type name.
      ByteSpan payload = record.payload;
      const auto flags_low = read_u32(payload);
      const auto flags_high = read_u32(payload);
      if (!flags_low || !flags_high) {
        return MetadataScan{};
      }
      const auto name_or = read_xlwidestring(payload);
      if (!name_or) {
        return MetadataScan{};
      }
      if (scan.type_ordinal == 0U && name_or.value() == kDynamicArrayMetadataType) {
        scan.type_ordinal = type_ordinal;
      }
      continue;
    }

    if (in_cell_table && record.type == kBrtMdb) {
      // MDB: a count of (type ordinal, id) pairs, then the pairs. An entry may
      // carry several types; it names the dynamic-array one if any pair does.
      ++entry_index;
      ByteSpan payload = record.payload;
      const auto pair_count_or = read_u32(payload);
      if (!pair_count_or) {
        return MetadataScan{};
      }
      for (std::uint32_t i = 0; i < pair_count_or.value(); ++i) {
        const auto pair_type_or = read_u32(payload);
        const auto pair_id_or = read_u32(payload);
        if (!pair_type_or || !pair_id_or) {
          return MetadataScan{};  // A count longer than the payload; trust neither.
        }
        if (scan.entry_index == 0U && scan.type_ordinal != 0U && pair_type_or.value() == scan.type_ordinal) {
          scan.entry_index = entry_index;
        }
      }
    }
  }

  scan.framing_complete =
      part_open && part_closed && type_table_closed && cell_table_closed && !in_type_table && !in_cell_table;
  return scan;
}

}  // namespace

std::vector<std::uint8_t> build_metadata_bin(bool dynamic_arrays, std::size_t rich_values) {
  const MetadataModel model = metadata_model(dynamic_arrays, rich_values);
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginMetadata, ByteSpan{});

  emit_u32(payload, model.type_count);
  emit_record(out, kBrtBeginEsmdtinfo, payload);
  for (std::uint32_t t = 1; t <= model.type_count; ++t) {
    const bool dynamic = t == model.dynamic_type;
    payload.clear();
    emit_u32(payload, dynamic ? kDynamicArrayTypeFlags : kRichValueTypeFlags);
    emit_u32(payload, kTypeFlagsHigh);
    emit_xlwidestring(payload, dynamic ? kDynamicArrayMetadataType : kRichValueMetadataType);
    emit_record(out, kBrtMdtinfo, payload);
  }
  emit_record(out, kBrtEndEsmdtinfo, ByteSpan{});

  const auto begin_future = [&](std::size_t blocks, std::string_view name) {
    payload.clear();
    emit_u32(payload, static_cast<std::uint32_t>(blocks));
    emit_xlwidestring(payload, name);
    emit_record(out, kBrtBeginEsfmdInfo, payload);
  };
  const auto begin_block = [&](std::uint32_t frt_version) {
    emit_record(out, kBrtBeginExt, ByteSpan{});
    payload.clear();
    emit_u32(payload, frt_version);
    emit_record(out, kFrtBegin, payload);
  };
  const auto end_block = [&]() {
    emit_record(out, kFrtEnd, ByteSpan{});
    emit_record(out, kBrtEndExt, ByteSpan{});
  };
  if (model.dynamic_blocks != 0U) {
    begin_future(model.dynamic_blocks, kDynamicArrayMetadataType);
    for (std::size_t i = 0; i < model.dynamic_blocks; ++i) {
      begin_block(kDynamicArrayFrtVersion);
      emit_record(out, kBrtBeginDynamicArrayExt, ByteSpan{});
      payload.clear();
      emit_u16(payload, 1U);
      emit_record(out, kBrtDynamicArrayProperties, payload);
      end_block();
    }
    emit_record(out, kBrtEndEsfmdInfo, ByteSpan{});
  }
  if (!model.rich_blocks.empty()) {
    begin_future(model.rich_blocks.size(), kRichValueMetadataType);
    for (const std::uint32_t rv : model.rich_blocks) {
      begin_block(kRichValueFrtVersion);
      emit_record(out, kBrtBeginRichValueExt, ByteSpan{});
      payload.clear();
      emit_u32(payload, 0U);
      emit_u32(payload, rv);
      emit_record(out, kBrtRichValueBlock, payload);
      end_block();
    }
    emit_record(out, kBrtEndEsfmdInfo, ByteSpan{});
  }

  const auto emit_list = [&](const std::vector<MetadataEntry>& entries, std::uint32_t list) {
    if (entries.empty()) {
      return;
    }
    payload.clear();
    emit_u32(payload, static_cast<std::uint32_t>(entries.size()));
    emit_u32(payload, list);
    emit_record(out, kBrtBeginEsfmd, payload);
    for (const MetadataEntry& e : entries) {
      payload.clear();
      emit_u32(payload, 1U);  // One (type, id) pair follows.
      emit_u32(payload, e.type);
      emit_u32(payload, e.id);
      emit_record(out, kBrtMdb, payload);
    }
    emit_record(out, kBrtEndEsfmd, ByteSpan{});
  };
  emit_list(model.cell_entries, kCellMetadataList);
  emit_list(model.value_entries, kValueMetadataList);
  emit_record(out, kBrtEndMetadata, ByteSpan{});
  return out;
}

std::vector<std::uint8_t> build_dynamic_array_metadata_bin() {
  return build_metadata_bin(true, 0U);
}

RichValueMetadata parse_metadata_bin(ByteSpan bytes) {
  std::vector<XlsbRecord> records;
  for (ByteSpan cursor = bytes; cursor.size > 0U;) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    records.push_back(record_or.value());
  }
  std::size_t at = 0;
  const auto take = [&](std::uint16_t type) -> const XlsbRecord* {
    return at < records.size() && records[at].type == type ? &records[at++] : nullptr;
  };
  const auto next_is = [&](std::uint16_t type) { return at < records.size() && records[at].type == type; };

  if (take(kBrtBeginMetadata) == nullptr) {
    return {};
  }
  const XlsbRecord* types = take(kBrtBeginEsmdtinfo);
  if (types == nullptr) {
    return {};
  }
  ByteSpan payload = types->payload;
  const auto type_count = read_u32(payload);
  if (!type_count) {
    return {};
  }
  MetadataModel model;
  for (std::uint32_t i = 1; i <= type_count.value(); ++i) {
    const XlsbRecord* type = take(kBrtMdtinfo);
    if (type == nullptr) {
      return {};
    }
    payload = type->payload;
    const auto low = read_u32(payload);
    const auto high = read_u32(payload);
    const auto name = read_xlwidestring(payload);
    if (!low || !high || !name || !add_metadata_type(model, name.value())) {
      return {};
    }
  }
  if (take(kBrtEndEsmdtinfo) == nullptr) {
    return {};
  }

  bool rich_blocks_seen = false;
  while (next_is(kBrtBeginEsfmdInfo)) {
    payload = take(kBrtBeginEsfmdInfo)->payload;
    const auto blocks = read_u32(payload);
    const auto name = read_xlwidestring(payload);
    if (!blocks || !name) {
      return {};
    }
    const bool dynamic = name.value() == kDynamicArrayMetadataType;
    if (!dynamic && name.value() != kRichValueMetadataType) {
      return {};
    }
    if (dynamic ? (model.dynamic_blocks != 0U || model.dynamic_type == 0U || blocks.value() != 1U)
                : (rich_blocks_seen || model.rich_type == 0U)) {
      return {};
    }
    rich_blocks_seen = rich_blocks_seen || !dynamic;
    for (std::uint32_t b = 0; b < blocks.value(); ++b) {
      if (take(kBrtBeginExt) == nullptr || take(kFrtBegin) == nullptr) {
        return {};
      }
      if (dynamic) {
        const XlsbRecord* properties =
            take(kBrtBeginDynamicArrayExt) != nullptr ? take(kBrtDynamicArrayProperties) : nullptr;
        if (properties == nullptr) {
          return {};
        }
        payload = properties->payload;
        const auto flags = read_u16(payload);
        if (!flags || flags.value() != 1U || payload.size != 0U) {
          return {};
        }
        ++model.dynamic_blocks;
      } else {
        const XlsbRecord* block = take(kBrtBeginRichValueExt) != nullptr ? take(kBrtRichValueBlock) : nullptr;
        if (block == nullptr) {
          return {};
        }
        payload = block->payload;
        const auto set = read_u32(payload);
        const auto index = read_u32(payload);
        if (!set || !index || set.value() != 0U || payload.size != 0U) {
          return {};
        }
        model.rich_blocks.push_back(index.value());
      }
      if (take(kFrtEnd) == nullptr || take(kBrtEndExt) == nullptr) {
        return {};
      }
    }
    if (take(kBrtEndEsfmdInfo) == nullptr) {
      return {};
    }
  }

  bool cell_list_seen = false;
  bool value_list_seen = false;
  while (next_is(kBrtBeginEsfmd)) {
    payload = take(kBrtBeginEsfmd)->payload;
    const auto entries = read_u32(payload);
    const auto list = read_u32(payload);
    if (!entries || !list) {
      return {};
    }
    const bool cell = list.value() == kCellMetadataList;
    if ((!cell && list.value() != kValueMetadataList) || (cell ? cell_list_seen : value_list_seen) ||
        (cell && (model.dynamic_type == 0U || entries.value() != 1U))) {
      return {};
    }
    (cell ? cell_list_seen : value_list_seen) = true;
    for (std::uint32_t e = 0; e < entries.value(); ++e) {
      const XlsbRecord* entry = take(kBrtMdb);
      if (entry == nullptr) {
        return {};
      }
      payload = entry->payload;
      const auto pairs = read_u32(payload);
      const auto type = read_u32(payload);
      const auto id = read_u32(payload);
      if (!pairs || !type || !id || pairs.value() != 1U || payload.size != 0U) {
        return {};
      }
      (cell ? model.cell_entries : model.value_entries).push_back({type.value(), id.value()});
    }
    if (take(kBrtEndEsfmd) == nullptr) {
      return {};
    }
  }
  if (take(kBrtEndMetadata) == nullptr || at != records.size()) {
    return {};
  }
  return resolve_metadata(model);
}

std::uint32_t find_dynamic_array_cell_meta_index(ByteSpan bytes) {
  const MetadataScan scan = ScanMetadataPart(bytes);
  if (!scan.framing_complete) {
    return 0;
  }
  return scan.entry_index;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
