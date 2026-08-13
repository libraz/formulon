#include "io/xlsb/metadata_bin.h"

#include <string>
#include <string_view>

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
constexpr std::uint16_t kBrtBeginFRT = 35;
constexpr std::uint16_t kBrtEndFRT = 36;
constexpr std::uint16_t kBrtBeginDynamicArrayExt = 4096;
constexpr std::uint16_t kBrtDynamicArrayProperties = 4097;

// Name Excel gives the dynamic-array metadata type.
constexpr std::string_view kDynamicArrayTypeName = "XLDAPR";

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
      if (scan.type_ordinal == 0U && name_or.value() == kDynamicArrayTypeName) {
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

std::vector<std::uint8_t> build_dynamic_array_metadata_bin() {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginMetadata, ByteSpan{});

  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsmdtinfo, payload);
  payload.clear();
  emit_u32(payload, 0xD86AC0B0U);
  emit_u32(payload, 0x0001D4C0U);
  emit_xlwidestring(payload, kDynamicArrayTypeName);
  emit_record(out, kBrtMdtinfo, payload);
  emit_record(out, kBrtEndEsmdtinfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, 1U);
  emit_xlwidestring(payload, kDynamicArrayTypeName);
  emit_record(out, kBrtBeginEsfmdInfo, payload);
  emit_record(out, kBrtBeginExt, ByteSpan{});
  payload.clear();
  emit_u32(payload, 0x00020002U);
  emit_record(out, kBrtBeginFRT, payload);
  emit_record(out, kBrtBeginDynamicArrayExt, ByteSpan{});
  payload.clear();
  emit_u16(payload, 1U);
  emit_record(out, kBrtDynamicArrayProperties, payload);
  emit_record(out, kBrtEndFRT, ByteSpan{});
  emit_record(out, kBrtEndExt, ByteSpan{});
  emit_record(out, kBrtEndEsfmdInfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, 1U);
  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsfmd, payload);
  payload.clear();
  emit_u32(payload, 1U);  // One (type, id) pair follows.
  emit_u32(payload, 1U);  // Type ordinal: the sole XLDAPR entry above.
  emit_u32(payload, 0U);
  emit_record(out, kBrtMdb, payload);
  emit_record(out, kBrtEndEsfmd, ByteSpan{});
  emit_record(out, kBrtEndMetadata, ByteSpan{});
  return out;
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
