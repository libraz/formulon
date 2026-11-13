/// @file
/// Record-level comparison of an `.xlsb` we write against Excel's own save of the same source.
///
/// A twin is `<case>.xlsb` plus `<case>.2.xlsb`, two Excel saves of one source. The bytes that differ between
/// them are fields Excel leaves uninitialised, so they form the mask a comparison ignores. Fixed uninitialised
/// fields that a single pair need not expose are added through `UninitialisedFieldHook`s; the default hook covers
/// the PtgAttrSemi data word of a formula whose token stream opens with it. PtgArray's reserved bytes and
/// PtgMemArea's u32 sit at arbitrary token positions and need an rgce walk, which no hook performs yet.
///
/// Records are matched per part by position among the selected types, and a mask position is keyed by the
/// record's index in Excel's full stream. The file also lists and converts the excel_open_gate bases, which the
/// guards share.

#pragma once

#include <cstddef>
#include <cstdint>
#include <functional>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <utility>
#include <vector>

namespace formulon::io::xlsb::test {

using Bytes = std::vector<std::uint8_t>;

/// The whole file at `path`; empty when it cannot be read.
Bytes ReadBytes(const std::string& path);

/// Entry names of the zip `package`, in archive order; empty when it does not open.
std::vector<std::string> ListParts(const Bytes& package);

/// The decompressed entry `name` of `package`, or nullopt.
std::optional<Bytes> ReadPart(const Bytes& package, const std::string& name);

/// Whether `name` is a `.bin` part holding a record stream (printer settings, embeddings, ActiveX and VBA
/// binaries are not).
bool IsRecordPart(const std::string& name);

/// One record of a part, with its index in the part's full record stream.
struct IndexedRecord {
  std::size_t index = 0;
  std::uint16_t type = 0;
  Bytes payload;
};

/// Every record of `part` in stream order; empty when the stream does not frame whole.
std::vector<IndexedRecord> SplitPart(const Bytes& part);

/// Record part name to its records whose type is in `types` (every type when `types` is empty), in order.
using RecordSet = std::map<std::string, std::vector<IndexedRecord>>;
RecordSet ExtractRecords(const Bytes& package, const std::set<std::uint16_t>& types);

/// Byte positions a comparison ignores, keyed by part, Excel record index and payload offset.
struct ByteMask {
  std::map<std::string, std::map<std::size_t, std::set<std::size_t>>> bytes;
  /// Records masked whole: their type or length differs between the two saves.
  std::set<std::pair<std::string, std::size_t>> whole_records;
  /// Parts masked whole: present in one save only, or not record streams and not byte-identical.
  std::set<std::string> whole_parts;

  bool Masks(const std::string& part, std::size_t index, std::size_t offset) const;
  std::size_t ByteCount() const;
  bool empty() const { return bytes.empty() && whole_records.empty() && whole_parts.empty(); }
  bool operator==(const ByteMask& other) const {
    return bytes == other.bytes && whole_records == other.whole_records && whole_parts == other.whole_parts;
  }
};

/// The bytes that differ between two Excel saves of one source.
ByteMask BuildTwinMask(const Bytes& first_save, const Bytes& second_save);

/// Appends to `offsets` the payload offsets of fields Excel leaves uninitialised in a record of `type`.
using UninitialisedFieldHook =
    std::function<void(std::uint16_t type, const Bytes& payload, std::vector<std::size_t>* offsets)>;

/// The PtgAttrSemi data word of a cell, array or shared formula whose rgce opens with PtgAttrSemi.
void MaskLeadingAttrSemi(std::uint16_t type, const Bytes& payload, std::vector<std::size_t>* offsets);

/// The hooks every twin comparison applies.
std::vector<UninitialisedFieldHook> DefaultUninitialisedHooks();

/// Adds to `mask` the positions `hooks` report for every record of `excel_package`.
void AddUninitialisedFields(const Bytes& excel_package, const std::vector<UninitialisedFieldHook>& hooks,
                            ByteMask* mask);

/// One readable line per difference between `ours` and `excel` over `types` (every type when empty) outside
/// `mask`: a missing part, a record count, type or length difference, or a run of differing bytes with both sides
/// in hex. Empty when they match.
std::vector<std::string> CompareRecordSets(const Bytes& ours, const Bytes& excel, const std::set<std::uint16_t>& types,
                                           const ByteMask& mask);

/// Paths relative to the fixtures root of the bases tools/oracle/excel_open_gate.txt converts; control lines are
/// skipped.
std::vector<std::string> GateBases(const std::string& gate_list_path);

/// `relative` (under `fixtures_root`) loaded as `.xlsx` or `.xlsb` and written as `.xlsb`.
struct Converted {
  Bytes bytes;
  std::string error;  ///< Empty on success.
};
Converted ConvertToXlsb(const std::string& fixtures_root, const std::string& relative);

}  // namespace formulon::io::xlsb::test
