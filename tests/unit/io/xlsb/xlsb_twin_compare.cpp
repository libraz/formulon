#include "xlsb_twin_compare.h"

#include <algorithm>
#include <cstdio>
#include <fstream>
#include <sstream>
#include <string_view>

#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

namespace formulon::io::xlsb::test {
namespace {

constexpr std::uint16_t kBrtFmlaString = 8;
constexpr std::uint16_t kBrtFmlaNum = 9;
constexpr std::uint16_t kBrtFmlaBool = 10;
constexpr std::uint16_t kBrtFmlaError = 11;
constexpr std::uint16_t kBrtArrFmla = 426;
constexpr std::uint16_t kBrtShrFmla = 427;
constexpr std::uint8_t kPtgAttr = 0x19;
constexpr std::uint8_t kAttrSemi = 0x01;
constexpr std::size_t kCellHeader = 8;
constexpr std::size_t kRfx = 16;

ByteSpan SpanOfBytes(const Bytes& v) {
  return ByteSpan{v.data(), v.size()};
}

bool EndsWith(std::string_view s, std::string_view suffix) {
  return s.size() >= suffix.size() && s.compare(s.size() - suffix.size(), suffix.size(), suffix) == 0;
}

bool StartsWith(std::string_view s, std::string_view prefix) {
  return s.compare(0, prefix.size(), prefix) == 0;
}

std::uint32_t U32(const Bytes& p, std::size_t at) {
  return static_cast<std::uint32_t>(p[at]) | (static_cast<std::uint32_t>(p[at + 1U]) << 8U) |
         (static_cast<std::uint32_t>(p[at + 2U]) << 16U) | (static_cast<std::uint32_t>(p[at + 3U]) << 24U);
}

/// Offset of `cce` in a formula record's payload, or nullopt for a record holding no formula.
std::optional<std::size_t> CceOffset(std::uint16_t type, const Bytes& p) {
  switch (type) {
    case kBrtFmlaNum:
      return kCellHeader + 8U + 2U;
    case kBrtFmlaBool:
    case kBrtFmlaError:
      return kCellHeader + 1U + 2U;
    case kBrtFmlaString:
      if (p.size() < kCellHeader + 4U) {
        return std::nullopt;
      }
      return kCellHeader + 4U + 2U * static_cast<std::size_t>(U32(p, kCellHeader)) + 2U;
    case kBrtArrFmla:
      return kRfx + 1U;
    case kBrtShrFmla:
      return kRfx;
    default:
      return std::nullopt;
  }
}

std::string Hex(const Bytes& p, std::size_t begin, std::size_t end) {
  static const char kDigits[] = "0123456789abcdef";
  std::string out;
  for (std::size_t i = begin; i < end && i < p.size(); ++i) {
    if (!out.empty()) {
      out.push_back(' ');
    }
    out.push_back(kDigits[p[i] >> 4U]);
    out.push_back(kDigits[p[i] & 0x0FU]);
  }
  return out;
}

std::string Where(const std::string& part, std::uint16_t type, std::size_t index) {
  return part + " record #" + std::to_string(index) + " type " + std::to_string(type);
}

/// Appends one line per run of differing, unmasked payload bytes.
void CompareRecord(const std::string& part, const IndexedRecord& ours, const IndexedRecord& excel, const ByteMask& mask,
                   std::vector<std::string>* out) {
  if (mask.whole_records.count({part, excel.index}) != 0U) {
    return;
  }
  if (ours.type != excel.type) {
    out->push_back(Where(part, excel.type, excel.index) + ": ours is type " + std::to_string(ours.type));
    return;
  }
  if (ours.payload.size() != excel.payload.size()) {
    out->push_back(Where(part, excel.type, excel.index) + ": payload " + std::to_string(ours.payload.size()) +
                   " bytes, Excel " + std::to_string(excel.payload.size()));
    return;
  }
  std::size_t i = 0;
  while (i < excel.payload.size()) {
    if (ours.payload[i] == excel.payload[i] || mask.Masks(part, excel.index, i)) {
      ++i;
      continue;
    }
    std::size_t end = i + 1U;
    while (end < excel.payload.size() && ours.payload[end] != excel.payload[end] &&
           !mask.Masks(part, excel.index, end)) {
      ++end;
    }
    out->push_back(Where(part, excel.type, excel.index) + " offset " + std::to_string(i) + ": ours " +
                   Hex(ours.payload, i, end) + ", Excel " + Hex(excel.payload, i, end));
    i = end;
  }
}

}  // namespace

Bytes ReadBytes(const std::string& path) {
  Bytes out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    out.clear();
  }
  std::fclose(f);
  return out;
}

std::vector<std::string> ListParts(const Bytes& package) {
  ZipReader zip;
  if (!zip.open(SpanOfBytes(package))) {
    return {};
  }
  return zip.list_entries();
}

std::optional<Bytes> ReadPart(const Bytes& package, const std::string& name) {
  ZipReader zip;
  if (!zip.open(SpanOfBytes(package))) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(name);
  if (!bytes) {
    return std::nullopt;
  }
  return std::move(bytes.value());
}

bool IsRecordPart(const std::string& name) {
  return EndsWith(name, ".bin") && !StartsWith(name, "xl/printerSettings/") && !StartsWith(name, "xl/embeddings/") &&
         !StartsWith(name, "xl/activeX/") && !EndsWith(name, "vbaProject.bin");
}

std::vector<IndexedRecord> SplitPart(const Bytes& part) {
  std::vector<IndexedRecord> out;
  const std::vector<FramedRecord> framed = split_records(part);
  out.reserve(framed.size());
  for (std::size_t i = 0; i < framed.size(); ++i) {
    const FramedRecord& rec = framed[i];
    out.push_back(IndexedRecord{i, rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size)});
  }
  return out;
}

RecordSet ExtractRecords(const Bytes& package, const std::set<std::uint16_t>& types) {
  RecordSet out;
  ZipReader zip;
  if (!zip.open(SpanOfBytes(package))) {
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    if (!IsRecordPart(name)) {
      continue;
    }
    auto bytes = zip.read_entry(name);
    if (!bytes) {
      continue;
    }
    std::vector<IndexedRecord>& kept = out[name];
    for (IndexedRecord& rec : SplitPart(bytes.value())) {
      if (types.empty() || types.count(rec.type) != 0U) {
        kept.push_back(std::move(rec));
      }
    }
  }
  return out;
}

bool ByteMask::Masks(const std::string& part, std::size_t index, std::size_t offset) const {
  if (whole_parts.count(part) != 0U || whole_records.count({part, index}) != 0U) {
    return true;
  }
  const auto p = bytes.find(part);
  if (p == bytes.end()) {
    return false;
  }
  const auto r = p->second.find(index);
  return r != p->second.end() && r->second.count(offset) != 0U;
}

std::size_t ByteMask::ByteCount() const {
  std::size_t n = 0;
  for (const auto& part : bytes) {
    for (const auto& rec : part.second) {
      n += rec.second.size();
    }
  }
  return n;
}

ByteMask BuildTwinMask(const Bytes& first_save, const Bytes& second_save) {
  ByteMask mask;
  const std::vector<std::string> first_parts = ListParts(first_save);
  const std::vector<std::string> second_parts = ListParts(second_save);
  const std::set<std::string> in_second(second_parts.begin(), second_parts.end());
  for (const std::string& name : first_parts) {
    if (in_second.count(name) == 0U) {
      mask.whole_parts.insert(name);
    }
  }
  const std::set<std::string> in_first(first_parts.begin(), first_parts.end());
  for (const std::string& name : second_parts) {
    if (in_first.count(name) == 0U) {
      mask.whole_parts.insert(name);
      continue;
    }
    const std::optional<Bytes> a = ReadPart(first_save, name);
    const std::optional<Bytes> b = ReadPart(second_save, name);
    if (!a || !b) {
      mask.whole_parts.insert(name);
      continue;
    }
    if (*a == *b) {
      continue;
    }
    if (!IsRecordPart(name)) {
      mask.whole_parts.insert(name);
      continue;
    }
    const std::vector<IndexedRecord> ra = SplitPart(*a);
    const std::vector<IndexedRecord> rb = SplitPart(*b);
    if (ra.empty() || rb.empty()) {
      mask.whole_parts.insert(name);
      continue;
    }
    for (std::size_t i = 0; i < std::max(ra.size(), rb.size()); ++i) {
      if (i >= ra.size() || i >= rb.size() || ra[i].type != rb[i].type ||
          ra[i].payload.size() != rb[i].payload.size()) {
        mask.whole_records.insert({name, i});
        continue;
      }
      for (std::size_t o = 0; o < ra[i].payload.size(); ++o) {
        if (ra[i].payload[o] != rb[i].payload[o]) {
          mask.bytes[name][i].insert(o);
        }
      }
    }
  }
  return mask;
}

void MaskLeadingAttrSemi(std::uint16_t type, const Bytes& payload, std::vector<std::size_t>* offsets) {
  const std::optional<std::size_t> cce_at = CceOffset(type, payload);
  if (!cce_at || *cce_at + 4U + 4U > payload.size()) {
    return;
  }
  const std::size_t rgce = *cce_at + 4U;
  if (U32(payload, *cce_at) >= 4U && payload[rgce] == kPtgAttr && payload[rgce + 1U] == kAttrSemi) {
    offsets->push_back(rgce + 2U);
    offsets->push_back(rgce + 3U);
  }
}

std::vector<UninitialisedFieldHook> DefaultUninitialisedHooks() {
  return {MaskLeadingAttrSemi};
}

void AddUninitialisedFields(const Bytes& excel_package, const std::vector<UninitialisedFieldHook>& hooks,
                            ByteMask* mask) {
  for (const auto& part : ExtractRecords(excel_package, {})) {
    for (const IndexedRecord& rec : part.second) {
      std::vector<std::size_t> offsets;
      for (const UninitialisedFieldHook& hook : hooks) {
        hook(rec.type, rec.payload, &offsets);
      }
      if (!offsets.empty()) {
        mask->bytes[part.first][rec.index].insert(offsets.begin(), offsets.end());
      }
    }
  }
}

std::vector<std::string> CompareRecordSets(const Bytes& ours, const Bytes& excel, const std::set<std::uint16_t>& types,
                                           const ByteMask& mask) {
  std::vector<std::string> out;
  const RecordSet a = ExtractRecords(ours, types);
  const RecordSet b = ExtractRecords(excel, types);
  for (const auto& part : b) {
    if (mask.whole_parts.count(part.first) != 0U) {
      continue;
    }
    const auto mine = a.find(part.first);
    if (mine == a.end()) {
      if (!part.second.empty()) {
        out.push_back(part.first + ": missing from ours");
      }
      continue;
    }
    if (mine->second.size() != part.second.size()) {
      out.push_back(part.first + ": " + std::to_string(mine->second.size()) + " selected records, Excel " +
                    std::to_string(part.second.size()));
    }
    for (std::size_t i = 0; i < std::min(mine->second.size(), part.second.size()); ++i) {
      CompareRecord(part.first, mine->second[i], part.second[i], mask, &out);
    }
  }
  for (const auto& part : a) {
    if (b.count(part.first) == 0U && !part.second.empty() && mask.whole_parts.count(part.first) == 0U) {
      out.push_back(part.first + ": not in Excel's save");
    }
  }
  return out;
}

std::vector<std::string> GateBases(const std::string& gate_list_path) {
  std::vector<std::string> out;
  std::ifstream in(gate_list_path);
  std::string line;
  while (std::getline(in, line)) {
    std::istringstream words(line);
    std::string first;
    if (!(words >> first) || first[0] == '#' || first.find(':') != std::string::npos) {
      continue;
    }
    if (EndsWith(first, ".xlsx") || EndsWith(first, ".xlsb")) {
      out.push_back(first);
    }
  }
  return out;
}

Converted ConvertToXlsb(const std::string& fixtures_root, const std::string& relative) {
  Converted out;
  const Bytes source = ReadBytes(fixtures_root + "/" + relative);
  if (source.empty()) {
    out.error = "cannot read " + relative;
    return out;
  }
  const auto write = [&](const Workbook& workbook) {
    auto written = write_xlsb_with_result(workbook);
    if (!written) {
      out.error = relative + ": " + written.error().message;
      return;
    }
    out.bytes = std::move(written.value().bytes);
  };
  if (EndsWith(relative, ".xlsb")) {
    auto loaded = read_xlsb(SpanOfBytes(source));
    if (!loaded) {
      out.error = relative + ": " + loaded.error().message;
      return out;
    }
    write(loaded.value().workbook);
  } else {
    auto loaded = read_ooxml(SpanOfBytes(source));
    if (!loaded) {
      out.error = relative + ": " + loaded.error().message;
      return out;
    }
    write(loaded.value().workbook);
  }
  return out;
}

}  // namespace formulon::io::xlsb::test
