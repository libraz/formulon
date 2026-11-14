//
// A pivot cache refreshed from an edited source survives a save: the .xlsx
// written after the refresh reads back to the refreshed cache and the pivot
// definitions remapped onto it. An .xlsb pivot is kept as loaded parts, which
// a refresh does not reach.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

// The pivot sits on the first sheet; its worksheet source is Sheet1!A1:E13
// (Region, Product, Date, Amt, Qty).
constexpr std::size_t kSourceSheet = 1U;
constexpr std::uint32_t kRegionCol = 0U;
constexpr std::uint32_t kAmtCol = 3U;

std::vector<std::uint8_t> ReadFixture(const char* name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_twins/" + name;
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "missing fixture " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string Show(const Value& v) {
  std::ostringstream out;
  out.precision(17);
  if (v.is_text()) {
    out << "s:" << v.as_text();
  } else if (v.is_number()) {
    out << "n:" << v.as_number();
  } else if (v.is_boolean()) {
    out << "b:" << (v.as_boolean() ? 1 : 0);
  } else if (v.is_error()) {
    out << "e:" << static_cast<int>(v.as_error());
  } else {
    out << "blank";
  }
  return out.str();
}

// Every cache field, shared item and record cell, and every pivot's field
// items, as text lines that compare whole.
std::vector<std::string> CacheAndPivotLines(const Workbook& wb) {
  std::vector<std::string> lines;
  for (const auto& cache : wb.pivot_caches()) {
    lines.push_back("cache " + std::to_string(cache->cache_id()) + " records " +
                    std::to_string(cache->records().size()));
    for (const pivot::PivotCacheField& field : cache->fields()) {
      std::string line = "field " + field.name + ":";
      for (const Value& item : field.shared_items) {
        line += " " + Show(item);
      }
      lines.push_back(line);
    }
    for (const pivot::PivotCacheRecord& record : cache->records()) {
      std::string line = "record:";
      for (std::size_t c = 0; c < record.cells.size(); ++c) {
        const Value& v = record.cells[c];
        const bool index = c < record.cell_is_index.size()
                               ? record.cell_is_index[c]
                               : c < cache->fields().size() && !cache->fields()[c].shared_items.empty();
        line += " " + (index && v.is_number()
                           ? Show(cache->fields()[c].shared_items.at(static_cast<std::size_t>(v.as_number())))
                           : Show(v));
      }
      lines.push_back(line);
    }
  }
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& table : wb.sheet(s).pivot_tables()) {
      lines.push_back("pivot " + table->name());
      for (const auto& field : table->fields()) {
        std::string line = "pivot field:";
        for (const auto& item : field.items) {
          line += " " + item.name + (item.visible ? "" : "(hidden)");
        }
        lines.push_back(line);
      }
    }
  }
  return lines;
}

// Loads `file`, renames one region and changes one amount in the source, and
// refreshes every cache.
template <typename Read>
Workbook LoadEditAndRefresh(const char* file, Read read) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  auto loaded = read(bytes);
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  Workbook wb = std::move(loaded.value().workbook);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kSourceSheet, 1U, kRegionCol, "West")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(kSourceSheet, 2U, kAmtCol, Value::number(999.0))));
  for (const auto& cache : wb.pivot_caches()) {
    const auto refreshed = wb.refresh_pivot_cache(cache->cache_id(), eval::default_registry());
    EXPECT_TRUE(static_cast<bool>(refreshed)) << (refreshed ? "" : refreshed.error().message);
  }
  return wb;
}

TEST(PivotRefreshRoundTrip, XlsxSaveReadsBackTheRefreshedCache) {
  Workbook wb = LoadEditAndRefresh("B01_base.xlsx", [](const std::vector<std::uint8_t>& bytes) {
    return io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  });
  const std::vector<std::string> refreshed = CacheAndPivotLines(wb);
  bool saw_west = false;
  for (const std::string& line : refreshed) {
    saw_west = saw_west || line.find("s:West") != std::string::npos;
  }
  ASSERT_TRUE(saw_west) << "the refresh did not pick up the edited source";

  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto reread = io::read_ooxml(io::ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  EXPECT_EQ(CacheAndPivotLines(reread.value().workbook), refreshed);
  EXPECT_FALSE(reread.value().workbook.pivot_caches().front()->refresh_on_load());
}

// An .xlsb pivot is not modelled: its cache and table parts are kept as
// loaded, so a refresh has nothing to rebuild and the save writes those parts
// unchanged next to the edited source.
TEST(PivotRefreshRoundTrip, XlsbPivotPartsPassThroughARefreshUnchanged) {
  const std::vector<std::uint8_t> bytes = ReadFixture("B01_base.xlsb");
  auto loaded = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  std::map<std::string, std::vector<std::uint8_t>> pivot_parts;
  for (const PassthroughPart& part : loaded.value().workbook.passthrough_parts()) {
    if (part.path.rfind("xl/pivot", 0) == 0) {
      pivot_parts[part.path] = part.bytes;
    }
  }
  ASSERT_FALSE(pivot_parts.empty());

  Workbook wb = LoadEditAndRefresh("B01_base.xlsb", [](const std::vector<std::uint8_t>& in) {
    return io::xlsb::read_xlsb(io::ByteSpan{in.data(), in.size()});
  });
  EXPECT_TRUE(wb.pivot_caches().empty());
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto reread = io::xlsb::read_xlsb(io::ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;

  std::map<std::string, std::vector<std::uint8_t>> saved_parts;
  for (const PassthroughPart& part : reread.value().workbook.passthrough_parts()) {
    if (part.path.rfind("xl/pivot", 0) == 0) {
      saved_parts[part.path] = part.bytes;
    }
  }
  EXPECT_EQ(saved_parts, pivot_parts);
  const Value edited = reread.value().workbook.sheet(kSourceSheet).resolve_cell_value(1U, kRegionCol);
  ASSERT_TRUE(edited.is_text());
  EXPECT_EQ(edited.as_text(), "West");
}

}  // namespace
}  // namespace formulon
