#pragma once

#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::sheet_ops_test {

// Sentinel sheet index well past `sheet_count()` for every workbook in this
// family. Used to force the out-of-range rejection paths.
inline constexpr std::uint32_t kOutOfRangeIndex = 99U;

// Convenience factory: workbook with three sheets, named `Alpha`, `Beta`,
// and `Gamma` (created via `add_sheet` so the default `Sheet1` stays out of
// the way).
inline Workbook ThreeSheetWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  wb.add_sheet("Gamma");
  return wb;
}

inline void SetMoveSensitiveMetadata(Sheet& sheet, std::string_view suffix) {
  sheet.set_drawing_rel_target("xl/drawings/drawing-" + std::string(suffix) + ".xml");
  sheet.set_auto_filter_xml("<autoFilter ref=\"A1:" + std::string(suffix) + "9\"/>");
  sheet.set_ext_lst_xml("<extLst><ext uri=\"" + std::string(suffix) + "\"/></extLst>");
  sheet.set_root_extra_ns_attrs(" xmlns:x" + std::string(suffix) + "=\"urn:" + std::string(suffix) + "\"");
}

inline void ExpectMoveSensitiveMetadata(const Sheet& sheet, std::string_view suffix) {
  EXPECT_EQ(sheet.drawing_rel_target(), "xl/drawings/drawing-" + std::string(suffix) + ".xml");
  EXPECT_EQ(sheet.auto_filter_xml(), "<autoFilter ref=\"A1:" + std::string(suffix) + "9\"/>");
  EXPECT_EQ(sheet.ext_lst_xml(), "<extLst><ext uri=\"" + std::string(suffix) + "\"/></extLst>");
  EXPECT_EQ(sheet.root_extra_ns_attrs(), " xmlns:x" + std::string(suffix) + "=\"urn:" + std::string(suffix) + "\"");
}

}  // namespace formulon::sheet_ops_test
