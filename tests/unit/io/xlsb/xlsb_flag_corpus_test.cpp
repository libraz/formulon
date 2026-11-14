// Flag, reserved and enum fields of every record our .xlsb writer emits, checked against the Excel-saved corpus
// in tests/fixtures/excel/xlsb_flag_corpus_manifest.txt: a written field may carry only bits (flag, reserved) or
// values (enum) some manifest-listed Excel file carries in the same field. Our outputs are the excel_open_gate
// bases converted to .xlsb and every manifest file that loads, saved again.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb::test {
namespace {

enum class FieldKind { kFlag, kReserved, kEnum, kDataOnly };

/// One field of a record payload. A data-only row (width 0) states that the record has no flag, reserved or
/// enum field.
struct Field {
  std::uint16_t type;
  std::size_t offset;
  std::size_t width;
  FieldKind kind;
  const char* note;
};

constexpr FieldKind kF = FieldKind::kFlag;
constexpr FieldKind kR = FieldKind::kReserved;
constexpr FieldKind kE = FieldKind::kEnum;

constexpr Field DataOnly(std::uint16_t type, const char* note) {
  return Field{type, 0U, 0U, FieldKind::kDataOnly, note};
}

// clang-format off
constexpr Field kFields[] = {
    // Cells: column, iStyleRef (3 bytes), then fPhShow and reserved bits.
    {1, 7, 1, kF, "BrtCellBlank fPhShow"},
    {2, 7, 1, kF, "BrtCellRk fPhShow"},
    {3, 7, 1, kF, "BrtCellError fPhShow"},
    {3, 8, 1, kE, "BrtCellError bErr"},
    {4, 7, 1, kF, "BrtCellBool fPhShow"},
    {4, 8, 1, kE, "BrtCellBool fBool"},
    {5, 7, 1, kF, "BrtCellReal fPhShow"},
    {6, 7, 1, kF, "BrtCellSt fPhShow"},
    {7, 7, 1, kF, "BrtCellIsst fPhShow"},
    {8, 7, 1, kF, "BrtFmlaString fPhShow"},
    {9, 7, 1, kF, "BrtFmlaNum fPhShow"},
    {9, 16, 2, kF, "BrtFmlaNum grbitFlags (fAlwaysCalc)"},
    {10, 7, 1, kF, "BrtFmlaBool fPhShow"},
    {10, 9, 2, kF, "BrtFmlaBool grbitFlags"},
    {11, 7, 1, kF, "BrtFmlaError fPhShow"},
    {11, 8, 1, kE, "BrtFmlaError bErr"},
    {11, 9, 2, kF, "BrtFmlaError grbitFlags"},
    {426, 16, 1, kF, "BrtArrFmla fAlwaysCalc"},
    {428, 32, 1, kF, "BrtTable input and calc flags"},
    // Rows and columns.
    {0, 10, 1, kF, "BrtRowHdr fExtraAsc, fExtraDsc"},
    {0, 11, 1, kF, "BrtRowHdr outline level, fCollapsed, fDyZero, fUnsynced, fGhostDirty"},
    {0, 12, 1, kF, "BrtRowHdr fPhShow"},
    {60, 16, 2, kF, "BrtColInfo flags"},
    // Shared strings.
    {19, 0, 1, kF, "BrtSSTItem fRichStr, fExtStr"},
    // Revision-uid wrapper and future-record header.
    {35, 0, 4, kE, "BrtFRTBegin product version"},
    {37, 0, 2, kE, "BrtACBegin cver"},
    {37, 2, 2, kE, "BrtACBegin product version"},
    {37, 4, 2, kF, "BrtACBegin product flags"},
    // Workbook part.
    {39, 0, 4, kF, "BrtName grbit"},
    {153, 0, 4, kF, "BrtWbProp flags"},
    {156, 0, 4, kE, "BrtBundleSh hsState"},
    {157, 4, 4, kE, "BrtCalcProp fAutoRecalc"},
    {157, 24, 2, kF, "BrtCalcProp grbit"},
    {158, 28, 1, kF, "BrtBookView flags"},
    // Styles part.
    {43, 2, 2, kF, "BrtFont grbit"},
    {43, 4, 2, kE, "BrtFont bls"},
    {43, 6, 2, kE, "BrtFont sss"},
    {43, 8, 1, kE, "BrtFont uls"},
    {43, 9, 1, kE, "BrtFont bFamily"},
    {43, 10, 1, kE, "BrtFont bCharSet"},
    {43, 11, 1, kR, "BrtFont unused"},
    {43, 12, 1, kF, "BrtFont colour fValidRGB, xColorType"},
    {43, 20, 1, kE, "BrtFont bFontScheme"},
    {45, 0, 4, kE, "BrtFill fls"},
    {45, 4, 1, kF, "BrtFill foreground colour selector"},
    {45, 12, 1, kF, "BrtFill background colour selector"},
    {45, 20, 4, kE, "BrtFill iGradientType"},
    {46, 0, 1, kF, "BrtBorder fBdrDiagDown, fBdrDiagUp"},
    {46, 1, 1, kE, "BrtBorder top dg"},
    {46, 2, 1, kR, "BrtBorder top reserved"},
    {46, 3, 1, kF, "BrtBorder top colour selector"},
    {46, 11, 1, kE, "BrtBorder bottom dg"},
    {46, 12, 1, kR, "BrtBorder bottom reserved"},
    {46, 13, 1, kF, "BrtBorder bottom colour selector"},
    {46, 21, 1, kE, "BrtBorder left dg"},
    {46, 22, 1, kR, "BrtBorder left reserved"},
    {46, 23, 1, kF, "BrtBorder left colour selector"},
    {46, 31, 1, kE, "BrtBorder right dg"},
    {46, 32, 1, kR, "BrtBorder right reserved"},
    {46, 33, 1, kF, "BrtBorder right colour selector"},
    {46, 41, 1, kE, "BrtBorder diagonal dg"},
    {46, 42, 1, kR, "BrtBorder diagonal reserved"},
    {46, 43, 1, kF, "BrtBorder diagonal colour selector"},
    {47, 12, 2, kF, "BrtXF alignment and protection flags"},
    {47, 14, 2, kF, "BrtXF xfGrbitAtr"},
    {48, 4, 2, kF, "BrtStyle grbitObj1"},
    {48, 6, 1, kE, "BrtStyle iStyBuiltIn"},
    {507, 0, 2, kF, "BrtDXF header flags"},
    {507, 2, 2, kR, "BrtDXF reserved"},
    // Worksheet prefix and views.
    {147, 0, 2, kF, "BrtWsProp flags"},
    {147, 2, 1, kF, "BrtWsProp flags2"},
    {147, 3, 1, kF, "BrtWsProp tab colour selector"},
    {147, 11, 4, kR, "BrtWsProp rwSync"},
    {147, 15, 4, kR, "BrtWsProp colSync"},
    {137, 0, 2, kF, "BrtBeginWsView flags"},
    {137, 2, 4, kE, "BrtBeginWsView xlView"},
    {137, 14, 1, kE, "BrtBeginWsView icvHdr"},
    {137, 15, 1, kR, "BrtBeginWsView reserved2"},
    {137, 16, 2, kR, "BrtBeginWsView reserved3"},
    {151, 24, 4, kE, "BrtPane pnnAct"},
    {151, 28, 1, kF, "BrtPane fFrozen, fFrozenNoSplit"},
    {152, 0, 4, kE, "BrtSel pnn"},
    {485, 8, 2, kF, "BrtWsFmtInfo fUnsynced, fDyZero, fExAsc, fExDsc"},
    {485, 10, 1, kE, "BrtWsFmtInfo iOutLevelRw"},
    {485, 11, 1, kE, "BrtWsFmtInfo iOutLevelCol"},
    // Sheet protection: fLocked and fifteen allow flags, one u32 each after the u16 password hash.
    {535, 2, 4, kE, "BrtSheetProtection fLocked"},
    {535, 6, 4, kE, "BrtSheetProtection fObjects"},
    {535, 10, 4, kE, "BrtSheetProtection fScenarios"},
    {535, 14, 4, kE, "BrtSheetProtection fFormatCells"},
    {535, 18, 4, kE, "BrtSheetProtection fFormatColumns"},
    {535, 22, 4, kE, "BrtSheetProtection fFormatRows"},
    {535, 26, 4, kE, "BrtSheetProtection fInsertColumns"},
    {535, 30, 4, kE, "BrtSheetProtection fInsertRows"},
    {535, 34, 4, kE, "BrtSheetProtection fInsertHyperlinks"},
    {535, 38, 4, kE, "BrtSheetProtection fDeleteColumns"},
    {535, 42, 4, kE, "BrtSheetProtection fDeleteRows"},
    {535, 46, 4, kE, "BrtSheetProtection fSelLockedCells"},
    {535, 50, 4, kE, "BrtSheetProtection fSort"},
    {535, 54, 4, kE, "BrtSheetProtection fAutoFilter"},
    {535, 58, 4, kE, "BrtSheetProtection fPivotTables"},
    {535, 62, 4, kE, "BrtSheetProtection fSelUnlockedCells"},
    // Data validation and conditional formatting.
    {64, 0, 4, kF, "BrtDVal flags"},
    {573, 0, 2, kF, "BrtBeginDVals flags"},
    {573, 10, 4, kR, "BrtBeginDVals unused"},
    {461, 4, 4, kE, "BrtBeginConditionalFormatting fPivot"},
    {463, 0, 4, kE, "BrtBeginCFRule iType"},
    {463, 4, 4, kE, "BrtBeginCFRule iTemplate"},
    {463, 16, 4, kE, "BrtBeginCFRule iParam"},
    {463, 20, 4, kR, "BrtBeginCFRule reserved1"},
    {463, 24, 4, kR, "BrtBeginCFRule reserved2"},
    {463, 28, 2, kF, "BrtBeginCFRule fStopTrue, fAbove, fBottom, fPercent"},
    {465, 0, 4, kE, "BrtBeginIconSet iSet"},
    {465, 4, 2, kF, "BrtBeginIconSet flags"},
    {467, 2, 1, kE, "BrtBeginDataBar fShowValue"},
    {471, 0, 4, kE, "BrtCFVO iType"},
    {471, 12, 4, kE, "BrtCFVO fSaveGTE"},
    {471, 16, 4, kE, "BrtCFVO fGTE"},
    {564, 0, 1, kF, "BrtColor colour selector"},
    {1146, 0, 4, kR, "BrtCFRuleExt reserved"},
    // Autofilter and sort state.
    {163, 4, 2, kF, "BrtBeginFilterColumn fHideArrow, fNoBtn"},
    {165, 0, 4, kE, "BrtBeginFilters fBlank"},
    {165, 4, 4, kE, "BrtBeginFilters iCalendarType"},
    {168, 4, 4, kE, "BrtColorFilter fCellColor"},
    {169, 0, 4, kE, "BrtIconFilter iIconSet"},
    {170, 0, 1, kF, "BrtTop10Filter fTop, fPercent"},
    {171, 0, 4, kE, "BrtDynamicFilter cft"},
    {171, 4, 1, kF, "BrtDynamicFilter value flags"},
    {172, 0, 4, kE, "BrtBeginCustomFilters fAnd"},
    {174, 0, 1, kE, "BrtCustomFilter vts"},
    {174, 1, 1, kE, "BrtCustomFilter grbitSgn"},
    {175, 14, 2, kR, "BrtAFilterDateGroupItem unused"},
    {175, 16, 4, kR, "BrtAFilterDateGroupItem unused"},
    {175, 20, 4, kE, "BrtAFilterDateGroupItem dateGroupingType"},
    {530, 0, 2, kF, "BrtBeginSortState fColumnSort, fCaseSensitive, fAltMethod"},
    {532, 0, 2, kF, "BrtBeginSortCond fSortDes, sortOn"},
    // Tables.
    {343, 16, 4, kE, "BrtBeginList lt"},
    {343, 24, 4, kE, "BrtBeginList crwHeader"},
    {343, 28, 4, kE, "BrtBeginList crwTotals"},
    {343, 32, 4, kF, "BrtBeginList flags"},
    {347, 4, 4, kE, "BrtBeginListCol ilta"},
    {347, 8, 4, kE, "BrtBeginListCol nDxfHdr (Excel refuses a header dxf)"},
    {351, 0, 1, kF, "BrtListCCFmla fArray"},
    {352, 0, 1, kF, "BrtListTrFmla fArray"},
    {513, 0, 2, kF, "BrtTableStyleClient flags"},
    {1111, 0, 4, kF, "BrtList14 flags"},
    // Records with no flag, reserved or enum field.
    DataOnly(36, "BrtFRTEnd"),
    DataOnly(38, "BrtACEnd"),
    DataOnly(44, "BrtFmt: id and format string"),
    DataOnly(49, "BrtCellMeta: metadata index"),
    DataOnly(50, "BrtValueMeta: metadata index"),
    DataOnly(128, "BrtFileVersion: GUID and strings"),
    DataOnly(129, "BrtBeginSheet"),
    DataOnly(130, "BrtEndSheet"),
    DataOnly(131, "BrtBeginBook"),
    DataOnly(132, "BrtEndBook"),
    DataOnly(133, "BrtBeginWsViews"),
    DataOnly(134, "BrtEndWsViews"),
    DataOnly(135, "BrtBeginBookViews"),
    DataOnly(136, "BrtEndBookViews"),
    DataOnly(138, "BrtEndWsView"),
    DataOnly(143, "BrtBeginBundleShs"),
    DataOnly(144, "BrtEndBundleShs"),
    DataOnly(145, "BrtBeginSheetData"),
    DataOnly(146, "BrtEndSheetData"),
    DataOnly(148, "BrtWsDim: RfX"),
    DataOnly(159, "BrtBeginSst: counts"),
    DataOnly(160, "BrtEndSst"),
    DataOnly(161, "BrtBeginAFilter: RfX"),
    DataOnly(162, "BrtEndAFilter"),
    DataOnly(164, "BrtEndFilterColumn"),
    DataOnly(166, "BrtEndFilters"),
    DataOnly(167, "BrtFilter: string"),
    DataOnly(173, "BrtEndCustomFilters"),
    DataOnly(176, "BrtMergeCell: RfX"),
    DataOnly(177, "BrtBeginMergeCells: count"),
    DataOnly(178, "BrtEndMergeCells"),
    DataOnly(278, "BrtBeginStyleSheet"),
    DataOnly(279, "BrtEndStyleSheet"),
    DataOnly(344, "BrtEndList"),
    DataOnly(345, "BrtBeginListCols: count"),
    DataOnly(346, "BrtEndListCols"),
    DataOnly(348, "BrtEndListCol"),
    DataOnly(353, "BrtBeginExternals"),
    DataOnly(354, "BrtEndExternals"),
    DataOnly(357, "BrtSupSelf"),
    DataOnly(362, "BrtExternSheet: XTI triples"),
    DataOnly(384, "BrtBeginPivotCacheIDs"),
    DataOnly(385, "BrtEndPivotCacheIDs"),
    DataOnly(386, "BrtBeginPivotCacheID: idSx and relationship id"),
    DataOnly(387, "BrtEndPivotCacheID"),
    DataOnly(390, "BrtBeginColInfos"),
    DataOnly(391, "BrtEndColInfos"),
    DataOnly(462, "BrtEndConditionalFormatting"),
    DataOnly(464, "BrtEndCFRule"),
    DataOnly(466, "BrtEndIconSet"),
    DataOnly(468, "BrtEndDataBar"),
    DataOnly(469, "BrtBeginColorScale"),
    DataOnly(470, "BrtEndColorScale"),
    DataOnly(494, "BrtHLink: RfX and strings"),
    DataOnly(505, "BrtBeginDXFs: count"),
    DataOnly(506, "BrtEndDXFs"),
    DataOnly(531, "BrtEndSortState"),
    DataOnly(533, "BrtEndSortCond"),
    DataOnly(550, "BrtDrawing: relationship id"),
    DataOnly(574, "BrtEndDVals"),
    DataOnly(603, "BrtBeginFills: count"),
    DataOnly(604, "BrtEndFills"),
    DataOnly(611, "BrtBeginFonts: count"),
    DataOnly(612, "BrtEndFonts"),
    DataOnly(613, "BrtBeginBorders: count"),
    DataOnly(614, "BrtEndBorders"),
    DataOnly(615, "BrtBeginFmts: count"),
    DataOnly(616, "BrtEndFmts"),
    DataOnly(617, "BrtBeginCellXFs: count"),
    DataOnly(618, "BrtEndCellXFs"),
    DataOnly(619, "BrtBeginStyles: count"),
    DataOnly(620, "BrtEndStyles"),
    DataOnly(626, "BrtBeginCellStyleXFs: count"),
    DataOnly(627, "BrtEndCellStyleXFs"),
    DataOnly(660, "BrtBeginListParts: count"),
    DataOnly(661, "BrtListPart: relationship id"),
    DataOnly(662, "BrtEndListParts"),
    DataOnly(2071, "BrtAbsPath15: string"),
    DataOnly(3072, "BrtUid: GUID"),
    // Compatibility Version stamp: reserved, warnBelowVersion (0 = none), setVersion.
    {5135, 0, 4, kR, "workbook Compatibility Version reserved"},
    {5135, 4, 4, kE, "workbook Compatibility Version warnBelowVersion"},
    {5135, 8, 4, kE, "workbook Compatibility Version setVersion"},
};

// Record types our outputs carry that no row classifies yet. Each leaves this list when its rows are added, and
// the list is empty once every record the writer emits is classified.
constexpr std::uint16_t kPendingClassification[] = {
    21, 24, 27, 31, 33, 51, 52, 53, 179, 180, 181, 182, 183, 184, 185, 186,
    187, 188, 189, 190, 191, 192, 193, 194, 219, 220, 221, 222, 223, 224, 243, 244,
    245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260,
    280, 281, 282, 283, 284, 285, 286, 287, 288, 289, 290, 291, 292, 293, 294, 295,
    296, 297, 298, 299, 300, 301, 302, 309, 310, 311, 312, 313, 314, 315, 332, 333,
    334, 335, 336, 337, 338, 339, 340, 355, 359, 360, 363, 364, 366, 368, 369, 370,
    371, 382, 383, 388, 389, 392, 393, 394, 395, 396, 397, 422, 423, 424, 425, 459, 460,
    476, 477, 478, 479, 480, 495, 496, 497, 498, 499, 500, 501, 502, 503, 504,
    508, 509, 514, 515, 516, 517, 518, 519, 520, 521, 522, 523, 525, 534, 536, 537,
    551, 554, 555, 556, 557, 577, 585, 586, 587, 588, 599, 600, 601, 602, 605, 606,
    607, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 648, 649, 650, 677, 678,
    1025, 1041, 1042, 1043, 1046, 1047, 1048, 1049, 1050, 1051, 1052, 1053, 1054, 1055, 1056, 1057,
    1058, 1059, 1062, 1063, 1066, 1067, 1075, 1076, 1077, 1078, 1079, 1080, 1081, 1082, 1083, 1084,
    1085, 1112, 1115, 1116, 1131, 1132, 1135, 1136, 1142, 1143, 1154, 1155, 1156, 2075, 2076, 2077,
    2078, 2096, 2097, 2098, 2099, 4096, 4097, 5000, 5001, 5002, 5003, 5081, 5082, 5084, 5085, 5094,
    5108, 5153,
};
// clang-format on

/// A field whose written values fall outside Excel's union because no manifest file shows the state yet.
/// For a flag or reserved field `values` holds the one mask of bits outside the union; for an enum field, the
/// values outside it. An entry goes once the corpus gains a file that carries the state, or the writer stops
/// writing it.
struct AwaitingEvidence {
  std::uint16_t type;
  std::size_t offset;
  std::set<std::uint32_t> values;
  const char* reason;
};

const AwaitingEvidence kAwaitingExcelEvidence[] = {
    {157,
     24,
     {0x01U},
     "BrtCalcProp fFullCalcOnLoad: Excel clears this one-shot request when it saves (claims/full_calc_on_load keeps "
     "0x6a although its source asks for a full calculation), so no Excel save carries it"},
};

std::string FixturesRoot() {
  return FORMULON_FIXTURES_DIR;
}

std::vector<std::string> ManifestPaths() {
  std::vector<std::string> out;
  std::ifstream in(FixturesRoot() + "/excel/xlsb_flag_corpus_manifest.txt");
  std::string line;
  while (std::getline(in, line)) {
    std::istringstream words(line.substr(0, line.find('#')));
    std::string path;
    if (words >> path) {
      out.push_back(path);
    }
  }
  return out;
}

std::uint32_t FieldValue(const Bytes& payload, const Field& field) {
  std::uint32_t v = 0;
  for (std::size_t i = 0; i < field.width; ++i) {
    v |= static_cast<std::uint32_t>(payload[field.offset + i]) << (8U * i);
  }
  return v;
}

bool Holds(const Bytes& payload, const Field& field) {
  return field.kind != FieldKind::kDataOnly && field.offset + field.width <= payload.size();
}

/// Calls `visit(record)` for every record of every record part of `package`.
template <class Visit>
void ForEachRecord(const Bytes& package, Visit visit) {
  for (const auto& part : ExtractRecords(package, {})) {
    for (const IndexedRecord& rec : part.second) {
      visit(part.first, rec);
    }
  }
}

struct Output {
  std::string name;
  Bytes bytes;
};

/// The converted gate bases and the re-saved manifest files, built once per test binary.
const std::vector<Output>& Outputs() {
  static const std::vector<Output> outputs = [] {
    std::vector<Output> out;
    const std::string gate_list = FixturesRoot() + "/../../tools/oracle/excel_open_gate.txt";
    for (const std::string& base : GateBases(gate_list)) {
      Converted converted = ConvertToXlsb(FixturesRoot(), base);
      EXPECT_TRUE(converted.error.empty()) << converted.error;
      out.push_back(Output{"base " + base, std::move(converted.bytes)});
    }
    for (const std::string& path : ManifestPaths()) {
      Converted converted = ConvertToXlsb(FixturesRoot(), path);
      if (converted.error.empty()) {
        out.push_back(Output{"resave " + path, std::move(converted.bytes)});
      }
    }
    return out;
  }();
  return outputs;
}

std::multimap<std::uint16_t, std::size_t> RowsByType() {
  std::multimap<std::uint16_t, std::size_t> out;
  for (std::size_t i = 0; i < std::size(kFields); ++i) {
    out.emplace(kFields[i].type, i);
  }
  return out;
}

std::string RowName(const Field& field) {
  return std::to_string(field.type) + "@" + std::to_string(field.offset) + " (" + field.note + ")";
}

TEST(XlsbFlagCorpus, ExcelUnionCoversWriterBits) {
  const std::multimap<std::uint16_t, std::size_t> rows = RowsByType();
  std::vector<std::uint32_t> union_bits(std::size(kFields), 0U);
  std::vector<std::set<std::uint32_t>> union_values(std::size(kFields));
  const std::vector<std::string> manifest = ManifestPaths();
  ASSERT_FALSE(manifest.empty());
  for (const std::string& path : manifest) {
    const Bytes bytes = ReadBytes(FixturesRoot() + "/" + path);
    EXPECT_FALSE(bytes.empty()) << "manifest file missing: " << path;
    ForEachRecord(bytes, [&](const std::string&, const IndexedRecord& rec) {
      const auto range = rows.equal_range(rec.type);
      for (auto it = range.first; it != range.second; ++it) {
        const Field& field = kFields[it->second];
        if (!Holds(rec.payload, field)) {
          continue;
        }
        const std::uint32_t v = FieldValue(rec.payload, field);
        union_bits[it->second] |= v;
        union_values[it->second].insert(v);
      }
    });
  }

  std::map<std::pair<std::uint16_t, std::size_t>, std::set<std::uint32_t>> outside;
  std::map<std::pair<std::uint16_t, std::size_t>, std::string> first_seen;
  ASSERT_FALSE(Outputs().empty());
  for (const Output& output : Outputs()) {
    ForEachRecord(output.bytes, [&](const std::string& part, const IndexedRecord& rec) {
      const auto range = rows.equal_range(rec.type);
      for (auto it = range.first; it != range.second; ++it) {
        const Field& field = kFields[it->second];
        if (!Holds(rec.payload, field)) {
          continue;
        }
        const std::uint32_t v = FieldValue(rec.payload, field);
        const auto key = std::make_pair(field.type, field.offset);
        if (field.kind == FieldKind::kEnum) {
          if (union_values[it->second].count(v) != 0U) {
            continue;
          }
          outside[key].insert(v);
        } else {
          const std::uint32_t offending = v & ~union_bits[it->second];
          if (offending == 0U) {
            continue;
          }
          std::set<std::uint32_t>& masks = outside[key];
          const std::uint32_t merged = (masks.empty() ? 0U : *masks.begin()) | offending;
          masks = {merged};
        }
        first_seen.emplace(key, RowName(field) + " in " + output.name + " " + part);
      }
    });
  }

  std::map<std::pair<std::uint16_t, std::size_t>, const AwaitingEvidence*> awaiting;
  for (const AwaitingEvidence& entry : kAwaitingExcelEvidence) {
    awaiting.emplace(std::make_pair(entry.type, entry.offset), &entry);
  }
  for (const auto& hit : outside) {
    const auto known = awaiting.find(hit.first);
    if (known == awaiting.end()) {
      ADD_FAILURE() << "written outside Excel's union: " << first_seen[hit.first] << ", values "
                    << ::testing::PrintToString(hit.second);
      continue;
    }
    EXPECT_EQ(hit.second, known->second->values) << known->second->reason;
  }
  for (const auto& entry : awaiting) {
    EXPECT_NE(outside.count(entry.first), 0U)
        << "no longer outside Excel's union, remove the entry: " << entry.second->reason;
  }
}

TEST(XlsbFlagCorpus, EveryEmittedRecordTypeHasARow) {
  std::set<std::uint16_t> classified;
  for (const Field& field : kFields) {
    classified.insert(field.type);
  }
  const std::set<std::uint16_t> pending(std::begin(kPendingClassification), std::end(kPendingClassification));
  EXPECT_EQ(pending.size(), std::size(kPendingClassification)) << "duplicate pending entry";
  for (const std::uint16_t type : pending) {
    EXPECT_EQ(classified.count(type), 0U) << "record " << type << " has a row, remove it from the pending list";
  }
  std::map<std::uint16_t, std::string> unclassified;
  ASSERT_FALSE(Outputs().empty());
  for (const Output& output : Outputs()) {
    ForEachRecord(output.bytes, [&](const std::string& part, const IndexedRecord& rec) {
      if (classified.count(rec.type) == 0U && pending.count(rec.type) == 0U) {
        unclassified.emplace(rec.type, output.name + " " + part);
      }
    });
  }
  for (const auto& hit : unclassified) {
    ADD_FAILURE() << "record " << hit.first << " has no field row (first seen in " << hit.second << ")";
  }
}

}  // namespace
}  // namespace formulon::io::xlsb::test
