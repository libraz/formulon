#include "chart/chart_template.h"

#include <algorithm>
#include <cstdint>
#include <initializer_list>
#include <regex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "styles.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using D = ChartSeriesDirection;

// ---------------------------------------------------------------------------
// Measured chart-part fragments (normalised: axis ids, unique ids and caches
// are placeholders; `@X@` marks a substitution point of the composer below).
// ---------------------------------------------------------------------------

constexpr const char* kHead = R"X(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="@L@"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart>)X";
constexpr const char* kTitle =
    R"X(<c:title><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr></c:title>)X";
constexpr const char* kSerBar =
    R"X(<c:ser><c:idx val="@I@"/><c:order val="@I@"/><c:tx><c:strRef><c:f>@NAME@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:invertIfNegative val="0"/><c:cat><c:strRef><c:f>@CAT@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>@VAL@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:val><c:extLst><c:ext uri="{UID}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{UID}"/></c:ext></c:extLst></c:ser>)X";
constexpr const char* kSerLine =
    R"X(<c:ser><c:idx val="@I@"/><c:order val="@I@"/><c:tx><c:strRef><c:f>@NAME@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:tx><c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:cat><c:strRef><c:f>@CAT@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>@VAL@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:val><c:smooth val="0"/><c:extLst><c:ext uri="{UID}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{UID}"/></c:ext></c:extLst></c:ser>)X";
constexpr const char* kSerArea =
    R"X(<c:ser><c:idx val="@I@"/><c:order val="@I@"/><c:tx><c:strRef><c:f>@NAME@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill><a:ln@W@><a:noFill/></a:ln><a:effectLst/></c:spPr><c:cat><c:strRef><c:f>@CAT@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>@VAL@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:val><c:extLst><c:ext uri="{UID}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{UID}"/></c:ext></c:extLst></c:ser>)X";
constexpr const char* kSerSc =
    R"X(<c:ser><c:idx val="@I@"/><c:order val="@I@"/><c:tx><c:strRef><c:f>@NAME@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:tx><c:spPr><a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln><a:effectLst/></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:marker><c:xVal><c:numRef><c:f>@CAT@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>@VAL@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:yVal><c:smooth val="0"/><c:extLst><c:ext uri="{UID}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{UID}"/></c:ext></c:extLst></c:ser>)X";
constexpr const char* kDpt =
    R"X(<c:dPt><c:idx val="@P@"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent@A@"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>)X";
constexpr const char* kSerPie =
    R"X(<c:ser><c:idx val="@I@"/><c:order val="@I@"/><c:tx><c:strRef><c:f>@NAME@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:tx>@DPTS@<c:cat><c:strRef><c:f>@CAT@</c:f><c:strCache>{CACHE}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>@VAL@</c:f><c:numCache>{CACHE}</c:numCache></c:numRef></c:val><c:extLst><c:ext uri="{UID}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{UID}"/></c:ext></c:extLst></c:ser>)X";
constexpr const char* kDlbls =
    R"X(<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>)X";
constexpr const char* kDlblsPie =
    R"X(<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>)X";
constexpr const char* kCatax =
    R"X(<c:catAx><c:axId val="@X1@"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="@POS@"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="@TICK@"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr><c:crossAx val="@X2@"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>)X";
constexpr const char* kValax =
    R"X(<c:valAx><c:axId val="@X2@"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="@POS@"/><c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines><c:numFmt formatCode="@FMT@" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr><c:crossAx val="@X1@"/><c:crosses val="autoZero"/><c:crossBetween val="@CB@"/></c:valAx>)X";
constexpr const char* kScax =
    R"X(<c:valAx><c:axId val="@X1@"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="@POS@"/><c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr><c:crossAx val="@X2@"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>)X";
constexpr const char* kTail =
    R"X(<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{UID}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="@L@"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>)X";

// ---------------------------------------------------------------------------
// Measured split table
// ---------------------------------------------------------------------------

struct SplitRow {
  const char* shape;
  const char* grid;  ///< Rows split by '/', cells by ','; U+00B7 is an empty cell.
  D dir;
  bool header_row;
  bool header_column;
  const char* plain;    ///< name/cat/val per series, "; "-separated.
  const char* scatter;  ///< name/xVal/yVal per series.
};

const SplitRow kSplitTable[] = {
    {"std", "·,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"std", "·,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"std", "·,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"wide", "·,S1,S2,S3,S4/a,1,2,3,4/b,5,6,7,8", D::kAuto, true, true, "A2/B1:E1/B2:E2; A3/B1:E1/B3:E3",
     "A2/B1:E1/B2:E2; A3/B1:E1/B3:E3"},
    {"wide", "·,S1,S2,S3,S4/a,1,2,3,4/b,5,6,7,8", D::kColumns, true, true,
     "B1/A2:A3/B2:B3; C1/A2:A3/C2:C3; D1/A2:A3/D2:D3; E1/A2:A3/E2:E3",
     "B1/A2:A3/B2:B3; C1/A2:A3/C2:C3; D1/A2:A3/D2:D3; E1/A2:A3/E2:E3"},
    {"wide", "·,S1,S2,S3,S4/a,1,2,3,4/b,5,6,7,8", D::kRows, true, true, "A2/B1:E1/B2:E2; A3/B1:E1/B3:E3",
     "A2/B1:E1/B2:E2; A3/B1:E1/B3:E3"},
    {"nohdr", "1,4/2,5/3,6/7,8", D::kAuto, false, false, "-/-/A1:A4; -/-/B1:B4", "-/A1:A4/B1:B4"},
    {"nohdr", "1,4/2,5/3,6/7,8", D::kColumns, false, false, "-/-/A1:A4; -/-/B1:B4", "-/A1:A4/B1:B4"},
    {"nohdr", "1,4/2,5/3,6/7,8", D::kRows, false, false, "-/-/A1:B1; -/-/A2:B2; -/-/A3:B3; -/-/A4:B4",
     "-/A1:B1/A2:B2; -/A1:B1/A3:B3; -/A1:B1/A4:B4"},
    {"numhdr", "·,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numhdr", "·,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numhdr", "·,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"a1text", "Item,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"a1text", "Item,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"a1text", "Item,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"numcat", "·,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numcat", "·,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numcat", "·,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"onecol", "S1/1/2/3/4", D::kAuto, true, false, "A1/-/A2:A5", "A1/-/A2:A5"},
    {"onecol", "S1/1/2/3/4", D::kColumns, true, false, "A1/-/A2:A5", "A1/-/A2:A5"},
    {"onecol", "S1/1/2/3/4", D::kRows, true, false, "-/A1/A2; -/A1/A3; -/A1/A4; -/A1/A5",
     "-/A1/A2; -/A1/A3; -/A1/A4; -/A1/A5"},
    {"square", "·,S1,S2,S3/a,1,2,3/b,4,5,6/c,7,8,9", D::kAuto, true, true,
     "A2/B1:D1/B2:D2; A3/B1:D1/B3:D3; A4/B1:D1/B4:D4", "A2/B1:D1/B2:D2; A3/B1:D1/B3:D3; A4/B1:D1/B4:D4"},
    {"square", "·,S1,S2,S3/a,1,2,3/b,4,5,6/c,7,8,9", D::kColumns, true, true,
     "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4; D1/A2:A4/D2:D4", "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4; D1/A2:A4/D2:D4"},
    {"square", "·,S1,S2,S3/a,1,2,3/b,4,5,6/c,7,8,9", D::kRows, true, true,
     "A2/B1:D1/B2:D2; A3/B1:D1/B3:D3; A4/B1:D1/B4:D4", "A2/B1:D1/B2:D2; A3/B1:D1/B3:D3; A4/B1:D1/B4:D4"},
    {"hdronly", "H1,H2/1,4/2,5/3,6/4,8", D::kAuto, true, false, "A1/-/A2:A5; B1/-/B2:B5", "B1/A2:A5/B2:B5"},
    {"hdronly", "H1,H2/1,4/2,5/3,6/4,8", D::kColumns, true, false, "A1/-/A2:A5; B1/-/B2:B5", "B1/A2:A5/B2:B5"},
    {"hdronly", "H1,H2/1,4/2,5/3,6/4,8", D::kRows, true, false,
     "-/A1:B1/A2:B2; -/A1:B1/A3:B3; -/A1:B1/A4:B4; -/A1:B1/A5:B5",
     "-/A1:B1/A2:B2; -/A1:B1/A3:B3; -/A1:B1/A4:B4; -/A1:B1/A5:B5"},
    {"datecat", "·,S1/2024-01-01,1/2024-01-02,2/2024-01-03,3", D::kAuto, true, true, "B1/A2:A4/B2:B4",
     "B1/A2:A4/B2:B4"},
    {"datecat", "·,S1/2024-01-01,1/2024-01-02,2/2024-01-03,3", D::kColumns, true, true, "B1/A2:A4/B2:B4",
     "B1/A2:A4/B2:B4"},
    {"datecat", "·,S1/2024-01-01,1/2024-01-02,2/2024-01-03,3", D::kRows, true, true, "A2/B1/B2; A3/B1/B3; A4/B1/B4",
     "A2/B1/B2; A3/B1/B3; A4/B1/B4"},
    {"numcat_a1", "X,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kAuto, true, false, "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numcat_a1", "X,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kColumns, true, false, "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"numcat_a1", "X,S1,S2/10,1,4/20,2,5/30,3,6/40,7,8", D::kRows, true, false,
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5",
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5"},
    {"textcol_nohdr", "a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, false, true, "-/A1:A4/B1:B4; -/A1:A4/C1:C4",
     "-/A1:A4/B1:B4; -/A1:A4/C1:C4"},
    {"textcol_nohdr", "a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, false, true, "-/A1:A4/B1:B4; -/A1:A4/C1:C4",
     "-/A1:A4/B1:B4; -/A1:A4/C1:C4"},
    {"textcol_nohdr", "a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, false, true,
     "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3; A4/-/B4:C4", "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4"},
    {"h_mixrow", "Item,S1,5/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, false, true, "-/A1:A5/B1:B5; -/A1:A5/C1:C5",
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"h_mixrow", "Item,S1,5/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, false, true, "-/A1:A5/B1:B5; -/A1:A5/C1:C5",
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"h_mixrow", "Item,S1,5/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, false, true,
     "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3; A4/-/B4:C4; A5/-/B5:C5",
     "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3; A4/-/B4:C4; A5/-/B5:C5"},
    {"h_mixcol", "Item,S1,S2/a,1,4/2,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_mixcol", "Item,S1,S2/a,1,4/2,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_mixcol", "Item,S1,S2/a,1,4/2,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_a1num", "1,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_a1num", "1,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_a1num", "1,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_a1txt_numrow", "Item,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, false, true, "-/A1:A5/B1:B5; -/A1:A5/C1:C5",
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"h_a1txt_numrow", "Item,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, false, true,
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5", "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"h_a1txt_numrow", "Item,2020,2021/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, false, true,
     "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3; A4/-/B4:C4; A5/-/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_blank_num", "·,1,2/3,1,4/5,2,5/6,3,6/9,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_blank_num", "·,1,2/3,1,4/5,2,5/6,3,6/9,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_blank_num", "·,1,2/3,1,4/5,2,5/6,3,6/9,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_datecol", "Date,S1,S2/2024-01-01,1,4/2024-01-02,2,5/2024-01-03,3,6/2024-01-04,7,8", D::kAuto, true, true,
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_datecol", "Date,S1,S2/2024-01-01,1,4/2024-01-02,2,5/2024-01-03,3,6/2024-01-04,7,8", D::kColumns, true, true,
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_datecol", "Date,S1,S2/2024-01-01,1,4/2024-01-02,2,5/2024-01-03,3,6/2024-01-04,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_bodytext", "Item,S1,S2/a,1,4/b,x,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_bodytext", "Item,S1,S2/a,1,4/b,x,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_bodytext", "Item,S1,S2/a,1,4/b,x,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"h_textcol_sq", "a,1,4/b,2,5/c,3,6", D::kAuto, false, true, "-/A1:A3/B1:B3; -/A1:A3/C1:C3",
     "-/A1:A3/B1:B3; -/A1:A3/C1:C3"},
    {"h_textcol_sq", "a,1,4/b,2,5/c,3,6", D::kColumns, false, true, "-/A1:A3/B1:B3; -/A1:A3/C1:C3",
     "-/A1:A3/B1:B3; -/A1:A3/C1:C3"},
    {"h_textcol_sq", "a,1,4/b,2,5/c,3,6", D::kRows, false, true, "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3"},
    {"h_hdr_sq", "H1,H2,H3/1,2,3/4,5,6", D::kAuto, true, false, "-/A1:C1/A2:C2; -/A1:C1/A3:C3",
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3"},
    {"h_hdr_sq", "H1,H2,H3/1,2,3/4,5,6", D::kColumns, true, false, "A1/-/A2:A3; B1/-/B2:B3; C1/-/C2:C3",
     "B1/A2:A3/B2:B3; C1/A2:A3/C2:C3"},
    {"h_hdr_sq", "H1,H2,H3/1,2,3/4,5,6", D::kRows, true, false, "-/A1:C1/A2:C2; -/A1:C1/A3:C3",
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3"},
    {"h_hdrblank", "Item,·,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_hdrblank", "Item,·,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"h_hdrblank", "Item,·,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A3/B1:C2/B3:C3; A4/B1:C2/B4:C4; A5/B1:C2/B5:C5"},
    {"h_blankcat", "·,S1,S2/·,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1:B2/A3:A5/B3:B5; C1:C2/A3:A5/C3:C5",
     "B1:B2/A3:A5/B3:B5; C1:C2/A3:A5/C3:C5"},
    {"h_blankcat", "·,S1,S2/·,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1:B2/A3:A5/B3:B5; C1:C2/A3:A5/C3:C5",
     "B1:B2/A3:A5/B3:B5; C1:C2/A3:A5/C3:C5"},
    {"h_blankcat", "·,S1,S2/·,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A3/B1:C2/B3:C3; A4/B1:C2/B4:C4; A5/B1:C2/B5:C5", "A3/B1:C2/B3:C3; A4/B1:C2/B4:C4; A5/B1:C2/B5:C5"},
    {"h_tall_nohdr_rows", "1,2,3,4,5/6,7,8,9,10", D::kAuto, false, false, "-/-/A1:E1; -/-/A2:E2", "-/A1:E1/A2:E2"},
    {"h_tall_nohdr_rows", "1,2,3,4,5/6,7,8,9,10", D::kColumns, false, false,
     "-/-/A1:A2; -/-/B1:B2; -/-/C1:C2; -/-/D1:D2; -/-/E1:E2",
     "-/A1:A2/B1:B2; -/A1:A2/C1:C2; -/A1:A2/D1:D2; -/A1:A2/E1:E2"},
    {"h_tall_nohdr_rows", "1,2,3,4,5/6,7,8,9,10", D::kRows, false, false, "-/-/A1:E1; -/-/A2:E2", "-/A1:E1/A2:E2"},
    {"c_onetext_mid", "Item,S1,S2/1,1,4/b,2,5/3,3,6/4,7,8", D::kAuto, true, false, "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"c_onetext_mid", "Item,S1,S2/1,1,4/b,2,5/3,3,6/4,7,8", D::kColumns, true, false,
     "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"c_onetext_mid", "Item,S1,S2/1,1,4/b,2,5/3,3,6/4,7,8", D::kRows, true, false,
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5",
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5"},
    {"c_onetext_last", "Item,S1,S2/1,1,4/2,2,5/3,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "C1/A2:B5/C2:C5"},
    {"c_onetext_last", "Item,S1,S2/1,1,4/2,2,5/3,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "C1/A2:B5/C2:C5"},
    {"c_onetext_last", "Item,S1,S2/1,1,4/2,2,5/3,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"r_firstnum", "Item,5,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"r_firstnum", "Item,5,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true, "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5",
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"r_firstnum", "Item,5,S2/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A3/B1:C2/B3:C3; A4/B1:C2/B4:C4; A5/B1:C2/B5:C5"},
    {"r_dates", "Item,2024-01-01,2024-02-01/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"r_dates", "Item,2024-01-01,2024-02-01/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, true, true,
     "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"r_dates", "Item,2024-01-01,2024-02-01/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"r_bool", "Item,TRUE,FALSE/a,1,4/b,2,5/c,3,6/d,7,8", D::kAuto, false, true, "-/A1:A5/B1:B5; -/A1:A5/C1:C5",
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"r_bool", "Item,TRUE,FALSE/a,1,4/b,2,5/c,3,6/d,7,8", D::kColumns, false, true, "-/A1:A5/B1:B5; -/A1:A5/C1:C5",
     "-/A1:A5/B1:B5; -/A1:A5/C1:C5"},
    {"r_bool", "Item,TRUE,FALSE/a,1,4/b,2,5/c,3,6/d,7,8", D::kRows, false, true,
     "A1/-/B1:C1; A2/-/B2:C2; A3/-/B3:C3; A4/-/B4:C4; A5/-/B5:C5",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4; A5/B1:C1/B5:C5"},
    {"c_bool", "Item,S1,S2/TRUE,1,4/FALSE,2,5/TRUE,3,6/FALSE,7,8", D::kAuto, true, false,
     "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"c_bool", "Item,S1,S2/TRUE,1,4/FALSE,2,5/TRUE,3,6/FALSE,7,8", D::kColumns, true, false,
     "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5", "B1/A2:A5/B2:B5; C1/A2:A5/C2:C5"},
    {"c_bool", "Item,S1,S2/TRUE,1,4/FALSE,2,5/TRUE,3,6/FALSE,7,8", D::kRows, true, false,
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5",
     "-/A1:C1/A2:C2; -/A1:C1/A3:C3; -/A1:C1/A4:C4; -/A1:C1/A5:C5"},
    {"c_alltext_rowsnum", "a,b,c/d,1,2/e,3,4", D::kAuto, true, true, "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3"},
    {"c_alltext_rowsnum", "a,b,c/d,1,2/e,3,4", D::kColumns, true, true, "B1/A2:A3/B2:B3; C1/A2:A3/C2:C3",
     "B1/A2:A3/B2:B3; C1/A2:A3/C2:C3"},
    {"c_alltext_rowsnum", "a,b,c/d,1,2/e,3,4", D::kRows, true, true, "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3",
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3"},
    {"blank2x2", "·,·,S1,S2/·,·,T1,T2/g1,a,1,4/·,b,2,5/g2,c,3,6/·,d,7,8", D::kAuto, true, true,
     "C1:C2/A3:B6/C3:C6; D1:D2/A3:B6/D3:D6", "C1:C2/A3:B6/C3:C6; D1:D2/A3:B6/D3:D6"},
    {"blank2x2", "·,·,S1,S2/·,·,T1,T2/g1,a,1,4/·,b,2,5/g2,c,3,6/·,d,7,8", D::kColumns, true, true,
     "C1:C2/A3:B6/C3:C6; D1:D2/A3:B6/D3:D6", "C1:C2/A3:B6/C3:C6; D1:D2/A3:B6/D3:D6"},
    {"blank2x2", "·,·,S1,S2/·,·,T1,T2/g1,a,1,4/·,b,2,5/g2,c,3,6/·,d,7,8", D::kRows, true, true,
     "A3:B3/C1:D2/C3:D3; A4:B4/C1:D2/C4:D4; A5:B5/C1:D2/C5:D5; A6:B6/C1:D2/C6:D6",
     "A3:B3/C1:D2/C3:D3; A4:B4/C1:D2/C4:D4; A5:B5/C1:D2/C5:D5; A6:B6/C1:D2/C6:D6"},
    {"m_mixedcol", "Cat,Mixed/a,1/b,·/c,3/d,#NA/e,txt/f,TRUE/g,#DIV/h,8", D::kAuto, true, true, "B1:B6/A7:A9/B7:B9",
     "B1:B6/A7:A9/B7:B9"},
    {"m_mixedcol", "Cat,Mixed/a,1/b,·/c,3/d,#NA/e,txt/f,TRUE/g,#DIV/h,8", D::kColumns, true, true, "B1:B6/A7:A9/B7:B9",
     "B1:B6/A7:A9/B7:B9"},
    {"m_mixedcol", "Cat,Mixed/a,1/b,·/c,3/d,#NA/e,txt/f,TRUE/g,#DIV/h,8", D::kRows, true, true,
     "A7/B1:B6/B7; A8/B1:B6/B8; A9/B1:B6/B9", "A7/B1:B6/B7; A8/B1:B6/B8; A9/B1:B6/B9"},
    {"m_lastcol_text", "Item,S1,S2/a,1,4/b,x,5/c,3,y/d,7,8", D::kAuto, true, true, "A5/B1:C4/B5:C5", "A5/B1:C4/B5:C5"},
    {"m_lastcol_text", "Item,S1,S2/a,1,4/b,x,5/c,3,y/d,7,8", D::kColumns, true, true, "B1:B4/A5/B5; C1:C4/A5/C5",
     "B1:B4/A5/B5; C1:C4/A5/C5"},
    {"m_lastcol_text", "Item,S1,S2/a,1,4/b,x,5/c,3,y/d,7,8", D::kRows, true, true, "A5/B1:C4/B5:C5", "A5/B1:C4/B5:C5"},
    {"m_bottom_text", "Item,S1,S2/a,1,4/b,2,5/d,z,8", D::kAuto, true, true, "C1/A2:B4/C2:C4", "C1/A2:B4/C2:C4"},
    {"m_bottom_text", "Item,S1,S2/a,1,4/b,2,5/d,z,8", D::kColumns, true, true, "C1/A2:B4/C2:C4", "C1/A2:B4/C2:C4"},
    {"m_bottom_text", "Item,S1,S2/a,1,4/b,2,5/d,z,8", D::kRows, true, true, "A2:B2/C1/C2; A3:B3/C1/C3; A4:B4/C1/C4",
     "A2:B2/C1/C2; A3:B3/C1/C3; A4:B4/C1/C4"},
    {"m_corner_body_text", "·,S1,S2/a,1,4/b,2,y/c,3,6", D::kAuto, true, true, "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4",
     "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4"},
    {"m_corner_body_text", "·,S1,S2/a,1,4/b,2,y/c,3,6", D::kColumns, true, true, "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4",
     "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4"},
    {"m_corner_body_text", "·,S1,S2/a,1,4/b,2,y/c,3,6", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4", "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4"},
    {"m_corner_bottom_text", "·,S1,S2/a,1,4/b,2,5/c,z,6", D::kAuto, true, true, "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4",
     "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4"},
    {"m_corner_bottom_text", "·,S1,S2/a,1,4/b,2,5/c,z,6", D::kColumns, true, true, "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4",
     "B1/A2:A4/B2:B4; C1/A2:A4/C2:C4"},
    {"m_corner_bottom_text", "·,S1,S2/a,1,4/b,2,5/c,z,6", D::kRows, true, true,
     "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4", "A2/B1:C1/B2:C2; A3/B1:C1/B3:C3; A4/B1:C1/B4:C4"}};

std::vector<std::string> split_on(const std::string& s, char sep) {
  std::vector<std::string> out;
  std::size_t start = 0;
  for (;;) {
    const std::size_t at = s.find(sep, start);
    out.push_back(s.substr(start, at == std::string::npos ? std::string::npos : at - start));
    if (at == std::string::npos) {
      return out;
    }
    start = at + 1U;
  }
}

std::int64_t days_from_civil(std::int64_t y, unsigned m, unsigned d) {
  y -= m <= 2U ? 1 : 0;
  const std::int64_t era = (y >= 0 ? y : y - 399) / 400;
  const unsigned yoe = static_cast<unsigned>(y - era * 400);
  const unsigned doy = (153U * (m + (m > 2U ? -3U : 9U)) + 2U) / 5U + d - 1U;
  const unsigned doe = yoe * 365U + yoe / 4U - yoe / 100U + doy;
  return era * 146097 + static_cast<std::int64_t>(doe) - 719468;
}

std::uint32_t date_xf(Workbook& wb) {
  StylesTable& st = wb.mutable_styles();
  CellXf xf;
  xf.num_fmt_id = 14;
  xf.apply_number_format = true;
  st.cell_xfs.push_back(xf);
  return static_cast<std::uint32_t>(st.cell_xfs.size() - 1U);
}

/// Writes `grid` at A1 of `sheet`; returns the A1 rectangle it covers.
std::string fill_grid(Workbook& wb, std::size_t sheet, const std::string& grid) {
  const std::uint32_t xf = date_xf(wb);
  const std::vector<std::string> rows = split_on(grid, '/');
  std::uint32_t cols = 0;
  for (std::uint32_t r = 0; r < rows.size(); ++r) {
    const std::vector<std::string> cells = split_on(rows[r], ',');
    cols = std::max<std::uint32_t>(cols, static_cast<std::uint32_t>(cells.size()));
    for (std::uint32_t c = 0; c < cells.size(); ++c) {
      const std::string& t = cells[c];
      Sheet& s = wb.sheet(sheet);
      if (t == "\xC2\xB7") {
        continue;
      }
      if (t == "TRUE" || t == "FALSE") {
        s.set_cell_value(r, c, Value::boolean(t == "TRUE"));
      } else if (t == "#NA") {
        s.set_cell_value(r, c, Value::error(ErrorCode::NA));
      } else if (t == "#DIV") {
        s.set_cell_value(r, c, Value::error(ErrorCode::Div0));
      } else if (t.size() == 10U && t[4] == '-' && t[7] == '-') {
        const std::int64_t days =
            days_from_civil(std::stoi(t.substr(0, 4)), static_cast<unsigned>(std::stoi(t.substr(5, 2))),
                            static_cast<unsigned>(std::stoi(t.substr(8, 2))));
        s.set_cell_value(r, c, Value::number(static_cast<double>(days + 25569)));
        EXPECT_TRUE(static_cast<bool>(wb.set_cell_xf_index(sheet, r, c, xf)));
      } else if (t.find_first_not_of("0123456789") == std::string::npos) {
        s.set_cell_value(r, c, Value::number(std::stod(t)));
      } else {
        s.set_cell_text(r, c, t);
      }
    }
  }
  return "A1:" + a1::encode_a1(static_cast<std::uint32_t>(rows.size() - 1U), cols - 1U);
}

std::string show(const ChartSplit& sp, bool scatter) {
  const auto bare = [](std::string s) {
    for (const std::string_view key : {std::string_view("Sheet1!"), std::string_view("$")}) {
      for (std::size_t at; (at = s.find(key)) != std::string::npos;) {
        s.erase(at, key.size());
      }
    }
    return s.empty() ? std::string("-") : s;
  };
  std::string out;
  for (const ChartSeriesRefs& s : sp.series) {
    out += out.empty() ? "" : "; ";
    out += bare(s.name) + '/' + bare(scatter ? s.x : s.cat) + '/' + bare(scatter ? s.y : s.val);
  }
  return out;
}

Expected<ChartSplit, Error> split(const Workbook& wb, const std::string& src,
                                  ChartInsertKind kind = ChartInsertKind::kColumn, D dir = D::kAuto,
                                  ChartHeaderMode hr = ChartHeaderMode::kAuto,
                                  ChartHeaderMode hc = ChartHeaderMode::kAuto) {
  return split_chart_source(wb, 0, src, kind, dir, hr, hc);
}

constexpr const char* kStdGrid = "\xC2\xB7,S1,S2/a,1,4/b,2,5/c,3,6/d,7,8";

TEST(ChartTemplateSplit, MeasuredTable) {
  for (const SplitRow& row : kSplitTable) {
    Workbook wb = Workbook::create();
    const std::string rect = fill_grid(wb, 0, row.grid);
    for (const ChartInsertKind kind : {ChartInsertKind::kColumn, ChartInsertKind::kBar, ChartInsertKind::kLine,
                                       ChartInsertKind::kArea, ChartInsertKind::kPie, ChartInsertKind::kScatter}) {
      const bool scatter = kind == ChartInsertKind::kScatter;
      SCOPED_TRACE(std::string(row.shape) + " dir=" + std::to_string(static_cast<int>(row.dir)) +
                   " kind=" + std::to_string(static_cast<int>(kind)));
      const auto sp = split(wb, rect, kind, row.dir);
      ASSERT_TRUE(static_cast<bool>(sp));
      EXPECT_EQ(show(sp.value(), scatter), scatter ? row.scatter : row.plain);
      EXPECT_EQ(sp.value().header_row, row.header_row);
      EXPECT_EQ(sp.value().header_column, row.header_column);
      EXPECT_NE(sp.value().direction, D::kAuto);
      if (row.dir != D::kAuto) {
        EXPECT_EQ(sp.value().direction, row.dir);
      }
    }
  }
}

TEST(ChartTemplateSplit, HeaderOverrides) {
  Workbook wb = Workbook::create();
  const std::string std_rect = fill_grid(wb, 0, kStdGrid);
  using M = ChartHeaderMode;
  // The corner blank is read as data once the header row is refused.
  auto sp = split(wb, std_rect, ChartInsertKind::kColumn, D::kAuto, M::kNo, M::kAuto);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(show(sp.value(), false), "-/A1:A5/B1:B5; -/A1:A5/C1:C5");
  EXPECT_FALSE(sp.value().header_row);
  EXPECT_TRUE(sp.value().header_column);

  sp = split(wb, std_rect, ChartInsertKind::kColumn, D::kAuto, M::kAuto, M::kNo);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(show(sp.value(), false), "A1/-/A2:A5; B1/-/B2:B5; C1/-/C2:C5");
  EXPECT_FALSE(sp.value().header_column);

  Workbook plain = Workbook::create();
  const std::string rect = fill_grid(plain, 0, "1,4/2,5/3,6/7,8");
  sp = split(plain, rect, ChartInsertKind::kColumn, D::kAuto, M::kYes, M::kAuto);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(show(sp.value(), false), "A1/-/A2:A4; B1/-/B2:B4");
  EXPECT_TRUE(sp.value().header_row);

  sp = split(plain, rect, ChartInsertKind::kColumn, D::kAuto, M::kAuto, M::kYes);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(show(sp.value(), false), "-/A1:A4/B1:B4");
  EXPECT_TRUE(sp.value().header_column);

  // A forced header leaves no data cell.
  Workbook one = Workbook::create();
  const std::string row_rect = fill_grid(one, 0, "1,2,3");
  EXPECT_FALSE(static_cast<bool>(split(one, row_rect, ChartInsertKind::kColumn, D::kAuto, M::kYes, M::kAuto)));
}

TEST(ChartTemplateSplit, QualifiedAndWholeColumnSources) {
  Workbook wb = Workbook::create();
  const std::size_t other = wb.add_sheet("My Sheet");
  fill_grid(wb, other, kStdGrid);
  auto sp = split_chart_source(wb, 0, "'My Sheet'!A1:C5", ChartInsertKind::kColumn, D::kColumns, ChartHeaderMode::kAuto,
                               ChartHeaderMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(sp.value().series[0].name, "'My Sheet'!$B$1");
  EXPECT_EQ(sp.value().series[0].cat, "'My Sheet'!$A$2:$A$5");
  EXPECT_EQ(sp.value().series[1].val, "'My Sheet'!$C$2:$C$5");

  sp = split_chart_source(wb, other, "A:C", ChartInsertKind::kColumn, D::kColumns, ChartHeaderMode::kAuto,
                          ChartHeaderMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(sp));
  EXPECT_EQ(sp.value().series[0].val, "'My Sheet'!$B$2:$B$5");
}

TEST(ChartTemplateSplit, RejectsBadSources) {
  Workbook wb = Workbook::create();
  fill_grid(wb, 0, kStdGrid);
  for (const char* bad : {"", "A1:B", "Nope!A1:B2", "[1]Sheet1!A1:B2", "A1,B2"}) {
    const auto sp = split(wb, bad);
    ASSERT_FALSE(static_cast<bool>(sp)) << bad;
    EXPECT_EQ(sp.error().code, FormulonErrorCode::kIoChartInvalidReference) << bad;
  }
  EXPECT_FALSE(static_cast<bool>(split_chart_source(wb, kChartNoHostSheet, "A1:C5", ChartInsertKind::kColumn, D::kAuto,
                                                    ChartHeaderMode::kAuto, ChartHeaderMode::kAuto)));
  Workbook empty = Workbook::create();
  EXPECT_FALSE(static_cast<bool>(split(empty, "A:B")));
}

// ---------------------------------------------------------------------------
// XML generation
// ---------------------------------------------------------------------------

void replace_all(std::string& s, const std::string& from, const std::string& to) {
  for (std::size_t at = 0; (at = s.find(from, at)) != std::string::npos; at += to.size()) {
    s.replace(at, from.size(), to);
  }
}

std::string fill(std::string s, std::initializer_list<std::pair<const char*, std::string>> kv) {
  for (const auto& [key, value] : kv) {
    replace_all(s, std::string("@") + key + "@", value);
  }
  return s;
}

std::string swap_axis_ids(std::string s) {
  replace_all(s, "@X1@", "@T@");
  replace_all(s, "@X2@", "@X1@");
  replace_all(s, "@T@", "@X2@");
  return s;
}

enum class TitleForm { kAuto, kDeleted };

/// The measured part for a two-series chart over the fixture ranges.
std::string expected_part(ChartInsertKind kind, ChartGrouping grouping, TitleForm title, const std::string& lang) {
  const bool scatter = kind == ChartInsertKind::kScatter;
  const bool ring = kind == ChartInsertKind::kDoughnut;
  const bool pie = kind == ChartInsertKind::kPie || ring;
  const bool bar = kind == ChartInsertKind::kBar;
  const bool column = kind == ChartInsertKind::kColumn;
  const bool line = kind == ChartInsertKind::kLine;
  const bool area = kind == ChartInsertKind::kArea;
  const bool percent = grouping == ChartGrouping::kPercentStacked;
  const bool stacked = grouping == ChartGrouping::kStacked || percent;
  std::string x = fill(kHead, {{"L", lang}});
  x += title == TitleForm::kAuto ? fill(kTitle, {{"L", lang}}) : "";
  x += std::string("<c:autoTitleDeleted val=\"") + (title == TitleForm::kAuto ? "0" : "1") +
       "\"/><c:plotArea><c:layout/>";
  const char* plot = ring      ? "doughnutChart"
                     : pie     ? "pieChart"
                     : scatter ? "scatterChart"
                     : line    ? "lineChart"
                     : area    ? "areaChart"
                               : "barChart";
  x += std::string("<c:") + plot + ">";
  x += column ? "<c:barDir val=\"col\"/>" : bar ? "<c:barDir val=\"bar\"/>" : "";
  if (column || bar || line || area) {
    x += std::string("<c:grouping val=\"") +
         (percent          ? "percentStacked"
          : stacked        ? "stacked"
          : (line || area) ? "standard"
                           : "clustered") +
         "\"/>";
  }
  x += scatter ? "<c:scatterStyle val=\"lineMarker\"/>" : "";
  x += pie ? "<c:varyColors val=\"1\"/>" : "<c:varyColors val=\"0\"/>";
  for (int i = 0; i < 2; ++i) {
    const std::string col(1, scatter ? "FG"[i] : "BC"[i]);
    const std::string name = "Sheet1!$" + col + "$1";
    const std::string cat = scatter ? "Sheet1!$E$2:$E$5" : "Sheet1!$A$2:$A$5";
    const std::string val = "Sheet1!$" + col + "$2:$" + col + "$5";
    std::string s = pie ? kSerPie : scatter ? kSerSc : line ? kSerLine : area ? kSerArea : kSerBar;
    std::string dpts;
    for (int p = 0; p < 4; ++p) {
      dpts += fill(kDpt, {{"P", std::to_string(p)}, {"A", std::to_string(p + 1)}});
    }
    replace_all(s, "@W@", area && i > 0 ? " w=\"25400\"" : "");
    replace_all(s, "@DPTS@", dpts);
    x += fill(s, {{"I", std::to_string(i)}, {"A", std::to_string(i + 1)}, {"NAME", name}, {"CAT", cat}, {"VAL", val}});
  }
  x += pie ? kDlblsPie : kDlbls;
  x += column ? "<c:gapWidth val=\"219\"/>" : bar ? "<c:gapWidth val=\"182\"/>" : "";
  x += (column || bar) && stacked ? "<c:overlap val=\"100\"/>" : "";
  x += line ? "<c:smooth val=\"0\"/>" : "";
  x += ring  ? "<c:firstSliceAng val=\"0\"/><c:holeSize val=\"50\"/>"
       : pie ? "<c:firstSliceAng val=\"0\"/>"
             : "<c:axId val=\"@X1@\"/><c:axId val=\"@X2@\"/>";
  x += std::string("</c:") + plot + ">";
  if (scatter) {
    x += fill(kScax, {{"POS", "b"}, {"L", lang}});
    x += swap_axis_ids(fill(kScax, {{"POS", "l"}, {"L", lang}}));
  } else if (!pie) {
    x += fill(kCatax, {{"POS", bar ? "l" : "b"}, {"TICK", area ? "out" : "none"}, {"L", lang}});
    x += fill(kValax, {{"POS", bar ? "b" : "l"},
                       {"FMT", percent ? "0%" : "General"},
                       {"CB", area ? "midCat" : "between"},
                       {"L", lang}});
  }
  x += fill(kTail, {{"L", lang}});
  return x;
}

/// Replaces generated ids and caches by the placeholders of the fragments.
std::string normalise(const std::string& xml) {
  std::string out = std::regex_replace(xml, std::regex("<c:(str|num|multiLvlStr)Cache>[\\s\\S]*?</c:\\1Cache>"),
                                       "<c:$1Cache>{CACHE}</c:$1Cache>");
  out = std::regex_replace(out, std::regex("\\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\\}"), "{UID}");
  std::vector<std::string> ids;
  std::string mapped;
  const std::regex axis("<c:(axId|crossAx) val=\"(\\d+)\"/>");
  auto last = out.cbegin();
  for (std::sregex_iterator it(out.cbegin(), out.cend(), axis), end; it != end; ++it) {
    const std::string id = (*it)[2];
    if (std::find(ids.begin(), ids.end(), id) == ids.end()) {
      ids.push_back(id);
    }
    mapped.append(last, (*it)[0].first);
    mapped += "<c:" + (*it)[1].str() + " val=\"@X" +
              std::to_string(std::find(ids.begin(), ids.end(), id) - ids.begin() + 1) + "@\"/>";
    last = (*it)[0].second;
  }
  mapped.append(last, out.cend());
  return mapped;
}

ChartSplit fixture_split(ChartInsertKind kind) {
  ChartSplit sp;
  const bool scatter = kind == ChartInsertKind::kScatter;
  for (int i = 0; i < 2; ++i) {
    const std::string col(1, scatter ? "FG"[i] : "BC"[i]);
    ChartSeriesRefs s;
    s.name = "Sheet1!$" + col + "$1";
    (scatter ? s.x : s.cat) = scatter ? "Sheet1!$E$2:$E$5" : "Sheet1!$A$2:$A$5";
    (scatter ? s.y : s.val) = "Sheet1!$" + col + "$2:$" + col + "$5";
    sp.series.push_back(std::move(s));
  }
  return sp;
}

std::string build(const Workbook& wb, const ChartInsertOptions& options) {
  const auto xml = build_chart_xml(wb, options, fixture_split(options.kind));
  EXPECT_TRUE(static_cast<bool>(xml));
  return xml ? xml.value() : std::string();
}

Workbook measured_workbook() {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

struct TemplateCase {
  ChartInsertKind kind;
  ChartGrouping grouping;
};

const TemplateCase kTemplateCases[] = {
    {ChartInsertKind::kColumn, ChartGrouping::kClustered},
    {ChartInsertKind::kColumn, ChartGrouping::kStacked},
    {ChartInsertKind::kColumn, ChartGrouping::kPercentStacked},
    {ChartInsertKind::kBar, ChartGrouping::kClustered},
    {ChartInsertKind::kBar, ChartGrouping::kStacked},
    {ChartInsertKind::kBar, ChartGrouping::kPercentStacked},
    {ChartInsertKind::kLine, ChartGrouping::kStandard},
    {ChartInsertKind::kLine, ChartGrouping::kStacked},
    {ChartInsertKind::kLine, ChartGrouping::kPercentStacked},
    {ChartInsertKind::kArea, ChartGrouping::kStandard},
    {ChartInsertKind::kArea, ChartGrouping::kStacked},
    {ChartInsertKind::kArea, ChartGrouping::kPercentStacked},
    {ChartInsertKind::kPie, ChartGrouping::kClustered},
    {ChartInsertKind::kScatter, ChartGrouping::kClustered},
    {ChartInsertKind::kDoughnut, ChartGrouping::kClustered},
};

TEST(ChartTemplateXml, MatchesMeasuredTemplates) {
  const Workbook wb = measured_workbook();
  for (const TemplateCase& tc : kTemplateCases) {
    SCOPED_TRACE(std::to_string(static_cast<int>(tc.kind)) + "/" + std::to_string(static_cast<int>(tc.grouping)));
    ChartInsertOptions options;
    options.kind = tc.kind;
    options.grouping = tc.grouping;
    options.legend = ChartLegendPosition::kBottom;
    const std::string xml = build(wb, options);
    EXPECT_EQ(normalise(xml), expected_part(tc.kind, tc.grouping, TitleForm::kAuto, "ja-JP"));
    pugi::xml_document doc;
    EXPECT_TRUE(doc.load_string(xml.c_str())) << "well-formed";
  }
}

TEST(ChartTemplateXml, LineAndAreaTakeEitherPlainGrouping) {
  const Workbook wb = measured_workbook();
  for (const ChartInsertKind kind : {ChartInsertKind::kLine, ChartInsertKind::kArea}) {
    ChartInsertOptions a;
    a.kind = kind;
    a.grouping = ChartGrouping::kClustered;
    ChartInsertOptions b = a;
    b.grouping = ChartGrouping::kStandard;
    a.legend = b.legend = ChartLegendPosition::kBottom;
    EXPECT_EQ(normalise(build(wb, a)), expected_part(kind, ChartGrouping::kStandard, TitleForm::kAuto, "ja-JP"));
    EXPECT_EQ(normalise(build(wb, a)), normalise(build(wb, b)));
  }
}

TEST(ChartTemplateXml, TitleForms) {
  const Workbook wb = measured_workbook();
  ChartInsertOptions options;
  options.legend = ChartLegendPosition::kBottom;

  options.title = std::string();
  EXPECT_EQ(normalise(build(wb, options)),
            expected_part(ChartInsertKind::kColumn, ChartGrouping::kClustered, TitleForm::kDeleted, "ja-JP"));

  options.title = std::string("Q1 & Q2");
  const std::string xml = build(wb, options);
  const std::size_t open = xml.find("<c:title>");
  const std::size_t close = xml.find("</c:title>");
  ASSERT_NE(open, std::string::npos);
  const std::string title = xml.substr(open, close - open);
  EXPECT_NE(title.find("<c:tx><c:rich>"), std::string::npos);
  EXPECT_NE(title.find("<a:t>Q1 &amp; Q2</a:t>"), std::string::npos);
  EXPECT_EQ(title.find("<c:txPr>"), std::string::npos);
  EXPECT_NE(xml.find("<c:autoTitleDeleted val=\"0\"/>"), std::string::npos);
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml.c_str()));
}

TEST(ChartTemplateXml, LegendDate1904AndLocale) {
  Workbook wb = Workbook::create();  // win-365-en_US
  wb.set_date1904(true);
  ChartInsertOptions options;
  options.legend = ChartLegendPosition::kBottom;
  std::string xml = build(wb, options);
  EXPECT_NE(xml.find("<c:date1904 val=\"1\"/><c:lang val=\"en-US\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<a:endParaRPr lang=\"en-US\"/>"), std::string::npos);
  EXPECT_EQ(xml.find("ja-JP"), std::string::npos);

  wb.set_date1904(false);
  EXPECT_EQ(normalise(build(wb, options)),
            expected_part(ChartInsertKind::kColumn, ChartGrouping::kClustered, TitleForm::kAuto, "en-US"));

  options.legend = ChartLegendPosition::kTopRight;
  EXPECT_NE(build(wb, options).find("<c:legendPos val=\"tr\"/>"), std::string::npos);
  options.legend = ChartLegendPosition::kNone;
  xml = build(wb, options);
  EXPECT_EQ(xml.find("<c:legend>"), std::string::npos);
  EXPECT_NE(xml.find("</c:plotArea><c:plotVisOnly"), std::string::npos);
}

TEST(ChartTemplateXml, ReferenceElementsFollowTheCells) {
  Workbook wb = measured_workbook();
  fill_grid(wb, 0, "\xC2\xB7,S1/10,1/20,2/30,3");
  ChartInsertOptions options;
  options.legend = ChartLegendPosition::kBottom;
  const auto sp = split(wb, "A1:B4");
  ASSERT_TRUE(static_cast<bool>(sp));
  const auto xml = build_chart_xml(wb, options, sp.value());
  ASSERT_TRUE(static_cast<bool>(xml));
  EXPECT_NE(xml.value().find("<c:cat><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>"), std::string::npos);

  Workbook multi = measured_workbook();
  const std::string rect = fill_grid(multi, 0, "\xC2\xB7,\xC2\xB7,S1/g,a,1/\xC2\xB7,b,2");
  const auto msp = split(multi, rect);
  ASSERT_TRUE(static_cast<bool>(msp));
  const auto mxml = build_chart_xml(multi, options, msp.value());
  ASSERT_TRUE(static_cast<bool>(mxml));
  EXPECT_NE(mxml.value().find("<c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$3</c:f><c:multiLvlStrCache>"),
            std::string::npos);

  Workbook text_x = measured_workbook();
  const std::string xr = fill_grid(text_x, 0, "\xC2\xB7,S1/a,1/b,2");
  options.kind = ChartInsertKind::kScatter;
  const auto xsp = split_chart_source(text_x, 0, xr, ChartInsertKind::kScatter, D::kAuto, ChartHeaderMode::kAuto,
                                      ChartHeaderMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(xsp));
  const auto xxml = build_chart_xml(text_x, options, xsp.value());
  ASSERT_TRUE(static_cast<bool>(xxml));
  EXPECT_NE(xxml.value().find("<c:xVal><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f>"), std::string::npos);
}

// Series past the sixth (and pie points) take the colour style's next variation, as Excel inserts them.
TEST(ChartTemplateXml, ColoursCycleThroughVariations) {
  pugi::xml_document style;
  ASSERT_TRUE(style.load_string(kDefaultChartColorStyle));
  const pugi::xml_node root = style.document_element();
  EXPECT_EQ(chart_cycle_color(root, 0), "<a:schemeClr val=\"accent1\"/>");
  EXPECT_EQ(chart_cycle_color(root, 6), "<a:schemeClr val=\"accent1\"><a:lumMod val=\"60000\"/></a:schemeClr>");
  EXPECT_EQ(chart_cycle_color(root, 13),
            "<a:schemeClr val=\"accent2\"><a:lumMod val=\"80000\"/><a:lumOff val=\"20000\"/></a:schemeClr>");
  EXPECT_EQ(chart_cycle_color(root, 54), "<a:schemeClr val=\"accent1\"/>");

  // Only colour elements are base colours; an extension list is not one.
  pugi::xml_document extended;
  ASSERT_TRUE(extended.load_string(
      "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\">"
      "<a:schemeClr val=\"accent1\"/><a:srgbClr val=\"112233\"/><cs:variation/><cs:extLst/></cs:colorStyle>"));
  EXPECT_EQ(chart_cycle_color(extended.document_element(), 2), "<a:schemeClr val=\"accent1\"/>");

  const Workbook wb = measured_workbook();
  ChartSplit sp;
  for (int i = 0; i < 8; ++i) {
    ChartSeriesRefs s;
    s.val = "Sheet1!$B$2:$B$11";
    sp.series.push_back(s);
  }
  for (const ChartInsertKind kind : {ChartInsertKind::kColumn, ChartInsertKind::kLine, ChartInsertKind::kScatter}) {
    ChartInsertOptions o;
    o.kind = kind;
    ChartSplit split = sp;
    for (ChartSeriesRefs& s : split.series) {
      s.y = s.val;
    }
    const auto xml = build_chart_xml(wb, o, split);
    ASSERT_TRUE(static_cast<bool>(xml));
    EXPECT_NE(xml.value().find("<a:schemeClr val=\"accent2\"><a:lumMod val=\"60000\"/></a:schemeClr>"),
              std::string::npos);
  }
  ChartInsertOptions pie;
  pie.kind = ChartInsertKind::kPie;
  const auto xml = build_chart_xml(wb, pie, ChartSplit{{sp.series[0]}});
  ASSERT_TRUE(static_cast<bool>(xml));
  EXPECT_NE(xml.value().find("<c:idx val=\"9\"/><c:bubble3D val=\"0\"/><c:spPr><a:solidFill><a:schemeClr "
                             "val=\"accent4\"><a:lumMod val=\"60000\"/></a:schemeClr>"),
            std::string::npos);
}

TEST(ChartTemplateXml, RejectsEmptySplit) {
  const Workbook wb = measured_workbook();
  const auto xml = build_chart_xml(wb, ChartInsertOptions{}, ChartSplit{});
  ASSERT_FALSE(static_cast<bool>(xml));
  EXPECT_EQ(xml.error().code, FormulonErrorCode::kIoChartInvalidReference);
}

// ---------------------------------------------------------------------------
// Source derivation
// ---------------------------------------------------------------------------

ChartDataSource ref_source(const std::string& f) {
  ChartDataSource d;
  if (!f.empty()) {
    d.kind = ChartDataSource::Kind::kRef;
    d.formula = f;
  }
  return d;
}

ChartModel model_of(const ChartSplit& sp, bool scatter) {
  ChartModel m;
  ChartPlot plot;
  plot.type = scatter ? ChartPlotType::kScatter : ChartPlotType::kBar;
  for (const ChartSeriesRefs& s : sp.series) {
    ChartSeries series;
    series.name = ref_source(s.name);
    series.cat = ref_source(s.cat);
    series.val = ref_source(s.val);
    series.x = ref_source(s.x);
    series.y = ref_source(s.y);
    plot.series.push_back(std::move(series));
  }
  m.plots.push_back(std::move(plot));
  return m;
}

TEST(ChartTemplateDerive, RoundTripsEverySplitOfTheTable) {
  for (const SplitRow& row : kSplitTable) {
    Workbook wb = Workbook::create();
    const std::string rect = fill_grid(wb, 0, row.grid);
    for (const bool scatter : {false, true}) {
      SCOPED_TRACE(std::string(row.shape) + (scatter ? " scatter" : " plain"));
      const auto sp = split(wb, rect, scatter ? ChartInsertKind::kScatter : ChartInsertKind::kColumn, row.dir);
      ASSERT_TRUE(static_cast<bool>(sp));
      const auto derived = derive_chart_source(wb, 0, model_of(sp.value(), scatter));
      ASSERT_TRUE(derived.has_value());
      const std::string corner = rect.substr(rect.find(':') + 1U);
      const std::size_t digits = corner.find_first_of("0123456789");
      const std::string want = "Sheet1!$A$1:$" + corner.substr(0, digits) + "$" + corner.substr(digits);
      EXPECT_EQ(derived->range, want);
      EXPECT_EQ(derived->direction, sp.value().direction);
      EXPECT_EQ(derived->header_row, sp.value().header_row);
      EXPECT_EQ(derived->header_column, sp.value().header_column);
    }
  }
}

// A filter rewrites c:f to the visible cells and lists filtered series last; the source
// still derives from c15:fullRef in c:order.
TEST(ChartTemplateDerive, FilteredSeriesDeriveFromFullRef) {
  Workbook wb = Workbook::create();
  fill_grid(wb, 0, kStdGrid);
  const auto sp = split(wb, "A1:C5");
  ASSERT_TRUE(static_cast<bool>(sp));
  ChartModel model = model_of(sp.value(), false);
  for (std::size_t i = 0; i < model.plots[0].series.size(); ++i) {
    model.plots[0].series[i].order = static_cast<std::uint32_t>(i);
  }
  for (ChartSeries& s : model.plots[0].series) {
    s.cat.full_formula = s.cat.formula;
    s.cat.formula = "(Sheet1!$A$2:$A$3,Sheet1!$A$5)";
  }
  std::swap(model.plots[0].series[0], model.plots[0].series[1]);
  model.plots[0].series[1].filtered = true;

  const auto derived = derive_chart_source(wb, 0, model);
  ASSERT_TRUE(derived.has_value());
  EXPECT_EQ(derived->range, "Sheet1!$A$1:$C$5");
  EXPECT_EQ(derived->direction, sp.value().direction);
  EXPECT_EQ(derived->header_row, sp.value().header_row);
  EXPECT_EQ(derived->header_column, sp.value().header_column);
}

TEST(ChartTemplateDerive, NonDerivableSeriesReturnEmpty) {
  Workbook wb = Workbook::create();
  fill_grid(wb, 0, kStdGrid);
  const auto sp = split(wb, "A1:C5");
  ASSERT_TRUE(static_cast<bool>(sp));

  EXPECT_FALSE(derive_chart_source(wb, 0, ChartModel{}).has_value());

  ChartModel other_sheet = model_of(sp.value(), false);
  wb.add_sheet("Other");
  other_sheet.plots[0].series[1].val = ref_source("Other!$C$2:$C$5");
  EXPECT_FALSE(derive_chart_source(wb, 0, other_sheet).has_value());

  ChartModel mismatched = model_of(sp.value(), false);
  mismatched.plots[0].series[1].val = ref_source("Sheet1!$C$2:$C$4");
  EXPECT_FALSE(derive_chart_source(wb, 0, mismatched).has_value());

  ChartModel union_ref = model_of(sp.value(), false);
  union_ref.plots[0].series[0].val = ref_source("(Sheet1!$B$2:$B$3,Sheet1!$B$5)");
  EXPECT_FALSE(derive_chart_source(wb, 0, union_ref).has_value());

  ChartModel literal = model_of(sp.value(), false);
  literal.plots[0].series[0].val.kind = ChartDataSource::Kind::kLiteral;
  EXPECT_FALSE(derive_chart_source(wb, 0, literal).has_value());

  ChartModel radar = model_of(sp.value(), false);
  radar.plots[0].type = ChartPlotType::kRadar;
  EXPECT_FALSE(derive_chart_source(wb, 0, radar).has_value());

  // A data range that is a split of no rectangle: the second series skips a column.
  ChartModel gap = model_of(sp.value(), false);
  gap.plots[0].series[1].name = ref_source("Sheet1!$D$1");
  gap.plots[0].series[1].val = ref_source("Sheet1!$D$2:$D$5");
  EXPECT_FALSE(derive_chart_source(wb, 0, gap).has_value());
}

TEST(ChartTemplateDerive, RecoversForcedHeaders) {
  Workbook wb = Workbook::create();
  const std::string rect = fill_grid(wb, 0, kStdGrid);
  const auto sp = split(wb, rect, ChartInsertKind::kColumn, D::kColumns, ChartHeaderMode::kNo, ChartHeaderMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(sp));
  const auto derived = derive_chart_source(wb, 0, model_of(sp.value(), false));
  ASSERT_TRUE(derived.has_value());
  EXPECT_EQ(derived->range, "Sheet1!$A$1:$C$5");
  EXPECT_FALSE(derived->header_row);
  EXPECT_TRUE(derived->header_column);
}

}  // namespace
}  // namespace formulon
