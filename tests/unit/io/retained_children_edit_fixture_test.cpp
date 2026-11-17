//
// Row / column edits against Mac Excel 365 (16.113 and 16.115, ja-JP) saves under
// `tests/fixtures/excel/retained_children_edit/`. The `children` book carries
// every retained `<worksheet>` child that names cells -- protected ranges,
// three scenarios, a sheet-level sort state, a consolidation source, a custom
// view, cell watches, ignored errors and a range web-publish item; the
// `autofilter` book an AutoFilter with two criteria and a nested sort state,
// plus row and column breaks, which `.xlsb` retains as records. Excel saved
// each as-is and after each edit, in both formats. Applying the same edit to
// the base file and saving must leave those coordinates where Excel's own
// save does. The ranges inside one sqref are compared as a set: Excel
// reorders them on save, and their order carries nothing. The `controls`,
// `ole_move` and `ole_size` books hold form controls and OLE objects, whose
// cell anchors Excel moves in the worksheet, the drawing part and the VML;
// the `pic` book a picture and a chart. A delete covering every row or column
// of a move-and-size object removes it with the parts only it used.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <functional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/worksheet_child_refs.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/retained_children_edit/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  const ByteSpan span{bytes.data(), bytes.size()};
  if (xlsb) {
    auto r = xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(r));
    return r ? std::move(r.value().workbook) : Workbook::create_empty();
  }
  auto r = read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(r));
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb, bool xlsb) {
  if (xlsb) {
    auto r = xlsb::write_xlsb(wb);
    EXPECT_TRUE(static_cast<bool>(r));
    return r ? std::move(r.value()) : std::vector<std::uint8_t>{};
  }
  auto r = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(r));
  return r ? std::move(r.value()) : std::vector<std::uint8_t>{};
}

std::string ReadEntry(const std::vector<std::uint8_t>& package, const std::string& name) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
}

/// `text`'s space-separated tokens, sorted.
std::string SortedTokens(const std::string& text) {
  std::vector<std::string> tokens;
  std::size_t at = 0;
  while (at < text.size()) {
    const std::size_t end = std::min(text.find(' ', at), text.size());
    if (end > at) {
      tokens.push_back(text.substr(at, end - at));
    }
    at = end + 1U;
  }
  std::sort(tokens.begin(), tokens.end());
  std::string out;
  for (const std::string& t : tokens) {
    out += (out.empty() ? "" : " ") + t;
  }
  return out;
}

/// The coordinate attributes of the fixture's retained children, one string
/// per element in document order: `tag attr=value ...`.
std::vector<std::string> XlsxCoordinates(const std::vector<std::uint8_t>& package) {
  struct Probe {
    const char* tag;
    std::vector<const char*> attrs;
  };
  const std::vector<Probe> probes = {
      {"protectedRange", {"sqref"}},
      {"scenarios", {"current", "show", "sqref"}},
      {"scenario", {"name", "count"}},
      {"inputCells", {"r"}},
      {"sortState", {"ref"}},
      {"sortCondition", {"ref"}},
      {"dataRef", {"ref"}},
      {"customSheetView", {"topLeftCell"}},
      {"cellWatch", {"r"}},
      {"ignoredError", {"sqref"}},
      {"webPublishItems", {"count"}},
      {"webPublishItem", {"sourceRef"}},
      {"autoFilter", {"ref"}},
      {"filterColumn", {"colId"}},
      {"rowBreaks", {"count", "manualBreakCount"}},
      {"colBreaks", {"count", "manualBreakCount"}},
      {"brk", {"id"}},
  };
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
  std::vector<std::pair<std::size_t, std::string>> found;
  for (const Probe& probe : probes) {
    const std::string open = std::string("<") + probe.tag;
    for (std::size_t at = sheet.find(open); at != std::string::npos; at = sheet.find(open, at + 1U)) {
      const char next = sheet[at + open.size()];
      if (next != ' ' && next != '>' && next != '/') {
        continue;
      }
      const std::string tag = sheet.substr(at, sheet.find('>', at) - at);
      std::string line = probe.tag;
      for (const char* attr : probe.attrs) {
        const std::string key = std::string(" ") + attr + "=\"";
        const std::size_t v = tag.find(key);
        if (v != std::string::npos) {
          const std::size_t begin = v + key.size();
          line += std::string(" ") + attr + "=" + SortedTokens(tag.substr(begin, tag.find('"', begin) - begin));
        }
      }
      found.emplace_back(at, line);
    }
  }
  std::sort(found.begin(), found.end());
  std::vector<std::string> out;
  for (auto& [at, line] : found) {
    out.push_back(std::move(line));
  }
  return out;
}

std::uint32_t U32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

std::string Hex(const std::uint8_t* p, std::size_t n) {
  std::string out;
  for (std::size_t i = 0; i < n; ++i) {
    char buf[4];
    std::snprintf(buf, sizeof(buf), " %02x", p[i]);
    out += buf;
  }
  return out;
}

/// The Sqrfx at `offset` as sorted `r0:r1/c0:c1` tokens, then the bytes after it.
std::string SqrfxAt(const xlsb::FramedRecord& rec, std::size_t offset) {
  const std::uint8_t* p = rec.payload.data;
  const std::uint32_t count = U32(p + offset);
  std::string ranges;
  for (std::uint32_t i = 0; i < count; ++i) {
    const std::uint8_t* r = p + offset + 4U + 16U * i;
    ranges += (i == 0 ? "" : " ") + std::to_string(U32(r)) + ":" + std::to_string(U32(r + 4)) + "/" +
              std::to_string(U32(r + 8)) + ":" + std::to_string(U32(r + 12));
  }
  const std::size_t rest = offset + 4U + 16U * count;
  return Hex(p, offset) + " [" + SortedTokens(ranges) + "]" + Hex(p + rest, rec.payload.size - rest);
}

/// The payloads of the records behind the same children, in stream order:
/// the AutoFilter block's range, columns and sort state, range protection,
/// the scenario block, the sheet sort state, the consolidation source, the
/// breaks, cell watches, ignored errors and the web-publish items.
std::vector<std::string> XlsbCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
  const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
  std::vector<std::string> out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
    const std::uint16_t t = rec.type;
    if (t == 536U) {
      out.push_back("536" + SqrfxAt(rec, 2U));
    } else if (t == 649U) {
      out.push_back("649" + SqrfxAt(rec, 4U));
    } else if ((t >= 500U && t <= 504U) || (t >= 530U && t <= 533U) || t == 499U || t == 607U ||
               (t >= 554U && t <= 557U) || t == 161U || t == 162U || t == 163U || (t >= 392U && t <= 396U)) {
      out.push_back(std::to_string(t) + Hex(rec.payload.data, rec.payload.size));
    }
  }
  return out;
}

std::vector<std::string> Coordinates(const std::vector<std::uint8_t>& package, bool xlsb) {
  return xlsb ? XlsbCoordinates(package) : XlsxCoordinates(package);
}

struct EditCase {
  const char* state;
  std::function<Expected<void, Error>(Workbook&)> apply;
};

const std::vector<EditCase> kChildrenEdits = {
    {"1_insert_row4", [](Workbook& wb) { return wb.insert_rows(0, 3, 1); }},
    {"2_delete_row5", [](Workbook& wb) { return wb.delete_rows(0, 4, 1); }},
    {"3_delete_rows3_6", [](Workbook& wb) { return wb.delete_rows(0, 2, 4); }},
    {"4_delete_row9", [](Workbook& wb) { return wb.delete_rows(0, 8, 1); }},
    {"5_delete_rows2_10", [](Workbook& wb) { return wb.delete_rows(0, 1, 9); }},
    {"6_insert_colC", [](Workbook& wb) { return wb.insert_cols(0, 2, 1); }},
    {"7_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
    {"8_delete_colsB_D", [](Workbook& wb) { return wb.delete_cols(0, 1, 3); }},
};

const std::vector<EditCase> kAutoFilterEdits = {
    {"1_insert_row4", [](Workbook& wb) { return wb.insert_rows(0, 3, 1); }},
    {"2_delete_row6", [](Workbook& wb) { return wb.delete_rows(0, 5, 1); }},
    {"3_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
    {"4_insert_colB", [](Workbook& wb) { return wb.insert_cols(0, 1, 1); }},
    {"5_delete_colC", [](Workbook& wb) { return wb.delete_cols(0, 2, 1); }},
};

void CheckBook(const std::string& book, const std::vector<EditCase>& edits, bool xlsb) {
  const char* ext = xlsb ? ".xlsb" : ".xlsx";
  const std::vector<std::uint8_t> base = ReadFixture(book + "_0_base" + ext);
  const std::vector<std::string> base_coords = Coordinates(base, xlsb);
  ASSERT_GE(base_coords.size(), 8U) << book << ext;
  EXPECT_EQ(Coordinates(Save(Load(base, xlsb), xlsb), xlsb), base_coords) << book << ext;
  for (const EditCase& edit : edits) {
    Workbook wb = Load(base, xlsb);
    ASSERT_TRUE(static_cast<bool>(edit.apply(wb))) << edit.state;
    const std::vector<std::string> excel = Coordinates(ReadFixture(book + "_" + edit.state + ext), xlsb);
    EXPECT_NE(excel, base_coords) << book << "_" << edit.state << ext;
    EXPECT_EQ(Coordinates(Save(wb, xlsb), xlsb), excel) << book << "_" << edit.state << ext;
  }
}

TEST(RetainedChildrenEditFixture, ChildrenXlsx) {
  CheckBook("children", kChildrenEdits, false);
}

TEST(RetainedChildrenEditFixture, ChildrenXlsb) {
  CheckBook("children", kChildrenEdits, true);
}

TEST(RetainedChildrenEditFixture, AutoFilterXlsx) {
  CheckBook("autofilter", kAutoFilterEdits, false);
}

TEST(RetainedChildrenEditFixture, AutoFilterXlsb) {
  CheckBook("autofilter", kAutoFilterEdits, true);
}

/// The text content of every element in `block` that is a bare integer, in
/// order: the col / colOff / row / rowOff values of an anchor's markers.
std::string MarkerNumbers(const std::string& block) {
  std::string out;
  for (std::size_t gt = block.find('>'); gt != std::string::npos; gt = block.find('>', gt + 1U)) {
    const std::size_t lt = block.find('<', gt);
    if (lt == std::string::npos || lt == gt + 1U) {
      continue;
    }
    const std::string text = block.substr(gt + 1U, lt - gt - 1U);
    if (text.find_first_not_of("-0123456789") == std::string::npos) {
      out += (out.empty() ? "" : ",") + text;
    }
  }
  return out;
}

/// Every `open`..`close` block of `text` as `label` plus its marker numbers.
void CollectBlocks(const std::string& text, const std::string& open, const std::string& close, const char* label,
                   std::vector<std::string>& out) {
  for (std::size_t at = text.find(open); at != std::string::npos; at = text.find(open, at + 1U)) {
    const std::size_t end = text.find(close, at);
    if (end == std::string::npos) {
      break;
    }
    const std::string block = text.substr(at, end - at);
    out.push_back(std::string(label) + block.substr(open.size(), block.find('>') - open.size()) + " " +
                  MarkerNumbers(block));
  }
}

/// The cell anchors of a sheet's form controls and OLE objects in all three
/// places a package keeps them: the worksheet's `controlPr` / `objectPr`
/// anchor (`.xlsx` only), the drawing part's two-cell anchors, and the VML
/// `x:Anchor` of each legacy shape.
std::vector<std::string> ObjectAnchors(const std::vector<std::uint8_t>& package, bool xlsb) {
  std::vector<std::string> out;
  if (!xlsb) {
    CollectBlocks(ReadEntry(package, "xl/worksheets/sheet1.xml"), "<anchor", "</anchor>", "sheet", out);
  }
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  for (const std::string& name : zip.list_entries()) {
    if (name.rfind("xl/drawings/drawing", 0) == 0 && name.substr(name.size() - 4U) == ".xml") {
      std::vector<std::string> anchors;
      CollectBlocks(ReadEntry(package, name), "<xdr:twoCellAnchor", "</xdr:to>", "drawing", anchors);
      // Excel adds a hidden AutoShape from A1 to the OLE books on every save;
      // it is not one of the objects under test.
      for (std::string& anchor : anchors) {
        if (anchor.find(" 0,0,0,0,") == std::string::npos) {
          out.push_back(std::move(anchor));
        }
      }
    } else if (name.size() > 4U && name.substr(name.size() - 4U) == ".vml") {
      const std::string vml = ReadEntry(package, name);
      for (std::size_t at = vml.find("<x:Anchor>"); at != std::string::npos; at = vml.find("<x:Anchor>", at + 1U)) {
        std::string anchor = vml.substr(at + 10U, vml.find('<', at + 10U) - at - 10U);
        anchor.erase(
            std::remove_if(anchor.begin(), anchor.end(), [](char c) { return c == ' ' || c == '\n' || c == '\r'; }),
            anchor.end());
        out.push_back("vml " + anchor);
      }
    }
  }
  // Part order in the package carries nothing.
  std::sort(out.begin(), out.end());
  return out;
}

/// What a sheet keeps of its drawing objects besides their anchors: the
/// drawing-related parts by kind (digits dropped, as Excel renumbers them),
/// the sheet's relationship types, and the worksheet elements (`.xlsx`) or
/// records (`.xlsb`) that hold or link them. The binary index Excel adds to
/// an `.xlsb` is left out, since a save regenerates the sheets without it.
std::vector<std::string> ObjectParts(const std::vector<std::uint8_t>& package, bool xlsb) {
  std::vector<std::string> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  for (const std::string& name : zip.list_entries()) {
    for (const char* dir : {"xl/drawings/", "xl/embeddings/", "xl/media/", "xl/ctrlProps/", "xl/charts/"}) {
      if (name.rfind(dir, 0) == 0) {
        std::string kind = name;
        kind.erase(std::remove_if(kind.begin(), kind.end(), [](char c) { return c >= '0' && c <= '9'; }), kind.end());
        out.push_back("part " + kind);
      }
    }
  }
  const std::string rels_path = xlsb ? "xl/worksheets/_rels/sheet1.bin.rels" : "xl/worksheets/_rels/sheet1.xml.rels";
  const std::string rels = zip.has_entry(rels_path) ? ReadEntry(package, rels_path) : std::string();
  for (std::size_t at = rels.find(" Type=\""); at != std::string::npos; at = rels.find(" Type=\"", at + 1U)) {
    const std::size_t end = rels.find('"', at + 7U);
    const std::string type = rels.substr(at + 7U, end - at - 7U);
    const std::string kind = type.substr(type.rfind('/') + 1U);
    if (kind != "xlBinaryIndex") {
      out.push_back("rel " + kind);
    }
  }
  if (xlsb) {
    const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
    const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
    for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
      if (rec.type == 550U || rec.type == 551U || (rec.type >= 638U && rec.type <= 640U)) {
        out.push_back("record " + std::to_string(rec.type));
      }
    }
  } else {
    const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
    for (const char* tag : {"<drawing ", "<legacyDrawing ", "<oleObjects>", "<controls>", "<mc:AlternateContent",
                            "<oleObject ", "<control "}) {
      for (std::size_t at = sheet.find(tag); at != std::string::npos; at = sheet.find(tag, at + 1U)) {
        out.push_back(std::string("element ") + tag);
      }
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

const std::vector<EditCase> kControlEdits = {
    {"R_ins_above", [](Workbook& wb) { return wb.insert_rows(0, 1, 2); }},
    {"R_ins_in_cb", [](Workbook& wb) { return wb.insert_rows(0, 5, 1); }},
    {"R_ins_in_bt", [](Workbook& wb) { return wb.insert_rows(0, 10, 1); }},
    {"R_del_top_bt", [](Workbook& wb) { return wb.delete_rows(0, 8, 2); }},
    {"R_del_all_cb", [](Workbook& wb) { return wb.delete_rows(0, 4, 2); }},
    {"R_del_all_bt", [](Workbook& wb) { return wb.delete_rows(0, 9, 3); }},
    {"C_del_in_cb", [](Workbook& wb) { return wb.delete_cols(0, 4, 1); }},
};

const std::vector<EditCase> kOleMoveEdits = {
    {"R_del_cover", [](Workbook& wb) { return wb.delete_rows(0, 3, 10); }},
};

const std::vector<EditCase> kOleSizeEdits = {
    {"C_ins_in_cb", [](Workbook& wb) { return wb.insert_cols(0, 4, 1); }},
    {"R_del_all_cb", [](Workbook& wb) { return wb.delete_rows(0, 4, 2); }},
    {"C_del_cover", [](Workbook& wb) { return wb.delete_cols(0, 2, 7); }},
};

const std::vector<EditCase> kPictureEdits = {
    {"R_del_exact", [](Workbook& wb) { return wb.delete_rows(0, 4, 2); }},
    {"C_del_exact", [](Workbook& wb) { return wb.delete_cols(0, 3, 2); }},
    {"R_del_last_plus", [](Workbook& wb) { return wb.delete_rows(0, 5, 2); }},
};

void CheckObjectAnchors(const std::string& book, const std::vector<EditCase>& edits, bool xlsb) {
  const char* ext = xlsb ? ".xlsb" : ".xlsx";
  const std::vector<std::uint8_t> base = ReadFixture(book + "_0_base" + ext);
  const std::vector<std::string> base_anchors = ObjectAnchors(base, xlsb);
  ASSERT_GE(base_anchors.size(), 2U) << book << ext;
  const std::vector<std::uint8_t> resaved = Save(Load(base, xlsb), xlsb);
  EXPECT_EQ(ObjectAnchors(resaved, xlsb), base_anchors) << book << ext;
  EXPECT_EQ(ObjectParts(resaved, xlsb), ObjectParts(base, xlsb)) << book << ext;
  for (const EditCase& edit : edits) {
    Workbook wb = Load(base, xlsb);
    ASSERT_TRUE(static_cast<bool>(edit.apply(wb))) << edit.state;
    const std::vector<std::uint8_t> excel = ReadFixture(book + "_" + edit.state + ext);
    const std::vector<std::uint8_t> saved = Save(wb, xlsb);
    EXPECT_NE(ObjectAnchors(excel, xlsb), base_anchors) << book << "_" << edit.state << ext;
    EXPECT_EQ(ObjectAnchors(saved, xlsb), ObjectAnchors(excel, xlsb)) << book << "_" << edit.state << ext;
    EXPECT_EQ(ObjectParts(saved, xlsb), ObjectParts(excel, xlsb)) << book << "_" << edit.state << ext;
  }
}

// A check box that moves with cells and a button that moves and sizes with
// them: the check box keeps its line span through an edit inside it, and
// deleting all of the button's rows deletes the button.
TEST(RetainedChildrenEditFixture, FormControlAnchorsXlsx) {
  CheckObjectAnchors("controls", kControlEdits, false);
}

TEST(RetainedChildrenEditFixture, FormControlAnchorsXlsb) {
  CheckObjectAnchors("controls", kControlEdits, true);
}

// A move-only OLE object whose top row is deleted with the rows above it
// lands on the deletion start at its own size; a move-and-size one grows
// with a column inserted inside it, and a delete covering all of its rows or
// columns deletes it with its embedding and preview image.
TEST(RetainedChildrenEditFixture, OleObjectAnchorsXlsx) {
  CheckObjectAnchors("ole_move", kOleMoveEdits, false);
  CheckObjectAnchors("ole_size", kOleSizeEdits, false);
}

TEST(RetainedChildrenEditFixture, OleObjectAnchorsXlsb) {
  CheckObjectAnchors("ole_move", kOleMoveEdits, true);
  CheckObjectAnchors("ole_size", kOleSizeEdits, true);
}

// A move-and-size picture and chart over D5:E6: deleting exactly their rows or
// columns deletes both, and the emptied drawing part goes with them; deleting
// their last row and the one below leaves them one row tall.
TEST(RetainedChildrenEditFixture, PictureAndChartAnchorsXlsx) {
  CheckObjectAnchors("pic", kPictureEdits, false);
}

TEST(RetainedChildrenEditFixture, PictureAndChartAnchorsXlsb) {
  CheckObjectAnchors("pic", kPictureEdits, true);
}

// What the comparisons hold for the two rules no other structure follows,
// spelt out: an insert cuts an ignored-error range instead of stretching it,
// and a watch on a deleted cell keeps its address.
// Removing every control empties `<controls>`, which takes the top-level
// `mc:AlternateContent` around it; the controls' relationship ids come back.
TEST(RetainedChildrenEditFixture, RemovingEveryControlDropsTheWrapper) {
  Workbook wb = Load(ReadFixture("controls_0_base.xlsx"), false);
  WorksheetRawExtensions& children = wb.sheet(0).mutable_raw_extensions();
  const auto holds_controls = [&children] {
    return std::any_of(children.begin(), children.end(),
                       [](const WorksheetRawChild& child) { return child.xml.find("controls>") != std::string::npos; });
  };
  ASSERT_TRUE(holds_controls());
  std::vector<std::string> ids = remove_worksheet_objects(children, {1025U});
  EXPECT_EQ(ids.size(), 1U);
  EXPECT_TRUE(holds_controls());
  std::vector<std::string> rest = remove_worksheet_objects(children, {1026U});
  EXPECT_EQ(rest.size(), 1U);
  EXPECT_NE(ids, rest);
  EXPECT_FALSE(holds_controls());
  EXPECT_TRUE(std::none_of(children.begin(), children.end(), [](const WorksheetRawChild& child) {
    return child.xml.find("AlternateContent") != std::string::npos;
  }));
}

// The deletion rule reads a control's worksheet anchor as well as its drawing
// frame: Button 2 (move-and-size, rows 9..11) goes, Check Box 1 (move-only)
// never does, and a sheet without the drawing still loses the button.
TEST(RetainedChildrenEditFixture, CoveredControlGoesWithoutItsDrawingFrame) {
  Workbook wb = Load(ReadFixture("controls_0_base.xlsx"), false);
  const WorksheetRawExtensions& children = wb.sheet(0).raw_extensions();
  EXPECT_EQ(deleted_cell_objects(children, StructuralEdit{9U, 3U, true, true}), std::vector<std::uint32_t>{1026U});
  EXPECT_TRUE(deleted_cell_objects(children, StructuralEdit{9U, 2U, true, true}).empty());
  EXPECT_TRUE(deleted_cell_objects(children, StructuralEdit{4U, 2U, true, true}).empty());
  EXPECT_TRUE(deleted_cell_objects(children, StructuralEdit{9U, 3U, false, true}).empty());

  wb.sheet(0).set_drawing_rel_target({});
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 9, 3)));
  const auto holds = [&children](std::string_view shape_id) {
    return std::any_of(children.begin(), children.end(), [shape_id](const WorksheetRawChild& child) {
      return child.xml.find(shape_id) != std::string::npos;
    });
  };
  EXPECT_FALSE(holds("shapeId=\"1026\""));
  EXPECT_TRUE(holds("shapeId=\"1025\""));
}

TEST(RetainedChildrenEditFixture, IgnoredErrorsCutAndWatchesStay) {
  Workbook inserted = Load(ReadFixture("children_0_base.xlsx"), false);
  ASSERT_TRUE(static_cast<bool>(inserted.insert_rows(0, 3, 1)));
  const std::vector<std::string> after_insert = XlsxCoordinates(Save(inserted, false));
  EXPECT_NE(std::find(after_insert.begin(), after_insert.end(), "ignoredError sqref=A2:A3 A5:A11 C5"),
            after_insert.end());

  Workbook deleted = Load(ReadFixture("children_0_base.xlsx"), false);
  ASSERT_TRUE(static_cast<bool>(deleted.delete_rows(0, 4, 1)));
  const std::vector<std::string> after_delete = XlsxCoordinates(Save(deleted, false));
  EXPECT_NE(std::find(after_delete.begin(), after_delete.end(), "cellWatch r=C5"), after_delete.end());
  EXPECT_NE(std::find(after_delete.begin(), after_delete.end(), "cellWatch r=E7"), after_delete.end());
}

// A slicer frame stores no size (its xfrm is zero). Under a row delete an absolute frame stays where it is; a
// move-only one keeps its height, its top clamped to the first deleted row when the delete covers it. Mac Excel
// 16.115 saves under `tests/fixtures/excel/slicer_row_delete/`.
TEST(RetainedChildrenEditFixture, SlicerFrameRowDelete) {
  const auto anchors = [](const Workbook& wb) {
    std::vector<std::pair<std::uint32_t, std::uint64_t>> out;
    auto objects = list_drawing_objects(wb, 0);
    EXPECT_TRUE(static_cast<bool>(objects));
    for (const DrawingObject& o : objects ? objects.value() : std::vector<DrawingObject>{}) {
      out.emplace_back(o.from.row, static_cast<std::uint64_t>(o.from.row_off));
      out.emplace_back(o.to.row, static_cast<std::uint64_t>(o.to.row_off));
    }
    return out;
  };
  const std::vector<std::pair<const char*, std::function<Expected<void, Error>(Workbook&)>>> edits = {
      {"del_all", [](Workbook& wb) { return wb.delete_rows(0, 4, 13); }},
      {"del_top", [](Workbook& wb) { return wb.delete_rows(0, 2, 6); }},
      {"del_bot", [](Workbook& wb) { return wb.delete_rows(0, 9, 11); }},
  };
  for (const char* kind : {"absolute", "oneCell"}) {
    const std::vector<std::uint8_t> base = ReadFixture(std::string("../slicer_row_delete/base_") + kind + ".xlsx");
    for (const auto& [name, apply] : edits) {
      SCOPED_TRACE(std::string(name) + "_" + kind);
      Workbook wb = Load(base, false);
      ASSERT_TRUE(static_cast<bool>(apply(wb)));
      const Workbook excel =
          Load(ReadFixture(std::string("../slicer_row_delete/") + name + "_" + kind + ".xlsx"), false);
      EXPECT_FALSE(anchors(excel).empty());
      EXPECT_EQ(anchors(wb), anchors(excel));
    }
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
