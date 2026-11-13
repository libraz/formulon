// Tests for `Sheet::clone_as`: every member of the copy is rebuilt
// independently of the source, and the identity-bearing fields (tab
// selection, package links, GUIDs) start fresh.

#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "cf/cf_types.h"
#include "data_table.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "phonetic.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "unknown_relationship.h"
#include "utils/derived_guid.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kCopyName = "Src (2)";
constexpr std::string_view kThreadId = "{11111111-2222-4333-8444-555555555555}";
constexpr std::string_view kReplyId = "{66666666-7777-4888-9999-AAAAAAAAAAAA}";
constexpr std::string_view kMentionId = "{BBBBBBBB-CCCC-4DDD-8EEE-FFFFFFFFFFFF}";
constexpr std::string_view kPersonId = "{01234567-89AB-4CDE-8F01-23456789ABCD}";
constexpr std::string_view kNoteUid = "{0F0E0D0C-0B0A-4909-8807-060504030201}";
constexpr std::string_view kValidationUid = "{A1B2C3D4-E5F6-4789-8ABC-DEF012345678}";

/// A workbook whose only sheet touches every member `clone_as` copies. A1's
/// text result is committed by recalc, so its bytes are cell-owned.
Workbook PopulatedWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Src");
  EXPECT_TRUE(wb.set_cell_formula(0, 0, 0, "=\"ab\"&\"cd\""));
  EXPECT_TRUE(wb.set_cell_formula(0, 0, 4, "=SEQUENCE(2)"));
  EXPECT_TRUE(wb.recalc(eval::default_registry()));

  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 1, "literal");
  s.set_cell_value(0, 2, Value::number(7.0));
  s.set_cell_xf_index(0, 2, 3);
  s.set_cell_phonetic(0, 1, "リテラル");
  s.set_cell_phonetic_props(0, 1, PhoneticProperties{2, 1, 2});
  s.set_cell_show_phonetic(0, 1, true);
  s.set_cell_formula(2, 0, "=ROW(A1:A2)");
  s.set_cell_array_block(2, 0, 2, 1);

  s.add_merge(MergeRange{5, 0, 5, 2});

  cf::ConditionalFormat format;
  cf::CFCellRange range;
  range.first = CellAddress{0, 0};
  range.last = CellAddress{9, 0};
  format.sqref.push_back(range);
  cf::CFRule rule;
  rule.formula1 = std::string("A1>0");
  format.rules.push_back(rule);
  s.mutable_conditional_formats().push_back(format);

  DataValidation dv;
  dv.ranges.push_back(MergeRange{0, 3, 3, 3});
  dv.type = 3;
  dv.formula1 = "\"a,b\"";
  dv.extra_attrs = std::string(" imeMode=\"off\" xr:uid=\"") + std::string(kValidationUid) + "\"";
  s.mutable_validations().push_back(dv);
  DataValidation plain;
  plain.ranges.push_back(MergeRange{4, 3, 4, 3});
  plain.extra_attrs = " imeMode=\"on\"";
  s.mutable_validations().push_back(plain);
  s.set_validations_extra_attrs(" disablePrompts=\"1\"");

  Hyperlink link;
  link.row = 6;
  link.col = 0;
  link.last_row = 6;
  link.last_col = 0;
  link.target = "https://example.com/";
  link.location = "Src!A1";
  link.rid = "rId4";
  s.mutable_hyperlinks().push_back(link);

  CellComment note(7, 0, "Author", "note text", std::string(kNoteUid));
  note.vml_shape = "<v:shape/>";
  s.mutable_comments().push_back(note);
  s.set_comment_authors({"Author", "Other"});

  ThreadedComment thread;
  thread.id = std::string(kThreadId);
  thread.row = 8;
  thread.col = 0;
  thread.person_id = std::string(kPersonId);
  thread.created = "2026-11-04T10:00:00.00";
  thread.text = "thread";
  s.mutable_threaded_comments().push_back(thread);
  ThreadedComment reply;
  reply.id = std::string(kReplyId);
  reply.row = 8;
  reply.col = 0;
  reply.person_id = std::string(kPersonId);
  reply.created = "2026-11-04T10:01:00.00";
  reply.text = "@Someone reply";
  reply.parent_id = std::string(kThreadId);
  reply.mentions.push_back(Mention{std::string(kPersonId), std::string(kMentionId), 0, 8});
  s.mutable_threaded_comments().push_back(reply);

  s.set_comment_vml_path("xl/drawings/vmlDrawing1.vml");
  s.set_comment_vml_shapes(s.comment_shapes());
  s.set_comment_vml_extra_xml("<v:shapetype id=\"_x0000_t201\"/>");

  AutoFilter filter;
  filter.range = MergeRange{10, 0, 20, 2};
  s.set_auto_filter(filter);

  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_name("PivotTable1");
  pivot->set_pivot_cache_id(3);
  pivot->set_anchor(30, 0, 4, 2);
  pivot::PivotField field;
  field.source_name = "Region";
  pivot->mutable_fields().push_back(field);
  pivot->mutable_raw_passthrough_xml() = "<extLst/>";
  pivot->set_last_result(pivot::PivotResult{});
  s.add_pivot_table(std::move(pivot));

  DataTable table;
  table.ref = MergeRange{40, 1, 42, 1};
  table.r1.present = true;
  table.r1.row = 39;
  table.r1.col = 0;
  s.mutable_data_tables().push_back(table);

  s.mutable_view().zoom_scale = 150;
  s.mutable_view().tab_selected = true;

  SheetPrintSettings& print = s.mutable_print_settings();
  print.sheet_pr_xml = "<sheetPr filterMode=\"1\" codeName=\"Sheet1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>";
  print.page_setup_xml = "<pageSetup orientation=\"landscape\" r:id=\"rId2\"/>";
  print.printer_settings_rid = "rId2";
  print.printer_settings_path = "xl/printerSettings/printerSettings1.bin";

  s.mutable_protection().enabled = true;
  s.mutable_protection().algorithm_name = "SHA-512";

  ColumnLayout column;
  column.first = 0;
  column.last = 1;
  column.width = 20.0;
  s.mutable_layout().columns.push_back(column);
  RowLayout row;
  row.row = 3;
  row.height = 30.0;
  row.has_height = true;
  s.mutable_layout().row_overrides.push_back(row);
  s.mutable_format_defaults().default_row_height = 18.0;
  s.mutable_format_defaults().has_default_row_height = true;

  s.mutable_raw_extensions().push_back(WorksheetRawChild{2, "<customSheetViews/>"});
  s.set_ext_lst_xml("<extLst><ext uri=\"x14\"/></extLst>");
  s.set_root_extra_ns_attrs(" xmlns:x14ac=\"urn:x14ac\"");
  XlsbSheetTail tail;
  tail.before_merges = {1, 2, 3};
  tail.extern_sheets.push_back(XlsbExternSheetEntry{"Src", "Src", false, 0});
  s.set_xlsb_tail(tail);

  s.set_source_sheet_id(5);
  s.set_drawing_rel_target("xl/drawings/drawing1.xml");
  return wb;
}

std::string Guid(std::string_view source) {
  return derived_guid(source, kCopyName);
}

TEST(SheetClone, CloneIsIndependent) {
  Workbook wb = PopulatedWorkbook();
  const Sheet& src = wb.sheet(0);
  Sheet copy = src.clone_as(std::string(kCopyName));

  // Content equality right after cloning.
  EXPECT_EQ(copy.name(), kCopyName);
  EXPECT_EQ(copy.cell_count(), src.cell_count() - 1U);  // the SEQUENCE phantom is not cloned
  const Cell* a1 = copy.cell_at(0, 0);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(src.cell_at(0, 0)->cached_text_owned, nullptr);
  EXPECT_EQ(a1->formula_text, src.cell_at(0, 0)->formula_text);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "abcd");
  ASSERT_NE(a1->cached_text_owned, nullptr);
  EXPECT_EQ(a1->cached_value.as_text().data(), a1->cached_text_owned->data());
  EXPECT_NE(a1->cached_value.as_text().data(), src.cell_at(0, 0)->cached_value.as_text().data());
  const Cell* b1 = copy.cell_at(0, 1);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->cached_value.as_text(), "literal");
  EXPECT_NE(b1->cached_value.as_text().data(), src.cell_at(0, 1)->cached_value.as_text().data());
  ASSERT_EQ(b1->phonetic_runs.size(), 1U);
  EXPECT_EQ(b1->phonetic_runs[0].text, "リテラル");
  EXPECT_EQ(b1->phonetic_props.type, 1U);
  EXPECT_TRUE(b1->phonetic_present);
  EXPECT_TRUE(b1->show_phonetic);
  const Cell* c1 = copy.cell_at(0, 2);
  ASSERT_NE(c1, nullptr);
  EXPECT_EQ(c1->xf_index, 3U);
  EXPECT_TRUE(c1->has_explicit_xf);
  const Cell* e1 = copy.cell_at(0, 4);
  ASSERT_NE(e1, nullptr);
  EXPECT_EQ(e1->dynamic_array, src.cell_at(0, 4)->dynamic_array);
  EXPECT_EQ(copy.cell_at(2, 0)->array_block_rows, 2U);
  EXPECT_EQ(copy.cell_at(2, 0)->array_block_cols, 1U);
  const auto src_formulas = src.formula_cells_in(0, 0, 100, 10);
  const auto copy_formulas = copy.formula_cells_in(0, 0, 100, 10);
  EXPECT_EQ(copy_formulas, src_formulas);
  EXPECT_EQ(copy.merges(), src.merges());
  ASSERT_EQ(copy.conditional_formats().size(), 1U);
  EXPECT_EQ(copy.conditional_formats()[0].rules[0].formula1, std::optional<std::string>("A1>0"));
  ASSERT_EQ(copy.validations().size(), 2U);
  EXPECT_EQ(copy.validations()[0].formula1, "\"a,b\"");
  EXPECT_EQ(copy.validations()[1].extra_attrs, " imeMode=\"on\"");
  EXPECT_EQ(copy.validations_extra_attrs(), " disablePrompts=\"1\"");
  ASSERT_EQ(copy.hyperlinks().size(), 1U);
  EXPECT_EQ(copy.hyperlinks()[0].target, "https://example.com/");
  EXPECT_EQ(copy.hyperlinks()[0].location, "Src!A1");
  ASSERT_EQ(copy.comments().size(), 1U);
  EXPECT_EQ(copy.comments()[0].text, "note text");
  EXPECT_EQ(copy.comments()[0].vml_shape, "<v:shape/>");
  EXPECT_EQ(copy.comment_authors(), src.comment_authors());
  EXPECT_EQ(copy.comment_vml_shapes(), src.comment_vml_shapes());
  EXPECT_EQ(copy.comment_vml_extra_xml(), src.comment_vml_extra_xml());
  ASSERT_EQ(copy.threaded_comments().size(), 2U);
  EXPECT_EQ(copy.threaded_comments()[1].text, "@Someone reply");
  ASSERT_NE(copy.auto_filter(), nullptr);
  EXPECT_EQ(copy.auto_filter()->range, (MergeRange{10, 0, 20, 2}));
  ASSERT_EQ(copy.pivot_tables().size(), 1U);
  EXPECT_NE(copy.pivot_tables()[0].get(), src.pivot_tables()[0].get());
  EXPECT_EQ(copy.pivot_tables()[0]->name(), "PivotTable1");
  EXPECT_EQ(copy.pivot_tables()[0]->pivot_cache_id(), 3U);
  EXPECT_EQ(copy.pivot_tables()[0]->anchor_row(), 30U);
  ASSERT_EQ(copy.pivot_tables()[0]->fields().size(), 1U);
  EXPECT_EQ(copy.pivot_tables()[0]->fields()[0].source_name, "Region");
  EXPECT_EQ(copy.pivot_tables()[0]->raw_passthrough_xml(), "<extLst/>");
  ASSERT_EQ(copy.data_tables().size(), 1U);
  EXPECT_EQ(copy.data_tables()[0].ref, (MergeRange{40, 1, 42, 1}));
  EXPECT_EQ(copy.data_tables()[0].r1, src.data_tables()[0].r1);
  EXPECT_EQ(copy.view().zoom_scale, 150U);
  EXPECT_EQ(copy.print_settings().page_setup_xml, src.print_settings().page_setup_xml);
  EXPECT_TRUE(copy.protection().enabled);
  EXPECT_EQ(copy.protection().algorithm_name, "SHA-512");
  ASSERT_EQ(copy.layout().columns.size(), 1U);
  EXPECT_EQ(copy.layout().columns[0].width, 20.0);
  ASSERT_EQ(copy.layout().row_overrides.size(), 1U);
  EXPECT_EQ(copy.layout().row_overrides[0].height, 30.0);
  EXPECT_EQ(copy.format_defaults().default_row_height, 18.0);
  ASSERT_EQ(copy.raw_extensions().size(), 1U);
  EXPECT_EQ(copy.raw_extensions()[0].xml, "<customSheetViews/>");
  EXPECT_EQ(copy.ext_lst_xml(), src.ext_lst_xml());
  EXPECT_EQ(copy.root_extra_ns_attrs(), src.root_extra_ns_attrs());
  EXPECT_EQ(copy.xlsb_tail().before_merges, src.xlsb_tail().before_merges);
  ASSERT_EQ(copy.xlsb_tail().extern_sheets.size(), 1U);

  // Mutating each member of the copy leaves the source untouched.
  copy.set_cell_cached_value(0, 0, Value::text("changed"));
  copy.set_cell_value(0, 1, Value::number(1.0));
  copy.set_cell_phonetic(0, 1, "");
  copy.set_cell_xf_index(0, 2, 9);
  copy.set_cell_value(2, 0, Value::blank());
  copy.add_merge(MergeRange{50, 0, 50, 1});
  copy.mutable_conditional_formats()[0].rules[0].formula1 = std::string("A1<0");
  copy.mutable_validations()[0].formula1 = "\"c\"";
  copy.set_validations_extra_attrs("");
  copy.mutable_hyperlinks()[0].target = "https://changed.example/";
  copy.mutable_comments()[0].text = "changed";
  copy.set_comment_authors({});
  copy.mutable_threaded_comments()[0].text = "changed";
  copy.set_comment_vml_shapes({});
  copy.set_comment_vml_extra_xml("");
  copy.clear_auto_filter_model();
  copy.mutable_pivot_tables()[0]->set_name("Renamed");
  copy.mutable_pivot_tables()[0]->mutable_fields()[0].source_name = "Changed";
  copy.mutable_data_tables()[0].ref = MergeRange{60, 1, 61, 1};
  copy.mutable_view().zoom_scale = 50;
  copy.mutable_print_settings().page_setup_xml.clear();
  copy.mutable_protection().enabled = false;
  copy.mutable_layout().columns[0].width = 1.0;
  copy.mutable_layout().row_overrides.clear();
  copy.mutable_format_defaults().default_row_height = 1.0;
  copy.mutable_raw_extensions().clear();
  copy.set_ext_lst_xml("");
  copy.set_root_extra_ns_attrs("");
  copy.set_xlsb_tail(XlsbSheetTail{});

  EXPECT_EQ(src.cell_at(0, 0)->cached_value.as_text(), "abcd");
  EXPECT_EQ(src.cell_at(0, 1)->cached_value.as_text(), "literal");
  ASSERT_EQ(src.cell_at(0, 1)->phonetic_runs.size(), 1U);
  EXPECT_EQ(src.cell_at(0, 2)->xf_index, 3U);
  EXPECT_EQ(src.formula_cells_in(0, 0, 100, 10), src_formulas);
  EXPECT_NE(copy.formula_cells_in(0, 0, 100, 10), src_formulas);
  EXPECT_EQ(src.cell_at(2, 0)->array_block_rows, 2U);
  EXPECT_EQ(src.merges().size(), 1U);
  EXPECT_EQ(src.conditional_formats()[0].rules[0].formula1, std::optional<std::string>("A1>0"));
  EXPECT_EQ(src.validations()[0].formula1, "\"a,b\"");
  EXPECT_EQ(src.validations_extra_attrs(), " disablePrompts=\"1\"");
  EXPECT_EQ(src.hyperlinks()[0].target, "https://example.com/");
  EXPECT_EQ(src.comments()[0].text, "note text");
  EXPECT_EQ(src.comment_authors().size(), 2U);
  EXPECT_EQ(src.threaded_comments()[0].text, "thread");
  EXPECT_EQ(src.comment_vml_shapes().size(), 2U);
  EXPECT_FALSE(src.comment_vml_extra_xml().empty());
  ASSERT_NE(src.auto_filter(), nullptr);
  EXPECT_EQ(src.pivot_tables()[0]->name(), "PivotTable1");
  EXPECT_EQ(src.pivot_tables()[0]->fields()[0].source_name, "Region");
  EXPECT_EQ(src.data_tables()[0].ref, (MergeRange{40, 1, 42, 1}));
  EXPECT_EQ(src.view().zoom_scale, 150U);
  EXPECT_FALSE(src.print_settings().page_setup_xml.empty());
  EXPECT_TRUE(src.protection().enabled);
  EXPECT_EQ(src.layout().columns[0].width, 20.0);
  EXPECT_EQ(src.layout().row_overrides.size(), 1U);
  EXPECT_EQ(src.format_defaults().default_row_height, 18.0);
  EXPECT_EQ(src.raw_extensions().size(), 1U);
  EXPECT_FALSE(src.ext_lst_xml().empty());
  EXPECT_FALSE(src.root_extra_ns_attrs().empty());
  EXPECT_FALSE(src.xlsb_tail().empty());
}

TEST(SheetClone, CopyOutlivesSource) {
  std::optional<Sheet> copy;
  {
    Workbook wb = PopulatedWorkbook();
    copy.emplace(wb.sheet(0).clone_as(std::string(kCopyName)));
  }
  ASSERT_NE(copy->cell_at(0, 0), nullptr);
  EXPECT_EQ(copy->cell_at(0, 0)->cached_value.as_text(), "abcd");
  EXPECT_EQ(copy->cell_at(0, 1)->cached_value.as_text(), "literal");
}

TEST(SheetClone, FreshIdentity) {
  Workbook wb = PopulatedWorkbook();
  const Sheet& src = wb.sheet(0);
  const Sheet copy = src.clone_as(std::string(kCopyName));

  EXPECT_FALSE(copy.view().tab_selected);
  EXPECT_TRUE(src.view().tab_selected);
  EXPECT_EQ(copy.source_sheet_id(), 0U);
  EXPECT_TRUE(copy.drawing_rel_target().empty());
  EXPECT_TRUE(copy.comment_vml_path().empty());
  EXPECT_FALSE(src.comment_vml_path().empty());
  EXPECT_TRUE(copy.hyperlinks()[0].rid.empty());
  EXPECT_EQ(src.hyperlinks()[0].rid, "rId4");
  EXPECT_TRUE(copy.print_settings().printer_settings_rid.empty());
  EXPECT_EQ(copy.print_settings().printer_settings_path, "xl/printerSettings/printerSettings1.bin");
  EXPECT_EQ(copy.print_settings().sheet_pr_xml, "<sheetPr filterMode=\"1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>");
  EXPECT_EQ(copy.cell_enumeration_revision(), 0U);

  // Spill regions are rebuilt by the next recalc, not copied.
  EXPECT_NE(src.spill_region_at_anchor(0, 4), nullptr);
  EXPECT_EQ(copy.spill_region_at_anchor(0, 4), nullptr);
  EXPECT_EQ(copy.pivot_tables()[0]->last_result(), nullptr);
  EXPECT_NE(src.pivot_tables()[0]->last_result(), nullptr);

  // Derived GUIDs: deterministic, distinct from the source.
  EXPECT_EQ(copy.comments()[0].uid, Guid(kNoteUid));
  EXPECT_NE(copy.comments()[0].uid, kNoteUid);
  EXPECT_EQ(copy.validations()[0].extra_attrs, " imeMode=\"off\" xr:uid=\"" + Guid(kValidationUid) + "\"");
  const ThreadedComment& thread = copy.threaded_comments()[0];
  const ThreadedComment& reply = copy.threaded_comments()[1];
  EXPECT_EQ(thread.id, Guid(kThreadId));
  EXPECT_NE(thread.id, kThreadId);
  EXPECT_TRUE(is_threaded_comment_guid(thread.id));
  EXPECT_EQ(reply.id, Guid(kReplyId));
  EXPECT_EQ(reply.parent_id, thread.id);
  EXPECT_TRUE(thread.parent_id.empty());
  ASSERT_EQ(reply.mentions.size(), 1U);
  EXPECT_EQ(reply.mentions[0].mention_id, Guid(kMentionId));
  EXPECT_EQ(reply.mentions[0].person_id, kPersonId);
  EXPECT_EQ(thread.person_id, kPersonId);

  const Sheet again = src.clone_as(std::string(kCopyName));
  EXPECT_EQ(again.threaded_comments()[0].id, thread.id);
  EXPECT_EQ(again.comments()[0].uid, copy.comments()[0].uid);
  const Sheet other = src.clone_as("Src (3)");
  EXPECT_NE(other.threaded_comments()[0].id, thread.id);
}

// The worksheet root's revision uid is derived like the other uids, so the copy does not share its source's.
TEST(SheetClone, RootUidIsDerived) {
  constexpr std::string_view kRootUid = "{6B2E0F4A-1C3D-4E5F-8A9B-0C1D2E3F4A5B}";
  Workbook wb = PopulatedWorkbook();
  wb.sheet(0).set_root_extra_ns_attrs(" xmlns:x14ac=\"urn:x14ac\" xr:uid=\"" + std::string(kRootUid) + "\"");
  const Sheet copy = wb.sheet(0).clone_as(std::string(kCopyName));
  EXPECT_EQ(copy.root_extra_ns_attrs(), " xmlns:x14ac=\"urn:x14ac\" xr:uid=\"" + Guid(kRootUid) + "\"");
}

}  // namespace
}  // namespace formulon
