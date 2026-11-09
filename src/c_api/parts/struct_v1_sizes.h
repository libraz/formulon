//
// First-layout ("v1") sizes of the C ABI records, the floor
// `check_struct_size` compares a caller's `struct_size` against.
//
// A record is a struct in `formulon_c.h` whose first member is
// `uint32_t struct_size`; this table lists every one of them and nothing
// else. v1 is the layout each record had in 0.19.0, so every entry is
// `sizeof(T)` today. When a field is appended to a record after 1.0, its
// entry becomes `offsetof(T, <first appended field>)` and never moves again,
// so a caller compiled against v1 is still accepted.
//
// `StructV1<T>` is declared but not defined for anything outside the table:
// checking a value type, or a record added to the header without an entry
// here, is a compile error rather than a silent zero floor.

#ifndef FORMULON_C_API_PARTS_STRUCT_V1_SIZES_H_
#define FORMULON_C_API_PARTS_STRUCT_V1_SIZES_H_

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"

namespace formulon {
namespace c_api {
namespace parts {

template <typename T>
struct StructV1;

#define FORMULON_STRUCT_V1(T, V1_SIZE)                                            \
  static_assert(offsetof(T, struct_size) == 0, #T " must lead with struct_size"); \
  template <>                                                                     \
  struct StructV1<T> {                                                            \
    static constexpr std::size_t kSize = (V1_SIZE);                               \
    static constexpr const char* kName = #T;                                      \
  };

FORMULON_STRUCT_V1(fm_read_diagnostics_info_t, sizeof(fm_read_diagnostics_info_t))
FORMULON_STRUCT_V1(fm_save_diagnostics_info_t, sizeof(fm_save_diagnostics_info_t))
FORMULON_STRUCT_V1(fm_parsed_reference_info_t, sizeof(fm_parsed_reference_info_t))
FORMULON_STRUCT_V1(fm_parse_reference_spec_t, sizeof(fm_parse_reference_spec_t))
FORMULON_STRUCT_V1(fm_defined_name_t, sizeof(fm_defined_name_t))
FORMULON_STRUCT_V1(fm_table_info_t, sizeof(fm_table_info_t))
FORMULON_STRUCT_V1(fm_table_spec_t, sizeof(fm_table_spec_t))
FORMULON_STRUCT_V1(fm_table_patch_t, sizeof(fm_table_patch_t))
FORMULON_STRUCT_V1(fm_table_column_info_t, sizeof(fm_table_column_info_t))
FORMULON_STRUCT_V1(fm_table_style_options_t, sizeof(fm_table_style_options_t))
FORMULON_STRUCT_V1(fm_table_style_options_patch_t, sizeof(fm_table_style_options_patch_t))
FORMULON_STRUCT_V1(fm_hyperlink_t, sizeof(fm_hyperlink_t))
FORMULON_STRUCT_V1(fm_comment_info_t, sizeof(fm_comment_info_t))
FORMULON_STRUCT_V1(fm_workbook_protection_t, sizeof(fm_workbook_protection_t))
FORMULON_STRUCT_V1(fm_protected_range_t, sizeof(fm_protected_range_t))
FORMULON_STRUCT_V1(fm_protected_range_patch_t, sizeof(fm_protected_range_patch_t))
FORMULON_STRUCT_V1(fm_data_validation_t, sizeof(fm_data_validation_t))
FORMULON_STRUCT_V1(fm_validation_outcome_info_t, sizeof(fm_validation_outcome_info_t))
FORMULON_STRUCT_V1(fm_threaded_comment_t, sizeof(fm_threaded_comment_t))
FORMULON_STRUCT_V1(fm_person_t, sizeof(fm_person_t))
FORMULON_STRUCT_V1(fm_threaded_comment_patch_t, sizeof(fm_threaded_comment_patch_t))
FORMULON_STRUCT_V1(fm_parallel_recalc_stats_info_t, sizeof(fm_parallel_recalc_stats_info_t))
FORMULON_STRUCT_V1(fm_iterative_settings_info_t, sizeof(fm_iterative_settings_info_t))
FORMULON_STRUCT_V1(fm_iterative_settings_patch_t, sizeof(fm_iterative_settings_patch_t))
FORMULON_STRUCT_V1(fm_sheet_protection_t, sizeof(fm_sheet_protection_t))
FORMULON_STRUCT_V1(fm_sheet_protection_patch_t, sizeof(fm_sheet_protection_patch_t))
FORMULON_STRUCT_V1(fm_cf_match_info_t, sizeof(fm_cf_match_info_t))
FORMULON_STRUCT_V1(fm_rect_pt_info_t, sizeof(fm_rect_pt_info_t))
FORMULON_STRUCT_V1(fm_paper_info_t, sizeof(fm_paper_info_t))
FORMULON_STRUCT_V1(fm_margins_pt_info_t, sizeof(fm_margins_pt_info_t))
FORMULON_STRUCT_V1(fm_print_titles_t, sizeof(fm_print_titles_t))
FORMULON_STRUCT_V1(fm_page_layout_info_t, sizeof(fm_page_layout_info_t))
FORMULON_STRUCT_V1(fm_row_break_t, sizeof(fm_row_break_t))
FORMULON_STRUCT_V1(fm_column_break_t, sizeof(fm_column_break_t))
FORMULON_STRUCT_V1(fm_page_setup_t, sizeof(fm_page_setup_t))
FORMULON_STRUCT_V1(fm_page_setup_patch_t, sizeof(fm_page_setup_patch_t))
FORMULON_STRUCT_V1(fm_page_margins_t, sizeof(fm_page_margins_t))
FORMULON_STRUCT_V1(fm_page_margins_patch_t, sizeof(fm_page_margins_patch_t))
FORMULON_STRUCT_V1(fm_print_options_t, sizeof(fm_print_options_t))
FORMULON_STRUCT_V1(fm_print_options_patch_t, sizeof(fm_print_options_patch_t))
FORMULON_STRUCT_V1(fm_header_footer_t, sizeof(fm_header_footer_t))
FORMULON_STRUCT_V1(fm_header_footer_patch_t, sizeof(fm_header_footer_patch_t))
FORMULON_STRUCT_V1(fm_cf_rule_t, sizeof(fm_cf_rule_t))
FORMULON_STRUCT_V1(fm_cell_node_info_t, sizeof(fm_cell_node_info_t))
FORMULON_STRUCT_V1(fm_pivot_cell_info_t, sizeof(fm_pivot_cell_info_t))
FORMULON_STRUCT_V1(fm_pivot_field_spec_t, sizeof(fm_pivot_field_spec_t))
FORMULON_STRUCT_V1(fm_pivot_data_field_t, sizeof(fm_pivot_data_field_t))
FORMULON_STRUCT_V1(fm_pivot_field_info_t, sizeof(fm_pivot_field_info_t))
FORMULON_STRUCT_V1(fm_pivot_item_info_t, sizeof(fm_pivot_item_info_t))
FORMULON_STRUCT_V1(fm_pivot_style_options_t, sizeof(fm_pivot_style_options_t))
FORMULON_STRUCT_V1(fm_pivot_style_options_patch_t, sizeof(fm_pivot_style_options_patch_t))
FORMULON_STRUCT_V1(fm_pivot_spec_t, sizeof(fm_pivot_spec_t))
FORMULON_STRUCT_V1(fm_pivot_date_group_spec_t, sizeof(fm_pivot_date_group_spec_t))
FORMULON_STRUCT_V1(fm_pivot_field_patch_t, sizeof(fm_pivot_field_patch_t))
FORMULON_STRUCT_V1(fm_pivot_item_patch_t, sizeof(fm_pivot_item_patch_t))
FORMULON_STRUCT_V1(fm_pivot_worksheet_source_t, sizeof(fm_pivot_worksheet_source_t))
FORMULON_STRUCT_V1(fm_pivot_filter_t, sizeof(fm_pivot_filter_t))
FORMULON_STRUCT_V1(fm_pivot_options_t, sizeof(fm_pivot_options_t))
FORMULON_STRUCT_V1(fm_pivot_options_patch_t, sizeof(fm_pivot_options_patch_t))
FORMULON_STRUCT_V1(fm_pivot_cache_options_t, sizeof(fm_pivot_cache_options_t))
FORMULON_STRUCT_V1(fm_pivot_cache_options_patch_t, sizeof(fm_pivot_cache_options_patch_t))
FORMULON_STRUCT_V1(fm_pivot_grand_totals_patch_t, sizeof(fm_pivot_grand_totals_patch_t))
FORMULON_STRUCT_V1(fm_pivot_calculated_item_info_t, sizeof(fm_pivot_calculated_item_info_t))
FORMULON_STRUCT_V1(fm_spill_info_t, sizeof(fm_spill_info_t))
FORMULON_STRUCT_V1(fm_array_formula_info_t, sizeof(fm_array_formula_info_t))
FORMULON_STRUCT_V1(fm_function_metadata_info_t, sizeof(fm_function_metadata_info_t))
FORMULON_STRUCT_V1(fm_locale_facts_info_t, sizeof(fm_locale_facts_info_t))
FORMULON_STRUCT_V1(fm_sheet_view_info_t, sizeof(fm_sheet_view_info_t))
FORMULON_STRUCT_V1(fm_sheet_view_patch_t, sizeof(fm_sheet_view_patch_t))
FORMULON_STRUCT_V1(fm_column_layout_info_t, sizeof(fm_column_layout_info_t))
FORMULON_STRUCT_V1(fm_row_layout_info_t, sizeof(fm_row_layout_info_t))
FORMULON_STRUCT_V1(fm_auto_filter_t, sizeof(fm_auto_filter_t))
FORMULON_STRUCT_V1(fm_sheet_format_defaults_t, sizeof(fm_sheet_format_defaults_t))
FORMULON_STRUCT_V1(fm_sheet_format_defaults_patch_t, sizeof(fm_sheet_format_defaults_patch_t))
FORMULON_STRUCT_V1(fm_width_model_info_t, sizeof(fm_width_model_info_t))
FORMULON_STRUCT_V1(fm_parse_error_info_t, sizeof(fm_parse_error_info_t))
FORMULON_STRUCT_V1(fm_styles_batch_spec_t, sizeof(fm_styles_batch_spec_t))
FORMULON_STRUCT_V1(fm_table_style_t, sizeof(fm_table_style_t))
FORMULON_STRUCT_V1(fm_table_style_defaults_patch_t, sizeof(fm_table_style_defaults_patch_t))
FORMULON_STRUCT_V1(fm_theme_colors_t, sizeof(fm_theme_colors_t))
FORMULON_STRUCT_V1(fm_theme_fonts_t, sizeof(fm_theme_fonts_t))
FORMULON_STRUCT_V1(fm_effective_style_info_t, sizeof(fm_effective_style_info_t))
FORMULON_STRUCT_V1(fm_external_link_info_t, sizeof(fm_external_link_info_t))
FORMULON_STRUCT_V1(fm_image_info_t, sizeof(fm_image_info_t))
FORMULON_STRUCT_V1(fm_drawing_object_info_t, sizeof(fm_drawing_object_info_t))
FORMULON_STRUCT_V1(fm_picture_insert_spec_t, sizeof(fm_picture_insert_spec_t))
FORMULON_STRUCT_V1(fm_shape_info_t, sizeof(fm_shape_info_t))
FORMULON_STRUCT_V1(fm_shape_patch_t, sizeof(fm_shape_patch_t))
FORMULON_STRUCT_V1(fm_slicer_info_t, sizeof(fm_slicer_info_t))
FORMULON_STRUCT_V1(fm_slicer_patch_t, sizeof(fm_slicer_patch_t))
FORMULON_STRUCT_V1(fm_sparkline_group_t, sizeof(fm_sparkline_group_t))
FORMULON_STRUCT_V1(fm_sparkline_group_patch_t, sizeof(fm_sparkline_group_patch_t))
FORMULON_STRUCT_V1(fm_chart_info_t, sizeof(fm_chart_info_t))
FORMULON_STRUCT_V1(fm_chart_plot_info_t, sizeof(fm_chart_plot_info_t))
FORMULON_STRUCT_V1(fm_chart_series_info_t, sizeof(fm_chart_series_info_t))
FORMULON_STRUCT_V1(fm_chart_literal_info_t, sizeof(fm_chart_literal_info_t))
FORMULON_STRUCT_V1(fm_chart_point_info_t, sizeof(fm_chart_point_info_t))
FORMULON_STRUCT_V1(fm_chart_axis_info_t, sizeof(fm_chart_axis_info_t))
FORMULON_STRUCT_V1(fm_chart_preview_overrides_spec_t, sizeof(fm_chart_preview_overrides_spec_t))
FORMULON_STRUCT_V1(fm_chart_data_info_t, sizeof(fm_chart_data_info_t))
FORMULON_STRUCT_V1(fm_chart_data_series_info_t, sizeof(fm_chart_data_series_info_t))
FORMULON_STRUCT_V1(fm_chart_data_points_info_t, sizeof(fm_chart_data_points_info_t))
FORMULON_STRUCT_V1(fm_chart_axis_title_text_info_t, sizeof(fm_chart_axis_title_text_info_t))
FORMULON_STRUCT_V1(fm_chart_unresolved_info_t, sizeof(fm_chart_unresolved_info_t))
FORMULON_STRUCT_V1(fm_chart_insert_spec_t, sizeof(fm_chart_insert_spec_t))
FORMULON_STRUCT_V1(fm_chart_series_patch_t, sizeof(fm_chart_series_patch_t))
FORMULON_STRUCT_V1(fm_chart_point_patch_t, sizeof(fm_chart_point_patch_t))
FORMULON_STRUCT_V1(fm_chart_plot_patch_t, sizeof(fm_chart_plot_patch_t))
FORMULON_STRUCT_V1(fm_chart_axis_patch_t, sizeof(fm_chart_axis_patch_t))
FORMULON_STRUCT_V1(fm_chart_legend_t, sizeof(fm_chart_legend_t))
FORMULON_STRUCT_V1(fm_chart_legend_patch_t, sizeof(fm_chart_legend_patch_t))
FORMULON_STRUCT_V1(fm_chart_display_patch_t, sizeof(fm_chart_display_patch_t))
FORMULON_STRUCT_V1(fm_data_table_t, sizeof(fm_data_table_t))

#undef FORMULON_STRUCT_V1

// Rejects a record whose `struct_size` is below its v1 size, naming the type
// in the error context.
template <typename T>
inline fm_status_t check_record_size(const T& record) {
  return check_struct_size(record.struct_size, StructV1<T>::kSize, StructV1<T>::kName);
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_STRUCT_V1_SIZES_H_
