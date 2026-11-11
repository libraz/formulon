//
// C ABI - layout pins of the public structs that hold no pointer.
//
// Such a struct carries only fixed-width fields, so its layout is the same on
// every target. This TU compiles into both the native and the wasm32 build,
// and a field that moves, grows or changes type fails both. A struct holding
// a pointer differs between targets and is not pinned here.

#include <cstddef>

#include "c_api/formulon_c.h"

static_assert(sizeof(fm_cell_rect_t) == 16U);
static_assert(offsetof(fm_cell_rect_t, first_row) == 0U);
static_assert(offsetof(fm_cell_rect_t, first_col) == 4U);
static_assert(offsetof(fm_cell_rect_t, last_row) == 8U);
static_assert(offsetof(fm_cell_rect_t, last_col) == 12U);

static_assert(sizeof(fm_sheet_rect_t) == 20U);
static_assert(offsetof(fm_sheet_rect_t, sheet_index) == 0U);
static_assert(offsetof(fm_sheet_rect_t, rect) == 4U);

static_assert(sizeof(fm_cell_address_t) == 8U);
static_assert(offsetof(fm_cell_address_t, row) == 0U);
static_assert(offsetof(fm_cell_address_t, col) == 4U);

static_assert(sizeof(fm_xf_remap_t) == 8U);
static_assert(offsetof(fm_xf_remap_t, from_xf) == 0U);
static_assert(offsetof(fm_xf_remap_t, to_xf) == 4U);

static_assert(sizeof(fm_patch_op_t) == 4U);

static_assert(sizeof(fm_read_diagnostics_info_t) == 24U);
static_assert(offsetof(fm_read_diagnostics_info_t, undecoded_formula_count) == 4U);
static_assert(offsetof(fm_read_diagnostics_info_t, undecoded_defined_name_count) == 8U);
static_assert(offsetof(fm_read_diagnostics_info_t, undecoded_part_count) == 12U);
static_assert(offsetof(fm_read_diagnostics_info_t, skipped_feature_count) == 16U);
static_assert(offsetof(fm_read_diagnostics_info_t, unknown_content_type_count) == 20U);

static_assert(sizeof(fm_save_diagnostics_info_t) == 24U);
static_assert(offsetof(fm_save_diagnostics_info_t, downgraded_formula_count) == 4U);
static_assert(offsetof(fm_save_diagnostics_info_t, deferred_feature_count) == 8U);
static_assert(offsetof(fm_save_diagnostics_info_t, dropped_part_count) == 12U);
static_assert(offsetof(fm_save_diagnostics_info_t, dropped_relationship_count) == 16U);
static_assert(offsetof(fm_save_diagnostics_info_t, renumbered_part_count) == 20U);

static_assert(sizeof(fm_parallel_recalc_stats_info_t) == 56U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, cells_evaluated) == 8U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, sccs_processed) == 16U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, parallel_steps) == 24U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, serial_fallback_steps) == 32U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, cycle_recoveries) == 40U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, worker_threads_started) == 48U);
static_assert(offsetof(fm_parallel_recalc_stats_info_t, worker_threads_used) == 52U);

static_assert(sizeof(fm_iterative_settings_info_t) == 24U);
static_assert(offsetof(fm_iterative_settings_info_t, enabled) == 4U);
static_assert(offsetof(fm_iterative_settings_info_t, max_iterations) == 8U);
static_assert(offsetof(fm_iterative_settings_info_t, max_change) == 16U);

static_assert(sizeof(fm_iterative_settings_patch_t) == 32U);
static_assert(offsetof(fm_iterative_settings_patch_t, enabled_op) == 4U);
static_assert(offsetof(fm_iterative_settings_patch_t, enabled) == 8U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_iterations_op) == 12U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_iterations) == 16U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_change_op) == 20U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_change) == 24U);

static_assert(sizeof(fm_civil_time_t) == 24U);
static_assert(offsetof(fm_civil_time_t, year) == 0U);
static_assert(offsetof(fm_civil_time_t, month) == 4U);
static_assert(offsetof(fm_civil_time_t, day) == 8U);
static_assert(offsetof(fm_civil_time_t, hour) == 12U);
static_assert(offsetof(fm_civil_time_t, minute) == 16U);
static_assert(offsetof(fm_civil_time_t, second) == 20U);

static_assert(sizeof(fm_cf_match_info_t) == 136U);
static_assert(offsetof(fm_cf_match_info_t, kind) == 4U);
static_assert(offsetof(fm_cf_match_info_t, priority) == 8U);
static_assert(offsetof(fm_cf_match_info_t, has_dxf_index) == 12U);
static_assert(offsetof(fm_cf_match_info_t, dxf_index) == 16U);
static_assert(offsetof(fm_cf_match_info_t, color) == 24U);
static_assert(offsetof(fm_cf_match_info_t, bar_length_pct) == 48U);
static_assert(offsetof(fm_cf_match_info_t, bar_axis_position_pct) == 56U);
static_assert(offsetof(fm_cf_match_info_t, bar_is_negative) == 64U);
static_assert(offsetof(fm_cf_match_info_t, bar_fill) == 72U);
static_assert(offsetof(fm_cf_match_info_t, bar_border) == 96U);
static_assert(offsetof(fm_cf_match_info_t, bar_gradient) == 120U);
static_assert(offsetof(fm_cf_match_info_t, icon_set) == 124U);
static_assert(offsetof(fm_cf_match_info_t, icon_index) == 128U);
static_assert(offsetof(fm_cf_match_info_t, bar_direction) == 132U);

static_assert(sizeof(fm_cf_range_evaluation_spec_t) == 24U);
static_assert(offsetof(fm_cf_range_evaluation_spec_t, time_period_rules) == 4U);
static_assert(offsetof(fm_cf_range_evaluation_spec_t, has_today_serial) == 8U);
static_assert(offsetof(fm_cf_range_evaluation_spec_t, today_serial) == 16U);

static_assert(sizeof(fm_rect_pt_info_t) == 40U);
static_assert(offsetof(fm_rect_pt_info_t, x) == 8U);
static_assert(offsetof(fm_rect_pt_info_t, y) == 16U);
static_assert(offsetof(fm_rect_pt_info_t, width_pt) == 24U);
static_assert(offsetof(fm_rect_pt_info_t, height_pt) == 32U);

static_assert(sizeof(fm_paper_info_t) == 32U);
static_assert(offsetof(fm_paper_info_t, width_pt) == 8U);
static_assert(offsetof(fm_paper_info_t, height_pt) == 16U);
static_assert(offsetof(fm_paper_info_t, landscape) == 24U);
static_assert(offsetof(fm_paper_info_t, known) == 28U);

static_assert(sizeof(fm_margins_pt_info_t) == 56U);
static_assert(offsetof(fm_margins_pt_info_t, left_in) == 8U);
static_assert(offsetof(fm_margins_pt_info_t, right_in) == 16U);
static_assert(offsetof(fm_margins_pt_info_t, top_in) == 24U);
static_assert(offsetof(fm_margins_pt_info_t, bottom_in) == 32U);
static_assert(offsetof(fm_margins_pt_info_t, header_in) == 40U);
static_assert(offsetof(fm_margins_pt_info_t, footer_in) == 48U);

static_assert(sizeof(fm_print_titles_t) == 28U);
static_assert(offsetof(fm_print_titles_t, has_rows) == 4U);
static_assert(offsetof(fm_print_titles_t, first_row) == 8U);
static_assert(offsetof(fm_print_titles_t, last_row) == 12U);
static_assert(offsetof(fm_print_titles_t, has_columns) == 16U);
static_assert(offsetof(fm_print_titles_t, first_col) == 20U);
static_assert(offsetof(fm_print_titles_t, last_col) == 24U);

static_assert(sizeof(fm_page_layout_info_t) == 56U);
static_assert(offsetof(fm_page_layout_info_t, area_index) == 4U);
static_assert(offsetof(fm_page_layout_info_t, rect) == 8U);
static_assert(offsetof(fm_page_layout_info_t, origin_x_pt) == 24U);
static_assert(offsetof(fm_page_layout_info_t, origin_y_pt) == 32U);
static_assert(offsetof(fm_page_layout_info_t, width_pt) == 40U);
static_assert(offsetof(fm_page_layout_info_t, height_pt) == 48U);

static_assert(sizeof(fm_parse_reference_spec_t) == 20U);
static_assert(offsetof(fm_parse_reference_spec_t, r1c1) == 4U);
static_assert(offsetof(fm_parse_reference_spec_t, has_base) == 8U);
static_assert(offsetof(fm_parse_reference_spec_t, base_row) == 12U);
static_assert(offsetof(fm_parse_reference_spec_t, base_col) == 16U);

static_assert(sizeof(fm_formula_notation_spec_t) == 20U);
static_assert(offsetof(fm_formula_notation_spec_t, r1c1) == 4U);
static_assert(offsetof(fm_formula_notation_spec_t, has_base) == 8U);
static_assert(offsetof(fm_formula_notation_spec_t, base_row) == 12U);
static_assert(offsetof(fm_formula_notation_spec_t, base_col) == 16U);

static_assert(sizeof(fm_phonetic_properties_t) == 20U);
static_assert(offsetof(fm_phonetic_properties_t, font_index) == 4U);
static_assert(offsetof(fm_phonetic_properties_t, type) == 8U);
static_assert(offsetof(fm_phonetic_properties_t, alignment) == 12U);
static_assert(offsetof(fm_phonetic_properties_t, stated) == 16U);

static_assert(sizeof(fm_row_break_t) == 20U);
static_assert(offsetof(fm_row_break_t, row) == 4U);
static_assert(offsetof(fm_row_break_t, first) == 8U);
static_assert(offsetof(fm_row_break_t, last) == 12U);
static_assert(offsetof(fm_row_break_t, manual) == 16U);

static_assert(sizeof(fm_column_break_t) == 20U);
static_assert(offsetof(fm_column_break_t, col) == 4U);
static_assert(offsetof(fm_column_break_t, first) == 8U);
static_assert(offsetof(fm_column_break_t, last) == 12U);
static_assert(offsetof(fm_column_break_t, manual) == 16U);

static_assert(sizeof(fm_page_setup_t) == 52U);
static_assert(offsetof(fm_page_setup_t, has_orientation) == 4U);
static_assert(offsetof(fm_page_setup_t, orientation) == 8U);
static_assert(offsetof(fm_page_setup_t, has_paper_size) == 12U);
static_assert(offsetof(fm_page_setup_t, paper_size) == 16U);
static_assert(offsetof(fm_page_setup_t, has_scale_pct) == 20U);
static_assert(offsetof(fm_page_setup_t, scale_pct) == 24U);
static_assert(offsetof(fm_page_setup_t, has_fit_to_width) == 28U);
static_assert(offsetof(fm_page_setup_t, fit_to_width) == 32U);
static_assert(offsetof(fm_page_setup_t, has_fit_to_height) == 36U);
static_assert(offsetof(fm_page_setup_t, fit_to_height) == 40U);
static_assert(offsetof(fm_page_setup_t, has_fit_to_page) == 44U);
static_assert(offsetof(fm_page_setup_t, fit_to_page) == 48U);

static_assert(sizeof(fm_page_setup_patch_t) == 52U);
static_assert(offsetof(fm_page_setup_patch_t, orientation_op) == 4U);
static_assert(offsetof(fm_page_setup_patch_t, orientation) == 8U);
static_assert(offsetof(fm_page_setup_patch_t, paper_size_op) == 12U);
static_assert(offsetof(fm_page_setup_patch_t, paper_size) == 16U);
static_assert(offsetof(fm_page_setup_patch_t, scale_pct_op) == 20U);
static_assert(offsetof(fm_page_setup_patch_t, scale_pct) == 24U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_width_op) == 28U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_width) == 32U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_height_op) == 36U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_height) == 40U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_page_op) == 44U);
static_assert(offsetof(fm_page_setup_patch_t, fit_to_page) == 48U);

static_assert(sizeof(fm_page_margins_t) == 96U);
static_assert(offsetof(fm_page_margins_t, has_left_in) == 4U);
static_assert(offsetof(fm_page_margins_t, left_in) == 8U);
static_assert(offsetof(fm_page_margins_t, has_right_in) == 16U);
static_assert(offsetof(fm_page_margins_t, right_in) == 24U);
static_assert(offsetof(fm_page_margins_t, has_top_in) == 32U);
static_assert(offsetof(fm_page_margins_t, top_in) == 40U);
static_assert(offsetof(fm_page_margins_t, has_bottom_in) == 48U);
static_assert(offsetof(fm_page_margins_t, bottom_in) == 56U);
static_assert(offsetof(fm_page_margins_t, has_header_in) == 64U);
static_assert(offsetof(fm_page_margins_t, header_in) == 72U);
static_assert(offsetof(fm_page_margins_t, has_footer_in) == 80U);
static_assert(offsetof(fm_page_margins_t, footer_in) == 88U);

static_assert(sizeof(fm_page_margins_patch_t) == 96U);
static_assert(offsetof(fm_page_margins_patch_t, left_in_op) == 4U);
static_assert(offsetof(fm_page_margins_patch_t, left_in) == 8U);
static_assert(offsetof(fm_page_margins_patch_t, right_in_op) == 16U);
static_assert(offsetof(fm_page_margins_patch_t, right_in) == 24U);
static_assert(offsetof(fm_page_margins_patch_t, top_in_op) == 32U);
static_assert(offsetof(fm_page_margins_patch_t, top_in) == 40U);
static_assert(offsetof(fm_page_margins_patch_t, bottom_in_op) == 48U);
static_assert(offsetof(fm_page_margins_patch_t, bottom_in) == 56U);
static_assert(offsetof(fm_page_margins_patch_t, header_in_op) == 64U);
static_assert(offsetof(fm_page_margins_patch_t, header_in) == 72U);
static_assert(offsetof(fm_page_margins_patch_t, footer_in_op) == 80U);
static_assert(offsetof(fm_page_margins_patch_t, footer_in) == 88U);

static_assert(sizeof(fm_print_options_t) == 36U);
static_assert(offsetof(fm_print_options_t, has_grid_lines) == 4U);
static_assert(offsetof(fm_print_options_t, grid_lines) == 8U);
static_assert(offsetof(fm_print_options_t, has_headings) == 12U);
static_assert(offsetof(fm_print_options_t, headings) == 16U);
static_assert(offsetof(fm_print_options_t, has_horizontal_centered) == 20U);
static_assert(offsetof(fm_print_options_t, horizontal_centered) == 24U);
static_assert(offsetof(fm_print_options_t, has_vertical_centered) == 28U);
static_assert(offsetof(fm_print_options_t, vertical_centered) == 32U);

static_assert(sizeof(fm_print_options_patch_t) == 36U);
static_assert(offsetof(fm_print_options_patch_t, grid_lines_op) == 4U);
static_assert(offsetof(fm_print_options_patch_t, grid_lines) == 8U);
static_assert(offsetof(fm_print_options_patch_t, headings_op) == 12U);
static_assert(offsetof(fm_print_options_patch_t, headings) == 16U);
static_assert(offsetof(fm_print_options_patch_t, horizontal_centered_op) == 20U);
static_assert(offsetof(fm_print_options_patch_t, horizontal_centered) == 24U);
static_assert(offsetof(fm_print_options_patch_t, vertical_centered_op) == 28U);
static_assert(offsetof(fm_print_options_patch_t, vertical_centered) == 32U);

static_assert(sizeof(fm_cell_node_info_t) == 16U);
static_assert(offsetof(fm_cell_node_info_t, sheet_index) == 4U);
static_assert(offsetof(fm_cell_node_info_t, row) == 8U);
static_assert(offsetof(fm_cell_node_info_t, col) == 12U);

static_assert(sizeof(fm_pivot_cache_options_t) == 20U);
static_assert(offsetof(fm_pivot_cache_options_t, save_data) == 4U);
static_assert(offsetof(fm_pivot_cache_options_t, has_missing_items_limit) == 8U);
static_assert(offsetof(fm_pivot_cache_options_t, missing_items_limit) == 12U);
static_assert(offsetof(fm_pivot_cache_options_t, refresh_on_load) == 16U);

static_assert(sizeof(fm_pivot_cache_field_types_info_t) == 36U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_number) == 4U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_integer) == 8U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_string) == 12U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_blank) == 16U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_date) == 20U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_mixed_types) == 24U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_semi_mixed_types) == 28U);
static_assert(offsetof(fm_pivot_cache_field_types_info_t, contains_non_date) == 32U);

static_assert(sizeof(fm_pivot_cache_options_patch_t) == 28U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, save_data_op) == 4U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, save_data) == 8U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, missing_items_limit_op) == 12U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, missing_items_limit) == 16U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, refresh_on_load_op) == 20U);
static_assert(offsetof(fm_pivot_cache_options_patch_t, refresh_on_load) == 24U);

static_assert(sizeof(fm_pivot_grand_totals_patch_t) == 20U);
static_assert(offsetof(fm_pivot_grand_totals_patch_t, rows_op) == 4U);
static_assert(offsetof(fm_pivot_grand_totals_patch_t, rows) == 8U);
static_assert(offsetof(fm_pivot_grand_totals_patch_t, columns_op) == 12U);
static_assert(offsetof(fm_pivot_grand_totals_patch_t, columns) == 16U);

static_assert(sizeof(fm_spill_info_t) == 36U);
static_assert(offsetof(fm_spill_info_t, anchor) == 4U);
static_assert(offsetof(fm_spill_info_t, rect) == 12U);
static_assert(offsetof(fm_spill_info_t, blocked) == 28U);
static_assert(offsetof(fm_spill_info_t, cause) == 32U);

static_assert(sizeof(fm_sheet_view_info_t) == 80U);
static_assert(offsetof(fm_sheet_view_info_t, zoom_pct) == 4U);
static_assert(offsetof(fm_sheet_view_info_t, freeze_rows) == 8U);
static_assert(offsetof(fm_sheet_view_info_t, freeze_cols) == 12U);
static_assert(offsetof(fm_sheet_view_info_t, show_grid_lines) == 16U);
static_assert(offsetof(fm_sheet_view_info_t, show_row_column_headers) == 20U);
static_assert(offsetof(fm_sheet_view_info_t, show_zeros) == 24U);
static_assert(offsetof(fm_sheet_view_info_t, right_to_left) == 28U);
static_assert(offsetof(fm_sheet_view_info_t, tab_selected) == 32U);
static_assert(offsetof(fm_sheet_view_info_t, view_mode) == 36U);
static_assert(offsetof(fm_sheet_view_info_t, top_left_row) == 40U);
static_assert(offsetof(fm_sheet_view_info_t, top_left_col) == 44U);
static_assert(offsetof(fm_sheet_view_info_t, pane_top_left_row) == 48U);
static_assert(offsetof(fm_sheet_view_info_t, pane_top_left_col) == 52U);
static_assert(offsetof(fm_sheet_view_info_t, split_x_twips) == 56U);
static_assert(offsetof(fm_sheet_view_info_t, split_y_twips) == 64U);
static_assert(offsetof(fm_sheet_view_info_t, show_formulas) == 72U);
static_assert(offsetof(fm_sheet_view_info_t, show_outline_symbols) == 76U);
static_assert(sizeof(fm_sheet_view_patch_t) == 152U);
static_assert(offsetof(fm_sheet_view_patch_t, zoom_pct_op) == 4U);
static_assert(offsetof(fm_sheet_view_patch_t, zoom_pct) == 8U);
static_assert(offsetof(fm_sheet_view_patch_t, freeze_rows_op) == 12U);
static_assert(offsetof(fm_sheet_view_patch_t, freeze_rows) == 16U);
static_assert(offsetof(fm_sheet_view_patch_t, freeze_cols_op) == 20U);
static_assert(offsetof(fm_sheet_view_patch_t, freeze_cols) == 24U);
static_assert(offsetof(fm_sheet_view_patch_t, show_grid_lines_op) == 28U);
static_assert(offsetof(fm_sheet_view_patch_t, show_grid_lines) == 32U);
static_assert(offsetof(fm_sheet_view_patch_t, show_row_column_headers_op) == 36U);
static_assert(offsetof(fm_sheet_view_patch_t, show_row_column_headers) == 40U);
static_assert(offsetof(fm_sheet_view_patch_t, show_zeros_op) == 44U);
static_assert(offsetof(fm_sheet_view_patch_t, show_zeros) == 48U);
static_assert(offsetof(fm_sheet_view_patch_t, right_to_left_op) == 52U);
static_assert(offsetof(fm_sheet_view_patch_t, right_to_left) == 56U);
static_assert(offsetof(fm_sheet_view_patch_t, tab_selected_op) == 60U);
static_assert(offsetof(fm_sheet_view_patch_t, tab_selected) == 64U);
static_assert(offsetof(fm_sheet_view_patch_t, view_mode_op) == 68U);
static_assert(offsetof(fm_sheet_view_patch_t, view_mode) == 72U);
static_assert(offsetof(fm_sheet_view_patch_t, top_left_row_op) == 76U);
static_assert(offsetof(fm_sheet_view_patch_t, top_left_row) == 80U);
static_assert(offsetof(fm_sheet_view_patch_t, top_left_col_op) == 84U);
static_assert(offsetof(fm_sheet_view_patch_t, top_left_col) == 88U);
static_assert(offsetof(fm_sheet_view_patch_t, pane_top_left_row_op) == 92U);
static_assert(offsetof(fm_sheet_view_patch_t, pane_top_left_row) == 96U);
static_assert(offsetof(fm_sheet_view_patch_t, pane_top_left_col_op) == 100U);
static_assert(offsetof(fm_sheet_view_patch_t, pane_top_left_col) == 104U);
static_assert(offsetof(fm_sheet_view_patch_t, split_x_twips_op) == 108U);
static_assert(offsetof(fm_sheet_view_patch_t, split_x_twips) == 112U);
static_assert(offsetof(fm_sheet_view_patch_t, split_y_twips_op) == 120U);
static_assert(offsetof(fm_sheet_view_patch_t, split_y_twips) == 128U);
static_assert(offsetof(fm_sheet_view_patch_t, show_formulas_op) == 136U);
static_assert(offsetof(fm_sheet_view_patch_t, show_formulas) == 140U);
static_assert(offsetof(fm_sheet_view_patch_t, show_outline_symbols_op) == 144U);
static_assert(offsetof(fm_sheet_view_patch_t, show_outline_symbols) == 148U);

static_assert(sizeof(fm_column_layout_info_t) == 48U);
static_assert(offsetof(fm_column_layout_info_t, first) == 4U);
static_assert(offsetof(fm_column_layout_info_t, last) == 8U);
static_assert(offsetof(fm_column_layout_info_t, width_chars) == 16U);
static_assert(offsetof(fm_column_layout_info_t, hidden) == 24U);
static_assert(offsetof(fm_column_layout_info_t, outline_level) == 28U);
static_assert(offsetof(fm_column_layout_info_t, _pad) == 29U);
static_assert(offsetof(fm_column_layout_info_t, collapsed) == 32U);
static_assert(offsetof(fm_column_layout_info_t, has_width_chars) == 36U);
static_assert(offsetof(fm_column_layout_info_t, has_xf_index) == 40U);
static_assert(offsetof(fm_column_layout_info_t, xf_index) == 44U);

static_assert(sizeof(fm_row_layout_info_t) == 48U);
static_assert(offsetof(fm_row_layout_info_t, row) == 4U);
static_assert(offsetof(fm_row_layout_info_t, height_pt) == 8U);
static_assert(offsetof(fm_row_layout_info_t, hidden) == 16U);
static_assert(offsetof(fm_row_layout_info_t, outline_level) == 20U);
static_assert(offsetof(fm_row_layout_info_t, _pad) == 21U);
static_assert(offsetof(fm_row_layout_info_t, collapsed) == 24U);
static_assert(offsetof(fm_row_layout_info_t, has_xf_index) == 28U);
static_assert(offsetof(fm_row_layout_info_t, xf_index) == 32U);
static_assert(offsetof(fm_row_layout_info_t, has_height_pt) == 36U);
static_assert(offsetof(fm_row_layout_info_t, custom_height) == 40U);

static_assert(sizeof(fm_cf_icon_t) == 8U);
static_assert(offsetof(fm_cf_icon_t, icon_set) == 0U);
static_assert(offsetof(fm_cf_icon_t, icon_index) == 4U);

static_assert(sizeof(fm_table_style_element_t) == 16U);
static_assert(offsetof(fm_table_style_element_t, type) == 0U);
static_assert(offsetof(fm_table_style_element_t, size) == 4U);
static_assert(offsetof(fm_table_style_element_t, has_dxf_index) == 8U);
static_assert(offsetof(fm_table_style_element_t, dxf_index) == 12U);

static_assert(sizeof(fm_date_group_item_t) == 12U);
static_assert(offsetof(fm_date_group_item_t, year) == 0U);
static_assert(offsetof(fm_date_group_item_t, month) == 2U);
static_assert(offsetof(fm_date_group_item_t, day) == 3U);
static_assert(offsetof(fm_date_group_item_t, hour) == 4U);
static_assert(offsetof(fm_date_group_item_t, minute) == 5U);
static_assert(offsetof(fm_date_group_item_t, second) == 6U);
static_assert(offsetof(fm_date_group_item_t, grouping) == 8U);

static_assert(sizeof(fm_sheet_format_defaults_t) == 64U);
static_assert(offsetof(fm_sheet_format_defaults_t, default_column_width_chars) == 8U);
static_assert(offsetof(fm_sheet_format_defaults_t, default_row_height_pt) == 16U);
static_assert(offsetof(fm_sheet_format_defaults_t, base_column_width_chars) == 24U);
static_assert(offsetof(fm_sheet_format_defaults_t, has_default_column_width_chars) == 32U);
static_assert(offsetof(fm_sheet_format_defaults_t, has_default_row_height_pt) == 36U);
static_assert(offsetof(fm_sheet_format_defaults_t, custom_height) == 40U);
static_assert(offsetof(fm_sheet_format_defaults_t, zero_height) == 44U);
static_assert(offsetof(fm_sheet_format_defaults_t, effective_default_column_width_chars) == 48U);
static_assert(offsetof(fm_sheet_format_defaults_t, effective_default_row_height_pt) == 56U);
static_assert(sizeof(fm_sheet_format_defaults_patch_t) == 64U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, default_column_width_chars_op) == 4U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, default_column_width_chars) == 8U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, default_row_height_pt_op) == 16U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, default_row_height_pt) == 24U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, base_column_width_chars_op) == 32U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, base_column_width_chars) == 40U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, custom_height_op) == 48U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, custom_height) == 52U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, zero_height_op) == 56U);
static_assert(offsetof(fm_sheet_format_defaults_patch_t, zero_height) == 60U);

static_assert(sizeof(fm_cell_xf_t) == 136U);
static_assert(offsetof(fm_cell_xf_t, font_index) == 0U);
static_assert(offsetof(fm_cell_xf_t, fill_index) == 4U);
static_assert(offsetof(fm_cell_xf_t, border_index) == 8U);
static_assert(offsetof(fm_cell_xf_t, num_fmt_id) == 12U);
static_assert(offsetof(fm_cell_xf_t, horizontal) == 16U);
static_assert(offsetof(fm_cell_xf_t, vertical) == 20U);
static_assert(offsetof(fm_cell_xf_t, wrap_text) == 24U);
static_assert(offsetof(fm_cell_xf_t, justify_last_line) == 28U);
static_assert(offsetof(fm_cell_xf_t, style_xf_index) == 32U);
static_assert(offsetof(fm_cell_xf_t, has_alignment) == 36U);
static_assert(offsetof(fm_cell_xf_t, has_text_rotation) == 40U);
static_assert(offsetof(fm_cell_xf_t, text_rotation) == 44U);
static_assert(offsetof(fm_cell_xf_t, has_indent) == 48U);
static_assert(offsetof(fm_cell_xf_t, indent) == 52U);
static_assert(offsetof(fm_cell_xf_t, has_relative_indent) == 56U);
static_assert(offsetof(fm_cell_xf_t, relative_indent) == 60U);
static_assert(offsetof(fm_cell_xf_t, has_shrink_to_fit) == 64U);
static_assert(offsetof(fm_cell_xf_t, shrink_to_fit) == 68U);
static_assert(offsetof(fm_cell_xf_t, has_reading_order) == 72U);
static_assert(offsetof(fm_cell_xf_t, reading_order) == 76U);
static_assert(offsetof(fm_cell_xf_t, has_horizontal) == 80U);
static_assert(offsetof(fm_cell_xf_t, has_vertical) == 84U);
static_assert(offsetof(fm_cell_xf_t, has_wrap_text) == 88U);
static_assert(offsetof(fm_cell_xf_t, has_justify_last_line) == 92U);
static_assert(offsetof(fm_cell_xf_t, apply_number_format) == 96U);
static_assert(offsetof(fm_cell_xf_t, apply_font) == 100U);
static_assert(offsetof(fm_cell_xf_t, apply_fill) == 104U);
static_assert(offsetof(fm_cell_xf_t, apply_border) == 108U);
static_assert(offsetof(fm_cell_xf_t, apply_alignment) == 112U);
static_assert(offsetof(fm_cell_xf_t, apply_protection) == 116U);
static_assert(offsetof(fm_cell_xf_t, quote_prefix) == 120U);
static_assert(offsetof(fm_cell_xf_t, has_protection) == 124U);
static_assert(offsetof(fm_cell_xf_t, locked) == 128U);
static_assert(offsetof(fm_cell_xf_t, hidden) == 132U);

static_assert(sizeof(fm_color_t) == 24U);
static_assert(offsetof(fm_color_t, tint) == 0U);
static_assert(offsetof(fm_color_t, argb) == 8U);
static_assert(offsetof(fm_color_t, theme) == 12U);
static_assert(offsetof(fm_color_t, indexed) == 16U);
static_assert(offsetof(fm_color_t, kind) == 20U);

static_assert(sizeof(fm_gradient_stop_t) == 32U);
static_assert(offsetof(fm_gradient_stop_t, position) == 0U);
static_assert(offsetof(fm_gradient_stop_t, color) == 8U);

static_assert(sizeof(fm_border_side_t) == 32U);
static_assert(offsetof(fm_border_side_t, style) == 0U);
static_assert(offsetof(fm_border_side_t, color) == 8U);

static_assert(sizeof(fm_border_t) == 168U);
static_assert(offsetof(fm_border_t, left) == 0U);
static_assert(offsetof(fm_border_t, right) == 32U);
static_assert(offsetof(fm_border_t, top) == 64U);
static_assert(offsetof(fm_border_t, bottom) == 96U);
static_assert(offsetof(fm_border_t, diagonal) == 128U);
static_assert(offsetof(fm_border_t, diagonal_up) == 160U);
static_assert(offsetof(fm_border_t, diagonal_down) == 164U);

static_assert(sizeof(fm_theme_colors_t) == 52U);
static_assert(offsetof(fm_theme_colors_t, argb) == 4U);

static_assert(sizeof(fm_image_info_t) == 16U);
static_assert(offsetof(fm_image_info_t, format) == 4U);
static_assert(offsetof(fm_image_info_t, width_px) == 8U);
static_assert(offsetof(fm_image_info_t, height_px) == 12U);

static_assert(sizeof(fm_drawing_anchor_t) == 72U);
static_assert(offsetof(fm_drawing_anchor_t, kind) == 0U);
static_assert(offsetof(fm_drawing_anchor_t, edit_as) == 4U);
static_assert(offsetof(fm_drawing_anchor_t, from_row) == 8U);
static_assert(offsetof(fm_drawing_anchor_t, from_col) == 12U);
static_assert(offsetof(fm_drawing_anchor_t, from_row_offset_emu) == 16U);
static_assert(offsetof(fm_drawing_anchor_t, from_col_offset_emu) == 24U);
static_assert(offsetof(fm_drawing_anchor_t, to_row) == 32U);
static_assert(offsetof(fm_drawing_anchor_t, to_col) == 36U);
static_assert(offsetof(fm_drawing_anchor_t, to_row_offset_emu) == 40U);
static_assert(offsetof(fm_drawing_anchor_t, to_col_offset_emu) == 48U);
static_assert(offsetof(fm_drawing_anchor_t, width_emu) == 56U);
static_assert(offsetof(fm_drawing_anchor_t, height_emu) == 64U);

static_assert(sizeof(fm_drawing_anchor_patch_t) == 144U);
static_assert(offsetof(fm_drawing_anchor_patch_t, kind_op) == 0U);
static_assert(offsetof(fm_drawing_anchor_patch_t, kind) == 4U);
static_assert(offsetof(fm_drawing_anchor_patch_t, edit_as_op) == 8U);
static_assert(offsetof(fm_drawing_anchor_patch_t, edit_as) == 12U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_row_op) == 16U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_row) == 20U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_col_op) == 24U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_col) == 28U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_row_offset_emu_op) == 32U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_row_offset_emu) == 40U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_col_offset_emu_op) == 48U);
static_assert(offsetof(fm_drawing_anchor_patch_t, from_col_offset_emu) == 56U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_row_op) == 64U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_row) == 68U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_col_op) == 72U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_col) == 76U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_row_offset_emu_op) == 80U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_row_offset_emu) == 88U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_col_offset_emu_op) == 96U);
static_assert(offsetof(fm_drawing_anchor_patch_t, to_col_offset_emu) == 104U);
static_assert(offsetof(fm_drawing_anchor_patch_t, width_emu_op) == 112U);
static_assert(offsetof(fm_drawing_anchor_patch_t, width_emu) == 120U);
static_assert(offsetof(fm_drawing_anchor_patch_t, height_emu_op) == 128U);
static_assert(offsetof(fm_drawing_anchor_patch_t, height_emu) == 136U);

static_assert(sizeof(fm_chart_fill_t) == 48U);
static_assert(offsetof(fm_chart_fill_t, color) == 0U);
static_assert(offsetof(fm_chart_fill_t, has_transparency_pct) == 24U);
static_assert(offsetof(fm_chart_fill_t, transparency_pct) == 32U);
static_assert(offsetof(fm_chart_fill_t, none) == 40U);
static_assert(offsetof(fm_chart_fill_t, other) == 44U);

static_assert(sizeof(fm_chart_line_t) == 72U);
static_assert(offsetof(fm_chart_line_t, color) == 0U);
static_assert(offsetof(fm_chart_line_t, has_transparency_pct) == 24U);
static_assert(offsetof(fm_chart_line_t, transparency_pct) == 32U);
static_assert(offsetof(fm_chart_line_t, has_width_pt) == 40U);
static_assert(offsetof(fm_chart_line_t, width_pt) == 48U);
static_assert(offsetof(fm_chart_line_t, has_dash) == 56U);
static_assert(offsetof(fm_chart_line_t, dash) == 60U);
static_assert(offsetof(fm_chart_line_t, none) == 64U);
static_assert(offsetof(fm_chart_line_t, other) == 68U);

static_assert(sizeof(fm_chart_marker_t) == 136U);
static_assert(offsetof(fm_chart_marker_t, symbol) == 0U);
static_assert(offsetof(fm_chart_marker_t, has_size_pt) == 4U);
static_assert(offsetof(fm_chart_marker_t, size_pt) == 8U);
static_assert(offsetof(fm_chart_marker_t, fill) == 16U);
static_assert(offsetof(fm_chart_marker_t, line) == 64U);

static_assert(sizeof(fm_chart_point_info_t) == 272U);
static_assert(offsetof(fm_chart_point_info_t, idx) == 4U);
static_assert(offsetof(fm_chart_point_info_t, fill) == 8U);
static_assert(offsetof(fm_chart_point_info_t, line) == 56U);
static_assert(offsetof(fm_chart_point_info_t, has_marker) == 128U);
static_assert(offsetof(fm_chart_point_info_t, marker) == 136U);

static_assert(sizeof(fm_chart_fill_patch_t) == 56U);
static_assert(offsetof(fm_chart_fill_patch_t, color_op) == 0U);
static_assert(offsetof(fm_chart_fill_patch_t, color) == 8U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_pct_op) == 32U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_pct) == 40U);
static_assert(offsetof(fm_chart_fill_patch_t, none_op) == 48U);
static_assert(offsetof(fm_chart_fill_patch_t, none) == 52U);

static_assert(sizeof(fm_chart_line_patch_t) == 80U);
static_assert(offsetof(fm_chart_line_patch_t, color_op) == 0U);
static_assert(offsetof(fm_chart_line_patch_t, color) == 8U);
static_assert(offsetof(fm_chart_line_patch_t, transparency_pct_op) == 32U);
static_assert(offsetof(fm_chart_line_patch_t, transparency_pct) == 40U);
static_assert(offsetof(fm_chart_line_patch_t, width_pt_op) == 48U);
static_assert(offsetof(fm_chart_line_patch_t, width_pt) == 56U);
static_assert(offsetof(fm_chart_line_patch_t, dash_op) == 64U);
static_assert(offsetof(fm_chart_line_patch_t, dash) == 68U);
static_assert(offsetof(fm_chart_line_patch_t, none_op) == 72U);
static_assert(offsetof(fm_chart_line_patch_t, none) == 76U);

static_assert(sizeof(fm_chart_marker_patch_t) == 160U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol_op) == 0U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol) == 4U);
static_assert(offsetof(fm_chart_marker_patch_t, size_pt_op) == 8U);
static_assert(offsetof(fm_chart_marker_patch_t, size_pt) == 16U);
static_assert(offsetof(fm_chart_marker_patch_t, fill) == 24U);
static_assert(offsetof(fm_chart_marker_patch_t, line) == 80U);

static_assert(sizeof(fm_chart_point_patch_t) == 312U);
static_assert(offsetof(fm_chart_point_patch_t, fill) == 8U);
static_assert(offsetof(fm_chart_point_patch_t, line) == 64U);
static_assert(offsetof(fm_chart_point_patch_t, marker_op) == 144U);
static_assert(offsetof(fm_chart_point_patch_t, marker) == 152U);

static_assert(sizeof(fm_chart_plot_patch_t) == 44U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width_pct_op) == 4U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width_pct) == 8U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap_pct_op) == 12U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap_pct) == 16U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle_deg_op) == 20U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle_deg) == 24U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors_op) == 28U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors) == 32U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size_pct_op) == 36U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size_pct) == 40U);

static_assert(sizeof(fm_chart_display_patch_t) == 36U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as_op) == 4U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as) == 8U);
static_assert(offsetof(fm_chart_display_patch_t, plot_visible_only_op) == 12U);
static_assert(offsetof(fm_chart_display_patch_t, plot_visible_only) == 16U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank_op) == 20U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank) == 24U);
static_assert(offsetof(fm_chart_display_patch_t, style_op) == 28U);
static_assert(offsetof(fm_chart_display_patch_t, style) == 32U);

static_assert(sizeof(fm_data_table_t) == 60U);
static_assert(offsetof(fm_data_table_t, rect) == 4U);
static_assert(offsetof(fm_data_table_t, row_input_row) == 20U);
static_assert(offsetof(fm_data_table_t, row_input_col) == 24U);
static_assert(offsetof(fm_data_table_t, column_input_row) == 28U);
static_assert(offsetof(fm_data_table_t, column_input_col) == 32U);
static_assert(offsetof(fm_data_table_t, has_row_input) == 36U);
static_assert(offsetof(fm_data_table_t, has_column_input) == 40U);
static_assert(offsetof(fm_data_table_t, two_dimensional) == 44U);
static_assert(offsetof(fm_data_table_t, row_oriented) == 48U);
static_assert(offsetof(fm_data_table_t, row_input_deleted) == 52U);
static_assert(offsetof(fm_data_table_t, column_input_deleted) == 56U);
