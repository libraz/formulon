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

static_assert(sizeof(fm_read_diagnostics_t) == 24U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_formula_count) == 4U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_defined_name_count) == 8U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_part_count) == 12U);
static_assert(offsetof(fm_read_diagnostics_t, skipped_feature_count) == 16U);
static_assert(offsetof(fm_read_diagnostics_t, unknown_content_type_count) == 20U);

static_assert(sizeof(fm_save_diagnostics_t) == 24U);
static_assert(offsetof(fm_save_diagnostics_t, downgraded_formula_count) == 4U);
static_assert(offsetof(fm_save_diagnostics_t, deferred_feature_count) == 8U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_part_count) == 12U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_relationship_count) == 16U);
static_assert(offsetof(fm_save_diagnostics_t, renumbered_part_count) == 20U);

static_assert(sizeof(fm_validation_outcome_t) == 20U);
static_assert(offsetof(fm_validation_outcome_t, has_rule) == 4U);
static_assert(offsetof(fm_validation_outcome_t, valid) == 8U);
static_assert(offsetof(fm_validation_outcome_t, rule_index) == 12U);
static_assert(offsetof(fm_validation_outcome_t, error_style) == 16U);
static_assert(offsetof(fm_validation_outcome_t, _pad) == 17U);

static_assert(sizeof(fm_parallel_recalc_stats_t) == 56U);
static_assert(offsetof(fm_parallel_recalc_stats_t, cells_evaluated) == 8U);
static_assert(offsetof(fm_parallel_recalc_stats_t, sccs_processed) == 16U);
static_assert(offsetof(fm_parallel_recalc_stats_t, parallel_steps) == 24U);
static_assert(offsetof(fm_parallel_recalc_stats_t, serial_fallback_steps) == 32U);
static_assert(offsetof(fm_parallel_recalc_stats_t, cycle_recoveries) == 40U);
static_assert(offsetof(fm_parallel_recalc_stats_t, worker_threads_started) == 48U);
static_assert(offsetof(fm_parallel_recalc_stats_t, worker_threads_used) == 52U);

static_assert(sizeof(fm_iterative_settings_patch_t) == 24U);
static_assert(offsetof(fm_iterative_settings_patch_t, enabled) == 4U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_iterations) == 8U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_change) == 16U);

static_assert(sizeof(fm_civil_time_t) == 24U);
static_assert(offsetof(fm_civil_time_t, year) == 0U);
static_assert(offsetof(fm_civil_time_t, month) == 4U);
static_assert(offsetof(fm_civil_time_t, day) == 8U);
static_assert(offsetof(fm_civil_time_t, hour) == 12U);
static_assert(offsetof(fm_civil_time_t, minute) == 16U);
static_assert(offsetof(fm_civil_time_t, second) == 20U);

static_assert(sizeof(fm_cf_match_t) == 144U);
static_assert(offsetof(fm_cf_match_t, kind) == 4U);
static_assert(offsetof(fm_cf_match_t, priority) == 8U);
static_assert(offsetof(fm_cf_match_t, dxf_index_engaged) == 12U);
static_assert(offsetof(fm_cf_match_t, dxf_index) == 16U);
static_assert(offsetof(fm_cf_match_t, color) == 24U);
static_assert(offsetof(fm_cf_match_t, bar_length_pct) == 48U);
static_assert(offsetof(fm_cf_match_t, bar_axis_position_pct) == 56U);
static_assert(offsetof(fm_cf_match_t, bar_is_negative) == 64U);
static_assert(offsetof(fm_cf_match_t, bar_fill) == 72U);
static_assert(offsetof(fm_cf_match_t, bar_border_engaged) == 96U);
static_assert(offsetof(fm_cf_match_t, bar_border) == 104U);
static_assert(offsetof(fm_cf_match_t, bar_gradient) == 128U);
static_assert(offsetof(fm_cf_match_t, icon_set_name) == 132U);
static_assert(offsetof(fm_cf_match_t, icon_index) == 136U);
static_assert(offsetof(fm_cf_match_t, bar_direction) == 137U);
static_assert(offsetof(fm_cf_match_t, _pad) == 138U);

static_assert(sizeof(fm_rect_pt_t) == 40U);
static_assert(offsetof(fm_rect_pt_t, x) == 8U);
static_assert(offsetof(fm_rect_pt_t, y) == 16U);
static_assert(offsetof(fm_rect_pt_t, width_pt) == 24U);
static_assert(offsetof(fm_rect_pt_t, height_pt) == 32U);

static_assert(sizeof(fm_paper_info_t) == 32U);
static_assert(offsetof(fm_paper_info_t, width_pt) == 8U);
static_assert(offsetof(fm_paper_info_t, height_pt) == 16U);
static_assert(offsetof(fm_paper_info_t, landscape) == 24U);
static_assert(offsetof(fm_paper_info_t, known) == 28U);

static_assert(sizeof(fm_margins_pt_t) == 56U);
static_assert(offsetof(fm_margins_pt_t, left_in) == 8U);
static_assert(offsetof(fm_margins_pt_t, right_in) == 16U);
static_assert(offsetof(fm_margins_pt_t, top_in) == 24U);
static_assert(offsetof(fm_margins_pt_t, bottom_in) == 32U);
static_assert(offsetof(fm_margins_pt_t, header_in) == 40U);
static_assert(offsetof(fm_margins_pt_t, footer_in) == 48U);

static_assert(sizeof(fm_print_titles_t) == 28U);
static_assert(offsetof(fm_print_titles_t, has_rows) == 4U);
static_assert(offsetof(fm_print_titles_t, first_row) == 8U);
static_assert(offsetof(fm_print_titles_t, last_row) == 12U);
static_assert(offsetof(fm_print_titles_t, has_columns) == 16U);
static_assert(offsetof(fm_print_titles_t, first_col) == 20U);
static_assert(offsetof(fm_print_titles_t, last_col) == 24U);

static_assert(sizeof(fm_page_layout_t) == 56U);
static_assert(offsetof(fm_page_layout_t, area_index) == 4U);
static_assert(offsetof(fm_page_layout_t, rect) == 8U);
static_assert(offsetof(fm_page_layout_t, origin_x_pt) == 24U);
static_assert(offsetof(fm_page_layout_t, origin_y_pt) == 32U);
static_assert(offsetof(fm_page_layout_t, width_pt) == 40U);
static_assert(offsetof(fm_page_layout_t, height_pt) == 48U);

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
static_assert(offsetof(fm_page_setup_t, orientation_engaged) == 4U);
static_assert(offsetof(fm_page_setup_t, orientation) == 8U);
static_assert(offsetof(fm_page_setup_t, paper_size_engaged) == 12U);
static_assert(offsetof(fm_page_setup_t, paper_size) == 16U);
static_assert(offsetof(fm_page_setup_t, scale_pct_engaged) == 20U);
static_assert(offsetof(fm_page_setup_t, scale_pct) == 24U);
static_assert(offsetof(fm_page_setup_t, fit_to_width_engaged) == 28U);
static_assert(offsetof(fm_page_setup_t, fit_to_width) == 32U);
static_assert(offsetof(fm_page_setup_t, fit_to_height_engaged) == 36U);
static_assert(offsetof(fm_page_setup_t, fit_to_height) == 40U);
static_assert(offsetof(fm_page_setup_t, fit_to_page_engaged) == 44U);
static_assert(offsetof(fm_page_setup_t, fit_to_page) == 48U);

static_assert(sizeof(fm_page_margins_t) == 96U);
static_assert(offsetof(fm_page_margins_t, left_in_engaged) == 4U);
static_assert(offsetof(fm_page_margins_t, left_in) == 8U);
static_assert(offsetof(fm_page_margins_t, right_in_engaged) == 16U);
static_assert(offsetof(fm_page_margins_t, right_in) == 24U);
static_assert(offsetof(fm_page_margins_t, top_in_engaged) == 32U);
static_assert(offsetof(fm_page_margins_t, top_in) == 40U);
static_assert(offsetof(fm_page_margins_t, bottom_in_engaged) == 48U);
static_assert(offsetof(fm_page_margins_t, bottom_in) == 56U);
static_assert(offsetof(fm_page_margins_t, header_in_engaged) == 64U);
static_assert(offsetof(fm_page_margins_t, header_in) == 72U);
static_assert(offsetof(fm_page_margins_t, footer_in_engaged) == 80U);
static_assert(offsetof(fm_page_margins_t, footer_in) == 88U);

static_assert(sizeof(fm_print_options_t) == 36U);
static_assert(offsetof(fm_print_options_t, grid_lines_engaged) == 4U);
static_assert(offsetof(fm_print_options_t, grid_lines) == 8U);
static_assert(offsetof(fm_print_options_t, headings_engaged) == 12U);
static_assert(offsetof(fm_print_options_t, headings) == 16U);
static_assert(offsetof(fm_print_options_t, horizontal_centered_engaged) == 20U);
static_assert(offsetof(fm_print_options_t, horizontal_centered) == 24U);
static_assert(offsetof(fm_print_options_t, vertical_centered_engaged) == 28U);
static_assert(offsetof(fm_print_options_t, vertical_centered) == 32U);

static_assert(sizeof(fm_cell_node_t) == 16U);
static_assert(offsetof(fm_cell_node_t, sheet_index) == 4U);
static_assert(offsetof(fm_cell_node_t, row) == 8U);
static_assert(offsetof(fm_cell_node_t, col) == 12U);

static_assert(sizeof(fm_pivot_cache_options_t) == 16U);
static_assert(offsetof(fm_pivot_cache_options_t, save_data) == 4U);
static_assert(offsetof(fm_pivot_cache_options_t, missing_items_limit) == 8U);
static_assert(offsetof(fm_pivot_cache_options_t, refresh_on_load) == 12U);

static_assert(sizeof(fm_spill_info_t) == 28U);
static_assert(offsetof(fm_spill_info_t, rect) == 4U);
static_assert(offsetof(fm_spill_info_t, engaged) == 20U);
static_assert(offsetof(fm_spill_info_t, blocked) == 24U);

static_assert(sizeof(fm_sheet_view_t) == 88U);
static_assert(offsetof(fm_sheet_view_t, zoom_pct) == 4U);
static_assert(offsetof(fm_sheet_view_t, freeze_rows) == 8U);
static_assert(offsetof(fm_sheet_view_t, freeze_cols) == 12U);
static_assert(offsetof(fm_sheet_view_t, visibility) == 16U);
static_assert(offsetof(fm_sheet_view_t, show_grid_lines) == 20U);
static_assert(offsetof(fm_sheet_view_t, show_row_column_headers) == 24U);
static_assert(offsetof(fm_sheet_view_t, show_zeros) == 28U);
static_assert(offsetof(fm_sheet_view_t, right_to_left) == 32U);
static_assert(offsetof(fm_sheet_view_t, tab_selected) == 36U);
static_assert(offsetof(fm_sheet_view_t, view_mode) == 40U);
static_assert(offsetof(fm_sheet_view_t, top_left_row) == 44U);
static_assert(offsetof(fm_sheet_view_t, top_left_col) == 48U);
static_assert(offsetof(fm_sheet_view_t, pane_top_left_row) == 52U);
static_assert(offsetof(fm_sheet_view_t, pane_top_left_col) == 56U);
static_assert(offsetof(fm_sheet_view_t, split_x) == 64U);
static_assert(offsetof(fm_sheet_view_t, split_y) == 72U);
static_assert(offsetof(fm_sheet_view_t, show_formulas) == 80U);
static_assert(offsetof(fm_sheet_view_t, show_outline_symbols) == 84U);

static_assert(sizeof(fm_column_layout_t) == 48U);
static_assert(offsetof(fm_column_layout_t, first) == 4U);
static_assert(offsetof(fm_column_layout_t, last) == 8U);
static_assert(offsetof(fm_column_layout_t, width_chars) == 16U);
static_assert(offsetof(fm_column_layout_t, hidden) == 24U);
static_assert(offsetof(fm_column_layout_t, outline_level) == 28U);
static_assert(offsetof(fm_column_layout_t, _pad) == 29U);
static_assert(offsetof(fm_column_layout_t, collapsed) == 32U);
static_assert(offsetof(fm_column_layout_t, has_width_chars) == 36U);
static_assert(offsetof(fm_column_layout_t, has_xf_index) == 40U);
static_assert(offsetof(fm_column_layout_t, xf_index) == 44U);

static_assert(sizeof(fm_row_layout_t) == 48U);
static_assert(offsetof(fm_row_layout_t, row) == 4U);
static_assert(offsetof(fm_row_layout_t, height_pt) == 8U);
static_assert(offsetof(fm_row_layout_t, hidden) == 16U);
static_assert(offsetof(fm_row_layout_t, outline_level) == 20U);
static_assert(offsetof(fm_row_layout_t, _pad) == 21U);
static_assert(offsetof(fm_row_layout_t, collapsed) == 24U);
static_assert(offsetof(fm_row_layout_t, has_xf_index) == 28U);
static_assert(offsetof(fm_row_layout_t, xf_index) == 32U);
static_assert(offsetof(fm_row_layout_t, has_height_pt) == 36U);
static_assert(offsetof(fm_row_layout_t, custom_height) == 40U);

static_assert(sizeof(fm_date_group_item_t) == 8U);
static_assert(offsetof(fm_date_group_item_t, year) == 0U);
static_assert(offsetof(fm_date_group_item_t, month) == 2U);
static_assert(offsetof(fm_date_group_item_t, day) == 3U);
static_assert(offsetof(fm_date_group_item_t, hour) == 4U);
static_assert(offsetof(fm_date_group_item_t, minute) == 5U);
static_assert(offsetof(fm_date_group_item_t, second) == 6U);
static_assert(offsetof(fm_date_group_item_t, grouping) == 7U);

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

static_assert(sizeof(fm_cell_xf_t) == 128U);
static_assert(offsetof(fm_cell_xf_t, font_index) == 0U);
static_assert(offsetof(fm_cell_xf_t, fill_index) == 4U);
static_assert(offsetof(fm_cell_xf_t, border_index) == 8U);
static_assert(offsetof(fm_cell_xf_t, num_fmt_id) == 12U);
static_assert(offsetof(fm_cell_xf_t, horizontal) == 14U);
static_assert(offsetof(fm_cell_xf_t, vertical) == 15U);
static_assert(offsetof(fm_cell_xf_t, wrap_text) == 16U);
static_assert(offsetof(fm_cell_xf_t, justify_last_line) == 20U);
static_assert(offsetof(fm_cell_xf_t, style_xf_index) == 24U);
static_assert(offsetof(fm_cell_xf_t, has_alignment) == 28U);
static_assert(offsetof(fm_cell_xf_t, has_text_rotation) == 32U);
static_assert(offsetof(fm_cell_xf_t, text_rotation) == 36U);
static_assert(offsetof(fm_cell_xf_t, has_indent) == 40U);
static_assert(offsetof(fm_cell_xf_t, indent) == 44U);
static_assert(offsetof(fm_cell_xf_t, has_relative_indent) == 48U);
static_assert(offsetof(fm_cell_xf_t, relative_indent) == 52U);
static_assert(offsetof(fm_cell_xf_t, has_shrink_to_fit) == 56U);
static_assert(offsetof(fm_cell_xf_t, shrink_to_fit) == 60U);
static_assert(offsetof(fm_cell_xf_t, has_reading_order) == 64U);
static_assert(offsetof(fm_cell_xf_t, reading_order) == 68U);
static_assert(offsetof(fm_cell_xf_t, has_horizontal) == 72U);
static_assert(offsetof(fm_cell_xf_t, has_vertical) == 76U);
static_assert(offsetof(fm_cell_xf_t, has_wrap_text) == 80U);
static_assert(offsetof(fm_cell_xf_t, has_justify_last_line) == 84U);
static_assert(offsetof(fm_cell_xf_t, apply_number_format) == 88U);
static_assert(offsetof(fm_cell_xf_t, apply_font) == 92U);
static_assert(offsetof(fm_cell_xf_t, apply_fill) == 96U);
static_assert(offsetof(fm_cell_xf_t, apply_border) == 100U);
static_assert(offsetof(fm_cell_xf_t, apply_alignment) == 104U);
static_assert(offsetof(fm_cell_xf_t, apply_protection) == 108U);
static_assert(offsetof(fm_cell_xf_t, quote_prefix) == 112U);
static_assert(offsetof(fm_cell_xf_t, has_protection) == 116U);
static_assert(offsetof(fm_cell_xf_t, locked) == 120U);
static_assert(offsetof(fm_cell_xf_t, hidden) == 124U);

static_assert(sizeof(fm_color_spec_t) == 24U);
static_assert(offsetof(fm_color_spec_t, tint) == 0U);
static_assert(offsetof(fm_color_spec_t, argb) == 8U);
static_assert(offsetof(fm_color_spec_t, theme) == 12U);
static_assert(offsetof(fm_color_spec_t, indexed) == 16U);
static_assert(offsetof(fm_color_spec_t, kind) == 20U);

static_assert(sizeof(fm_gradient_stop_t) == 32U);
static_assert(offsetof(fm_gradient_stop_t, position) == 0U);
static_assert(offsetof(fm_gradient_stop_t, color) == 8U);

static_assert(sizeof(fm_border_side_t) == 32U);
static_assert(offsetof(fm_border_side_t, style) == 0U);
static_assert(offsetof(fm_border_side_t, color) == 8U);

static_assert(sizeof(fm_border_record_t) == 168U);
static_assert(offsetof(fm_border_record_t, left) == 0U);
static_assert(offsetof(fm_border_record_t, right) == 32U);
static_assert(offsetof(fm_border_record_t, top) == 64U);
static_assert(offsetof(fm_border_record_t, bottom) == 96U);
static_assert(offsetof(fm_border_record_t, diagonal) == 128U);
static_assert(offsetof(fm_border_record_t, diagonal_up) == 160U);
static_assert(offsetof(fm_border_record_t, diagonal_down) == 164U);

static_assert(sizeof(fm_theme_colors_t) == 52U);
static_assert(offsetof(fm_theme_colors_t, argb) == 4U);

static_assert(sizeof(fm_image_info_t) == 16U);
static_assert(offsetof(fm_image_info_t, format) == 4U);
static_assert(offsetof(fm_image_info_t, width_px) == 8U);
static_assert(offsetof(fm_image_info_t, height_px) == 12U);

static_assert(sizeof(fm_drawing_anchor_t) == 48U);
static_assert(offsetof(fm_drawing_anchor_t, anchor_kind) == 0U);
static_assert(offsetof(fm_drawing_anchor_t, edit_as) == 4U);
static_assert(offsetof(fm_drawing_anchor_t, row) == 8U);
static_assert(offsetof(fm_drawing_anchor_t, col) == 12U);
static_assert(offsetof(fm_drawing_anchor_t, row_offset_emu) == 16U);
static_assert(offsetof(fm_drawing_anchor_t, col_offset_emu) == 24U);
static_assert(offsetof(fm_drawing_anchor_t, width_emu) == 32U);
static_assert(offsetof(fm_drawing_anchor_t, height_emu) == 40U);

static_assert(sizeof(fm_chart_fill_t) == 32U);
static_assert(offsetof(fm_chart_fill_t, color) == 0U);
static_assert(offsetof(fm_chart_fill_t, transparency_pct) == 24U);

static_assert(sizeof(fm_chart_line_t) == 56U);
static_assert(offsetof(fm_chart_line_t, color) == 0U);
static_assert(offsetof(fm_chart_line_t, transparency_pct) == 24U);
static_assert(offsetof(fm_chart_line_t, width_pt) == 32U);
static_assert(offsetof(fm_chart_line_t, has_dash) == 40U);
static_assert(offsetof(fm_chart_line_t, dash) == 44U);
static_assert(offsetof(fm_chart_line_t, none) == 48U);

static_assert(sizeof(fm_chart_marker_t) == 104U);
static_assert(offsetof(fm_chart_marker_t, symbol) == 0U);
static_assert(offsetof(fm_chart_marker_t, has_size_pt) == 4U);
static_assert(offsetof(fm_chart_marker_t, size_pt) == 8U);
static_assert(offsetof(fm_chart_marker_t, fill) == 16U);
static_assert(offsetof(fm_chart_marker_t, line) == 48U);

static_assert(sizeof(fm_chart_point_t) == 208U);
static_assert(offsetof(fm_chart_point_t, idx) == 4U);
static_assert(offsetof(fm_chart_point_t, fill) == 8U);
static_assert(offsetof(fm_chart_point_t, line) == 40U);
static_assert(offsetof(fm_chart_point_t, has_marker) == 96U);
static_assert(offsetof(fm_chart_point_t, marker) == 104U);

static_assert(sizeof(fm_chart_fill_patch_t) == 48U);
static_assert(offsetof(fm_chart_fill_patch_t, color_op) == 0U);
static_assert(offsetof(fm_chart_fill_patch_t, color) == 8U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_pct_op) == 32U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_pct) == 40U);

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

static_assert(sizeof(fm_chart_marker_patch_t) == 152U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol_op) == 0U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol) == 4U);
static_assert(offsetof(fm_chart_marker_patch_t, size_pt_op) == 8U);
static_assert(offsetof(fm_chart_marker_patch_t, size_pt) == 16U);
static_assert(offsetof(fm_chart_marker_patch_t, fill) == 24U);
static_assert(offsetof(fm_chart_marker_patch_t, line) == 72U);

static_assert(sizeof(fm_chart_point_patch_t) == 296U);
static_assert(offsetof(fm_chart_point_patch_t, fill) == 8U);
static_assert(offsetof(fm_chart_point_patch_t, line) == 56U);
static_assert(offsetof(fm_chart_point_patch_t, marker_op) == 136U);
static_assert(offsetof(fm_chart_point_patch_t, marker) == 144U);

static_assert(sizeof(fm_chart_plot_patch_t) == 44U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width_pct_op) == 4U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width_pct) == 8U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap_op) == 12U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap) == 16U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle_deg_op) == 20U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle_deg) == 24U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors_op) == 28U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors) == 32U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size_op) == 36U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size) == 40U);

static_assert(sizeof(fm_chart_display_patch_t) == 28U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as_op) == 4U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as) == 8U);
static_assert(offsetof(fm_chart_display_patch_t, plot_visible_only_op) == 12U);
static_assert(offsetof(fm_chart_display_patch_t, plot_visible_only) == 16U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank_op) == 20U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank) == 24U);

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
