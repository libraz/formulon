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

static_assert(sizeof(fm_read_diagnostics_t) == 20U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_formula_count) == 0U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_defined_name_count) == 4U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_part_count) == 8U);
static_assert(offsetof(fm_read_diagnostics_t, skipped_feature_count) == 12U);
static_assert(offsetof(fm_read_diagnostics_t, unknown_content_type_count) == 16U);

static_assert(sizeof(fm_save_diagnostics_t) == 20U);
static_assert(offsetof(fm_save_diagnostics_t, downgraded_formula_count) == 0U);
static_assert(offsetof(fm_save_diagnostics_t, deferred_feature_count) == 4U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_part_count) == 8U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_relationship_count) == 12U);
static_assert(offsetof(fm_save_diagnostics_t, renumbered_part_count) == 16U);

static_assert(sizeof(fm_validation_outcome_t) == 16U);
static_assert(offsetof(fm_validation_outcome_t, has_rule) == 0U);
static_assert(offsetof(fm_validation_outcome_t, valid) == 4U);
static_assert(offsetof(fm_validation_outcome_t, rule_index) == 8U);
static_assert(offsetof(fm_validation_outcome_t, error_style) == 12U);
static_assert(offsetof(fm_validation_outcome_t, _pad) == 13U);

static_assert(sizeof(fm_parallel_recalc_stats_t) == 48U);
static_assert(offsetof(fm_parallel_recalc_stats_t, cells_evaluated) == 0U);
static_assert(offsetof(fm_parallel_recalc_stats_t, sccs_processed) == 8U);
static_assert(offsetof(fm_parallel_recalc_stats_t, parallel_steps) == 16U);
static_assert(offsetof(fm_parallel_recalc_stats_t, serial_fallback_steps) == 24U);
static_assert(offsetof(fm_parallel_recalc_stats_t, cycle_recoveries) == 32U);
static_assert(offsetof(fm_parallel_recalc_stats_t, worker_threads_started) == 40U);
static_assert(offsetof(fm_parallel_recalc_stats_t, worker_threads_used) == 44U);

static_assert(sizeof(fm_iterative_settings_patch_t) == 16U);
static_assert(offsetof(fm_iterative_settings_patch_t, enabled) == 0U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_iterations) == 4U);
static_assert(offsetof(fm_iterative_settings_patch_t, max_change) == 8U);

static_assert(sizeof(fm_civil_time_t) == 24U);
static_assert(offsetof(fm_civil_time_t, year) == 0U);
static_assert(offsetof(fm_civil_time_t, month) == 4U);
static_assert(offsetof(fm_civil_time_t, day) == 8U);
static_assert(offsetof(fm_civil_time_t, hour) == 12U);
static_assert(offsetof(fm_civil_time_t, minute) == 16U);
static_assert(offsetof(fm_civil_time_t, second) == 20U);

static_assert(sizeof(fm_cf_color_t) == 4U);
static_assert(offsetof(fm_cf_color_t, r) == 0U);
static_assert(offsetof(fm_cf_color_t, g) == 1U);
static_assert(offsetof(fm_cf_color_t, b) == 2U);
static_assert(offsetof(fm_cf_color_t, a) == 3U);

static_assert(sizeof(fm_cf_match_t) == 72U);
static_assert(offsetof(fm_cf_match_t, kind) == 0U);
static_assert(offsetof(fm_cf_match_t, priority) == 4U);
static_assert(offsetof(fm_cf_match_t, dxf_id_engaged) == 8U);
static_assert(offsetof(fm_cf_match_t, dxf_id) == 12U);
static_assert(offsetof(fm_cf_match_t, color) == 16U);
static_assert(offsetof(fm_cf_match_t, bar_length_pct) == 24U);
static_assert(offsetof(fm_cf_match_t, bar_axis_position_pct) == 32U);
static_assert(offsetof(fm_cf_match_t, bar_is_negative) == 40U);
static_assert(offsetof(fm_cf_match_t, bar_fill) == 44U);
static_assert(offsetof(fm_cf_match_t, bar_border_engaged) == 48U);
static_assert(offsetof(fm_cf_match_t, bar_border) == 52U);
static_assert(offsetof(fm_cf_match_t, bar_gradient) == 56U);
static_assert(offsetof(fm_cf_match_t, icon_set_name) == 60U);
static_assert(offsetof(fm_cf_match_t, icon_index) == 64U);
static_assert(offsetof(fm_cf_match_t, bar_direction) == 65U);
static_assert(offsetof(fm_cf_match_t, _pad) == 66U);

static_assert(sizeof(fm_rect_pt_t) == 32U);
static_assert(offsetof(fm_rect_pt_t, x) == 0U);
static_assert(offsetof(fm_rect_pt_t, y) == 8U);
static_assert(offsetof(fm_rect_pt_t, width) == 16U);
static_assert(offsetof(fm_rect_pt_t, height) == 24U);

static_assert(sizeof(fm_paper_info_t) == 24U);
static_assert(offsetof(fm_paper_info_t, width_pt) == 0U);
static_assert(offsetof(fm_paper_info_t, height_pt) == 8U);
static_assert(offsetof(fm_paper_info_t, landscape) == 16U);
static_assert(offsetof(fm_paper_info_t, known) == 20U);

static_assert(sizeof(fm_margins_pt_t) == 48U);
static_assert(offsetof(fm_margins_pt_t, left) == 0U);
static_assert(offsetof(fm_margins_pt_t, right) == 8U);
static_assert(offsetof(fm_margins_pt_t, top) == 16U);
static_assert(offsetof(fm_margins_pt_t, bottom) == 24U);
static_assert(offsetof(fm_margins_pt_t, header) == 32U);
static_assert(offsetof(fm_margins_pt_t, footer) == 40U);

static_assert(sizeof(fm_print_titles_t) == 24U);
static_assert(offsetof(fm_print_titles_t, has_rows) == 0U);
static_assert(offsetof(fm_print_titles_t, first_row) == 4U);
static_assert(offsetof(fm_print_titles_t, last_row) == 8U);
static_assert(offsetof(fm_print_titles_t, has_cols) == 12U);
static_assert(offsetof(fm_print_titles_t, first_col) == 16U);
static_assert(offsetof(fm_print_titles_t, last_col) == 20U);

static_assert(sizeof(fm_page_layout_t) == 56U);
static_assert(offsetof(fm_page_layout_t, area_index) == 0U);
static_assert(offsetof(fm_page_layout_t, rect) == 4U);
static_assert(offsetof(fm_page_layout_t, origin_x_pt) == 24U);
static_assert(offsetof(fm_page_layout_t, origin_y_pt) == 32U);
static_assert(offsetof(fm_page_layout_t, width_pt) == 40U);
static_assert(offsetof(fm_page_layout_t, height_pt) == 48U);

static_assert(sizeof(fm_page_break_t) == 16U);
static_assert(offsetof(fm_page_break_t, id) == 0U);
static_assert(offsetof(fm_page_break_t, min) == 4U);
static_assert(offsetof(fm_page_break_t, max) == 8U);
static_assert(offsetof(fm_page_break_t, manual) == 12U);

static_assert(sizeof(fm_page_setup_t) == 48U);
static_assert(offsetof(fm_page_setup_t, orientation_engaged) == 0U);
static_assert(offsetof(fm_page_setup_t, orientation) == 4U);
static_assert(offsetof(fm_page_setup_t, paper_size_engaged) == 8U);
static_assert(offsetof(fm_page_setup_t, paper_size) == 12U);
static_assert(offsetof(fm_page_setup_t, scale_engaged) == 16U);
static_assert(offsetof(fm_page_setup_t, scale) == 20U);
static_assert(offsetof(fm_page_setup_t, fit_to_width_engaged) == 24U);
static_assert(offsetof(fm_page_setup_t, fit_to_width) == 28U);
static_assert(offsetof(fm_page_setup_t, fit_to_height_engaged) == 32U);
static_assert(offsetof(fm_page_setup_t, fit_to_height) == 36U);
static_assert(offsetof(fm_page_setup_t, fit_to_page_engaged) == 40U);
static_assert(offsetof(fm_page_setup_t, fit_to_page) == 44U);

static_assert(sizeof(fm_page_margins_t) == 96U);
static_assert(offsetof(fm_page_margins_t, left_engaged) == 0U);
static_assert(offsetof(fm_page_margins_t, left) == 8U);
static_assert(offsetof(fm_page_margins_t, right_engaged) == 16U);
static_assert(offsetof(fm_page_margins_t, right) == 24U);
static_assert(offsetof(fm_page_margins_t, top_engaged) == 32U);
static_assert(offsetof(fm_page_margins_t, top) == 40U);
static_assert(offsetof(fm_page_margins_t, bottom_engaged) == 48U);
static_assert(offsetof(fm_page_margins_t, bottom) == 56U);
static_assert(offsetof(fm_page_margins_t, header_engaged) == 64U);
static_assert(offsetof(fm_page_margins_t, header) == 72U);
static_assert(offsetof(fm_page_margins_t, footer_engaged) == 80U);
static_assert(offsetof(fm_page_margins_t, footer) == 88U);

static_assert(sizeof(fm_print_options_t) == 32U);
static_assert(offsetof(fm_print_options_t, grid_lines_engaged) == 0U);
static_assert(offsetof(fm_print_options_t, grid_lines) == 4U);
static_assert(offsetof(fm_print_options_t, headings_engaged) == 8U);
static_assert(offsetof(fm_print_options_t, headings) == 12U);
static_assert(offsetof(fm_print_options_t, horizontal_centered_engaged) == 16U);
static_assert(offsetof(fm_print_options_t, horizontal_centered) == 20U);
static_assert(offsetof(fm_print_options_t, vertical_centered_engaged) == 24U);
static_assert(offsetof(fm_print_options_t, vertical_centered) == 28U);

static_assert(sizeof(fm_cell_node_t) == 12U);
static_assert(offsetof(fm_cell_node_t, sheet) == 0U);
static_assert(offsetof(fm_cell_node_t, row) == 4U);
static_assert(offsetof(fm_cell_node_t, col) == 8U);

static_assert(sizeof(fm_pivot_cache_options_t) == 12U);
static_assert(offsetof(fm_pivot_cache_options_t, save_data) == 0U);
static_assert(offsetof(fm_pivot_cache_options_t, missing_items_limit) == 4U);
static_assert(offsetof(fm_pivot_cache_options_t, refresh_on_load) == 8U);

static_assert(sizeof(fm_spill_info_t) == 24U);
static_assert(offsetof(fm_spill_info_t, rect) == 0U);
static_assert(offsetof(fm_spill_info_t, engaged) == 16U);
static_assert(offsetof(fm_spill_info_t, blocked) == 20U);

static_assert(sizeof(fm_sheet_view_t) == 80U);
static_assert(offsetof(fm_sheet_view_t, zoom_scale) == 0U);
static_assert(offsetof(fm_sheet_view_t, freeze_rows) == 4U);
static_assert(offsetof(fm_sheet_view_t, freeze_cols) == 8U);
static_assert(offsetof(fm_sheet_view_t, visibility) == 12U);
static_assert(offsetof(fm_sheet_view_t, show_grid_lines) == 16U);
static_assert(offsetof(fm_sheet_view_t, show_row_col_headers) == 20U);
static_assert(offsetof(fm_sheet_view_t, show_zeros) == 24U);
static_assert(offsetof(fm_sheet_view_t, right_to_left) == 28U);
static_assert(offsetof(fm_sheet_view_t, tab_selected) == 32U);
static_assert(offsetof(fm_sheet_view_t, view_mode) == 36U);
static_assert(offsetof(fm_sheet_view_t, top_left_row) == 40U);
static_assert(offsetof(fm_sheet_view_t, top_left_col) == 44U);
static_assert(offsetof(fm_sheet_view_t, pane_top_left_row) == 48U);
static_assert(offsetof(fm_sheet_view_t, pane_top_left_col) == 52U);
static_assert(offsetof(fm_sheet_view_t, split_x) == 56U);
static_assert(offsetof(fm_sheet_view_t, split_y) == 64U);
static_assert(offsetof(fm_sheet_view_t, show_formulas) == 72U);
static_assert(offsetof(fm_sheet_view_t, show_outline_symbols) == 76U);

static_assert(sizeof(fm_column_layout_t) == 40U);
static_assert(offsetof(fm_column_layout_t, first) == 0U);
static_assert(offsetof(fm_column_layout_t, last) == 4U);
static_assert(offsetof(fm_column_layout_t, width) == 8U);
static_assert(offsetof(fm_column_layout_t, hidden) == 16U);
static_assert(offsetof(fm_column_layout_t, outline_level) == 20U);
static_assert(offsetof(fm_column_layout_t, collapsed) == 21U);
static_assert(offsetof(fm_column_layout_t, _pad) == 22U);
static_assert(offsetof(fm_column_layout_t, has_width) == 24U);
static_assert(offsetof(fm_column_layout_t, has_style) == 28U);
static_assert(offsetof(fm_column_layout_t, style_xf) == 32U);

static_assert(sizeof(fm_row_layout_t) == 40U);
static_assert(offsetof(fm_row_layout_t, row) == 0U);
static_assert(offsetof(fm_row_layout_t, height) == 8U);
static_assert(offsetof(fm_row_layout_t, hidden) == 16U);
static_assert(offsetof(fm_row_layout_t, outline_level) == 20U);
static_assert(offsetof(fm_row_layout_t, collapsed) == 21U);
static_assert(offsetof(fm_row_layout_t, _pad) == 22U);
static_assert(offsetof(fm_row_layout_t, has_style) == 24U);
static_assert(offsetof(fm_row_layout_t, style_xf) == 28U);
static_assert(offsetof(fm_row_layout_t, has_height) == 32U);
static_assert(offsetof(fm_row_layout_t, custom_height) == 36U);

static_assert(sizeof(fm_date_group_item_t) == 8U);
static_assert(offsetof(fm_date_group_item_t, year) == 0U);
static_assert(offsetof(fm_date_group_item_t, month) == 2U);
static_assert(offsetof(fm_date_group_item_t, day) == 3U);
static_assert(offsetof(fm_date_group_item_t, hour) == 4U);
static_assert(offsetof(fm_date_group_item_t, minute) == 5U);
static_assert(offsetof(fm_date_group_item_t, second) == 6U);
static_assert(offsetof(fm_date_group_item_t, grouping) == 7U);

static_assert(sizeof(fm_sheet_format_defaults_t) == 56U);
static_assert(offsetof(fm_sheet_format_defaults_t, default_col_width) == 0U);
static_assert(offsetof(fm_sheet_format_defaults_t, default_row_height) == 8U);
static_assert(offsetof(fm_sheet_format_defaults_t, base_col_width) == 16U);
static_assert(offsetof(fm_sheet_format_defaults_t, has_default_col_width) == 24U);
static_assert(offsetof(fm_sheet_format_defaults_t, has_default_row_height) == 28U);
static_assert(offsetof(fm_sheet_format_defaults_t, custom_height) == 32U);
static_assert(offsetof(fm_sheet_format_defaults_t, zero_height) == 36U);
static_assert(offsetof(fm_sheet_format_defaults_t, effective_default_col_width) == 40U);
static_assert(offsetof(fm_sheet_format_defaults_t, effective_default_row_height) == 48U);

static_assert(sizeof(fm_cell_xf_t) == 128U);
static_assert(offsetof(fm_cell_xf_t, font_index) == 0U);
static_assert(offsetof(fm_cell_xf_t, fill_index) == 4U);
static_assert(offsetof(fm_cell_xf_t, border_index) == 8U);
static_assert(offsetof(fm_cell_xf_t, num_fmt_id) == 12U);
static_assert(offsetof(fm_cell_xf_t, horizontal_align) == 14U);
static_assert(offsetof(fm_cell_xf_t, vertical_align) == 15U);
static_assert(offsetof(fm_cell_xf_t, wrap_text) == 16U);
static_assert(offsetof(fm_cell_xf_t, justify_last_line) == 20U);
static_assert(offsetof(fm_cell_xf_t, xf_id) == 24U);
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
static_assert(offsetof(fm_cell_xf_t, has_horizontal_align) == 72U);
static_assert(offsetof(fm_cell_xf_t, has_vertical_align) == 76U);
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
static_assert(offsetof(fm_color_spec_t, rgb) == 8U);
static_assert(offsetof(fm_color_spec_t, theme) == 12U);
static_assert(offsetof(fm_color_spec_t, indexed) == 16U);
static_assert(offsetof(fm_color_spec_t, kind) == 20U);

static_assert(sizeof(fm_gradient_stop_t) == 40U);
static_assert(offsetof(fm_gradient_stop_t, position) == 0U);
static_assert(offsetof(fm_gradient_stop_t, color_argb) == 8U);
static_assert(offsetof(fm_gradient_stop_t, color) == 16U);

static_assert(sizeof(fm_border_side_t) == 32U);
static_assert(offsetof(fm_border_side_t, style) == 0U);
static_assert(offsetof(fm_border_side_t, color_argb) == 4U);
static_assert(offsetof(fm_border_side_t, color) == 8U);

static_assert(sizeof(fm_border_record_t) == 168U);
static_assert(offsetof(fm_border_record_t, left) == 0U);
static_assert(offsetof(fm_border_record_t, right) == 32U);
static_assert(offsetof(fm_border_record_t, top) == 64U);
static_assert(offsetof(fm_border_record_t, bottom) == 96U);
static_assert(offsetof(fm_border_record_t, diagonal) == 128U);
static_assert(offsetof(fm_border_record_t, diagonal_up) == 160U);
static_assert(offsetof(fm_border_record_t, diagonal_down) == 164U);

static_assert(sizeof(fm_theme_colors_t) == 48U);
static_assert(offsetof(fm_theme_colors_t, argb) == 0U);

static_assert(sizeof(fm_image_info_t) == 12U);
static_assert(offsetof(fm_image_info_t, format) == 0U);
static_assert(offsetof(fm_image_info_t, px_width) == 4U);
static_assert(offsetof(fm_image_info_t, px_height) == 8U);

static_assert(sizeof(fm_drawing_anchor_t) == 48U);
static_assert(offsetof(fm_drawing_anchor_t, anchor_kind) == 0U);
static_assert(offsetof(fm_drawing_anchor_t, edit_as) == 4U);
static_assert(offsetof(fm_drawing_anchor_t, row) == 8U);
static_assert(offsetof(fm_drawing_anchor_t, col) == 12U);
static_assert(offsetof(fm_drawing_anchor_t, row_offset_emu) == 16U);
static_assert(offsetof(fm_drawing_anchor_t, col_offset_emu) == 24U);
static_assert(offsetof(fm_drawing_anchor_t, width_emu) == 32U);
static_assert(offsetof(fm_drawing_anchor_t, height_emu) == 40U);

static_assert(sizeof(fm_chart_fill_t) == 16U);
static_assert(offsetof(fm_chart_fill_t, has_color) == 0U);
static_assert(offsetof(fm_chart_fill_t, color) == 4U);
static_assert(offsetof(fm_chart_fill_t, transparency_pct) == 8U);

static_assert(sizeof(fm_chart_line_t) == 40U);
static_assert(offsetof(fm_chart_line_t, has_color) == 0U);
static_assert(offsetof(fm_chart_line_t, color) == 4U);
static_assert(offsetof(fm_chart_line_t, transparency_pct) == 8U);
static_assert(offsetof(fm_chart_line_t, width_pt) == 16U);
static_assert(offsetof(fm_chart_line_t, has_dash) == 24U);
static_assert(offsetof(fm_chart_line_t, dash) == 28U);
static_assert(offsetof(fm_chart_line_t, none) == 32U);

static_assert(sizeof(fm_chart_marker_t) == 72U);
static_assert(offsetof(fm_chart_marker_t, symbol) == 0U);
static_assert(offsetof(fm_chart_marker_t, has_size) == 4U);
static_assert(offsetof(fm_chart_marker_t, size) == 8U);
static_assert(offsetof(fm_chart_marker_t, fill) == 16U);
static_assert(offsetof(fm_chart_marker_t, line) == 32U);

static_assert(sizeof(fm_chart_point_t) == 144U);
static_assert(offsetof(fm_chart_point_t, idx) == 0U);
static_assert(offsetof(fm_chart_point_t, fill) == 8U);
static_assert(offsetof(fm_chart_point_t, line) == 24U);
static_assert(offsetof(fm_chart_point_t, has_marker) == 64U);
static_assert(offsetof(fm_chart_point_t, marker) == 72U);

static_assert(sizeof(fm_chart_fill_patch_t) == 24U);
static_assert(offsetof(fm_chart_fill_patch_t, color_op) == 0U);
static_assert(offsetof(fm_chart_fill_patch_t, color) == 4U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_op) == 8U);
static_assert(offsetof(fm_chart_fill_patch_t, transparency_pct) == 16U);

static_assert(sizeof(fm_chart_line_patch_t) == 56U);
static_assert(offsetof(fm_chart_line_patch_t, color_op) == 0U);
static_assert(offsetof(fm_chart_line_patch_t, color) == 4U);
static_assert(offsetof(fm_chart_line_patch_t, transparency_op) == 8U);
static_assert(offsetof(fm_chart_line_patch_t, transparency_pct) == 16U);
static_assert(offsetof(fm_chart_line_patch_t, width_op) == 24U);
static_assert(offsetof(fm_chart_line_patch_t, width_pt) == 32U);
static_assert(offsetof(fm_chart_line_patch_t, dash_op) == 40U);
static_assert(offsetof(fm_chart_line_patch_t, dash) == 44U);
static_assert(offsetof(fm_chart_line_patch_t, none_op) == 48U);
static_assert(offsetof(fm_chart_line_patch_t, none) == 52U);

static_assert(sizeof(fm_chart_marker_patch_t) == 104U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol_op) == 0U);
static_assert(offsetof(fm_chart_marker_patch_t, symbol) == 4U);
static_assert(offsetof(fm_chart_marker_patch_t, size_op) == 8U);
static_assert(offsetof(fm_chart_marker_patch_t, size) == 16U);
static_assert(offsetof(fm_chart_marker_patch_t, fill) == 24U);
static_assert(offsetof(fm_chart_marker_patch_t, line) == 48U);

static_assert(sizeof(fm_chart_point_patch_t) == 192U);
static_assert(offsetof(fm_chart_point_patch_t, fill) == 0U);
static_assert(offsetof(fm_chart_point_patch_t, line) == 24U);
static_assert(offsetof(fm_chart_point_patch_t, marker_op) == 80U);
static_assert(offsetof(fm_chart_point_patch_t, marker) == 88U);

static_assert(sizeof(fm_chart_plot_patch_t) == 40U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width_op) == 0U);
static_assert(offsetof(fm_chart_plot_patch_t, gap_width) == 4U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap_op) == 8U);
static_assert(offsetof(fm_chart_plot_patch_t, overlap) == 12U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle_op) == 16U);
static_assert(offsetof(fm_chart_plot_patch_t, first_slice_angle) == 20U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors_op) == 24U);
static_assert(offsetof(fm_chart_plot_patch_t, vary_colors) == 28U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size_op) == 32U);
static_assert(offsetof(fm_chart_plot_patch_t, hole_size) == 36U);

static_assert(sizeof(fm_chart_display_patch_t) == 24U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as_op) == 0U);
static_assert(offsetof(fm_chart_display_patch_t, blanks_as) == 4U);
static_assert(offsetof(fm_chart_display_patch_t, plot_vis_only_op) == 8U);
static_assert(offsetof(fm_chart_display_patch_t, plot_vis_only) == 12U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank_op) == 16U);
static_assert(offsetof(fm_chart_display_patch_t, na_as_blank) == 20U);

static_assert(sizeof(fm_data_table_t) == 56U);
static_assert(offsetof(fm_data_table_t, rect) == 0U);
static_assert(offsetof(fm_data_table_t, row_input_row) == 16U);
static_assert(offsetof(fm_data_table_t, row_input_col) == 20U);
static_assert(offsetof(fm_data_table_t, col_input_row) == 24U);
static_assert(offsetof(fm_data_table_t, col_input_col) == 28U);
static_assert(offsetof(fm_data_table_t, has_row_input) == 32U);
static_assert(offsetof(fm_data_table_t, has_col_input) == 36U);
static_assert(offsetof(fm_data_table_t, two_dimensional) == 40U);
static_assert(offsetof(fm_data_table_t, row_oriented) == 44U);
static_assert(offsetof(fm_data_table_t, row_input_deleted) == 48U);
static_assert(offsetof(fm_data_table_t, col_input_deleted) == 52U);
