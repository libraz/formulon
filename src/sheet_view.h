//
// Sheet tab visibility, view, selection and protection state.

#ifndef FORMULON_SHEET_VIEW_H_
#define FORMULON_SHEET_VIEW_H_

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace formulon {

/// Tab visibility, i.e. OOXML `ST_SheetState` (ECMA-376 §18.18.68) and
/// the `hsState` field of the XLSB `BrtBundleSh` record.
///
/// The two containers spell the same three states, and the enumerator
/// values are the `hsState` numbering so the binary path needs no table.
/// `kVeryHidden` differs from `kHidden` in that Excel leaves such a sheet
/// out of the "Unhide" dialog, which is how a workbook keeps a settings
/// or lookup sheet out of a user's reach; collapsing it to `kHidden`
/// weakens that on the author's behalf.
enum class SheetVisibility : std::uint8_t {
  kVisible = 0,
  kHidden = 1,
  kVeryHidden = 2,
};

/// One `<selection>` of a `<sheetView>`: the active cell and selected ranges
/// of one pane. Stored as the OOXML spellings so unusual `sqref` forms
/// (whole rows or columns, several ranges) round-trip verbatim.
struct SheetSelection {
  std::string pane;                  ///< `pane`; empty is the schema default `topLeft`, as is "topLeft".
  std::string active_cell;           ///< `activeCell` (A1 form); empty when absent.
  std::string sqref;                 ///< `sqref`, space-separated ranges; empty is the active cell alone.
  std::uint32_t active_cell_id = 0;  ///< `activeCellId`: the `sqref` range holding the active cell.
};

/// Per-sheet view state (zoom, frozen panes, tab visibility).
///
/// Mirrors the OOXML `<sheetView>` and `<sheetPr>` attributes that survive
/// a save/load round-trip without affecting evaluation. The schema may be
/// extended by later bundles, but the field names declared here are stable.
struct SheetView {
  /// Default zoom percentage when no `<sheetView zoomScale="...">` was
  /// stored. Matches Excel's "100%" default.
  static constexpr std::uint32_t kDefaultZoomScale = 100U;

  std::uint32_t zoom_scale = kDefaultZoomScale;
  std::uint32_t freeze_rows = 0;  // 0 = no row freeze
  std::uint32_t freeze_cols = 0;  // 0 = no column freeze
  /// A frozen view saved as `frozenSplit`: a freeze that also keeps a
  /// movable split, restored when the panes are unfrozen in Excel.
  bool frozen_split = false;
  /// An unfrozen split's position in twips (`<pane xSplit ySplit>` with no
  /// freeze state); 0 for no split on that axis. Ignored while frozen.
  double split_x = 0.0;
  double split_y = 0.0;
  bool tab_hidden = false;
  /// Refines `tab_hidden` to OOXML `veryHidden`. Visibility is three
  /// states but `tab_hidden` is the one bit every binding reads, so the
  /// third state rides alongside it instead of replacing it: a caller
  /// that knows only the bool still sees a very-hidden sheet as hidden.
  /// Never read this field on its own — `visibility()` resolves the pair
  /// to exactly one state, including the combination no writer should
  /// produce.
  bool tab_very_hidden = false;
  // `<sheetView>` display attributes. Names mirror the OOXML attributes;
  // the values shown are the ECMA-376 §18.3.1.87 schema defaults, so a
  // freshly created sheet round-trips without emitting them. Losing these
  // on save silently re-shows hidden gridlines/headers, drops the active
  // tab, and resets the view mode.
  bool show_grid_lines = true;       ///< `showGridLines` (default true).
  bool show_row_col_headers = true;  ///< `showRowColHeaders` (default true).
  bool show_zeros = true;            ///< `showZeros` (default true).
  bool right_to_left = false;        ///< `rightToLeft` (default false).
  bool tab_selected = false;         ///< `tabSelected` (default false).
  bool window_protection = false;    ///< `windowProtection` (default false).
  bool show_formulas = false;        ///< `showFormulas`: cells show their formula text (default false).
  bool show_ruler = true;            ///< `showRuler` (default true).
  bool show_outline_symbols = true;  ///< `showOutlineSymbols` (default true).
  bool default_grid_color = true;    ///< `defaultGridColor` (default true).
  bool show_white_space = true;      ///< `showWhiteSpace` (default true).
  /// `colorId`: palette index of a custom grid colour (default 64, automatic).
  static constexpr std::uint32_t kDefaultColorId = 64U;
  std::uint32_t color_id = kDefaultColorId;
  /// `zoomScaleNormal` / `zoomScaleSheetLayoutView` / `zoomScalePageLayoutView`;
  /// 0 is absent.
  std::uint32_t zoom_scale_normal = 0;
  std::uint32_t zoom_scale_sheet_layout_view = 0;
  std::uint32_t zoom_scale_page_layout_view = 0;
  /// Raw `<pivotSelection>` and `<extLst>` children of `<sheetView>`, kept
  /// verbatim and re-emitted after the selections.
  std::string tail_xml;
  /// Raw `<sheetView>` elements after the first (one per extra workbook
  /// window), kept verbatim and re-emitted after the modelled view.
  std::string extra_views_xml;
  /// `view` mode: empty (== "normal"), "pageBreakPreview", or
  /// "pageLayout". Stored verbatim so unknown future values round-trip.
  std::string view_mode;
  /// `<sheetView topLeftCell>`: the scroll origin of the top-left pane. Empty
  /// is the default `A1`.
  std::string top_left_cell;
  /// `<pane topLeftCell>` and `activePane` of a frozen or split view. Empty
  /// means derived: the first unfrozen cell (A1 for a split), and the pane
  /// the splits imply (`bottomRight`, `topRight` or `bottomLeft`). A frozen
  /// view's active pane is always derived, as Excel saves it; a split's
  /// is derived only when it names a pane the view lacks.
  std::string pane_top_left_cell;
  std::string active_pane;
  /// Every `<selection>`, in document order (at most one per pane).
  std::vector<SheetSelection> selections;

  /// Panes compare by name with the default `topLeft` spelled out.
  static std::string_view pane_key(std::string_view pane) noexcept {
    return pane.empty() ? std::string_view("topLeft") : pane;
  }

  bool is_frozen() const noexcept { return freeze_rows != 0U || freeze_cols != 0U; }
  /// Whether a freeze or split divides the rows / the columns.
  bool divides_rows() const noexcept { return is_frozen() ? freeze_rows != 0U : split_y > 0.0; }
  bool divides_cols() const noexcept { return is_frozen() ? freeze_cols != 0U : split_x > 0.0; }

  /// Whether the view has `pane`: `topLeft` always, the others only on the
  /// axes a freeze or split divides. Excel refuses a file with a selection in
  /// a pane the view lacks.
  bool has_pane(std::string_view pane) const noexcept {
    const std::string_view key = pane_key(pane);
    if (key == "topLeft") {
      return true;
    }
    if (key == "topRight") {
      return divides_cols();
    }
    if (key == "bottomLeft") {
      return divides_rows();
    }
    return key == "bottomRight" && divides_rows() && divides_cols();
  }

  /// The pane a frozen or split view is active in: the pane the splits
  /// imply, or a split's `active_pane` when the view has it. Empty without
  /// panes.
  std::string_view effective_active_pane() const noexcept {
    if (!divides_rows() && !divides_cols()) {
      return {};
    }
    if (!is_frozen() && !active_pane.empty() && has_pane(active_pane)) {
      return active_pane;
    }
    return divides_rows() ? (divides_cols() ? "bottomRight" : "bottomLeft") : "topRight";
  }

  /// Freezes `rows` x `cols` (zero for none). A change replaces any split,
  /// re-derives the pane origin and active pane, and re-homes the active
  /// selection to the new active pane; the other panes' selections are dropped.
  void set_freeze(std::uint32_t rows, std::uint32_t cols) {
    if (rows == freeze_rows && cols == freeze_cols) {
      return;
    }
    replace_panes([&] {
      freeze_rows = rows;
      freeze_cols = cols;
      split_x = 0.0;
      split_y = 0.0;
    });
  }

  /// Splits the view at `x` / `y` twips (zero for none on that axis; both zero
  /// removes the split). A negative or non-finite value is rejected and leaves
  /// the view unchanged. A change replaces any freeze and re-derives the panes
  /// as `set_freeze` does.
  /// @return false when an argument was rejected.
  bool set_split(double x, double y) {
    if (!std::isfinite(x) || !std::isfinite(y) || x < 0.0 || y < 0.0) {
      return false;
    }
    if (!is_frozen() && x == split_x && y == split_y) {
      return true;
    }
    replace_panes([&] {
      freeze_rows = 0;
      freeze_cols = 0;
      split_x = x;
      split_y = y;
    });
    return true;
  }

  /// The one visibility state the two bits stand for.
  ///
  /// `tab_very_hidden` decides first, so a caller that set it without
  /// `tab_hidden` still gets `kVeryHidden` rather than a sheet that
  /// silently became visible. Every writer goes through here, which is
  /// what keeps the saved state equal to the loaded one for all three
  /// values.
  SheetVisibility visibility() const noexcept {
    if (tab_very_hidden) {
      return SheetVisibility::kVeryHidden;
    }
    return tab_hidden ? SheetVisibility::kHidden : SheetVisibility::kVisible;
  }

  /// Sets both bits from one state, so they cannot drift apart.
  void set_visibility(SheetVisibility state) noexcept {
    tab_hidden = state != SheetVisibility::kVisible;
    tab_very_hidden = state == SheetVisibility::kVeryHidden;
  }

 private:
  /// Runs `change` on the pane layout, then re-derives the pane origin and
  /// active pane and re-homes the active selection to the new active pane.
  template <typename Change>
  void replace_panes(Change&& change) {
    std::vector<SheetSelection> kept;
    for (SheetSelection& selection : selections) {
      if (pane_key(selection.pane) == pane_key(effective_active_pane())) {
        kept.push_back(std::move(selection));
        break;
      }
    }
    change();
    frozen_split = false;
    pane_top_left_cell.clear();
    active_pane.clear();
    selections = std::move(kept);
    for (SheetSelection& selection : selections) {
      selection.pane = pane_key(effective_active_pane());
    }
  }
};

/// Mirror of OOXML `<sheetProtection>` (ECMA-376 §18.3.1.85). Stored as
/// passive round-trip metadata: the engine does not enforce locks at
/// evaluation time. The host UI inspects these flags to mirror Excel's
/// "Protect Sheet" dialog state.
///
/// Field semantics follow the OOXML attribute names verbatim. Each
/// boolean represents the attribute's value as it appears in the
/// document; `false` means "attribute absent or `0`". Each attribute's
/// effect (whether `true` blocks or allows the operation) is dictated
/// by the spec, not inverted here.
///
/// `enabled` is the only synthetic field: it tracks whether the
/// `<sheetProtection>` element is present at all. A workbook with
/// `enabled == false` round-trips with no element emitted, regardless
/// of the other field values.
struct SheetProtection {
  /// True when the `<sheetProtection>` element is present.
  bool enabled = false;
  /// Hash algorithm name (e.g. "SHA-512"). Empty when no modern
  /// password is set.
  std::string algorithm_name;
  /// Base64-encoded password hash. Empty when no modern password is
  /// set.
  std::string hash_value;
  /// Base64-encoded salt for the password hash.
  std::string salt_value;
  /// Iteration count for the password hash.
  std::uint32_t spin_count = 0;
  /// Legacy 16-bit hashed password, written as a hex string (e.g.
  /// `"CC53"`). Empty when no legacy password is set. Excel writes
  /// either the legacy or the modern password attributes, not both.
  std::string legacy_password;

  // OOXML attribute flags. Names mirror the attributes verbatim; see
  // ECMA-376 §18.3.1.85 for per-attribute defaults and effects. A flag of
  // `true` means the corresponding action is LOCKED. Eleven of these
  // (format*, insert*, delete*, sort, autoFilter, pivotTables) default to
  // `true` in the schema, so the member initialisers match those defaults:
  // a default-constructed protection that is then `enabled` mirrors Excel's
  // "Protect Sheet" defaults. The five action flags whose schema default is
  // `false` (sheet, objects, scenarios, select*) stay false.
  bool sheet = false;
  bool objects = false;
  bool scenarios = false;
  bool format_cells = true;
  bool format_columns = true;
  bool format_rows = true;
  bool insert_columns = true;
  bool insert_rows = true;
  bool insert_hyperlinks = true;
  bool delete_columns = true;
  bool delete_rows = true;
  bool select_locked_cells = false;
  bool select_unlocked_cells = false;
  bool sort = true;
  bool auto_filter = true;
  bool pivot_tables = true;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_VIEW_H_
