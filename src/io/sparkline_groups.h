//
// Typed view of the `x14:sparklineGroups` extension in a worksheet's retained
// `<extLst>` (the string `Sheet::ext_lst_xml()` holds).
//
// The XML stays the source of truth: reads parse it and writes patch the
// parsed DOM, so attributes and children the model does not carry (`uid`,
// `rightToLeft`, unknown extensions) survive an edit, and an `ext` holding
// other payloads keeps them. A group is addressed by its position among all
// groups of the sheet in document order; add and remove shift later indices.
// Row/column edits move and drop sparklines through the extension rewrite in
// `io/ext_lst_refs.h`, which this module does not duplicate.

#ifndef FORMULON_IO_SPARKLINE_GROUPS_H_
#define FORMULON_IO_SPARKLINE_GROUPS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// One sparkline: the data `formula` (`xm:f`, sheet-qualified range text
/// without `=`) and the single cell `location` (`xm:sqref`) that shows it.
struct SparklineItem {
  std::string formula;
  std::string location;
};

enum class SparklineType : std::int32_t { kLine = 0, kColumn = 1, kStacked = 2 };
enum class SparklineAxisType : std::int32_t { kIndividual = 0, kGroup = 1, kCustom = 2 };
enum class SparklineEmptyCells : std::int32_t { kGap = 0, kZero = 1, kSpan = 2 };

/// A sparkline group. A colour of kind `kNone` is an absent element.
struct SparklineGroup {
  SparklineType type = SparklineType::kLine;
  ColorSpec color_series;
  ColorSpec color_negative;
  ColorSpec color_axis;
  ColorSpec color_markers;
  ColorSpec color_first;
  ColorSpec color_last;
  ColorSpec color_high;
  ColorSpec color_low;
  bool markers = false;
  bool high = false;
  bool low = false;
  bool first = false;
  bool last = false;
  bool negative = false;
  bool display_x_axis = false;
  bool display_hidden = false;
  SparklineAxisType min_axis_type = SparklineAxisType::kIndividual;
  SparklineAxisType max_axis_type = SparklineAxisType::kIndividual;
  bool has_manual_min = false;
  bool has_manual_max = false;
  double manual_min = 0.0;
  double manual_max = 0.0;
  bool date_axis = false;
  std::string date_axis_ref;  ///< The group's own `xm:f`; empty when absent.
  double line_weight = 0.75;
  SparklineEmptyCells empty_cells_as = SparklineEmptyCells::kGap;
  std::vector<SparklineItem> items;
  /// `rightToLeft` and the revision `uid` (`{GUID}`, empty when absent): read,
  /// and written by `add_sparkline_group`; an update keeps the group's own.
  bool right_to_left = false;
  std::string uid;
};

/// The groups of `ext_lst_xml` in document order; empty for an empty or
/// unparseable extension.
std::vector<SparklineGroup> read_sparkline_groups(std::string_view ext_lst_xml);

/// Appends `group` with its `items`. Fails with `kInvalidArgument` for no
/// items, an item whose location is not one cell or whose formula is empty, a
/// non-finite number, or an out-of-range enumeration.
Expected<void, Error> add_sparkline_group(std::string& ext_lst_xml, const SparklineGroup& group);

/// Rewrites the modelled attributes, colours and date-axis reference of group
/// `index` from `group` (its `items` are ignored). Fails like
/// `add_sparkline_group`, and for an `index` past the last group.
Expected<void, Error> update_sparkline_group(std::string& ext_lst_xml, std::size_t index, const SparklineGroup& group);

/// Replaces the sparklines of group `index`. Fails like `update_sparkline_group`.
Expected<void, Error> set_group_sparklines(std::string& ext_lst_xml, std::size_t index,
                                           const std::vector<SparklineItem>& items);

/// Removes group `index`, and the `ext` and `extLst` once nothing remains in
/// them. Fails with `kInvalidArgument` for an `index` past the last group.
Expected<void, Error> remove_sparkline_group(std::string& ext_lst_xml, std::size_t index);

}  // namespace formulon::io

#endif  // FORMULON_IO_SPARKLINE_GROUPS_H_
