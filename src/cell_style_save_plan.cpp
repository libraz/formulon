#include "cell_style_save_plan.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "builtin_style_names.h"
#include "eval/collation.h"
#include "new_workbook_defaults.h"
#include "row_cells.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "sheet_passthrough.h"
#include "theme.h"
#include "utils/index_sort.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kNone = 0xFFFFFFFFU;
constexpr std::uint32_t kNormalBuiltinId = 0U;
constexpr std::uint32_t kRowLevelBuiltinId = 1U;
constexpr std::uint32_t kColLevelBuiltinId = 2U;
// RowLevel_n / ColLevel_n exist for n = 1..7.

bool is_builtin(const CellStyleRecord& record) {
  return record.builtin_id != CellStyleRecord::kBuiltinIdNone;
}

bool is_outline(std::uint32_t builtin_id) {
  return builtin_id == kRowLevelBuiltinId || builtin_id == kColLevelBuiltinId;
}

// One merge group: records whose saved names are equal under kCompare.
struct Group {
  std::uint32_t first = kNone;
  std::uint32_t last = kNone;
  std::uint32_t normal_first = kNone;
  std::uint32_t normal_last = kNone;
  // Filled once every record is grouped.
  std::uint32_t rep = kNone;
  std::uint32_t rep_xf = 0;
  bool survives = false;
};

void mark(UsedCellXfs& out, std::uint32_t xf) {
  if (xf < out.used.size()) {
    out.used[xf] = true;
  }
}

}  // namespace

UsedCellXfs collect_used_cell_xfs(const Workbook& wb) {
  UsedCellXfs out;
  out.used.assign(wb.styles().cell_xfs.size(), false);
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    const Sheet& sheet = wb.sheet(si);
    if (sheet.is_opaque_ooxml_sheet()) {
      out.all_used = true;
      continue;
    }
    for (const auto& row : sheet.rows()) {
      for (const RowCells::Entry entry : row.second) {
        mark(out, entry.cell.xf_index);
      }
      for (const StyleRun& run : row.second.runs()) {
        mark(out, run.xf);
      }
    }
    for (const RowLayout& row : sheet.layout().row_overrides) {
      if (row.has_style) {
        mark(out, row.style_xf);
      }
    }
    for (const ColumnLayout& column : sheet.layout().columns) {
      if (column.has_style) {
        mark(out, column.style_xf);
      }
    }
  }
  return out;
}

CellStyleSavePlan build_cell_style_save_plan(const StylesTable& styles, const UsedCellXfs& used, ExcelProfile profile) {
  const ExcelLocale locale = profile.locale;
  const bool synthesized_xf = styles.cell_style_xfs.empty();
  const auto xf_count = static_cast<std::uint32_t>(synthesized_xf ? 1U : styles.cell_style_xfs.size());
  // Out-of-range style xf ids read as 0, as the writers clamp them.
  const auto effective_xf = [xf_count](std::uint32_t xf) { return xf < xf_count ? xf : 0U; };

  // Records in cellStyles order, a synthesized Normal first when none is present.
  std::vector<CellStyleRecord> records;
  bool has_normal = false;
  for (const CellStyleRecord& record : styles.cell_styles) {
    has_normal = has_normal || record.builtin_id == kNormalBuiltinId;
  }
  if (!has_normal) {
    CellStyleRecord normal;
    normal.name = builtin_style_name(kNormalBuiltinId, 0U, locale);
    normal.builtin_id = kNormalBuiltinId;
    records.push_back(std::move(normal));
  }
  records.insert(records.end(), styles.cell_styles.begin(), styles.cell_styles.end());
  const auto record_count = static_cast<std::uint32_t>(records.size());

  std::vector<std::string> saved_names(record_count);
  for (std::uint32_t r = 0; r < record_count; ++r) {
    const CellStyleRecord& record = records[r];
    saved_names[r] = is_builtin(record) ? builtin_style_name(record.builtin_id, record.i_level, locale) : record.name;
  }

  // Step 1: merge groups, numbered by first appearance.
  std::vector<Group> groups;
  std::vector<std::uint32_t> group_of(record_count);
  {
    std::unordered_map<std::string, std::uint32_t> by_key;
    for (std::uint32_t r = 0; r < record_count; ++r) {
      std::string key = eval::collation_key(saved_names[r], profile, eval::CollationStrength::kCompare);
      const auto inserted = by_key.emplace(std::move(key), static_cast<std::uint32_t>(groups.size()));
      if (inserted.second) {
        groups.emplace_back();
      }
      const std::uint32_t g = inserted.first->second;
      group_of[r] = g;
      Group& group = groups[g];
      if (group.first == kNone) {
        group.first = r;
      }
      group.last = r;
      if (records[r].builtin_id == kNormalBuiltinId) {
        if (group.normal_first == kNone) {
          group.normal_first = r;
        }
        group.normal_last = r;
      }
    }
  }
  const auto group_count = static_cast<std::uint32_t>(groups.size());
  std::uint32_t normal_group = 0;
  for (std::uint32_t g = 0; g < group_count; ++g) {
    Group& group = groups[g];
    // Normal represents any group it is in, with its own style xf.
    if (group.normal_first != kNone) {
      normal_group = g;
      group.rep = group.normal_first;
      group.rep_xf = effective_xf(records[group.normal_last].xf_id);
    } else {
      group.rep = group.first;
      group.rep_xf = effective_xf(records[group.last].xf_id);
    }
  }

  // Step 2: a style xf goes where the first record naming it goes; an orphan stays.
  std::vector<std::uint32_t> dest(xf_count, kNone);
  for (std::uint32_t r = 0; r < record_count; ++r) {
    const std::uint32_t x = effective_xf(records[r].xf_id);
    if (dest[x] == kNone) {
      dest[x] = groups[group_of[r]].rep_xf;
    }
  }
  std::vector<bool> orphan(xf_count, false);
  for (std::uint32_t x = 0; x < xf_count; ++x) {
    if (dest[x] == kNone) {
      dest[x] = x;
      orphan[x] = true;
    }
  }

  // Step 3: the destinations some used cell xf reaches.
  std::vector<bool> reached(xf_count, false);
  for (std::size_t c = 0; c < styles.cell_xfs.size(); ++c) {
    if (used.all_used || (c < used.used.size() && used.used[c])) {
      reached[dest[effective_xf(styles.cell_xfs[c].xf_id)]] = true;
    }
  }

  // Step 4: survivors by the representative's flags or use.
  for (Group& group : groups) {
    const CellStyleRecord& rep = records[group.rep];
    group.survives = rep.builtin_id == kNormalBuiltinId || !is_builtin(rep) || rep.custom_builtin || rep.hidden ||
                     used.all_used || reached[group.rep_xf];
  }

  // Step 5: Normal's xf, then the outline styles by (level, Row before Col), then model order.
  CellStyleSavePlan plan;
  plan.profile_ = profile;
  std::vector<std::uint32_t> new_index(xf_count, kNone);
  const auto emit = [&plan, &new_index](std::uint32_t x) {
    if (new_index[x] == kNone) {
      new_index[x] = static_cast<std::uint32_t>(plan.style_xf_order_.size());
      plan.style_xf_order_.push_back(x);
    }
  };
  emit(groups[normal_group].rep_xf);
  std::vector<std::uint32_t> outlines;
  for (std::uint32_t g = 0; g < group_count; ++g) {
    if (groups[g].survives && is_outline(records[groups[g].rep].builtin_id)) {
      outlines.push_back(g);
    }
  }
  const auto outline_less = [&groups, &records](std::uint32_t lhs, std::uint32_t rhs) {
    const CellStyleRecord& a = records[groups[lhs].rep];
    const CellStyleRecord& b = records[groups[rhs].rep];
    return a.i_level != b.i_level ? a.i_level < b.i_level : a.builtin_id < b.builtin_id;
  };
  const auto outline_index_less = [&outlines, &outline_less](std::uint32_t lhs, std::uint32_t rhs) {
    return outline_less(outlines[lhs], outlines[rhs]);
  };
  std::vector<std::uint32_t> outline_order;
  sorted_index_order(outline_order, static_cast<std::uint32_t>(outlines.size()), make_index_less(outline_index_less));
  for (const std::uint32_t i : outline_order) {
    emit(groups[outlines[i]].rep_xf);
  }
  std::vector<bool> kept_xf(xf_count, false);
  for (const Group& group : groups) {
    if (group.survives) {
      kept_xf[group.rep_xf] = true;
    }
  }
  for (std::uint32_t x = 0; x < xf_count; ++x) {
    if (kept_xf[x] || (orphan[x] && reached[x])) {
      emit(x);
    }
  }
  plan.parent_of_.resize(xf_count);
  for (std::uint32_t x = 0; x < xf_count; ++x) {
    const std::uint32_t to = new_index[dest[x]];
    plan.parent_of_[x] = to == kNone ? 0U : to;
  }

  // Step 6: survivors by saved name under the sort collation, ties in group order.
  std::vector<std::uint32_t> survivors;
  std::vector<std::string> sort_keys;
  plan.planned_names_.reserve(group_count);
  for (std::uint32_t g = 0; g < group_count; ++g) {
    plan.planned_names_.push_back(saved_names[groups[g].rep]);
    if (groups[g].survives) {
      survivors.push_back(g);
      sort_keys.push_back(eval::collation_key(saved_names[groups[g].rep], profile, eval::CollationStrength::kDynamic));
    }
  }
  const auto name_less = [&sort_keys](std::uint32_t lhs, std::uint32_t rhs) {
    return eval::collate_key_compare(sort_keys[lhs], sort_keys[rhs]) < 0;
  };
  std::vector<std::uint32_t> name_order;
  sorted_index_order(name_order, static_cast<std::uint32_t>(survivors.size()), make_index_less(name_less));
  plan.cell_styles_.reserve(survivors.size());
  for (const std::uint32_t i : name_order) {
    const Group& group = groups[survivors[i]];
    CellStyleRecord out = records[group.rep];
    out.name = saved_names[group.rep];
    out.xf_id = new_index[group.rep_xf];
    plan.cell_styles_.push_back(std::move(out));
  }

  // Step 9: identity when no record was synthesized, merged, dropped or renamed and both orders hold.
  bool identity = has_normal && !synthesized_xf && group_count == record_count && survivors.size() == group_count &&
                  plan.style_xf_order_.size() == xf_count;
  for (std::uint32_t i = 0; identity && i < name_order.size(); ++i) {
    identity = name_order[i] == i && saved_names[i] == records[i].name;
  }
  for (std::uint32_t x = 0; identity && x < xf_count; ++x) {
    identity = plan.style_xf_order_[x] == x;
  }
  // A table with no named styles has nothing to plan; the writers' default Normal is not a change.
  plan.identity_ = identity || (styles.cell_styles.empty() && styles.cell_style_xfs.empty());
  return plan;
}

bool saved_file_names_cell_style(const CellStyleSavePlan& plan, std::string_view name) {
  const eval::CollateQuery query(name, plan.profile(), eval::CollationStrength::kCompare);
  for (const std::string& planned : plan.planned_names_) {
    if (query.compare(planned) == 0) {
      return true;
    }
  }
  return false;
}

namespace {

constexpr std::uint8_t kSchemeMajor = 1U;
constexpr std::uint8_t kSchemeMinor = 2U;

// The theme script whose face Excel writes for a theme-linked font, by UI
// locale; "" for the Latin face. Mac 16.115; Windows reuses the Mac rows.
const char* theme_font_script(ExcelLocale locale) {
  switch (locale) {
    case ExcelLocale::kJaJP:
      return "Jpan";
    case ExcelLocale::kZhCN:
      return "Hans";
    case ExcelLocale::kZhTW:
      return "Hant";
    case ExcelLocale::kKoKR:
      return "Hang";
    case ExcelLocale::kThTH:
      return "Thai";
    default:
      return "";
  }
}

// `<phoneticPr>` ordinals (`PhoneticProperties`) and the phonetic font's `<family>`.
constexpr std::uint8_t kPhoneticFullwidthKatakana = 1U;
constexpr std::uint8_t kPhoneticNoConversion = 3U;
constexpr std::uint8_t kPhoneticAlignLeft = 1U;
constexpr std::uint8_t kPhoneticFontFamily = 3U;

bool any_sheet_without_phonetic(const Workbook& wb) {
  const auto slot = static_cast<std::uint32_t>(worksheet_child::slot_of("phoneticPr"));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const WorksheetRawExtensions& children = wb.sheet(i).raw_extensions();
    if (std::none_of(children.begin(), children.end(),
                     [slot](const WorksheetRawChild& child) { return child.slot == slot; })) {
      return true;
    }
  }
  return false;
}

// True when `font` writes the same `<font>` as `phonetic`, which states only a face, size, family, charset and scheme.
bool writes_phonetic_font(const FontRecord& font, const FontRecord& phonetic) {
  return font.name == phonetic.name && font.size == phonetic.size && font.has_family &&
         font.family == phonetic.family && font.has_charset && font.charset == phonetic.charset &&
         font.scheme == phonetic.scheme && !font.bold && !font.italic && !font.strike && !font.outline &&
         !font.shadow && !font.condense && !font.extend && font.underline == 0U && font.vert_align == 0U &&
         font.color.kind == ColorSpec::Kind::kNone && font.color_argb == 0xFF000000U;
}

// The theme's major or minor face for `script`; null when the theme names none.
const std::string* theme_face(const Theme& theme, bool major, const char* script) {
  if (*script == '\0') {
    return &(major ? theme.fonts.major_latin : theme.fonts.minor_latin);
  }
  for (const ThemeScriptFont& font : major ? theme.major_scripts : theme.minor_scripts) {
    if (font.script == script) {
      return &font.typeface;
    }
  }
  return nullptr;
}

}  // namespace

FontSavePlan build_font_save_plan(const Workbook& wb) {
  FontSavePlan plan;
  plan.fonts_ = wb.styles().fonts;
  const ExcelLocale locale = wb.excel_profile().locale;
  const LoadedTheme loaded = wb.load_theme();
  const char* script = theme_font_script(locale);
  for (FontRecord& font : plan.fonts_) {
    if (font.scheme != kSchemeMajor && font.scheme != kSchemeMinor) {
      continue;
    }
    const std::string* face = theme_face(loaded.theme, font.scheme == kSchemeMajor, script);
    if (face != nullptr && !face->empty() && *face != font.name) {
      font.name = *face;
      plan.identity_ = false;
    }
  }

  const NewWorkbookDefaults& row = new_workbook_defaults(locale);
  if (row.phonetic_size == 0U || plan.fonts_.empty() || plan.fonts_[0].scheme == 0U ||
      !any_sheet_without_phonetic(wb)) {
    return plan;
  }
  FontRecord phonetic;
  phonetic.name = plan.fonts_[0].name;
  phonetic.size = row.phonetic_size;
  phonetic.has_family = true;
  phonetic.family = kPhoneticFontFamily;
  phonetic.has_charset = true;
  phonetic.charset = row.charset;
  phonetic.scheme = plan.fonts_[0].scheme;
  std::size_t id = 0;
  while (id < plan.fonts_.size() && !writes_phonetic_font(plan.fonts_[id], phonetic)) {
    ++id;
  }
  if (id == plan.fonts_.size()) {
    plan.fonts_.push_back(std::move(phonetic));
    plan.identity_ = false;
  }
  PhoneticProperties props;
  props.font_id = static_cast<std::uint16_t>(id);
  props.type = row.phonetic_no_conversion ? kPhoneticNoConversion : kPhoneticFullwidthKatakana;
  props.alignment = kPhoneticAlignLeft;
  plan.absent_sheet_phonetic_ = props;
  return plan;
}

}  // namespace formulon
