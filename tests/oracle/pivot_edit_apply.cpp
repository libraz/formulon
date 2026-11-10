// Applies one pivot-edit oracle case through the C API: loads before.xlsx,
// runs each `--op` against Sheet2's first pivot, projects it if pending, and
// saves. Fields and items are named by their labels.
//
// Ops (UTF-8, ':'-separated):
//   axis:<field>:row|column|page|absent   data_add:<field>:sum
//   date_group:<field>:months             num_group:<field>:<interval>
//   label_begins:<field>:<text>           top_count:<field>:<n>:<data_field>
//   page_only:<field>:<item>              hide_item:<field>:<item>
//   subtotal_location:top|bottom          default_subtotal:<field>:0|1
//   layout:compact|tabular|outline        grand_rows:0|1
//   collapse:<field>:<item>               refresh
//
// Exit 0 on success, 2 on usage, 3 when the engine refuses an op, 4 on I/O.

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <iterator>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace {

constexpr std::size_t kNotFound = static_cast<std::size_t>(-1);

std::vector<std::string> split(const std::string& text) {
  std::vector<std::string> parts;
  std::size_t start = 0;
  for (;;) {
    const std::size_t colon = text.find(':', start);
    parts.push_back(text.substr(start, colon == std::string::npos ? std::string::npos : colon - start));
    if (colon == std::string::npos) {
      return parts;
    }
    start = colon + 1;
  }
}

struct Target {
  fm_workbook_t* wb = nullptr;
  std::size_t sheet = 0;
  std::size_t pivot = 0;
};

std::size_t find_sheet(const fm_workbook_t* book, const char* name) {
  std::size_t count = 0;
  if (fm_workbook_get_sheet_count(book, &count) != FM_STATUS_OK) {
    return kNotFound;
  }
  for (std::size_t i = 0; i < count; ++i) {
    const char* got = nullptr;
    if (fm_workbook_get_sheet_name(book, i, &got) == FM_STATUS_OK && got != nullptr && std::strcmp(got, name) == 0) {
      return i;
    }
  }
  return kNotFound;
}

std::size_t find_field(const Target& target, const std::string& name) {
  std::size_t count = 0;
  if (fm_sheet_get_pivot_field_count(target.wb, target.sheet, target.pivot, &count) != FM_STATUS_OK) {
    return kNotFound;
  }
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_field_info_t info{};
    info.struct_size = sizeof(info);
    if (fm_sheet_get_pivot_field_at(target.wb, target.sheet, target.pivot, i, &info) == FM_STATUS_OK &&
        info.source_name != nullptr && name == info.source_name) {
      return i;
    }
  }
  return kNotFound;
}

std::size_t item_count(const Target& target, std::size_t field) {
  std::size_t count = 0;
  return fm_sheet_get_pivot_field_item_count(target.wb, target.sheet, target.pivot, field, &count) == FM_STATUS_OK
             ? count
             : 0;
}

std::size_t find_item(const Target& target, std::size_t field, const std::string& name) {
  const std::size_t count = item_count(target, field);
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_item_info_t info{};
    info.struct_size = sizeof(info);
    if (fm_sheet_get_pivot_field_item_at(target.wb, target.sheet, target.pivot, field, i, &info) == FM_STATUS_OK &&
        info.name != nullptr && name == info.name) {
      return i;
    }
  }
  return kNotFound;
}

fm_status_t patch_item(const Target& target, std::size_t field, std::size_t item, bool visible_op, int32_t value) {
  fm_pivot_item_patch_t patch{};
  patch.struct_size = sizeof(patch);
  if (visible_op) {
    patch.visible_op = FM_PATCH_OP_SET;
    patch.visible = value;
  } else {
    patch.show_detail_op = FM_PATCH_OP_SET;
    patch.show_detail = value;
  }
  return fm_sheet_update_pivot_field_item(target.wb, target.sheet, target.pivot, field, item, &patch);
}

int32_t axis_of(const std::string& name) {
  if (name == "row") {
    return FM_PIVOT_AXIS_ROW;
  }
  if (name == "column") {
    return FM_PIVOT_AXIS_COLUMN;
  }
  if (name == "page") {
    return FM_PIVOT_AXIS_PAGE;
  }
  if (name == "absent") {
    return FM_PIVOT_AXIS_ABSENT;
  }
  return -2;
}

int32_t layout_of(const std::string& name) {
  if (name == "compact") {
    return FM_PIVOT_REPORT_LAYOUT_COMPACT;
  }
  if (name == "tabular") {
    return FM_PIVOT_REPORT_LAYOUT_TABULAR;
  }
  if (name == "outline") {
    return FM_PIVOT_REPORT_LAYOUT_OUTLINE;
  }
  return -1;
}

// Returns FM_STATUS_OK, an engine status, or -1 for an op this tool cannot parse.
int apply_op(const Target& target, const std::string& text) {
  const std::vector<std::string> args = split(text);
  const std::string& kind = args[0];
  if (kind == "refresh" && args.size() == 1) {
    uint32_t cache_id = 0;
    const fm_status_t status = fm_sheet_get_pivot_cache_id(target.wb, target.sheet, target.pivot, &cache_id);
    return status != FM_STATUS_OK ? status : fm_workbook_refresh_pivot_cache(target.wb, cache_id);
  }
  if (kind == "subtotal_location" && args.size() == 2 && (args[1] == "top" || args[1] == "bottom")) {
    // A pivot-wide setting: every field takes it.
    std::size_t count = 0;
    fm_status_t status = fm_sheet_get_pivot_field_count(target.wb, target.sheet, target.pivot, &count);
    fm_pivot_field_patch_t patch{};
    patch.struct_size = sizeof(patch);
    patch.subtotal_top_op = FM_PATCH_OP_SET;
    patch.subtotal_top = args[1] == "top" ? 1 : 0;
    for (std::size_t i = 0; status == FM_STATUS_OK && i < count; ++i) {
      status = fm_sheet_update_pivot_field(target.wb, target.sheet, target.pivot, i, &patch);
    }
    return status;
  }
  if (kind == "layout" && args.size() == 2 && layout_of(args[1]) >= 0) {
    return fm_sheet_set_pivot_report_layout(target.wb, target.sheet, target.pivot, layout_of(args[1]));
  }
  if (kind == "grand_rows" && args.size() == 2) {
    fm_pivot_grand_totals_patch_t patch{};
    patch.struct_size = sizeof(patch);
    patch.rows_op = FM_PATCH_OP_SET;
    patch.rows = std::atoi(args[1].c_str());
    return fm_sheet_update_pivot_grand_totals(target.wb, target.sheet, target.pivot, &patch);
  }
  if (args.size() < 3) {
    return -1;
  }
  const std::size_t field = find_field(target, args[1]);
  if (field == kNotFound) {
    std::fprintf(stderr, "no pivot field %s\n", args[1].c_str());
    return -1;
  }
  if (kind == "axis" && axis_of(args[2]) != -2) {
    return fm_sheet_set_pivot_field_axis(target.wb, target.sheet, target.pivot, field, axis_of(args[2]));
  }
  if (kind == "data_add" && args[2] == "sum") {
    fm_pivot_data_field_t spec{};
    spec.struct_size = sizeof(spec);
    spec.field_index = static_cast<uint32_t>(field);
    spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
    std::size_t index = 0;
    return fm_sheet_add_pivot_data_field(target.wb, target.sheet, target.pivot, &spec, &index);
  }
  if (kind == "date_group" && args[2] == "months") {
    fm_pivot_date_group_spec_t spec{};
    spec.struct_size = sizeof(spec);
    spec.periods = FM_PIVOT_PERIOD_MONTHS;
    spec.calendar = FM_PIVOT_CALENDAR_GREGORIAN;
    spec.interval_days = 1;
    return fm_sheet_set_pivot_field_date_group(target.wb, target.sheet, target.pivot, field, &spec);
  }
  if (kind == "num_group") {
    fm_pivot_numeric_group_spec_t spec{};
    spec.struct_size = sizeof(spec);
    spec.interval = std::atof(args[2].c_str());
    return fm_sheet_set_pivot_field_numeric_group(target.wb, target.sheet, target.pivot, field, &spec);
  }
  if (kind == "label_begins" || kind == "top_count") {
    fm_pivot_filter_t spec{};
    spec.struct_size = sizeof(spec);
    spec.field_index = static_cast<uint32_t>(field);
    spec.value_high_kind = FM_PIVOT_FILTER_OPERAND_KIND_ABSENT;
    if (kind == "label_begins") {
      spec.kind = FM_PIVOT_FILTER_KIND_LABEL_BEGINS_WITH;
      spec.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
      spec.value_text = args[2].c_str();
    } else {
      if (args.size() != 4) {
        return -1;
      }
      spec.kind = FM_PIVOT_FILTER_KIND_VALUE_TOP_COUNT;
      spec.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_INT;
      spec.value_int = std::atoi(args[2].c_str());
      spec.data_field_index = static_cast<uint32_t>(std::atoi(args[3].c_str()));
    }
    std::size_t index = 0;
    return fm_sheet_add_pivot_filter(target.wb, target.sheet, target.pivot, &spec, &index);
  }
  if (kind == "default_subtotal") {
    fm_pivot_field_patch_t patch{};
    patch.struct_size = sizeof(patch);
    patch.default_subtotal_op = FM_PATCH_OP_SET;
    patch.default_subtotal = std::atoi(args[2].c_str());
    return fm_sheet_update_pivot_field(target.wb, target.sheet, target.pivot, field, &patch);
  }
  if (kind == "hide_item" || kind == "collapse" || kind == "page_only") {
    const std::size_t item = find_item(target, field, args[2]);
    if (item == kNotFound) {
      std::fprintf(stderr, "no item %s in field %s\n", args[2].c_str(), args[1].c_str());
      return -1;
    }
    if (kind == "hide_item") {
      return patch_item(target, field, item, true, 0);
    }
    if (kind == "collapse") {
      return patch_item(target, field, item, false, 0);
    }
    // The C API has no current-page setter; one selected page item is every other item hidden.
    const std::size_t count = item_count(target, field);
    for (std::size_t i = 0; i < count; ++i) {
      if (i == item) {
        continue;
      }
      const fm_status_t status = patch_item(target, field, i, true, 0);
      if (status != FM_STATUS_OK) {
        return status;
      }
    }
    return FM_STATUS_OK;
  }
  return -1;
}

bool read_file(const char* path, std::vector<std::uint8_t>* out) {
  std::ifstream in(path, std::ios::binary);
  if (!in) {
    return false;
  }
  out->assign(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
  return true;
}

bool write_file(const char* path, const std::uint8_t* bytes, std::size_t len) {
  std::ofstream out(path, std::ios::binary);
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  return static_cast<bool>(out);
}

int usage(const char* argv0) {
  std::fprintf(stderr, "usage: %s <before.xlsx> <out.xlsx> --profile <id> [--op <op>]...\n", argv0);
  return 2;
}

}  // namespace

int main(int argc, char** argv) {
  if (argc < 3) {
    return usage(argv[0]);
  }
  const char* profile = nullptr;
  std::vector<std::string> ops;
  for (int i = 3; i < argc; ++i) {
    if (std::strcmp(argv[i], "--profile") == 0 && i + 1 < argc) {
      profile = argv[++i];
    } else if (std::strcmp(argv[i], "--op") == 0 && i + 1 < argc) {
      ops.emplace_back(argv[++i]);
    } else {
      return usage(argv[0]);
    }
  }
  std::vector<std::uint8_t> input;
  if (!read_file(argv[1], &input)) {
    std::fprintf(stderr, "cannot read %s\n", argv[1]);
    return 4;
  }
  Target target;
  if (fm_workbook_load(input.data(), input.size(), &target.wb) != FM_STATUS_OK) {
    std::fprintf(stderr, "load failed: %s\n", argv[1]);
    return 4;
  }
  if (profile != nullptr && fm_workbook_set_profile_id(target.wb, profile) != FM_STATUS_OK) {
    std::fprintf(stderr, "unknown profile %s\n", profile);
    fm_workbook_destroy(target.wb);
    return 2;
  }
  target.sheet = find_sheet(target.wb, "Sheet2");
  if (target.sheet == kNotFound) {
    std::fprintf(stderr, "no Sheet2\n");
    fm_workbook_destroy(target.wb);
    return 4;
  }
  int exit_code = 0;
  for (const std::string& op : ops) {
    const int status = apply_op(target, op);
    if (status == -1) {
      std::fprintf(stderr, "bad op %s\n", op.c_str());
      fm_workbook_destroy(target.wb);
      return 2;
    }
    if (status != FM_STATUS_OK) {
      std::printf("op-status %s %d\n", op.c_str(), status);
      exit_code = 3;
    }
  }
  int32_t pending = 0;
  if (fm_sheet_is_pivot_projection_pending(target.wb, target.sheet, target.pivot, &pending) == FM_STATUS_OK &&
      pending != 0) {
    const fm_status_t status = fm_sheet_project_pivot(target.wb, target.sheet, target.pivot);
    if (status != FM_STATUS_OK) {
      std::printf("project-status %d\n", status);
      exit_code = 3;
    }
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  const bool saved = fm_workbook_save(target.wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len) == FM_STATUS_OK &&
                     write_file(argv[2], bytes, len);
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(target.wb);
  if (!saved) {
    std::fprintf(stderr, "save failed: %s\n", argv[2]);
    return 4;
  }
  return exit_code;
}
