#include "io/data_table_xml.h"

#include <cstdint>
#include <vector>

#include "io/sqref_text.h"
#include "io/xsd_bool.h"
#include "utils/resource_budget.h"

namespace formulon::io {

namespace {

// One cell reference; a missing or malformed one leaves the input absent.
DataTableInput parse_input(std::string_view text, bool deleted) {
  DataTableInput input;
  input.deleted = deleted;
  std::vector<MergeRange> cells;
  if (!text.empty() && parse_sqref_text(text, cells) && cells.size() == 1U) {
    input.present = true;
    input.row = cells[0].first_row;
    input.col = cells[0].first_col;
  }
  return input;
}

void append_input_attr(std::string& out, const char* name, const DataTableInput& input) {
  if (!input.present) {
    return;
  }
  out.push_back(' ');
  out.append(name);
  out.append("=\"");
  out.append(format_sqref_text({MergeRange{input.row, input.col, input.row, input.col}}));
  out.push_back('"');
}

}  // namespace

bool parse_data_table(const DataTableAttrs& attrs, DataTable* out) {
  std::vector<MergeRange> ref;
  if (attrs.ref.empty() || !parse_sqref_text(attrs.ref, ref) || ref.size() != 1U) {
    return false;
  }
  DataTable table;
  table.ref = normalized(ref[0]);
  const std::uint64_t cells = static_cast<std::uint64_t>(table.ref.last_row - table.ref.first_row + 1U) *
                              (table.ref.last_col - table.ref.first_col + 1U);
  if (ResourceBudget(kMaxDataTableCells).would_exceed(cells)) {
    return false;
  }
  const auto flag = [](std::string_view text, std::uint8_t bit, std::uint8_t* explicit_attrs) {
    if (!text.empty()) {
      *explicit_attrs = static_cast<std::uint8_t>(*explicit_attrs | bit);
    }
    return parse_xsd_bool(text, false);
  };
  table.two_dimensional = flag(attrs.dt2d, kDataTableExplicitDt2D, &table.explicit_attrs);
  table.row_oriented = flag(attrs.dtr, kDataTableExplicitDtr, &table.explicit_attrs);
  const bool deleted1 = flag(attrs.del1, kDataTableExplicitDel1, &table.explicit_attrs);
  const bool deleted2 = flag(attrs.del2, kDataTableExplicitDel2, &table.explicit_attrs);
  table.always_calc = flag(attrs.ca, kDataTableExplicitCa, &table.explicit_attrs);
  table.always_calc_array = flag(attrs.aca, kDataTableExplicitAca, &table.explicit_attrs);
  table.r1 = parse_input(attrs.r1, deleted1);
  table.r2 = parse_input(attrs.r2, deleted2);
  *out = table;
  return true;
}

void append_data_table_formula(std::string& out, const DataTable& table) {
  const auto explicit_or = [&](std::uint8_t bit, bool value) { return value || (table.explicit_attrs & bit) != 0U; };
  out.append("<f t=\"dataTable\" ref=\"");
  out.append(format_sqref_text({table.ref}));
  out.push_back('"');
  if (explicit_or(kDataTableExplicitDt2D, table.two_dimensional)) {
    emit_xsd_bool_attr(out, "dt2D", table.two_dimensional);
  }
  if (explicit_or(kDataTableExplicitDtr, table.row_oriented)) {
    emit_xsd_bool_attr(out, "dtr", table.row_oriented);
  }
  append_input_attr(out, "r1", table.r1);
  append_input_attr(out, "r2", table.r2);
  if (explicit_or(kDataTableExplicitDel1, table.r1.deleted)) {
    emit_xsd_bool_attr(out, "del1", table.r1.deleted);
  }
  if (explicit_or(kDataTableExplicitDel2, table.r2.deleted)) {
    emit_xsd_bool_attr(out, "del2", table.r2.deleted);
  }
  if (explicit_or(kDataTableExplicitCa, table.always_calc)) {
    emit_xsd_bool_attr(out, "ca", table.always_calc);
  }
  if (explicit_or(kDataTableExplicitAca, table.always_calc_array)) {
    emit_xsd_bool_attr(out, "aca", table.always_calc_array);
  }
  out.append("/>");
}

}  // namespace formulon::io
