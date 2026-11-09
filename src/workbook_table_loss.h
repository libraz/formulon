//
// What a structural edit takes from tables, and the reference transform that
// sends structured references to them to `#REF!`.

#ifndef FORMULON_WORKBOOK_TABLE_LOSS_H_
#define FORMULON_WORKBOOK_TABLE_LOSS_H_

#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "parser/ref_transforms.h"
#include "sheet_name.h"

namespace formulon {

/// A table an edit deletes (`whole`), or the columns of one it removes.
struct TableLoss {
  std::string name;
  std::string display_name;
  bool whole = false;
  std::vector<std::string> columns;
  /// Every column of the table before the edit, when `columns` is not empty: the names a column spelling resolves
  /// among.
  std::vector<std::string> all_columns;
};

/// A structural edit's reference transform that also sends a structured
/// reference to a table the edit deletes, or to a column it deletes, to
/// `#REF!`, as Excel rewrites them. A table spelling resolves among
/// `table_list` (the workbook's tables before the edit); without a list it
/// matches under the case fold.
class TableLossTransform final : public parser::RefTransform {
 public:
  TableLossTransform(const parser::RefTransform& inner, const std::vector<TableLoss>& losses,
                     const sheet_names::NameList* table_list = nullptr)
      : inner_(inner), losses_(losses), table_list_(table_list) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return inner_.apply(ref); }
  std::optional<std::pair<parser::Reference, parser::Reference>> apply_range(
      const parser::Reference& lhs, const parser::Reference& rhs) const override {
    return inner_.apply_range(lhs, rhs);
  }
  bool preserves_ref3d_coordinates() const noexcept override { return inner_.preserves_ref3d_coordinates(); }
  parser::RefCollapse collapse_form() const noexcept override { return inner_.collapse_form(); }
  std::optional<Ref3DSheetSpan> apply_ref3d_span(std::string_view begin, std::string_view end) const override {
    return inner_.apply_ref3d_span(begin, end);
  }
  std::optional<std::string_view> apply_name_sheet(std::string_view sheet) const override {
    return inner_.apply_name_sheet(sheet);
  }
  bool rebinds_name_to_book(std::string_view sheet, std::string_view name) const noexcept override {
    return inner_.rebinds_name_to_book(sheet, name);
  }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view column) const override;

 private:
  const parser::RefTransform& inner_;
  const std::vector<TableLoss>& losses_;
  const sheet_names::NameList* table_list_;
};

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_TABLE_LOSS_H_
