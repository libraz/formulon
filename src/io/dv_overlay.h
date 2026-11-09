//
// Placement of data validations between the worksheet `<dataValidations>`
// element and the x14 extension (`<x14:dataValidations>`). Excel stores a
// validation whose formulas name another sheet in the extension alone; the
// reader moves such entries into the model and the writer rebuilds them from it.

#ifndef FORMULON_IO_DV_OVERLAY_H_
#define FORMULON_IO_DV_OVERLAY_H_

#include <string>
#include <string_view>
#include <vector>

#include "sheet_metadata.h"

namespace formulon::parser {
struct ExternalBookIndexer;
}  // namespace formulon::parser

namespace formulon::io {

/// True when `v` is stored in the x14 extension: it was read from there, or a
/// formula names a sheet other than `own_sheet` (see
/// `parser::references_other_sheet`, which resolves among `sheet_list`).
bool data_validation_in_x14(const DataValidation& v, std::string_view own_sheet,
                            const std::vector<std::string>* sheet_list = nullptr);

/// Appends the attributes of `v` shared by both forms (everything but
/// `sqref`), in Excel's order.
void append_data_validation_attrs(std::string& out, const DataValidation& v);

/// Writes the validations of `validations` that `data_validation_in_x14`
/// selects into the `{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}` extension of
/// `ext_lst_xml`, appended after the other extensions when it is absent, and
/// returns the merged `<extLst>`. `container_extra_attrs` (escaped
/// ` name="value"` pairs) go on `<x14:dataValidations>`. Returns the input
/// unchanged when no validation is selected or it does not parse.
std::string merge_x14_dv_entries(const std::string& ext_lst_xml, const std::vector<DataValidation>& validations,
                                 std::string_view own_sheet, const parser::ExternalBookIndexer* indexer = nullptr,
                                 std::string_view container_extra_attrs = {},
                                 const std::vector<std::string>* sheet_list = nullptr);

/// Removes every `<x14:dataValidations>` -- the entries the reader moved into
/// the model -- and the `<ext>` elements that leaves empty. Returns the input
/// unchanged when there is none or it does not parse, and an empty string
/// when nothing survives.
std::string strip_x14_data_validations(const std::string& ext_lst_xml);

}  // namespace formulon::io

#endif  // FORMULON_IO_DV_OVERLAY_H_
