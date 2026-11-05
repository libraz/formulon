//
// Implementation of the x14 data-validation placement declared in
// `dv_overlay.h`. Like `cf_overlay.cpp`, it works on the raw `<extLst>`
// string; the undeclared `x14:` / `xm:` prefixes parse as plain names.

#include "io/dv_overlay.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/cf_writer.h"
#include "io/data_validation_attr_names.h"
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml_defs.h"
#include "io/sheet_overlay_reader.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "pugixml.hpp"
#include "utils/arena.h"

namespace formulon::io {
namespace {

constexpr const char* kX14DvExtUri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

bool FormulaReferencesOtherSheet(std::string_view formula, std::string_view own_sheet) {
  if (formula.empty()) {
    return false;
  }
  Arena arena;
  parser::Parser parser(formula, arena);
  const parser::AstNode* root = parser.parse();
  return root != nullptr && parser.errors().empty() && parser::references_other_sheet(*root, own_sheet);
}

void AppendX14Entry(std::string& out, const DataValidation& v, const parser::ExternalBookIndexer* indexer) {
  out.append("<x14:dataValidation");
  append_data_validation_attrs(out, v);
  out.append(">");
  const auto append_formula = [&](const char* tag, const std::string& formula) {
    if (formula.empty()) {
      return;
    }
    out.append("<").append(tag).append("><xm:f>");
    AppendXmlEscaped(out, storage_feature_formula(formula, indexer));
    out.append("</xm:f></").append(tag).append(">");
  };
  append_formula("x14:formula1", v.formula1);
  append_formula("x14:formula2", v.formula2);
  out.append("<xm:sqref>");
  for (std::size_t i = 0; i < v.ranges.size(); ++i) {
    if (i > 0) {
      out.push_back(' ');
    }
    AppendRangeRef(out, v.ranges[i]);
  }
  out.append("</xm:sqref></x14:dataValidation>");
}

pugi::xml_node FindContainer(const pugi::xml_node& ext_lst) {
  for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
    if (pugi::xml_node c = ext.child("x14:dataValidations"); c) {
      return c;
    }
  }
  return {};
}

}  // namespace

bool data_validation_in_x14(const DataValidation& v, std::string_view own_sheet) {
  return FormulaReferencesOtherSheet(v.formula1, own_sheet) || FormulaReferencesOtherSheet(v.formula2, own_sheet);
}

void append_data_validation_attrs(std::string& out, const DataValidation& v) {
  if (const std::string_view t = enum_name(kDataValidationTypeNames, v.type); !t.empty()) {
    append_xml_attr(out, "type", t);
  }
  if (const std::string_view op = enum_name(kDataValidationOperatorNames, v.op); !op.empty()) {
    append_xml_attr(out, "operator", op);
  }
  if (const std::string_view es = enum_name(kDataValidationErrorStyleNames, v.error_style); !es.empty()) {
    append_xml_attr(out, "errorStyle", es);
  }
  if (v.allow_blank) {
    emit_xsd_bool_attr(out, "allowBlank", true);
  }
  if (v.show_input_message) {
    emit_xsd_bool_attr(out, "showInputMessage", true);
  }
  if (v.show_error_message) {
    emit_xsd_bool_attr(out, "showErrorMessage", true);
  }
  // `showDropDown` is inverted per ECMA-376: writing "1" SUPPRESSES the
  // in-cell arrow, so it is only emitted when the arrow should be
  // hidden. Omitting the attribute preserves Excel's default (shown).
  if (!v.show_dropdown) {
    emit_xsd_bool_attr(out, "showDropDown", true);
  }
  if (!v.error_title.empty()) {
    append_xml_attr(out, "errorTitle", v.error_title);
  }
  if (!v.error_message.empty()) {
    append_xml_attr(out, "error", v.error_message);
  }
  if (!v.prompt_title.empty()) {
    append_xml_attr(out, "promptTitle", v.prompt_title);
  }
  if (!v.prompt_message.empty()) {
    append_xml_attr(out, "prompt", v.prompt_message);
  }
  out.append(v.extra_attrs);
}

std::string merge_x14_dv_entries(const std::string& ext_lst_xml, const std::vector<DataValidation>& validations,
                                 std::string_view own_sheet, const parser::ExternalBookIndexer* indexer,
                                 std::string_view container_extra_attrs) {
  std::string entries;
  std::size_t count = 0;
  for (const DataValidation& v : validations) {
    if (data_validation_in_x14(v, own_sheet)) {
      AppendX14Entry(entries, v, indexer);
      ++count;
    }
  }
  if (count == 0) {
    return ext_lst_xml;
  }

  pugi::xml_document doc;
  pugi::xml_node ext_lst;
  if (ext_lst_xml.empty()) {
    ext_lst = doc.append_child("extLst");
  } else {
    if (!doc.load_string(ext_lst_xml.c_str())) {
      return ext_lst_xml;
    }
    ext_lst = doc.child("extLst");
    if (!ext_lst) {
      return ext_lst_xml;
    }
  }

  std::string container = "<x14:dataValidations";
  append_xml_attr_uint(container, "count", count);
  container.append(" xmlns:xm=\"").append(kNsExcelMain).append("\"");
  container.append(container_extra_attrs);
  container.append(">").append(entries).append("</x14:dataValidations>");
  pugi::xml_document built;
  if (!built.load_string(container.c_str())) {
    return ext_lst_xml;
  }
  const pugi::xml_node fresh = built.child("x14:dataValidations");

  if (pugi::xml_node existing = FindContainer(ext_lst); existing) {
    for (pugi::xml_node entry = fresh.first_child(); entry; entry = entry.next_sibling()) {
      existing.append_copy(entry);
    }
    std::size_t total = 0;
    for (pugi::xml_node e = existing.child("x14:dataValidation"); e; e = e.next_sibling("x14:dataValidation")) {
      ++total;
    }
    set_attr(existing, "count", static_cast<std::int64_t>(total));
  } else {
    pugi::xml_node ext = ext_lst.append_child("ext");
    ext.append_attribute("uri") = kX14DvExtUri;
    ext.append_attribute("xmlns:x14") = kNsX14.data();
    ext.append_copy(fresh);
  }
  return raw_xml(ext_lst);
}

std::string strip_x14_data_validations(const std::string& ext_lst_xml) {
  if (ext_lst_xml.empty()) {
    return ext_lst_xml;
  }
  pugi::xml_document doc;
  const pugi::xml_parse_result rc = doc.load_string(ext_lst_xml.c_str());
  pugi::xml_node ext_lst = doc.child("extLst");
  if (!rc || !ext_lst) {
    return ext_lst_xml;
  }
  // Only entries the reader adopted leave the raw copy.
  bool changed = false;
  for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
    for (pugi::xml_node c = ext.child("x14:dataValidations"); c; c = c.next_sibling("x14:dataValidations")) {
      std::vector<pugi::xml_node> adopted;
      for (pugi::xml_node dv = c.child("x14:dataValidation"); dv; dv = dv.next_sibling("x14:dataValidation")) {
        if (data_validation_sqref_readable(dv.child("xm:sqref").text().get())) {
          adopted.push_back(dv);
        }
      }
      for (const pugi::xml_node& dv : adopted) {
        c.remove_child(dv);
        changed = true;
      }
    }
    for (pugi::xml_node c = ext.child("x14:dataValidations"); c;) {
      const pugi::xml_node next = c.next_sibling("x14:dataValidations");
      if (!c.child("x14:dataValidation")) {
        ext.remove_child(c);
      }
      c = next;
    }
  }
  if (!changed) {
    return ext_lst_xml;
  }
  for (pugi::xml_node ext = ext_lst.child("ext"); ext;) {
    const pugi::xml_node next = ext.next_sibling("ext");
    if (!has_element_child(ext)) {
      ext_lst.remove_child(ext);
    }
    ext = next;
  }
  return has_element_child(ext_lst) ? raw_xml(ext_lst) : std::string();
}

}  // namespace formulon::io
