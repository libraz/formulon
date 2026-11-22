//
// Shared XML-fragment helpers for the print-settings entry points, which
// live in print_settings.cpp and sheet_tab_color.cpp.

#ifndef FORMULON_C_API_PARTS_PRINT_SETTINGS_INTERNAL_H_
#define FORMULON_C_API_PARTS_PRINT_SETTINGS_INTERNAL_H_

#include <cstdint>
#include <string_view>

#include "pugixml.hpp"

namespace formulon {
namespace c_api {
namespace parts {

/// Loads `fragment` into `doc` and returns its root, creating an empty
/// `<element_name/>` when the fragment is absent or unparseable.
pugi::xml_node load_or_create_root(pugi::xml_document& doc, std::string_view fragment, const char* element_name);

/// Upserts `name="value"` on `node`.
void set_attr(pugi::xml_node node, const char* name, std::string_view value);

/// Upserts `name` as a decimal unsigned integer.
void set_attr_u32(pugi::xml_node node, const char* name, std::uint32_t value);

/// Upserts `name` as the shortest round-trip decimal of `value`.
void set_attr_double(pugi::xml_node node, const char* name, double value);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_PRINT_SETTINGS_INTERNAL_H_
