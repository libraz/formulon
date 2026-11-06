// Helpers shared by the style-pool, cell/table-style and theme bindings.

#ifndef FORMULON_NODE_ADDON_PARTS_STYLES_INTERNAL_H_
#define FORMULON_NODE_ADDON_PARTS_STYLES_INTERNAL_H_

#include <cstddef>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace detail {

/// Builds the `getCellXf` / `getCellStyleXf` record.
Napi::Object CellXfResultToJs(Napi::Env env, const fm_cell_xf_t& xf);

// Shared body of the style-pool count getters that take no argument.
using StyleCountFn = fm_status_t (*)(const fm_workbook_t*, std::size_t*);

Napi::Value StyleCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, StyleCountFn fn);

}  // namespace detail
}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_STYLES_INTERNAL_H_
