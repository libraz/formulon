#ifndef FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_
#define FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_

// Patch-object input helpers shared by the chart read and patch translation units.

#include <cstddef>
#include <cstdint>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace detail {

/// A patch field: absent keeps, `null` clears, anything else sets `value`.
struct Field {
  fm_patch_op_t op = FM_PATCH_OP_KEEP;
  Napi::Value value;
};

Field Probe(CheckedSpecReader& reader, const Napi::Object& owner, const char* key);

/// Reads the patch object argument `info[idx]`; absent reads as `{}`. False after
/// throwing when a value other than a plain object was supplied.
bool PatchObjectArg(const Napi::CallbackInfo& info, size_t idx, const char* api, Napi::Object* out);

}  // namespace detail

using detail::Field;
using detail::PatchObjectArg;
using detail::Probe;

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_
