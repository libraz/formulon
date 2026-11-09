#ifndef FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_
#define FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_

// Patch-object input helpers and the fill / line paint records shared by the
// chart, shape, slicer and drawing translation units.

#include <cstddef>
#include <cstdint>
#include <string>

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

/// Reads `owner.fill` / `owner.line` into a paint patch: absent keeps every member, `null` clears every member
/// (the style's paint), and an object patches the members it carries.
void ReadFillPatch(CheckedSpecReader& reader, const Napi::Object& owner, fm_chart_fill_patch_t& out);
void ReadLinePatch(CheckedSpecReader& reader, const Napi::Object& owner, fm_chart_line_patch_t& out);

/// Reads `owner.fill` / `owner.line` as the paint of a new object: absent or `null` is the style's paint (all
/// zero), and a `null` member is an absent one.
fm_chart_fill_t ReadFillSpec(CheckedSpecReader& reader, const Napi::Object& owner);
fm_chart_line_t ReadLineSpec(CheckedSpecReader& reader, const Napi::Object& owner);

/// The JS `ChartFill` / `ChartLine` record of a paint, or `null` when it states nothing (the style's paint).
Napi::Value FillOrNull(Napi::Env env, const fm_chart_fill_t& fill);
Napi::Value LineOrNull(Napi::Env env, const fm_chart_line_t& line);

/// Reads the keys every drawing object takes (`name`, `description`, `hidden`, `anchor`) out of a patch object;
/// the stores own the strings the patch points at.
fm_drawing_patch_t PullDrawingPatch(CheckedSpecReader& reader, const Napi::Object& spec, std::string& name_store,
                                    std::string& description_store);

}  // namespace detail

using detail::Field;
using detail::FillOrNull;
using detail::LineOrNull;
using detail::PatchObjectArg;
using detail::Probe;
using detail::PullDrawingPatch;
using detail::ReadFillPatch;
using detail::ReadFillSpec;
using detail::ReadLinePatch;
using detail::ReadLineSpec;

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_CHART_INTERNAL_H_
