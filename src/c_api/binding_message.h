/// @file binding_message.h
/// One builder for the argument-error messages the JS bindings raise from their positional guards and record readers.
/// Header-only and free of emscripten / napi so the WASM and native addons word every reject alike.
// @size-budget: 16 KB
#ifndef FORMULON_C_API_BINDING_MESSAGE_H_
#define FORMULON_C_API_BINDING_MESSAGE_H_

#include <cstddef>
#include <string>
#include <string_view>

namespace formulon::binding {

namespace detail {
#include "c_api/generated/binding_param_names.inc"
#include "c_api/generated/binding_record_params.inc"
}  // namespace detail

/// `<call> <path> <constraint>`; an empty `path` gives `<call> <constraint>`.
inline std::string argument_message(std::string_view call, std::string_view path, std::string_view constraint) {
  std::string message(call);
  if (!path.empty()) {
    message.push_back(' ');
    message.append(path);
  }
  message.push_back(' ');
  message.append(constraint);
  return message;
}

/// `<parent>.<field>`, or `<field>` when `parent` is empty.
inline std::string field_path(std::string_view parent, std::string_view field) {
  std::string path(parent);
  if (!path.empty()) {
    path.push_back('.');
  }
  path.append(field);
  return path;
}

/// `<array_path>[<index>]`.
inline std::string element_path(std::string_view array_path, std::size_t index) {
  std::string path(array_path);
  path.push_back('[');
  path.append(std::to_string(index));
  path.push_back(']');
  return path;
}

/// The record parameter of `call` (`"patch"` for `Sheet.updateView`), `""` when it takes none.
inline std::string_view record_param(std::string_view call) {
  for (const detail::RecordParam& entry : detail::kRecordParams) {
    if (call == entry.call) {
      return entry.param;
    }
  }
  return {};
}

/// The d.ts name of positional parameter `index` of `call` (`"row"` for `Sheet.setNumber`, 0), `""` when `call`
/// declares no parameter at `index`.
inline std::string_view param_name(std::string_view call, std::size_t index) {
  for (const detail::ParamNames& entry : detail::kParamNames) {
    if (call != entry.call) {
      continue;
    }
    std::string_view names(entry.names);
    for (; index > 0; --index) {
      const std::size_t space = names.find(' ');
      if (space == std::string_view::npos) {
        return {};
      }
      names.remove_prefix(space + 1);
    }
    return names.substr(0, names.find(' '));
  }
  return {};
}

/// The M1 path of `field` read inside `call`: `field` itself when its first segment already names a parameter of
/// `call`, else `field` under the call's record parameter (`patch.zoomPct` for `Sheet.updateView`, `zoomPct`).
inline std::string record_path(std::string_view call, std::string_view field) {
  const std::string_view head = field.substr(0, field.find_first_of(".["));
  for (const detail::ParamNames& entry : detail::kDeclaredParamNames) {
    if (call != entry.call) {
      continue;
    }
    std::string_view names(entry.names);
    while (!names.empty()) {
      const std::size_t space = names.find(' ');
      if (names.substr(0, space) == head) {
        return std::string(field);
      }
      names = space == std::string_view::npos ? std::string_view() : names.substr(space + 1);
    }
    break;
  }
  return field_path(record_param(call), field);
}

}  // namespace formulon::binding

#endif  // FORMULON_C_API_BINDING_MESSAGE_H_
