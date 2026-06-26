// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "c_api/parts/common.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <type_traits>
#include <utility>

#include "c_api/formulon_c.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace c_api {
namespace parts {

// ABI identity: the `FM_VAL_*` discriminators are numerically equal to the
// `formulon::ValueKind` ordinals. `value_to_fm` below maps each ValueKind
// to its `FM_VAL_*` by name; these static_asserts pin the parity so a
// reorder of either enum is a build break rather than a silent wire-format
// drift across the C ABI.
static_assert(FM_VAL_BLANK == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Blank),
              "FM_VAL_BLANK must match ValueKind::Blank");
static_assert(FM_VAL_NUMBER == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Number),
              "FM_VAL_NUMBER must match ValueKind::Number");
static_assert(FM_VAL_BOOL == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Bool),
              "FM_VAL_BOOL must match ValueKind::Bool");
static_assert(FM_VAL_TEXT == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Text),
              "FM_VAL_TEXT must match ValueKind::Text");
static_assert(FM_VAL_ERROR == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Error),
              "FM_VAL_ERROR must match ValueKind::Error");
static_assert(FM_VAL_ARRAY == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Array),
              "FM_VAL_ARRAY must match ValueKind::Array");
static_assert(FM_VAL_REF == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Ref),
              "FM_VAL_REF must match ValueKind::Ref");
static_assert(FM_VAL_LAMBDA == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Lambda),
              "FM_VAL_LAMBDA must match ValueKind::Lambda");

namespace {

// Thread-local diagnostics for the most recent C API call on this thread.
//
// Both buffers are always non-empty in the sense that they own valid string
// storage, but their `c_str()` may point to an empty string. Returning
// `c_str()` therefore satisfies the public "never NULL" contract.
thread_local std::string g_last_error_message;
thread_local std::string g_last_error_context;

}  // namespace

void clear_last_error() {
  g_last_error_message.clear();
  g_last_error_context.clear();
}

fm_status_t set_last_error(const formulon::Error& err) {
  g_last_error_message = err.message;
  g_last_error_context = err.context;
  return static_cast<fm_status_t>(err.code);
}

fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context) {
  formulon::Error err;
  err.code = code;
  err.message = message != nullptr ? message : "";
  err.context = std::move(context);
  return set_last_error(err);
}

const char* last_error_message() {
  return g_last_error_message.c_str();
}

const char* last_error_context() {
  return g_last_error_context.c_str();
}

bool check_range_count(std::uint32_t n, const char* api) {
  if (n > kMaxRangesPerCApiCall) {
    set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "range_count exceeds per-call cap",
        std::string(api) + ": range_count=" + std::to_string(n) + " cap=" + std::to_string(kMaxRangesPerCApiCall));
    return false;
  }
  return true;
}

std::string_view intern_text(TextStore& store, std::string_view text) {
  store.emplace_back(text.data(), text.size());
  return std::string_view(store.back());
}

void value_to_fm(const formulon::Value& v, TextStore& store, fm_value_t* out) {
  switch (v.kind()) {
    case formulon::ValueKind::Blank:
      out->kind = FM_VAL_BLANK;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Number:
      out->kind = FM_VAL_NUMBER;
      out->u.number = v.as_number();
      return;
    case formulon::ValueKind::Bool:
      out->kind = FM_VAL_BOOL;
      out->u.boolean = v.as_boolean() ? 1 : 0;
      return;
    case formulon::ValueKind::Text: {
      const std::string_view text = v.as_text();
      store.emplace_back(text.data(), text.size());
      out->kind = FM_VAL_TEXT;
      out->u.text = store.back().c_str();
      return;
    }
    case formulon::ValueKind::Error:
      out->kind = FM_VAL_ERROR;
      // The C-ABI wire format for errors is the raw `ErrorCode` ordinal
      // (Null=0, Div0=1, Value=2, ...), NOT the OOXML wire code. This is
      // the single source of truth for the boundary; see the
      // `fm_value_t::u.error_code` contract in `formulon_c.h`.
      out->u.error_code = static_cast<int32_t>(v.as_error());
      return;
    case formulon::ValueKind::Array:
      // Array passthrough is reserved; the kind is reported but no
      // payload is exposed across the boundary in this bundle.
      out->kind = FM_VAL_ARRAY;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Ref:
      out->kind = FM_VAL_REF;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Lambda:
      out->kind = FM_VAL_LAMBDA;
      out->u.number = 0.0;
      return;
  }
  // Defensive default: surface as Blank rather than leaving uninitialised
  // bytes on the boundary. The switch above is exhaustive over every
  // `ValueKind` enumerator, so this branch is unreachable in practice.
  out->kind = FM_VAL_BLANK;
  out->u.number = 0.0;
}

fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

fm_status_t check_sheet_u32(const fm_workbook_t* wb, std::uint32_t sheet, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (static_cast<std::size_t>(sheet) >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet=" + std::to_string(sheet) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
