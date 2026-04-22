// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the function dispatch table. See `function_registry.h`
// for the public contract.

#include "eval/function_registry.h"

#include <cctype>
#include <cstddef>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "eval/builtins.h"

namespace formulon {
namespace eval {
namespace {

// Uppercases ASCII letters; non-ASCII bytes pass through verbatim. We loop
// through `unsigned char` to avoid the undefined behaviour of `std::toupper`
// on `char` values >= 0x80 with the default locale.
std::string to_upper_ascii(std::string_view s) {
  std::string out;
  out.resize(s.size());
  for (std::size_t i = 0; i < s.size(); ++i) {
    const auto byte = static_cast<unsigned char>(s[i]);
    if (byte < 0x80) {
      out[i] = static_cast<char>(std::toupper(byte));
    } else {
      out[i] = s[i];
    }
  }
  return out;
}

}  // namespace

struct FunctionRegistry::Impl {
  std::unordered_map<std::string, FunctionDef> table;
};

FunctionRegistry::FunctionRegistry() : impl_(std::make_unique<Impl>()) {}

FunctionRegistry::~FunctionRegistry() = default;

FunctionRegistry::FunctionRegistry(FunctionRegistry&&) noexcept = default;
FunctionRegistry& FunctionRegistry::operator=(FunctionRegistry&&) noexcept = default;

bool FunctionRegistry::register_function(const FunctionDef& def) {
  std::string key = to_upper_ascii(def.canonical_name);
  // unordered_map::emplace returns {iterator, inserted}; we never overwrite
  // an existing entry, so a duplicate key is a no-op.
  auto result = impl_->table.emplace(std::move(key), def);
  return result.second;
}

const FunctionDef* FunctionRegistry::lookup(std::string_view name) const noexcept {
  // Allocation may fail under -fno-exceptions; in practice the key is short
  // and the lookup is on the cold path of dispatch (we already failed to find
  // a builtin via the hot bytecode path).
  std::string key = to_upper_ascii(name);
  auto it = impl_->table.find(key);
  if (it == impl_->table.end()) {
    return nullptr;
  }
  return &it->second;
}

std::size_t FunctionRegistry::size() const noexcept {
  return impl_->table.size();
}

const FunctionRegistry& default_registry() {
  // Meyers singleton: thread-safe lazy initialization since C++11.
  static const FunctionRegistry instance = [] {
    FunctionRegistry r;
    register_builtins(r);
    return r;
  }();
  return instance;
}

}  // namespace eval
}  // namespace formulon
