//
// Implementation of the function dispatch table. See `function_registry.h`
// for the public contract.

#include "eval/function_registry.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "eval/builtins.h"
#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {
namespace {

// Hashes the ASCII-case-folded bytes directly from the caller's view. The
// registry stores the hash separately from its owned display key so lookup
// never needs a temporary normalized string.
std::size_t folded_hash(std::string_view value) noexcept {
  std::uint64_t hash = 14695981039346656037ULL;
  for (const char byte : value) {
    hash ^= static_cast<unsigned char>(strings::ascii_to_lower(byte));
    hash *= 1099511628211ULL;
  }
  return static_cast<std::size_t>(hash);
}

}  // namespace

struct FunctionRegistry::Impl {
  struct Entry {
    // The canonical key is owned by the registry because for_each_name() and
    // the default catalog expose it for the registry lifetime. FunctionDef's
    // canonical_name stays the caller's view.
    std::string key;
    FunctionDef def;
  };

  // A multimap is intentional: the folded hash is the lookup key, and any
  // hash collision is resolved by the exact ASCII case-insensitive compare
  // against Entry::key. Nodes remain stable across rehash and container move.
  std::unordered_multimap<std::size_t, Entry> table;
  // Lets lookup cheaply reject an unknown oversized name. Registered host
  // functions remain unrestricted: a matching long name still reaches the
  // normal lookup path.
  std::size_t longest_name = 0;
};

FunctionRegistry::FunctionRegistry() : impl_(std::make_unique<Impl>()) {}

FunctionRegistry::~FunctionRegistry() = default;

FunctionRegistry::FunctionRegistry(FunctionRegistry&&) noexcept = default;
FunctionRegistry& FunctionRegistry::operator=(FunctionRegistry&&) noexcept = default;

bool FunctionRegistry::register_function(const FunctionDef& def) {
  const std::string_view name = def.canonical_name;
  const std::size_t hash = folded_hash(name);
  const auto candidates = impl_->table.equal_range(hash);
  for (auto it = candidates.first; it != candidates.second; ++it) {
    if (strings::case_insensitive_eq(it->second.key, name)) {
      return false;
    }
  }

  // The registry owns the upper-case key; the FunctionDef is copied unchanged.
  FunctionRegistry::Impl::Entry entry{strings::to_ascii_upper(name), def};
  impl_->table.emplace(hash, std::move(entry));
  if (name.size() > impl_->longest_name) {
    impl_->longest_name = name.size();
  }
  return true;
}

const FunctionDef* FunctionRegistry::lookup(std::string_view name) const noexcept {
  if (name.size() > impl_->longest_name) {
    return nullptr;
  }
  const auto candidates = impl_->table.equal_range(folded_hash(name));
  for (auto it = candidates.first; it != candidates.second; ++it) {
    if (strings::case_insensitive_eq(it->second.key, name)) {
      return &it->second.def;
    }
  }
  return nullptr;
}

std::size_t FunctionRegistry::size() const noexcept {
  return impl_->table.size();
}

void FunctionRegistry::for_each_name(void (*cb)(std::string_view, void*), void* ctx) const {
  if (cb == nullptr) {
    return;
  }
  for (const auto& kv : impl_->table) {
    cb(std::string_view(kv.second.key), ctx);
  }
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

const FunctionDef* lookup_on_current_host(const FunctionRegistry& registry, std::string_view name) noexcept {
  const FunctionDef* def = registry.lookup(name);
  if (def != nullptr && def->absent_on_mac && current_eval_profile().host == ExcelHost::kMac365) {
    return nullptr;
  }
  return def;
}

}  // namespace eval
}  // namespace formulon
