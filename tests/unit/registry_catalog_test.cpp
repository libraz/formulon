// Copyright 2026 libraz. Licensed under the MIT License.
//
// Drift detection: every function name reachable through `default_registry()`,
// the tree walker's lazy-dispatch table, or `parser_special_form_names()`
// (LET today, LAMBDA later) must appear in the canonical catalog at
// `tools/catalog/functions.txt`. Catalog entries that are NOT reachable are
// fine (they represent the implementation backlog); the only failure mode
// is an ORPHAN — a registered name missing from the catalog.
//
// The catalog path is injected at compile time via
// `-DFORMULON_CATALOG_PATH="..."` from `tests/CMakeLists.txt`, so the test
// binary does not need to reason about `__FILE__` or working directories.

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdio>
#include <fstream>
#include <sstream>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"

#ifndef FORMULON_CATALOG_PATH
#error "FORMULON_CATALOG_PATH must be defined by the build system."
#endif

namespace formulon {
namespace eval {
namespace {

struct CatalogParse {
  std::unordered_set<std::string> names;
  std::size_t raw_line_count = 0;  // Non-comment, non-blank lines (pre-dedupe).
};

// Parses `tools/catalog/functions.txt`: strips comments (`#` to EOL) and
// blank lines, collects the surviving tokens as canonical UPPERCASE names.
CatalogParse load_catalog(const char* path) {
  CatalogParse out;
  std::ifstream in(path);
  if (!in) {
    ADD_FAILURE() << "Failed to open catalog at: " << path;
    return out;
  }

  std::string line;
  while (std::getline(in, line)) {
    // Strip in-line `#` comments first, then surrounding whitespace.
    const auto hash = line.find('#');
    if (hash != std::string::npos) {
      line.resize(hash);
    }
    auto not_space = [](unsigned char c) { return !std::isspace(c); };
    const auto begin = std::find_if(line.begin(), line.end(), not_space);
    const auto end = std::find_if(line.rbegin(), line.rend(), not_space).base();
    if (begin >= end) {
      continue;  // Blank or comment-only line.
    }
    std::string name(begin, end);
    ++out.raw_line_count;
    out.names.insert(std::move(name));
  }
  return out;
}

// Collects names from `default_registry()` via the C-pointer iteration hook.
std::vector<std::string> collect_registered_names() {
  std::vector<std::string> names;
  default_registry().for_each_name(
      [](std::string_view name, void* ctx) {
        auto* sink = static_cast<std::vector<std::string>*>(ctx);
        sink->emplace_back(name);
      },
      &names);
  return names;
}

// Collects names from the tree walker's lazy-dispatch table (IF, SUMIF,
// XLOOKUP, ...). These are routed before `default_registry()` is consulted.
std::vector<std::string> collect_lazy_names() {
  std::vector<std::string> names;
  const char* const* p = lazy_form_names();
  for (; p != nullptr && *p != nullptr; ++p) {
    names.emplace_back(*p);
  }
  return names;
}

// Collects names for parser-integrated special forms (LET today; LAMBDA in
// the future). These never reach the registry or `kLazyDispatch` — the
// parser lowers them to dedicated AST shapes during parsing.
std::vector<std::string> collect_special_form_names() {
  std::vector<std::string> names;
  const char* const* p = parser_special_form_names();
  for (; p != nullptr && *p != nullptr; ++p) {
    names.emplace_back(*p);
  }
  return names;
}

TEST(RegistryCatalog, AllRegisteredInCatalog) {
  const CatalogParse catalog = load_catalog(FORMULON_CATALOG_PATH);
  ASSERT_FALSE(catalog.names.empty()) << "Catalog parsed as empty at " << FORMULON_CATALOG_PATH;

  std::vector<std::string> orphans;
  for (const auto& name : collect_registered_names()) {
    if (catalog.names.find(name) == catalog.names.end()) {
      orphans.push_back(name);
    }
  }
  for (const auto& name : collect_lazy_names()) {
    if (catalog.names.find(name) == catalog.names.end()) {
      orphans.push_back(name);
    }
  }
  for (const auto& name : collect_special_form_names()) {
    if (catalog.names.find(name) == catalog.names.end()) {
      orphans.push_back(name);
    }
  }

  if (!orphans.empty()) {
    std::sort(orphans.begin(), orphans.end());
    orphans.erase(std::unique(orphans.begin(), orphans.end()), orphans.end());
    std::ostringstream msg;
    msg << orphans.size() << " function name(s) are registered at runtime but "
        << "missing from tools/catalog/functions.txt:\n";
    for (const auto& name : orphans) {
      msg << "  - " << name << "\n";
    }
    msg << "Add them to the catalog (under the appropriate `# 11.3.x` "
        << "section) to silence this test.";
    FAIL() << msg.str();
  }
}

TEST(RegistryCatalog, CatalogHasNoDuplicates) {
  const CatalogParse catalog = load_catalog(FORMULON_CATALOG_PATH);
  EXPECT_EQ(catalog.names.size(), catalog.raw_line_count)
      << "tools/catalog/functions.txt contains duplicate entries: " << catalog.raw_line_count << " non-blank lines vs "
      << catalog.names.size() << " unique names.";
}

TEST(RegistryCatalog, CoverageReport) {
  // Informational only. This test is never expected to fail; it prints a
  // one-liner coverage summary that ctest --output-on-failure will surface
  // when something else in the suite breaks, and that developers can eyeball
  // during local runs.
  const CatalogParse catalog = load_catalog(FORMULON_CATALOG_PATH);
  ASSERT_FALSE(catalog.names.empty());

  std::unordered_set<std::string> implemented;
  for (const auto& n : collect_registered_names()) {
    if (catalog.names.find(n) != catalog.names.end()) {
      implemented.insert(n);
    }
  }
  for (const auto& n : collect_lazy_names()) {
    if (catalog.names.find(n) != catalog.names.end()) {
      implemented.insert(n);
    }
  }
  for (const auto& n : collect_special_form_names()) {
    if (catalog.names.find(n) != catalog.names.end()) {
      implemented.insert(n);
    }
  }

  const std::size_t impl_count = implemented.size();
  const std::size_t target = catalog.names.size();
  const double pct = target == 0 ? 0.0 : (100.0 * static_cast<double>(impl_count) / static_cast<double>(target));
  std::printf("[coverage] Formulon function coverage: %zu / %zu implemented (%.1f%%)\n", impl_count, target, pct);
  std::fflush(stdout);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
