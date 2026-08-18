//
// Every XML part of a saved package must parse.
//
// This is the cheapest possible assertion about the writer's output, and
// until now nothing made it. The read<->write symmetry tests compare a
// model against a model, so a part the reader can no longer parse is only
// noticed if the reader is asked to parse it again -- and no test loaded an
// Excel-authored pivot workbook and saved it. That gap let the writer emit
// a namespace-qualified attribute (`mc:Ignorable`, `xr:uid`) with nothing
// binding its prefix: not merely lossy, but XML no parser accepts.
//
// The check deliberately stops short of schema validity, which needs the
// ECMA-376 XSDs, not a dependency here; and Excel's acceptance is stricter
// than either -- see the pivot subtotal markers, which are schema-optional
// and still required. Well-formedness is the floor, and a floor that is
// mechanically enforceable on every host is worth more than a ceiling that
// is not.
//
// Parsing alone is not that floor. pugixml is not namespace-aware: it reads
// `mc:Ignorable` as an attribute whose name happens to contain a colon and
// never asks whether `mc` is bound. Handing it the malformed output this
// file exists to catch produces a clean parse. So the binding check below
// is written out explicitly rather than delegated to the parser -- without
// it these tests pass against the defect, which was confirmed by injecting
// it.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

/// True when `prefix` is declared on `node` or any of its ancestors.
/// `xml` is bound by the XML specification itself and never declared.
bool PrefixIsBound(const pugi::xml_node& node, const std::string& prefix) {
  if (prefix == "xml" || prefix == "xmlns") {
    return true;
  }
  const std::string declaration = "xmlns:" + prefix;
  for (pugi::xml_node scope = node; scope; scope = scope.parent()) {
    if (scope.attribute(declaration.c_str())) {
      return true;
    }
  }
  return false;
}

std::string PrefixOf(std::string_view qualified_name) {
  const std::size_t colon = qualified_name.find(':');
  return colon == std::string_view::npos ? std::string() : std::string(qualified_name.substr(0, colon));
}

/// Reports every qualified element or attribute name at or below `node`
/// whose prefix nothing binds, as human-readable locations.
void CollectUnboundPrefixes(const pugi::xml_node& node, const std::string& part, std::vector<std::string>& problems) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string element_prefix = PrefixOf(child.name());
    if (!element_prefix.empty() && !PrefixIsBound(child, element_prefix)) {
      problems.push_back(part + ": <" + child.name() + "> prefix '" + element_prefix + "' is not bound");
    }
    for (pugi::xml_attribute attr = child.first_attribute(); attr; attr = attr.next_attribute()) {
      const std::string attr_prefix = PrefixOf(attr.name());
      if (attr_prefix.empty() || PrefixIsBound(child, attr_prefix)) {
        continue;
      }
      problems.push_back(part + ": <" + child.name() + " " + attr.name() + "> prefix '" + attr_prefix +
                         "' is not bound");
    }
    CollectUnboundPrefixes(child, part, problems);
  }
}

/// Loads `fixture`, saves it, and parses every `.xml` part of the result.
/// Reports the offending part and pugixml's own description on failure,
/// because "a part did not parse" is not actionable on its own.
void ExpectSavedPartsParse(const char* fixture) {
  const std::vector<std::uint8_t> original = test::read_file_bytes(FixturePath(fixture));
  ASSERT_FALSE(original.empty()) << "fixture not readable: " << fixture;

  auto read = io::read_ooxml(test::span_of(original));
  ASSERT_TRUE(static_cast<bool>(read)) << fixture << ": read failed: " << read.error().message;
  Workbook wb = std::move(read.value().workbook);

  const auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << fixture << ": save failed: " << saved.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(saved.value()))))
      << fixture << ": saved package is not a readable archive";

  std::size_t parsed = 0;
  for (const std::string& name : zip.list_entries()) {
    if (name.size() < 4 || name.compare(name.size() - 4, 4, ".xml") != 0) {
      continue;
    }
    auto body = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(body)) << fixture << ": cannot read " << name;
    pugi::xml_document doc;
    const pugi::xml_parse_result result = doc.load_buffer(body.value().data(), body.value().size());
    EXPECT_TRUE(result) << fixture << ": " << name << " does not parse: " << result.description();
    if (!result) {
      continue;
    }
    std::vector<std::string> unbound;
    CollectUnboundPrefixes(doc, name, unbound);
    for (const std::string& problem : unbound) {
      ADD_FAILURE() << fixture << ": " << problem;
    }
    ++parsed;
  }
  EXPECT_GT(parsed, 0U) << fixture << ": saved package contained no XML parts";
}

// The pivot fixtures are the ones that matter most here: they are the only
// Excel-authored packages carrying `mc:`/`xr:`-qualified attributes on parts
// the reader consumes, which is exactly where a captured attribute can
// outlive its namespace binding.
TEST(SavedPackageWellFormed, PivotCaptionFilterFixture) {
  ExpectSavedPartsParse("pivot_caption_filter.xlsx");
}

TEST(SavedPackageWellFormed, PivotValueDateFilterFixture) {
  ExpectSavedPartsParse("pivot_value_date_filters.xlsx");
}

TEST(SavedPackageWellFormed, StylesAndFormulaFixtures) {
  ExpectSavedPartsParse("xlsb_fidelity_base.xlsx");
  ExpectSavedPartsParse("formula_corpus.xlsx");
  ExpectSavedPartsParse("storage_prefix_probe.xlsx");
}

}  // namespace
}  // namespace formulon
