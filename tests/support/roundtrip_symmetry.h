// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Read<->write symmetry test infrastructure.
//
// Self round-trip tests (feed a hand-built model to the writer, read it back
// with the reader) cannot detect the "reader understands an attribute the
// writer never emits (or vice versa)" defect class: both halves share the
// same blind spot. This header provides utilities to exercise the *whole*
// OOXML pipeline against real Excel-produced XML and assert that a declared
// set of attributes survives a `load -> save` cycle.
//
// Two granularities are supported:
//
//   * Package level: `load_save_cycle` runs `read_ooxml -> Workbook::save`,
//     `extract_part` pulls a named part (e.g. "xl/styles.xml") out of an
//     xlsx zip, and `part_attributes_survive_save` asserts that the
//     attributes present at a given XPath in the ORIGINAL part are still
//     present with equal values in the SAVED output.
//
//   * Fragment level: `parse_xml` + `attributes_preserved` compare two
//     already-parsed documents (e.g. a component reader/writer's input and
//     re-serialised output).
//
// XPath note: pugixml is namespace-naive, so an unprefixed node test
// (`//workbookPr`, `//numFmt`) matches default-namespace OOXML elements by
// their literal stored name. Prefer such unprefixed paths.
//
// Header-only and test-only: it uses `<fstream>` (fixtures live on disk) and
// gtest `AssertionResult`, neither of which belongs in library code.

#ifndef FORMULON_TESTS_SUPPORT_ROUNDTRIP_SYMMETRY_H_
#define FORMULON_TESTS_SUPPORT_ROUNDTRIP_SYMMETRY_H_

#include <cstdint>
#include <fstream>
#include <initializer_list>
#include <iterator>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "workbook.h"

namespace formulon {
namespace test {

/// Wraps a byte vector in an `io::ByteSpan`.
inline io::ByteSpan span_of(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// Reads the whole file at `path` into a byte vector. On failure raises a
/// non-fatal gtest failure and returns an empty vector.
inline std::vector<std::uint8_t> read_file_bytes(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  if (!in) {
    ADD_FAILURE() << "cannot open fixture: " << path;
    return {};
  }
  return std::vector<std::uint8_t>(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

/// Extracts part `name` (e.g. "xl/styles.xml") from the xlsx package `pkg` as
/// text. Returns an `AssertionFailure` describing the first failing step.
inline ::testing::AssertionResult extract_part(io::ByteSpan pkg, std::string_view name, std::string* out) {
  io::ZipReader zip;
  auto opened = zip.open(pkg);
  if (!opened) {
    return ::testing::AssertionFailure() << "zip open failed: " << opened.error().message;
  }
  if (!zip.has_entry(name)) {
    return ::testing::AssertionFailure() << "package has no part: " << name;
  }
  auto bytes = zip.read_entry(name);
  if (!bytes) {
    return ::testing::AssertionFailure() << "read part failed (" << name << "): " << bytes.error().message;
  }
  out->assign(bytes.value().begin(), bytes.value().end());
  return ::testing::AssertionSuccess();
}

/// Runs one full package cycle: `pkg -> read_ooxml -> Workbook::save`, writing
/// the freshly serialised bytes to `*out`.
inline ::testing::AssertionResult load_save_cycle(io::ByteSpan pkg, std::vector<std::uint8_t>* out) {
  auto loaded = io::read_ooxml(pkg);
  if (!loaded) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << loaded.error().message;
  }
  auto saved = loaded.value().workbook.save();
  if (!saved) {
    return ::testing::AssertionFailure() << "Workbook::save failed: " << saved.error().message;
  }
  *out = std::move(saved.value());
  return ::testing::AssertionSuccess();
}

/// Parses `xml` into `*doc`. Returns an `AssertionFailure` on a parse error.
inline ::testing::AssertionResult parse_xml(std::string_view xml, pugi::xml_document* doc) {
  const pugi::xml_parse_result rc = doc->load_buffer(xml.data(), xml.size());
  if (!rc) {
    return ::testing::AssertionFailure() << "xml parse failed: " << rc.description();
  }
  return ::testing::AssertionSuccess();
}

/// Asserts that for the element matched by `xpath` in `after`, every attribute
/// in `attrs` that the SAME-xpath element in `before` carries is present with
/// an equal value. Attributes absent from `before` are ignored (nothing to
/// preserve); a dropped or changed attribute is reported. A missing node on
/// either side is a failure.
inline ::testing::AssertionResult attributes_preserved(const pugi::xml_document& before,
                                                       const pugi::xml_document& after, const char* xpath,
                                                       std::initializer_list<const char*> attrs) {
  const pugi::xpath_node bn = before.select_node(xpath);
  const pugi::xpath_node an = after.select_node(xpath);
  if (!bn) {
    return ::testing::AssertionFailure() << "xpath not found in 'before': " << xpath;
  }
  if (!an) {
    return ::testing::AssertionFailure() << "xpath not found in 'after': " << xpath;
  }
  for (const char* attr : attrs) {
    const pugi::xml_attribute ba = bn.node().attribute(attr);
    if (!ba) {
      continue;  // Original does not carry it; nothing to preserve.
    }
    const pugi::xml_attribute aa = an.node().attribute(attr);
    if (!aa) {
      return ::testing::AssertionFailure()
             << xpath << ": attribute '" << attr << "' was dropped (original value '" << ba.value() << "')";
    }
    if (std::string_view(aa.value()) != std::string_view(ba.value())) {
      return ::testing::AssertionFailure()
             << xpath << ": attribute '" << attr << "' changed '" << ba.value() << "' -> '" << aa.value() << "'";
    }
  }
  return ::testing::AssertionSuccess();
}

/// End-to-end symmetry check: load-save `pkg`, then assert that at
/// `part`/`xpath` every attribute in `attrs` present in the ORIGINAL survives
/// into the saved output. This is the core "the writer preserves what the
/// reader understood" assertion.
inline ::testing::AssertionResult part_attributes_survive_save(io::ByteSpan pkg, std::string_view part,
                                                               const char* xpath,
                                                               std::initializer_list<const char*> attrs) {
  std::string before_xml;
  const ::testing::AssertionResult e1 = extract_part(pkg, part, &before_xml);
  if (!e1) {
    return e1;
  }
  std::vector<std::uint8_t> cycled;
  const ::testing::AssertionResult e2 = load_save_cycle(pkg, &cycled);
  if (!e2) {
    return e2;
  }
  std::string after_xml;
  const ::testing::AssertionResult e3 = extract_part(span_of(cycled), part, &after_xml);
  if (!e3) {
    return e3;
  }
  pugi::xml_document before;
  pugi::xml_document after;
  const ::testing::AssertionResult p1 = parse_xml(before_xml, &before);
  if (!p1) {
    return p1;
  }
  const ::testing::AssertionResult p2 = parse_xml(after_xml, &after);
  if (!p2) {
    return p2;
  }
  return attributes_preserved(before, after, xpath, attrs);
}

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_SUPPORT_ROUNDTRIP_SYMMETRY_H_
