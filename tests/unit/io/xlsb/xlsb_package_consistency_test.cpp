// Package structure of every excel_open_gate base converted to .xlsb: each relationship target exists, each part
// other than [Content_Types].xml, the .rels parts and docProps/* is the target of exactly one relationship, and
// each part has a content type.

#include <algorithm>
#include <cctype>
#include <map>
#include <set>
#include <string>
#include <tuple>
#include <vector>

#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb::test {
namespace {

constexpr const char* kDanglingTarget = "dangling-target";
constexpr const char* kSingleReference = "single-reference";
constexpr const char* kContentType = "content-type";

/// A rule a converted base breaks today. Each entry is asserted to still fail, so a fix shows up as a failure
/// that removes the entry; the list is empty once the writer consumes every part it keeps.
struct KnownFailure {
  const char* base;
  const char* rule;
  const char* part;
  const char* reason;
};

const std::vector<KnownFailure> kKnownFailures = {};

std::string Lower(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
  return s;
}

bool StartsWith(const std::string& s, const std::string& prefix) {
  return s.compare(0, prefix.size(), prefix) == 0;
}

bool IsRelsPart(const std::string& name) {
  return name.size() > 5U && name.compare(name.size() - 5U, 5U, ".rels") == 0;
}

/// The part a relationship in `rels_part` points at, with `.` and `..` resolved.
std::string ResolveTarget(const std::string& rels_part, const std::string& target) {
  if (!target.empty() && target[0] == '/') {
    return target.substr(1);
  }
  // `dir/_rels/name.rels` describes `dir/name`; targets are relative to `dir`.
  const std::size_t rels_dir = rels_part.rfind("_rels/");
  std::string base = rels_dir == std::string::npos ? std::string() : rels_part.substr(0, rels_dir);
  std::vector<std::string> segments;
  const std::string joined = base + target;
  std::size_t start = 0;
  while (start <= joined.size()) {
    const std::size_t slash = std::min(joined.find('/', start), joined.size());
    const std::string segment = joined.substr(start, slash - start);
    if (segment == "..") {
      if (!segments.empty()) {
        segments.pop_back();
      }
    } else if (!segment.empty() && segment != ".") {
      segments.push_back(segment);
    }
    start = slash + 1U;
  }
  std::string out;
  for (const std::string& segment : segments) {
    out += (out.empty() ? "" : "/") + segment;
  }
  return out;
}

/// Every (rule, part) the package breaks.
std::set<std::pair<std::string, std::string>> Violations(const Bytes& package) {
  std::set<std::pair<std::string, std::string>> out;
  const std::vector<std::string> parts = ListParts(package);
  const std::set<std::string> present(parts.begin(), parts.end());

  std::set<std::string> default_extensions;
  std::set<std::string> overrides;
  if (const std::optional<Bytes> types = ReadPart(package, "[Content_Types].xml")) {
    pugi::xml_document doc;
    if (doc.load_buffer(types->data(), types->size())) {
      for (const pugi::xml_node node : doc.document_element().children()) {
        const std::string name = node.name();
        if (name == "Default") {
          default_extensions.insert(Lower(node.attribute("Extension").value()));
        } else if (name == "Override") {
          overrides.insert(Lower(ResolveTarget("", node.attribute("PartName").value())));
        }
      }
    }
  } else {
    out.emplace(kContentType, "[Content_Types].xml");
  }

  std::map<std::string, int> references;
  for (const std::string& name : parts) {
    if (!IsRelsPart(name)) {
      continue;
    }
    const std::optional<Bytes> rels = ReadPart(package, name);
    pugi::xml_document doc;
    if (!rels || !doc.load_buffer(rels->data(), rels->size())) {
      out.emplace(kDanglingTarget, name);
      continue;
    }
    for (const pugi::xml_node rel : doc.document_element().children("Relationship")) {
      if (std::string(rel.attribute("TargetMode").value()) == "External") {
        continue;
      }
      const std::string target = ResolveTarget(name, rel.attribute("Target").value());
      ++references[target];
      if (present.count(target) == 0U) {
        out.emplace(kDanglingTarget, target);
      }
    }
  }

  for (const std::string& name : parts) {
    if (name == "[Content_Types].xml") {
      continue;
    }
    const std::size_t dot = name.rfind('.');
    const std::string extension = dot == std::string::npos ? std::string() : Lower(name.substr(dot + 1U));
    if (default_extensions.count(extension) == 0U && overrides.count(Lower(name)) == 0U) {
      out.emplace(kContentType, name);
    }
    if (IsRelsPart(name) || StartsWith(name, "docProps/")) {
      continue;
    }
    if (references[name] != 1) {
      out.emplace(kSingleReference, name);
    }
  }
  return out;
}

TEST(XlsbPackageConsistency, ResolveTargetHandlesRelativeForms) {
  EXPECT_EQ(ResolveTarget("_rels/.rels", "xl/workbook.bin"), "xl/workbook.bin");
  EXPECT_EQ(ResolveTarget("xl/_rels/workbook.bin.rels", "worksheets/sheet1.bin"), "xl/worksheets/sheet1.bin");
  EXPECT_EQ(ResolveTarget("xl/worksheets/_rels/sheet1.bin.rels", "../drawings/vmlDrawing1.vml"),
            "xl/drawings/vmlDrawing1.vml");
  EXPECT_EQ(ResolveTarget("xl/worksheets/_rels/sheet1.bin.rels", "/xl/comments1.bin"), "xl/comments1.bin");
}

TEST(XlsbPackageConsistency, ConvertedGateBases) {
  const std::string root = FORMULON_FIXTURES_DIR;
  const std::vector<std::string> bases = GateBases(root + "/../../tools/oracle/excel_open_gate.txt");
  ASSERT_FALSE(bases.empty());
  std::set<std::tuple<std::string, std::string, std::string>> expected;
  for (const KnownFailure& known : kKnownFailures) {
    expected.emplace(known.base, known.rule, known.part);
  }
  std::set<std::tuple<std::string, std::string, std::string>> seen;
  for (const std::string& base : bases) {
    const Converted converted = ConvertToXlsb(root, base);
    ASSERT_TRUE(converted.error.empty()) << converted.error;
    for (const auto& violation : Violations(converted.bytes)) {
      const auto key = std::make_tuple(base, violation.first, violation.second);
      seen.insert(key);
      EXPECT_NE(expected.count(key), 0U) << base << ": " << violation.first << " " << violation.second;
    }
  }
  for (const KnownFailure& known : kKnownFailures) {
    EXPECT_NE(seen.count(std::make_tuple(std::string(known.base), std::string(known.rule), std::string(known.part))),
              0U)
        << "fixed, remove the entry: " << known.base << " " << known.rule << " " << known.part;
  }
}

}  // namespace
}  // namespace formulon::io::xlsb::test
