// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `xl/_rels/workbook.xml.rels` reader. Walks the workbook-level
// relationship file once and returns an aggregated lookup keyed by
// relationship id (worksheet, sharedStrings, styles, pivot cache,
// external link). Unrecognised relationship types are captured
// verbatim so the writer can re-emit them and keep the matching
// `<Override>`-listed part reachable in the package graph.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_WORKBOOK_RELS_READER_H_
#define FORMULON_IO_OOXML_WORKBOOK_RELS_READER_H_

#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "io/unknown_relationship.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Aggregated workbook-relationship lookup: per-sheet `rId -> path` map
/// plus optional resolved paths for the `sharedStrings` and `styles`
/// parts. Empty `sst_path` / `styles_path` mean "no such relationship",
/// which is legal — the package can omit either part.
///
/// `pivot_cache_definition_paths_by_rid` carries the resolved part path
/// for every `<Relationship Type=".../pivotCacheDefinition">` entry,
/// keyed by relationship id. The workbook's `<pivotCaches>` element
/// joins each `cacheId` to its definition path through this map.
struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;
  std::string sst_path;
  std::string styles_path;
  std::unordered_map<std::string, std::string> pivot_cache_definition_paths_by_rid;
  std::unordered_map<std::string, std::string> external_link_paths_by_rid;
  // Relationship entries with Type URIs we don't recognise (theme,
  // calcChain, vbaProject, customXml, ...). Captured verbatim so the
  // writer can re-emit them, keeping passthrough-listed parts reachable
  // in the package graph.
  std::vector<UnknownRelationship> unknown_rels;
};

/// Loads `<workbook_dir>/_rels/<workbook_filename>.rels` (Excel always
/// emits this; absence is treated as a broken package) and returns the
/// aggregated relationship lookup. Each in-package target is resolved
/// against the workbook directory, with Zip-Slip path-traversal
/// hardening applied.
Expected<WorkbookRels, Error> load_workbook_rels(const ZipReader& zip, std::string_view workbook_path);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKBOOK_RELS_READER_H_
