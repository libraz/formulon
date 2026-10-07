//
// External-link parts of the OOXML writer: the body generated from an
// `ExternalLinkRecord` when no current loaded body exists, the per-link
// rels file, and the `[N]` numbering saved formulas use.
//
// A stored formula names a book by the 1-based position of its link in
// `<externalReferences>`, which differs from `ExternalLinkRecord::index`
// once the indices have gaps, so the numbering is derived from the links
// this package actually writes. Internal to `src/io/ooxml/`.

#ifndef FORMULON_IO_OOXML_EXTERNAL_LINK_WRITER_H_
#define FORMULON_IO_OOXML_EXTERNAL_LINK_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast_format.h"

namespace formulon {

class Workbook;
struct ExternalLinkRecord;

namespace io {

/// Maps a cross-workbook reference to the `[N]` a saved formula carries:
/// the 1-based position of its link among the written links.
class ExternalLinkOrdinals {
 public:
  ExternalLinkOrdinals() = default;
  /// `written_indices` holds `ExternalLinkRecord::index` of each written
  /// link, in `<externalReferences>` order.
  ExternalLinkOrdinals(const Workbook& wb, std::vector<std::uint32_t> written_indices);

  /// An indexer over this object; valid while it is neither moved nor destroyed.
  parser::ExternalBookIndexer indexer() const noexcept { return parser::ExternalBookIndexer{&Ordinal, this}; }

 private:
  static std::uint32_t Ordinal(const void* ctx, std::string_view path, std::string_view book);

  parser::ExternalBookIndexer links_{nullptr, nullptr};
  std::vector<std::uint32_t> written_indices_;
};

/// The `externalLink<N>.xml` body for `rec`, generated from its cached
/// book: every sheet name, every cached name (`sheetId` on a sheet-scope
/// one), a `<sheetData>` for each sheet that has data, and the absolute
/// path's relationship when `rec.absolute_target` is set.
std::string BuildExternalLinkXml(const ExternalLinkRecord& rec);

/// The per-link rels file: the target relationship, and the absolute
/// path's when `rec.absolute_target` is set. Empty when the record has
/// neither.
std::string BuildExternalLinkRels(const ExternalLinkRecord& rec);

/// True when a formula must be written as `storage` rather than as its
/// held text: the storage form adds a prefix, names a book by index, or
/// quotes a sheet the held text leaves bare.
bool NeedsStorageSpelling(const parser::AstNode& root, std::string_view storage);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_EXTERNAL_LINK_WRITER_H_
