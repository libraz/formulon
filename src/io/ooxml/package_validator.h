//
// Package-level structural validation for OOXML reads. Bundles
// `[Content_Types].xml` parsing, root-level rels lookup, and the path
// helpers (`resolve_relative_path`, `dir_of`, `rels_path_for_part`,
// `relationship_ref_id`, `extension_of_part`) that the OOXML reader, the
// OOXML writer and the XLSB reader/writer all share.
//
// Zip-Slip hardening: `resolve_relative_path` refuses any input that
// escapes the package root via excessive `..` segments or that begins
// with an absolute-path slash. Callers must propagate the error and
// refuse to open the package.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_PACKAGE_VALIDATOR_H_
#define FORMULON_IO_OOXML_PACKAGE_VALIDATOR_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/default_content_type.h"
#include "io/workbook_kind.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// One entry from `[Content_Types].xml`'s `<Override>` list, paired with
/// its declared content type. The reader uses the content type (a) to
/// decide whether the part is interesting at all (we only consume
/// recognised content types) and (b) so the writer slice can re-emit
/// the `<Override>` for passthrough parts verbatim.
struct OverrideEntry {
  std::string part_name;     ///< package-relative, no leading slash
  std::string content_type;  ///< verbatim ContentType= attribute value
};

/// Verifies that `[Content_Types].xml` references a workbook content
/// type at least once and returns the corresponding `WorkbookKind`.
///
/// Accepts the four canonical Excel workbook content types
/// (xlsx / xlsm / xltx / xltm). When the package declares a
/// workbook-shaped override (PartName=`/xl/workbook.xml`) whose
/// ContentType is not recognised, surfaces a structured-log warning
/// and falls back to `WorkbookKind::kXlsx` rather than failing the
/// read (Excel-compatibility-first behaviour).
Expected<WorkbookKind, Error> verify_content_types(const std::vector<std::uint8_t>& ct_bytes);

/// Lists every part name advertised by `[Content_Types].xml`'s
/// `<Override>` elements together with its content type. `<Default>`
/// entries are ignored: they describe extensions rather than specific
/// parts, and the passthrough flow only carries Override-registered
/// parts.
Expected<std::vector<OverrideEntry>, Error> list_override_part_entries(const std::vector<std::uint8_t>& ct_bytes);

/// Lists every `<Default>` entry declared in `[Content_Types].xml`,
/// pairing each file extension with its content type. Excel registers
/// binary and media parts (vbaProject.bin, images, VML, printer
/// settings) by extension here rather than through per-part
/// `<Override>` elements. The reader captures these so the writer can
/// re-emit the `<Default>` registration for any Default-typed part it
/// round-trips through passthrough. Extensions are lower-cased so
/// comparison against archive entry suffixes is case-insensitive.
Expected<std::vector<DefaultContentType>, Error> list_default_content_types(const std::vector<std::uint8_t>& ct_bytes);

/// Returns the part path the package-level rels file points at for
/// `OfficeDocument`. The OOXML spec allows arbitrary placement (Excel
/// always uses `/xl/workbook.xml`), so we follow the relationship
/// rather than hard-coding the path. The path is normalised to drop
/// any leading slash so the result is directly consumable as a ZIP
/// entry name.
Expected<std::string, Error> resolve_office_document_path(const std::vector<std::uint8_t>& rels_bytes);

/// Builds a path relative to `base_dir`. OOXML rels Target attributes
/// are relative to the part that owns the rels file. We collapse `..`
/// segments so `worksheets/sheet1.xml` resolved against `xl/` yields
/// `xl/worksheets/sheet1.xml`.
///
/// Path-traversal hardening: refuses any input whose `..` segments
/// outnumber the prefix directories (would escape the package root)
/// and refuses absolute-path targets that begin with `/`. Both cases
/// surface `kIoZipSlip`.
Expected<std::string, Error> resolve_relative_path(std::string_view base_dir, std::string_view target);

/// Returns true when `part_name` is a safe, canonical OPC part name that
/// may be carried through passthrough and re-emitted verbatim. Rejects
/// anything a downstream extractor could interpret as an escape from the
/// package root: absolute paths (leading `/`), backslash separators,
/// drive-letter / scheme colons, `.`/`..` path segments, and empty
/// segments (`//`). OPC §9.1.1 requires part names to be `/`-rooted with
/// no `.`/`..` segments, so a well-formed Excel package never trips this;
/// the guard exists purely to refuse hostile archives.
bool is_safe_part_name(std::string_view part_name) noexcept;

/// Returns the directory portion of `path` (everything up to the last
/// `/`, exclusive). Empty for top-level paths like `_rels/.rels`.
std::string dir_of(std::string_view path);

/// Returns the `.rels` file path corresponding to any OOXML part path.
/// For example, `xl/worksheets/sheet1.xml` becomes
/// `xl/worksheets/_rels/sheet1.xml.rels`.
std::string rels_path_for_part(std::string_view part_path);

/// Returns the Office relationship id from nodes that may spell it as
/// either `r:id` or bare `id`.
std::string relationship_ref_id(const pugi::xml_node& node);

/// Lowercases an ASCII part extension. `[Content_Types].xml` matches
/// `Default Extension=` case-insensitively, so both the reader and the
/// writer normalise before comparing.
std::string lowercase_extension(std::string_view extension);

/// Returns the lowercased extension of a part path, without the dot.
/// Empty when the path has no extension, ends in a dot, or the only dot
/// belongs to a directory segment.
std::string extension_of_part(std::string_view path);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PACKAGE_VALIDATOR_H_
