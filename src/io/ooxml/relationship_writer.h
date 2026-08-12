//
// Relationship / Override XML emission helpers for the OOXML writer.
// Path-shaping helpers that turn package paths into the various
// relative forms OOXML expects, plus the actual `<Relationship>` /
// `<Override>` element writers. Internal to `src/io/ooxml/`; not part
// of the public API.
//
// The shared content-type and relationship-type URI constants live in
// the writer-side TUs themselves (anonymous namespace of
// `ooxml_writer.cpp`) rather than this header: keeping them with
// internal linkage avoids ODR-emit overhead in WASM for sibling TUs
// that do not actually consume the strings (e.g.
// `emission_plan.cpp`).

#ifndef FORMULON_IO_OOXML_RELATIONSHIP_WRITER_H_
#define FORMULON_IO_OOXML_RELATIONSHIP_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace io {

// ---------------------------------------------------------------------------
// Path-shaping helpers
// ---------------------------------------------------------------------------

/// Strips a leading `"xl/"` segment when present; the value is returned
/// as a view into the original path. Used to convert package paths
/// into the relative form Excel emits for workbook-level relationships.
std::string_view WithoutXlPrefix(std::string_view path);

/// Builds the `Target=` value for a worksheet rels file: prefixes the
/// path with `"../"` so it resolves against the `xl/worksheets/`
/// directory (e.g. `"xl/printerSettings/printerSettings1.bin"` ->
/// `"../printerSettings/printerSettings1.bin"`).
std::string TargetRelativeToWorksheet(std::string_view package_path);

// ---------------------------------------------------------------------------
// Relationship / Override XML emitters
// ---------------------------------------------------------------------------

/// Appends a single `<Override PartName="/<path>" ContentType="<ct>"/>`
/// entry plus its trailing newline. Used by the package-level Content
/// Types builder for the per-table / per-pivot / per-comments /
/// passthrough Override blocks. `path` is escaped to defend against
/// passthrough paths carrying XML-critical characters; `ct` is a
/// writer-controlled string view and is emitted verbatim.
void AppendOverride(std::string& out, std::string_view path, std::string_view ct, bool escape_path = false);

/// Appends a single `<Relationship Id="<id>" Type="<type>"
/// Target="<target>"/>` entry plus its trailing newline.
void AppendRelationship(std::string& out, std::string_view id, std::string_view type, std::string_view target,
                        bool target_external = false, bool escape_target = false);

/// Same as the string-view overload but formats `"rId<rid>"` as the
/// relationship id.
void AppendRelationship(std::string& out, std::uint32_t rid, std::string_view type, std::string_view target,
                        bool target_external = false, bool escape_target = false);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_RELATIONSHIP_WRITER_H_
