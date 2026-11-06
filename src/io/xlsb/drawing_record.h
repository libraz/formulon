//
// Worksheet-tail BrtDrawing handling: reading the relationship id of the one
// record a sheet carries, removing retained records and splicing in the
// record for a model-owned drawing at its grammar slot, plus the check that
// a drawing part holds only frames an XLSB package can carry.

#ifndef FORMULON_IO_XLSB_DRAWING_RECORD_H_
#define FORMULON_IO_XLSB_DRAWING_RECORD_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {
namespace io {
namespace xlsb {

/// BrtDrawing records of a retained `after_hyperlinks` tail slot.
struct DrawingRecordScan {
  /// Number of BrtDrawing records.
  std::size_t count = 0;
  /// The relationship id, only when there is exactly one record and its
  /// payload is one XLWideString with no bytes left over.
  std::optional<std::string> rel_id;
  /// False when the buffer is non-empty and does not split into records.
  bool parsed = true;
};

/// Counts the BrtDrawing records of `after_hyperlinks` and reads the id of a
/// single one.
DrawingRecordScan scan_drawing_records(const std::vector<std::uint8_t>& after_hyperlinks);

/// Removes every BrtDrawing record. A non-empty buffer that does not split
/// into records is left untouched and yields false.
bool strip_drawing_records(std::vector<std::uint8_t>& after_hyperlinks);

/// Replaces the BrtDrawing records with one naming `rel_id`, placed after the
/// last record the grammar puts ahead of the drawing (page setup through
/// smart tags), or first when there is none. A non-empty buffer that does
/// not split into records is left untouched and yields false.
bool splice_drawing(std::vector<std::uint8_t>& after_hyperlinks, std::string_view rel_id);

/// True when `drawing_xml` parses and every `graphicData` in it, inside
/// `mc:AlternateContent` included, names a chart. Pictures, shapes and
/// groups carry no `graphicData` and always pass.
bool drawing_frames_are_writable(std::string_view drawing_xml);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_DRAWING_RECORD_H_
