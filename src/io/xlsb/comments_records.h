//
// The worksheet comments part of an xlsb package (`xl/commentsN.bin`): the
// author list, then per note a revision-uid wrapper, its cell, and its text as
// a `RichStr` whose runs name fonts in the styles part's font table.

#ifndef FORMULON_IO_XLSB_COMMENTS_RECORDS_H_
#define FORMULON_IO_XLSB_COMMENTS_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/record.h"
#include "sheet_metadata.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes a comments part. A run's font index resolves through `fonts`; one
/// outside it leaves the run's font unset. `authors_out` receives the author
/// list in stored order.
Expected<std::vector<CellComment>, Error> read_comments_bin(ByteSpan part, const std::vector<FontRecord>& fonts,
                                                            std::vector<std::string>* authors_out);

/// The font table index of each note run's font: the table's own entries,
/// then the run fonts it lacks, appended in note order.
class NoteRunFonts {
 public:
  /// Indexes `table` (the fonts the styles part writes) and appends the run
  /// fonts of `notes` it lacks.
  NoteRunFonts(const std::vector<FontRecord>& table, const std::vector<const std::vector<CellComment>*>& notes);

  /// The index `font` is written with; 0 for a run without a font.
  std::uint16_t index_of(const std::optional<FontRecord>& font) const;

  /// The fonts appended after the table.
  const std::vector<FontRecord>& appended() const noexcept { return appended_; }

 private:
  /// The BrtFont bytes of every font, in table order; a font's index is its position.
  std::vector<std::string> keys_;
  std::vector<FontRecord> appended_;
};

/// Encodes `notes` (a sheet's notes and thread stubs) as a comments part, in
/// row-major cell order. A note without a uid takes the one Excel assigns,
/// built from `sheet_index` and its position.
std::vector<std::uint8_t> write_comments_bin(std::vector<CellComment> notes, std::size_t sheet_index,
                                             const NoteRunFonts& fonts);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_COMMENTS_RECORDS_H_
