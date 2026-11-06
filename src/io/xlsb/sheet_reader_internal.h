// Declarations shared between the XLSB worksheet reader's record loop, its
// cell-record decoding and its tail-record retention.

#ifndef FORMULON_IO_XLSB_SHEET_READER_INTERNAL_H_
#define FORMULON_IO_XLSB_SHEET_READER_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/sheet_reader.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

/// What became of one source record. The four values are exhaustive: a
/// record is decoded into the model, kept verbatim for re-emission,
/// left to the writer to re-derive, or reported as content this load
/// could not carry. There is no fifth outcome, and in particular no
/// outcome that discards a record without saying so.
///
/// `[[nodiscard]]` is deliberate. The defect this classification exists
/// to prevent is a record reaching the end of the dispatch with nothing
/// done to it, so the dispatch's answer must not be droppable either.
enum class [[nodiscard]] RecordDisposition {
  /// Decoded into the `Workbook` / `Sheet` model. Also covers a record
  /// whose content the reader took responsibility for and whose partial
  /// loss it reported through a more specific counter -- a formula whose
  /// Ptg stream fell outside the supported set is counted once, as an
  /// undecoded formula, not twice.
  kModelled,
  /// Kept verbatim in `XlsbSheetTail` and re-emitted byte-for-byte.
  kRetained,
  /// Not kept, because the writer emits an equivalent record of its own
  /// from the model. The membership test is
  /// `IsWriterRegeneratedSheetRecord`.
  kRegenerated,
  /// Neither decoded nor preserved: the content is reported through
  /// `XlsbReadResult::dropped_record_count` and a structured warning.
  /// A record decoded only in part resolves here too, so the residue is
  /// visible rather than implied.
  kAccounted,
};

/// Retains one worksheet-tail record in its grammar slot.
void RetainTailRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed, std::size_t size);

/// Decides what happens to a record the dispatch did not decode.
RecordDisposition ResolveUnmodelledRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed,
                                          std::size_t size);

/// Decodes a cell's `rgce` Ptg stream into formula text, or empty when it cannot.
std::string DecodeFormulaText(ByteSpan ptg_bytes, ByteSpan rgcb, const std::vector<std::string>& sheet_names,
                              const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                              const XlsbExternalBooks& external_books, std::size_t sheet_index, std::uint32_t row,
                              std::uint32_t col, std::uint32_t* undecoded_formula_count,
                              const std::vector<TableMetadata>* tables);

/// Column + style-table index decoded by `ReadCellHeader`.
struct CellHeaderInfo {
  std::uint32_t col = 0;
  /// 0-based index into `StylesTable::cell_xfs`. `0` is the default xf;
  /// value records skip storing it, blank and formula records store it as an
  /// explicit assignment.
  std::uint32_t xf_index = 0;
  /// fPhShow: the cell shows its phonetic guide.
  bool show_phonetic = false;
};

/// Reads the common cell header and advances `cursor` past it.
Expected<CellHeaderInfo, Error> ReadCellHeader(ByteSpan& cursor);

/// Stores the cell's xf and counts it as decoded.
Expected<void, Error> CommitCell(Workbook& wb, std::size_t sheet_index, SheetDecodeState& state,
                                 const CellHeaderInfo& header, bool explicit_xf);

/// Slices a `CellParsedFormula` into its rgce and rgcb spans.
Expected<void, Error> ReadCellParsedFormula(ByteSpan& p, const char* what, ByteSpan* rgce, ByteSpan* rgcb);

/// Decodes one literal cell record (blank, RK, real, bool, error, inline or shared string).
Expected<RecordDisposition, Error> DecodeLiteralCell(const XlsbRecord& rec, XlsbRecordType type,
                                                     SheetDecodeState& state, std::size_t sheet_index, Workbook& wb,
                                                     const std::vector<std::string_view>& sst_entries,
                                                     const std::vector<std::vector<PhoneticRun>>& sst_phonetic,
                                                     const std::vector<PhoneticProperties>& sst_phonetic_props,
                                                     std::deque<std::string>& text_storage);

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_READER_INTERNAL_H_
