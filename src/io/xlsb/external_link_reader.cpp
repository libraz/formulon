
#include "io/xlsb/external_link_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "external_book.h"
#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Opcodes of the two reference Ptgs a supporting-workbook defined name
/// stores. Both appear unclassed (no value/array class bit) in every
/// sample measured.
constexpr std::uint8_t kPtgRef3d = 0x3A;
constexpr std::uint8_t kPtgArea3d = 0x3B;

/// Byte length of each, inside an external link part: an opcode plus
/// four (Ref3d) or six (Area3d) 16-bit fields. These are *not* the
/// worksheet layouts of the same opcodes, which use 32-bit rows and a
/// single `ixti`.
constexpr std::size_t kExternRef3dBytes = 9U;
constexpr std::size_t kExternArea3dBytes = 13U;

Error CorruptError(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbCorrupt, message, "context=xlsb_external_link_reader");
}

/// Decodes a defined name's stored formula into the rectangle it names.
///
/// The two accepted shapes are a single reference and a rectangle, each
/// naming exactly one sheet of the supporting workbook. A multi-sheet
/// span, or any other token, leaves `out->resolvable` false: the name
/// exists but this cache cannot say what it points at, and the reference
/// reads `#REF!` rather than resolving against guessed coordinates.
void DecodeNameFormula(ByteSpan rgce, ExternalBookName* out) {
  if (rgce.size == 0) {
    return;
  }
  const std::uint8_t opcode = rgce.data[0];
  const auto field = [&rgce](std::size_t index) -> std::uint32_t {
    const std::size_t offset = 1U + index * 2U;
    return static_cast<std::uint32_t>(rgce.data[offset]) | (static_cast<std::uint32_t>(rgce.data[offset + 1U]) << 8U);
  };
  if (opcode == kPtgRef3d && rgce.size == kExternRef3dBytes) {
    if (field(0) != field(1)) {
      return;  // A span across sheets names no single rectangle.
    }
    out->sheet = field(0);
    out->row = field(2);
    out->col = field(3);
    out->row_end = out->row;
    out->col_end = out->col;
    out->is_range = false;
    out->resolvable = true;
    return;
  }
  if (opcode == kPtgArea3d && rgce.size == kExternArea3dBytes) {
    if (field(0) != field(1)) {
      return;
    }
    out->sheet = field(0);
    out->row = field(2);
    out->row_end = field(3);
    out->col = field(4);
    out->col_end = field(5);
    out->is_range = true;
    out->resolvable = true;
  }
}

}  // namespace

Expected<ExternalBook, Error> read_external_link_bin(ByteSpan cursor, std::string* book_rel_id) {
  ExternalBook book;
  // The record stream is flat: a name's `BrtExternNameFmla` applies to
  // the `BrtExternNameStart` before it, and a cell record to the most
  // recent row header inside the most recent sheet table.
  std::uint32_t current_sheet = ExternalBook::kNoSheet;
  std::uint32_t current_row = 0;
  bool row_seen = false;

  const auto put_cell = [&book, &current_sheet, &current_row, &row_seen](std::uint32_t col, ExternalCell cell) -> bool {
    if (current_sheet == ExternalBook::kNoSheet || !row_seen) {
      return false;
    }
    book.cells.emplace(ExternalBook::cell_key(current_sheet, current_row, col), std::move(cell));
    return true;
  };

  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    ByteSpan payload = rec.payload;
    switch (static_cast<XlsbRecordType>(rec.type)) {
      case XlsbRecordType::BrtBeginExternalBook: {
        // `sbt` (u16, 0 for a workbook), then the rel id as a wide string.
        if (book_rel_id == nullptr) {
          break;
        }
        RETURN_IF_ERROR(read_u16(payload));
        ASSIGN_OR_RETURN(auto rel, read_xlwidestring(payload));
        *book_rel_id = std::move(rel);
        break;
      }
      case XlsbRecordType::BrtSupTabs: {
        ASSIGN_OR_RETURN(auto count, read_u32(payload));
        for (std::uint32_t i = 0; i < count; ++i) {
          ASSIGN_OR_RETURN(auto name, read_xlwidestring(payload));
          book.sheet_names.push_back(std::move(name));
        }
        break;
      }
      case XlsbRecordType::BrtExternNameStart: {
        ASSIGN_OR_RETURN(auto name, read_xlwidestring(payload));
        ExternalBookName entry;
        entry.name = std::move(name);
        book.names.push_back(std::move(entry));
        break;
      }
      case XlsbRecordType::BrtExternNameFmla: {
        if (book.names.empty()) {
          return CorruptError("xlsb external name formula with no name to attach to");
        }
        ASSIGN_OR_RETURN(auto cce, read_u32(payload));
        if (cce > payload.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb external name formula truncated",
                            "context=xlsb_external_link_reader");
        }
        // No body at all: the supporting book declares no such name.
        book.names.back().exists = cce != 0U;
        DecodeNameFormula(ByteSpan{payload.data, cce}, &book.names.back());
        break;
      }
      case XlsbRecordType::BrtBeginExternTable: {
        ASSIGN_OR_RETURN(auto sheet, read_u32(payload));
        if (sheet >= book.sheet_names.size()) {
          return CorruptError("xlsb external cached sheet index is outside the supporting book's sheet table");
        }
        current_sheet = sheet;
        book.sheet_data.resize(book.sheet_names.size(), false);
        book.sheet_data[current_sheet] = true;
        row_seen = false;
        break;
      }
      case XlsbRecordType::BrtEndExternTable:
        current_sheet = ExternalBook::kNoSheet;
        row_seen = false;
        break;
      case XlsbRecordType::BrtExternRowHdr: {
        ASSIGN_OR_RETURN(current_row, read_u32(payload));
        row_seen = true;
        break;
      }
      case XlsbRecordType::BrtExternCellReal: {
        ASSIGN_OR_RETURN(auto col, read_u32(payload));
        ASSIGN_OR_RETURN(auto value, read_double(payload));
        ExternalCell cell;
        cell.value = Value::number(value);
        if (!put_cell(col, std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellBool: {
        ASSIGN_OR_RETURN(auto col, read_u32(payload));
        ASSIGN_OR_RETURN(auto flag, read_u8(payload));
        ExternalCell cell;
        cell.value = Value::boolean(flag != 0U);
        if (!put_cell(col, std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellError: {
        ASSIGN_OR_RETURN(auto col, read_u32(payload));
        ASSIGN_OR_RETURN(auto code, read_u8(payload));
        ExternalCell cell;
        // The byte is the OOXML wire code, the same one a cached error
        // cell carries inside a worksheet.
        cell.value = Value::error(error_from_ooxml_code(static_cast<std::int32_t>(code)));
        if (!put_cell(col, std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellString: {
        ASSIGN_OR_RETURN(auto col, read_u32(payload));
        ASSIGN_OR_RETURN(auto text, read_xlwidestring(payload));
        ExternalCell cell;
        // The kind is carried on `value` and the bytes on `text`; see
        // `ExternalCell`. Binding a `Value::text` to the local string
        // here would leave the cell aliasing a dead buffer.
        cell.value = Value::text({});
        cell.text = std::move(text);
        if (!put_cell(col, std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      default:
        // The part also carries its own framing and the alternate-URL
        // future records. Neither contributes to the cache.
        break;
    }
  }
  return book;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
