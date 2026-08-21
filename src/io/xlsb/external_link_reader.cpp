
#include "io/xlsb/external_link_reader.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <utility>

#include "io/external_book.h"
#include "io/xlsb/record.h"
#include "utils/error.h"
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

Expected<double, Error> ReadDouble(ByteSpan& cursor) {
  if (cursor.size < sizeof(double)) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb extern cell double truncated",
                      "context=xlsb_external_link_reader");
  }
  double value = 0.0;
  std::memcpy(&value, cursor.data, sizeof(value));
  cursor.data += sizeof(double);
  cursor.size -= sizeof(double);
  return value;
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

Expected<ExternalBook, Error> read_external_link_bin(ByteSpan cursor) {
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
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    ByteSpan payload = rec.payload;
    switch (static_cast<XlsbRecordType>(rec.type)) {
      case XlsbRecordType::BrtSupTabs: {
        auto count_or = read_u32(payload);
        if (!count_or) {
          return count_or.error();
        }
        for (std::uint32_t i = 0; i < count_or.value(); ++i) {
          auto name_or = read_xlwidestring(payload);
          if (!name_or) {
            return name_or.error();
          }
          book.sheet_names.push_back(std::move(name_or.value()));
        }
        break;
      }
      case XlsbRecordType::BrtExternNameStart: {
        auto name_or = read_xlwidestring(payload);
        if (!name_or) {
          return name_or.error();
        }
        ExternalBookName entry;
        entry.name = std::move(name_or.value());
        book.names.push_back(std::move(entry));
        break;
      }
      case XlsbRecordType::BrtExternNameFmla: {
        if (book.names.empty()) {
          return CorruptError("xlsb external name formula with no name to attach to");
        }
        auto cce_or = read_u32(payload);
        if (!cce_or) {
          return cce_or.error();
        }
        if (cce_or.value() > payload.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb external name formula truncated",
                            "context=xlsb_external_link_reader");
        }
        DecodeNameFormula(ByteSpan{payload.data, cce_or.value()}, &book.names.back());
        break;
      }
      case XlsbRecordType::BrtBeginExternTable: {
        auto sheet_or = read_u32(payload);
        if (!sheet_or) {
          return sheet_or.error();
        }
        if (sheet_or.value() >= book.sheet_names.size()) {
          return CorruptError("xlsb external cached sheet index is outside the supporting book's sheet table");
        }
        current_sheet = sheet_or.value();
        row_seen = false;
        break;
      }
      case XlsbRecordType::BrtEndExternTable:
        current_sheet = ExternalBook::kNoSheet;
        row_seen = false;
        break;
      case XlsbRecordType::BrtExternRowHdr: {
        auto row_or = read_u32(payload);
        if (!row_or) {
          return row_or.error();
        }
        current_row = row_or.value();
        row_seen = true;
        break;
      }
      case XlsbRecordType::BrtExternCellReal: {
        auto col_or = read_u32(payload);
        if (!col_or) {
          return col_or.error();
        }
        auto value_or = ReadDouble(payload);
        if (!value_or) {
          return value_or.error();
        }
        ExternalCell cell;
        cell.value = Value::number(value_or.value());
        if (!put_cell(col_or.value(), std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellBool: {
        auto col_or = read_u32(payload);
        if (!col_or) {
          return col_or.error();
        }
        auto flag_or = read_u8(payload);
        if (!flag_or) {
          return flag_or.error();
        }
        ExternalCell cell;
        cell.value = Value::boolean(flag_or.value() != 0U);
        if (!put_cell(col_or.value(), std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellError: {
        auto col_or = read_u32(payload);
        if (!col_or) {
          return col_or.error();
        }
        auto code_or = read_u8(payload);
        if (!code_or) {
          return code_or.error();
        }
        ExternalCell cell;
        // The byte is the OOXML wire code, the same one a cached error
        // cell carries inside a worksheet.
        cell.value = Value::error(error_from_ooxml_code(static_cast<std::int32_t>(code_or.value())));
        if (!put_cell(col_or.value(), std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      case XlsbRecordType::BrtExternCellString: {
        auto col_or = read_u32(payload);
        if (!col_or) {
          return col_or.error();
        }
        auto text_or = read_xlwidestring(payload);
        if (!text_or) {
          return text_or.error();
        }
        ExternalCell cell;
        // The kind is carried on `value` and the bytes on `text`; see
        // `ExternalCell`. Binding a `Value::text` to the local string
        // here would leave the cell aliasing a dead buffer.
        cell.value = Value::text({});
        cell.text = std::move(text_or.value());
        if (!put_cell(col_or.value(), std::move(cell))) {
          return CorruptError("xlsb external cached cell appears outside a sheet table");
        }
        break;
      }
      default:
        // The part also carries its own framing, the link's relationship
        // id, and the alternate-URL future records. None of them
        // contributes to the cache.
        break;
    }
  }
  return book;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
