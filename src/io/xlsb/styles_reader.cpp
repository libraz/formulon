// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the `xl/styles.bin` reader. See
// `io/xlsb/styles_reader.h` for the contract.
//
// Record shapes below were derived from a real Excel-365-produced
// `xl/styles.bin` (byte-level verification, not transcribed from a
// third-party summary):
//
//   BrtFmt  (44): u16 ifmt, u32 cch, cch x UTF-16LE code units.
//   BrtXF   (47): 8 x u16 — [xfId_or_parent, numFmtId, fontId, fillId,
//                 borderId, reserved, alignmentFlags, applyFlags].
//                 Appears once per entry in both the `<cellStyleXfs>`
//                 block (bracketed by BrtBeginCellStyleXFs /
//                 BrtEndCellStyleXFs) and the `<cellXfs>` block
//                 (bracketed by BrtBeginCellXFs / BrtEndCellXFs); which
//                 block a given BrtXF belongs to is tracked by the most
//                 recently seen Begin marker.
//
// BrtFont / BrtFill / BrtBorder are counted but not decoded (see the
// header's doc comment); this reader pushes one default-constructed
// record per occurrence so index-based lookups from `BrtXF` stay valid.

#include "io/xlsb/styles_reader.h"

#include <cstdint>
#include <string>
#include <utility>

#include "io/xlsb/record.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Which `<cellXfs>`-family table is currently being populated. `BrtXF`
// records outside both brackets (should not happen in a well-formed
// part, but bounds-checked input requires a defined behaviour) are
// dropped rather than misattributed.
enum class XfTarget { kNone, kCellStyleXfs, kCellXfs };

Expected<void, Error> DecodeFmt(ByteSpan payload, StylesTable& table) {
  ByteSpan p = payload;
  auto ifmt_or = read_u16(p);
  if (!ifmt_or) {
    return ifmt_or.error();
  }
  auto name_or = read_xlwidestring(p);
  if (!name_or) {
    return name_or.error();
  }
  NumFmtRecord rec;
  rec.id = ifmt_or.value();
  rec.format_string_index = static_cast<std::uint32_t>(table.num_fmt_strings.size());
  table.num_fmt_strings.push_back(std::move(name_or.value()));
  table.num_fmts.push_back(rec);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeXf(ByteSpan payload, XfTarget target, StylesTable& table) {
  ByteSpan p = payload;
  // 8 x u16: [xfId_or_parent, numFmtId, fontId, fillId, borderId,
  // reserved, alignmentFlags, applyFlags]. Only the format/font/fill/
  // border fields are consumed here; alignment and apply-flags are
  // round-tripped via the raw `xl/styles.bin` passthrough copy instead
  // of being modelled in `CellXf`.
  auto skip_parent = read_u16(p);
  if (!skip_parent) {
    return skip_parent.error();
  }
  auto num_fmt_id_or = read_u16(p);
  if (!num_fmt_id_or) {
    return num_fmt_id_or.error();
  }
  auto font_id_or = read_u16(p);
  if (!font_id_or) {
    return font_id_or.error();
  }
  auto fill_id_or = read_u16(p);
  if (!fill_id_or) {
    return fill_id_or.error();
  }
  auto border_id_or = read_u16(p);
  if (!border_id_or) {
    return border_id_or.error();
  }
  CellXf xf;
  xf.num_fmt_id = num_fmt_id_or.value();
  xf.font_index = font_id_or.value();
  xf.fill_index = fill_id_or.value();
  xf.border_index = border_id_or.value();
  switch (target) {
    case XfTarget::kCellStyleXfs:
      table.cell_style_xfs.push_back(xf);
      break;
    case XfTarget::kCellXfs:
      table.cell_xfs.push_back(xf);
      break;
    case XfTarget::kNone:
      // A `BrtXF` outside any bracket has no destination table; drop it
      // rather than guessing. Real Excel output never hits this path.
      break;
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<StylesTable, Error> read_styles_bin(ByteSpan bytes) {
  StylesTable table;
  XfTarget xf_target = XfTarget::kNone;
  ByteSpan cursor = bytes;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    switch (static_cast<XlsbRecordType>(rec.type)) {
      case XlsbRecordType::BrtFmt: {
        if (auto r = DecodeFmt(rec.payload, table); !r) {
          return r.error();
        }
        break;
      }
      case XlsbRecordType::BrtFont:
        table.fonts.push_back(FontRecord{});
        break;
      case XlsbRecordType::BrtFill:
        table.fills.push_back(FillRecord{});
        break;
      case XlsbRecordType::BrtBorder:
        table.borders.push_back(BorderRecord{});
        break;
      case XlsbRecordType::BrtBeginCellStyleXFs:
        xf_target = XfTarget::kCellStyleXfs;
        break;
      case XlsbRecordType::BrtEndCellStyleXFs:
        xf_target = XfTarget::kNone;
        break;
      case XlsbRecordType::BrtBeginCellXFs:
        xf_target = XfTarget::kCellXfs;
        break;
      case XlsbRecordType::BrtEndCellXFs:
        xf_target = XfTarget::kNone;
        break;
      case XlsbRecordType::BrtXF: {
        if (auto r = DecodeXf(rec.payload, xf_target, table); !r) {
          return r.error();
        }
        break;
      }
      default:
        break;
    }
  }
  // Empty-document contract mirrors `io::read_styles`: every consumer
  // indexes `cell_xfs[xf_index]` / `fonts[font_index]` / etc. without a
  // bounds check once `xf_index` itself has been validated, so a part
  // that carried none of the relevant records still needs a resolvable
  // default row in each vector.
  if (table.fonts.empty()) {
    table.fonts.push_back(FontRecord{});
  }
  if (table.fills.empty()) {
    table.fills.push_back(FillRecord{});
  }
  if (table.borders.empty()) {
    table.borders.push_back(BorderRecord{});
  }
  if (table.cell_xfs.empty()) {
    table.cell_xfs.push_back(CellXf{});
  }
  return table;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
