#include "io/xlsb/pivot_table_records.h"

#include <algorithm>
#include <cstddef>
#include <optional>
#include <string_view>
#include <utility>

#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace ids = pivot_records;

constexpr const char* kContext = "context=xlsb_pivot_table_records";
constexpr std::size_t kViewHeaderSize = 32;
constexpr std::size_t kLocationSize = 36;
constexpr std::size_t kFieldHeaderSize = 20;
constexpr std::size_t kItemHeaderSize = 7;
constexpr std::size_t kPageFieldSize = 13;
constexpr std::size_t kDataFieldHeaderSize = 25;
constexpr std::size_t kUidSize = 16;
/// BrtBeginSXDI byte 24: a name follows.
constexpr std::uint8_t kDataFieldHasName = 0x01;
constexpr std::uint16_t kKnownItemFlags =
    kSxItemHidden | kSxItemCollapsed | kSxItemHasCaption | kSxItemCalculated | kSxItemMissing;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, kContext);
}

Error Unexpected(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, message, kContext);
}

/// Sequential reader over the part's framed records.
class Records {
 public:
  explicit Records(std::vector<XlsbRecord> records) : records_(std::move(records)) {}

  bool peek(std::uint16_t type) const { return pos_ < records_.size() && records_[pos_].type == type; }

  /// Payload of the next record, which must be of `type`.
  Expected<ByteSpan, Error> take(std::uint16_t type, const char* missing) {
    if (!peek(type)) {
      return Unexpected(missing);
    }
    return records_[pos_++].payload;
  }

  /// Consumes the next record, which must be an empty record of `type`.
  Expected<void, Error> take_end(std::uint16_t type, const char* missing) {
    ASSIGN_OR_RETURN(const ByteSpan payload, take(type, missing));
    if (payload.size != 0U) {
      return Corrupt("xlsb pivot table end record carries a payload");
    }
    return Expected<void, Error>::Ok();
  }

  /// Copies the records from the next one through the first of type `end`.
  Expected<std::vector<ids::RawRecord>, Error> take_raw_block(std::uint16_t end) {
    std::vector<ids::RawRecord> out;
    while (pos_ < records_.size()) {
      const XlsbRecord& rec = records_[pos_++];
      out.push_back(
          ids::RawRecord{rec.type, std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size)});
      if (rec.type == end) {
        return out;
      }
    }
    return Unexpected("xlsb pivot table raw block is not closed");
  }

  const std::vector<XlsbRecord>& all() const { return records_; }
  std::size_t* pos() { return &pos_; }
  bool done() const { return pos_ == records_.size(); }

 private:
  std::vector<XlsbRecord> records_;
  std::size_t pos_ = 0;
};

/// Reads a u32 count record payload.
Expected<std::uint32_t, Error> CountOf(ByteSpan payload) {
  if (payload.size != 4U) {
    return Corrupt("xlsb pivot table count record has an unexpected size");
  }
  return load_u32(payload.data);
}

std::vector<std::uint8_t> Copy(ByteSpan payload) {
  return std::vector<std::uint8_t>(payload.data, payload.data + payload.size);
}

Expected<SxView, Error> DecodeView(ByteSpan payload) {
  if (payload.size < kViewHeaderSize) {
    return Corrupt("xlsb pivot table view header truncated");
  }
  const std::uint8_t* p = payload.data;
  SxView view;
  view.created_version = p[0];
  view.flags1 = p[1];
  view.flags2 = p[2];
  view.flags3 = p[3];
  view.flags4 = p[4];
  view.flags5 = p[5];
  view.flags6 = p[6];
  view.flags7 = p[7];
  view.flags8 = p[8];
  view.flags9 = p[9];
  view.flags10 = p[10];
  view.flags11 = p[11];
  view.data_axis = p[12];
  view.flags13 = p[13];
  view.updated_version = p[14];
  view.min_refreshable_version = p[15];
  view.data_position = load_u32(p + 16);
  view.reserved20 = load_u32(p + 20);
  view.reserved24 = load_u32(p + 24);
  view.cache_id = load_u32(p + 28);
  ByteSpan rest{p + kViewHeaderSize, payload.size - kViewHeaderSize};
  ASSIGN_OR_RETURN(view.name, read_xlwidestring(rest));
  ASSIGN_OR_RETURN(view.data_caption, read_xlwidestring(rest));
  const bool grand = (view.flags6 & kSxViewHasGrandCaption) != 0U;
  const bool error = (view.flags8 & kSxViewNoErrorCaption) == 0U;
  const bool missing = (view.flags8 & kSxViewNoMissingCaption) == 0U;
  if (static_cast<int>(grand) + static_cast<int>(error) + static_cast<int>(missing) > 1) {
    return Corrupt("xlsb pivot table view carries several captions whose order is uncharacterised");
  }
  if (grand) {
    ASSIGN_OR_RETURN(view.grand_total_caption, read_xlwidestring(rest));
  } else if (error) {
    ASSIGN_OR_RETURN(view.error_caption, read_xlwidestring(rest));
  } else if (missing) {
    ASSIGN_OR_RETURN(view.missing_caption, read_xlwidestring(rest));
  }
  if (rest.size != 0U) {
    return Corrupt("xlsb pivot table view carries uncharacterised trailing data");
  }
  return view;
}

Expected<SxLocation, Error> DecodeLocation(ByteSpan payload) {
  if (payload.size != kLocationSize) {
    return Corrupt("xlsb pivot table location has an unexpected size");
  }
  const std::uint8_t* p = payload.data;
  SxLocation loc;
  loc.first_row = load_u32(p);
  loc.last_row = load_u32(p + 4);
  loc.first_col = load_u32(p + 8);
  loc.last_col = load_u32(p + 12);
  loc.first_header_row = load_u32(p + 16);
  loc.first_data_row = load_u32(p + 20);
  loc.first_data_col = load_u32(p + 24);
  loc.row_page_count = load_u32(p + 28);
  loc.col_page_count = load_u32(p + 32);
  return loc;
}

Expected<SxItem, Error> DecodeItem(ByteSpan payload) {
  if (payload.size < kItemHeaderSize) {
    return Corrupt("xlsb pivot field item truncated");
  }
  SxItem item;
  item.kind = payload.data[0];
  item.flags = load_u16(payload.data + 1);
  item.index = load_u32(payload.data + 3);
  if (item.kind > kSxItemKindMax || (item.flags & ~kKnownItemFlags) != 0U) {
    return Corrupt("xlsb pivot field item carries a kind or flag this codec has not characterised");
  }
  ByteSpan rest{payload.data + kItemHeaderSize, payload.size - kItemHeaderSize};
  if ((item.flags & kSxItemHasCaption) != 0U) {
    ASSIGN_OR_RETURN(item.caption, read_xlwidestring(rest));
  }
  if (rest.size != 0U) {
    return Corrupt("xlsb pivot field item carries uncharacterised trailing data");
  }
  return item;
}

Expected<SxField, Error> DecodeField(Records& recs) {
  ASSIGN_OR_RETURN(const ByteSpan payload, recs.take(ids::kBeginSXVD, "xlsb pivot field expected"));
  if (payload.size < kFieldHeaderSize) {
    return Corrupt("xlsb pivot field header truncated");
  }
  const std::uint8_t* p = payload.data;
  SxField field;
  field.axis = p[0];
  field.subtotals = load_u16(p + 1);
  field.name_flags = p[3];
  field.num_fmt = load_u32(p + 4);
  field.flags = load_u32(p + 8);
  field.item_page_count = load_u32(p + 12);
  field.reserved16 = load_u32(p + 16);
  ByteSpan rest{p + kFieldHeaderSize, payload.size - kFieldHeaderSize};
  if ((field.name_flags & kSxFieldHasDisplayName) != 0U) {
    ASSIGN_OR_RETURN(field.display_name, read_xlwidestring(rest));
  }
  if (rest.size != 0U) {
    return Corrupt("xlsb pivot field header carries uncharacterised trailing data");
  }
  if (recs.peek(ids::kBeginSXVIs)) {
    ASSIGN_OR_RETURN(const ByteSpan begin, recs.take(ids::kBeginSXVIs, "xlsb pivot field items expected"));
    ASSIGN_OR_RETURN(const std::uint32_t declared, CountOf(begin));
    std::vector<SxItem> items;
    while (recs.peek(ids::kBeginSXVI)) {
      ASSIGN_OR_RETURN(const ByteSpan item_payload, recs.take(ids::kBeginSXVI, "xlsb pivot field item expected"));
      ASSIGN_OR_RETURN(SxItem item, DecodeItem(item_payload));
      RETURN_IF_ERROR(recs.take_end(ids::kEndSXVI, "xlsb pivot field item is not closed"));
      items.push_back(std::move(item));
    }
    RETURN_IF_ERROR(recs.take_end(ids::kEndSXVIs, "xlsb pivot field item list carries an uncharacterised record"));
    if (items.size() != declared) {
      return Corrupt("xlsb pivot field item count does not match the items decoded");
    }
    field.items = std::move(items);
  }
  if (recs.peek(ids::kBeginAutoSortScope)) {
    ASSIGN_OR_RETURN(field.auto_sort_scope, recs.take_raw_block(ids::kEndAutoSortScope));
  }
  RETURN_IF_ERROR(recs.take_end(ids::kEndSXVD, "xlsb pivot field carries an uncharacterised record"));
  return field;
}

Expected<std::vector<std::uint32_t>, Error> DecodeFieldOrder(ByteSpan payload) {
  if (payload.size < 4U) {
    return Corrupt("xlsb pivot axis field list truncated");
  }
  const std::uint32_t count = load_u32(payload.data);
  if (payload.size != 4U + static_cast<std::size_t>(count) * 4U) {
    return Corrupt("xlsb pivot axis field list length does not match its count");
  }
  std::vector<std::uint32_t> order;
  order.reserve(count);
  for (std::uint32_t k = 0; k < count; ++k) {
    order.push_back(load_u32(payload.data + 4U + static_cast<std::size_t>(k) * 4U));
  }
  return order;
}

Expected<SxPageField, Error> DecodePageField(ByteSpan payload) {
  if (payload.size != kPageFieldSize || payload.data[12] != 0U) {
    return Corrupt("xlsb pivot page field carries a name or caption this codec has not characterised");
  }
  SxPageField page;
  page.field = load_u32(payload.data);
  page.item = load_u32(payload.data + 4);
  page.hierarchy = load_u32(payload.data + 8);
  return page;
}

Expected<SxDataField, Error> DecodeDataField(ByteSpan payload) {
  if (payload.size < kDataFieldHeaderSize) {
    return Corrupt("xlsb pivot data field header truncated");
  }
  const std::uint8_t* p = payload.data;
  if (p[24] != kDataFieldHasName) {
    return Corrupt("xlsb pivot data field without a name is uncharacterised");
  }
  SxDataField data;
  data.field = load_u32(p);
  data.subtotal = load_u32(p + 4);
  data.show_data_as = load_u32(p + 8);
  data.base_field = load_u32(p + 12);
  data.base_item = load_u32(p + 16);
  data.num_fmt = load_u32(p + 20);
  ByteSpan rest{p + kDataFieldHeaderSize, payload.size - kDataFieldHeaderSize};
  ASSIGN_OR_RETURN(data.name, read_xlwidestring(rest));
  if (rest.size != 0U) {
    return Corrupt("xlsb pivot data field carries uncharacterised trailing data");
  }
  return data;
}

Expected<SxStyleClient, Error> DecodeStyle(ByteSpan payload) {
  if (payload.size < 6U) {
    return Corrupt("xlsb pivot style client truncated");
  }
  SxStyleClient style;
  style.flags = load_u16(payload.data);
  ByteSpan rest{payload.data + 2, payload.size - 2U};
  if (load_u32(rest.data) == 0xFFFFFFFFU) {
    rest.data += 4;
    rest.size -= 4U;
  } else {
    ASSIGN_OR_RETURN(style.name, read_xlwidestring(rest));
  }
  if (rest.size != 0U) {
    return Corrupt("xlsb pivot style client carries uncharacterised trailing data");
  }
  return style;
}

/// BrtFRTBegin, the `begin` / `end` record pair, BrtFRTEnd.
Expected<SxFrtBlock, Error> DecodeFrt(Records& recs, std::uint16_t begin, std::uint16_t end) {
  SxFrtBlock block;
  ASSIGN_OR_RETURN(const ByteSpan frt, recs.take(kFrtBegin, "xlsb pivot table future record block expected"));
  block.frt_payload = Copy(frt);
  ASSIGN_OR_RETURN(const ByteSpan payload, recs.take(begin, "xlsb pivot table future record expected"));
  block.payload = Copy(payload);
  RETURN_IF_ERROR(recs.take_end(end, "xlsb pivot table future record is not closed"));
  RETURN_IF_ERROR(recs.take_end(kFrtEnd, "xlsb pivot table future record block is not closed"));
  return block;
}

void EmitEmpty(std::vector<std::uint8_t>& dst, std::uint16_t type) {
  emit_record(dst, type, ByteSpan{});
}

void EmitCount(std::vector<std::uint8_t>& dst, std::uint16_t type, std::size_t count) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(count));
  emit_record(dst, type, payload);
}

void EmitView(std::vector<std::uint8_t>& dst, const SxView& view) {
  std::uint8_t flags6 = static_cast<std::uint8_t>(view.flags6 & ~kSxViewHasGrandCaption);
  if (view.grand_total_caption) {
    flags6 = static_cast<std::uint8_t>(flags6 | kSxViewHasGrandCaption);
  }
  std::uint8_t flags8 = static_cast<std::uint8_t>(view.flags8 | kSxViewNoErrorCaption | kSxViewNoMissingCaption);
  if (view.error_caption) {
    flags8 = static_cast<std::uint8_t>(flags8 & ~kSxViewNoErrorCaption);
  }
  if (view.missing_caption) {
    flags8 = static_cast<std::uint8_t>(flags8 & ~kSxViewNoMissingCaption);
  }
  std::vector<std::uint8_t> p;
  for (const std::uint8_t b : {view.created_version, view.flags1, view.flags2, view.flags3, view.flags4, view.flags5,
                               flags6, view.flags7, flags8, view.flags9, view.flags10, view.flags11, view.data_axis,
                               view.flags13, view.updated_version, view.min_refreshable_version}) {
    emit_u8(p, b);
  }
  emit_u32(p, view.data_position);
  emit_u32(p, view.reserved20);
  emit_u32(p, view.reserved24);
  emit_u32(p, view.cache_id);
  emit_xlwidestring(p, view.name);
  emit_xlwidestring(p, view.data_caption);
  for (const std::optional<std::string>* caption :
       {&view.grand_total_caption, &view.error_caption, &view.missing_caption}) {
    if (*caption) {
      emit_xlwidestring(p, **caption);
    }
  }
  emit_record(dst, ids::kBeginSXView, p);
}

void EmitLocation(std::vector<std::uint8_t>& dst, const SxLocation& loc) {
  std::vector<std::uint8_t> p;
  for (const std::uint32_t v : {loc.first_row, loc.last_row, loc.first_col, loc.last_col, loc.first_header_row,
                                loc.first_data_row, loc.first_data_col, loc.row_page_count, loc.col_page_count}) {
    emit_u32(p, v);
  }
  emit_record(dst, ids::kBeginSXLocation, p);
  EmitEmpty(dst, ids::kEndSXLocation);
}

void EmitRaw(std::vector<std::uint8_t>& dst, const std::vector<ids::RawRecord>& block) {
  for (const ids::RawRecord& rec : block) {
    emit_record(dst, rec.type, rec.payload);
  }
}

void EmitField(std::vector<std::uint8_t>& dst, const SxField& field) {
  std::uint8_t name_flags = static_cast<std::uint8_t>(field.name_flags & ~kSxFieldHasDisplayName);
  if (field.display_name) {
    name_flags = static_cast<std::uint8_t>(name_flags | kSxFieldHasDisplayName);
  }
  std::vector<std::uint8_t> p;
  emit_u8(p, field.axis);
  emit_u16(p, field.subtotals);
  emit_u8(p, name_flags);
  emit_u32(p, field.num_fmt);
  emit_u32(p, field.flags);
  emit_u32(p, field.item_page_count);
  emit_u32(p, field.reserved16);
  if (field.display_name) {
    emit_xlwidestring(p, *field.display_name);
  }
  emit_record(dst, ids::kBeginSXVD, p);
  if (field.items) {
    EmitCount(dst, ids::kBeginSXVIs, field.items->size());
    for (const SxItem& item : *field.items) {
      std::uint16_t flags = static_cast<std::uint16_t>(item.flags & ~kSxItemHasCaption);
      if (item.caption) {
        flags = static_cast<std::uint16_t>(flags | kSxItemHasCaption);
      }
      p.clear();
      emit_u8(p, item.kind);
      emit_u16(p, flags);
      emit_u32(p, item.index);
      if (item.caption) {
        emit_xlwidestring(p, *item.caption);
      }
      emit_record(dst, ids::kBeginSXVI, p);
      EmitEmpty(dst, ids::kEndSXVI);
    }
    EmitEmpty(dst, ids::kEndSXVIs);
  }
  EmitRaw(dst, field.auto_sort_scope);
  EmitEmpty(dst, ids::kEndSXVD);
}

void EmitFieldOrder(std::vector<std::uint8_t>& dst, std::uint16_t begin, std::uint16_t end,
                    const std::vector<std::uint32_t>& order) {
  std::vector<std::uint8_t> p;
  emit_u32(p, static_cast<std::uint32_t>(order.size()));
  for (const std::uint32_t v : order) {
    emit_u32(p, v);
  }
  emit_record(dst, begin, p);
  EmitEmpty(dst, end);
}

void EmitFrt(std::vector<std::uint8_t>& dst, const SxFrtBlock& block, std::uint16_t begin, std::uint16_t end) {
  emit_record(dst, kFrtBegin, block.frt_payload);
  emit_record(dst, begin, block.payload);
  EmitEmpty(dst, end);
  EmitEmpty(dst, kFrtEnd);
}

}  // namespace

Expected<PivotTableDoc, Error> decode_pivot_table(ByteSpan part) {
  std::vector<XlsbRecord> framed;
  while (part.size > 0U) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(part));
    framed.push_back(rec);
  }
  Records recs(std::move(framed));
  PivotTableDoc doc;

  if (recs.peek(ids::kACBegin)) {
    ASSIGN_OR_RETURN(const ByteSpan ac, recs.take(ids::kACBegin, "xlsb pivot table uid block expected"));
    if (ac.size != kAcUidBeginPayload.size() ||
        !std::equal(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end(), ac.data)) {
      return Corrupt("xlsb pivot table uid block carries an uncharacterised header");
    }
    ASSIGN_OR_RETURN(const ByteSpan uid, recs.take(ids::kUid, "xlsb pivot table uid expected"));
    if (uid.size != kUidSize) {
      return Corrupt("xlsb pivot table uid has an unexpected size");
    }
    std::array<std::uint8_t, kUidSize> bytes{};
    std::copy(uid.data, uid.data + kUidSize, bytes.begin());
    doc.uid = bytes;
    RETURN_IF_ERROR(recs.take_end(ids::kACEnd, "xlsb pivot table uid block is not closed"));
  }

  ASSIGN_OR_RETURN(const ByteSpan view, recs.take(ids::kBeginSXView, "xlsb pivot table view expected"));
  ASSIGN_OR_RETURN(doc.view, DecodeView(view));
  ASSIGN_OR_RETURN(const ByteSpan location, recs.take(ids::kBeginSXLocation, "xlsb pivot table location expected"));
  ASSIGN_OR_RETURN(doc.location, DecodeLocation(location));
  RETURN_IF_ERROR(recs.take_end(ids::kEndSXLocation, "xlsb pivot table location is not closed"));

  ASSIGN_OR_RETURN(const ByteSpan vds, recs.take(ids::kBeginSXVDs, "xlsb pivot table field list expected"));
  ASSIGN_OR_RETURN(const std::uint32_t field_count, CountOf(vds));
  while (recs.peek(ids::kBeginSXVD)) {
    ASSIGN_OR_RETURN(SxField field, DecodeField(recs));
    doc.fields.push_back(std::move(field));
  }
  RETURN_IF_ERROR(recs.take_end(ids::kEndSXVDs, "xlsb pivot table field list carries an uncharacterised record"));
  if (doc.fields.size() != field_count) {
    return Corrupt("xlsb pivot table field count does not match the fields decoded");
  }

  if (recs.peek(ids::kBeginISXVDRws)) {
    ASSIGN_OR_RETURN(const ByteSpan rows, recs.take(ids::kBeginISXVDRws, "xlsb pivot row fields expected"));
    ASSIGN_OR_RETURN(doc.row_fields, DecodeFieldOrder(rows));
    RETURN_IF_ERROR(recs.take_end(ids::kEndISXVDRws, "xlsb pivot row fields are not closed"));
  }
  ASSIGN_OR_RETURN(doc.row_items, decode_line_items(recs.all(), recs.pos(), SxLineItemAxis::kRows));
  if (recs.peek(ids::kBeginISXVDCols)) {
    ASSIGN_OR_RETURN(const ByteSpan cols, recs.take(ids::kBeginISXVDCols, "xlsb pivot column fields expected"));
    ASSIGN_OR_RETURN(doc.col_fields, DecodeFieldOrder(cols));
    RETURN_IF_ERROR(recs.take_end(ids::kEndISXVDCols, "xlsb pivot column fields are not closed"));
  }
  ASSIGN_OR_RETURN(doc.col_items, decode_line_items(recs.all(), recs.pos(), SxLineItemAxis::kCols));

  if (recs.peek(ids::kBeginSXPIs)) {
    ASSIGN_OR_RETURN(const ByteSpan pis, recs.take(ids::kBeginSXPIs, "xlsb pivot page fields expected"));
    ASSIGN_OR_RETURN(const std::uint32_t page_count, CountOf(pis));
    std::vector<SxPageField> pages;
    while (recs.peek(ids::kBeginSXPI)) {
      ASSIGN_OR_RETURN(const ByteSpan pi, recs.take(ids::kBeginSXPI, "xlsb pivot page field expected"));
      ASSIGN_OR_RETURN(SxPageField page, DecodePageField(pi));
      RETURN_IF_ERROR(recs.take_end(ids::kEndSXPI, "xlsb pivot page field is not closed"));
      pages.push_back(page);
    }
    RETURN_IF_ERROR(recs.take_end(ids::kEndSXPIs, "xlsb pivot page field list carries an uncharacterised record"));
    if (pages.size() != page_count) {
      return Corrupt("xlsb pivot page field count does not match the fields decoded");
    }
    doc.page_fields = std::move(pages);
  }

  ASSIGN_OR_RETURN(const ByteSpan dis, recs.take(ids::kBeginSXDIs, "xlsb pivot data fields expected"));
  ASSIGN_OR_RETURN(const std::uint32_t data_count, CountOf(dis));
  while (recs.peek(ids::kBeginSXDI)) {
    ASSIGN_OR_RETURN(const ByteSpan di, recs.take(ids::kBeginSXDI, "xlsb pivot data field expected"));
    ASSIGN_OR_RETURN(SxDataField data, DecodeDataField(di));
    RETURN_IF_ERROR(recs.take_end(ids::kEndSXDI, "xlsb pivot data field is not closed"));
    doc.data_fields.push_back(std::move(data));
  }
  RETURN_IF_ERROR(recs.take_end(ids::kEndSXDIs, "xlsb pivot data field list carries an uncharacterised record"));
  if (doc.data_fields.size() != data_count) {
    return Corrupt("xlsb pivot data field count does not match the fields decoded");
  }

  ASSIGN_OR_RETURN(const ByteSpan style, recs.take(ids::kTableStyleClient, "xlsb pivot style client expected"));
  ASSIGN_OR_RETURN(doc.style, DecodeStyle(style));
  if (recs.peek(ids::kBeginSXFilters)) {
    ASSIGN_OR_RETURN(doc.filters, recs.take_raw_block(ids::kEndSXFilters));
  }
  ASSIGN_OR_RETURN(doc.view14, DecodeFrt(recs, ids::kBeginSXView14, ids::kEndSXView14));
  ASSIGN_OR_RETURN(doc.view16, DecodeFrt(recs, ids::kBeginSXView16, ids::kEndSXView16));
  RETURN_IF_ERROR(recs.take_end(ids::kEndSXView, "xlsb pivot table carries an uncharacterised record"));
  if (!recs.done()) {
    return Unexpected("xlsb pivot table part continues after the view");
  }
  return doc;
}

Expected<std::vector<std::uint8_t>, Error> encode_pivot_table(const PivotTableDoc& doc) {
  const SxView& view = doc.view;
  if (static_cast<int>(view.grand_total_caption.has_value()) + static_cast<int>(view.error_caption.has_value()) +
          static_cast<int>(view.missing_caption.has_value()) >
      1) {
    return refuse_write("XLSB pivot table view sets several captions whose order is uncharacterised");
  }
  std::vector<std::uint8_t> out;
  if (doc.uid) {
    emit_record(out, ids::kACBegin, ByteSpan{kAcUidBeginPayload.data(), kAcUidBeginPayload.size()});
    emit_record(out, ids::kUid, ByteSpan{doc.uid->data(), doc.uid->size()});
    EmitEmpty(out, ids::kACEnd);
  }
  EmitView(out, view);
  EmitLocation(out, doc.location);
  EmitCount(out, ids::kBeginSXVDs, doc.fields.size());
  for (const SxField& field : doc.fields) {
    EmitField(out, field);
  }
  EmitEmpty(out, ids::kEndSXVDs);
  if (doc.row_fields) {
    EmitFieldOrder(out, ids::kBeginISXVDRws, ids::kEndISXVDRws, *doc.row_fields);
  }
  encode_line_items(out, doc.row_items, SxLineItemAxis::kRows);
  if (doc.col_fields) {
    EmitFieldOrder(out, ids::kBeginISXVDCols, ids::kEndISXVDCols, *doc.col_fields);
  }
  encode_line_items(out, doc.col_items, SxLineItemAxis::kCols);
  if (doc.page_fields) {
    EmitCount(out, ids::kBeginSXPIs, doc.page_fields->size());
    std::vector<std::uint8_t> p;
    for (const SxPageField& page : *doc.page_fields) {
      p.clear();
      emit_u32(p, page.field);
      emit_u32(p, page.item);
      emit_u32(p, page.hierarchy);
      emit_u8(p, 0);
      emit_record(out, ids::kBeginSXPI, p);
      EmitEmpty(out, ids::kEndSXPI);
    }
    EmitEmpty(out, ids::kEndSXPIs);
  }
  EmitCount(out, ids::kBeginSXDIs, doc.data_fields.size());
  std::vector<std::uint8_t> p;
  for (const SxDataField& data : doc.data_fields) {
    p.clear();
    for (const std::uint32_t v :
         {data.field, data.subtotal, data.show_data_as, data.base_field, data.base_item, data.num_fmt}) {
      emit_u32(p, v);
    }
    emit_u8(p, kDataFieldHasName);
    emit_xlwidestring(p, data.name);
    emit_record(out, ids::kBeginSXDI, p);
    EmitEmpty(out, ids::kEndSXDI);
  }
  EmitEmpty(out, ids::kEndSXDIs);
  p.clear();
  emit_u16(p, doc.style.flags);
  emit_xlnullablewidestring(p, doc.style.name ? std::optional<std::string_view>(*doc.style.name) : std::nullopt);
  emit_record(out, ids::kTableStyleClient, p);
  EmitRaw(out, doc.filters);
  EmitFrt(out, doc.view14, ids::kBeginSXView14, ids::kEndSXView14);
  EmitFrt(out, doc.view16, ids::kBeginSXView16, ids::kEndSXView16);
  EmitEmpty(out, ids::kEndSXView);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
