// Pictures in cells: a cell holds a picture as a `Value::picture` naming an entry of the
// workbook's picture table, which interns each distinct (media bytes, alt) once.

// @size-budget: 3 KB

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/image_header.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/rich_values.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "picture_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {
namespace {

// The largest image accepted into a cell; a larger one is refused with kSecurityResourceLimit.
// @limit: engine
constexpr std::size_t kMaxPictureBytes = 32ULL * 1024ULL * 1024ULL;
// The CalcOrigin of a placed picture.
constexpr std::uint8_t kPlacedOrigin = 5;
constexpr std::string_view kModule = "cell_picture";

/// The picture table's media getter over the part store of `ctx`, a `Workbook`.
MediaSpan part_store_bytes(const void* ctx, std::string_view media_path) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(*static_cast<const Workbook*>(ctx), media_path);
  return part == nullptr || part->bytes.empty() ? MediaSpan{} : MediaSpan{part->bytes.data(), part->bytes.size()};
}

/// The media part a picture-table media already holds `bytes` in, or null.
const PassthroughPart* media_with_bytes(const Workbook& wb, const std::uint8_t* bytes, std::size_t len) {
  const PictureTable& table = wb.picture_table();
  for (const std::uint32_t media : table.media_with_key(media_key_of(bytes, len))) {
    const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, table.media_path(media));
    if (part != nullptr && part->bytes.size() == len && std::equal(part->bytes.begin(), part->bytes.end(), bytes)) {
      return part;
    }
  }
  return nullptr;
}

/// Adds the passthrough part `path` holding `xml` unless the book has it.
Expected<void, Error> ensure_part(Workbook& wb, std::string path, std::string content_type, const std::string& xml) {
  if (io::ooxml::find_passthrough_part(wb, path) != nullptr) {
    return Expected<void, Error>::Ok();
  }
  return wb.add_passthrough_part(
      PassthroughPart(std::move(path), std::move(content_type), std::vector<std::uint8_t>(xml.begin(), xml.end())));
}

/// Registers the workbook relationship to `richValueRel` a read book carries, without which
/// a reload cannot resolve any `_rvRel:` key.
void register_rich_value_rel(Workbook& wb) {
  std::vector<UnknownRelationship> rels = wb.unknown_workbook_rels();
  if (io::rich_value_rel_registered(rels)) {
    return;
  }
  rels.push_back(
      {next_unknown_rel_id(rels), std::string(io::kRelRichValueRel), std::string(io::kRichValueRelPartPath), false});
  wb.set_unknown_workbook_rels(std::move(rels));
}

/// Lists `media_path` in the `richValueRel` relationships kept as passthrough, creating the
/// pair and its workbook relationship when the book has none, so the book saves as a read one
/// does and a save drops the part once no picture value names it.
Expected<void, Error> list_media(Workbook& wb, const std::string& media_path) {
  const std::string rels_path(io::kRichValueRelRelsPartPath);
  RETURN_IF_ERROR(
      ensure_part(wb, std::string(io::kRichValueRelPartPath), std::string(io::kCtRichValueRel),
                  std::string(io::kXmlDecl) + "<richValueRels xmlns=\"" + std::string(io::kNsRichValueRel) + "\"/>"));
  RETURN_IF_ERROR(ensure_part(
      wb, rels_path, std::string(),
      std::string(io::kXmlDecl) + "<Relationships xmlns=\"" + std::string(io::kNsPackageRelationships) + "\"/>"));
  register_rich_value_rel(wb);
  pugi::xml_document doc;
  RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, rels_path, kModule, doc));
  pugi::xml_node root = doc.document_element();
  const auto taken = [&root](const std::string& id) {
    for (pugi::xml_node rel : root.children("Relationship")) {
      if (id == rel.attribute("Id").value()) {
        return true;
      }
    }
    return false;
  };
  std::size_t n = 1;
  while (taken("rId" + std::to_string(n))) {
    ++n;
  }
  pugi::xml_node rel = root.append_child("Relationship");
  rel.append_attribute("Id").set_value(("rId" + std::to_string(n)).c_str());
  rel.append_attribute("Type").set_value(std::string(io::kImageRelType).c_str());
  rel.append_attribute("Target").set_value(("../" + media_path.substr(3)).c_str());
  return io::ooxml::store_part_dom(wb, rels_path, doc);
}

}  // namespace

PicturePayload Workbook::intern_picture(std::string_view media_path, MediaSpan bytes,
                                        std::optional<std::string_view> alt, MediaBytes getter,
                                        const void* getter_ctx) {
  const MediaKey key = media_key_of(bytes.data, bytes.size);
  const auto ranked = [&](std::uint32_t media) {
    return getter == nullptr ? MediaSpan{} : getter(getter_ctx, picture_table_.media_path(media));
  };
  std::optional<std::uint32_t> media_id;
  for (const std::uint32_t media : picture_table_.media_with_key(key)) {
    const MediaSpan held = ranked(media);
    if (held.data != nullptr && compare_media_bytes(bytes.data, bytes.size, held.data, held.size) == 0) {
      media_id = media;
      break;
    }
  }
  if (!media_id.has_value()) {
    // Upper bound in ascending picture order; a ranked media with no bytes compares by its key.
    std::size_t lo = 0;
    std::size_t hi = picture_table_.media_count();
    while (lo < hi) {
      const std::size_t mid = lo + (hi - lo) / 2U;
      const std::uint32_t media = picture_table_.media_at_rank(mid);
      const MediaSpan held = ranked(media);
      const int order = held.data != nullptr ? compare_media_bytes(bytes.data, bytes.size, held.data, held.size)
                                             : compare_media_keys(key, picture_table_.media_key(media));
      if (order < 0) {
        hi = mid;
      } else {
        lo = mid + 1U;
      }
    }
    media_id = picture_table_.add_media_at(lo, key, std::string(media_path));
  }
  const std::optional<std::string_view> normal = normalise_picture_alt(alt);
  if (const std::optional<std::uint32_t> known = picture_table_.find_entry(*media_id, normal)) {
    return PicturePayload{*known, *media_id};
  }
  const std::uint32_t picture_id = picture_table_.add_entry(
      *media_id, normal.has_value() ? std::optional<std::string>(std::string(*normal)) : std::nullopt);
  return PicturePayload{picture_id, *media_id};
}

std::optional<CellPicture> Workbook::cell_picture(std::size_t sheet, std::uint32_t row, std::uint32_t col) const {
  if (sheet >= sheet_count()) {
    return std::nullopt;
  }
  const Value value = sheets_[sheet].resolve_cell_value(row, col);
  if (!value.is_picture()) {
    return std::nullopt;
  }
  const PictureEntry* entry = picture_entry(value.as_picture().picture_id);
  if (entry == nullptr) {
    return std::nullopt;
  }
  const PassthroughPart* media = io::ooxml::find_passthrough_part(*this, entry->media_path);
  if (media == nullptr) {
    return std::nullopt;
  }
  const auto info = probe_image(media->bytes.data(), media->bytes.size());
  if (!info) {
    return std::nullopt;
  }
  CellPicture out;
  out.bytes = media->bytes;
  out.format = info.value().format;
  out.width = info.value().px_width;
  out.height = info.value().px_height;
  out.alt_text = entry->alt;
  return out;
}

Expected<void, Error> Workbook::set_cell_picture(std::size_t sheet, std::uint32_t row, std::uint32_t col,
                                                 const std::uint8_t* bytes, std::size_t len,
                                                 std::optional<std::string_view> alt_text) {
  if (bytes == nullptr) {
    return make_error(FormulonErrorCode::kBindingNullPointer, "bytes must not be NULL");
  }
  if (len > kMaxPictureBytes) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "cell picture: image is larger than 32 MiB",
                      "size=" + std::to_string(len));
  }
  if (len == 0) {
    return make_error(FormulonErrorCode::kIoImageUnsupported, "cell picture: image is empty");
  }
  const auto info = probe_image(bytes, len);
  if (!info) {
    return Expected<void, Error>(info.error());
  }
  const ImageFormat format = info.value().format;
  if (format != ImageFormat::kPng && format != ImageFormat::kJpeg) {
    return make_error(FormulonErrorCode::kIoImageUnsupported, "cell picture: only PNG and JPEG images are accepted");
  }

  // Refuses before any side effect, so the call stays atomic.
  RETURN_IF_ERROR(detail::check_cell_writable(sheets_, sheet, row, col));

  std::string media_path;
  if (const PassthroughPart* shared = media_with_bytes(*this, bytes, len)) {
    media_path = shared->path;
  } else {
    const char* ext = image_extension(format);
    // A path the table names stays its media's even after the part is gone.
    std::vector<std::string> named;
    for (std::uint32_t m = 0; m < picture_table_.media_count(); ++m) {
      named.push_back(picture_table_.media_path(m));
    }
    media_path = drawing_edit::unused_part_path(*this, "xl/media/image", ext, named);
    drawing_edit::ensure_default_content_type(*this, ext, image_content_type(format));
    RETURN_IF_ERROR(add_passthrough_part(
        PassthroughPart(media_path, std::string(), std::vector<std::uint8_t>(bytes, bytes + len))));
    RETURN_IF_ERROR(list_media(*this, media_path));
  }

  const PicturePayload payload =
      intern_picture(media_path, MediaSpan{bytes, len}, alt_text, &part_store_bytes, static_cast<const void*>(this));
  put_cell_value(sheet, row, col, Value::picture(payload, kPlacedOrigin));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_cell_picture(std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  if (sheet >= sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range");
  }
  const Cell* cell = sheets_[sheet].cell_at(row, col);
  if (cell == nullptr || !cell->formula_text.empty() || !cell->cached_value.is_picture() ||
      sheets_[sheet].spill_region_covering(row, col) != nullptr) {
    return Expected<void, Error>::Ok();
  }
  return set_cell_value(sheet, row, col, Value::blank());
}

}  // namespace formulon
