// Pictures in cells: a picture is a rich value (`_localImage`) on a cell that caches `#VALUE!`.

// @size-budget: 3 KB

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/drawing_edit_internal.h"
#include "drawing/image_header.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/rich_values.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// The largest image accepted into a cell; a larger one is refused with kSecurityResourceLimit.
// @limit: engine
constexpr std::size_t kMaxPictureBytes = 32ULL * 1024ULL * 1024ULL;
constexpr std::string_view kPictureStructure = "_localImage";
constexpr std::string_view kImageKey = "_rvRel:LocalImageIdentifier";
constexpr std::string_view kOriginKey = "CalcOrigin";
constexpr std::string_view kTextKey = "Text";
constexpr std::string_view kPlacedOrigin = "5";
constexpr std::string_view kImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
constexpr std::string_view kModule = "cell_picture";

const OpaqueRichValueKey* find_key(const OpaqueRichValue& value, std::string_view name) {
  for (const OpaqueRichValueKey& key : value.keys) {
    if (key.name == name) {
      return &key;
    }
  }
  return nullptr;
}

/// The media part a picture value points at, or null.
const PassthroughPart* picture_media(const Workbook& wb, const OpaqueRichValue& value) {
  if (value.structure_type != kPictureStructure) {
    return nullptr;
  }
  const OpaqueRichValueKey* key = find_key(value, kImageKey);
  return key == nullptr ? nullptr : io::ooxml::find_passthrough_part(wb, key->rel_target);
}

/// The media part a picture value in the book already holds `bytes` in, or null.
const PassthroughPart* media_with_bytes(const Workbook& wb, const std::uint8_t* bytes, std::size_t len) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const CellRichValue& cell : wb.sheet(s).rich_values()) {
      const PassthroughPart* part = picture_media(wb, cell.value);
      if (part != nullptr && part->bytes.size() == len && std::equal(part->bytes.begin(), part->bytes.end(), bytes)) {
        return part;
      }
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
  std::uint32_t top = 0;
  for (const UnknownRelationship& rel : rels) {
    if (rel.id.size() > 3U && rel.id.compare(0, 3, "rId") == 0) {
      top = std::max(top, static_cast<std::uint32_t>(std::strtoul(rel.id.c_str() + 3, nullptr, 10)));
    }
  }
  rels.push_back({"rId" + std::to_string(top + 1U), std::string(io::kRelRichValueRel),
                  std::string(io::kRichValueRelPartPath), false});
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
  rel.append_attribute("Type").set_value(std::string(kImageRelType).c_str());
  rel.append_attribute("Target").set_value(("../" + media_path.substr(3)).c_str());
  return io::ooxml::store_part_dom(wb, rels_path, doc);
}

OpaqueRichValueKey make_key(std::string_view name, std::string_view type, std::string_view value) {
  OpaqueRichValueKey key;
  key.name = std::string(name);
  key.type = std::string(type);
  key.value = std::string(value);
  return key;
}

}  // namespace

std::optional<CellPicture> Workbook::cell_picture(std::size_t sheet, std::uint32_t row, std::uint32_t col) const {
  if (sheet >= sheet_count()) {
    return std::nullopt;
  }
  const std::optional<OpaqueRichValue> value = sheets_[sheet].rich_value_at(row, col);
  if (!value.has_value()) {
    return std::nullopt;
  }
  const PassthroughPart* media = picture_media(*this, *value);
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
  if (const OpaqueRichValueKey* text = find_key(*value, kTextKey)) {
    out.alt_text = text->value;
  }
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
  if (alt_text.has_value() && alt_text->empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "alt_text must not be empty");
  }

  // The write comes first and refuses without changing anything; it also drops
  // the cell's earlier rich value, so the picture is recorded after it.
  RETURN_IF_ERROR(set_cell_value(sheet, row, col, Value::error(ErrorCode::Value)));

  std::string media_path;
  if (const PassthroughPart* shared = media_with_bytes(*this, bytes, len)) {
    media_path = shared->path;
  } else {
    const char* ext = image_extension(format);
    media_path = drawing_edit::unused_part_path(*this, "xl/media/image", ext);
    drawing_edit::ensure_default_content_type(*this, ext, image_content_type(format));
    RETURN_IF_ERROR(add_passthrough_part(
        PassthroughPart(media_path, std::string(), std::vector<std::uint8_t>(bytes, bytes + len))));
    RETURN_IF_ERROR(list_media(*this, media_path));
  }

  OpaqueRichValue value;
  value.structure_type = std::string(kPictureStructure);
  OpaqueRichValueKey image = make_key(kImageKey, "i", "");
  image.rel_target = media_path;
  image.rel_type = std::string(kImageRelType);
  value.keys.push_back(std::move(image));
  value.keys.push_back(make_key(kOriginKey, "i", kPlacedOrigin));
  if (alt_text.has_value()) {
    value.keys.push_back(make_key(kTextKey, "s", *alt_text));
  }
  sheets_[sheet].set_rich_value(row, col, std::move(value));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_cell_picture(std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  if (sheet >= sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range");
  }
  const std::optional<OpaqueRichValue> value = sheets_[sheet].rich_value_at(row, col);
  if (!value.has_value() || value->structure_type != kPictureStructure) {
    return Expected<void, Error>::Ok();
  }
  RETURN_IF_ERROR(set_cell_value(sheet, row, col, Value::blank()));
  sheets_[sheet].clear_rich_value(row, col);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
