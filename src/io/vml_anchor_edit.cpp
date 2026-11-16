#include "io/vml_anchor_edit.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cstddef>
#include <string_view>
#include <utility>

#include "io/worksheet_child_refs.h"
#include "utils/number_text.h"

namespace formulon::io {
namespace {

constexpr std::string_view kShapeOpen = "<v:shape ";
constexpr std::string_view kShapeClose = "</v:shape>";
constexpr std::string_view kAnchorOpen = "<x:Anchor>";
constexpr std::string_view kAnchorClose = "</x:Anchor>";

/// Points per VML anchor offset pixel (96 DPI).
constexpr double kPointsPerPixel = 0.75;

/// `x:Anchor` fields: left column, its offset, top row, its offset, right
/// column, its offset, bottom row, its offset.
using AnchorFields = std::array<std::int64_t, 8>;

bool ParseAnchor(std::string_view text, AnchorFields& fields) {
  std::size_t at = 0;
  for (std::size_t i = 0; i < fields.size(); ++i) {
    while (at < text.size() && (text[at] == ' ' || text[at] == '\n' || text[at] == '\r' || text[at] == '\t')) {
      ++at;
    }
    const auto [end, ec] = std::from_chars(text.data() + at, text.data() + text.size(), fields[i]);
    if (ec != std::errc() || fields[i] < 0) {
      return false;
    }
    at = static_cast<std::size_t>(end - text.data());
    if (i + 1U < fields.size()) {
      if (at >= text.size() || text[at] != ',') {
        return false;
      }
      ++at;
    }
  }
  return true;
}

std::string FormatAnchor(const AnchorFields& fields) {
  std::string out;
  for (std::size_t i = 0; i < fields.size(); ++i) {
    out.append(i == 0U ? "" : ", ").append(std::to_string(fields[i]));
  }
  return out;
}

// Replaces the value of CSS property `name` in `style` with `points`pt.
void SetStyleLength(std::string& style, std::string_view name, double points) {
  for (std::size_t at = style.find(name); at != std::string::npos; at = style.find(name, at + 1U)) {
    const bool starts = at == 0U || style[at - 1U] == ';' || style[at - 1U] == ' ' || style[at - 1U] == '\n';
    if (!starts || at + name.size() >= style.size() || style[at + name.size()] != ':') {
      continue;
    }
    const std::size_t value = at + name.size() + 1U;
    const std::size_t end = style.find(';', value);
    std::string text;
    append_xml_number(text, points);
    text.append("pt");
    style.replace(value, (end == std::string::npos ? style.size() : end) - value, text);
    return;
  }
}

// Re-derives `margin-left`, `margin-top`, `width` and `height` in the shape
// open tag `tag` from `fields`.
void SetStyleGeometry(std::string& tag, const AnchorFields& fields, const LineOffsetPt& offset_pt) {
  std::size_t at = tag.find(" style=");
  if (at == std::string::npos || at + 8U >= tag.size()) {
    return;
  }
  const char quote = tag[at + 7U];
  const std::size_t begin = at + 8U;
  const std::size_t end = tag.find(quote, begin);
  if (end == std::string::npos) {
    return;
  }
  const auto position = [&](bool row_axis, std::size_t field) {
    return offset_pt(row_axis, static_cast<std::uint32_t>(fields[field])) +
           static_cast<double>(fields[field + 1U]) * kPointsPerPixel;
  };
  const double left = position(false, 0U);
  const double top = position(true, 2U);
  std::string style = tag.substr(begin, end - begin);
  SetStyleLength(style, "margin-left", left);
  SetStyleLength(style, "margin-top", top);
  SetStyleLength(style, "width", position(false, 4U) - left);
  SetStyleLength(style, "height", position(true, 6U) - top);
  tag.replace(begin, end - begin, style);
}

// Moves one shape's anchor; returns false when the shape is left as it was.
bool RemapShape(std::string& shape, const StructuralEdit& edit, const LineOffsetPt& offset_pt) {
  const std::size_t data = shape.find("<x:ClientData");
  if (data == std::string::npos || shape.find("ObjectType=\"Note\"", data) != std::string::npos ||
      shape.find("<x:MoveWithCells", data) != std::string::npos) {
    return false;
  }
  const std::size_t open = shape.find(kAnchorOpen, data);
  const std::size_t text_at = open == std::string::npos ? open : open + kAnchorOpen.size();
  const std::size_t close = open == std::string::npos ? open : shape.find(kAnchorClose, text_at);
  AnchorFields fields{};
  if (close == std::string::npos || !ParseAnchor(std::string_view(shape).substr(text_at, close - text_at), fields)) {
    return false;
  }
  const std::size_t axis = edit.row_axis ? 2U : 0U;
  AnchorLine from{static_cast<std::uint32_t>(fields[axis]), fields[axis + 1U]};
  AnchorLine to{static_cast<std::uint32_t>(fields[axis + 4U]), fields[axis + 5U]};
  move_cell_anchor(from, to, shape.find("<x:SizeWithCells", data) == std::string::npos, edit);
  const AnchorFields before = fields;
  fields[axis] = from.line;
  fields[axis + 1U] = from.offset;
  fields[axis + 4U] = to.line;
  fields[axis + 5U] = to.offset;
  if (fields == before) {
    return false;
  }
  // Keep the whitespace Excel writes before the numbers.
  std::size_t first = text_at;
  while (first < close &&
         (shape[first] == ' ' || shape[first] == '\n' || shape[first] == '\r' || shape[first] == '\t')) {
    ++first;
  }
  shape.replace(first, close - first, FormatAnchor(fields));
  const std::size_t tag_end = shape.find('>');
  std::string tag = shape.substr(0, tag_end);
  SetStyleGeometry(tag, fields, offset_pt);
  shape.replace(0, tag_end, tag);
  return true;
}

}  // namespace

std::vector<std::string> remove_vml_shapes(std::string& shapes, const std::vector<std::uint32_t>& shape_ids) {
  constexpr std::string_view kIdPrefix = "id=\"_x0000_s";
  constexpr std::string_view kRelId = "o:relid=\"";
  std::vector<std::string> image_rel_ids;
  std::string out;
  std::size_t copied = 0;
  for (std::size_t at = shapes.find(kShapeOpen); at != std::string::npos; at = shapes.find(kShapeOpen, at)) {
    const std::size_t close = shapes.find(kShapeClose, at);
    if (close == std::string::npos) {
      break;
    }
    const std::size_t end = close + kShapeClose.size();
    const std::string_view shape = std::string_view(shapes).substr(at, end - at);
    const std::size_t id_at = shape.substr(0, shape.find('>')).find(kIdPrefix);
    std::uint32_t id = 0;
    const char* digits = id_at == std::string_view::npos ? nullptr : shape.data() + id_at + kIdPrefix.size();
    if (digits == nullptr || std::from_chars(digits, shape.data() + shape.size(), id).ec != std::errc() ||
        std::find(shape_ids.begin(), shape_ids.end(), id) == shape_ids.end()) {
      at = end;
      continue;
    }
    if (const std::size_t rel = shape.find(kRelId); rel != std::string_view::npos) {
      const std::size_t value = rel + kRelId.size();
      image_rel_ids.emplace_back(shape.substr(value, shape.find('"', value) - value));
    }
    out.append(shapes, copied, at - copied);
    // The line break Excel writes after the shape goes with it.
    copied = end;
    while (copied < shapes.size() && (shapes[copied] == '\r' || shapes[copied] == '\n')) {
      ++copied;
    }
    at = copied;
  }
  if (copied != 0U) {
    out.append(shapes, copied, std::string::npos);
    shapes = std::move(out);
  }
  return image_rel_ids;
}

bool vml_has_shapes(std::string_view shapes) {
  return shapes.find(kShapeOpen) != std::string_view::npos;
}

bool remap_vml_anchors(std::string& shapes, const StructuralEdit& edit, const LineOffsetPt& offset_pt) {
  if (edit.band.has_value() || edit.count == 0U) {
    return false;
  }
  bool changed = false;
  std::string out;
  out.reserve(shapes.size());
  std::size_t copied = 0;
  for (std::size_t at = shapes.find(kShapeOpen); at != std::string::npos; at = shapes.find(kShapeOpen, at)) {
    const std::size_t close = shapes.find(kShapeClose, at);
    if (close == std::string::npos) {
      break;
    }
    const std::size_t end = close + kShapeClose.size();
    std::string shape = shapes.substr(at, end - at);
    if (RemapShape(shape, edit, offset_pt)) {
      out.append(shapes, copied, at - copied).append(shape);
      copied = end;
      changed = true;
    }
    at = end;
  }
  if (changed) {
    out.append(shapes, copied, std::string::npos);
    shapes = std::move(out);
  }
  return changed;
}

}  // namespace formulon::io
