#include "workbook_theme_package.h"

// @size-budget: 5 KB

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "drawing/drawing_edit_internal.h"
#include "drawing/image_header.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/theme_part.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon::io {
namespace {

constexpr std::size_t kPackageBudget = 64ULL * 1024ULL * 1024ULL;
constexpr const char* kManagerContentType = "application/vnd.openxmlformats-officedocument.themeManager+xml";
constexpr const char* kPackageThemeDir = "theme/theme/";
constexpr const char* kPackageThemePath = "theme/theme/theme1.xml";
constexpr const char* kPackageManagerPath = "theme/theme/themeManager.xml";
constexpr const char* kManagerXml =
    "<a:themeManager xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";

using Bytes = std::vector<std::uint8_t>;

/// A theme relationship with its payload. `media_path` is the destination
/// path of an internal target; `bytes` is empty for a repeat of an earlier target.
struct CarriedRel {
  ooxml::DrawingRel rel;
  std::string media_path;
  Bytes bytes;
  std::string ext;
  std::string content_type;
};

std::string content_type_for_ext(const Workbook& wb, const std::string& ext) {
  for (const DefaultContentType& def : wb.default_content_types()) {
    if (def.extension == ext) {
      return def.content_type;
    }
  }
  if (ext == "png" || ext == "gif" || ext == "bmp") {
    return "image/" + ext;
  }
  if (ext == "jpeg" || ext == "jpg") {
    return "image/jpeg";
  }
  if (ext == "emf") {
    return "image/x-emf";
  }
  if (ext == "wdp") {
    return "image/vnd.ms-photo";
  }
  return "application/octet-stream";
}

/// `to` relative to directory `from_dir` (both package paths, no leading slash).
std::string relative_path(std::string_view from_dir, std::string_view to) {
  std::vector<std::string_view> base;
  std::vector<std::string_view> dest;
  const auto split = [](std::string_view path, std::vector<std::string_view>& out) {
    while (!path.empty()) {
      const std::size_t slash = path.find('/');
      out.push_back(path.substr(0, slash));
      path = slash == std::string_view::npos ? std::string_view() : path.substr(slash + 1);
    }
  };
  split(from_dir, base);
  split(to, dest);
  std::size_t common = 0;
  while (common < base.size() && common + 1 < dest.size() && base[common] == dest[common]) {
    ++common;
  }
  std::string out;
  for (std::size_t i = common; i < base.size(); ++i) {
    out += "../";
  }
  for (std::size_t i = common; i < dest.size(); ++i) {
    out.append(dest[i]);
    if (i + 1 < dest.size()) {
      out += '/';
    }
  }
  return out;
}

std::string rels_xml(const std::vector<ooxml::DrawingRel>& rels) {
  std::string xml(kXmlDecl);
  xml.append("<Relationships xmlns=\"").append(kNsPackageRelationships).append("\">");
  for (const ooxml::DrawingRel& rel : rels) {
    xml += "<Relationship Id=\"";
    AppendXmlAttrEscaped(xml, rel.id);
    xml += "\" Type=\"";
    AppendXmlAttrEscaped(xml, rel.type);
    xml += "\" Target=\"";
    AppendXmlAttrEscaped(xml, rel.target);
    xml += rel.external ? "\" TargetMode=\"External\"/>" : "\"/>";
  }
  xml += "</Relationships>";
  return xml;
}

Expected<Bytes, Error> read_part(const ZipReader& zip, const std::string& name) {
  if (!zip.has_entry(name)) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "theme package: part is missing", "part=" + name);
  }
  return zip.read_entry(name);
}

Expected<std::vector<ooxml::DrawingRel>, Error> read_rels(const Bytes& bytes, std::string_view owner_dir) {
  auto rels = ooxml::parse_part_rels(bytes, owner_dir);
  if (!rels) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "theme package: relationships are unreadable",
                      "dir=" + std::string(owner_dir));
  }
  return std::move(rels.value());
}

Expected<std::string, Error> internal_target(const std::vector<ooxml::DrawingRel>& rels, std::string_view type) {
  for (const ooxml::DrawingRel& rel : rels) {
    if (!rel.external && rel.type == type) {
      return rel.target;
    }
  }
  return make_error(FormulonErrorCode::kIoRelationshipBroken, "theme package: relationship is missing",
                    "type=" + std::string(type));
}

void ensure_default_content_type(Workbook& wb, const std::string& extension, const std::string& content_type) {
  for (const DefaultContentType& def : wb.default_content_types()) {
    if (def.extension == extension) {
      return;
    }
  }
  std::vector<DefaultContentType> defaults = wb.default_content_types();
  defaults.push_back({extension, content_type});
  wb.set_default_content_types(std::move(defaults));
}

/// Reads the theme part's own relationships out of the package, loading each
/// internal target and choosing its destination path in `wb`.
Expected<std::vector<CarriedRel>, Error> read_theme_rels(const Workbook& wb, const ZipReader& zip,
                                                         const std::string& theme_src) {
  std::vector<CarriedRel> carried;
  const std::string rels_src = ooxml::rels_path_for_part(theme_src);
  if (!zip.has_entry(rels_src)) {
    return carried;
  }
  ASSIGN_OR_RETURN(const Bytes rels_bytes, read_part(zip, rels_src));
  ASSIGN_OR_RETURN(std::vector<ooxml::DrawingRel> rels, read_rels(rels_bytes, ooxml::dir_of(theme_src)));
  std::unordered_map<std::string, std::size_t> by_target;
  std::vector<std::string> reserved;
  for (ooxml::DrawingRel& rel : rels) {
    CarriedRel out;
    if (!rel.external) {
      const auto seen = by_target.find(rel.target);
      if (seen != by_target.end()) {
        out.media_path = carried[seen->second].media_path;
        out.rel = std::move(rel);
        carried.push_back(std::move(out));
        continue;
      }
      ASSIGN_OR_RETURN(out.bytes, read_part(zip, rel.target));
      auto info = probe_image(out.bytes.data(), out.bytes.size());
      if (info) {
        out.ext = image_extension(info.value().format);
        out.content_type = image_content_type(info.value().format);
      } else {
        out.ext = ooxml::extension_of_part(rel.target);
        if (out.ext != "emf" && out.ext != "wdp") {
          return Expected<std::vector<CarriedRel>, Error>(std::move(info.error()));
        }
        out.content_type = content_type_for_ext(wb, out.ext);
      }
      out.media_path = drawing_edit::unused_part_path(wb, "xl/media/image", out.ext.c_str(), reserved);
      reserved.push_back(out.media_path);
      by_target.emplace(rel.target, carried.size());
    }
    out.rel = std::move(rel);
    carried.push_back(std::move(out));
  }
  return carried;
}

/// Excel's apply removes `defaultThemeVersion` from `<workbookPr>`. An empty
/// fragment becomes `<workbookPr/>` so the xlsb writer does not default it back.
void drop_default_theme_version(Workbook& wb) {
  const std::string& raw = wb.workbook_pr_xml();
  pugi::xml_document doc;
  const std::string_view source = raw.empty() ? std::string_view("<workbookPr/>") : std::string_view(raw);
  if (!doc.load_buffer(source.data(), source.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return;
  }
  pugi::xml_node root = doc.first_child();
  if (!root || std::string_view(root.name()) != "workbookPr") {
    return;
  }
  root.remove_attribute("defaultThemeVersion");
  wb.set_workbook_pr_xml(raw_xml(root));
}

}  // namespace

Expected<void, Error> apply_theme_package(Workbook& wb, const std::uint8_t* data, std::size_t size) {
  if (data == nullptr) {
    return make_error(FormulonErrorCode::kBindingNullPointer, "theme package: data is NULL");
  }
  if (size == 0) {
    return make_error(FormulonErrorCode::kInvalidArgument, "theme package: data is empty");
  }
  ZipReader zip;
  RETURN_IF_ERROR(zip.open(ByteSpan{data, size}, kPackageBudget));

  ASSIGN_OR_RETURN(const Bytes root_rels_bytes, read_part(zip, "_rels/.rels"));
  ASSIGN_OR_RETURN(std::vector<ooxml::DrawingRel> root_rels, read_rels(root_rels_bytes, ""));
  ASSIGN_OR_RETURN(const std::string manager, internal_target(root_rels, kRelOfficeDocument));
  ASSIGN_OR_RETURN(const Bytes manager_rels_bytes, read_part(zip, ooxml::rels_path_for_part(manager)));
  ASSIGN_OR_RETURN(std::vector<ooxml::DrawingRel> manager_rels, read_rels(manager_rels_bytes, ooxml::dir_of(manager)));
  ASSIGN_OR_RETURN(const std::string theme_src, internal_target(manager_rels, kRelTheme));
  ASSIGN_OR_RETURN(Bytes theme_bytes, read_part(zip, theme_src));
  if (!is_parseable_theme(theme_bytes)) {
    return make_error(FormulonErrorCode::kIoXmlParse, "theme package: theme is not a parseable theme",
                      "part=" + theme_src);
  }
  ASSIGN_OR_RETURN(std::vector<CarriedRel> carried, read_theme_rels(wb, zip, theme_src));

  // Everything is validated; from here the workbook changes.
  const std::string path = theme_part_path(wb);
  const std::string rels_path = ooxml::rels_path_for_part(path);
  const Theme before = load_theme(wb).theme;
  const PassthroughPart* old_rels = ooxml::find_passthrough_part(wb, rels_path);
  std::vector<std::string> orphans;
  if (old_rels != nullptr) {
    if (auto old = ooxml::parse_part_rels(old_rels->bytes, ooxml::dir_of(path))) {
      for (const ooxml::DrawingRel& rel : old.value()) {
        if (!rel.external) {
          orphans.push_back(rel.target);
        }
      }
    }
  }

  std::vector<ooxml::DrawingRel> new_rels;
  for (CarriedRel& item : carried) {
    if (!item.bytes.empty()) {
      RETURN_IF_ERROR(wb.add_passthrough_part(PassthroughPart(item.media_path, "", std::move(item.bytes))));
      ensure_default_content_type(wb, item.ext, item.content_type);
    }
    ooxml::DrawingRel rel = std::move(item.rel);
    if (!rel.external) {
      rel.target = relative_path(ooxml::dir_of(path), item.media_path);
    }
    new_rels.push_back(std::move(rel));
  }
  if (!new_rels.empty()) {
    const std::string xml = rels_xml(new_rels);
    Bytes bytes(xml.begin(), xml.end());
    RETURN_IF_ERROR(old_rels != nullptr ? wb.replace_passthrough_part(rels_path, std::move(bytes))
                                        : wb.add_passthrough_part(PassthroughPart(rels_path, "", std::move(bytes))));
  } else if (old_rels != nullptr) {
    orphans.push_back(rels_path);
  }
  if (ooxml::find_passthrough_part(wb, path) != nullptr) {
    RETURN_IF_ERROR(wb.replace_passthrough_part(path, std::move(theme_bytes)));
  } else {
    RETURN_IF_ERROR(wb.add_passthrough_part(PassthroughPart(path, kThemeContentType, std::move(theme_bytes))));
    wb.add_workbook_relationship(std::string(kRelTheme), path);
  }
  ooxml::remove_orphans(wb, std::move(orphans));
  mark_formulas_dirty_if_fonts_changed(wb, before);
  drop_default_theme_version(wb);
  return Expected<void, Error>();
}

Expected<std::vector<std::uint8_t>, Error> save_theme_package(const Workbook& wb) {
  const std::string path = theme_part_path(wb);
  const PassthroughPart* part = ooxml::find_passthrough_part(wb, path);
  Bytes theme_bytes;
  if (part != nullptr) {
    theme_bytes = part->bytes;
    if (!is_parseable_theme(theme_bytes)) {
      return make_error(FormulonErrorCode::kIoXmlParse, "theme: part is not a parseable theme", "part=" + path);
    }
  } else {
    const std::string xml = build_default_theme_xml();
    theme_bytes.assign(xml.begin(), xml.end());
  }

  // Media keep the theme's rel ids; each distinct target is renamed under theme/media/.
  std::vector<ooxml::DrawingRel> rels;
  std::vector<std::pair<std::string, const Bytes*>> media;
  std::vector<std::string> exts;
  const PassthroughPart* rels_part =
      part != nullptr ? ooxml::find_passthrough_part(wb, ooxml::rels_path_for_part(path)) : nullptr;
  if (rels_part != nullptr) {
    ASSIGN_OR_RETURN(rels, read_rels(rels_part->bytes, ooxml::dir_of(path)));
    std::unordered_map<std::string, std::string> renamed;
    for (ooxml::DrawingRel& rel : rels) {
      if (rel.external) {
        continue;
      }
      auto seen = renamed.find(rel.target);
      if (seen == renamed.end()) {
        const PassthroughPart* source = ooxml::find_passthrough_part(wb, rel.target);
        if (source == nullptr) {
          return make_error(FormulonErrorCode::kIoRelationshipBroken, "theme: media part is missing",
                            "part=" + rel.target);
        }
        const std::string ext = ooxml::extension_of_part(rel.target);
        const std::string name = "image" + std::to_string(media.size() + 1U) + "." + ext;
        media.emplace_back("theme/media/" + name, &source->bytes);
        if (std::find(exts.begin(), exts.end(), ext) == exts.end()) {
          exts.push_back(ext);
        }
        seen = renamed.emplace(rel.target, "../media/" + name).first;
      }
      rel.target = seen->second;
    }
  }

  std::string types(kXmlDecl);
  types.append("<Types xmlns=\"").append(kNsContentTypes).append("\">");
  types.append("<Default Extension=\"rels\" ContentType=\"").append(kCtPackageRels).append("\"/>");
  types.append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
  for (const std::string& ext : exts) {
    types.append("<Default Extension=\"").append(ext).append("\" ContentType=\"");
    AppendXmlAttrEscaped(types, content_type_for_ext(wb, ext));
    types.append("\"/>");
  }
  types.append("<Override PartName=\"/").append(kPackageThemePath).append("\" ContentType=\"");
  types.append(kThemeContentType).append("\"/>");
  types.append("<Override PartName=\"/").append(kPackageManagerPath).append("\" ContentType=\"");
  types.append(kManagerContentType).append("\"/>");
  types.append("</Types>");

  const auto rel_to = [](const char* type, const char* target) {
    return std::vector<ooxml::DrawingRel>{{"rId1", type, target, false}};
  };
  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed",
                      "context=save_theme_package");
  }
  std::unordered_set<std::string> written;
  mz_zip_archive* zip = writer.get();
  RETURN_IF_ERROR(AddPart(zip, "[Content_Types].xml", types, &written));
  RETURN_IF_ERROR(AddPart(zip, "_rels/.rels",
                          rels_xml(rel_to(std::string(kRelOfficeDocument).c_str(), kPackageManagerPath)), &written));
  RETURN_IF_ERROR(AddPart(zip, kPackageManagerPath, std::string(kXmlDecl) + kManagerXml, &written));
  RETURN_IF_ERROR(AddPart(zip, std::string(kPackageThemeDir) + "_rels/themeManager.xml.rels",
                          rels_xml(rel_to(std::string(kRelTheme).c_str(), "theme1.xml")), &written));
  RETURN_IF_ERROR(AddPartBytes(zip, kPackageThemePath, theme_bytes, &written));
  if (!rels.empty()) {
    RETURN_IF_ERROR(AddPart(zip, std::string(kPackageThemeDir) + "_rels/theme1.xml.rels", rels_xml(rels), &written));
  }
  for (const auto& [name, bytes] : media) {
    RETURN_IF_ERROR(AddPartBytes(zip, name, *bytes, &written));
  }
  return FinalizeArchive(writer, "save_theme_package");
}

}  // namespace formulon::io
