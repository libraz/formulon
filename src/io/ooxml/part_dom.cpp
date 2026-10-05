#include "io/ooxml/part_dom.h"

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {

const PassthroughPart* find_passthrough_part(const Workbook& wb, std::string_view path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
}

Expected<void, Error> load_part_dom(const Workbook& wb, std::string_view path, std::string_view reader_module,
                                    pugi::xml_document& doc) {
  const PassthroughPart* part = find_passthrough_part(wb, path);
  if (part == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "part_dom: no such part", "part=" + std::string(path));
  }
  return load_xml_buffer(doc, part->bytes, reader_module, path);
}

Expected<void, Error> store_part_dom(Workbook& wb, std::string_view path, const pugi::xml_document& doc) {
  std::string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  append_raw_xml(xml, doc.document_element());
  return wb.replace_passthrough_part(path, std::vector<std::uint8_t>(xml.begin(), xml.end()));
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
