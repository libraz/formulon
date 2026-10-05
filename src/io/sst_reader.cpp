//
// Implementation of the shared-strings reader. See sst_reader.h for the
// public contract.
//
// The walker is intentionally conservative: it concatenates every `<t>`
// descendant of a given `<si>` (whether direct or under `<r>`) into the
// surface text and reports `kIoSheetCorrupt` when a `<si>` carries
// neither a direct `<t>` nor any `<r><t>` payload. Rich-text formatting
// attributes on `<r>`/`<rPr>` are not preserved (this layer is plain-
// text only, by design). Phonetic-guide subtrees (`<rPh>`) are walked
// separately and land in `phonetic_for_entries[i]` as one run per block,
// span offsets included, so PHONETIC() can surface the kana over the
// characters it actually covers.

#include "io/sst_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "io/xml_utils.h"
#include "phonetic.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {

Expected<SharedStringTable, Error> read_shared_strings(std::vector<std::uint8_t> sst_bytes,
                                                       std::deque<std::string>& text_storage) {
  // `doc` is a body local and `sst_bytes` a parameter, so the buffer the
  // DOM aliases is destroyed after the DOM that points into it.
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer_inplace(doc, sst_bytes, "sst_reader", "sharedStrings.xml"));
  pugi::xml_node root = doc.child("sst");
  if (!root) {
    return make_error(FormulonErrorCode::kIoXmlParse, "sharedStrings.xml: missing <sst> root",
                      "context=sst_reader part=xl/sharedStrings.xml");
  }

  SharedStringTable table;

  std::size_t index = 0;
  for (pugi::xml_node si = root.child("si"); si; si = si.next_sibling("si"), ++index) {
    text_storage.emplace_back();
    std::string& payload = text_storage.back();
    const std::size_t t_count = append_rich_text(si, payload);
    if (t_count == 0) {
      // No <t> descendant at all. Roll the placeholder back so
      // text_storage stays in sync with the (failing) result and report
      // the offending index in context. We deliberately error out rather
      // than silently storing "" because Excel never emits a <t>-less
      // <si>; a SST without any text payload is data loss waiting to
      // happen on the next write.
      text_storage.pop_back();
      std::string ctx("context=sst_reader part=xl/sharedStrings.xml index=");
      ctx.append(std::to_string(index));
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "sharedStrings.xml: <si> with no <t> descendant",
                        std::move(ctx));
    }
    table.entries.emplace_back(payload);

    // Capture phonetic runs from any <rPh> children. These own their
    // kana rather than aliasing `text_storage`, so an unannotated entry
    // costs an empty vector and no allocation. The slot is pushed
    // unconditionally to keep the index alignment invariant
    // `phonetic_for_entries.size() == entries.size()`.
    table.phonetic_for_entries.emplace_back();
    append_phonetic_runs(si, table.phonetic_for_entries.back());
    table.phonetic_props_for_entries.push_back(read_phonetic_properties(si));
  }

  return table;
}

}  // namespace io
}  // namespace formulon
