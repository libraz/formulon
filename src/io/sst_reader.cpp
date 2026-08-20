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

#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "phonetic.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace {

/// Collects every `<rPh>` direct child of `si_node` into `out`, one run
/// per block, in document order. Each block looks like
/// `<rPh sb="0" eb="2"><t>トウキョウ</t></rPh>`: `sb`/`eb` delimit the
/// surface-text span the kana covers, in UTF-16 code units, and the
/// block's `<t>` descendants concatenate into that run's kana. The spans
/// are kept because PHONETIC leaves the unannotated remainder of the
/// string in place, so collapsing multi-block annotations into one kana
/// string would lose observable content.
void CollectPhoneticRuns(const pugi::xml_node& si_node, std::vector<PhoneticRun>& out) {
  for (pugi::xml_node rph = si_node.child("rPh"); rph; rph = rph.next_sibling("rPh")) {
    PhoneticRun run;
    run.sb = static_cast<std::uint32_t>(rph.attribute("sb").as_uint(0U));
    run.eb = static_cast<std::uint32_t>(rph.attribute("eb").as_uint(0U));
    for (pugi::xml_node t = rph.child("t"); t; t = t.next_sibling("t")) {
      AppendOoxmlTextUnescaped(run.text, t.text().get());
    }
    out.push_back(std::move(run));
  }
}

}  // namespace

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
    CollectPhoneticRuns(si, table.phonetic_for_entries.back());
  }

  return table;
}

}  // namespace io
}  // namespace formulon
