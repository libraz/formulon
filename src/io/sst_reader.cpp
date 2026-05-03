// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
// separately and their concatenated `<t>` payloads land in
// `phonetic_for_entries[i]` so PHONETIC() can surface the kana.

#include "io/sst_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

/// Appends every `<t>` text descendant of `node` to `out`, walking into
/// `<r>` (rich-text run) children but skipping `<rPh>` (phonetic guide)
/// subtrees entirely. Returns the number of `<t>` elements consumed so
/// callers can flag empty `<si>` entries as corrupt.
std::size_t AppendSiText(const pugi::xml_node& node, std::string& out) {
  std::size_t count = 0;

  // Direct <t> child: a plain (non-rich) string item.
  for (pugi::xml_node t = node.child("t"); t; t = t.next_sibling("t")) {
    out.append(t.text().get());
    ++count;
  }

  // Rich-text runs: each <r> may contain its own <rPr> (formatting,
  // ignored at this layer) and a single <t>. Concatenate <t> payloads in
  // document order across runs.
  for (pugi::xml_node r = node.child("r"); r; r = r.next_sibling("r")) {
    for (pugi::xml_node t = r.child("t"); t; t = t.next_sibling("t")) {
      out.append(t.text().get());
      ++count;
    }
  }

  // <rPh> (phonetic guides) are walked by `AppendPhoneticText`; this
  // helper deliberately skips them so the surface text in `entries[i]`
  // never picks up kana annotations.

  return count;
}

/// Walks every `<rPh>` direct child of `si_node` and concatenates their
/// `<t>` descendants into `out` in document order. Each `<rPh>` block
/// looks like `<rPh sb="0" eb="2"><t>やまだ</t></rPh>`; multi-block
/// annotations on a single `<si>` (one per kanji span) are flattened
/// into a single kana string. This is lossy with respect to the sb/eb
/// span boundaries, which is acceptable: PHONETIC()'s observable result
/// is the concatenated kana, so the spans are unobservable through the
/// engine's surface today.
void AppendPhoneticText(const pugi::xml_node& si_node, std::string& out) {
  for (pugi::xml_node rph = si_node.child("rPh"); rph; rph = rph.next_sibling("rPh")) {
    for (pugi::xml_node t = rph.child("t"); t; t = t.next_sibling("t")) {
      out.append(t.text().get());
    }
  }
}

}  // namespace

Expected<SharedStringTable, Error> read_shared_strings(const std::vector<std::uint8_t>& sst_bytes,
                                                       std::deque<std::string>& text_storage) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(sst_bytes.data(), sst_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=sst_reader part=xl/sharedStrings.xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "sharedStrings.xml: pugixml parse failed", std::move(ctx));
  }
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
    const std::size_t t_count = AppendSiText(si, payload);
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

    // Capture phonetic kana from any <rPh> children. We append into
    // `text_storage` (which is the same pointer-stable deque used for
    // `entries`) so the resulting `string_view` shares the same
    // lifetime. When the entry has no <rPh> we skip the append entirely
    // and store an empty `string_view{}` to keep the index alignment
    // invariant `phonetic_for_entries.size() == entries.size()`.
    text_storage.emplace_back();
    std::string& phonetic_payload = text_storage.back();
    AppendPhoneticText(si, phonetic_payload);
    if (phonetic_payload.empty()) {
      text_storage.pop_back();
      table.phonetic_for_entries.emplace_back();
    } else {
      table.phonetic_for_entries.emplace_back(phonetic_payload);
    }
  }

  return table;
}

}  // namespace io
}  // namespace formulon
