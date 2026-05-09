// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the comments-part reader. See comments_reader.h
// for the contract; the format is a small fragment of OOXML so the
// implementation is intentionally short and self-contained.

#include "io/comments_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon::io {

Expected<std::vector<CellComment>, Error> read_comments(const std::vector<std::uint8_t>& bytes) {
  std::vector<CellComment> out;
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, "comments_reader", "comments part"));
  pugi::xml_node root = doc.child("comments");
  if (!root) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "comments part: missing <comments> root",
                      "context=comments_reader");
  }

  // Author table: `<authors><author>Alice</author>...</authors>`. The
  // table is referenced by `<comment authorId="N">` indices.
  std::vector<std::string> authors;
  if (pugi::xml_node a = root.child("authors"); a) {
    for (pugi::xml_node au = a.child("author"); au; au = au.next_sibling("author")) {
      authors.emplace_back(au.text().get());
    }
  }

  pugi::xml_node list = root.child("commentList");
  if (!list) {
    // Empty author table with no <commentList> is technically legal;
    // surface the empty list rather than failing.
    return out;
  }
  for (pugi::xml_node c = list.child("comment"); c; c = c.next_sibling("comment")) {
    const std::string_view ref = c.attribute("ref").value();
    if (ref.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "comment: missing/empty ref", "context=comments_reader");
    }
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=comments_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "comment: ref unparseable", std::move(ctx));
    }
    CellComment cc;
    cc.row = rc.value().first;
    cc.col = rc.value().second;
    const std::int64_t author_id = c.attribute("authorId").as_llong(-1);
    if (author_id >= 0 && static_cast<std::size_t>(author_id) < authors.size()) {
      cc.author = authors[static_cast<std::size_t>(author_id)];
    }
    if (pugi::xml_node text = c.child("text"); text) {
      // Comment XML never carries `<rPh>` in practice (kana phonetic
      // guides only attach to shared-string entries), so the unified
      // walker's rPh-skip is a no-op for this path.
      (void)append_rich_text(text, cc.text);
    }
    out.push_back(std::move(cc));
  }
  return out;
}

}  // namespace formulon::io
