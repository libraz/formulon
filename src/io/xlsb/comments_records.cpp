// @size-budget: 4 KB
//
// Implementation of the xlsb comments part. See `io/xlsb/comments_records.h`.

#include "io/xlsb/comments_records.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/threaded_comments_io.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/rich_str.h"
#include "io/xlsb/styles_writer.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// BrtACBegin payload of a note's revision-uid wrapper.
constexpr std::array<std::uint8_t, 6> kNoteUidBeginPayload = {0x01, 0x00, 0x02, 0x11, 0x00, 0x80};

/// Sheets whose notes without a uid get Excel's `{00000000-0006-0000-SSSS-0000NN000000}`, `SSSS` the sheet's
/// position and `NN` the note's; measured on the first two sheets.
constexpr std::size_t kMeasuredUidSheets = 2;

std::string FontKey(const FontRecord& font) {
  std::vector<std::uint8_t> bytes;
  emit_font_record(bytes, font);
  return std::string(bytes.begin(), bytes.end());
}

}  // namespace

Expected<std::vector<CellComment>, Error> read_comments_bin(ByteSpan part, const std::vector<FontRecord>& fonts,
                                                            std::vector<std::string>* authors_out) {
  std::vector<std::string> authors;
  std::vector<CellComment> out;
  std::string uid;
  while (part.size > 0U) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(part));
    ByteSpan p = rec.payload;
    if (rec.type == kBrtCommentAuthor) {
      ASSIGN_OR_RETURN(std::string author, read_xlwidestring(p));
      authors.push_back(std::move(author));
    } else if (rec.type == kUid && p.size == kGuidBytes) {
      uid = FormatGuid(p.data);
    } else if (rec.type == kBrtBeginComment) {
      ASSIGN_OR_RETURN(const std::uint32_t author, read_u32(p));
      ASSIGN_OR_RETURN(const MergeRange cell, read_rfx(p));
      CellComment& note = out.emplace_back();
      note.row = cell.first_row;
      note.col = cell.first_col;
      note.author = author < authors.size() ? authors[author] : std::string();
      note.uid = std::move(uid);
      uid.clear();
    } else if (rec.type == kBrtCommentText && !out.empty()) {
      ASSIGN_OR_RETURN(RichStr text, read_rich_str(p));
      CellComment& note = out.back();
      for (std::size_t i = 0; i < text.runs.size(); ++i) {
        const std::uint32_t start = text.runs[i].ich;
        const std::uint32_t end = i + 1U < text.runs.size() ? text.runs[i + 1U].ich : utf16_units_in(text.text);
        TextRun& run = note.runs.emplace_back();
        run.text = utf16_substring(text.text, start, end > start ? end - start : 0U);
        if (text.runs[i].ifnt < fonts.size()) {
          run.font = fonts[text.runs[i].ifnt];
        }
      }
      note.text = std::move(text.text);
    }
  }
  if (authors_out != nullptr) {
    *authors_out = std::move(authors);
  }
  return out;
}

NoteRunFonts::NoteRunFonts(const std::vector<FontRecord>& table,
                           const std::vector<const std::vector<CellComment>*>& notes) {
  for (const FontRecord& font : table) {
    keys_.push_back(FontKey(font));
  }
  for (const std::vector<CellComment>* list : notes) {
    for (const CellComment& note : *list) {
      for (const TextRun& run : note.runs) {
        if (!run.font) {
          continue;
        }
        std::string key = FontKey(*run.font);
        if (std::find(keys_.begin(), keys_.end(), key) == keys_.end()) {
          keys_.push_back(std::move(key));
          appended_.push_back(*run.font);
        }
      }
    }
  }
}

std::uint16_t NoteRunFonts::index_of(const std::optional<FontRecord>& font) const {
  if (!font) {
    return 0U;
  }
  const auto found = std::find(keys_.begin(), keys_.end(), FontKey(*font));
  return found == keys_.end() ? 0U : static_cast<std::uint16_t>(found - keys_.begin());
}

std::vector<std::uint8_t> write_comments_bin(std::vector<CellComment> notes, std::size_t sheet_index,
                                             const NoteRunFonts& fonts) {
  sort_by_index(
      notes, [](const CellComment& a, const CellComment& b) { return a.row != b.row ? a.row < b.row : a.col < b.col; });
  std::vector<std::string> authors;
  std::unordered_map<std::string, std::uint32_t> author_index;
  for (const CellComment& note : notes) {
    if (author_index.emplace(note.author, static_cast<std::uint32_t>(authors.size())).second) {
      authors.push_back(note.author);
    }
  }
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginComments, ByteSpan{});
  emit_record(out, kBrtBeginCommentAuthors, ByteSpan{});
  for (const std::string& author : authors) {
    payload.clear();
    emit_xlwidestring(payload, author);
    emit_record(out, kBrtCommentAuthor, payload);
  }
  emit_record(out, kBrtEndCommentAuthors, ByteSpan{});
  emit_record(out, kBrtBeginCommentList, ByteSpan{});
  for (std::size_t i = 0; i < notes.size(); ++i) {
    const CellComment& note = notes[i];
    std::array<std::uint8_t, kGuidBytes> guid{};
    if (!ParseGuid(note.uid, guid)) {
      // u32 0, u32 6, u32 sheet position, u32 the note's one-based position.
      guid = {};
      guid[4] = 6U;
      for (std::size_t b = 0; b < 4U; ++b) {
        guid[8U + b] = static_cast<std::uint8_t>((sheet_index >> (8U * b)) & 0xFFU);
        guid[12U + b] = static_cast<std::uint8_t>(((i + 1U) >> (8U * b)) & 0xFFU);
      }
    }
    emit_record(out, kAcBegin, ByteSpan{kNoteUidBeginPayload.data(), kNoteUidBeginPayload.size()});
    emit_record(out, kUid, ByteSpan{guid.data(), guid.size()});
    emit_record(out, kAcEnd, ByteSpan{});
    payload.clear();
    emit_u32(payload, author_index[note.author]);
    emit_rfx(payload, MergeRange{note.row, note.col, note.row, note.col});
    payload.resize(payload.size() + kGuidBytes, 0U);
    emit_record(out, kBrtBeginComment, payload);
    // A thread's stub is plain text; a note is one run per formatted stretch, a plain note one run in font 0.
    std::vector<StrRun> runs;
    if (!is_thread_stub_author(note.author)) {
      std::uint32_t at = 0;
      for (const TextRun& run : note.runs) {
        runs.push_back(StrRun{static_cast<std::uint16_t>(at), fonts.index_of(run.font)});
        at += utf16_units_in(run.text);
      }
      if (runs.empty()) {
        runs.push_back(StrRun{});
      }
    }
    payload.clear();
    emit_rich_str(payload, note.text, runs, {}, PhoneticProperties{});
    emit_record(out, kBrtCommentText, payload);
    emit_record(out, kBrtEndComment, ByteSpan{});
  }
  emit_record(out, kBrtEndCommentList, ByteSpan{});
  emit_record(out, kBrtEndComments, ByteSpan{});
  return out;
}

bool generated_note_uid_is_measured(std::size_t sheet_index, bool first_sheet_has_notes) {
  return sheet_index == 0U || (sheet_index < kMeasuredUidSheets && first_sheet_has_notes);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
