//
// C ABI - threaded comments (threads, replies, mentions) and the workbook
// person list they name.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;

namespace {

// Copies `count` caller mentions into the model shape.
fm_status_t mentions_from_fm(const fm_mention_t* mentions, uint32_t count, std::vector<formulon::Mention>* out) {
  if (mentions == nullptr && count != 0U) {
    return invalid_argument("NULL mentions with a non-zero count", "mention_count=" + std::to_string(count));
  }
  for (uint32_t i = 0; i < count; ++i) {
    if (auto rc = check_utf8({mentions[i].person_id, mentions[i].mention_id}); rc != 0) {
      return rc;
    }
    formulon::Mention m;
    m.person_id = text_or_empty(mentions[i].person_id);
    m.mention_id = text_or_empty(mentions[i].mention_id);
    m.start = mentions[i].start;
    m.length = mentions[i].length;
    out->push_back(std::move(m));
  }
  return 0;
}

// Resets the read scratch and the arenas this file publishes through, and
// returns the handle they belong to.
fm_workbook_t* reset_scratch(const fm_workbook_t* wb) {
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  fresh_read_scratch(wb);
  scratch->mention_scratch.clear();
  return scratch;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_threaded_comment_count(const fm_workbook_t* wb, size_t sheet_index,
                                                           size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out_count is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).threaded_comments().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_threaded_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                        fm_threaded_comment_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& list = wb->workbook().sheet(sheet_index).threaded_comments();
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ThreadedComment& c = list[idx];
  fm_workbook_t* scratch = reset_scratch(wb);
  std::vector<fm_mention_t> mentions;
  mentions.reserve(c.mentions.size());
  for (const formulon::Mention& m : c.mentions) {
    mentions.push_back(fm_mention_t{store_cstr(scratch->read_scratch, m.person_id),
                                    store_cstr(scratch->read_scratch, m.mention_id), m.start, m.length});
  }
  *out = fm_threaded_comment_t{};
  out->id = store_cstr(scratch->read_scratch, c.id);
  out->row = c.row;
  out->col = c.col;
  out->person_id = store_cstr(scratch->read_scratch, c.person_id);
  out->created = store_cstr(scratch->read_scratch, c.created);
  out->text = store_cstr(scratch->read_scratch, c.text);
  out->parent_id = store_cstr(scratch->read_scratch, c.parent_id);
  out->done = c.done ? 1 : 0;
  out->mention_count = static_cast<uint32_t>(mentions.size());
  out->mentions = scratch->mention_scratch.adopt(std::move(mentions));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_threaded_comment(const fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                     int32_t* out_present, fm_threaded_comment_t* out) {
  clear_last_error();
  if (id == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out = fm_threaded_comment_t{};
  const auto& list = wb->workbook().sheet(sheet_index).threaded_comments();
  const std::size_t idx = formulon::find_threaded_comment(list, id);
  if (idx == list.size()) {
    return 0;
  }
  *out_present = 1;
  return fm_sheet_get_threaded_comment_at(wb, sheet_index, idx, out);
}

extern "C" fm_status_t fm_sheet_add_threaded_comment(fm_workbook_t* wb, size_t sheet_index,
                                                     const fm_threaded_comment_t* comment) {
  clear_last_error();
  if (comment == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "comment is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({comment->id, comment->person_id, comment->created, comment->text, comment->parent_id});
      rc != 0) {
    return rc;
  }
  formulon::ThreadedComment model;
  if (auto rc = mentions_from_fm(comment->mentions, comment->mention_count, &model.mentions); rc != 0) {
    return rc;
  }
  model.id = text_or_empty(comment->id);
  model.row = comment->row;
  model.col = comment->col;
  model.person_id = text_or_empty(comment->person_id);
  model.created = text_or_empty(comment->created);
  model.text = text_or_empty(comment->text);
  model.parent_id = text_or_empty(comment->parent_id);
  model.done = comment->done != 0;
  return finish_status(wb->workbook().add_threaded_comment(sheet_index, std::move(model)));
}

extern "C" fm_status_t fm_sheet_update_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                        const char* text, const fm_mention_t* mentions,
                                                        uint32_t mention_count) {
  clear_last_error();
  if (id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "id is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id, text}); rc != 0) {
    return rc;
  }
  std::vector<formulon::Mention> model;
  if (auto rc = mentions_from_fm(mentions, mention_count, &model); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().edit_threaded_comment(sheet_index, id, text_or_empty(text), std::move(model)));
}

extern "C" fm_status_t fm_sheet_set_thread_resolved(fm_workbook_t* wb, size_t sheet_index, const char* thread_id,
                                                    int32_t done) {
  clear_last_error();
  if (thread_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "thread_id is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({thread_id}); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().set_thread_resolved(sheet_index, thread_id, done != 0));
}

extern "C" fm_status_t fm_sheet_remove_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id) {
  clear_last_error();
  if (id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "id is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  const auto& comments = wb->workbook().sheet(sheet_index).threaded_comments();
  if (formulon::find_threaded_comment(comments, id) == comments.size()) {
    return not_found("no threaded comment has that id", std::string("id=") + id);
  }
  return finish_status(wb->workbook().remove_threaded_comment(sheet_index, id));
}

extern "C" fm_status_t fm_workbook_get_person_count(const fm_workbook_t* wb, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().persons().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_person_at(const fm_workbook_t* wb, size_t idx, fm_person_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const auto& persons = wb->workbook().persons();
  if (auto rc = check_index(idx, persons.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::Person& p = persons[idx];
  fm_workbook_t* scratch = reset_scratch(wb);
  out->id = store_cstr(scratch->read_scratch, p.id);
  out->display_name = store_cstr(scratch->read_scratch, p.display_name);
  out->user_id = store_cstr(scratch->read_scratch, p.user_id);
  out->provider_id = store_cstr(scratch->read_scratch, p.provider_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_person(const fm_workbook_t* wb, const char* id, int32_t* out_present,
                                              fm_person_t* out) {
  clear_last_error();
  if (wb == nullptr || id == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out = fm_person_t{};
  const auto& persons = wb->workbook().persons();
  for (std::size_t i = 0; i < persons.size(); ++i) {
    if (persons[i].id == id) {
      *out_present = 1;
      return fm_workbook_get_person_at(wb, i, out);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_add_person(fm_workbook_t* wb, const fm_person_t* person) {
  clear_last_error();
  if (wb == nullptr || person == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({person->id, person->display_name, person->user_id, person->provider_id}); rc != 0) {
    return rc;
  }
  formulon::Person model;
  model.id = text_or_empty(person->id);
  model.display_name = text_or_empty(person->display_name);
  model.user_id = text_or_empty(person->user_id);
  model.provider_id = text_or_empty(person->provider_id);
  return finish_status(wb->workbook().add_person(std::move(model)));
}

extern "C" fm_status_t fm_workbook_remove_person(fm_workbook_t* wb, const char* id) {
  clear_last_error();
  if (wb == nullptr || id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  const auto& people = wb->workbook().persons();
  if (std::none_of(people.begin(), people.end(), [&](const formulon::Person& p) { return p.id == id; })) {
    return not_found("no person has that id", std::string("id=") + id);
  }
  return finish_status(wb->workbook().remove_person(id));
}
