//
// C ABI - threaded comments (threads, replies, mentions) and the workbook
// person list they name.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;

namespace {

// Copies `count` caller mentions into the model shape; a span ending before
// it starts is refused. An empty mention id is left for the engine to generate.
fm_status_t mentions_from_fm(const fm_mention_t* mentions, uint32_t count, std::vector<formulon::Mention>* out) {
  if (mentions == nullptr && count != 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "mentions is NULL while count > 0",
                             "mentions_count=" + std::to_string(count));
  }
  for (uint32_t i = 0; i < count; ++i) {
    if (auto rc = check_utf8({mentions[i].person_id, mentions[i].mention_id}); rc != 0) {
      return rc;
    }
    if (mentions[i].end < mentions[i].start) {
      return invalid_argument("mention ends before it starts", "mentions[" + std::to_string(i) +
                                                                   "] start=" + std::to_string(mentions[i].start) +
                                                                   " end=" + std::to_string(mentions[i].end));
    }
    formulon::Mention m;
    m.person_id = text_or_empty(mentions[i].person_id);
    m.mention_id = text_or_empty(mentions[i].mention_id);
    m.start = mentions[i].start;
    m.length = mentions[i].end - mentions[i].start;
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

// `text` as a read-scratch-backed view, or NULL when it is empty (an absent optional string).
const char* store_optional_cstr(fm_workbook_t* scratch, const std::string& text) {
  return text.empty() ? nullptr : store_cstr(scratch->read_scratch, text);
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
  if (auto rc = check_record_size(*out); rc != 0) {
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
                                    store_cstr(scratch->read_scratch, m.mention_id), m.start, m.start + m.length});
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_threaded_comment_t{};
  out->struct_size = struct_size;
  out->id = store_cstr(scratch->read_scratch, c.id);
  out->row = c.row;
  out->col = c.col;
  out->person_id = store_cstr(scratch->read_scratch, c.person_id);
  out->created = store_optional_cstr(scratch, c.created);
  out->text = store_cstr(scratch->read_scratch, c.text);
  out->parent_id = store_optional_cstr(scratch, c.parent_id);
  out->done = c.done ? 1 : 0;
  out->mentions_count = static_cast<uint32_t>(mentions.size());
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
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_threaded_comment_t{};
  out->struct_size = struct_size;
  const auto& list = wb->workbook().sheet(sheet_index).threaded_comments();
  const std::size_t idx = formulon::find_threaded_comment(list, id);
  if (idx == list.size()) {
    return 0;
  }
  *out_present = 1;
  return fm_sheet_get_threaded_comment_at(wb, sheet_index, idx, out);
}

extern "C" fm_status_t fm_sheet_add_threaded_comment(fm_workbook_t* wb, size_t sheet_index,
                                                     const fm_threaded_comment_t* comment, const char** out_id) {
  clear_last_error();
  if (comment == nullptr || out_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             comment == nullptr ? "comment is NULL" : "out_id is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*comment); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({comment->id, comment->person_id, comment->created, comment->text, comment->parent_id});
      rc != 0) {
    return rc;
  }
  // NULL opens a thread; an empty parent id names nothing.
  if (comment->parent_id != nullptr && comment->parent_id[0] == '\0') {
    return invalid_argument("parent_id is empty", "field=parent_id");
  }
  formulon::ThreadedComment model;
  if (auto rc = mentions_from_fm(comment->mentions, comment->mentions_count, &model.mentions); rc != 0) {
    return rc;
  }
  if (comment->parent_id != nullptr) {
    const auto& list = wb->workbook().sheet(sheet_index).threaded_comments();
    if (formulon::find_threaded_comment(list, comment->parent_id) == list.size()) {
      return not_found("no threaded comment has that id", "parent_id=" + std::string(comment->parent_id));
    }
  }
  model.id = text_or_empty(comment->id);
  model.row = comment->row;
  model.col = comment->col;
  model.person_id = text_or_empty(comment->person_id);
  model.created = text_or_empty(comment->created);
  model.text = text_or_empty(comment->text);
  model.parent_id = text_or_empty(comment->parent_id);
  model.done = comment->done != 0;
  auto added = wb->workbook().add_threaded_comment(sheet_index, std::move(model));
  if (!added) {
    return set_last_error(added.error());
  }
  *out_id = formulon::c_api::parts::publish_scratch(wb, std::move(added.value()));
  return 0;
}

extern "C" fm_status_t fm_sheet_update_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                        const fm_threaded_comment_patch_t* patch) {
  clear_last_error();
  if (id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "id is NULL");
  }
  if (patch == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "patch is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  // No field can be cleared; the whole patch is checked before anything is written.
  if (!patch_op_ok_without_clear(patch->text_op)) {
    return invalid_argument("invalid text_op", "text_op=" + std::to_string(patch->text_op));
  }
  if (!patch_op_ok_without_clear(patch->mentions_op)) {
    return invalid_argument("invalid mentions_op", "mentions_op=" + std::to_string(patch->mentions_op));
  }
  if (!patch_op_ok_without_clear(patch->done_op)) {
    return invalid_argument("invalid done_op", "done_op=" + std::to_string(patch->done_op));
  }
  const bool set_text = patch->text_op == FM_PATCH_OP_SET;
  const bool set_mentions = patch->mentions_op == FM_PATCH_OP_SET;
  const bool set_done = patch->done_op == FM_PATCH_OP_SET;
  if (set_text && patch->text == nullptr) {
    return invalid_argument("SET text is NULL", "field=text");
  }
  if (auto rc = check_utf8({id, set_text ? patch->text : nullptr}); rc != 0) {
    return rc;
  }
  if (set_done) {
    if (auto rc = check_enum_domain(patch->done, 0, 1, "done"); rc != 0) {
      return rc;
    }
  }
  std::vector<formulon::Mention> mentions;
  if (set_mentions) {
    if (patch->mentions == nullptr && patch->mentions_count != 0U) {
      return null_argument("mentions");
    }
    if (auto rc = mentions_from_fm(patch->mentions, patch->mentions_count, &mentions); rc != 0) {
      return rc;
    }
  }
  formulon::Workbook& book = wb->workbook();
  const auto& comments = book.sheet(sheet_index).threaded_comments();
  const std::size_t index = formulon::find_threaded_comment(comments, id);
  if (index == comments.size()) {
    return not_found("no threaded comment has that id", std::string("id=") + id);
  }
  if (set_done && !comments[index].parent_id.empty()) {
    return invalid_argument("done is set on a reply", std::string("id=") + id);
  }
  if (set_text || set_mentions) {
    std::string text = set_text ? std::string(patch->text) : comments[index].text;
    if (!set_mentions) {
      mentions = comments[index].mentions;
    }
    if (auto rc = finish_status(book.edit_threaded_comment(sheet_index, id, std::move(text), std::move(mentions)));
        rc != 0) {
      return rc;
    }
  }
  if (set_done) {
    return finish_status(book.set_thread_resolved(sheet_index, id, patch->done != 0));
  }
  return 0;
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
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto& persons = wb->workbook().persons();
  if (auto rc = check_index(idx, persons.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::Person& p = persons[idx];
  fm_workbook_t* scratch = reset_scratch(wb);
  out->id = store_cstr(scratch->read_scratch, p.id);
  out->display_name = store_cstr(scratch->read_scratch, p.display_name);
  out->user_id = store_optional_cstr(scratch, p.user_id);
  out->provider_id = store_optional_cstr(scratch, p.provider_id);
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
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_person_t{};
  out->struct_size = struct_size;
  const auto& persons = wb->workbook().persons();
  for (std::size_t i = 0; i < persons.size(); ++i) {
    if (formulon::strings::case_insensitive_eq(persons[i].id, id)) {
      *out_present = 1;
      return fm_workbook_get_person_at(wb, i, out);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_add_person(fm_workbook_t* wb, const fm_person_t* person,
                                              const char** out_person_id) {
  clear_last_error();
  if (wb == nullptr || person == nullptr || out_person_id == nullptr) {
    return null_argument();
  }
  *out_person_id = nullptr;
  if (auto rc = check_record_size(*person); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({person->id, person->display_name, person->user_id, person->provider_id}); rc != 0) {
    return rc;
  }
  formulon::Person model;
  model.id = text_or_empty(person->id);
  model.display_name = text_or_empty(person->display_name);
  model.user_id = text_or_empty(person->user_id);
  model.provider_id = text_or_empty(person->provider_id);
  auto added = wb->workbook().add_person(std::move(model));
  if (!added) {
    return set_last_error(added.error());
  }
  *out_person_id = formulon::c_api::parts::publish_scratch(wb, std::string(wb->workbook().persons().back().id));
  return 0;
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
  if (std::none_of(people.begin(), people.end(),
                   [&](const formulon::Person& p) { return formulon::strings::case_insensitive_eq(p.id, id); })) {
    return not_found("no person has that id", std::string("id=") + id);
  }
  return finish_status(wb->workbook().remove_person(id));
}
