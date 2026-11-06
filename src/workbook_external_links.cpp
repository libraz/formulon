//
// Cross-workbook links for `Workbook`: the external-link table lookups, the
// binding of external-book references in formulas, formula ingestion and the
// self-book-name spelling.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "external_link.h"
#include "io/format_detect.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/external_book_spelling.h"
#include "parser/formula_prefix.h"
#include "parser/parser_detail.h"
#include "parser/ref_transforms.h"
#include "parser/structured_ref_payload.h"
#include "parser/structured_ref_spelling.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

using detail::MayNameExternalBook;

// The file name a link target ends in; a `file:` URL is percent-decoded.
std::string LinkFileName(std::string_view target) {
  const std::size_t sep = target.find_last_of("/\\");
  const std::string_view name = sep == std::string_view::npos ? target : target.substr(sep + 1);
  if (!strings::starts_with(target, "file:")) {
    return std::string(name);
  }
  return parser::percent_decode(name);
}

// Directory the link displays: from the alternate absolute URL when Excel
// recorded one, otherwise from the target itself.
std::string LinkDisplayPath(const ExternalLinkRecord& rec) {
  return parser::display_path_for_link_target(rec.absolute_target.empty() ? rec.target : rec.absolute_target);
}

std::string LinkBookName(const ExternalLinkRecord& rec) {
  return LinkFileName(rec.target.empty() ? rec.absolute_target : rec.target);
}

bool IsPathSeparator(char c) noexcept {
  return c == '/' || c == '\\';
}

// Path equality under ASCII case folding, with `/` and `\` interchangeable.
bool SamePath(std::string_view a, std::string_view b) noexcept {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (IsPathSeparator(a[i]) && IsPathSeparator(b[i])) {
      continue;
    }
    if (strings::ascii_to_lower(a[i]) != strings::ascii_to_lower(b[i])) {
      return false;
    }
  }
  return true;
}

// The target a new link records for `path` + `book`: the file name alone,
// a POSIX path as written, and a Windows drive or UNC path as a `file:` URL.
std::string TargetForNewLink(std::string_view path, std::string_view book) {
  std::string out;
  const bool drive = path.size() >= 2 && strings::is_ascii_alpha(path[0]) && path[1] == ':';
  const bool unc = path.size() >= 2 && path[0] == '\\' && path[1] == '\\';
  if (drive) {
    out = "file:///";
    out.append(path);
  } else if (unc) {
    out = "file:";
    for (const char c : path) {
      out.push_back(c == '\\' ? '/' : c);
    }
  } else {
    out.append(path);
  }
  out.append(book);
  return out;
}

// Position in `links` of the link `path` + `book` names (see
// `Workbook::find_external_link`), or `links.size()` when none does.
std::size_t FindLinkPosition(const std::vector<ExternalLinkRecord>& links, std::string_view path,
                             std::string_view book) noexcept {
  std::size_t found = links.size();
  const auto take_lowest = [&links, &found](std::size_t i) {
    if (found == links.size() || links[i].index < links[found].index) {
      found = i;
    }
  };
  if (!path.empty()) {
    std::string key(path);
    key.append(book);
    std::vector<std::string> dirs;
    std::vector<std::string> names;
    dirs.reserve(links.size());
    names.reserve(links.size());
    for (const ExternalLinkRecord& link : links) {
      dirs.push_back(LinkDisplayPath(link));
      names.push_back(LinkBookName(link));
    }
    for (std::size_t i = 0; i < links.size(); ++i) {
      if (!dirs[i].empty() && SamePath(dirs[i] + names[i], key)) {
        return i;
      }
    }
    for (std::size_t i = 0; i < links.size(); ++i) {
      if (dirs[i].empty() && strings::case_insensitive_eq(names[i], book)) {
        take_lowest(i);
      }
    }
    return found;
  }
  for (std::size_t i = 0; i < links.size(); ++i) {
    if (strings::case_insensitive_eq(LinkFileName(links[i].target), book) ||
        strings::case_insensitive_eq(LinkFileName(links[i].absolute_target), book)) {
      take_lowest(i);
    }
  }
  return found;
}

std::uint32_t IndexForBook(const void* ctx, std::string_view path, std::string_view book) {
  const ExternalLinkRecord* rec = static_cast<const Workbook*>(ctx)->find_external_link(path, book);
  return rec == nullptr ? 0U : rec->index;
}

std::uint32_t IndexForSheetQualifier(const void* ctx, std::string_view qualifier) {
  const ExternalLinkRecord* rec = static_cast<const Workbook*>(ctx)->link_for_sheet_qualifier(qualifier);
  return rec == nullptr ? 0U : rec->index;
}

bool ResolveLinkIndex(const void* ctx, std::uint32_t index, parser::ExternalBookDisplay* out) {
  for (const ExternalLinkRecord& rec : *static_cast<const std::vector<ExternalLinkRecord>*>(ctx)) {
    if (rec.index != index) {
      continue;
    }
    out->book = LinkBookName(rec);
    if (out->book.empty()) {
      return false;
    }
    out->path = LinkDisplayPath(rec);
    return true;
  }
  return false;
}

}  // namespace

std::vector<const ExternalLinkRecord*> Workbook::external_links_by_index(bool skip_ole_dde) const {
  std::vector<const ExternalLinkRecord*> out;
  out.reserve(external_links_.size());
  for (const ExternalLinkRecord& rec : external_links_) {
    if (skip_ole_dde &&
        (rec.kind == ExternalLinkRecord::Kind::kOleLink || rec.kind == ExternalLinkRecord::Kind::kDdeLink)) {
      continue;
    }
    // Insertion sort: stable, and the list is almost always in order already.
    std::size_t at = out.size();
    out.push_back(&rec);
    for (; at > 0 && out[at - 1]->index > rec.index; --at) {
      out[at] = out[at - 1];
    }
    out[at] = &rec;
  }
  return out;
}

const ExternalLinkRecord* Workbook::find_external_link(std::string_view path, std::string_view book) const noexcept {
  const std::size_t at = FindLinkPosition(external_links_, path, book);
  return at == external_links_.size() ? nullptr : &external_links_[at];
}

const ExternalLinkRecord* Workbook::link_for_sheet_qualifier(std::string_view qualifier) const noexcept {
  if (sheet_by_name(qualifier) != nullptr) {
    return nullptr;
  }
  const ExternalLinkRecord* rec = find_external_link({}, qualifier);
  if (rec == nullptr || rec->kind == ExternalLinkRecord::Kind::kOleLink ||
      rec->kind == ExternalLinkRecord::Kind::kDdeLink) {
    return nullptr;
  }
  return rec;
}

parser::ExternalBookIndexer Workbook::external_book_indexer() const noexcept {
  return parser::ExternalBookIndexer{&IndexForBook, this, &IndexForSheetQualifier};
}

void Workbook::bind_external_books(const parser::AstNode& root) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  bind_external_books_unlocked(root);
}

void Workbook::bind_external_books(std::string_view formula) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  bind_external_books_unlocked(formula);
}

void Workbook::bind_external_books_unlocked(std::string_view formula) {
  if (!MayNameExternalBook(formula)) {
    return;
  }
  formula = strip_formula_prefix(formula);
  Arena arena;
  if (const parser::AstNode* root = parse_indexable_formula(formula, arena); root != nullptr) {
    bind_external_books_unlocked(*root);
  }
}

void Workbook::bind_external_books_unlocked(const parser::AstNode& root) {
  std::vector<const parser::AstNode*> pending{&root};
  while (!pending.empty()) {
    const parser::AstNode* node = pending.back();
    pending.pop_back();
    if (node->kind() != parser::NodeKind::ExternalRef) {
      // Reversed, so references bind in source order.
      const std::vector<const parser::AstNode*> children = parser::child_nodes(*node);
      pending.insert(pending.end(), children.rbegin(), children.rend());
      continue;
    }
    if (parser::is_self_book_name_ref(*node)) {
      continue;
    }
    const std::string_view path = node->as_external_ref_path();
    const std::string_view book = node->as_external_ref_book();
    const std::size_t at = FindLinkPosition(external_links_, path, book);
    ExternalLinkRecord* rec = at == external_links_.size() ? nullptr : &external_links_[at];
    if (rec == nullptr) {
      std::uint32_t next = 1;
      for (const ExternalLinkRecord& existing : external_links_) {
        next = existing.index >= next ? existing.index + 1U : next;
      }
      ExternalLinkRecord created;
      created.index = next;
      created.kind = ExternalLinkRecord::Kind::kExternalBook;
      created.target_external = true;
      created.target = TargetForNewLink(path, book);
      external_links_.push_back(std::move(created));
      rec = &external_links_.back();
    } else if (!path.empty() && LinkDisplayPath(*rec).empty()) {
      // Excel keeps one link for the book and retargets it at the absolute path.
      rec->target = TargetForNewLink(path, book);
      rec->absolute_target = rec->target;
      rec->body_stale = rec->body_stale || !rec->part_path.empty();
    }
    for (const std::string_view sheet : {node->as_external_ref_sheet(), node->as_external_ref_sheet_end()}) {
      if (sheet.empty() || rec->book.sheet_index(sheet) != ExternalBook::kNoSheet) {
        continue;
      }
      rec->book.sheet_names.emplace_back(sheet);
      // A loaded body no longer lists every sheet the model does.
      rec->body_stale = rec->body_stale || !rec->part_path.empty();
    }
  }
}

std::string Workbook::ingest_stored_formula(std::string_view stored, FormulaSource source) {
  std::string text = normalize_formula_text(std::string(stored), source);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (text.find('[') != std::string::npos) {
    // A link whose rels named no target is the book its index spells.
    for (ExternalLinkRecord& rec : external_links_) {
      const bool book_kind =
          rec.kind == ExternalLinkRecord::Kind::kExternalBook || rec.kind == ExternalLinkRecord::Kind::kUnknown;
      if (book_kind && rec.target.empty() && rec.absolute_target.empty()) {
        rec.target = std::to_string(rec.index);
      }
    }
    text = parser::spell_external_books(text, parser::ExternalBookResolver{&ResolveLinkIndex, &external_links_});
  }
  bind_external_books_unlocked(std::string_view(text));
  return text;
}

namespace {

/// Where a formula is entered, for the bare-table-name predicate: the
/// defined names visible there are workbook-scope ones and `sheet`'s locals.
struct EnteredNameScope {
  const Workbook* workbook;
  std::int32_t sheet;  ///< -1 for workbook scope only.
};

bool IsBareTableName(const void* ctx, std::string_view name) {
  const EnteredNameScope& scope = *static_cast<const EnteredNameScope*>(ctx);
  if (!scope.workbook->is_table_name(name)) {
    return false;
  }
  for (const DefinedName& entry : scope.workbook->defined_names()) {
    if ((entry.local_sheet_id < 0 || entry.local_sheet_id == scope.sheet) &&
        strings::case_insensitive_eq(entry.name, name)) {
      return false;
    }
  }
  return true;
}

}  // namespace

std::string Workbook::bind_entered_structured_refs(std::string_view text, std::size_t sheet, std::uint32_t row,
                                                   std::uint32_t col) const {
  const std::optional<std::size_t> own = table_containing(sheet, row, col);
  const std::string_view own_table = own ? std::string_view(tables_[*own].display_name) : std::string_view();
  const EnteredNameScope scope{this, static_cast<std::int32_t>(sheet)};
  return parser::bind_structured_refs(text, own_table, parser::BareTableNamePredicate{&IsBareTableName, &scope});
}

void Workbook::ingest_entered_formula(std::string& formula) {
  formula = normalize_formula_text(std::move(formula), FormulaSource::kEntered);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const EnteredNameScope scope{this, -1};
  formula = parser::bind_structured_refs(formula, {}, parser::BareTableNamePredicate{&IsBareTableName, &scope});
  if (!file_name_.empty()) {
    formula = parser::bind_self_book_names(formula, file_name_);
  }
  bind_external_books_unlocked(std::string_view(formula));
}

Expected<void, Error> Workbook::set_file_name(std::string_view name) {
  // Excel's limit on a file name.
  constexpr std::size_t kMaxFileNameBytes = 218;
  if (!name.empty()) {
    const auto refuse = [name](const char* reason) {
      return make_error(FormulonErrorCode::kInvalidArgument, std::string("set_file_name: ") + reason,
                        "name=" + std::string(name));
    };
    if (!parser::detail::HasWorkbookExtension(name)) {
      return refuse("no workbook extension");
    }
    if (name.find_first_of(std::string_view("[]:/\\?*\0", 8)) != std::string_view::npos) {
      return refuse("forbidden character");
    }
    if (!sheet_names::valid_utf8(name)) {
      return refuse("not valid UTF-8");
    }
    if (name.size() > kMaxFileNameBytes) {
      return refuse("longer than 218 bytes");
    }
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    for (const ExternalLinkRecord& link : external_links_) {
      if (strings::case_insensitive_eq(LinkBookName(link), name)) {
        return refuse("an external link names the same book");
      }
    }
  }
  file_name_.assign(name);
  return Expected<void, Error>::Ok();
}

std::string Workbook::spell_self_book_names(std::string_view text) const {
  return parser::spell_self_book_names(text, file_name_);
}

bool Workbook::spells_self_book_names(std::string_view text) const noexcept {
  return !file_name_.empty() && text.find("[0]!") != std::string_view::npos;
}

std::string Workbook::bind_self_book_names(std::string_view text) const {
  return parser::bind_self_book_names(text, file_name_);
}

}  // namespace formulon
