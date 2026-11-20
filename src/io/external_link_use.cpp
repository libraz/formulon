//
// Implementation of `io/external_link_use.h`.

#include "io/external_link_use.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_ref.h"
#include "defined_name.h"
#include "external_link.h"
#include "io/sheet_formula_sites.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/external_book_spelling.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// The widest `[n]` read as a link index, as `Workbook::external_book_indexer()` reads it; keeps n in 32 bits.
// @limit: engine
constexpr std::size_t kMaxIndexDigits = 9;

/// Collects the links a save keeps: those a rewritten holder references, and every link up to the highest index a
/// verbatim holder names. The rewritten holders' links are also kept in the order they were first named.
class LinkUse {
 public:
  LinkUse(const Workbook& wb, std::uint32_t last)
      : wb_(wb), last_(last), referenced_(last + 1U, false), indexer_(wb.external_book_indexer()) {}

  /// True once every link is kept whatever else is found.
  bool done() const noexcept { return verbatim_max_ == last_; }

  void mark(std::uint32_t index) {
    if (index != 0U && index <= last_ && !referenced_[index]) {
      referenced_[index] = true;
      first_use_.push_back(index);
    }
  }
  void mark_verbatim(std::uint32_t index) noexcept {
    if (index > verbatim_max_ && index <= last_) {
      verbatim_max_ = index;
    }
  }

  /// Marks, as verbatim, every `[n]` in `text` whose `n` names a link.
  void scan_raw(std::string_view text) {
    for (std::size_t at = text.find('['); at != std::string_view::npos && !done(); at = text.find('[', at + 1U)) {
      std::size_t end = at + 1U;
      std::uint32_t n = 0;
      while (end < text.size() && end - at - 1U < kMaxIndexDigits && text[end] >= '0' && text[end] <= '9') {
        n = n * 10U + static_cast<std::uint32_t>(text[end] - '0');
        ++end;
      }
      if (end > at + 1U && end < text.size() && text[end] == ']' && has_link(n)) {
        mark_verbatim(n);
      }
    }
  }

  /// Parses `text` as a formula or name body and marks the links its references name, as verbatim when the writer
  /// copies `text` unchanged.
  void scan_formula(std::string_view text, bool verbatim) {
    text = strip_formula_prefix(text);
    if (done() || text.find_first_of("[!") == std::string_view::npos) {
      return;
    }
    Arena arena;
    // Name bodies come through here too; their grammar only adds the top-level union.
    parser::Parser parser(text, arena, parser::name_body_options());
    const parser::AstNode* root = parser.parse();
    if (root == nullptr || !parser.errors().empty()) {
      // Text the scan cannot read may name any link.
      mark_verbatim(last_);
      return;
    }
    const auto use = [this, verbatim](std::uint32_t index) {
      if (verbatim) {
        mark_verbatim(index);
      } else {
        mark(index);
      }
    };
    std::vector<const parser::AstNode*> pending{root};
    while (!pending.empty()) {
      const parser::AstNode* node = pending.back();
      pending.pop_back();
      if (node->kind() == parser::NodeKind::ExternalRef) {
        if (!parser::is_self_book_name_ref(*node)) {
          use(indexer_.index(indexer_.ctx, node->as_external_ref_path(), node->as_external_ref_book()));
        }
        continue;
      }
      if (node->kind() == parser::NodeKind::ErrorLiteral) {
        const parser::RefErrorSite* site = node->as_ref_error_site();
        if (site != nullptr && site->form == parser::RefErrorSite::Form::kExternal) {
          use(indexer_.index(indexer_.ctx, site->path, site->book));
          continue;
        }
      }
      const std::vector<const parser::AstNode*> children = parser::child_nodes(*node);
      pending.insert(pending.end(), children.rbegin(), children.rend());
    }
  }

  /// The highest index a verbatim holder names; 0 when none does.
  std::uint32_t verbatim_max() const noexcept { return verbatim_max_; }

  /// The indices rewritten holders named, each once, in the order first named.
  const std::vector<std::uint32_t>& first_use() const noexcept { return first_use_; }

  std::vector<std::uint32_t> kept() const {
    std::vector<std::uint32_t> out;
    for (const ExternalLinkRecord& rec : wb_.external_links()) {
      if (rec.index <= verbatim_max_ || referenced_[rec.index]) {
        out.push_back(rec.index);
      }
    }
    sort_ascending(out);
    return out;
  }

 private:
  bool has_link(std::uint32_t index) const {
    const auto& links = wb_.external_links();
    return std::any_of(links.begin(), links.end(), [index](const ExternalLinkRecord& r) { return r.index == index; });
  }

  const Workbook& wb_;
  std::uint32_t last_;
  std::uint32_t verbatim_max_ = 0;
  std::vector<bool> referenced_;
  std::vector<std::uint32_t> first_use_;
  parser::ExternalBookIndexer indexer_;
};

bool IsXmlPartPath(const std::string& path) {
  return strings::ends_with(path, ".xml") || strings::ends_with(path, ".vml") || strings::ends_with(path, ".rels");
}

/// What a walk over the workbook's link holders is for.
enum class Purpose : std::uint8_t {
  kWritten,   ///< The links a save keeps.
  kRenumber,  ///< The links a renumbering must leave at their index.
};

std::uint32_t LastLinkIndex(const Workbook& wb) {
  std::uint32_t last = 0;
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    last = std::max(last, rec.index);
  }
  return last;
}

/// Walks every holder of a link index into `use`: first those that cannot follow a renumbering (verbatim), then
/// defined names in stored order and the sheets' formulas from the active sheet, which fix the first-use order.
void ScanHolders(LinkUse& use, const Workbook& wb, std::uint32_t last, Purpose purpose) {
  const bool renumber = purpose == Purpose::kRenumber;
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    // An OLE or DDE link's references live in raw `<oleObjects>` XML nothing parses. A link with no target is spelled
    // by its number in formulas, and ingestion names it by that number.
    const bool no_target =
        rec.absolute_target.empty() && (rec.target.empty() || rec.target == std::to_string(rec.index));
    if (rec.kind == ExternalLinkRecord::Kind::kOleLink || rec.kind == ExternalLinkRecord::Kind::kDdeLink ||
        (renumber && no_target)) {
      use.mark_verbatim(rec.index);
    }
  }
  // A binary macro or dialog sheet's formulas cannot be read.
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    if (strings::ends_with(wb.sheet(i).opaque_ooxml_part_path(), ".bin")) {
      use.mark_verbatim(last);
    }
  }
  // A retained xlsb sheet tail names links by index, which the xlsb writer maps to positions.
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    for (const XlsbExternSheetEntry& entry : wb.sheet(i).xlsb_tail().extern_sheets) {
      if (renumber) {
        use.mark_verbatim(entry.external_book);
      } else {
        use.mark(entry.external_book);
      }
    }
  }
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    use.scan_raw(sheet.ext_lst_xml());
    use.scan_raw(sheet.comment_vml_extra_xml());
    for (const WorksheetRawChild& child : sheet.raw_extensions()) {
      use.scan_raw(child.xml);
    }
  }
  use.scan_raw(wb.workbook_ext_lst_xml());
  for (const WorksheetRawChild& child : wb.raw_children()) {
    use.scan_raw(child.xml);
  }
  // A renumbering follows the series formulas of a chart reached from a drawing, which are walked with their sheet.
  std::set<std::string> walked_charts;
  if (renumber) {
    for (std::string& path : chart_ref::chart_part_paths(wb)) {
      walked_charts.insert(std::move(path));
    }
  }
  // A save keeps a chart's links at their index, so every link up to the highest one a chart names stays written.
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (use.done()) {
      break;
    }
    if (strings::starts_with(part.path, "xl/externalLinks/") || !IsXmlPartPath(part.path) ||
        walked_charts.count(part.path) != 0U) {
      continue;
    }
    use.scan_raw(std::string_view(reinterpret_cast<const char*>(part.bytes.data()), part.bytes.size()));
  }
  // Kept links: a table formula's links keep their indices. Renumbering: tables follow the whole sheet walk.
  if (!renumber) {
    ForEachTableFormula(wb, [&use](std::string_view text, std::size_t) { use.scan_formula(text, /*verbatim=*/true); });
  }
  for (const DefinedName& name : wb.defined_names()) {
    use.scan_formula(name.formula, /*verbatim=*/false);
  }
  // Sheets are walked from the active one, then the rest in tab order. A conditional-format or validation formula
  // is held by book name like a cell formula, so a renumbering follows it.
  const std::size_t active = wb.active_sheet();
  for (std::size_t step = 0; step < wb.sheet_count() && !use.done(); ++step) {
    const std::size_t i = step == 0U ? active : (step <= active ? step - 1U : step);
    ForEachSheetFormula(wb.sheet(i), [&use](std::string_view text, SheetFormulaSite) {
      if (!text.empty()) {
        use.scan_formula(text, /*verbatim=*/false);
      }
    });
    if (renumber && !use.done() && !chart_ref::for_each_sheet_chart_site(wb, i, [&use](std::string_view text) {
          use.scan_formula(text, /*verbatim=*/false);
        })) {
      use.mark_verbatim(last);
    }
  }
  if (renumber && !use.done()) {
    ForEachTableFormula(wb, [&use](std::string_view text, std::size_t) { use.scan_formula(text, /*verbatim=*/false); });
  }
}

}  // namespace

std::string external_link_file_stem(const ExternalLinkRecord& link) {
  const std::string_view target = link.target;
  const std::size_t slash = target.find_last_of("/\\");
  const std::string_view file = slash == std::string_view::npos ? target : target.substr(slash + 1U);
  return parser::percent_decode(file.substr(0, file.rfind('.')));
}

std::vector<std::uint32_t> table_reference_link_indices(const Workbook& wb) {
  const parser::ExternalBookIndexer indexer = wb.external_book_indexer();
  std::vector<std::uint32_t> out;
  const auto scan = [&indexer, &out](std::string_view text) {
    text = strip_formula_prefix(text);
    if (text.find('[') == std::string_view::npos) {
      return;
    }
    Arena arena;
    parser::Parser parser(text, arena, parser::name_body_options());
    const parser::AstNode* root = parser.parse();
    if (root == nullptr || !parser.errors().empty()) {
      return;
    }
    std::vector<const parser::AstNode*> pending{root};
    while (!pending.empty()) {
      const parser::AstNode* node = pending.back();
      pending.pop_back();
      if (node->kind() == parser::NodeKind::ExternalRef) {
        if (!node->as_external_ref_structured().empty() && !parser::is_self_book_name_ref(*node)) {
          out.push_back(indexer.index(indexer.ctx, node->as_external_ref_path(), node->as_external_ref_book()));
        }
        continue;
      }
      const std::vector<const parser::AstNode*> children = parser::child_nodes(*node);
      pending.insert(pending.end(), children.begin(), children.end());
    }
  };
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    ForEachSheetFormula(wb.sheet(i), [&scan](std::string_view text, SheetFormulaSite) { scan(text); });
  }
  ForEachTableFormula(wb, [&scan](std::string_view text, std::size_t) { scan(text); });
  for (const DefinedName& name : wb.defined_names()) {
    scan(name.formula);
  }
  sort_ascending(out);
  out.erase(std::unique(out.begin(), out.end()), out.end());
  return out;
}

std::vector<std::uint32_t> written_external_link_indices(const Workbook& wb) {
  const std::uint32_t last = LastLinkIndex(wb);
  if (last == 0U) {
    return {};
  }
  LinkUse use(wb, last);
  ScanHolders(use, wb, last, Purpose::kWritten);
  return use.kept();
}

void renumber_external_links_by_first_use(Workbook& wb) {
  if (wb.external_links().size() < 2U) {
    return;
  }
  const std::uint32_t last = LastLinkIndex(wb);
  LinkUse use(wb, last);
  ScanHolders(use, wb, last, Purpose::kRenumber);
  const std::uint32_t pinned = use.verbatim_max();
  // The movable links' own indices are the slots they are dealt back into.
  std::vector<std::uint32_t> slots;
  std::vector<bool> movable(last + 1U, false);
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    if (rec.index > pinned) {
      slots.push_back(rec.index);
      movable[rec.index] = true;
    }
  }
  if (slots.size() < 2U) {
    return;
  }
  sort_ascending(slots);
  std::vector<std::uint32_t> order;
  std::vector<bool> placed(last + 1U, false);
  for (const std::uint32_t index : use.first_use()) {
    if (movable[index] && !placed[index]) {
      placed[index] = true;
      order.push_back(index);
    }
  }
  for (const std::uint32_t index : slots) {
    if (!placed[index]) {
      order.push_back(index);
    }
  }
  std::vector<std::uint32_t> remap(last + 1U);
  for (std::uint32_t i = 0; i <= last; ++i) {
    remap[i] = i;
  }
  bool identity = true;
  for (std::size_t k = 0; k < order.size(); ++k) {
    remap[order[k]] = slots[k];
    identity = identity && order[k] == slots[k];
  }
  if (identity) {
    return;
  }
  const auto map_index = [&remap](std::uint32_t index) { return index <= remap.size() - 1U ? remap[index] : index; };
  std::vector<ExternalLinkRecord> links = wb.external_links();
  for (ExternalLinkRecord& rec : links) {
    rec.index = map_index(rec.index);
  }
  sort_by_index(links, [](const ExternalLinkRecord& a, const ExternalLinkRecord& b) { return a.index < b.index; });
  wb.set_external_links(std::move(links));
  std::vector<XlsbExternSheetEntry> xti = wb.xti_order();
  for (XlsbExternSheetEntry& entry : xti) {
    entry.external_book = map_index(entry.external_book);
  }
  wb.set_xti_order(std::move(xti));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    XlsbSheetTail tail = wb.sheet(i).xlsb_tail();
    for (XlsbExternSheetEntry& entry : tail.extern_sheets) {
      entry.external_book = map_index(entry.external_book);
    }
    wb.sheet(i).set_xlsb_tail(std::move(tail));
  }
  chart_ref::respell_all_book_indices(wb, remap);
}

}  // namespace io
}  // namespace formulon
