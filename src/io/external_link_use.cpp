//
// Implementation of `io/external_link_use.h`.

#include "io/external_link_use.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "external_link.h"
#include "io/sheet_formula_sites.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// The widest `[n]` read as a link index, as `Workbook::external_book_indexer()` reads it; keeps n in 32 bits.
// @limit: engine
constexpr std::size_t kMaxIndexDigits = 9;

/// Tracks the highest referenced link index, stopping once every link is used.
class LinkUse {
 public:
  LinkUse(const Workbook& wb, std::uint32_t last) : wb_(wb), last_(last), indexer_(wb.external_book_indexer()) {}

  bool done() const noexcept { return kept_ == last_; }
  std::uint32_t kept() const noexcept { return kept_; }
  void mark(std::uint32_t index) noexcept {
    if (index > kept_ && index <= last_) {
      kept_ = index;
    }
  }

  /// Marks every `[n]` in `text` whose `n` names a link.
  void scan_raw(std::string_view text) {
    for (std::size_t at = text.find('['); at != std::string_view::npos && !done(); at = text.find('[', at + 1U)) {
      std::size_t end = at + 1U;
      std::uint32_t n = 0;
      while (end < text.size() && end - at - 1U < kMaxIndexDigits && text[end] >= '0' && text[end] <= '9') {
        n = n * 10U + static_cast<std::uint32_t>(text[end] - '0');
        ++end;
      }
      if (end > at + 1U && end < text.size() && text[end] == ']' && has_link(n)) {
        mark(n);
      }
    }
  }

  /// Parses `text` as a formula or name body and marks the links its references name.
  void scan_formula(std::string_view text) {
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
      mark(last_);
      return;
    }
    std::vector<const parser::AstNode*> pending{root};
    while (!pending.empty() && !done()) {
      const parser::AstNode* node = pending.back();
      pending.pop_back();
      if (node->kind() == parser::NodeKind::ExternalRef) {
        if (!parser::is_self_book_name_ref(*node)) {
          mark(indexer_.index(indexer_.ctx, node->as_external_ref_path(), node->as_external_ref_book()));
        }
        continue;
      }
      if (node->kind() == parser::NodeKind::ErrorLiteral) {
        const parser::RefErrorSite* site = node->as_ref_error_site();
        if (site != nullptr && site->form == parser::RefErrorSite::Form::kExternal) {
          mark(indexer_.index(indexer_.ctx, site->path, site->book));
          continue;
        }
      }
      if (node->kind() == parser::NodeKind::NameRef && !node->as_name_sheet().empty()) {
        mark(indexer_.qualifier_index(indexer_.ctx, node->as_name_sheet()));
      }
      const std::vector<const parser::AstNode*> children = parser::child_nodes(*node);
      pending.insert(pending.end(), children.rbegin(), children.rend());
    }
  }

 private:
  bool has_link(std::uint32_t index) const {
    const auto& links = wb_.external_links();
    return std::any_of(links.begin(), links.end(), [index](const ExternalLinkRecord& r) { return r.index == index; });
  }

  const Workbook& wb_;
  std::uint32_t last_;
  std::uint32_t kept_ = 0;
  parser::ExternalBookIndexer indexer_;
};

bool IsXmlPartPath(const std::string& path) {
  return strings::ends_with(path, ".xml") || strings::ends_with(path, ".vml") || strings::ends_with(path, ".rels");
}

}  // namespace

std::uint32_t last_written_external_link(const Workbook& wb) {
  std::uint32_t last = 0;
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    last = std::max(last, rec.index);
  }
  if (last == 0U) {
    return 0U;
  }
  LinkUse use(wb, last);
  // An OLE or DDE link's references live in raw `<oleObjects>` XML nothing parses.
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    if (rec.kind == ExternalLinkRecord::Kind::kOleLink || rec.kind == ExternalLinkRecord::Kind::kDdeLink) {
      use.mark(rec.index);
    }
  }
  // A binary macro or dialog sheet's formulas cannot be read.
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    if (strings::ends_with(wb.sheet(i).opaque_ooxml_part_path(), ".bin")) {
      use.mark(last);
    }
  }
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    for (const XlsbExternSheetEntry& entry : wb.sheet(i).xlsb_tail().extern_sheets) {
      use.mark(entry.external_book);
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
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (use.done()) {
      break;
    }
    if (strings::starts_with(part.path, "xl/externalLinks/") || !IsXmlPartPath(part.path)) {
      continue;
    }
    use.scan_raw(std::string_view(reinterpret_cast<const char*>(part.bytes.data()), part.bytes.size()));
  }
  for (std::size_t i = 0; i < wb.sheet_count() && !use.done(); ++i) {
    ForEachSheetFormula(wb.sheet(i), [&use](std::string_view text, SheetFormulaSite) {
      if (!text.empty()) {
        use.scan_formula(text);
      }
    });
  }
  ForEachTableFormula(wb, [&use](std::string_view text, std::size_t) { use.scan_formula(text); });
  for (const DefinedName& name : wb.defined_names()) {
    use.scan_formula(name.formula);
  }
  return use.kept();
}

}  // namespace io
}  // namespace formulon
