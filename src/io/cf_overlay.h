//
// Reconciliation of the raw worksheet-level `<extLst>` x14
// conditional-formatting overlay against the in-memory CF model.
//
// The OOXML reader captures the worksheet `<extLst>` verbatim
// (`Sheet::ext_lst_xml()`) and the writer re-emits it unchanged, so the
// Excel 2010+ `<x14:conditionalFormattings>` overlay survives a save
// cycle without being modelled. Each `<x14:cfRule id="{GUID}">` in that
// overlay is reached from a legacy `<cfRule>` through a nested
// `<extLst><ext><x14:id>{GUID}</x14:id>` link, which the reader decodes
// into `CFRule::id` (see `src/io/cf_reader.cpp`).
//
// Both directions of drift have to be handled at save time. A mutation
// that removes a model rule leaves its x14 payload behind, which would
// be re-emitted as a dangling reference and resurface the rule on
// reopen — `reconcile_x14_cf_overlay` prunes it. A rule whose data-bar
// settings were set programmatically has no payload at all, and one whose
// captured payload was edited through the model no longer matches it —
// `merge_x14_cf_entries` builds the first and rewrites the second.
//
// Design references:
//   * src/io/cf_reader.h (overlay decode; the nested-id link)
//   * src/io/cf_writer.h (legacy CF emission; overlay entry construction)

#ifndef FORMULON_IO_CF_OVERLAY_H_
#define FORMULON_IO_CF_OVERLAY_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "io/cf_writer.h"

namespace formulon::parser {
struct ExternalBookIndexer;
}  // namespace formulon::parser

namespace formulon::io {

/// Prunes from `ext_lst_xml` every `<x14:cfRule id="...">` whose id no
/// longer matches any `CFRule::id` in `formats`, then drops each
/// `<x14:conditionalFormatting>` / `<x14:conditionalFormattings>` /
/// `<ext>` element left without meaningful children by that pruning.
/// Returns the reconciled raw `<extLst>` element, the input unchanged
/// (byte-for-byte) when no pruning was needed, or an empty string when
/// nothing survives.
///
/// `<x14:cfRule>` elements without an `id` attribute carry no legacy
/// cross-reference, can never dangle, and are kept verbatim. Neither is a
/// rule the model cannot hold (`x14_cf_rule_model_can_hold`): it was never
/// in `formats`, so its absence there says nothing about a removal. Extension
/// blocks other than `x14:conditionalFormattings` are never touched.
///
/// Conservative fallback: when `ext_lst_xml` does not parse or is not a
/// single `<extLst>` element, the referenced ids cannot be enumerated,
/// so the whole overlay is dropped (empty string) rather than risking a
/// dangling GUID surviving a mutation.
std::string reconcile_x14_cf_overlay(const std::string& ext_lst_xml, const std::vector<cf::ConditionalFormat>& formats);

/// Folds the model's x14 data-bar settings into the worksheet `<extLst>`
/// given by `ext_lst_xml` and returns the merged raw `<extLst>` element.
///
/// A captured `<x14:cfRule>` whose id matches a model data bar is kept
/// byte-for-byte while its `<x14:dataBar>` decodes to the model's
/// settings; once it does not (the model was edited after load), the
/// model-owned attributes and colours are rewritten from the model, and
/// the thresholds and any unmodelled child are kept. A rule
/// that needs a payload and has none gets one built by
/// `build_x14_cf_overlay_entries`, placed at the front of the first
/// `<x14:conditionalFormattings>` the overlay already has (or a new
/// `<ext>`, ahead of the other extensions, or `<extLst>` around it). So does
/// every rule `sheet` places in the extension alone; an id minted for one
/// avoids the ids the overlay already holds.
///
/// Returns `ext_lst_xml` byte-for-byte when nothing changes, so a save
/// that needs no new extension content cannot perturb the captured
/// overlay's serialisation. An unparseable `ext_lst_xml` is likewise
/// returned unchanged. `indexer` numbers the external books a built entry's
/// formulas name (see `storage_feature_formula`).
std::string merge_x14_cf_entries(const std::string& ext_lst_xml, const std::vector<cf::ConditionalFormat>& formats,
                                 const parser::ExternalBookIndexer* indexer = nullptr,
                                 const CfSheetContext& sheet = {});

/// Removes from `ext_lst_xml` the `<x14:cfRule>` elements at `adopted` (ascending
/// positions among them in document order, as `read_conditional_formats`
/// reports them) -- the rules the reader moved into the model, which the
/// writer rebuilds from it -- and the containers that leaves empty. Returns
/// the input unchanged when nothing matches or it does not parse, and an
/// empty string when nothing survives.
std::string strip_x14_cf_rules(const std::string& ext_lst_xml, const std::vector<std::size_t>& adopted);

/// Returns `ext_lst_xml` with every `<x14:cfRule>` id upper case, as Excel
/// saves it and as `read_conditional_formats` reads the ids and links it
/// matches them against. Returns the input unchanged when no id needs it or
/// it does not parse.
std::string upper_x14_cf_rule_ids(const std::string& ext_lst_xml);

/// Breaks the sharing of an x14 link id between legacy rules of `formats`, as Excel does on load: the last rule
/// keeps the id and each earlier one gets a fresh id (see `cf::x14_rule_id`), in its captured link and in the
/// matching payload of `ext_lst_xml`, taken in document order. A payload an earlier rule has no counterpart for
/// is left alone. Returns `ext_lst_xml` with the renamed payload ids, unchanged when no id is shared or it does
/// not parse.
std::string separate_shared_x14_links(std::vector<cf::ConditionalFormat>& formats, const std::string& ext_lst_xml,
                                      std::size_t sheet_index);

/// Largest `priority` of the `<x14:cfRule>` elements of `ext_lst_xml` that carry
/// one (the rules the model cannot hold; a legacy rule's payload has none), or
/// 0 when there is none. Nothing when the XML does not parse as one `<extLst>`
/// or a priority is not an integer.
std::optional<std::int32_t> max_x14_cf_priority(const std::string& ext_lst_xml);

/// Returns `ext_lst_xml` with every `<x14:cfRule>` priority >= `from` raised by
/// one. Returns the input byte-for-byte when nothing shifts, and nothing when
/// it does not parse, a priority is not an integer or one >= `from` is
/// INT32_MAX.
std::optional<std::string> shift_x14_cf_priorities(const std::string& ext_lst_xml, std::int32_t from);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_OVERLAY_H_
