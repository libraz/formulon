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

#include <string>
#include <vector>

#include "cf/cf_types.h"

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
/// cross-reference, can never dangle, and are kept verbatim. Extension
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
/// `build_x14_cf_overlay_entries`, placed in the first
/// `<x14:conditionalFormattings>` the overlay already has (or a new
/// `<ext>` / `<extLst>` around it).
///
/// Returns `ext_lst_xml` byte-for-byte when nothing changes, so a save
/// that needs no new extension content cannot perturb the captured
/// overlay's serialisation. An unparseable `ext_lst_xml` is likewise
/// returned unchanged.
std::string merge_x14_cf_entries(const std::string& ext_lst_xml, const std::vector<cf::ConditionalFormat>& formats);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_OVERLAY_H_
