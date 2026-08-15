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
// settings were set programmatically has no payload at all, and those
// settings are simply lost on save unless one is built —
// `merge_x14_cf_entries` folds it in.
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

/// Folds `entries` — `<x14:conditionalFormatting>` elements produced by
/// `build_x14_cf_overlay_entries` — into the worksheet `<extLst>` given
/// by `ext_lst_xml`, and returns the merged raw `<extLst>` element.
///
/// An entry whose `<x14:cfRule id>` already appears anywhere in
/// `ext_lst_xml` is dropped rather than appended: that id came from a
/// loaded file, so the overlay already holds the real payload, including
/// whatever parts of it this engine does not model. The rebuilt entry
/// would be a lossy duplicate.
///
/// Surviving entries go into the first `<x14:conditionalFormattings>`
/// the overlay already has; when there is none, the enclosing `<ext>`
/// (and `<extLst>`, when `ext_lst_xml` is empty) is created around them.
///
/// Returns `ext_lst_xml` byte-for-byte when `entries` is empty or every
/// entry was dropped, so a save that needs no new extension content
/// cannot perturb the captured overlay's serialisation. An unparseable
/// `ext_lst_xml` is likewise returned unchanged, with the entries
/// dropped: preserving bytes that are known to round-trip beats
/// rewriting them from a parse that already failed.
std::string merge_x14_cf_entries(const std::string& ext_lst_xml, const std::string& entries);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_OVERLAY_H_
