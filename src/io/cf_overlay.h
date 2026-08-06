//
// Reconciliation of the raw worksheet-level `<extLst>` x14
// conditional-formatting overlay against the in-memory CF model.
//
// The OOXML reader captures the worksheet `<extLst>` verbatim
// (`Sheet::ext_lst_xml()`) and the writer re-emits it unchanged, so the
// Excel 2010+ `<x14:conditionalFormattings>` overlay survives a save
// cycle without being modelled. Each `<x14:cfRule id="{GUID}">` in that
// overlay cross-references a legacy `<cfRule id="{GUID}">` in the model
// (see `src/io/cf_reader.cpp`). When a mutation removes a model rule,
// the raw overlay must be pruned in step, or the deleted rule's x14
// payload would be re-emitted on save as a dangling reference and the
// rule would resurface on reopen.
//
// Design references:
//   * src/io/cf_reader.h (overlay decode; GUID cross-reference)
//   * src/io/cf_writer.h (legacy CF emission; overlay passthrough note)

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

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_OVERLAY_H_
