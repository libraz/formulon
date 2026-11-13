//
// MS-XLSB conditional-formatting block codec: one
// BrtBeginConditionalFormatting .. BrtEndConditionalFormatting run to and
// from a `cf::ConditionalFormat`. Layouts follow Excel 365
// (macOS) re-saves, cross-checked against the .xlsx Excel wrote for the
// same workbook; `cf_records.cpp` carries the field tables.
//
// Decoding is all-or-nothing per block: a block holding any value outside
// the supported set (a theme colour, an unsupported template, a pivot-scoped
// block, a formula the Ptg codec refuses) decodes to nothing, and the
// caller keeps its bytes verbatim instead.

#ifndef FORMULON_IO_XLSB_CF_RECORDS_H_
#define FORMULON_IO_XLSB_CF_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "cf/cf_types.h"
#include "io/cf_layout.h"
#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "sheet_passthrough.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtBeginConditionalFormatting = 461;
inline constexpr std::uint16_t kBrtEndConditionalFormatting = 462;

/// Parses a conditional-format threshold given as text; empty text is 0.
Expected<double, Error> parse_threshold_number(const std::string& text);

/// The rule-kind fields BrtBeginCFRule and BrtBeginCFRule14 share: iType,
/// iTemplate, iParam and the flag word (stopIfTrue 0x02, above 0x04, bottom
/// 0x08, percent 0x10).
struct CfRuleFields {
  std::uint32_t itype = 0;
  std::uint32_t tmpl = 0;
  std::uint32_t param = 0;
  std::uint16_t flags = 0;
};

/// Sets `rule`'s type, its type-specific fields, stopIfTrue and the text of
/// a text rule from `fields` and `text`. False for a combination outside the
/// supported set.
bool decode_cf_rule_fields(const CfRuleFields& fields, std::string text, cf::CFRule& rule);

/// The fields `decode_cf_rule_fields` reads back as `rule`. Fails for a rule
/// type or value the records cannot carry.
Expected<CfRuleFields, Error> encode_cf_rule_fields(const cf::CFRule& rule);

/// False for an iType whose rules state no dxf (colour scale, data bar, icon
/// set): their dxfId slot names nothing, and Excel writes 0 there in an x14
/// rule whatever the styles part holds.
bool cf_rule_type_has_dxf(std::uint32_t itype);

/// Whether the rule type stores formula slots (the rank / average /
/// uniqueness and visual rules carry none).
bool cf_rule_has_formulas(cf::RuleType type);

/// Decodes one framed block, starting at its BrtBeginConditionalFormatting
/// and ending with its BrtEndConditionalFormatting.
std::optional<cf::ConditionalFormat> decode_cf_block(ByteSpan block, const FeatureFormulaReadContext& ctx);

/// Folds the x14 data-bar settings the legacy records cannot carry
/// (negative fill and border, axis position and colour, border, gradient,
/// bar lengths) onto the linked rules of `formats`, as the .xlsx reader
/// does with `<x14:dataBar>`. `records` is a run of framed tail records;
/// the x14 blocks stay retained verbatim. An x14 rule whose records hold
/// anything unsupported leaves its legacy rule as decoded.
void apply_x14_data_bar_overlays(ByteSpan records, std::vector<cf::ConditionalFormat>& formats);

/// Brings the x14 data-bar rules in a sheet's retained tail `records` in
/// line with the model: a rule linked to a model data bar has its
/// model-owned settings rewritten (thresholds are kept), a
/// data-bar rule whose model rule is gone is dropped, along with any block
/// or container that leaves empty. Every data-bar id still present is
/// added to `linked`. Records this does not recognise pass through.
void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked);

/// Adds an x14 data-bar rule to `records` (the tail slot after the
/// hyperlinks) for every model data bar that needs one and is not in
/// `linked`, inside the retained x14 container when there is one, and adds
/// its id to `linked`. Fails for a formula threshold, whose x14 form is
/// unsupported.
Expected<void, Error> add_x14_data_bars(std::vector<std::uint8_t>& records,
                                        const std::vector<cf::ConditionalFormat>& formats,
                                        std::unordered_set<std::string>& linked);

/// True for an icon-set rule with per-level custom icons. Such a rule has no
/// legacy record form: it is written as a standalone x14 rule
/// (BrtBeginCFRule14 with iPri > 0) and never as a BrtBeginCFRule.
bool is_x14_icon_set_rule(const cf::CFRule& rule);

/// True for a rule written only as a standalone x14 rule: a custom icon set,
/// or a rule with formula slots (or none, as top10 / aboveAverage /
/// uniqueness) whose formulas name a sheet
/// other than `own_sheet` (spellings resolved among `sheet_list`, see
/// `parser::references_other_sheet`). Colour scales, data bars and plain
/// icon sets stay legacy records.
bool is_x14_only_rule(const cf::CFRule& rule, std::string_view own_sheet,
                      const std::vector<std::string>* sheet_list = nullptr);

/// Moves every standalone x14 rule (iPri > 0) of the retained tail `records`
/// the model can hold into `formats`. A custom icon set joins a block with
/// the same ranges (or its legacy shadow) and its records are dropped, along
/// with a block or container left empty. Any other rule comes in as
/// one new block per x14 block, as the .xlsx reader adds
/// them; its records stay retained, so the sheet references of the tail keep
/// their table, and `reconcile_x14_rules` rewrites them from the model. A
/// rule holding anything outside the supported shape stays retained.
void take_x14_rules(std::vector<std::uint8_t>& records, const FeatureFormulaReadContext& ctx,
                    std::vector<cf::ConditionalFormat>& formats);

/// True when the retained sheet tail `tail` holds a conditional-format rule
/// `formats` does not: a legacy block `decode_cf_block` refused, or an x14
/// rule whose id names no rule of `formats`. Also true for a tail whose
/// records do not frame.
bool tail_holds_unmodelled_cf(const XlsbSheetTail& tail, const std::vector<cf::ConditionalFormat>& formats);

/// Rewrites, in place, every retained standalone formula rule of `records`
/// from the model rule with its id, keeping Excel's dxf copy while the dxf is
/// unchanged, and adds the id to `linked`. A rule whose model ranges differ
/// from its block's is dropped, as is one gone from the model that
/// `take_x14_rules` could adopt (decoded against the source
/// `extern_sheets`); one it could not adopt stays verbatim. A block or
/// container left empty goes. Fails for a formula the records cannot carry.
Expected<void, Error> reconcile_x14_rules(std::vector<std::uint8_t>& records,
                                          const std::vector<cf::ConditionalFormat>& formats,
                                          const FeatureFormulaWriteContext& ctx, std::string_view own_sheet,
                                          const std::vector<XlsbExternSheetEntry>& extern_sheets,
                                          std::unordered_set<std::string>& linked);

/// Adds every `is_x14_only_rule` rule of `formats` not in `linked` to the
/// tail slot `records`: custom icon sets one x14 block per model block, any
/// other rule one x14 block each, keeping its priority, ahead of the blocks
/// of a retained x14 container. A rule without an id gets
/// `cf::x14_rule_id(sheet_index, priority, ...)` for the write only. Fails
/// for a threshold, icon or formula the records cannot carry, or a level
/// count that does not match the set.
Expected<void, Error> add_x14_rules(std::vector<std::uint8_t>& records,
                                    const std::vector<cf::ConditionalFormat>& formats,
                                    const FeatureFormulaWriteContext& ctx, std::string_view own_sheet,
                                    std::size_t sheet_index, const std::unordered_set<std::string>& linked);

/// Emits `view` as one framed block, leaving out the rules
/// `is_x14_only_rule` places in the x14 records. A data bar whose id is in
/// `linked` has an x14 counterpart carrying its lengths, and keeps the
/// pre-2010 10/90 in the legacy record as Excel does. Fails on a rule the
/// records cannot carry (a formula the Ptg codec refuses, an x14 link id that
/// is not a GUID, a malformed number threshold).
Expected<void, Error> emit_cf_block(std::vector<std::uint8_t>& dst, const CfBlockView& view,
                                    const FeatureFormulaWriteContext& ctx,
                                    const std::unordered_set<std::string>& linked, std::string_view own_sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_RECORDS_H_
