//
// OOXML token vocabularies for the `styles.xml` enums that `styles.h`
// stores as ordinals. Each array is indexed by the stored ordinal, so the
// writer maps ordinal -> token by subscript and the reader maps token ->
// ordinal by position. Out-of-range ordinals and unknown tokens are each
// side's own policy.
//
// Header-only, `inline constexpr` — no translation unit needed.

#ifndef FORMULON_IO_STYLES_VOCAB_H_
#define FORMULON_IO_STYLES_VOCAB_H_

#include <cstddef>

namespace formulon {
namespace io {

/// `ST_HorizontalAlignment`, by `CellXf::horizontal_align`.
inline constexpr const char* kHorizontalAlignNames[] = {
    "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed",
};
inline constexpr std::size_t kHorizontalAlignCount = sizeof(kHorizontalAlignNames) / sizeof(kHorizontalAlignNames[0]);

/// `ST_VerticalAlignment`, by `CellXf::vertical_align`.
inline constexpr const char* kVerticalAlignNames[] = {
    "top", "center", "bottom", "justify", "distributed",
};
inline constexpr std::size_t kVerticalAlignCount = sizeof(kVerticalAlignNames) / sizeof(kVerticalAlignNames[0]);

/// `ST_UnderlineValues`, by `FontRecord::underline`.
inline constexpr const char* kUnderlineNames[] = {
    "none", "single", "double", "singleAccounting", "doubleAccounting",
};
inline constexpr std::size_t kUnderlineCount = sizeof(kUnderlineNames) / sizeof(kUnderlineNames[0]);

/// `ST_VerticalAlignRun`, by `FontRecord::vert_align`.
inline constexpr const char* kVertAlignNames[] = {
    "baseline",
    "superscript",
    "subscript",
};
inline constexpr std::size_t kVertAlignCount = sizeof(kVertAlignNames) / sizeof(kVertAlignNames[0]);

/// `ST_FontScheme`, by `FontRecord::scheme`.
inline constexpr const char* kFontSchemeNames[] = {
    "none",
    "major",
    "minor",
};
inline constexpr std::size_t kFontSchemeCount = sizeof(kFontSchemeNames) / sizeof(kFontSchemeNames[0]);

/// `ST_PatternType`, by `FillRecord::pattern`.
inline constexpr const char* kFillPatternNames[] = {
    "none",     "solid",     "mediumGray",   "darkGray",    "lightGray",       "darkHorizontal", "darkVertical",
    "darkDown", "darkUp",    "darkGrid",     "darkTrellis", "lightHorizontal", "lightVertical",  "lightDown",
    "lightUp",  "lightGrid", "lightTrellis", "gray125",     "gray0625",
};
inline constexpr std::size_t kFillPatternCount = sizeof(kFillPatternNames) / sizeof(kFillPatternNames[0]);

/// `ST_BorderStyle`, by `BorderSide::style`.
inline constexpr const char* kBorderStyleNames[] = {
    "none",         "thin",    "medium",        "dashed",     "dotted",           "thick",        "double", "hair",
    "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot",
};
inline constexpr std::size_t kBorderStyleCount = sizeof(kBorderStyleNames) / sizeof(kBorderStyleNames[0]);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_VOCAB_H_
