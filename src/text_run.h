//
// One formatted run of a note's text.

#ifndef FORMULON_TEXT_RUN_H_
#define FORMULON_TEXT_RUN_H_

#include <optional>
#include <string>

#include "styles.h"

namespace formulon {

/// A stretch of text in one font; an unset font is the default comment font.
struct TextRun {
  std::string text;
  std::optional<FontRecord> font;
};

}  // namespace formulon

#endif  // FORMULON_TEXT_RUN_H_
