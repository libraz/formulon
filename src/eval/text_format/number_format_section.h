// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Section splitter and classifier interface. The actual entry-point
// declarations (`split_sections` and `classify`) live in
// `number_format_types.h` because they are also referenced from
// `number_format.cpp` (the public dispatch site); this header re-exports
// those declarations through the same include and exists primarily to
// document the seam between the section-level analysis stage and the
// per-token tokenizer stage in `number_format_tokenizer.cpp`.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SECTION_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SECTION_H_

#include "eval/text_format/number_format_types.h"

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SECTION_H_
