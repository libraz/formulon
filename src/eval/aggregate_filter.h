#ifndef FORMULON_EVAL_AGGREGATE_FILTER_H_
#define FORMULON_EVAL_AGGREGATE_FILTER_H_

#include <vector>

#include "value.h"

namespace formulon {
namespace eval {
namespace aggregate_filter_detail {

void compact_by_flags(std::vector<Value>* cells, const std::vector<bool>& hidden, bool use_hidden,
                      const std::vector<bool>& nested, bool use_nested);

}  // namespace aggregate_filter_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_AGGREGATE_FILTER_H_
