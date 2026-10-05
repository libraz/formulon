#ifndef FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_INTERNAL_H_
#define FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_INTERNAL_H_

#include <memory>
#include <string>

#include "tests/oracle/json_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace workbook_builder_detail {

Error invalid(std::string message);
Expected<std::unique_ptr<Workbook>, Error> build_workbook(const JsonValue& spec);

}  // namespace workbook_builder_detail
}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_INTERNAL_H_
