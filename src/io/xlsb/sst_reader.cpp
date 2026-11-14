#include "io/xlsb/sst_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/rich_str.h"
#include "phonetic.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {

Expected<std::vector<std::string_view>, Error> DecodeSharedStringsBin(
    const std::vector<std::uint8_t>& body, std::deque<std::string>& text_storage,
    std::vector<std::vector<PhoneticRun>>& out_phonetic, std::vector<PhoneticProperties>& out_phonetic_props) {
  std::vector<std::string_view> entries;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem)) {
      continue;
    }
    // BrtSSTItem ([MS-XLSB] §2.4.293) is a `RichStr`; cells keep the text flat, so its runs are not used.
    ByteSpan p = rec.payload;
    ASSIGN_OR_RETURN(RichStr item, read_rich_str(p));
    text_storage.push_back(std::move(item.text));
    entries.push_back(text_storage.back());
    out_phonetic.push_back(std::move(item.phonetic));
    out_phonetic_props.push_back(item.phonetic_props);
  }
  return entries;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
