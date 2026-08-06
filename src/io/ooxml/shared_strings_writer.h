//
// OOXML shared-string table builder and writer.

#ifndef FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_
#define FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace formulon {
class Workbook;
namespace io {

struct SharedStringEntry {
  std::string text;
  std::string phonetic;
};

class SharedStrings {
 public:
  std::uint32_t intern(std::string_view text, std::string_view phonetic);
  std::uint32_t index_of(std::string_view text, std::string_view phonetic) const;
  bool empty() const noexcept { return entries_.empty(); }
  std::uint32_t total_count() const noexcept { return total_count_; }
  const std::vector<SharedStringEntry>& entries() const noexcept { return entries_; }

 private:
  static std::string key_for(std::string_view text, std::string_view phonetic);

  std::uint32_t total_count_ = 0;
  std::vector<SharedStringEntry> entries_;
  std::unordered_map<std::string, std::uint32_t> index_;
};

SharedStrings BuildSharedStrings(const Workbook& workbook);
std::string WriteSharedStrings(const SharedStrings& strings);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_
