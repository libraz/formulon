//
// OOXML shared-string table builder and writer.

#ifndef FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_
#define FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "phonetic.h"

namespace formulon {
class Workbook;
namespace io {

struct SharedStringEntry {
  std::string text;
  /// One `<rPh>` block per run, each covering the surface-text span it
  /// names. Empty for an unannotated entry.
  std::vector<PhoneticRun> phonetic;
  /// The `<phoneticPr>` block emitted beside those runs. Ignored for an
  /// entry with no runs, which gets no element at all.
  PhoneticProperties phonetic_props;
};

class SharedStrings {
 public:
  std::uint32_t intern(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                       PhoneticProperties phonetic_props);
  std::uint32_t index_of(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                         PhoneticProperties phonetic_props) const;
  bool empty() const noexcept { return entries_.empty(); }
  std::uint32_t total_count() const noexcept { return total_count_; }
  const std::vector<SharedStringEntry>& entries() const noexcept { return entries_; }

 private:
  /// Interning key. Two cells share an `<si>` only when both their
  /// surface text and their full run list agree, so the spans are part
  /// of the key: `東京都` annotated over `[0,2)` and the same text
  /// annotated over `[0,3)` are different string items. The
  /// `<phoneticPr>` block joins the key for the same reason: one reading
  /// rendered as hiragana and the other as katakana are two items.
  static std::string key_for(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                             PhoneticProperties phonetic_props);

  std::uint32_t total_count_ = 0;
  std::vector<SharedStringEntry> entries_;
  std::unordered_map<std::string, std::uint32_t> index_;
};

SharedStrings BuildSharedStrings(const Workbook& workbook);
std::string WriteSharedStrings(const SharedStrings& strings);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHARED_STRINGS_WRITER_H_
