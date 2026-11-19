// Excel equality and ascending order of two pictures, read from the picture
// table the current evaluation published on this thread. With no table
// published, or for an id the table does not hold, both fall back to
// `picture_id`; a Debug build asserts that a table is published whenever two
// different ids are compared. The table is not mutated during an evaluation
// call, so workers read it without a lock.

#ifndef FORMULON_EVAL_PICTURE_COMPARE_H_
#define FORMULON_EVAL_PICTURE_COMPARE_H_

#include "picture_table.h"
#include "value.h"

namespace formulon::eval {
namespace detail {

inline thread_local const PictureTable* g_current_picture_table = nullptr;

}  // namespace detail

/// The picture table published on this thread, or nullptr.
inline const PictureTable* current_picture_table() noexcept {
  return detail::g_current_picture_table;
}

/// Publishes `table` on this thread for the scope's life and restores the
/// previous one on exit. A null `table` keeps the outer one.
class ScopedPictureTable {
 public:
  explicit ScopedPictureTable(const PictureTable* table) noexcept : previous_(detail::g_current_picture_table) {
    if (table != nullptr) {
      detail::g_current_picture_table = table;
    }
  }

  ~ScopedPictureTable() noexcept { detail::g_current_picture_table = previous_; }

  ScopedPictureTable(const ScopedPictureTable&) = delete;
  ScopedPictureTable& operator=(const ScopedPictureTable&) = delete;
  ScopedPictureTable(ScopedPictureTable&&) = delete;
  ScopedPictureTable& operator=(ScopedPictureTable&&) = delete;

 private:
  const PictureTable* previous_;
};

/// True iff `a` and `b` show the same media and their alts are both absent,
/// or both present and equal under the active profile's text `=` collation.
bool pictures_equal(PicturePayload a, PicturePayload b);

/// Returns -1, 0 or 1 as `a` sorts before, ties with or sorts after `b`
/// ascending: media rank, then a present alt before an absent one, then the
/// alts in the active profile's `kDynamic` collation.
int pictures_order(PicturePayload a, PicturePayload b);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_PICTURE_COMPARE_H_
