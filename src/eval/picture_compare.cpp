#include "eval/picture_compare.h"

#include <cassert>

#include "eval/collation.h"
#include "eval/eval_profile_scope.h"

namespace formulon::eval {
namespace {

int id_order(PicturePayload a, PicturePayload b) noexcept {
  if (a.picture_id == b.picture_id) {
    return 0;
  }
  return a.picture_id < b.picture_id ? -1 : 1;
}

}  // namespace

bool pictures_equal(PicturePayload a, PicturePayload b) {
  if (a.picture_id == b.picture_id) {
    return true;
  }
  const PictureTable* table = current_picture_table();
  assert(table != nullptr && "pictures compared with no ScopedPictureTable open");
  const PictureEntry* ea = table != nullptr ? table->entry(a.picture_id) : nullptr;
  const PictureEntry* eb = table != nullptr ? table->entry(b.picture_id) : nullptr;
  if (ea == nullptr || eb == nullptr) {
    return false;
  }
  if (ea->media_id != eb->media_id || ea->alt.has_value() != eb->alt.has_value()) {
    return false;
  }
  if (!ea->alt.has_value()) {
    return true;
  }
  return collate(*ea->alt, *eb->alt, current_eval_profile(), CollationStrength::kCompare) == 0;
}

int pictures_order(PicturePayload a, PicturePayload b) {
  if (a.picture_id == b.picture_id) {
    return 0;
  }
  const PictureTable* table = current_picture_table();
  assert(table != nullptr && "pictures compared with no ScopedPictureTable open");
  const PictureEntry* ea = table != nullptr ? table->entry(a.picture_id) : nullptr;
  const PictureEntry* eb = table != nullptr ? table->entry(b.picture_id) : nullptr;
  if (ea == nullptr || eb == nullptr) {
    return id_order(a, b);
  }
  const std::uint32_t ra = table->media_rank(ea->media_id);
  const std::uint32_t rb = table->media_rank(eb->media_id);
  if (ra != rb) {
    return ra < rb ? -1 : 1;
  }
  if (ea->alt.has_value() != eb->alt.has_value()) {
    return ea->alt.has_value() ? -1 : 1;
  }
  if (!ea->alt.has_value()) {
    return 0;
  }
  return collate(*ea->alt, *eb->alt, current_eval_profile(), CollationStrength::kDynamic);
}

}  // namespace formulon::eval
