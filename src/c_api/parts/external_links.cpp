//
// C ABI - external links enumeration (`xl/externalLinks/`).

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "styles.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::null_argument;

extern "C" fm_status_t fm_workbook_get_external_link_count(const fm_workbook_t* wb, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().external_links().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_external_link_at(const fm_workbook_t* wb, size_t external_link_index,
                                                        fm_external_link_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto& links = wb->workbook().external_links();
  if (auto rc = check_index(external_link_index, links.size(), "external_link_index"); rc != 0) {
    return rc;
  }
  const formulon::ExternalLinkRecord& rec = links[external_link_index];
  out->book_number = rec.index;
  out->rel_id = rec.rel_id.c_str();
  out->part_path = rec.part_path.c_str();
  out->target = rec.target.empty() ? nullptr : rec.target.c_str();
  out->target_external = rec.target_external ? 1 : 0;
  switch (rec.kind) {
    case formulon::ExternalLinkRecord::Kind::kExternalBook:
      out->kind = FM_EXTERNAL_LINK_KIND_EXTERNAL_BOOK;
      break;
    case formulon::ExternalLinkRecord::Kind::kOleLink:
      out->kind = FM_EXTERNAL_LINK_KIND_OLE;
      break;
    case formulon::ExternalLinkRecord::Kind::kDdeLink:
      out->kind = FM_EXTERNAL_LINK_KIND_DDE;
      break;
    case formulon::ExternalLinkRecord::Kind::kUnknown:
    default:
      out->kind = FM_EXTERNAL_LINK_KIND_UNKNOWN;
      break;
  }
  return 0;
}
