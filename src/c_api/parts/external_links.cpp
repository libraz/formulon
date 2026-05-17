// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - external links enumeration (`xl/externalLinks/`).

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "io/styles_reader.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

extern "C" fm_status_t fm_workbook_external_link_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_external_link_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().external_links().size());
  return 0;
}

extern "C" fm_status_t fm_workbook_external_link_at(fm_workbook_t* wb, uint32_t index, fm_external_link_record_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_external_link_at: NULL argument");
  }
  const auto& links = wb->workbook().external_links();
  if (index >= links.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_external_link_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(links.size()));
  }
  const formulon::io::ExternalLinkRecord& rec = links[index];
  out->index = rec.index;
  out->rel_id = rec.rel_id.c_str();
  out->part_path = rec.part_path.c_str();
  out->target = rec.target.c_str();
  out->target_external = rec.target_external ? 1 : 0;
  switch (rec.kind) {
    case formulon::io::ExternalLinkRecord::Kind::kExternalBook:
      out->kind = FM_EXTERNAL_LINK_KIND_EXTERNAL_BOOK;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kOleLink:
      out->kind = FM_EXTERNAL_LINK_KIND_OLE;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kDdeLink:
      out->kind = FM_EXTERNAL_LINK_KIND_DDE;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kUnknown:
    default:
      out->kind = FM_EXTERNAL_LINK_KIND_UNKNOWN;
      break;
  }
  return 0;
}
