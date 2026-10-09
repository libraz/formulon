#!/usr/bin/env python3
"""Binding drift guard.

Several files that describe the C-ABI / embind / N-API surface are
hand-synchronized rather than generated, and have gone out of sync more
than once: a newly added `fm_*` entry point lands in one binding but is
forgotten in another, or a `.d.ts` stops matching what is actually
registered. This script cross-checks those pairs and exits non-zero on
any mismatch, so the drift fails fast in CTest instead of shipping as a
binding-specific runtime regression.

Checks:
  python-exports  Every `LIB.fm_*` symbol called from
                  packages/python/formulon/*.py is present in
                  tools/wasm/capi_exports.txt (the staged WASM export
                  list the Python binding loads), and every symbol in
                  that list is declared in src/c_api/formulon_c.h. The
                  literal `_STATUS_RETURNING_EXPORT_NAMES` tuple in `_c.py`
                  must also equal exactly the intersection of the manifest
                  and header declarations returning `fm_status_t`.
  python-call-signatures
                  Every `LIB.fm_*(...)` call in packages/python/formulon
                  passes as many arguments as src/c_api/formulon_c.h declares
                  parameters, and every scratch block passed to a pointer
                  parameter is at least (for a struct, exactly) the wasm32
                  size of that parameter's pointee. Argument types at
                  non-pointer positions are not covered; see the function's
                  docstring for the exact boundary.
  python-inline-structs
                  The C structs the Python binding decodes with a bare
                  `struct.unpack` instead of a `_structs.Struct` entry --
                  `fm_value_t`, `fm_print_range_t` -- have their size
                  literals and their decoders' offsets and field widths
                  measured against the header.
  dts-wasm        src/wasm/formulon.d.ts (Workbook / WorkbookCtor /
                  FormulonModule method surface) matches what is
                  registered in src/wasm/parts/bindings_register.cpp.
  dts-node        packages/npm-native/index.d.ts (Workbook /
                  WorkbookCtor / free-function surface) matches what is
                  registered in src/node_addon/parts/workbook_class.cc
                  and src/node_addon/addon.cc. It also enforces the exact
                  intentional WASM-only / Node-only method allowlists.
  dts-shared-shapes
                  The two published declaration files agree on the *shape*
                  of everything they both declare, not merely on the names:
                  a method present in both `Workbook` (or `WorkbookCtor`)
                  interfaces declares the same return type, and a record
                  type present in both declares the same field set, with
                  optionality counted as part of the field. A type declared
                  on one surface only needs an entry in
                  `_DTS_SURFACE_ONLY_TYPES`; a deliberately divergent return
                  type needs one in `_DTS_RETURN_TYPE_EXEMPT_METHODS`.
  pure-js-helpers Helpers with no native entry point behind them
                  (`NODE_PURE_JS_FREE_FUNCTIONS`) are exported by both npm
                  packages (packages/npm/common.mjs, re-exported by each
                  entry point, and packages/npm-native/index.mjs), declared
                  in both declaration files, and implemented by the canonical
                  helper source. The native entry point must explicitly
                  re-export that source.
  readme-counts   The instance-method count quoted in
                  packages/npm-native/README.md matches the actual count
                  registered in workbook_class.cc and its shared/WASM-only
                  projection is current.
  dts-enums       Every `export enum` in src/wasm/formulon.d.ts that
                  mirrors a C/C++ enum (embind passes these through as
                  plain numbers rather than registering a real
                  `enum_<T>`, so the `.d.ts` copy is the only place the
                  ordinal values live on the JS side) matches its
                  source enum's ordinal sequence. The frozen ordinal
                  tables the two published ESM entry points export
                  (packages/npm/common.mjs via index.mjs / threads.mjs, and
                  packages/npm-native/index.mjs) must then carry the same
                  names and the same values as each other, as that
                  canonical `.d.ts`, and as the declaration file their own
                  package ships -- swapping one JS package for the other
                  is only safe if a named constant means the same thing in
                  both.
  style-record-fields
                  Every public type that projects a style record
                  (`ColorSpec` / `FontRecord` / `FillRecord` /
                  `BorderSide`) carries the same field set as its C ABI
                  struct, across the WASM `.d.ts`, the Node `.d.ts` and
                  the Python dataclasses. Deliberate omissions live in
                  `_STYLE_RECORD_EXEMPT_TYPES`.
  abi-baseline    Every entry point in the released C ABI surface
                  (tools/dev/c_abi_baseline.txt) is still declared in
                  src/c_api/formulon_c.h with the same signature, unless
                  the divergence is recorded in tools/dev/c_abi_breaks.txt.
                  Five of the eight base/`_ex` families are reached only
                  through their `_ex` variant, so deleting or renaming the
                  base leaves all four surfaces green and the break ships
                  silently -- this is the only artifact that notices. A
                  stale ledger entry fails too, so the ledger cannot
                  outlive the break it excuses.
  staged-dist     packages/npm-native/dist/{index.d.ts,index.mjs} are
                  the expected staged copies of the package-root sources:
                  `index.d.ts` and `common.mjs` are byte-identical, while
                  `index.mjs` has exactly its source-relative common import
                  rewritten to the staged sibling. This prevents a published
                  declaration, constant table or shim from going stale.
                  Reports "SKIPPED" when the package has not been staged,
                  because `dist/` is gitignored and absent in a fresh clone.
  header-error-codes
                  Every numeric status code a doc comment in
                  src/c_api/formulon_c.h prints beside an enumerator name
                  equals that enumerator's value in src/utils/error.h.
                  `fm_status_t` is a bare `int32_t`, so those numbers are
                  the whole contract a binding author reading only the
                  header can compare against.
  all             Run every check above (default).

Stdlib only; no build artifacts or network access required.
"""

from __future__ import annotations

import argparse
import sys
from typing import List

from binding_drift.abi_checks import check_abi_baseline, check_header_error_codes
from binding_drift.c_header import c_abi_declarations  # noqa: F401  (re-exported for gen_c_abi_baseline.py)
from binding_drift.declaration_checks import (
    check_dts_enums,
    check_dts_node,
    check_dts_wasm,
    check_pure_js_helpers,
    check_readme_counts,
    check_staged_dist,
)
from binding_drift.python_checks import (
    check_python_call_signatures,
    check_python_exports,
    check_python_inline_structs,
    check_python_struct_layouts,
)
from binding_drift.shape_checks import check_dts_shared_shapes, check_style_record_fields
from binding_drift.surface_files import _SKIPPED

CHECKS = {
    "python-exports": check_python_exports,
    "python-struct-layouts": check_python_struct_layouts,
    "python-call-signatures": check_python_call_signatures,
    "python-inline-structs": check_python_inline_structs,
    "dts-wasm": check_dts_wasm,
    "dts-node": check_dts_node,
    "dts-shared-shapes": check_dts_shared_shapes,
    "pure-js-helpers": check_pure_js_helpers,
    "readme-counts": check_readme_counts,
    "dts-enums": check_dts_enums,
    "style-record-fields": check_style_record_fields,
    "staged-dist": check_staged_dist,
    "abi-baseline": check_abi_baseline,
    "header-error-codes": check_header_error_codes,
}


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "check",
        nargs="?",
        default="all",
        choices=[*CHECKS.keys(), "all"],
        help="which drift check to run (default: all)",
    )
    args = parser.parse_args(argv)

    names = list(CHECKS.keys()) if args.check == "all" else [args.check]
    _SKIPPED.clear()
    problems: List[str] = []
    for name in names:
        problems.extend(CHECKS[name]())

    # Printed either way: a skip is not a pass, and the reader has to be able
    # to tell which parts of the run actually looked at anything.
    for note in _SKIPPED:
        print(f"check_binding_drift ({args.check}): SKIPPED {note}")

    if problems:
        print(f"check_binding_drift ({args.check}): DRIFT DETECTED")
        for problem in problems:
            print(problem)
        return 1

    print(f"check_binding_drift ({args.check}): no drift detected")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
