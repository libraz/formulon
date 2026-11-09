"""Drift checks across the WASM and N-API `.d.ts` surfaces, the npm helpers and the staged package."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from .api_conventions import (
    LEDGER_PATH,
    SHEET_HANDLE_CLASS,
    Ledger,
    c_verb,
    entry_handle,
    handle_classes,
    handle_name,
    load_ledger,
)
from .api_conventions_bindings import FACADE_EXTRA, OOXML_CLASSES, SHEET_CLASS, WORKBOOK_CLASS
from .enum_decl import (
    _cpp_enum_class_values,
    _extract_c_typedef_enum,
    _extract_cpp_enum_class,
    _extract_ts_enum,
    _parse_ts_named_enum,
)
from .surface_files import (
    _SKIPPED,
    CALC_MODE_H,
    CAPI_HEADER,
    CF_MATCH_H,
    EXTERNAL_LINKS_H,
    NODE_ADDON_CC,
    NODE_DIST_DIR,
    NODE_DTS,
    NODE_INDEX_MJS,
    NODE_README,
    NODE_STAGE_MJS,
    NODE_WORKBOOK_CLASS_CC,
    NPM_COMMON_MJS,
    NPM_ENTRY_MJS,
    REPO_ROOT,
    VALUE_H,
    WASM_BINDINGS_CPP,
    WASM_DTS,
    _format_diff,
    _read,
    _read_bytes,
)
from .ts_decl import (
    _extract_braced_block,
    _extract_ts_callable_members,
    _extract_ts_methods,
    _find_interface_body,
    _js_common_imports,
    _js_common_namespaces,
    _js_common_wired_function_source,
    _js_exported_function_source,
    _parse_js_constants,
)

# Intentional binding differences are kept exact. This makes a newly added
# method or a removed method fail drift checking instead of becoming a stale
# exception in one binding.
WASM_ONLY_METHODS: set = set()
NODE_ONLY_METHODS: set = set()

# Members of the `Formulon` interface that are intentionally NOT backed by a
# native `exports.Set(...)` / embind `function(...)` registration -- host-side
# helpers implemented once in packages/npm/common.mjs (e.g. the
# function-metadata provider merge helper). Excluded before comparing the
# d.ts free-function surface against the registrations.
NODE_PURE_JS_FREE_FUNCTIONS = {"mergeFunctionMetadata", "toNative"}

# `Formulon` members `buildFormulon` derives from a registered runtime member of another name
# (`evaluateTypedFormula` tags the result of the raw `evaluateFormula`); nothing registers them.
FACADE_FREE_FUNCTIONS = {"evaluateTypedFormula"}

# `Workbook` methods the pure-JS layer adds (`installWorkbookExtensions`, or the
# `makeFacade` classes); no registration stands behind them in either runtime.
PURE_JS_WORKBOOK_METHODS = {"setValue", "sheet"}

# Read-only `Workbook` properties, registered as an accessor rather than a method.
WORKBOOK_ACCESSORS = {"disposed"}

# Read-only `Workbook` properties the facade builds; no runtime registers them.
FACADE_WORKBOOK_ACCESSORS = {"ooxml"}

# Addon exports that are internal wiring and not part of the `Formulon` interface.
NODE_ADDON_INTERNAL_EXPORTS = {"__setErrorClass"}

# The interface `createFormulon` resolves to; its methods are the free functions.
ENTRY_INTERFACE = "Formulon"

_READONLY_FIELD_RE = re.compile(r"^\s*readonly\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*:", re.MULTILINE)


def _optional_interface_body(text: str, name: str) -> str:
    """The body of `export interface <name>`, or "" when the file does not declare it."""
    match = re.search(r"export interface %s(?:\s+extends\s+[^{]+?)?\s*\{" % re.escape(name), text)
    return _find_interface_body(text, name, Path(name)) if match else ""


def _ledger() -> Ledger:
    ledger = load_ledger(LEDGER_PATH)
    return ledger if not ledger.problems else Ledger()


def _js_names(ledger: Ledger, keys) -> Set[str]:
    out: Set[str] = set()
    for key in keys:
        entry = ledger.entry(key)
        was = entry.get("was") if isinstance(entry.get("was"), dict) else {}
        out.update(n for n in (entry.get("js"), was.get("js")) if isinstance(n, str))
    return out


def _pascal(snake: str) -> str:
    return "".join(w[:1].upper() + w[1:] for w in snake.split("_") if w)


def facade_split(ledger: Ledger) -> Tuple[Set[str], Set[str]]:
    """`(raw_only, facade_only)` JS method names the ledger implies.

    Raw-only names stay registered while the published declarations drop them: the count and element getter
    of a `list` / `handles` / `property` collection, every `drop:` JS name, and the raw name of an `internal`
    op or list or of an op a handle class publishes. Facade-only names are declared
    but built in JS: `FACADE_EXTRA`, the `list<Plural>()` of a `list` / `handles` collection, the
    `get<Plural>()` of a `property` collection, every `kind: list` entry, every `facade: handwritten` op
    without a C target, and the `Sheet` handle methods (`rename`, `move`).
    """
    raw_only: Set[str] = set()
    facade_only: Set[str] = set(PURE_JS_WORKBOOK_METHODS)
    for members in FACADE_EXTRA["js"].values():
        facade_only |= members
    for coll in ledger.collections.values():
        access = coll.get("access") if isinstance(coll, dict) else None
        if access not in ("list", "handles", "property") or not isinstance(coll.get("plural"), str):
            continue
        getters = [
            op
            for op in coll.get("ops") or []
            if isinstance(ledger.entry(op).get("c"), str)
            and c_verb(ledger.entry(op)["c"]) == "get"
            and not ledger.entry(op)["c"].endswith("_at")
        ]
        raw_only |= _js_names(ledger, getters)
        prefix = "get" if access == "property" else "list"
        facade_only.add(prefix + _pascal(coll["plural"]))
    for entry in ledger.entries.values():
        drop = entry.get("drop") if isinstance(entry, dict) else None
        if isinstance(drop, dict) and isinstance(drop.get("js"), str):
            raw_only.add(drop["js"])
        if isinstance(entry, dict) and (entry.get("internal") is True or entry_handle(entry) is not None):
            if isinstance(entry.get("js"), str):
                raw_only.add(entry["js"])
            # A `Sheet` handle method is declared on `Sheet` and forwards to a raw of another name.
            name = handle_name(entry, "js", ledger.collections)
            if name and SHEET_HANDLE_CLASS in handle_classes(entry):
                facade_only.add(name)
            continue
        if _facade_built(entry) and isinstance(entry.get("js"), str):
            facade_only.add(entry["js"])
    return raw_only, facade_only


def handle_declaration_problems(
    check: str, ledger: Ledger, registered: Set[str], dts_text: str, dts: Path
) -> List[str]:
    """Each handle method is declared on `interface <class>`, and the raw method it forwards to is registered."""
    problems: List[str] = []
    for key, entry in sorted(ledger.entries.items()):
        if not isinstance(entry, dict) or entry_handle(entry) is None or entry.get("kind") not in ("op", "list"):
            continue
        surfaces = entry.get("surfaces")
        if isinstance(surfaces, list) and not {"wasm", "native"} & set(surfaces):
            continue
        drop = entry.get("drop")
        if isinstance(drop, dict) and drop.get("js") is not None and drop.get("js") == entry.get("js"):
            continue
        name = handle_name(entry, "js", ledger.collections)
        handle = entry_handle(entry) or {}
        for cls in handle_classes(entry):
            body = _optional_interface_body(dts_text, cls)
            # An identity property is a readonly field of the class; every other handle op is a method.
            declared = (
                set(_READONLY_FIELD_RE.findall(body)) if handle.get("property") is True else _extract_ts_methods(body)
            )
            if name and name not in declared:
                problems.append(f"{check}: {key}: `{cls}.{name}` is not declared in {dts.relative_to(REPO_ROOT)}")
        raw = entry.get("js")
        if not _facade_built(entry) and isinstance(raw, str) and raw not in registered:
            problems.append(f"{check}: {key}: the raw `{raw}` a handle method forwards to is not registered")
    return problems


def _facade_built(entry) -> bool:
    """A `kind: list` entry, or a `facade: handwritten` op with no C target: no raw class registers it."""
    if not isinstance(entry, dict):
        return False
    if entry.get("kind") == "list":
        return True
    return entry.get("kind") == "op" and entry.get("facade") == "handwritten" and not isinstance(entry.get("c"), str)


def _declared_instance_methods(dts_text: str, source: Path) -> Set[str]:
    """Methods of the `Workbook` and `Sheet` interfaces and of their `ooxml` layers; the sheet-level ones omit
    the raw sheet argument."""
    methods = _extract_ts_methods(_find_interface_body(dts_text, WORKBOOK_CLASS, source))
    for cls in (SHEET_CLASS, *OOXML_CLASSES.values()):
        methods |= _extract_ts_methods(_optional_interface_body(dts_text, cls))
    return methods


def _registration_problems(
    check: str,
    label: str,
    registered: Set[str],
    declared: Set[str],
    allowed: Tuple[Set[str], Set[str]],
    reg_label: str,
    dts_label: str,
) -> List[str]:
    """A registration missing from the declarations, or a declaration nothing registers, outside `allowed`
    (`(raw_only, facade_only)`)."""
    raw_only, facade_only = allowed
    diff = _format_diff(reg_label, registered - declared - raw_only, dts_label, declared - registered - facade_only)
    return [f"{check}: {label} mismatch:\n" + "\n".join(diff)] if diff else []


# ---------------------------------------------------------------------------
# Check 2a: WASM embind registration <-> src/wasm/formulon.d.ts.
# ---------------------------------------------------------------------------


def check_dts_wasm() -> List[str]:
    problems: List[str] = []
    cpp_text = _read(WASM_BINDINGS_CPP)

    instance_cpp = set(re.findall(r'\.function\("([^"]+)"', cpp_text))
    static_cpp = set(re.findall(r'\.class_function\("([^"]+)"', cpp_text))
    free_cpp = set(re.findall(r'(?<!\.)\bfunction\("([^"]+)"', cpp_text))

    dts_text = _read(WASM_DTS)
    split = facade_split(_ledger())
    instance_dts = _declared_instance_methods(dts_text, WASM_DTS)
    static_dts = _extract_ts_methods(_find_interface_body(dts_text, "WorkbookCtor", WASM_DTS))
    free_dts = _extract_ts_methods(_find_interface_body(dts_text, ENTRY_INTERFACE, WASM_DTS))
    free_dts -= NODE_PURE_JS_FREE_FUNCTIONS
    accessor_cpp = set(re.findall(r'\.property\("([^"]+)"', cpp_text))
    accessor_dts = set(_READONLY_FIELD_RE.findall(_find_interface_body(dts_text, "Workbook", WASM_DTS)))

    problems += handle_declaration_problems("dts-wasm", _ledger(), instance_cpp, dts_text, WASM_DTS)
    for label, cpp_set, dts_set, allowed in (
        ("Workbook / Sheet instance methods", instance_cpp, instance_dts, split),
        ("Workbook static factories", static_cpp, static_dts, (set(), set())),
        ("free functions", free_cpp, free_dts, (set(), FACADE_FREE_FUNCTIONS)),
        ("Workbook accessors", accessor_cpp & WORKBOOK_ACCESSORS, accessor_dts, (set(), FACADE_WORKBOOK_ACCESSORS)),
    ):
        problems += _registration_problems(
            "dts-wasm",
            label,
            cpp_set,
            dts_set,
            allowed,
            f"{WASM_BINDINGS_CPP.relative_to(REPO_ROOT)}",
            f"{WASM_DTS.relative_to(REPO_ROOT)}",
        )
    if accessor_cpp & WORKBOOK_ACCESSORS != WORKBOOK_ACCESSORS:
        problems.append(
            f"dts-wasm: Workbook accessors {sorted(WORKBOOK_ACCESSORS - accessor_cpp)} are not registered "
            f"with `.property(...)` in {WASM_BINDINGS_CPP.relative_to(REPO_ROOT)}"
        )

    return problems


# ---------------------------------------------------------------------------
# Check 2b: Node N-API registration <-> packages/npm-native/index.d.ts.
# ---------------------------------------------------------------------------


def check_dts_node() -> List[str]:
    problems: List[str] = []
    cc_text = _read(NODE_WORKBOOK_CLASS_CC)

    instance_cc = set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text))
    static_cc = set(re.findall(r'StaticMethod<[^>]+>\("([^"]+)"\)', cc_text))

    addon_text = _read(NODE_ADDON_CC)
    exported_cc = set(re.findall(r'exports\.Set\("([^"]+)"', addon_text)) - {"Workbook"}
    exported_cc |= set(re.findall(r'SetGuardedModuleFunction<[^>]*>\(\s*env,\s*exports,\s*"([^"]+)"\)', cc_text))
    free_cc = exported_cc - NODE_ADDON_INTERNAL_EXPORTS
    accessor_cc = set(re.findall(r'InstanceAccessor(?:<[^>]*>)?\(\s*"([^"]+)"', cc_text))

    dts_text = _read(NODE_DTS)
    split = facade_split(_ledger())
    instance_dts = _declared_instance_methods(dts_text, NODE_DTS)
    static_dts = _extract_ts_methods(_find_interface_body(dts_text, "WorkbookCtor", NODE_DTS))
    free_dts = _extract_ts_methods(_find_interface_body(dts_text, ENTRY_INTERFACE, NODE_DTS))
    free_dts -= NODE_PURE_JS_FREE_FUNCTIONS
    accessor_dts = set(_READONLY_FIELD_RE.findall(_find_interface_body(dts_text, "Workbook", NODE_DTS)))
    missing_internal = NODE_ADDON_INTERNAL_EXPORTS - exported_cc
    if missing_internal:
        problems.append(
            f"dts-node: {NODE_ADDON_CC.relative_to(REPO_ROOT)} does not export the internal "
            f"{sorted(missing_internal)} the package entry point calls"
        )

    wasm_cpp_text = _read(WASM_BINDINGS_CPP)
    wasm_instance = set(re.findall(r'\.function\("([^"]+)"', wasm_cpp_text))
    actual_wasm_only = wasm_instance - instance_cc
    actual_node_only = instance_cc - wasm_instance
    if actual_wasm_only != WASM_ONLY_METHODS:
        problems.append(
            "dts-node: WASM-only instance-method allowlist mismatch; "
            f"actual={sorted(actual_wasm_only)}, allowlist={sorted(WASM_ONLY_METHODS)}"
        )
    if actual_node_only != NODE_ONLY_METHODS:
        problems.append(
            "dts-node: Node-only instance-method allowlist mismatch; "
            f"actual={sorted(actual_node_only)}, allowlist={sorted(NODE_ONLY_METHODS)}"
        )
    stale_wasm_only = WASM_ONLY_METHODS - actual_wasm_only
    stale_node_only = NODE_ONLY_METHODS - actual_node_only
    if stale_wasm_only:
        problems.append(f"dts-node: stale WASM-only allowlist entries: {sorted(stale_wasm_only)}")
    if stale_node_only:
        problems.append(f"dts-node: stale Node-only allowlist entries: {sorted(stale_node_only)}")

    problems += handle_declaration_problems("dts-node", _ledger(), instance_cc, dts_text, NODE_DTS)
    for label, cc_set, dts_set, allowed in (
        ("Workbook / Sheet instance methods", instance_cc, instance_dts, split),
        ("Workbook static factories", static_cc, static_dts, (set(), set())),
        ("free functions", free_cc, free_dts, (set(), FACADE_FREE_FUNCTIONS)),
        ("Workbook accessors", accessor_cc & WORKBOOK_ACCESSORS, accessor_dts, (set(), FACADE_WORKBOOK_ACCESSORS)),
    ):
        problems += _registration_problems(
            "dts-node",
            label,
            cc_set,
            dts_set,
            allowed,
            f"{NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} / {NODE_ADDON_CC.relative_to(REPO_ROOT)}",
            f"{NODE_DTS.relative_to(REPO_ROOT)}",
        )
    if accessor_cc & WORKBOOK_ACCESSORS != WORKBOOK_ACCESSORS:
        problems.append(
            f"dts-node: Workbook accessors {sorted(WORKBOOK_ACCESSORS - accessor_cc)} are not registered "
            f"with `InstanceAccessor` in {NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)}"
        )

    return problems


_COMMON_REEXPORT_RE = re.compile(r"^export (?:\*|\{[^}]*\}) from '\./common\.js';", re.MULTILINE)


def _npm_entries_reexport_common(check: str) -> List[str]:
    """Every WASM npm entry point must re-export from common.mjs."""
    return [
        f"{check}: {entry.relative_to(REPO_ROOT)} does not re-export from './common.js'"
        for entry in NPM_ENTRY_MJS
        if not _COMMON_REEXPORT_RE.search(_read(entry))
    ]


def check_pure_js_helpers() -> List[str]:
    problems: List[str] = _npm_entries_reexport_common("pure-js-helpers")
    sources = {
        "npm": (NPM_COMMON_MJS, WASM_DTS),
        "npm-native": (NODE_INDEX_MJS, NODE_DTS),
    }
    common_text = _read(NPM_COMMON_MJS)

    for name in sorted(NODE_PURE_JS_FREE_FUNCTIONS):
        bodies: dict[str, str] = {}
        for package, (mjs_path, dts_path) in sources.items():
            if package == "npm-native":
                body = _js_common_wired_function_source(mjs_path, name)
            else:
                body = _js_exported_function_source(_read(mjs_path), name)
            if body is None:
                if package == "npm-native":
                    problems.append(
                        f"pure-js-helpers: {mjs_path.relative_to(REPO_ROOT)} must import ../npm/common.mjs "
                        f"as a namespace, call `buildFormulon`, and common.mjs must list `{name}` "
                        "in FORMULON_MEMBERS"
                    )
                else:
                    problems.append(
                        f"pure-js-helpers: {mjs_path.relative_to(REPO_ROOT)} does not export a "
                        f"`{name}` function. A helper only one package ships cannot be named in a "
                        "declaration file both packages publish."
                    )
            else:
                bodies[package] = body
            if name not in _extract_ts_callable_members(
                _find_interface_body(_read(dts_path), ENTRY_INTERFACE, dts_path)
            ):
                problems.append(
                    f"pure-js-helpers: {dts_path.relative_to(REPO_ROOT)} does not declare `{ENTRY_INTERFACE}.{name}`"
                )
        if len(bodies) == len(sources) and len(set(bodies.values())) != 1:
            problems.append(
                f"pure-js-helpers: the `{name}` implementations have diverged between "
                f"{NPM_COMMON_MJS.relative_to(REPO_ROOT)} and {NODE_INDEX_MJS.relative_to(REPO_ROOT)}. "
                "The native entry point must import the canonical helper from "
                "packages/npm/common.mjs so both packages stay tied to one implementation."
            )

    if _MAKE_FACADE_RE.search(common_text):
        return problems + _facade_wiring_problems(common_text, _read(NODE_INDEX_MJS))
    extensions = _js_exported_function_source(common_text, "installWorkbookExtensions")
    if extensions is None:
        problems.append(
            f"pure-js-helpers: {NPM_COMMON_MJS.relative_to(REPO_ROOT)} does not export `installWorkbookExtensions`"
        )
    else:
        for method in sorted(PURE_JS_WORKBOOK_METHODS):
            if not re.search(r"\b%s\b" % re.escape(method), extensions):
                problems.append(f"pure-js-helpers: `installWorkbookExtensions` does not add `Workbook.{method}`")
    native_text = _read(NODE_INDEX_MJS)
    installed = "installWorkbookExtensions" in _js_common_imports(native_text).values() or any(
        re.search(r"\b%s\.installWorkbookExtensions\s*\(" % re.escape(ns), native_text)
        for ns in _js_common_namespaces(native_text)
    )
    if not installed:
        problems.append(
            f"pure-js-helpers: {NODE_INDEX_MJS.relative_to(REPO_ROOT)} must take "
            "`installWorkbookExtensions` from ../npm/common.mjs"
        )

    return problems


_MAKE_FACADE_RE = re.compile(r"^(?:export\s+)?function\s+makeFacade\s*\(", re.MULTILINE)
_RUNTIME_MEMBERS_RE = re.compile(r"\bRUNTIME_MEMBERS\s*=\s*new Set\(\[(.*?)\]\)", re.DOTALL)


def _function_body(text: str, name: str) -> str:
    match = re.search(r"^(?:export\s+)?function\s+%s\s*\(" % re.escape(name), text, re.MULTILINE)
    if not match:
        return ""
    return _extract_braced_block(text, text.index("{", match.end()) + 1)


def _facade_wiring_problems(common_text: str, native_text: str) -> List[str]:
    """The facade wiring: `buildFormulon` hands the raw `Workbook` to `makeFacade`, `RUNTIME_MEMBERS` no longer
    copies the raw class onto the module object, `installWorkbookExtensions` is gone, and the native entry
    passes its raw runtime to `buildFormulon` / `loadFormulon` without installing extensions.
    """
    common_label = NPM_COMMON_MJS.relative_to(REPO_ROOT)
    native_label = NODE_INDEX_MJS.relative_to(REPO_ROOT)
    problems: List[str] = []
    if re.search(r"\bfunction\s+installWorkbookExtensions\b", common_text):
        problems.append(f"pure-js-helpers: {common_label} defines `makeFacade` and still `installWorkbookExtensions`")
    if "makeFacade(" not in _function_body(common_text, "buildFormulon"):
        problems.append(f"pure-js-helpers: `buildFormulon` in {common_label} does not call `makeFacade`")
    runtime = _RUNTIME_MEMBERS_RE.search(common_text)
    if runtime is not None and re.search(r"['\"]Workbook['\"]", runtime.group(1)):
        problems.append(f"pure-js-helpers: `RUNTIME_MEMBERS` in {common_label} still copies the raw `Workbook`")
    body = _function_body(common_text, "makeFacade")
    for method in sorted(PURE_JS_WORKBOOK_METHODS):
        if not re.search(r"\b%s\b" % re.escape(method), body):
            problems.append(f"pure-js-helpers: `makeFacade` does not build `{method}`")
    if "installWorkbookExtensions" in native_text:
        problems.append(f"pure-js-helpers: {native_label} still calls `installWorkbookExtensions`")
    wired = any(
        re.search(r"\b%s\.(?:buildFormulon|loadFormulon)\s*\(" % re.escape(ns), native_text)
        for ns in _js_common_namespaces(native_text)
    )
    if not wired:
        problems.append(
            f"pure-js-helpers: {native_label} must pass its runtime to `buildFormulon` from ../npm/common.mjs"
        )
    return problems


# ---------------------------------------------------------------------------
# Check 3: README instance-method count <-> actual registration count.
# ---------------------------------------------------------------------------

_NUMBER_WORDS = {
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
}

_README_COUNT_RE = re.compile(r"register (\d+) instance methods plus the\s+(\w+) static factories")
_README_SHARED_COUNT_RE = re.compile(r"instance methods, (\d+) are shared with WASM; (\w+) remain WASM-only")


def check_readme_counts() -> List[str]:
    problems: List[str] = []
    cc_text = _read(NODE_WORKBOOK_CLASS_CC)
    actual_instance = len(set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text)))
    actual_static = len(set(re.findall(r'StaticMethod<[^>]+>\("([^"]+)"\)', cc_text)))

    readme_text = _read(NODE_README)
    match = _README_COUNT_RE.search(readme_text)
    if not match:
        problems.append(
            f"readme-counts: could not find the instance/static method-count sentence in "
            f"{NODE_README.relative_to(REPO_ROOT)} (pattern: {_README_COUNT_RE.pattern!r})"
        )
        return problems

    quoted_instance = int(match.group(1))
    quoted_static_word = match.group(2).lower()
    quoted_static = _NUMBER_WORDS.get(quoted_static_word)

    if quoted_instance != actual_instance:
        problems.append(
            f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {quoted_instance} instance methods, "
            f"actual count in {NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} is {actual_instance}"
        )
    if quoted_static is None or quoted_static != actual_static:
        problems.append(
            f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {match.group(2)!r} static factories, "
            f"actual count in {NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} is {actual_static}"
        )

    wasm_cpp_text = _read(WASM_BINDINGS_CPP)
    wasm_instance = set(re.findall(r'\.function\("([^"]+)"', wasm_cpp_text))
    actual_node_methods = set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text))
    actual_shared = len(wasm_instance & actual_node_methods)
    actual_wasm_only = len(wasm_instance - actual_node_methods)
    shared_match = _README_SHARED_COUNT_RE.search(readme_text)
    if not shared_match:
        problems.append(
            f"readme-counts: could not find the shared/WASM-only method-count sentence in "
            f"{NODE_README.relative_to(REPO_ROOT)}"
        )
    else:
        quoted_shared = int(shared_match.group(1))
        quoted_wasm_only = _NUMBER_WORDS.get(shared_match.group(2).lower())
        if quoted_shared != actual_shared:
            problems.append(
                f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {quoted_shared} shared methods, "
                f"actual shared count is {actual_shared}"
            )
        if quoted_wasm_only is None or quoted_wasm_only != actual_wasm_only:
            problems.append(
                f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {shared_match.group(2)!r} WASM-only methods, "
                f"actual count is {actual_wasm_only}"
            )

    return problems


# Maps each `.d.ts` enum name to the (source description, extractor) pair
# that finds its mirrored C/C++ enum's ordinal sequence. `fm_*_t` names
# resolve against formulon_c.h; the remaining few are plain C++ `enum
# class` types in their own small headers (comparing against the giant
# formulon_c.h text for those would just add false-positive risk from an
# unrelated same-named enum elsewhere).
_DTS_ENUM_SOURCES = {
    "ValueKind": (CAPI_HEADER, "fm_value_kind_t", _extract_c_typedef_enum),
    "WorkbookFormat": (CAPI_HEADER, "fm_workbook_format_t", _extract_c_typedef_enum),
    "PivotCellKind": (CAPI_HEADER, "fm_pivot_cell_kind_t", _extract_c_typedef_enum),
    "PivotAxis": (CAPI_HEADER, "fm_pivot_axis_t", _extract_c_typedef_enum),
    "PivotAggregation": (CAPI_HEADER, "fm_pivot_aggregation_t", _extract_c_typedef_enum),
    "PivotShowValuesAs": (CAPI_HEADER, "fm_pivot_show_values_as_t", _extract_c_typedef_enum),
    "PivotFilterKind": (CAPI_HEADER, "fm_pivot_filter_kind_t", _extract_c_typedef_enum),
    "PivotCalendar": (CAPI_HEADER, "fm_pivot_calendar_t", _extract_c_typedef_enum),
    "PivotSortOrder": (CAPI_HEADER, "fm_pivot_sort_order_t", _extract_c_typedef_enum),
    "PivotReportLayout": (CAPI_HEADER, "fm_pivot_report_layout_t", _extract_c_typedef_enum),
    "SheetVisibility": (CAPI_HEADER, "fm_sheet_visibility_t", _extract_c_typedef_enum),
    "CollationStrength": (CAPI_HEADER, "fm_collation_strength_t", _extract_c_typedef_enum),
    "ParseErrorCode": (CAPI_HEADER, "fm_parse_error_code_t", _extract_c_typedef_enum),
    "CfRuleType": (CAPI_HEADER, "fm_cf_rule_type_t", _extract_c_typedef_enum),
    "CfValueObjectType": (CAPI_HEADER, "fm_cf_value_object_type_t", _extract_c_typedef_enum),
    "IconSetName": (CAPI_HEADER, "fm_icon_set_name_t", _extract_c_typedef_enum),
    "ValidationType": (CAPI_HEADER, "fm_validation_type_t", _extract_c_typedef_enum),
    "CfMatchKind": (CF_MATCH_H, "CFMatchKind", _extract_cpp_enum_class),
    "CalcMode": (CALC_MODE_H, "CalcMode", _extract_cpp_enum_class),
    "RefMode": (CALC_MODE_H, "RefMode", _extract_cpp_enum_class),
    "ErrorCode": (VALUE_H, "ErrorCode", _extract_cpp_enum_class),
    "ExternalLinkKind": (EXTERNAL_LINKS_H, "Kind", _extract_cpp_enum_class),
}


def _enum_sources(dts_text: str) -> Dict[str, tuple]:
    """`_DTS_ENUM_SOURCES`, plus every `.d.ts` table whose ledger enum names a C typedef enum of the header.

    A table added for a new C enum is therefore checked without a hand-kept entry.
    """
    sources = dict(_DTS_ENUM_SOURCES)
    declared = set(re.findall(r"^export const (\w+): Readonly<\{", dts_text, re.MULTILINE))
    header = _read(CAPI_HEADER)
    for entry in _ledger().entries.values():
        if not isinstance(entry, dict) or entry.get("kind") != "enum":
            continue
        name, c_name = entry.get("js"), entry.get("c")
        if name in declared and name not in sources and isinstance(c_name, str):
            if _extract_c_typedef_enum(header, c_name) is not None:
                sources[name] = (CAPI_HEADER, c_name, _extract_c_typedef_enum)
    return sources


def _present(values: List[int]) -> List[int]:
    """Ordinals without `-1`: an "absent" member is `-1` in C and `null` in the binding, never a table member."""
    return [v for v in values if v != -1]


# C members that only instruct a patch, keyed by the `.d.ts` table that leaves them out: the binding spells the
# instruction as a `null` record (`title: null` clears the title), so the table never carries the member.
_PATCH_ONLY_ORDINALS = {
    "ChartTitleKind": frozenset({0}),
}


def check_dts_enums() -> List[str]:
    problems: List[str] = []
    dts_text = _read(WASM_DTS)

    for ts_name, (source_path, source_enum, extractor) in sorted(_enum_sources(dts_text).items()):
        ts_values = _extract_ts_enum(dts_text, ts_name)
        if ts_values is None:
            problems.append(
                f"dts-enums: could not find/parse the `{ts_name}` table in {WASM_DTS.relative_to(REPO_ROOT)}"
            )
            continue

        source_text = _read(source_path)
        source_values = extractor(source_text, source_enum)
        if source_values is None:
            problems.append(
                f"dts-enums: could not find/parse `{source_enum}` in {source_path.relative_to(REPO_ROOT)} "
                f"(needed to check {WASM_DTS.relative_to(REPO_ROOT)}'s {ts_name})"
            )
            continue

        patch_only = _PATCH_ONLY_ORDINALS.get(ts_name, frozenset())
        stale = sorted(v for v in patch_only if v not in source_values or v in ts_values)
        if stale:
            problems.append(
                f"dts-enums: {ts_name} lists patch-only ordinals {stale} that {source_enum} lacks or the table carries"
            )
        if _present(ts_values) != [v for v in _present(source_values) if v not in patch_only]:
            problems.append(
                f"dts-enums: {ts_name} in {WASM_DTS.relative_to(REPO_ROOT)} has ordinals {ts_values}, "
                f"but its source {source_enum} in {source_path.relative_to(REPO_ROOT)} has {source_values}"
            )

    problems.extend(_check_js_constant_tables(dts_text))
    common_text, value_h_text = _read(NPM_COMMON_MJS), _read(VALUE_H)
    problems.extend(check_string_tables(common_text, dts_text, value_h_text))
    problems.extend(check_string_tables(common_text, _read(NODE_DTS), value_h_text, NODE_DTS.name))
    return problems


# String tables keyed by an enum's members: `export const X = Object.freeze({ [Enum.Member]: '...', ... })` at
# runtime, `export const X: Readonly<Record<Enum, string>>;` in the declarations.
_JS_STRING_TABLE_RE = re.compile(r"^export const (\w+) = Object\.freeze\(\{(.*?)\}\);", re.DOTALL | re.MULTILINE)
_JS_STRING_MEMBER_RE = re.compile(r"\[(\w+)\.(\w+)\]\s*:\s*(?:'([^']*)'|\"([^\"]*)\")")
_TS_STRING_TABLE_RE = re.compile(r"^export const (\w+)\s*:\s*Readonly<Record<(\w+),\s*string>>\s*;", re.MULTILINE)
_ERROR_TABLE_RE = re.compile(r"kErrorTable\[\]\s*=\s*\{(.*?)\n\};", re.DOTALL)


def _js_string_tables(text: str) -> Dict[str, Tuple[str, List[Tuple[str, str]]]]:
    """`{name: (enum, [(member, text), ...])}` of the frozen tables whose every entry is `[Enum.Member]: '...'`."""
    out: Dict[str, Tuple[str, List[Tuple[str, str]]]] = {}
    for match in _JS_STRING_TABLE_RE.finditer(text):
        entries = [e for e in match.group(2).split(",\n") if e.strip()]
        members = _JS_STRING_MEMBER_RE.findall(match.group(2))
        enums = {enum for enum, _, _, _ in members}
        if members and len(members) == len(entries) and len(enums) == 1:
            out[match.group(1)] = (enums.pop(), [(member, a or b) for _, member, a, b in members])
    return out


def string_table_names(text: str) -> Set[str]:
    return set(_js_string_tables(text))


def _error_display_names(value_h_text: str) -> Optional[List[Tuple[str, str]]]:
    """`(ErrorCode member, display name)` in `kErrorTable` order, or None when either is unreadable."""
    members = _cpp_enum_class_values(value_h_text, "ErrorCode")
    table = _ERROR_TABLE_RE.search(value_h_text)
    if members is None or table is None:
        return None
    names = re.findall(r'\{\s*"([^"]*)"\s*,', table.group(1))
    return list(zip(members, names)) if len(names) == len(members) else None


# The source of the texts of a string table, by the enum that keys it.
_STRING_TABLE_SOURCES = {"ErrorCode": ("kErrorTable", _error_display_names)}


def check_string_tables(common_text: str, dts_text: str, value_h_text: str, dts_name: str = WASM_DTS.name) -> List[str]:
    """Each `Readonly<Record<Enum, string>>` the declarations type has a runtime table in `common.mjs` keyed by every
    member of `Enum` in order, with the texts of its source table; and every runtime string table is declared."""
    problems: List[str] = []
    runtime = _js_string_tables(common_text)
    declared = dict(_TS_STRING_TABLE_RE.findall(dts_text))
    for name in sorted(set(runtime) - set(declared)):
        problems.append(f"dts-enums: common.mjs exports the string table {name}, which {dts_name} does not declare")
    for name, enum in sorted(declared.items()):
        if name not in runtime:
            problems.append(f"dts-enums: {dts_name} declares the string table {name}, which common.mjs does not export")
            continue
        keyed_by, rows = runtime[name]
        members = _parse_ts_named_enum(dts_text, enum)
        if keyed_by != enum or members is None or [m for m, _ in rows] != list(members):
            problems.append(
                f"dts-enums: {name} in common.mjs is keyed by {keyed_by} members {[m for m, _ in rows]}, "
                f"not by every {enum} member in order {list(members or [])}"
            )
            continue
        source = _STRING_TABLE_SOURCES.get(enum)
        want = source[1](value_h_text) if source else None
        if source and want is None:
            problems.append(f"dts-enums: could not read {source[0]} to check {name}")
        elif source and rows != want:
            problems.append(
                f"dts-enums: {name} in common.mjs is {[t for _, t in rows]}, but {source[0]} has {[t for _, t in want]}"
            )
    return problems


def _check_js_constant_tables(wasm_dts_text: str) -> List[str]:
    problems: List[str] = _npm_entries_reexport_common("dts-enums")
    strings = string_table_names(_read(NPM_COMMON_MJS))
    npm = {k: v for k, v in _parse_js_constants(NPM_COMMON_MJS).items() if k not in strings}
    native = {k: v for k, v in _parse_js_constants(NODE_INDEX_MJS).items() if k not in strings}
    npm_label = NPM_COMMON_MJS.relative_to(REPO_ROOT)
    native_label = NODE_INDEX_MJS.relative_to(REPO_ROOT)

    only_npm = sorted(set(npm) - set(native))
    only_native = sorted(set(native) - set(npm))
    if only_npm:
        problems.append(f"dts-enums: exported by {npm_label} but not {native_label}: {only_npm}")
    if only_native:
        problems.append(f"dts-enums: exported by {native_label} but not {npm_label}: {only_native}")

    for name in sorted(set(npm) & set(native)):
        if npm[name] != native[name]:
            problems.append(f"dts-enums: {name} is {npm[name]} in {npm_label} but {native[name]} in {native_label}")

    # Each runtime table must also match the declaration file its own
    # package ships, and the tables (not the bare scalars) must match the
    # canonical WASM `.d.ts`.
    node_dts_text = _read(NODE_DTS)
    for label, constants, dts_text, dts_path in (
        (npm_label, npm, wasm_dts_text, WASM_DTS),
        (native_label, native, node_dts_text, NODE_DTS),
    ):
        for name, value in sorted(constants.items()):
            if not isinstance(value, dict):
                if not re.search(r"export const\s+" + re.escape(name) + r"\s*=\s*" + str(value) + r"\b", dts_text):
                    problems.append(
                        f"dts-enums: {label} exports {name} = {value}, "
                        f"not declared with that value in {dts_path.relative_to(REPO_ROOT)}"
                    )
                continue
            declared = _parse_ts_named_enum(dts_text, name)
            if declared is None:
                problems.append(
                    f"dts-enums: {label} exports the table {name}, "
                    f"which {dts_path.relative_to(REPO_ROOT)} does not declare as a value"
                )
            elif declared != value:
                problems.append(
                    f"dts-enums: {name} is {value} in {label} but {declared} in {dts_path.relative_to(REPO_ROOT)}"
                )

    canonical_only = sorted(
        name for name in _DTS_ENUM_SOURCES if name not in npm or not isinstance(npm.get(name), dict)
    )
    if canonical_only:
        problems.append(
            f"dts-enums: {WASM_DTS.relative_to(REPO_ROOT)} declares these enums but {npm_label} "
            f"does not export them as runtime tables: {canonical_only}"
        )
    return problems


# ---------------------------------------------------------------------------
# Check 7: the npm-native package's staged copies <-> their sources.
#
# `stage.mjs` publishes `packages/npm-native/dist/` by copying `index.d.ts` and
# `common.mjs`, while it rewrites the native index's source-relative import to
# the staged sibling. Every other check in this file reads the source side
# only, so an edit that lands in the source but is never re-staged publishes a
# declaration file, shared constant table or shim that no longer describes the
# runtime. The WASM package already guards its own staged `.d.ts`
# (packages/npm/scripts/check-dts.mjs); npm-native needs the same guard for the
# copied files and the deterministic import rewrite.
# ---------------------------------------------------------------------------

# (source, staged copy) pairs `stage.mjs` produces. The native index is the
# one transformed copy; `_staged_bytes` applies that exact transformation.
_NODE_STAGED_COPIES = (
    (NODE_DTS, NODE_DIST_DIR / "index.d.ts"),
    (NODE_INDEX_MJS, NODE_DIST_DIR / "index.mjs"),
    (NPM_COMMON_MJS, NODE_DIST_DIR / "common.mjs"),
)

_NATIVE_COMMON_IMPORT = b"from '../npm/common.mjs'"
_STAGED_COMMON_IMPORT = b"from './common.mjs'"


_STAGE_SITES_RE = re.compile(r"commonImportSites\s*!==\s*(\d+)")


def _expected_common_import_sites() -> int | None:
    """The import-site count `stage.mjs` insists on before it rewrites, or None when unreadable."""
    match = _STAGE_SITES_RE.search(_read(NODE_STAGE_MJS)) if NODE_STAGE_MJS.is_file() else None
    return int(match.group(1)) if match else None


def _staged_bytes(source: Path, data: bytes) -> tuple[bytes | None, str | None]:
    """Returns expected bytes and an error for a malformed native shim source."""
    if source != NODE_INDEX_MJS:
        return data, None
    expected_sites = _expected_common_import_sites()
    if expected_sites is None:
        return (
            None,
            f"staged-dist: cannot read the common import-site count from {NODE_STAGE_MJS.relative_to(REPO_ROOT)}",
        )
    sites = data.count(_NATIVE_COMMON_IMPORT)
    if sites != expected_sites:
        return None, (
            f"staged-dist: {NODE_INDEX_MJS.relative_to(REPO_ROOT)} must contain exactly {expected_sites} canonical "
            f"common imports for staging (as {NODE_STAGE_MJS.relative_to(REPO_ROOT)} rewrites), found {sites}"
        )
    return data.replace(_NATIVE_COMMON_IMPORT, _STAGED_COMMON_IMPORT), None


def check_staged_dist() -> List[str]:
    problems: List[str] = []
    staged_present = any(staged.is_file() for _, staged in _NODE_STAGED_COPIES)
    for source, staged in _NODE_STAGED_COPIES:
        if not staged.is_file():
            if not staged_present:
                # `dist/` is gitignored, so a fresh clone has nothing to
                # compare until the package is staged. Record that as
                # not-checked instead of passing; a partially staged package
                # is a real drift and must fail below.
                continue
            problems.append(
                f"staged-dist: {staged.relative_to(REPO_ROOT)} is absent from a partial staging; "
                "re-stage the package with `make node-package`"
            )
            continue
        expected, malformed = _staged_bytes(source, _read_bytes(source))
        if malformed:
            problems.append(malformed)
            continue
        if expected != _read_bytes(staged):
            problems.append(
                f"staged-dist: {staged.relative_to(REPO_ROOT)} differs from "
                f"the expected staged form of {source.relative_to(REPO_ROOT)}; the published copy is stale -- "
                "re-stage the package with `make node-package`"
            )
    if not staged_present:
        _SKIPPED.append("staged-dist: npm-native dist has not been staged in this tree (`make node-package`)")
    return problems
