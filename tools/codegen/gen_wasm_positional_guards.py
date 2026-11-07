#!/usr/bin/env python3
"""Generate the embind positional-argument guard pre-js.

The N-API binding is the canonical declaration of the primitive positional
types. Its GuardedInstanceMethod masks are parsed here instead of being
copied into a second hand-maintained table. The parser is deliberately
strict: a malformed registration, duplicate name, mask overlap, or binding
surface drift stops generation before an artifact can be built.

The generated file is passed to emscripten as --pre-js. It contains no
native code, so the guard lives in the JS glue and does not change the wasm
module. The same wrapper throws what the native code records during the call:
a nested-field rejection as TypeError/RangeError, a callback's own exception
as itself, and a failed status as `FormulonError`, as on the Node surface.
Every Workbook method, every Workbook static function and every free function
is wrapped; installation fails if one is left unwrapped.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

REPO_ROOT = Path(__file__).resolve().parents[2]

# The sheet-scope method list and the handle-method table gen_facade_tables.py writes from the ledger;
# reading the generated copy keeps the WASM build free of a YAML parser.
FACADE_SCOPE_INC = REPO_ROOT / "src/node_addon/generated/facade_scope.inc"
DEFAULT_NODE_SOURCE = REPO_ROOT / "src/node_addon/parts/workbook_class.cc"
DEFAULT_WASM_SOURCE = REPO_ROOT / "src/wasm/parts/bindings_register.cpp"
EXPECTED_NODE_REGISTRATIONS = 429

MASK_NAMES = (
    "u32",
    "i32",
    "u64",
    "double",
    "string",
    "optionalU32",
    "optionalU64",
    "optionalNull",
    "object",
    "required",
)
MASK_TOKEN = re.compile(r"0x([0-9a-fA-F]+)ULL\Z")
METHOD_MARKER = re.compile(r"\bGuardedInstanceMethod(?=\s*<)")
WASM_METHOD = re.compile(r"\.function\(\s*\"([A-Za-z0-9_]+)\"")
WASM_FREE_FUNCTION = re.compile(r"(?<![.\w])function\(\s*\"([A-Za-z0-9_]+)\"")
WASM_STATIC_FUNCTION = re.compile(r"\.class_function\(\s*\"([A-Za-z0-9_]+)\"")

# Node methods intentionally not exposed by the embind class; today there are
# none. Every method both surfaces share takes its spec from the Node
# registration. Keeping the set explicit makes a future binding drift fail at
# generation time instead of silently losing validation.
NODE_ONLY_METHODS: frozenset[str] = frozenset()

# Positional specs of the free functions that take checked primitives. Every
# other registered free function is still wrapped, with an empty spec.
FREE_FUNCTION_GUARDS = {
    "evaluateFormula": {"string": 0x1},
    "getStatusName": {"i32": 0x1},
    "setLogMinLevel": {"i32": 0x1},
    "canonicalizeFormula": {"string": 0x3},
    "canonicalizeFunctionName": {"string": 0x3},
    "compareText": {"u32": 0x4, "string": 0xB},
    "convertCivilToSerial": {"object": 0x1},
    "convertSerialToCivil": {"double": 0x1},
    "getCollationClasses": {"u32": 0x2, "string": 0x4, "object": 0x1},
    "getFunctionMetadata": {"string": 0x1},
    "getImageInfo": {"object": 0x1},
    "getLocaleFacts": {"string": 0x1},
    "listLocaleErrorNames": {"string": 0x1},
    "localizeFormula": {"string": 0x3},
    "localizeFunctionName": {"string": 0x3},
}

# Positional specs of the Workbook static functions. The set must equal the
# registered `class_function` names, so a new static cannot go unwrapped.
STATIC_FUNCTION_GUARDS = {
    "createDefault": {},
    "createEmpty": {},
    "createForProfile": {"string": 0x1},
    "load": {"object": 0x1},
}


class GuardGenerationError(ValueError):
    """An input binding source is inconsistent or could not be parsed."""


@dataclass(frozen=True)
class Registration:
    cpp_name: str
    js_name: str
    masks: tuple[int, ...]

    def as_js_spec(self) -> dict[str, int]:
        return {name: value for name, value in zip(MASK_NAMES, self.masks) if value}


def _parse_template_inner(inner: str, source: Path, marker_offset: int) -> tuple[str, tuple[int, ...]]:
    parts = [part.strip() for part in inner.split(",")]
    if not parts or not parts[0].startswith("&Workbook::"):
        raise GuardGenerationError(f"{source}:{marker_offset}: expected &Workbook::Method registration")
    cpp_name = parts[0][len("&Workbook::") :]
    if not re.fullmatch(r"[A-Za-z0-9_]+", cpp_name):
        raise GuardGenerationError(f"{source}:{marker_offset}: invalid C++ method name {cpp_name!r}")
    if len(parts) > len(MASK_NAMES) + 1:
        raise GuardGenerationError(
            f"{source}:{marker_offset}: {cpp_name} has {len(parts) - 1} masks; at most {len(MASK_NAMES)} are allowed"
        )

    masks = [0] * len(MASK_NAMES)
    for index, token in enumerate(parts[1:]):
        match = MASK_TOKEN.fullmatch(token)
        if match is None:
            raise GuardGenerationError(f"{source}:{marker_offset}: invalid positional mask {token!r}")
        masks[index] = int(match.group(1), 16)
    return cpp_name, tuple(masks)


def _find_closing_angle(text: str, opening: int, source: Path) -> int:
    # The template argument list is intentionally simple today, but balancing
    # here makes the parser fail closed if a future registration introduces a
    # nested template rather than silently truncating at its first angle.
    depth = 0
    for index in range(opening, len(text)):
        char = text[index]
        if char == "<":
            depth += 1
        elif char == ">":
            depth -= 1
            if depth == 0:
                return index
            if depth < 0:
                break
    raise GuardGenerationError(f"{source}:{opening}: unterminated GuardedInstanceMethod template")


def _parse_node_registrations(source: Path) -> list[Registration]:
    text = source.read_text(encoding="utf-8")
    try:
        registration_start = text.index("Napi::Function constructor = DefineClass")
        registration_end = text.index("});", registration_start)
    except ValueError as exc:
        raise GuardGenerationError(f"{source}: cannot locate DefineClass registration block") from exc
    registration_block = text[registration_start:registration_end]
    for marker in re.finditer(r"\bGuardedInstanceMethod\b", registration_block):
        if not re.match(r"\s*<", registration_block[marker.end() :]):
            raise GuardGenerationError(
                f"{source}:{registration_start + marker.start()}: unparsed GuardedInstanceMethod in DefineClass"
            )

    registrations: list[Registration] = []
    seen_cpp: set[str] = set()
    seen_js: set[str] = set()
    marker_count = 0
    for match in METHOD_MARKER.finditer(text):
        marker_count += 1
        opening = text.find("<", match.start(), match.end() + 1)
        if opening < 0:
            raise GuardGenerationError(f"{source}:{match.start()}: missing angle after GuardedInstanceMethod")
        closing = _find_closing_angle(text, opening, source)
        suffix = text[closing + 1 :]
        name_match = re.match(r"\s*\(\s*\"([A-Za-z0-9_]+)\"\s*\)", suffix)
        if name_match is None:
            raise GuardGenerationError(
                f"{source}:{closing}: GuardedInstanceMethod must end with a quoted JS method name"
            )
        cpp_name, masks = _parse_template_inner(text[opening + 1 : closing], source, match.start())
        js_name = name_match.group(1)
        if cpp_name in seen_cpp:
            raise GuardGenerationError(f"{source}:{match.start()}: duplicate C++ method {cpp_name}")
        if js_name in seen_js:
            raise GuardGenerationError(f"{source}:{match.start()}: duplicate JS method {js_name}")
        seen_cpp.add(cpp_name)
        seen_js.add(js_name)
        _validate_masks(js_name, masks, source, match.start())
        registrations.append(Registration(cpp_name, js_name, masks))

    if marker_count != len(registrations):
        raise GuardGenerationError(
            f"{source}: parsed {len(registrations)} GuardedInstanceMethod registrations out of {marker_count} markers"
        )
    if marker_count != EXPECTED_NODE_REGISTRATIONS:
        raise GuardGenerationError(
            f"{source}: expected {EXPECTED_NODE_REGISTRATIONS} GuardedInstanceMethod registrations, found {marker_count}"
        )
    return registrations


def _validate_masks(js_name: str, masks: tuple[int, ...], source: Path, offset: int) -> None:
    optional_null_index = MASK_NAMES.index("optionalNull")
    optional = masks[MASK_NAMES.index("optionalU32")] | masks[MASK_NAMES.index("optionalU64")]
    for index, mask in enumerate(masks):
        if mask >= (1 << 32):
            raise GuardGenerationError(
                f"{source}:{offset}: {js_name} mask {MASK_NAMES[index]} uses a positional bit >= 32"
            )
    if masks[optional_null_index] & ~optional:
        raise GuardGenerationError(f"{source}:{offset}: {js_name} optional-null mask is not optional")
    # Every mask but optionalNull states a position's whole rule, so a
    # position may carry at most one of them.
    for index, mask in enumerate(masks):
        if index == optional_null_index:
            continue
        for other_index in range(index + 1, len(masks)):
            if other_index != optional_null_index and mask & masks[other_index]:
                raise GuardGenerationError(
                    f"{source}:{offset}: {js_name} mask {MASK_NAMES[index]} overlaps {MASK_NAMES[other_index]}"
                )


def _parse_wasm_methods(source: Path) -> set[str]:
    text = source.read_text(encoding="utf-8")
    methods = set(WASM_METHOD.findall(text))
    if not methods:
        raise GuardGenerationError(f"{source}: no embind methods found")
    return methods


def _parse_free_functions(source: Path) -> set[str]:
    text = source.read_text(encoding="utf-8")
    functions = set(WASM_FREE_FUNCTION.findall(text))
    missing = sorted(set(FREE_FUNCTION_GUARDS) - functions)
    if missing:
        raise GuardGenerationError(f"{source}: positional free-function registration drift: missing={missing}")
    return functions


def _parse_static_functions(source: Path) -> set[str]:
    text = source.read_text(encoding="utf-8")
    functions = set(WASM_STATIC_FUNCTION.findall(text))
    if functions != set(STATIC_FUNCTION_GUARDS):
        missing = sorted(set(STATIC_FUNCTION_GUARDS) - functions)
        extra = sorted(functions - set(STATIC_FUNCTION_GUARDS))
        raise GuardGenerationError(f"{source}: static function registration drift: missing={missing}, extra={extra}")
    return functions


def _workbook_specs(node: Iterable[Registration]) -> dict[str, dict[str, int]]:
    """The positional spec of every WASM `Workbook` method: the Node mask."""
    return {entry.js_name: entry.as_js_spec() for entry in node if entry.js_name not in NODE_ONLY_METHODS}


def _check_surface(specs: dict[str, dict[str, int]], wasm_methods: set[str]) -> None:
    expected_wasm = set(specs)
    if wasm_methods != expected_wasm:
        missing = sorted(expected_wasm - wasm_methods)
        extra = sorted(wasm_methods - expected_wasm)
        raise GuardGenerationError(f"WASM/Node binding surface drift: missing={missing}, extra={extra}")


_INC_ARRAY = r"\b%s\[\]\s*=\s*\{(.*?)\n\};"
_HANDLE_ROW = re.compile(r'^\s*\{"([A-Za-z0-9_]+)",\s*"([A-Za-z0-9_.]+)",\s*(\d+)\},$', re.MULTILINE)


def _inc_block(name: str, text: str) -> str:
    found = re.search(_INC_ARRAY % name, text, re.S)
    if found is None:
        raise GuardGenerationError(f"{FACADE_SCOPE_INC}: no `{name}` table")
    return found.group(1)


def _sheet_methods(wasm_methods: set[str], text: str | None = None) -> list[str]:
    """Raw methods whose first argument is the sheet index the `Sheet` facade supplies."""
    text = FACADE_SCOPE_INC.read_text(encoding="utf-8") if text is None else text
    names = re.findall(r'^\s*"([A-Za-z0-9_]+)",$', _inc_block("kSheetMethods", text), re.MULTILINE)
    return sorted(set(names) & wasm_methods)


def _handle_methods(wasm_methods: set[str], text: str | None = None) -> dict[str, list[object]]:
    """`{raw: [label, supplied]}` of the raw methods a handle class calls: the `Class.method` its argument
    errors carry and the count of leading raw arguments the handle supplies."""
    text = FACADE_SCOPE_INC.read_text(encoding="utf-8") if text is None else text
    rows = _HANDLE_ROW.findall(_inc_block("kHandleMethods", text))
    return {raw: [label, int(supplied)] for raw, label, supplied in sorted(rows) if raw in wasm_methods}


def _build_metadata(
    node: list[Registration],
    specs: dict[str, dict[str, int]],
    sheet: list[str],
    free: set[str],
    static: set[str],
    handle: dict[str, list[object]] | None = None,
) -> dict[str, object]:
    return {
        "nodeRegistrationCount": len(node),
        "workbook": dict(sorted(specs.items())),
        "sheet": sheet,
        "handle": dict(sorted((handle or {}).items())),
        "static": {name: dict(STATIC_FUNCTION_GUARDS[name]) for name in sorted(static)},
        "free": {name: dict(FREE_FUNCTION_GUARDS.get(name, {})) for name in sorted(free)},
        "wasmExcluded": sorted(NODE_ONLY_METHODS),
    }


def _render(metadata: dict[str, object]) -> str:
    data = json.dumps(metadata, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return f"""// AUTO-GENERATED by tools/codegen/gen_wasm_positional_guards.py.
// Do not edit by hand; regenerate from the Node and WASM binding sources.
(function (module) {{
  'use strict';

  const metadata = Object.freeze({data});
  const installKey = Symbol.for('formulon.checkedPositionalArguments');
  const activeWorkbookCalls = new WeakMap();
  // Native addresses of the workbooks whose iterative progress callback is
  // running, kept by the native trampoline. The callback runs inside the
  // recalc, so every method of such a workbook is refused, as on Node.
  const progressCallbackKey = Symbol.for('formulon.progressCallbackWorkbooks');
  const UINT32_MAX = 0xffffffff;
  const INT32_MIN = -0x80000000;
  const INT32_MAX = 0x7fffffff;
  const UINT64_MAX_EXACT = Number.MAX_SAFE_INTEGER;
  // The native code records the call's failure here while a wrapped call is
  // active: kind 0 / 1 a nested-field TypeError / RangeError, 2 an exception a
  // getter or callback threw, 3 a failed status. The wrapper throws it after
  // the method returns, so native frames unwind normally and both surfaces
  // raise the same error.
  const argumentErrors = {{ depth: 0, pending: undefined, thrown: undefined }};
  module[Symbol.for('formulon.argumentErrors')] = argumentErrors;

  function pendingError(pending) {{
    if (pending.kind === 3) {{
      const FormulonError = module.FormulonError;
      if (typeof FormulonError !== 'function') {{
        return new TypeError('FormulonError was not provided to the module factory');
      }}
      return new FormulonError(pending.code, pending.message, pending.context, pending.parseError);
    }}
    if (pending.kind === 1) return new RangeError(pending.message);
    if (pending.kind === 2 && pending.thrown !== undefined) return pending.thrown.value;
    return new TypeError(pending.message);
  }}

  // The raw methods the `Sheet` facade calls with the sheet index first. Their
  // errors name the facade method and count arguments as it sees them; the
  // sheet index itself only reaches a guard through a direct raw call.
  const sheetMethods = new Set(metadata.sheet);
  // The raw methods a handle class calls: `[label, supplied]`, the `Class.method` its errors name and the
  // count of leading arguments the handle supplies. A handle method takes precedence over a sheet method.
  const handleMethods = new Map(Object.entries(metadata.handle));

  // `owner` is 'Workbook' for the class and its statics, null for a free function.
  function argumentLabel(owner, name, index) {{
    if (owner === null) return name + ' argument ' + (index + 1);
    const handle = owner === 'Workbook' ? handleMethods.get(name) : undefined;
    if (handle !== undefined) {{
      const [label, supplied] = handle;
      return index < supplied
        ? label + ' supplied argument ' + (index + 1)
        : label + ' argument ' + (index - supplied + 1);
    }}
    if (owner === 'Workbook' && sheetMethods.has(name)) {{
      return index === 0 ? 'Sheet.' + name + ' sheet index' : 'Sheet.' + name + ' argument ' + index;
    }}
    return owner + '.' + name + ' argument ' + (index + 1);
  }}

  function argumentType(label, expected) {{
    throw new TypeError(label + ' must be a primitive ' + expected);
  }}

  function argumentRequired(label) {{
    throw new TypeError(label + ' is required');
  }}

  function argumentRange(label, expected) {{
    throw new RangeError(label + ' is outside ' + expected + ' range');
  }}

  function validateNumber(value, label, kind) {{
    if (typeof value !== 'number') {{
      argumentType(label, 'number');
    }}
    if (!Number.isInteger(value)) {{
      argumentRange(label, kind);
    }}
    if (kind === 'uint32' && (value < 0 || value > UINT32_MAX)) {{
      argumentRange(label, kind);
    }}
    if (kind === 'int32' && (value < INT32_MIN || value > INT32_MAX)) {{
      argumentRange(label, kind);
    }}
    if (kind === 'uint64' && (value < 0 || value > UINT64_MAX_EXACT)) {{
      argumentRange(label, kind);
    }}
  }}

  function validateArguments(owner, name, args, spec) {{
    const masks = [
      ['u32', spec.u32 ?? 0, 'uint32'],
      ['i32', spec.i32 ?? 0, 'int32'],
      ['u64', spec.u64 ?? 0, 'uint64'],
      ['double', spec.double ?? 0, 'double'],
      ['string', spec.string ?? 0, 'string'],
      ['optionalU32', spec.optionalU32 ?? 0, 'uint32'],
      ['optionalU64', spec.optionalU64 ?? 0, 'uint64'],
    ];
    const optionalNull = spec.optionalNull ?? 0;
    const objectMask = spec.object ?? 0;
    const requiredMask = spec.required ?? 0;
    // Positions every mask but the optional ones treat as present, so an
    // omitted trailing argument fails its own check instead of being skipped.
    const presentMask = ((spec.u32 ?? 0) | (spec.i32 ?? 0) | (spec.u64 ?? 0) | (spec.double ?? 0) |
                         (spec.string ?? 0) | objectMask | requiredMask) >>> 0;
    const checked = Math.max(args.length, 32 - Math.clz32(presentMask));
    for (let index = 0; index < checked; index += 1) {{
      const bit = 2 ** index;
      const value = args[index];
      const label = () => argumentLabel(owner, name, index);
      if ((objectMask & bit) !== 0 &&
          (value === null || (typeof value !== 'object' && typeof value !== 'function'))) {{
        argumentType(label(), 'object');
      }}
      if ((requiredMask & bit) !== 0 && value === undefined) argumentRequired(label());
      for (const [kind, mask, numericKind] of masks) {{
        if ((mask & bit) === 0) continue;
        if (kind === 'double') {{
          if (typeof value !== 'number') argumentType(label(), 'number');
        }} else if (kind === 'string') {{
          if (typeof value !== 'string') argumentType(label(), 'string');
        }} else if (kind === 'optionalU32' || kind === 'optionalU64') {{
          if (value === undefined || (value === null && (optionalNull & bit) !== 0)) continue;
          if (value === null) argumentType(label(), 'number');
          validateNumber(value, label(), numericKind);
        }} else {{
          validateNumber(value, label(), numericKind);
        }}
      }}
    }}
  }}

  // Embind would forward a lone UTF-16 surrogate as WTF-8 bytes; replace it
  // with U+FFFD as N-API's UTF-8 conversion does.
  const loneSurrogate = /[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|(?<![\\uD800-\\uDBFF])[\\uDC00-\\uDFFF]/g;
  function wellFormed(text) {{
    return typeof text.toWellFormed === 'function' ? text.toWellFormed() : text.replace(loneSurrogate, '\\uFFFD');
  }}

  // Returns `args` with every string-mask position made well-formed. The
  // array is the wrapper's own rest parameter, never the caller's.
  function wellFormedStrings(args, spec) {{
    const stringMask = spec.string ?? 0;
    for (let index = 0; index < args.length; index += 1) {{
      if ((stringMask & 2 ** index) !== 0) args[index] = wellFormed(args[index]);
    }}
    return args;
  }}

  function receiverStateKey(receiver) {{
    if (receiver === null || receiver === undefined ||
        (typeof receiver !== 'object' && typeof receiver !== 'function')) return null;
    // Embind proxies expose the same own `$$` data property as the original
    // handle. Reading its descriptor avoids invoking a user accessor while
    // giving every wrapper for the same native state one WeakMap key.
    const descriptor = Object.getOwnPropertyDescriptor(receiver, '$$');
    if (descriptor && Object.prototype.hasOwnProperty.call(descriptor, 'value')) {{
      const value = descriptor.value;
      if (value !== null && (typeof value === 'object' || typeof value === 'function')) return value;
    }}
    return receiver;
  }}

  function makeCallableWrapper(owner, name, original, spec, trackWorkbookCall) {{
    const wrapped = function (...args) {{
      validateArguments(owner, name, args, spec);
      const receiverKey = trackWorkbookCall ? receiverStateKey(this) : null;
      const busy = module[progressCallbackKey];
      if (busy !== undefined && busy.size > 0 && receiverKey !== null && busy.has(receiverKey.ptr)) {{
        throw pendingError({{ kind: 3, code: 4005, context: '',
          message: 'cannot call a Workbook method from its iterative progress callback ' +
                   '(kGraphRecalcReentrant, 4005)' }});
      }}
      const receiverTracked = receiverKey !== null;
      if (receiverTracked) activeWorkbookCalls.set(receiverKey, (activeWorkbookCalls.get(receiverKey) ?? 0) + 1);
      // A getter or callback may re-enter another wrapped method mid-call;
      // each call owns only the failure recorded during its own native frame.
      const outerPending = argumentErrors.pending;
      argumentErrors.pending = undefined;
      argumentErrors.depth += 1;
      let pending;
      let result;
      try {{
        result = Reflect.apply(original, this, wellFormedStrings(args, spec));
      }} finally {{
        argumentErrors.depth -= 1;
        pending = argumentErrors.pending;
        argumentErrors.pending = outerPending;
        if (receiverTracked) {{
          const depth = activeWorkbookCalls.get(receiverKey) ?? 1;
          if (depth <= 1) activeWorkbookCalls.delete(receiverKey);
          else activeWorkbookCalls.set(receiverKey, depth - 1);
        }}
      }}
      if (pending !== undefined) throw pendingError(pending);
      return result;
    }};
    Object.defineProperty(wrapped, 'name', {{ value: original.name, configurable: true }});
    const argCountDescriptor = Object.getOwnPropertyDescriptor(original, 'argCount');
    if (argCountDescriptor) Object.defineProperty(wrapped, 'argCount', argCountDescriptor);
    return wrapped;
  }}

  function guardOverloadTable(owner, name, original, table, spec, trackWorkbookCall) {{
    if (!Array.isArray(table)) {{
      throw new TypeError('cannot guard non-array overload table ' + name);
    }}
    const guarded = new Array(table.length);
    for (let index = 0; index < table.length; index += 1) {{
      if (!(index in table)) continue;
      const entry = table[index];
      if (typeof entry !== 'function') {{
        throw new TypeError('cannot guard non-function overload entry ' + name + '[' + index + ']');
      }}
      if (entry === original) {{
        throw new Error('overload table recursively references dispatcher ' + name);
      }}
      // Leaf wrappers intentionally do not copy or inspect another
      // overloadTable. Current embind leaves are terminal functions; keeping
      // this non-recursive makes cycles fail closed at the dispatcher edge.
      guarded[index] = makeCallableWrapper(owner, name, entry, spec, trackWorkbookCall);
    }}
    return guarded;
  }}

  function wrap(owner, target, name, spec, trackWorkbookCall = false) {{
    const descriptor = Object.getOwnPropertyDescriptor(target, name);
    if (!descriptor) return false;
    if (typeof descriptor.value !== 'function') {{
      throw new TypeError('cannot guard non-function binding ' + name);
    }}
    const original = descriptor.value;
    const wrapped = makeCallableWrapper(owner, name, original, spec, trackWorkbookCall);
    const overloadDescriptor = Object.getOwnPropertyDescriptor(original, 'overloadTable');
    if (overloadDescriptor) {{
      if (!Object.prototype.hasOwnProperty.call(overloadDescriptor, 'value')) {{
        throw new TypeError('cannot guard accessor overload table ' + name);
      }}
      const guardedTable =
        guardOverloadTable(owner, name, original, overloadDescriptor.value, spec, trackWorkbookCall);
      Object.defineProperty(wrapped, 'overloadTable', {{ ...overloadDescriptor, value: guardedTable }});
    }}
    Object.defineProperty(target, name, {{ ...descriptor, value: wrapped }});
    return true;
  }}

  // Fails closed: an own function of `target` that `wrapped` does not name
  // would return its default value on failure instead of throwing.
  function requireWrapped(target, wrapped, what) {{
    for (const name of Object.getOwnPropertyNames(target)) {{
      if (name === 'constructor' || wrapped.has(name)) continue;
      const descriptor = Object.getOwnPropertyDescriptor(target, name);
      if (descriptor && typeof descriptor.value === 'function') {{
        throw new Error('unwrapped ' + what + ' ' + name);
      }}
    }}
  }}

  function install() {{
    if (module[installKey]) return;
    const workbook = module.Workbook;
    if (!workbook || !workbook.prototype) return;
    const wrappedMethods = new Set();
    for (const [name, spec] of Object.entries(metadata.workbook)) {{
      if (wrap('Workbook', workbook.prototype, name, spec, true)) wrappedMethods.add(name);
      else if (!metadata.wasmExcluded.includes(name)) throw new Error('missing positional binding ' + name);
    }}
    requireWrapped(workbook.prototype, wrappedMethods, 'Workbook method');
    for (const [name, spec] of Object.entries(metadata.static)) {{
      if (!wrap('Workbook', workbook, name, spec)) throw new Error('missing Workbook static function ' + name);
    }}
    requireWrapped(workbook, new Set(Object.keys(metadata.static)), 'Workbook static function');
    for (const [name, spec] of Object.entries(metadata.free)) {{
      if (!wrap(null, module, name, spec)) throw new Error('missing positional free function ' + name);
    }}
    const lifecycleKeys = ['dispose', 'delete', 'deleteLater'];
    if (typeof Symbol.dispose === 'symbol') lifecycleKeys.push(Symbol.dispose);
    for (const key of lifecycleKeys) {{
      let owner = workbook.prototype;
      let descriptor;
      while (owner !== null && descriptor === undefined) {{
        descriptor = Object.getOwnPropertyDescriptor(owner, key);
        if (descriptor === undefined) owner = Object.getPrototypeOf(owner);
      }}
      if (!descriptor || typeof descriptor.value !== 'function') {{
        throw new Error('missing Workbook lifecycle binding ' + String(key));
      }}
      const original = descriptor.value;
      const guarded = function (...args) {{
        const receiverKey = receiverStateKey(this);
        if (receiverKey !== null && activeWorkbookCalls.get(receiverKey) > 0) {{
          throw pendingError({{ kind: 3, code: 4005, context: '',
            message: 'cannot dispose Workbook during an active method call (kGraphRecalcReentrant, 4005)' }});
        }}
        return Reflect.apply(original, this, args);
      }};
      Object.defineProperty(guarded, 'name', {{ value: original.name, configurable: true }});
      for (const property of ['argCount', 'overloadTable']) {{
        const propertyDescriptor = Object.getOwnPropertyDescriptor(original, property);
        if (propertyDescriptor) Object.defineProperty(guarded, property, propertyDescriptor);
      }}
      // Keep the guard on the descriptor's actual owner as well as on the
      // Workbook shadow. A caller can retain an inherited descriptor and
      // invoke it with Function#call, bypassing a Workbook-only shadow.
      Object.defineProperty(owner, key, {{ ...descriptor, value: guarded }});
      if (owner !== workbook.prototype) {{
        Object.defineProperty(workbook.prototype, key, {{ ...descriptor, value: guarded }});
      }}
    }}
    // `using` calls dispose(); embind's own [Symbol.dispose] detaches the
    // wrapper instead, so the Workbook rebinds the symbol to the guarded
    // dispose(). The inherited embind method stays guarded above.
    if (typeof Symbol.dispose === 'symbol') {{
      Object.defineProperty(workbook.prototype, Symbol.dispose, {{
        value: workbook.prototype.dispose,
        writable: true,
        configurable: true,
      }});
    }}
    module[installKey] = true;
  }}

  // Emscripten invokes onRuntimeInitialized before postRun. The caller can
  // replace either hook from preRun, so armInstall is appended after all
  // existing preRun callbacks and re-wraps the final hooks at that point.
  const runtimeWrapperKey = Symbol('formulon.positionalRuntimeWrapper');
  function armRuntimeHook() {{
    const runtimeHook = module.onRuntimeInitialized;
    if (runtimeHook && runtimeHook[runtimeWrapperKey]) return;
    const wrapped = function (...args) {{
      install();
      if (runtimeHook === undefined || runtimeHook === null) return undefined;
      return Reflect.apply(runtimeHook, this, args);
    }};
    Object.defineProperty(wrapped, runtimeWrapperKey, {{ value: true }});
    module.onRuntimeInitialized = wrapped;
  }}

  function armPostRun() {{
    const postRun = module.postRun;
    if (postRun === undefined) {{
      module.postRun = [install];
    }} else if (Array.isArray(postRun)) {{
      module.postRun = [install, ...postRun.filter((callback) => callback !== install)];
    }} else {{
      module.postRun = [install, postRun];
    }}
  }}

  function armInstall() {{
    armRuntimeHook();
    armPostRun();
  }}

  function appendPreRunArm() {{
    const preRun = module.preRun;
    if (preRun === undefined) module.preRun = [armInstall];
    else if (Array.isArray(preRun)) {{
      module.preRun = [...preRun.filter((callback) => callback !== armInstall), armInstall];
    }} else module.preRun = [preRun, armInstall];
  }}

  armInstall();
  appendPreRunArm();
  const preInit = module.preInit;
  if (preInit === undefined) module.preInit = [appendPreRunArm];
  else if (Array.isArray(preInit)) module.preInit = [...preInit, appendPreRunArm];
  else module.preInit = [preInit, appendPreRunArm];
}})(Module);
"""


def generate(node_source: Path, wasm_source: Path, output: Path) -> None:
    node = _parse_node_registrations(node_source)
    wasm = _parse_wasm_methods(wasm_source)
    specs = _workbook_specs(node)
    _check_surface(specs, wasm)
    free = _parse_free_functions(wasm_source)
    static = _parse_static_functions(wasm_source)
    metadata = _build_metadata(node, specs, _sheet_methods(wasm), free, static, _handle_methods(wasm))
    rendered = _render(metadata)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(rendered, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--node-source", type=Path, default=DEFAULT_NODE_SOURCE)
    parser.add_argument("--wasm-source", type=Path, default=DEFAULT_WASM_SOURCE)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        generate(args.node_source, args.wasm_source, args.output)
    except (OSError, GuardGenerationError) as exc:
        print(f"gen_wasm_positional_guards: {exc}", file=sys.stderr)
        return 1
    print(f"generated {args.output} from {EXPECTED_NODE_REGISTRATIONS} Node registrations")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
