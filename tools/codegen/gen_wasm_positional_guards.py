#!/usr/bin/env python3
"""Generate the embind positional-argument guard pre-js.

The N-API binding is the canonical declaration of the primitive positional
types. Its GuardedInstanceMethod masks are parsed here instead of being
copied into a second hand-maintained table. The parser is deliberately
strict: a malformed registration, duplicate name, mask overlap, or binding
surface drift stops generation before an artifact can be built.

The generated file is passed to emscripten as --pre-js. It contains no
native code, so the guard lives in the JS glue and does not change the wasm
module. The same wrapper throws the nested-field rejections the native readers
record, so a bad field raises TypeError/RangeError as on the Node surface.
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
DEFAULT_NODE_SOURCE = REPO_ROOT / "src/node_addon/parts/workbook_class.cc"
DEFAULT_WASM_SOURCE = REPO_ROOT / "src/wasm/parts/bindings_register.cpp"
EXPECTED_NODE_REGISTRATIONS = 275

MASK_NAMES = (
    "u32",
    "i32",
    "u64",
    "double",
    "string",
    "optionalU32",
    "optionalU64",
    "optionalNull",
)
MASK_TOKEN = re.compile(r"0x([0-9a-fA-F]+)ULL\Z")
METHOD_MARKER = re.compile(r"\bGuardedInstanceMethod(?=\s*<)")
WASM_METHOD = re.compile(r"\.function\(\s*\"([A-Za-z0-9_]+)\"")
WASM_FREE_FUNCTION = re.compile(r"\bfunction\(\s*\"([A-Za-z0-9_]+)\"")

# The Node class has two lifecycle helpers that are intentionally not exposed
# by the embind class. Two embind methods have no Node counterpart and carry
# only nested/object arguments, so they do not need a primitive positional
# guard. Keeping these sets explicit makes a future binding drift fail at
# generation time instead of silently losing validation.
NODE_ONLY_METHODS = frozenset(("dispose", "memoryUsage"))
WASM_UNGUARDED_METHODS = frozenset(("addCellStyleXf", "createTable"))

WASM_ONLY_GUARDS = {
    "getSheetAutoFilterXml": {"u32": 0x1},
    "setSheetAutoFilterXml": {"u32": 0x1, "string": 0x2},
    "updateTable": {"u32": 0x1},
    "removeTable": {"u32": 0x1},
}

FREE_FUNCTION_GUARDS = {
    "evalFormula": {"string": 0x1},
    "statusString": {"i32": 0x1},
    "errorDisplayName": {"i32": 0x1},
    "setLogMinLevel": {"i32": 0x1},
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
    required = masks[:5]
    optional = masks[5:7]
    optional_null = masks[7]
    for index, mask in enumerate(masks):
        if mask >= (1 << 32):
            raise GuardGenerationError(
                f"{source}:{offset}: {js_name} mask {MASK_NAMES[index]} uses a positional bit >= 32"
            )
    if optional[0] & optional[1]:
        raise GuardGenerationError(f"{source}:{offset}: {js_name} optional u32/u64 masks overlap")
    if optional_null & ~(optional[0] | optional[1]):
        raise GuardGenerationError(f"{source}:{offset}: {js_name} optional-null mask is not optional")
    for index, mask in enumerate(required):
        other_mask = 0
        for other_index, other in enumerate(masks):
            if other_index != index:
                other_mask |= other
        if mask & other_mask:
            raise GuardGenerationError(f"{source}:{offset}: {js_name} mask {MASK_NAMES[index]} overlaps another type")


def _parse_wasm_methods(source: Path) -> set[str]:
    text = source.read_text(encoding="utf-8")
    methods = set(WASM_METHOD.findall(text))
    if not methods:
        raise GuardGenerationError(f"{source}: no embind methods found")
    return methods


def _check_free_functions(source: Path) -> None:
    text = source.read_text(encoding="utf-8")
    functions = set(WASM_FREE_FUNCTION.findall(text))
    missing = sorted(set(FREE_FUNCTION_GUARDS) - functions)
    if missing:
        raise GuardGenerationError(f"{source}: positional free-function registration drift: missing={missing}")


def _check_surface(node: Iterable[Registration], wasm_methods: set[str]) -> None:
    registrations = list(node)
    node_names = {entry.js_name for entry in registrations}
    expected_wasm = (node_names - NODE_ONLY_METHODS) | set(WASM_ONLY_GUARDS) | WASM_UNGUARDED_METHODS
    if wasm_methods != expected_wasm:
        missing = sorted(expected_wasm - wasm_methods)
        extra = sorted(wasm_methods - expected_wasm)
        raise GuardGenerationError(f"WASM/Node binding surface drift: missing={missing}, extra={extra}")


def _build_metadata(node: list[Registration]) -> dict[str, object]:
    workbook: dict[str, dict[str, int]] = {
        entry.js_name: entry.as_js_spec() for entry in sorted(node, key=lambda entry: entry.js_name)
    }
    # These embind-only methods have no primitive positional arguments, but
    # they still cross a native Workbook frame. Include them in the generated
    # wrapper with an empty spec so nested getters cannot destroy the receiver
    # while the call is active.
    for name in WASM_UNGUARDED_METHODS:
        workbook[name] = {}
    for name, spec in WASM_ONLY_GUARDS.items():
        workbook[name] = dict(spec)
    return {
        "nodeRegistrationCount": len(node),
        "workbook": dict(sorted(workbook.items())),
        "free": {name: dict(spec) for name, spec in sorted(FREE_FUNCTION_GUARDS.items())},
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
  const UINT32_MAX = 0xffffffff;
  const INT32_MIN = -0x80000000;
  const INT32_MAX = 0x7fffffff;
  const UINT64_MAX_EXACT = Number.MAX_SAFE_INTEGER;
  // The native readers record a nested-field rejection here while a wrapped
  // call is active; the wrapper throws it after the method returns, so native
  // frames unwind normally and both surfaces raise the same error class.
  const argumentErrors = {{ depth: 0, pending: undefined, thrown: undefined }};
  module[Symbol.for('formulon.argumentErrors')] = argumentErrors;

  function nestedArgumentError(pending) {{
    if (pending.kind === 1) return new RangeError(pending.message);
    if (pending.kind === 2 && pending.thrown !== undefined) return pending.thrown.value;
    return new TypeError(pending.message);
  }}

  function argumentType(name, index, expected) {{
    throw new TypeError(name + ' argument ' + (index + 1) + ' must be a primitive ' + expected);
  }}

  function argumentRange(name, index, expected) {{
    throw new RangeError(name + ' argument ' + (index + 1) + ' is outside ' + expected + ' range');
  }}

  function validateNumber(value, name, index, kind) {{
    if (typeof value !== 'number') {{
      argumentType(name, index, 'number');
    }}
    if (!Number.isInteger(value)) {{
      argumentRange(name, index, kind);
    }}
    if (kind === 'uint32' && (value < 0 || value > UINT32_MAX)) {{
      argumentRange(name, index, kind);
    }}
    if (kind === 'int32' && (value < INT32_MIN || value > INT32_MAX)) {{
      argumentRange(name, index, kind);
    }}
    if (kind === 'uint64' && (value < 0 || value > UINT64_MAX_EXACT)) {{
      argumentRange(name, index, kind);
    }}
  }}

  function validateArguments(name, args, spec) {{
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
    for (let index = 0; index < args.length; index += 1) {{
      const bit = 2 ** index;
      for (const [kind, mask, numericKind] of masks) {{
        if ((mask & bit) === 0) continue;
        const value = args[index];
        if (kind === 'double') {{
          if (typeof value !== 'number') argumentType(name, index, 'number');
        }} else if (kind === 'string') {{
          if (typeof value !== 'string') argumentType(name, index, 'string');
        }} else if (kind === 'optionalU32' || kind === 'optionalU64') {{
          if (value === undefined || (value === null && (optionalNull & bit) !== 0)) continue;
          if (value === null) argumentType(name, index, 'number');
          validateNumber(value, name, index, numericKind);
        }} else {{
          validateNumber(value, name, index, numericKind);
        }}
      }}
    }}
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

  function makeCallableWrapper(name, original, spec, trackWorkbookCall) {{
    const wrapped = function (...args) {{
      validateArguments(name, args, spec);
      const receiverKey = trackWorkbookCall ? receiverStateKey(this) : null;
      const receiverTracked = receiverKey !== null;
      if (receiverTracked) activeWorkbookCalls.set(receiverKey, (activeWorkbookCalls.get(receiverKey) ?? 0) + 1);
      // A getter may re-enter another wrapped method mid-read; each call
      // owns only the rejection recorded during its own native frame.
      const outerPending = argumentErrors.pending;
      argumentErrors.pending = undefined;
      argumentErrors.depth += 1;
      let pending;
      let result;
      try {{
        result = Reflect.apply(original, this, args);
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
      if (pending !== undefined) throw nestedArgumentError(pending);
      return result;
    }};
    Object.defineProperty(wrapped, 'name', {{ value: original.name, configurable: true }});
    const argCountDescriptor = Object.getOwnPropertyDescriptor(original, 'argCount');
    if (argCountDescriptor) Object.defineProperty(wrapped, 'argCount', argCountDescriptor);
    return wrapped;
  }}

  function guardOverloadTable(name, original, table, spec, trackWorkbookCall) {{
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
      guarded[index] = makeCallableWrapper(name, entry, spec, trackWorkbookCall);
    }}
    return guarded;
  }}

  function wrap(target, name, spec, trackWorkbookCall = false) {{
    const descriptor = Object.getOwnPropertyDescriptor(target, name);
    if (!descriptor) return false;
    if (typeof descriptor.value !== 'function') {{
      throw new TypeError('cannot guard non-function binding ' + name);
    }}
    const original = descriptor.value;
    const wrapped = makeCallableWrapper(name, original, spec, trackWorkbookCall);
    const overloadDescriptor = Object.getOwnPropertyDescriptor(original, 'overloadTable');
    if (overloadDescriptor) {{
      if (!Object.prototype.hasOwnProperty.call(overloadDescriptor, 'value')) {{
        throw new TypeError('cannot guard accessor overload table ' + name);
      }}
      const guardedTable = guardOverloadTable(name, original, overloadDescriptor.value, spec, trackWorkbookCall);
      Object.defineProperty(wrapped, 'overloadTable', {{ ...overloadDescriptor, value: guardedTable }});
    }}
    Object.defineProperty(target, name, {{ ...descriptor, value: wrapped }});
    return true;
  }}

  function install() {{
    if (module[installKey]) return;
    const workbook = module.Workbook;
    if (!workbook || !workbook.prototype) return;
    for (const [name, spec] of Object.entries(metadata.workbook)) {{
      if (!wrap(workbook.prototype, name, spec, true) && !metadata.wasmExcluded.includes(name)) {{
        throw new Error('missing positional binding ' + name);
      }}
    }}
    for (const [name, spec] of Object.entries(metadata.free)) {{
      if (!wrap(module, name, spec)) throw new Error('missing positional free function ' + name);
    }}
    const lifecycleKeys = ['delete', 'deleteLater'];
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
          throw new TypeError('cannot dispose Workbook during an active method call');
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
    _check_surface(node, wasm)
    _check_free_functions(wasm_source)
    metadata = _build_metadata(node)
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
