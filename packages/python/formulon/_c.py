"""Low-level WASM binding for the Formulon C ABI.

Public users should import the ``Workbook`` / ``Value`` / ``FormulonError``
symbols re-exported by the top-level :mod:`formulon` package instead.

Architecture
------------

The binding loads ``formulon_capi.wasm`` (a standalone reactor-style
WebAssembly module that exports the ``fm_*`` C ABI from
``src/c_api/formulon_c.h``) via the ``wasmtime`` runtime. A single
module instance is created lazily at first use and shared across every
:class:`formulon.Workbook` in the process; each ``Workbook`` instance
owns an opaque ``fm_workbook_t*`` (i32 in WASM) handle.

Pointers are 32-bit offsets into the WASM linear memory. Strings cross
the boundary as ``malloc``-allocated UTF-8 NUL-terminated buffers in
that memory; the caller is responsible for ``free``-ing them after the
WASM function returns. Borrowed return pointers (e.g. text from
``fm_workbook_get_value``) are read directly from WASM memory and
decoded eagerly into Python ``str`` so the result outlives any
subsequent WASM mutation.

The ``fm_value_t`` POD has the wasm32 layout::

    offset 0: int32  kind
    offset 4: int32  (padding for 8-byte alignment of the union)
    offset 8: union  { double number ; int32 boolean ; int32 error_code ; ptr text }
    total: 16 bytes

The union member is selected by ``kind``; reading any other member is
undefined per the C ABI contract.
"""

from __future__ import annotations

import struct
import threading
from enum import IntEnum
from pathlib import Path
from typing import Optional, Tuple

import wasmtime

__all__ = [
    "LIB",
    "ValueKind",
    "decode_cstr",
    "fm_value_t_size",
]

# fm_value_t layout: int32 kind + 4 pad + 8 union = 16 bytes.
fm_value_t_size = 16


# ---------------------------------------------------------------------------
# ValueKind enum (matches fm_value_kind_t)
# ---------------------------------------------------------------------------


class ValueKind(IntEnum):
    """Mirror of ``fm_value_kind_t`` in ``formulon_c.h``."""

    BLANK = 0
    NUMBER = 1
    BOOL = 2
    TEXT = 3
    ERROR = 4
    ARRAY = 5
    REF = 6
    LAMBDA = 7


# ---------------------------------------------------------------------------
# WASM module location
# ---------------------------------------------------------------------------


def _locate_wasm() -> Path:
    """Return the on-disk path to ``formulon_capi.wasm``.

    Search order:
      1. ``packages/python/formulon/_wasm/formulon_capi.wasm`` -- the
         package-data location populated by ``make python-package``
         and shipped inside the wheel.
      2. ``$FORMULON_WASM_PATH`` -- explicit override for development.

    Raises:
      FileNotFoundError: when the WASM is not on disk in either
        location. The error message names both candidates.
    """
    here = Path(__file__).resolve().parent
    bundled = here / "_wasm" / "formulon_capi.wasm"
    if bundled.is_file():
        return bundled

    import os

    override = os.environ.get("FORMULON_WASM_PATH")
    if override:
        p = Path(override)
        if p.is_file():
            return p

    raise FileNotFoundError(
        "formulon: failed to locate formulon_capi.wasm. "
        f"Tried: {bundled}. "
        "Run `make python-package` to stage the artifact, or install a "
        "wheel that ships it under formulon/_wasm/."
    )


# ---------------------------------------------------------------------------
# WASM module / store wrapper
# ---------------------------------------------------------------------------


class _WasmInstance:
    """Owns a single ``wasmtime`` engine, store, and module instance.

    The instance is created lazily on first attribute access; subsequent
    Workbook creations reuse it. This trades cold-start latency for
    repeat-call speed and keeps the engine cache around for the life of
    the Python process.

    The store is `not` thread-safe (per the wasmtime-py docs); a single
    process-wide ``_call_lock`` serialises every WASM invocation. The
    underlying calculation engine is already safe for one outstanding
    recalc per ``Workbook`` handle, so the additional lock only prevents
    cross-handle reentry on the wasmtime store itself.
    """

    def __init__(self) -> None:
        self._engine: Optional[wasmtime.Engine] = None
        self._store: Optional[wasmtime.Store] = None
        self._instance: Optional[wasmtime.Instance] = None
        self._memory: Optional[wasmtime.Memory] = None
        self._exports: dict = {}
        self._init_lock = threading.Lock()
        self._call_lock = threading.RLock()
        self._last_diagnostic = threading.local()

    def _ensure(self) -> None:
        if self._instance is not None:
            return
        with self._init_lock:
            if self._instance is not None:
                return
            engine = wasmtime.Engine()
            store = wasmtime.Store(engine)

            # WASI: provide a minimal config. The engine never reads
            # files or stdin; stdout/stderr inherit so any diagnostic
            # libc calls (e.g. trap reasons) surface to the host.
            wasi = wasmtime.WasiConfig()
            wasi.inherit_stdout()
            wasi.inherit_stderr()
            store.set_wasi(wasi)

            wasm_path = _locate_wasm()
            module = wasmtime.Module.from_file(engine, str(wasm_path))

            linker = wasmtime.Linker(engine)
            linker.define_wasi()

            # Stub for env.emscripten_notify_memory_growth. emcc emits
            # this import even under STANDALONE_WASM; it is called on
            # memory.grow but the host has nothing useful to do with it.
            ty = wasmtime.FuncType([wasmtime.ValType.i32()], [])
            linker.define(
                store,
                "env",
                "emscripten_notify_memory_growth",
                wasmtime.Func(store, ty, lambda _i: None),
            )

            instance = linker.instantiate(store, module)
            exports = dict(instance.exports(store).items())

            # Reactor init must run before any export is callable.
            init_fn = exports.get("_initialize")
            if init_fn is not None:
                init_fn(store)

            self._engine = engine
            self._store = store
            self._instance = instance
            self._memory = exports["memory"]
            self._exports = exports

    # ----- raw export accessor --------------------------------------------
    def __getattr__(self, name: str):
        self._ensure()
        fn = self._exports.get(name)
        if fn is None:
            raise AttributeError(f"WASM export '{name}' not found")
        store = self._store
        lock = self._call_lock

        # Wrap so the caller can use a ctypes-like call syntax.
        def _wrapped(*args):
            with lock:
                result = fn(store, *args)
                # Capture diagnostics before another thread can enter the
                # shared, no-pthread WASM instance. Some successful scalar
                # APIs are also non-zero; an unused snapshot is harmless.
                if isinstance(result, int) and result != 0:
                    message_fn = self._exports.get("fm_last_error_message")
                    context_fn = self._exports.get("fm_last_error_context")
                    if message_fn is not None and context_fn is not None:
                        self._last_diagnostic.value = (
                            result,
                            self._read_cstr_unlocked(message_fn(store)),
                            self._read_cstr_unlocked(context_fn(store)),
                        )
                return result

        return _wrapped

    # ----- memory primitives ----------------------------------------------
    @property
    def store(self) -> wasmtime.Store:
        self._ensure()
        assert self._store is not None
        return self._store

    @property
    def memory(self) -> wasmtime.Memory:
        self._ensure()
        assert self._memory is not None
        return self._memory

    @property
    def call_lock(self) -> threading.RLock:
        return self._call_lock

    def read_bytes(self, ptr: int, length: int) -> bytes:
        """Copy ``length`` bytes from WASM memory starting at ``ptr``.

        Takes ``_call_lock`` for the duration of the read: the
        wasmtime ``Store`` is not thread-safe, so a concurrent WASM
        call on another thread (which can grow linear memory) must not
        race with this read.
        """
        self._ensure()
        assert self._memory is not None
        with self._call_lock:
            return bytes(self._memory.read(self._store, ptr, ptr + length))

    def write_bytes(self, ptr: int, data: bytes) -> None:
        """Write ``data`` into WASM memory starting at ``ptr``.

        Takes ``_call_lock`` for the same reason as :meth:`read_bytes`.
        """
        self._ensure()
        assert self._memory is not None
        with self._call_lock:
            self._memory.write(self._store, data, ptr)

    def read_u32(self, ptr: int) -> int:
        return struct.unpack("<I", self.read_bytes(ptr, 4))[0]

    def read_i32(self, ptr: int) -> int:
        return struct.unpack("<i", self.read_bytes(ptr, 4))[0]

    def read_f64(self, ptr: int) -> float:
        return struct.unpack("<d", self.read_bytes(ptr, 8))[0]

    def read_cstr(self, ptr: int) -> str:
        """Decode a NUL-terminated UTF-8 C string from ``ptr``.

        Returns the empty string when ``ptr`` is 0. The whole scan runs
        under ``_call_lock`` (see :meth:`read_bytes`) so a concurrent
        WASM call on another thread cannot grow memory mid-read.
        """
        if ptr == 0:
            return ""
        self._ensure()
        assert self._memory is not None
        # Stream a chunk at a time to avoid copying the whole memory.
        chunks: list[bytes] = []
        offset = ptr
        chunk_size = 256
        with self._call_lock:
            mem_len = self._memory.data_len(self._store)
            while offset < mem_len:
                end = min(offset + chunk_size, mem_len)
                buf = bytes(self._memory.read(self._store, offset, end))
                nul = buf.find(b"\x00")
                if nul >= 0:
                    chunks.append(buf[:nul])
                    break
                chunks.append(buf)
                offset = end
        return b"".join(chunks).decode("utf-8", errors="replace")

    def _read_cstr_unlocked(self, ptr: int) -> str:
        """Decode a C string while the caller already owns ``_call_lock``."""
        if ptr == 0:
            return ""
        assert self._memory is not None
        chunks: list[bytes] = []
        offset = ptr
        mem_len = self._memory.data_len(self._store)
        while offset < mem_len:
            end = min(offset + 256, mem_len)
            buf = bytes(self._memory.read(self._store, offset, end))
            nul = buf.find(b"\x00")
            if nul >= 0:
                chunks.append(buf[:nul])
                break
            chunks.append(buf)
            offset = end
        return b"".join(chunks).decode("utf-8", errors="replace")

    def last_diagnostic(self, status: int) -> Tuple[str, str]:
        """Return the diagnostic atomically captured for ``status`` if any."""
        snapshot = getattr(self._last_diagnostic, "value", None)
        if snapshot is not None and snapshot[0] == status:
            return snapshot[1], snapshot[2]
        with self._call_lock:
            return (
                self._read_cstr_unlocked(self._exports["fm_last_error_message"](self._store)),
                self._read_cstr_unlocked(self._exports["fm_last_error_context"](self._store)),
            )

    def alloc(self, size: int) -> int:
        """Allocate ``size`` bytes in WASM memory; return the pointer.

        Raises:
          MemoryError: when the WASM-side allocator returns NULL.
        """
        if size <= 0:
            return 0
        self._ensure()
        with self._call_lock:
            ptr = self._exports["malloc"](self._store, size)
        if ptr == 0:
            raise MemoryError(f"formulon: WASM malloc({size}) returned NULL")
        return ptr

    def free(self, ptr: int) -> None:
        if ptr == 0:
            return
        self._ensure()
        with self._call_lock:
            self._exports["free"](self._store, ptr)

    def alloc_utf8(self, s: str) -> Tuple[int, int]:
        """Encode ``s`` as UTF-8 and copy it into WASM memory.

        Returns ``(ptr, length_with_nul)``. The caller MUST free ``ptr``
        with :meth:`free` once the call that consumed it returns.
        """
        if not isinstance(s, str):
            raise TypeError(f"expected str, got {type(s).__name__}")
        buf = s.encode("utf-8") + b"\x00"
        ptr = self.alloc(len(buf))
        self.write_bytes(ptr, buf)
        return ptr, len(buf)

    def alloc_bytes(self, data: bytes) -> int:
        """Copy ``data`` into WASM memory; return the pointer."""
        if len(data) == 0:
            return 0
        ptr = self.alloc(len(data))
        self.write_bytes(ptr, data)
        return ptr


# ---------------------------------------------------------------------------
# Module-level singleton + thin compatibility shim
# ---------------------------------------------------------------------------


LIB = _WasmInstance()


def decode_cstr(ptr_or_bytes) -> str:
    """Backwards-compat shim for :class:`formulon.workbook.FormulonError`.

    Accepts either an int WASM pointer (decoded via :class:`LIB`) or a
    bytes object (legacy ctypes path). Returns the empty string when
    the input is ``None``, ``0``, or empty.
    """
    if ptr_or_bytes is None:
        return ""
    if isinstance(ptr_or_bytes, bytes):
        return ptr_or_bytes.decode("utf-8", errors="replace")
    if isinstance(ptr_or_bytes, int):
        return LIB.read_cstr(ptr_or_bytes)
    raise TypeError(f"decode_cstr: unexpected type {type(ptr_or_bytes).__name__}")
