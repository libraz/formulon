"""Tests for the Python binding's wasm32 struct marshalling primitives."""

from __future__ import annotations

import struct
import unittest

from formulon._structs import F64, I32, I64, PTR, U8, U16, U32, VALUE_BLOB, Struct


class RecordingMemory:
    def __init__(self, size: int = 256) -> None:
        self.memory = bytearray(size)
        self.writes: list[tuple[int, bytearray]] = []

    def write_bytes(self, ptr: int, data: bytearray) -> None:
        self.writes.append((ptr, data))
        self.memory[ptr : ptr + len(data)] = data

    def read_bytes(self, ptr: int, length: int) -> bytes:
        return bytes(self.memory[ptr : ptr + length])


class StructPackTest(unittest.TestCase):
    def test_all_scalar_kinds_and_signed_boundaries_use_one_write(self) -> None:
        layout = Struct(
            "all_scalars",
            [
                ("ptr", PTR),
                ("u32", U32),
                ("i32_min", I32),
                ("i32_max", I32),
                ("i64_min", I64),
                ("i64_max", I64),
                ("u16", U16),
                ("u8", U8),
                ("f64", F64),
            ],
        )
        values = {
            "ptr": 0xFFFFFFFF,
            "u32": 0xFFFFFFFF,
            "i32_min": -(2**31),
            "i32_max": 2**31 - 1,
            "i64_min": -(2**63),
            "i64_max": 2**63 - 1,
            "u16": 0xFFFF,
            "u8": 0xFF,
            "f64": -3.5,
        }
        lib = RecordingMemory()

        layout.pack(lib, 16, values)

        self.assertEqual(len(lib.writes), 1)
        ptr, payload = lib.writes[0]
        self.assertEqual(ptr, 16)
        self.assertIs(type(payload), bytearray)
        self.assertEqual(layout.unpack(lib, 16), values)

    def test_padding_omitted_fields_and_opaque_blobs_are_zeroed(self) -> None:
        layout = Struct(
            "padding_and_blob",
            [("byte", U8), ("wide", I64), ("blob", VALUE_BLOB), ("tail", U8)],
        )
        lib = RecordingMemory()
        lib.memory[16 : 16 + layout.size] = b"\xa5" * layout.size

        layout.pack(lib, 16, {"byte": 7, "tail": 9})

        raw = lib.read_bytes(16, layout.size)
        self.assertEqual(raw[layout.offsets["byte"][1]], 7)
        self.assertEqual(raw[layout.offsets["tail"][1]], 9)
        wide_start = layout.offsets["wide"][1]
        self.assertEqual(raw[1:wide_start], b"\x00" * (wide_start - 1))
        blob_start = layout.offsets["blob"][1]
        self.assertEqual(raw[blob_start : blob_start + VALUE_BLOB[1]], b"\x00" * VALUE_BLOB[1])
        tail_start = layout.offsets["tail"][1]
        self.assertEqual(raw[tail_start + 1 :], b"\x00" * (layout.size - tail_start - 1))
        self.assertEqual(layout.unpack(lib, 16), {"byte": 7, "wide": 0, "tail": 9})

    def test_pack_errors_do_not_partially_write(self) -> None:
        layout = Struct("errors", [("value", I32), ("blob", VALUE_BLOB)])
        for values, expected in (
            ({"missing": 1}, KeyError),
            ({"blob": b"\x00" * VALUE_BLOB[1]}, KeyError),
            ({"value": 2**31}, struct.error),
        ):
            lib = RecordingMemory()
            with self.subTest(values=values):
                with self.assertRaises(expected):
                    layout.pack(lib, 16, values)
                self.assertEqual(lib.writes, [])


if __name__ == "__main__":
    unittest.main()
