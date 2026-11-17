#!/usr/bin/env python3
"""Compile the DrawingML preset shape definitions into src/io/ooxml/preset_geometry_table.inc.

Input: `presetShapeDefinitions.xml` from ECMA-376 Part 1 (Office Open XML File Formats, the
DrawingML geometries annex), given as a path. The file is not part of this tree; the generated table
records the edition and the file's SHA-256 in its header.

For each preset, in byte order of the names (the order of the `ST_ShapeType` name table in
src/io/ooxml/preset_geometry.cpp), the table keeps:
  - the adjust guides (`avLst`): name and default value;
  - the guides (`gdLst`) the connection sites depend on, pruned of the rest, as bytecode;
  - the connection sites (`cxnLst`): x, y and angle operands.

Operands are one byte, `tag << 6 | index`, or two when the index is 63 or more (`tag << 6 | 63`, then
`index - 63`). Tags: 0 a shared value (`kPresetValues`: a constant, or w, h, ss, ls divided by n),
1 an adjust guide of the preset, 2 an earlier guide of its program. An instruction is the operator
byte followed by its operands; identical programs and site lists are stored once.

  gen_preset_geometry.py --source <presetShapeDefinitions.xml> [--check]
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "src" / "io" / "ooxml" / "preset_geometry_table.inc"
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EDITION = "ECMA-376 Part 1, 5th edition (December 2016)"

# Operator byte and operand count, in the order of ECMA-376 Part 1, 20.1.9.11.
OPS = {
    "*/": (0, 3),
    "+-": (1, 3),
    "+/": (2, 3),
    "?:": (3, 3),
    "abs": (4, 1),
    "at2": (5, 2),
    "cat2": (6, 3),
    "cos": (7, 2),
    "max": (8, 2),
    "min": (9, 2),
    "mod": (10, 3),
    "pin": (11, 3),
    "sat2": (12, 3),
    "sin": (13, 2),
    "sqrt": (14, 1),
    "tan": (15, 2),
    "val": (16, 1),
}
# Value kinds (PresetValue::kind): a constant, or a shape dimension divided by `arg`.
CONST, W, H, SS, LS = 0, 1, 2, 3, 4
ANGLES = {
    "cd2": 10800000,
    "cd4": 5400000,
    "cd8": 2700000,
    "3cd4": 16200000,
    "3cd8": 8100000,
    "5cd8": 13500000,
    "7cd8": 18900000,
}


def builtin(name: str) -> tuple[int, int] | None:
    """The value kind and argument of a built-in guide name, or None."""
    if name in ANGLES:
        return CONST, ANGLES[name]
    fixed = {
        "w": (W, 1),
        "h": (H, 1),
        "ss": (SS, 1),
        "ls": (LS, 1),
        "hc": (W, 2),
        "vc": (H, 2),
        "l": (CONST, 0),
        "t": (CONST, 0),
        "r": (W, 1),
        "b": (H, 1),
    }
    if name in fixed:
        return fixed[name]
    # One published site angle names `cd3`, read by the `cdN` naming as a third of the circle.
    m = re.fullmatch(r"([1-9]?)cd(\d+)", name)
    if m:
        return CONST, int(m.group(1) or 1) * 21600000 // int(m.group(2))
    m = re.fullmatch(r"(wd|hd|ssd)(\d+)", name)
    if m:
        return {"wd": W, "hd": H, "ssd": SS}[m.group(1)], int(m.group(2))
    return None


class Compiler:
    def __init__(self) -> None:
        self.values: list[tuple[int, int]] = []
        self.adjust_names: list[str] = []
        self.adjusts: list[tuple[int, int]] = []
        self.code = bytearray()
        self.blobs: dict[bytes, int] = {}
        self.entries: list[tuple[int, int, int, int, int]] = []

    def value(self, kind: int, arg: int) -> int:
        if (kind, arg) not in self.values:
            self.values.append((kind, arg))
        return self.values.index((kind, arg))

    @staticmethod
    def operand(tag: int, index: int) -> bytes:
        if index < 63:
            return bytes([tag << 6 | index])
        if index - 63 > 255:
            raise SystemExit(f"operand index {index} does not fit")
        return bytes([tag << 6 | 63, index - 63])

    def preset(self, node: ET.Element) -> None:
        av = [(g.get("name"), g.get("fmla")) for g in node.findall(f"{NS}avLst/{NS}gd")]
        gd = [(g.get("name"), g.get("fmla")) for g in node.findall(f"{NS}gdLst/{NS}gd")]
        sites = [
            (c.find(f"{NS}pos").get("x"), c.find(f"{NS}pos").get("y"), c.get("ang"))
            for c in node.findall(f"{NS}cxnLst/{NS}cxn")
        ]
        av_slot = {name: i for i, (name, _) in enumerate(av)}

        # A guide name may be defined again later; a use reads the latest definition before it.
        def latest(name: str, before: int) -> int | None:
            for i in range(before - 1, -1, -1):
                if gd[i][0] == name:
                    return i
            return None

        needed: set[int] = set()
        stack = [(v, len(gd)) for site in sites for v in site]
        while stack:
            name, before = stack.pop()
            i = latest(name, before)
            if i is not None and i not in needed:
                needed.add(i)
                stack.extend((arg, i) for arg in gd[i][1].split()[1:])
        program = sorted(needed)
        slot = {entry: k for k, entry in enumerate(program)}

        def ref(arg: str, before: int) -> bytes:
            i = latest(arg, before)
            if i is not None:
                return self.operand(2, slot[i])
            if arg in av_slot:
                return self.operand(1, av_slot[arg])
            known = builtin(arg)
            if known is not None:
                return self.operand(0, self.value(*known))
            if re.fullmatch(r"-?\d+", arg):
                return self.operand(0, self.value(CONST, int(arg)))
            raise SystemExit(f"{node.tag}: unknown operand {arg}")

        blob = bytearray()
        for entry in program:
            op, *args = gd[entry][1].split()
            code, count = OPS[op]
            # One formula in the published file (circularArrow) carries a trailing extra operand.
            if len(args) < count:
                raise SystemExit(f"{node.tag}: {gd[entry][1]}")
            args = args[:count]
            blob.append(code)
            for arg in args:
                blob += ref(arg, entry)
        for site in sites:
            for operand in site:
                blob += ref(operand, len(gd))
        key = bytes(blob)
        if key not in self.blobs:
            self.blobs[key] = len(self.code)
            self.code += key
        adjust_first = len(self.adjusts)
        for name, fmla in av:
            op, number = fmla.split()
            if op != "val":
                raise SystemExit(f"{node.tag}: adjust default {fmla}")
            if name not in self.adjust_names:
                self.adjust_names.append(name)
            self.adjusts.append((self.adjust_names.index(name), int(number)))
        self.entries.append((adjust_first, len(av), self.blobs[key], len(program), len(sites)))


def repaired(root: ET.Element) -> list[ET.Element]:
    """The presets of `root`, with the published file's one known defect corrected.

    Every edition lists `upDownArrow` twice, identically, and has no `upArrow`. `upArrow` is
    `downArrow` mirrored top to bottom: the same guides, with the side sites at `dy1` from the top
    instead of `b - dy1`, in the top, left, bottom, right site order of the other block arrows.
    """
    presets: list[ET.Element] = []
    for node in root:
        if any(node.tag == seen.tag for seen in presets):
            if ET.tostring(node) != ET.tostring(next(seen for seen in presets if seen.tag == node.tag)):
                raise SystemExit(f"{node.tag} is defined twice with different content")
            continue
        presets.append(node)
    if not any(node.tag == "upArrow" for node in presets):
        down = next(node for node in presets if node.tag == "downArrow")
        up = copy.deepcopy(down)
        up.tag = "upArrow"
        for pos in up.findall(f"{NS}cxnLst/{NS}cxn/{NS}pos"):
            if pos.get("y") == "y1":
                pos.set("y", "dy1")
        presets.append(up)
    return presets


def render(source: Path) -> str:
    data = source.read_bytes()
    root = ET.fromstring(data)
    presets = sorted(repaired(root), key=lambda node: node.tag)
    compiler = Compiler()
    for node in presets:
        compiler.preset(node)
    if len(compiler.code) > 0xFFFF or len(compiler.adjusts) > 0xFFFF:
        raise SystemExit("table offsets are 16-bit")
    lines = [
        f"// Generated by tools/codegen/gen_preset_geometry.py from {EDITION},",
        "// presetShapeDefinitions.xml (the DrawingML geometries annex),",
        f"// SHA-256 {hashlib.sha256(data).hexdigest()}. Do not edit.",
        "// clang-format off",
        "inline constexpr PresetValue kPresetValues[] = {",
    ]
    lines += [f"    {{{kind}, {arg}}}," for kind, arg in compiler.values]
    lines += ["};", "inline constexpr std::string_view kPresetAdjustNames[] = {"]
    lines += [f'    "{name}",' for name in compiler.adjust_names]
    lines += ["};", "inline constexpr PresetAdjust kPresetAdjusts[] = {"]
    lines += [f"    {{{name}, {default}}}," for name, default in compiler.adjusts]
    lines += ["};", "inline constexpr std::uint8_t kPresetCode[] = {"]
    for i in range(0, len(compiler.code), 24):
        lines.append("    " + " ".join(f"{b}," for b in compiler.code[i : i + 24]))
    lines += [
        "};",
        "// {adjust_first, adjust_count, code_offset, guide_count, site_count}, in name order.",
        "inline constexpr PresetEntry kPresetEntries[] = {",
    ]
    for (first, count, offset, guides, sites), node in zip(compiler.entries, presets):
        lines.append(f"    {{{first}, {count}, {offset}, {guides}, {sites}}},  // {node.tag.replace(NS, '')}")
    lines += [
        "};",
        "#ifdef FORMULON_PRESET_GEOMETRY_TABLE_NAMES",
        "inline constexpr std::string_view kPresetTableNames[] = {",
    ]
    lines += [f'    "{node.tag.replace(NS, "")}",' for node in presets]
    lines += ["};", "#endif", "// clang-format on", ""]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--source", type=Path, help="presetShapeDefinitions.xml; without it --check is skipped")
    parser.add_argument("--check", action="store_true", help="fail when the checked-in table differs")
    args = parser.parse_args()
    if args.source is None:
        if args.check:
            print("gen_preset_geometry: no --source given; check skipped")
            return 0
        parser.error("--source is required to write the table")
    text = render(args.source)
    if args.check:
        if OUT.read_text(encoding="utf-8") != text:
            print(f"gen_preset_geometry: {OUT.relative_to(REPO)} is stale; rerun without --check")
            return 1
        return 0
    OUT.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
