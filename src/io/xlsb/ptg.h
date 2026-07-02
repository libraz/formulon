// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB Ptg (Parse Tag) dispatch table. The XLSB binary formula stream
// is a sequence of single-byte-tagged tokens; this module provides the
// O(log N) lookup that the reader's Ptg decoder uses to branch on each
// tag, plus tiny helpers for decoding the class bits that distinguish
// Reference / Value / Array variants of class-marked Ptgs.
//
// The full enumeration summarises [MS-XLSB] §2.5.97. Class-marked Ptgs
// (the `0x20`/`0x40`/`0x60` trio) collapse into a single `PtgKind` here;
// the `class_from_byte` helper extracts the class so callers do not have
// to duplicate the bit-twiddling per call site.
//
// This file ships only the dispatch table — no Ptg → AST conversion is
// performed here. `ptg_reader.{h,cpp}` consumes this table to build AST
// nodes (decode) and `ptg_writer.{h,cpp}` performs the inverse (encode).

#ifndef FORMULON_IO_XLSB_PTG_H_
#define FORMULON_IO_XLSB_PTG_H_

#include <array>
#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

/// Implementation status of a Ptg in the v1.0 reader/writer.
///
///   * `Full`             — round-trips Reader + Writer (Ptg ↔ AST).
///   * `PreserveOnly`     — Reader reads opaque bytes, Writer must
///                          re-emit them verbatim (no AST round-trip).
///   * `Unsupported`      — Reader surfaces `#NAME?` + structured-log
///                          diagnostic; Writer never emits.
enum class PtgStatus : std::uint8_t {
  Full = 0,
  PreserveOnly = 1,
  Unsupported = 2,
};

/// Class mark carried by class-marked Ptgs (`0x20`/`0x40`/`0x60` trio).
///
/// The class bits live in positions 5..6 of the first byte; they decide
/// how the operand is consumed by the surrounding expression context
/// (e.g. a `PtgRef` inside a `SUM(...)` is value-class because the
/// enclosing function expects a value, while the same `PtgRef` standing
/// alone evaluates to a reference).
enum class PtgClass : std::uint8_t {
  Reference = 0,  ///< `0x20`-marked: passed through as a reference.
  Value = 1,      ///< `0x40`-marked: dereferenced to a scalar value.
  Array = 2,      ///< `0x60`-marked: dereferenced to an array.
};

/// Logical Ptg identity. Class-marked Ptgs collapse the
/// `0x20`/`0x40`/`0x60` trio into a single value; use `class_from_byte`
/// on the raw first byte to recover the class.
///
/// Names mirror the [MS-XLSB] §2.5.97 enumeration. The enum is
/// intentionally opaque w.r.t. the wire byte so callers reach for
/// `PtgInfo::base_byte` rather than bit-twiddling the kind.
enum class PtgKind : std::uint8_t {
  Unknown = 0,

  // ---- Operators (no class mark) ------------------------------------------
  Exp,      ///< 0x01 — shared/array formula reference (si head).
  Tbl,      ///< 0x02 — data-table formula (preserve only).
  Add,      ///< 0x03 — binary `+`.
  Sub,      ///< 0x04 — binary `-`.
  Mul,      ///< 0x05 — `*`.
  Div,      ///< 0x06 — `/`.
  Power,    ///< 0x07 — `^`.
  Concat,   ///< 0x08 — `&`.
  Lt,       ///< 0x09 — `<`.
  Le,       ///< 0x0A — `<=`.
  Eq,       ///< 0x0B — `=`.
  Ge,       ///< 0x0C — `>=`.
  Gt,       ///< 0x0D — `>`.
  Ne,       ///< 0x0E — `<>`.
  Isect,    ///< 0x0F — range intersection (space).
  Union,    ///< 0x10 — range union (`,`).
  Range,    ///< 0x11 — range operator (`:`).
  Uplus,    ///< 0x12 — unary `+`.
  Uminus,   ///< 0x13 — unary `-`.
  Percent,  ///< 0x14 — postfix `%`.
  Paren,    ///< 0x15 — parenthesis (writer-only).
  MissArg,  ///< 0x16 — omitted argument (e.g. `IF(,x,y)`).
  Str,      ///< 0x17 — string literal.
  ElfLel,   ///< 0x18 — reserved (XLM); unsupported.
  Attr,     ///< 0x19 — function-call attribute (sub-kinded).
  Err,      ///< 0x1C — error literal (`#DIV/0!` …).
  Bool,     ///< 0x1D — TRUE / FALSE.
  Int,      ///< 0x1E — 16-bit unsigned int (0..65535).
  Num,      ///< 0x1F — IEEE 754 double.

  // ---- Class-marked Ptgs (0x20/0x40/0x60 trio collapsed) ------------------
  Array,      ///< 0x20/0x40/0x60 — array literal `{1,2;3,4}`.
  Func,       ///< 0x21/0x41/0x61 — fixed-arity built-in (id + class).
  FuncVar,    ///< 0x22/0x42/0x62 — variable-arity built-in.
  Name,       ///< 0x23/0x43/0x63 — DefinedName reference.
  Ref,        ///< 0x24/0x44/0x64 — single-cell reference.
  Area,       ///< 0x25/0x45/0x65 — range reference.
  MemArea,    ///< 0x26/0x46/0x66 — constant range (union).
  MemErr,     ///< 0x27/0x47/0x67 — error-form MemArea.
  MemNoMem,   ///< 0x28/0x48/0x68 — out-of-mem placeholder (preserve).
  MemFunc,    ///< 0x29/0x49/0x69 — function-returns-range marker.
  RefErr,     ///< 0x2A/0x4A/0x6A — `#REF!`-form Ref.
  AreaErr,    ///< 0x2B/0x4B/0x6B — `#REF!`-form Area.
  RefN,       ///< 0x2C/0x4C/0x6C — relative ref (shared formulas).
  AreaN,      ///< 0x2D/0x4D/0x6D — relative range (shared formulas).
  MemAreaN,   ///< 0x2E/0x4E/0x6E — relative MemArea.
  MemNoMemN,  ///< 0x2F/0x4F/0x6F — reserved (preserve).
  NameX,      ///< 0x39/0x59/0x79 — external-book DefinedName.
  Ref3d,      ///< 0x3A/0x5A/0x7A — 3D single-cell reference.
  Area3d,     ///< 0x3B/0x5B/0x7B — 3D range reference.
  RefErr3d,   ///< 0x3C/0x5C/0x7C — `#REF!`-form 3D ref.
  AreaErr3d,  ///< 0x3D/0x5D/0x7D — `#REF!`-form 3D area.

  // ---- Extension Ptgs (0xE0..0xFF) ----------------------------------------
  FuncCE,    ///< 0xE0 — reserved; unsupported.
  FuncVar2,  ///< 0xE1 — reserved; unsupported.
  IfError,   ///< 0xEA — IFERROR optimisation marker (transparent).
};

/// PtgAttr (0x19) sub-kind. Stored as a separate byte after the 0x19
/// dispatch byte. Values come from [MS-XLSB] §2.5.97.46 (PtgAttr).
enum class PtgAttrKind : std::uint8_t {
  Semi = 0x01,       ///< Volatile marker (`SEMI`).
  If = 0x02,         ///< IF branch offset.
  Choose = 0x04,     ///< CHOOSE jump table.
  Goto = 0x08,       ///< Unconditional jump.
  Sum = 0x10,        ///< Unary SUM optimisation.
  Baxcel = 0x20,     ///< Baseline reference (preserve).
  Space = 0x40,      ///< Inter-token whitespace (round-trip).
  SpaceSemi = 0x41,  ///< Whitespace + volatile.
};

/// One row of the dispatch table.
///
///   * `kind`       — logical Ptg identity (class-marked variants
///                    collapse to one entry, indexed by the
///                    Reference-class first byte).
///   * `base_byte`  — the Reference-class (or unclassed) first byte; the
///                    wire byte for the Value / Array variants is
///                    `base_byte | 0x20` and `base_byte | 0x40`
///                    respectively.
///   * `name`       — short symbolic name for diagnostics (e.g. `"Ref"`).
///                    The pointer references a static string literal
///                    with program lifetime.
///   * `status`     — v1.0 implementation status (see `PtgStatus`).
struct PtgInfo {
  PtgKind kind;
  std::uint8_t base_byte;
  const char* name;
  PtgStatus status;
};

/// Number of rows in the dispatch table. Exposed for tests so they can
/// assert the table is sorted without hard-coding the count.
inline constexpr std::size_t kPtgInfoCount = 51;

/// The full Ptg dispatch table, sorted ascending by `base_byte` for
/// O(log N) binary-search lookup. Defined `constexpr` in the header so
/// the static-assert that the table is sorted can run at compile time
/// from the implementation TU.
inline constexpr std::array<PtgInfo, kPtgInfoCount> kPtgInfoTable = {{
    // ---- Operators (no class mark) -----------------------------------------
    // `Exp` (shared/array-formula shell): the reader does not decode this
    // token directly -- `decode_ptgs` has no `PtgKind::Exp` case, so a
    // cell whose `rgce` is a bare `PtgExp` logs
    // `xlsb.formula.not_decoded` and keeps its cached value only. The
    // *cell*-level workaround (`DecodeSheetBin`'s `BrtArrFmla` handling,
    // which supplies the array/CSE formula's real Ptg tokens out-of-band
    // and registers the spill) recovers the correct formula text and
    // value despite this, but that is a record-layer mechanism, not a
    // `PtgKind::Exp` decode.
    {PtgKind::Exp, 0x01, "Exp", PtgStatus::Unsupported},
    {PtgKind::Tbl, 0x02, "Tbl", PtgStatus::Unsupported},
    {PtgKind::Add, 0x03, "Add", PtgStatus::Full},
    {PtgKind::Sub, 0x04, "Sub", PtgStatus::Full},
    {PtgKind::Mul, 0x05, "Mul", PtgStatus::Full},
    {PtgKind::Div, 0x06, "Div", PtgStatus::Full},
    {PtgKind::Power, 0x07, "Power", PtgStatus::Full},
    {PtgKind::Concat, 0x08, "Concat", PtgStatus::Full},
    {PtgKind::Lt, 0x09, "Lt", PtgStatus::Full},
    {PtgKind::Le, 0x0A, "Le", PtgStatus::Full},
    {PtgKind::Eq, 0x0B, "Eq", PtgStatus::Full},
    {PtgKind::Ge, 0x0C, "Ge", PtgStatus::Full},
    {PtgKind::Gt, 0x0D, "Gt", PtgStatus::Full},
    {PtgKind::Ne, 0x0E, "Ne", PtgStatus::Full},
    {PtgKind::Isect, 0x0F, "Isect", PtgStatus::Full},
    {PtgKind::Union, 0x10, "Union", PtgStatus::Full},
    {PtgKind::Range, 0x11, "Range", PtgStatus::Full},
    {PtgKind::Uplus, 0x12, "Uplus", PtgStatus::Full},
    {PtgKind::Uminus, 0x13, "Uminus", PtgStatus::Full},
    {PtgKind::Percent, 0x14, "Percent", PtgStatus::Full},
    // Neither the reader nor the writer handles `Paren` despite the
    // "writer-only" framing in `PtgKind::Paren`'s doc comment: source
    // parenthesisation is captured structurally (via `AstNode` operator
    // precedence / explicit grouping), and the writer never re-derives
    // a standalone `PtgParen` byte from that structure.
    {PtgKind::Paren, 0x15, "Paren", PtgStatus::Unsupported},
    {PtgKind::MissArg, 0x16, "MissArg", PtgStatus::Full},
    {PtgKind::Str, 0x17, "Str", PtgStatus::Full},
    {PtgKind::ElfLel, 0x18, "ElfLel", PtgStatus::Unsupported},
    // Reader-only: `decode_ptgs` handles every `PtgAttr` sub-kind
    // (Sum/Space/SpaceSemi/If/Choose/Goto/Semi/Baxcel), but the writer
    // never emits `0x19` -- it re-derives the semantically equivalent
    // canonical form instead (a single-arg `SUM(x)` call node for
    // `PtgAttrSum`; whitespace fidelity for `PtgAttrSpace` is not
    // preserved on write). `Full` still holds because both directions
    // of the round-trip are covered, just through different wire forms.
    {PtgKind::Attr, 0x19, "Attr", PtgStatus::Full},
    {PtgKind::Err, 0x1C, "Err", PtgStatus::Full},
    {PtgKind::Bool, 0x1D, "Bool", PtgStatus::Full},
    {PtgKind::Int, 0x1E, "Int", PtgStatus::Full},
    {PtgKind::Num, 0x1F, "Num", PtgStatus::Full},

    // ---- Class-marked Ptgs (Reference-class base bytes 0x20..0x3D) --------
    {PtgKind::Array, 0x20, "Array", PtgStatus::Full},
    {PtgKind::Func, 0x21, "Func", PtgStatus::Full},
    {PtgKind::FuncVar, 0x22, "FuncVar", PtgStatus::Full},
    {PtgKind::Name, 0x23, "Name", PtgStatus::Full},
    {PtgKind::Ref, 0x24, "Ref", PtgStatus::Full},
    {PtgKind::Area, 0x25, "Area", PtgStatus::Full},
    // `MemArea` / `MemErr` / `MemNoMem` / `MemFunc` / `RefN` / `AreaN` /
    // `MemAreaN` / `MemNoMemN` / `NameX`: none of these has a
    // `decode_ptgs` case (falls through to the `default:` ->
    // `unsupported_ptg` branch) or an `encode_ptgs` emission site.
    {PtgKind::MemArea, 0x26, "MemArea", PtgStatus::Unsupported},
    {PtgKind::MemErr, 0x27, "MemErr", PtgStatus::Unsupported},
    {PtgKind::MemNoMem, 0x28, "MemNoMem", PtgStatus::Unsupported},
    {PtgKind::MemFunc, 0x29, "MemFunc", PtgStatus::Unsupported},
    // `RefErr` / `AreaErr` / `RefErr3d` / `AreaErr3d`: the reader decodes
    // these to an `#REF!` `ErrorLiteral` node (see `ptg_reader.cpp`'s
    // combined `RefErr`/`RefErr3d` and `AreaErr`/`AreaErr3d` cases). The
    // writer never emits them because the AST never distinguishes "a
    // reference that is `#REF!`" from a plain `#REF!` error literal --
    // both directions of the round-trip are covered, just through two
    // different Ptg bytes (`PtgErr`, 0x1C, on write), so `Full` reflects
    // the actual read+write contract despite the asymmetric wire form.
    {PtgKind::RefErr, 0x2A, "RefErr", PtgStatus::Full},
    {PtgKind::AreaErr, 0x2B, "AreaErr", PtgStatus::Full},
    {PtgKind::RefN, 0x2C, "RefN", PtgStatus::Unsupported},
    {PtgKind::AreaN, 0x2D, "AreaN", PtgStatus::Unsupported},
    {PtgKind::MemAreaN, 0x2E, "MemAreaN", PtgStatus::Unsupported},
    {PtgKind::MemNoMemN, 0x2F, "MemNoMemN", PtgStatus::Unsupported},
    {PtgKind::NameX, 0x39, "NameX", PtgStatus::Unsupported},
    {PtgKind::Ref3d, 0x3A, "Ref3d", PtgStatus::Full},
    {PtgKind::Area3d, 0x3B, "Area3d", PtgStatus::Full},
    {PtgKind::RefErr3d, 0x3C, "RefErr3d", PtgStatus::Full},
    {PtgKind::AreaErr3d, 0x3D, "AreaErr3d", PtgStatus::Full},

    // ---- Extension Ptgs (0xE0..0xFF) ---------------------------------------
    {PtgKind::IfError, 0xEA, "IfError", PtgStatus::Full},
}};

/// Returns the table row whose `base_byte` matches the class-stripped
/// first byte of the wire token, or `nullptr` when no row matches. The
/// caller is responsible for handing in the *reference-class* base byte
/// (i.e. for class-marked Ptgs, `first_byte & ~0x60`); use
/// `lookup_ptg_from_wire` if the raw wire byte is what's available.
const PtgInfo* lookup_ptg(std::uint8_t base_byte);

/// Convenience over `lookup_ptg` that also handles the class-marked
/// trio (0x20/0x40/0x60) by stripping the class bits before lookup.
/// Bytes outside the table return `nullptr`.
const PtgInfo* lookup_ptg_from_wire(std::uint8_t first_byte);

/// Returns the class encoded in the upper bits of the wire byte for a
/// class-marked Ptg. Caller is responsible for ensuring the byte
/// corresponds to a class-marked Ptg (`is_class_marked` returns true);
/// for unclassed bytes this still returns `Reference` (0x00 class bits)
/// which is a harmless default.
PtgClass class_from_byte(std::uint8_t first_byte);

/// Whether the Ptg kind is class-marked (i.e. its wire byte has variants
/// at `+0x00` / `+0x20` / `+0x40` / `+0x60`).
bool is_class_marked(PtgKind kind);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_H_
