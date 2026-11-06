"""Public API naming / shape conventions, checked against a per-surface baseline.

Every rule is a function registered under a rule-id with `@rule`. A rule
reads a `Context` (the parsed C header, the name ledger, the repo root) and
yields `Violation`s. A violation's identity is `(rule-id, op-key,
discriminator)` on one surface; its message is informational only.

The checker is RED on any violation missing from the baseline and on any
baseline line no longer violated, so a baseline can only shrink. Baseline
files live in `tools/dev/api_convention_baseline/{c,wasm,native,py}.txt`,
one `<rule-id>\\t<op-key>\\t<discriminator>\\t<message>` line per violation,
no comments; a missing file reads as empty.

The op-key is the ledger key (`tools/api/names.yaml`) of the symbol, found
through its current or former name, so a rename does not move its baseline
lines; without a ledger entry it is the symbol itself. Rules that classify a
function by its name (verb, `_count`, `_destroy`) read the ledger's target
name, while the naming rules read the name in the tree.

Further rule modules register into the same registry: the sibling modules
`api_conventions_bindings.py` and `api_conventions_shapes.py`, if present,
are imported before rules run.
"""

from __future__ import annotations

import importlib
import importlib.util
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, List, NamedTuple, Optional, Sequence, Tuple

from .c_header import _C_FUNCTION_DECL_RE, _C_STRUCT_RE, _split_c_params
from .surface_files import CAPI_HEADER, REPO_ROOT

BASELINE_DIR = Path(__file__).resolve().parents[1] / "api_convention_baseline"
LEDGER_PATH = REPO_ROOT / "tools" / "api" / "names.yaml"
SURFACES = ("c", "wasm", "native", "py")
NO_DISCRIMINATOR = "-"

# ---------------------------------------------------------------------------
# Vocabulary.
# ---------------------------------------------------------------------------

# Scopes a C function name may carry after `fm_`; longest match wins.
SCOPES = (
    "workbook",
    "sheet",
    "cell_range",
    "cf_results",
    "cell_nodes",
    "pivot_cells",
    "pagination",
    "chart",
    "chart_data",
    "value_array",
    "buffer",
)
_SCOPES_LONGEST_FIRST = tuple(sorted(SCOPES, key=len, reverse=True))

# The closed verb vocabulary. A new verb lands in the same change as its use.
C_VERBS = frozenset(
    {
        "get", "set", "is", "has", "clear",
        "add", "update", "remove", "insert", "delete",
        "create", "destroy",
        "copy", "move", "rename", "load", "save", "recalc", "evaluate",
        "parse", "format", "convert", "resolve",
        "snapshot", "restore", "remap", "apply", "refresh", "validate", "match", "next",
        "localize", "canonicalize", "compare", "group", "ungroup", "project", "show",
    }
)  # fmt: skip
# Verbs that exist only on the binding surfaces.
BINDING_ONLY_VERBS = frozenset({"list"})

# Word casing exceptions used by the binding-name derivation.
ACRONYMS = {"r1c1": "R1C1", "a1": "A1", "xml": "Xml", "id": "Id", "xf": "Xf", "dxf": "Dxf", "cf": "Cf"}

# Process functions allowed to return a static or thread-local `const char*`.
PROCESS_STRING_FUNCTIONS = frozenset(
    {
        "fm_get_version",
        "fm_get_status_name",
        "fm_get_last_error_message",
        "fm_get_last_error_context",
        "fm_get_error_display_name",
    }
)

# Process functions allowed to return a plain numeric constant.
PROCESS_NUMBER_FUNCTIONS = frozenset({"fm_get_version_number"})

# The lifetime classes a `const char**` output's doc must name.
LIFETIME_CLASSES = ("model-backed", "read-scratch-backed", "owned snapshot", "static", "thread-local view")
_LIFETIME_RE = re.compile(
    r"(?<![\w-])(?:" + "|".join(re.escape(c).replace(r"\ ", r"\s+") for c in LIFETIME_CLASSES) + r")(?![\w-])",
    re.IGNORECASE,
)

CELL_RECT = "fm_cell_rect_t"
CELL_RECT_FIELDS = ("first_row", "first_col", "last_row", "last_col")
# Coordinate quadruples that spell a rectangle without `fm_cell_rect_t`.
_RECT_QUADS = (
    CELL_RECT_FIELDS,
    ("row", "col", "last_row", "last_col"),
    ("row0", "col0", "row1", "col1"),
    ("top", "left", "rows", "cols"),
    ("row", "col", "rows", "cols"),
    ("anchor_row", "anchor_col", "rows", "cols"),
    ("anchor_row", "anchor_col", "span_rows", "span_cols"),
)
_INTEGER_TYPES = frozenset({"uint8_t", "uint16_t", "uint32_t", "uint64_t", "int32_t", "int64_t", "size_t"})
_BARE_C_SCALARS = frozenset({"int", "unsigned", "unsigned int", "long", "short", "bool", "_Bool"})
_ID_PARAMS = frozenset({"cache_id", "num_fmt_id", "object_id"})
_ROW_COL_PARAMS = frozenset({"row", "col", "first_row", "last_row", "first_col", "last_col"})

_C_NAME_RE = re.compile(r"fm_[a-z0-9]+(?:_[a-z0-9]+)*")
_C_TYPEDEF_NAME_RE = re.compile(r"fm_[a-z0-9]+(?:_[a-z0-9]+)*_t")

# ---------------------------------------------------------------------------
# Per-name API.
# ---------------------------------------------------------------------------


def split_c_name(c_name: str) -> Tuple[Optional[str], List[str]]:
    """Splits `fm_<scope>_<rest>` into `(scope or None, rest words)`."""
    body = c_name[3:] if c_name.startswith("fm_") else c_name
    for scope in _SCOPES_LONGEST_FIRST:
        if body == scope:
            return scope, []
        if body.startswith(scope + "_"):
            return scope, body[len(scope) + 1 :].split("_")
    return None, body.split("_") if body else []


def c_verb(c_name: str) -> Optional[str]:
    """The word in verb position of a C function name, or None."""
    _, words = split_c_name(c_name)
    return words[0] if words else None


def check_c_name(c_name: str) -> List[Tuple[str, str]]:
    """Name-only problems of one C function name, as `(rule-id, message)` pairs.

    Covers `R1.form` (shape `fm_[<scope>_]<verb>[_<noun>...]`, lower snake
    case) and `R2.verb` (the verb is in the closed C vocabulary). A name that
    fails the form is not also judged on its verb.
    """
    if not _C_NAME_RE.fullmatch(c_name):
        return [("R1.form", f"{c_name} is not lower snake_case under the `fm_` prefix")]
    scope, words = split_c_name(c_name)
    if scope is None and (not words or (words[0] not in C_VERBS and words[0] not in BINDING_ONLY_VERBS)):
        head = words[0] if words else ""
        return [
            (
                "R1.form",
                f"{c_name}: `{head}` after `fm_` is neither a scope nor a verb",
            )
        ]
    if not words:
        return [("R1.form", f"{c_name}: no verb follows the `{scope}` scope")]
    if words[0] not in C_VERBS:
        hint = " (binding-only)" if words[0] in BINDING_ONLY_VERBS else ""
        return [("R2.verb", f"{c_name}: `{words[0]}`{hint} is not in the C verb vocabulary")]
    return []


def _cap(word: str) -> str:
    return ACRONYMS.get(word, word[:1].upper() + word[1:])


def derive_py(c_name: str) -> str:
    """Python name of a C function: the part after `fm_[<scope>_]`, snake_case."""
    _, words = split_c_name(c_name)
    return "_".join(words)


def derive_js(c_name: str) -> str:
    """JS name of a C function: the part after `fm_[<scope>_]`, lowerCamelCase."""
    _, words = split_c_name(c_name)
    if not words:
        return ""
    return words[0] + "".join(_cap(word) for word in words[1:])


_ROLE_SUFFIX = {"input": "Input", "patch": "Patch"}


def derive_type(c_typedef: str, role: Optional[str] = None) -> str:
    """JS / Python record or enum name of a C typedef: `fm_<name>_t` -> PascalCase.

    A record of role `input` / `patch` takes an `Input` / `Patch` suffix unless the name already ends with it.
    """
    body = c_typedef[3:] if c_typedef.startswith("fm_") else c_typedef
    if body.endswith("_t"):
        body = body[:-2]
    name = "".join(_cap(word) for word in body.split("_") if word)
    suffix = _ROLE_SUFFIX.get(role or "", "")
    return name if name.endswith(suffix) else name + suffix


# Record-name suffixes that add no meaning (R21.name); a role suffix after them is kept.
MEANINGLESS_SUFFIXES = ("Entry", "Record", "Info", "Spec")


def strip_record_suffix(name: str) -> str:
    """`name` without a trailing `Entry` / `Record` / `Info` / `Spec`, keeping an `Input` / `Patch` after it."""
    role = next((r for r in _ROLE_SUFFIX.values() if name.endswith(r) and name != r), "")
    base = name[: len(name) - len(role)]
    for suffix in MEANINGLESS_SUFFIXES:
        if base.endswith(suffix) and base != suffix:
            return base[: -len(suffix)] + role
    return name


def derive_record_type(c_typedef: str, role: Optional[str] = None) -> str:
    """R21 record name of a C typedef: `derive_type` without a meaningless suffix."""
    return strip_record_suffix(derive_type(c_typedef, role))


# ---------------------------------------------------------------------------
# Ledger.
# ---------------------------------------------------------------------------


class Ledger:
    """The parsed name ledger; empty when `tools/api/names.yaml` is absent."""

    def __init__(self, data: Optional[dict] = None) -> None:
        data = data or {}
        self.data = data
        self.entries: Dict[str, dict] = dict(data.get("entries") or {})
        self.collections: Dict[str, dict] = dict(data.get("collections") or {})
        self.problems: List[str] = []
        self._index: Dict[str, Dict[str, str]] = {}
        self.c_param_counts: Dict[str, int] = {}

    def _handed_over(self, a: str, b: str, name: str) -> bool:
        """One entry renames C function `name` away and the other takes the name, with another parameter count."""
        for giver, taker in ((a, b), (b, a)):
            gave = (self.entry(giver).get("was") or {}).get("c")
            took = (self.entry(taker).get("was") or {}).get("c")
            # Before the rename the header declares both former names, after it both current ones.
            if (took or "") in self.c_param_counts:
                counts = (self.c_param_counts.get(gave or ""), self.c_param_counts.get(took or ""))
            else:
                counts = (self.c_param_counts.get(self.entry(giver).get("c") or ""), self.c_param_counts.get(name))
            if (
                gave == name != self.entry(giver).get("c")
                and self.entry(taker).get("c") == name
                and None not in counts
                and counts[0] != counts[1]
            ):
                return True
        return False

    @property
    def present(self) -> bool:
        return bool(self.data)

    def key_for(self, name_field: str, symbol: str) -> Optional[str]:
        """Ledger key whose `<name_field>` or `was.<name_field>` names `symbol`; a current name wins a hand-over.

        A `kind: field` entry names a member of its record, not a symbol, so it never claims a name here.
        """
        index = self._index.get(name_field)
        if index is None:
            index = {}
            claims = []
            for key, entry in self.entries.items():
                if not isinstance(entry, dict) or entry.get("kind") == "field":
                    continue
                was = entry.get("was") if isinstance(entry.get("was"), dict) else {}
                claims.append((0, key, entry.get(name_field)))
                claims.append((1, key, was.get(name_field)))
            for _order, key, name in sorted(claims, key=lambda c: c[0]):
                if isinstance(name, str):
                    owner = index.setdefault(name, key)
                    if owner != key and not (name_field == "c" and self._handed_over(owner, key, name)):
                        self.problems.append(
                            f"api-conventions: ledger entries `{owner}` and `{key}` both claim "
                            f"{name_field} name `{name}`"
                        )
            self._index[name_field] = index
        return index.get(symbol)

    def entry(self, key: Optional[str]) -> dict:
        found = self.entries.get(key) if key is not None else None
        return found if isinstance(found, dict) else {}


def load_ledger(path: Path = LEDGER_PATH) -> Ledger:
    """Reads the ledger; an absent file is an empty ledger, an unreadable one a problem."""
    if not path.is_file():
        return Ledger()
    try:
        import yaml  # type: ignore
    except ImportError:
        ledger = Ledger()
        ledger.problems.append(
            f"api-conventions: {path} exists but PyYAML is not importable; run under "
            "tools/oracle/.venv or `pip install pyyaml`"
        )
        return ledger
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or data.get("schema") != 1 or not isinstance(data.get("entries"), dict):
        ledger = Ledger()
        ledger.problems.append(f"api-conventions: {path} is not a `schema: 1` ledger with an `entries:` map")
        return ledger
    return Ledger(data)


# ---------------------------------------------------------------------------
# C header model.
# ---------------------------------------------------------------------------


class CParam(NamedTuple):
    ctype: str
    name: str

    @property
    def role_name(self) -> str:
        """The name without an `out_` prefix."""
        return self.name[4:] if self.name.startswith("out_") else self.name

    @property
    def base_type(self) -> str:
        return self.ctype.replace("const ", "").replace("*", "").strip()

    @property
    def pointer_depth(self) -> int:
        return self.ctype.count("*")


class CFunction(NamedTuple):
    name: str
    ret: str
    params: Tuple[CParam, ...]
    doc: str
    line: int


class CTypedef(NamedTuple):
    name: str
    kind: str  # struct | enum | opaque | scalar
    fields: Tuple[CParam, ...]
    line: int


_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_ENUM_TYPEDEF_RE = re.compile(r"typedef\s+enum\s*\{[^}]*\}\s*([A-Za-z_][A-Za-z0-9_]*)\s*;")
_OPAQUE_TYPEDEF_RE = re.compile(r"typedef\s+struct\s+[A-Za-z_][A-Za-z0-9_]*\s+([A-Za-z_][A-Za-z0-9_]*)\s*;")
_SCALAR_TYPEDEF_RE = re.compile(r"typedef\s+(?:u?int(?:8|16|32|64)_t|size_t|double)\s+([A-Za-z_][A-Za-z0-9_]*)\s*;")
_UNION_OPEN_RE = re.compile(r"union\s*\{")
_UNION_CLOSE_RE = re.compile(r"\}\s*[A-Za-z_][A-Za-z0-9_]*\s*;")


def _parse_param(text: str) -> Optional[CParam]:
    text = re.sub(r"\[[^\]]*\]", "", " ".join(text.split()))
    if not text or text == "void":
        return None
    tokens = text.replace("*", "* ").split()
    name = tokens[-1]
    ctype = " ".join(tokens[:-1]).replace("* ", "*").replace(" *", "*")
    return CParam(ctype, name)


def _struct_fields(body: str) -> Tuple[CParam, ...]:
    body = _UNION_CLOSE_RE.sub(";", _UNION_OPEN_RE.sub("", _COMMENT_RE.sub(" ", body)))
    fields: List[CParam] = []
    for declaration in body.split(";"):
        declaration = " ".join(declaration.split())
        if not declaration:
            continue
        first, *rest = declaration.split(",")
        head = _parse_param(first)
        if head is None:
            continue
        fields.append(head)
        for extra in rest:
            name = re.sub(r"\[[^\]]*\]", "", extra).strip()
            if name:
                fields.append(CParam(head.ctype, name))
    return tuple(fields)


def _doc_before(header: str, start: int) -> str:
    """The block comment ending right before `start`, without comment markers."""
    head = header[:start].rstrip()
    if not head.endswith("*/"):
        return ""
    open_at = head.rfind("/*")
    if open_at < 0:
        return ""
    text = head[open_at + 2 : -2]
    return " ".join(line.strip().lstrip("*").strip() for line in text.splitlines())


def _line_of(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def parse_header(header: str) -> Tuple[List[CFunction], List[CTypedef]]:
    functions: List[CFunction] = []
    for match in _C_FUNCTION_DECL_RE.finditer(header):
        ret, name, params = match.groups()
        params = " ".join(params.split())
        parsed = [] if params in ("", "void") else [_parse_param(p) for p in _split_c_params(params)]
        functions.append(
            CFunction(
                name=name,
                ret=" ".join(ret.split()),
                params=tuple(p for p in parsed if p is not None),
                doc=_doc_before(header, match.start()),
                line=_line_of(header, match.start()),
            )
        )
    typedefs: List[CTypedef] = []
    for match in _C_STRUCT_RE.finditer(header):
        typedefs.append(
            CTypedef(match.group(2), "struct", _struct_fields(match.group(1)), _line_of(header, match.start()))
        )
    for regex, kind in ((_ENUM_TYPEDEF_RE, "enum"), (_OPAQUE_TYPEDEF_RE, "opaque"), (_SCALAR_TYPEDEF_RE, "scalar")):
        for match in regex.finditer(header):
            typedefs.append(CTypedef(match.group(1), kind, (), _line_of(header, match.start())))
    typedefs.sort(key=lambda t: t.line)
    return functions, typedefs


# ---------------------------------------------------------------------------
# Framework.
# ---------------------------------------------------------------------------


class Violation(NamedTuple):
    rule: str
    surface: str
    op_key: str
    discriminator: str
    message: str
    site: str = ""

    @property
    def identity(self) -> Tuple[str, str, str, str]:
        return (self.surface, self.rule, self.op_key, self.discriminator)


@dataclass
class Context:
    """What a rule reads. Built from the tree, or from fixture text in tests."""

    header: str
    ledger: Ledger = field(default_factory=Ledger)
    repo_root: Path = REPO_ROOT
    header_name: str = CAPI_HEADER.name

    def __post_init__(self) -> None:
        self.functions, self.typedefs = parse_header(self.header)
        self.ledger.c_param_counts = {fn.name: len(fn.params) for fn in self.functions}
        self.struct_names = {t.name for t in self.typedefs if t.kind == "struct"}
        self.enum_names = {t.name for t in self.typedefs if t.kind == "enum"}

    @classmethod
    def from_tree(cls, repo_root: Path = REPO_ROOT) -> "Context":
        header_path = repo_root / CAPI_HEADER.relative_to(REPO_ROOT)
        ledger_path = repo_root / LEDGER_PATH.relative_to(REPO_ROOT)
        return cls(header_path.read_text(encoding="utf-8"), load_ledger(ledger_path), repo_root, header_path.name)

    def op_key(self, c_symbol: str) -> str:
        return self.ledger.key_for("c", c_symbol) or c_symbol

    def target_c(self, c_symbol: str) -> str:
        """The ledger's C name for `c_symbol`, else `c_symbol` itself."""
        name = self.ledger.entry(self.ledger.key_for("c", c_symbol)).get("c")
        return name if isinstance(name, str) else c_symbol

    def site(self, line: int) -> str:
        return f"{self.header_name}:{line}"

    def c_violation(self, rule_id: str, symbol: str, discriminator: str, message: str, line: int) -> Violation:
        return Violation(rule_id, "c", self.op_key(symbol), discriminator, message, self.site(line))


RuleFn = Callable[[Context], Iterable[Violation]]
_RULES: Dict[str, RuleFn] = {}
_EXTENSIONS = ("api_conventions_bindings", "api_conventions_shapes")


def rule(rule_id: str) -> Callable[[RuleFn], RuleFn]:
    """Registers a rule under `rule_id`; its docstring states what it requires."""

    def register(fn: RuleFn) -> RuleFn:
        if rule_id in _RULES and _RULES[rule_id] is not fn:
            raise ValueError(f"rule {rule_id} registered twice")
        _RULES[rule_id] = fn
        return fn

    return register


def _natural_key(rule_id: str) -> List[object]:
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", rule_id)]


def all_rules() -> Dict[str, RuleFn]:
    """Every registered rule, after importing the sibling rule modules."""
    for module in _EXTENSIONS:
        name = f"{__package__}.{module}"
        if importlib.util.find_spec(name) is not None:
            importlib.import_module(name)
    return dict(sorted(_RULES.items(), key=lambda item: _natural_key(item[0])))


def collect(ctx: Context, rules: Optional[Dict[str, RuleFn]] = None) -> List[Violation]:
    """All violations of `rules` (default: every rule), one per identity."""
    seen: Dict[Tuple[str, str, str, str], Violation] = {}
    for rule_id, fn in (rules if rules is not None else all_rules()).items():
        for violation in fn(ctx):
            if violation.rule != rule_id or violation.surface not in SURFACES:
                raise ValueError(f"rule {rule_id} produced {violation!r}")
            seen.setdefault(violation.identity, violation)
    return sorted(seen.values(), key=lambda v: v.identity)


# ---------------------------------------------------------------------------
# Baseline.
# ---------------------------------------------------------------------------


def _clean(text: str) -> str:
    return " ".join(text.replace("\t", " ").split())


def read_baseline(baseline_dir: Path) -> Tuple[Dict[Tuple[str, str, str, str], str], List[str]]:
    """Baseline identities -> file name, plus parse problems."""
    lines: Dict[Tuple[str, str, str, str], str] = {}
    problems: List[str] = []
    for surface in SURFACES:
        path = baseline_dir / f"{surface}.txt"
        if not path.is_file():
            continue
        for number, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            parts = raw.split("\t")
            if len(parts) != 4 or not all(parts[:3]):
                problems.append(
                    f"api-conventions: {path.name}:{number} is not "
                    f"`<rule-id>\\t<op-key>\\t<discriminator>\\t<message>`: {raw!r}"
                )
                continue
            identity = (surface, parts[0], parts[1], parts[2])
            if identity in lines:
                problems.append(f"api-conventions: {path.name}:{number} repeats `{' '.join(parts[:3])}`")
            lines[identity] = path.name
    return lines, problems


def write_baseline(violations: Sequence[Violation], baseline_dir: Path) -> Dict[str, int]:
    """Writes all four baseline files; returns the line count per file."""
    baseline_dir.mkdir(parents=True, exist_ok=True)
    counts: Dict[str, int] = {}
    for surface in SURFACES:
        rows = sorted(
            "\t".join((v.rule, v.op_key, v.discriminator, _clean(v.message)))
            for v in violations
            if v.surface == surface
        )
        (baseline_dir / f"{surface}.txt").write_text("".join(row + "\n" for row in rows), encoding="utf-8")
        counts[f"{surface}.txt"] = len(rows)
    return counts


def compare(violations: Sequence[Violation], baseline_dir: Path) -> List[str]:
    """New violations and stale baseline lines, as problem lines."""
    baseline, problems = read_baseline(baseline_dir)
    current = {v.identity for v in violations}
    for violation in violations:
        if violation.identity not in baseline:
            site = f" ({violation.site})" if violation.site else ""
            problems.append(
                f"api-conventions: {violation.rule} [{violation.surface}] {violation.op_key} "
                f"{violation.discriminator}: {violation.message}{site}"
            )
    for identity, file_name in sorted(baseline.items()):
        if identity not in current:
            problems.append(
                f"api-conventions: stale line in {file_name}: `{' '.join(identity[1:])}` is no longer "
                "violated; delete the line"
            )
    return problems


def counts_report(violations: Sequence[Violation], baseline_dir: Path, rule_ids: Iterable[str]) -> List[str]:
    baseline, _ = read_baseline(baseline_dir)
    current = {v.identity for v in violations}
    out = []
    for rule_id in rule_ids:
        now = {i for i in current if i[1] == rule_id}
        base = {i for i in baseline if i[1] == rule_id}
        out.append(
            f"{rule_id}\tviolations={len(now)}\tbaseline={len(base)}\tnew={len(now - base)}\tstale={len(base - now)}"
        )
    per_file = {f"{s}.txt": 0 for s in SURFACES}
    for file_name in baseline.values():
        per_file[file_name] += 1
    out.append("baseline lines: " + ", ".join(f"{name}={count}" for name, count in per_file.items()))
    return out


def check_api_conventions(baseline_dir: Optional[Path] = None, ctx: Optional[Context] = None) -> List[str]:
    """Drift selector entry point: violations against the baseline."""
    ctx = ctx or Context.from_tree()
    violations = collect(ctx)
    return [*ctx.ledger.problems, *compare(violations, baseline_dir or BASELINE_DIR)]


def run_action(action: str, baseline_dir: Optional[Path] = None, ctx: Optional[Context] = None) -> int:
    """`list-rules`, `counts` or `write-baseline`; prints its output, returns an exit code."""
    rules = all_rules()
    if action == "list-rules":
        for rule_id in rules:
            print(rule_id)
        return 0
    ctx = ctx or Context.from_tree()
    violations = collect(ctx, rules)
    if ctx.ledger.problems:
        for problem in ctx.ledger.problems:
            print(problem)
        return 1
    baseline_dir = baseline_dir or BASELINE_DIR
    if action == "counts":
        for line in counts_report(violations, baseline_dir, rules):
            print(line)
        return 0
    if action == "write-baseline":
        written = write_baseline(violations, baseline_dir)
        print("api-conventions: wrote " + ", ".join(f"{name}={count}" for name, count in written.items()))
        return 0
    raise ValueError(action)


# ---------------------------------------------------------------------------
# Rules on the C header.
# ---------------------------------------------------------------------------


@rule("R1.form")
def _rule_name_form(ctx: Context) -> Iterable[Violation]:
    """A C function name is `fm_[<scope>_]<verb>[_<noun>...]` in lower snake case."""
    for fn in ctx.functions:
        for rule_id, message in check_c_name(fn.name):
            if rule_id == "R1.form":
                yield ctx.c_violation(rule_id, fn.name, NO_DISCRIMINATOR, message, fn.line)


@rule("R2.verb")
def _rule_verb(ctx: Context) -> Iterable[Violation]:
    """The word after the scope of a C function name is in the closed C verb vocabulary."""
    for fn in ctx.functions:
        for rule_id, message in check_c_name(fn.name):
            if rule_id == "R2.verb":
                yield ctx.c_violation(rule_id, fn.name, NO_DISCRIMINATOR, message, fn.line)


def _param_type_problem(param: CParam, enums: set) -> Optional[str]:
    role = param.role_name
    depth = param.pointer_depth
    want = None
    if role in ("sheet", "sheet_index") and param.base_type in _INTEGER_TYPES and depth <= 1:
        if role != "sheet_index" or param.base_type != "size_t":
            return f"`{param.ctype} {param.name}` is a sheet index; declare it `size_t{'*' * depth} sheet_index`"
        return None
    if role == "index" or role.endswith("_index") or role == "idx" or role.endswith("_idx"):
        want = "size_t"
    elif role in _ID_PARAMS:
        want = "uint32_t"
    elif role in _ROW_COL_PARAMS:
        want = "uint32_t"
    if want is not None and depth <= 1 and param.base_type != want:
        return f"`{param.ctype} {param.name}` should be `{want}{'*' * depth}`"
    if param.base_type in enums:
        return f"`{param.ctype} {param.name}` is enum-typed; enums cross the ABI as `int32_t{'*' * depth}`"
    if param.base_type in _BARE_C_SCALARS:
        return f"`{param.ctype} {param.name}` uses a bare C scalar; use a fixed-width type"
    return None


def _caller_buffer_problems(fn: CFunction) -> Iterable[Tuple[str, str]]:
    params = fn.params
    for i, param in enumerate(params[:-1]):
        if param.ctype not in ("char*", "uint8_t*") or params[i + 1].ctype != "size_t":
            continue
        expected = [("buf", param), ("cap", params[i + 1])]
        if i + 2 < len(params) and params[i + 2].ctype == "size_t*":
            expected.append(("out_len", params[i + 2]))
        for want, actual in expected:
            if actual.name != want:
                yield actual.name, f"caller buffer parameter `{actual.name}` should be named `{want}`"


@rule("R6.c")
def _rule_c_types(ctx: Context) -> Iterable[Violation]:
    """C scalar types: `size_t sheet_index`, `size_t` indices, `uint32_t` ids and rows / columns,
    enums as `int32_t`, no bare C scalars; a `*_count` getter returns `fm_status_t` through
    `size_t* out_count`; a `get` / `is` / `has` takes `const fm_workbook_t*`; a caller buffer is
    `buf, cap, out_len`.
    """
    for fn in ctx.functions:
        target = ctx.target_c(fn.name)
        seen = set()
        for param in fn.params:
            problem = _param_type_problem(param, ctx.enum_names)
            if problem and param.name not in seen:
                seen.add(param.name)
                yield ctx.c_violation("R6.c", fn.name, param.name, f"{fn.name}: {problem}", fn.line)
        for name, problem in _caller_buffer_problems(fn):
            if name not in seen:
                seen.add(name)
                yield ctx.c_violation("R6.c", fn.name, name, f"{fn.name}: {problem}", fn.line)
        if target.endswith("_count") and c_verb(target) != "set":
            last = fn.params[-1] if fn.params else None
            if fn.ret != "fm_status_t" or last is None or (last.ctype, last.name) != ("size_t*", "out_count"):
                yield ctx.c_violation(
                    "R6.c",
                    fn.name,
                    "count",
                    f"{fn.name}: a count returns `fm_status_t` with a trailing `size_t* out_count`",
                    fn.line,
                )
        if c_verb(target) in ("get", "is", "has") and fn.params and fn.params[0].ctype == "fm_workbook_t*":
            yield ctx.c_violation(
                "R6.c", fn.name, "const-handle", f"{fn.name}: a getter takes `const fm_workbook_t*`", fn.line
            )
    for typedef in ctx.typedefs:
        for member in typedef.fields:
            if member.base_type in ctx.enum_names:
                yield ctx.c_violation(
                    "R6.c",
                    typedef.name,
                    member.name,
                    f"{typedef.name}.{member.name} is enum-typed; enum fields are `int32_t`",
                    typedef.line,
                )


@rule("R7.c.byptr")
def _rule_record_by_pointer(ctx: Context) -> Iterable[Violation]:
    """A record argument is passed as `const T*`, never by value."""
    for fn in ctx.functions:
        for param in fn.params:
            if param.pointer_depth == 0 and param.base_type in ctx.struct_names:
                yield ctx.c_violation(
                    "R7.c.byptr",
                    fn.name,
                    param.name,
                    f"{fn.name}: `{param.ctype} {param.name}` passes a record by value; take `const {param.base_type}*`",
                    fn.line,
                )


@rule("R7.c.typedef")
def _rule_typedef_suffix(ctx: Context) -> Iterable[Violation]:
    """A struct, enum, handle or scalar typedef is named `fm_<name>_t`.

    Typedefs the ledger marks `replaced_by` are exempt: they are removed, not renamed.
    """
    for typedef in ctx.typedefs:
        if _C_TYPEDEF_NAME_RE.fullmatch(typedef.name):
            continue
        if "replaced_by" in ctx.ledger.entry(ctx.ledger.key_for("c", typedef.name)):
            continue
        yield ctx.c_violation(
            "R7.c.typedef",
            typedef.name,
            NO_DISCRIMINATOR,
            f"{typedef.kind} typedef {typedef.name} is not named `fm_<name>_t`",
            typedef.line,
        )


@rule("R8.c")
def _rule_c_returns(ctx: Context) -> Iterable[Violation]:
    """A C function returns `fm_status_t`; only `*_destroy` returns `void`, only the process string
    functions return `const char*`, and only the process number functions return `uint32_t`. Counts are
    judged by `R6.c`.
    """
    for fn in ctx.functions:
        target = ctx.target_c(fn.name)
        if fn.ret == "fm_status_t" or (target.endswith("_count") and c_verb(target) != "set"):
            continue
        if fn.ret == "void" and target.endswith("_destroy"):
            continue
        if fn.ret == "const char*" and target in PROCESS_STRING_FUNCTIONS:
            continue
        if fn.ret == "uint32_t" and target in PROCESS_NUMBER_FUNCTIONS:
            continue
        yield ctx.c_violation(
            "R8.c", fn.name, NO_DISCRIMINATOR, f"{fn.name} returns `{fn.ret}`, not `fm_status_t`", fn.line
        )


@rule("R10.c.doc")
def _rule_lifetime_doc(ctx: Context) -> Iterable[Violation]:
    """A function with a `const char**` output names its lifetime class in its doc:
    model-backed, read-scratch-backed, owned snapshot, static or thread-local view.
    """
    for fn in ctx.functions:
        if any(p.ctype == "const char**" for p in fn.params) and not _LIFETIME_RE.search(fn.doc):
            yield ctx.c_violation(
                "R10.c.doc",
                fn.name,
                NO_DISCRIMINATOR,
                f"{fn.name} returns a `const char**` view whose doc names none of: {', '.join(LIFETIME_CLASSES)}",
                fn.line,
            )


# Row properties whose column counterpart has a different name; a row's height,
# cached automatic height and cleared height all mirror the column width.
_ROW_TO_COLUMN_PROPERTY = {"height": "width", "auto_height": "width"}


def _counterpart_column(row_setter: str) -> Optional[str]:
    match = re.fullmatch(r"(.*_)(set|clear)_row_(\w+)", row_setter)
    if not match:
        return None
    prop = _ROW_TO_COLUMN_PROPERTY.get(match.group(3), match.group(3))
    return f"{match.group(1)}{match.group(2)}_column_{prop}"


def _takes_span(fn: CFunction) -> bool:
    names = [p.name for p in fn.params]
    return "first" in names and "last" in names


@rule("R11.axis")
def _rule_axis_symmetry(ctx: Context) -> Iterable[Violation]:
    """A row setter or clearer takes the same `(first, last)` span as its column counterpart."""
    by_target = {ctx.target_c(fn.name): fn for fn in ctx.functions}
    for target, fn in by_target.items():
        column = by_target.get(_counterpart_column(target) or "")
        if column is not None and _takes_span(column) and not _takes_span(fn):
            yield ctx.c_violation(
                "R11.axis",
                fn.name,
                NO_DISCRIMINATOR,
                f"{fn.name} takes a single row while {column.name} takes a `(first, last)` span",
                fn.line,
            )


def _rect_quads(members: Sequence[CParam]) -> Iterable[str]:
    """First coordinate name of every rectangle spelled out as four adjacent integers.

    Adjacency keeps two independent spans (repeat rows, repeat columns) from reading as one rectangle.
    """
    position = {m.role_name: i for i, m in enumerate(members) if m.base_type in _INTEGER_TYPES}
    found = set()
    for quad in _RECT_QUADS:
        for name in list(position):
            if not name.endswith(quad[0]):
                continue
            prefix = name[: len(name) - len(quad[0])]
            if prefix and not prefix.endswith("_"):
                continue
            names = [prefix + part for part in quad]
            if not all(n in position for n in names) or any(n in found for n in names):
                continue
            slots = sorted(position[n] for n in names)
            if slots[-1] - slots[0] == 3:
                found.update(names)
                yield members[position[names[0]]].name


@rule("R11.rect")
def _rule_single_rect(ctx: Context) -> Iterable[Violation]:
    """A cell rectangle is one `fm_cell_rect_t {first_row, first_col, last_row, last_col}`,
    never four loose coordinates in a struct or a parameter list.
    """
    for typedef in ctx.typedefs:
        if typedef.kind != "struct":
            continue
        if typedef.name == CELL_RECT:
            fields = tuple((m.ctype, m.name) for m in typedef.fields)
            if fields != tuple(("uint32_t", f) for f in CELL_RECT_FIELDS):
                yield ctx.c_violation(
                    "R11.rect",
                    typedef.name,
                    NO_DISCRIMINATOR,
                    f"{CELL_RECT} must be exactly `uint32_t {', '.join(CELL_RECT_FIELDS)}`",
                    typedef.line,
                )
            continue
        for first in _rect_quads(typedef.fields):
            yield ctx.c_violation(
                "R11.rect",
                typedef.name,
                first,
                f"{typedef.name} spells a rectangle from `{first}`; use `{CELL_RECT}`",
                typedef.line,
            )
    for fn in ctx.functions:
        for first in _rect_quads(fn.params):
            yield ctx.c_violation(
                "R11.rect",
                fn.name,
                first,
                f"{fn.name} takes a rectangle as four coordinates from `{first}`; take `const {CELL_RECT}*`",
                fn.line,
            )


# Comment markers that describe a struct field's value set: an enum typedef or a `FM_<FAMILY>_*` macro family.
_FIELD_LINE_RE = re.compile(r"^\s*([A-Za-z_][\w ]*?[\w*])\s*\b([A-Za-z_]\w*)\s*(?:\[[^\]]*\])?\s*;(.*)$")
_ENUM_TYPEDEF_REF_RE = re.compile(r"\bfm_[a-z0-9_]+_t\b")
_MACRO_FAMILY_RE = re.compile(r"\b(FM_[A-Z0-9_]+?)_?\*")


def _struct_field_comments(header: str) -> Dict[str, List[Tuple[str, str, str]]]:
    """Struct name -> `(ctype, field, trailing comment)` for every field declared on one line."""
    out: Dict[str, List[Tuple[str, str, str]]] = {}
    for match in _C_STRUCT_RE.finditer(header):
        fields = []
        for line in match.group(1).splitlines():
            code, _, comment = line.partition("/*")
            found = _FIELD_LINE_RE.match(code + ";" if ";" not in code else code)
            if found is None or found.group(1).strip() in ("union", "struct", "}"):
                continue
            ctype = " ".join(found.group(1).replace("*", " * ").split()).replace(" *", "*")
            fields.append((ctype, found.group(2), comment.split("*/")[0] if comment else found.group(3)))
        out[match.group(2)] = fields
    return out


def _enum_family(comment: str, enums: set, macros: set) -> Optional[str]:
    """The value set a field comment names: an enum typedef, or a `FM_<FAMILY>_*` family with a `#define`."""
    for name in _ENUM_TYPEDEF_REF_RE.findall(comment):
        if name in enums:
            return name
    for prefix in _MACRO_FAMILY_RE.findall(comment):
        if any(m.startswith(prefix.rstrip("_") + "_") for m in macros):
            return prefix.rstrip("_") + "_*"
    return None


_HEADER_DEFINE_RE = re.compile(r"^#define\s+(FM_[A-Z0-9_]+)", re.M)


@rule("R6.c.field")
def _rule_c_field_types(ctx: Context) -> Iterable[Violation]:
    """A public struct field is fixed-width: no `size_t` (an index or count is `uint32_t`, a byte length
    `uint64_t`), and a field whose comment names its value set (an enum typedef or a `FM_<FAMILY>_*` macro
    family) is `int32_t`, or a `uint8_t` whose comment names an `fm_*_t` enum. A field declared with the enum
    typedef itself is judged by `R6.c`.
    """
    macros = set(_HEADER_DEFINE_RE.findall(ctx.header))
    lines = {t.name: t.line for t in ctx.typedefs}
    for struct, fields in _struct_field_comments(ctx.header).items():
        for ctype, name, comment in fields:
            line = lines.get(struct, 0)
            if ctype == "size_t":
                yield ctx.c_violation(
                    "R6.c.field",
                    struct,
                    name,
                    f"{struct}.{name} is `size_t`; an index or count field is `uint32_t`, a byte length `uint64_t`",
                    line,
                )
                continue
            family = _enum_family(comment, ctx.enum_names, macros)
            if family is None or ctype in ctx.enum_names or ctype == "int32_t":
                continue
            if ctype == "uint8_t" and family in ctx.enum_names:
                continue
            yield ctx.c_violation(
                "R6.c.field",
                struct,
                name,
                f"{struct}.{name} carries `{family}` as `{ctype}`; an enum field is `int32_t`",
                line,
            )


# Work-process phrases a header comment must not carry; `later` alone has legitimate uses.
_PROCESS_PHRASES = (
    ("a later bundle", re.compile(r"\ba later bundle\b", re.I)),
    ("for now", re.compile(r"\bfor now\b", re.I)),
    ("TODO", re.compile(r"\bTODO\b")),
    ("FIXME", re.compile(r"\bFIXME\b")),
    ("temporar", re.compile(r"\btemporar", re.I)),
)
_HEADER_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_HEADER_REF_RE = re.compile(r"(?<![\w])(fm_[a-z0-9_]*[a-z0-9_])(\*?)")
_STRUCT_TAG_RE = re.compile(r"\b(?:struct|enum|union)\s+(fm_[a-z0-9_]+)")


@rule("R15.header")
def _rule_header_references(ctx: Context) -> Iterable[Violation]:
    """Every `fm_*` token in a header comment names a declared function, type or struct tag; a wildcard
    `fm_x_*` matches at least one. No header comment carries a work-process phrase (`a later bundle`,
    `for now`, `TODO`, `FIXME`, `temporar...`).
    """
    known = {fn.name for fn in ctx.functions} | {t.name for t in ctx.typedefs}
    known |= set(_STRUCT_TAG_RE.findall(ctx.header))
    declarations = sorted((fn.line, fn.name) for fn in ctx.functions)
    declarations += sorted((t.line, t.name) for t in ctx.typedefs)
    declarations.sort()
    for match in _HEADER_COMMENT_RE.finditer(ctx.header):
        text = match.group(0)
        line = _line_of(ctx.header, match.start())
        for token, star in _HEADER_REF_RE.findall(text):
            wildcard = star and token.endswith("_")
            resolved = any(n.startswith(token) for n in known) if wildcard else token in known
            if not resolved:
                shown = token + ("*" if wildcard else "")
                yield ctx.c_violation(
                    "R15.header", shown, "unresolved", f"a header comment names `{shown}`, which is not declared", line
                )
        for phrase, pattern in _PROCESS_PHRASES:
            if pattern.search(text):
                after = next((name for at, name in declarations if at >= line), "header")
                yield ctx.c_violation(
                    "R15.header", after, phrase, f"a header comment says `{phrase}`; describe the current state", line
                )
