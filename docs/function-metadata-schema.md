# Function metadata provider schema

Formulon's calculation engine deliberately ships **no** human-readable
function documentation. `functionMetadata(name, locale)` returns only the
data the engine actually owns — canonical name, arity bounds, and an
availability class — and leaves `signatureTemplate` / `description` as
`NULL` (empty). Non-primary-locale display-name tables are likewise empty.

That content is a UI concern, not a calculation concern. Rather than bake
several hundred localized strings into the WASM binary, the bindings expose
a **provider seam**: a host (an editor, an MCP server, a docs site) supplies
its own metadata document and merges it over the engine's structural result
at display time. This file is the canonical contract for that document.

## Scope

This metadata is **display-only**. It never affects how a formula is
parsed or evaluated.

- Formula **input parsing is fixed to English canonical function names.**
  A formula written with a localized function name (formula-language
  localization) is **not** interpreted by this seam — that is a separate
  engine-level concern and is out of scope here.
- The C ABI entry points `fm_function_localize` / `fm_function_canonicalize`
  are **not** driven by this provider. They remain canonical-fallback
  (they return the canonical name unchanged for non-primary locales).
- The primary locale is `ja-JP`, whose function names are identical to the
  English canonical names, so the primary locale needs no `aliases` entry
  at all (the alias is the identity mapping).

## Document shape

```json
{
  "version": 1,
  "functions": {
    "XLOOKUP": {
      "signature": "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
      "description": "Searches a range or an array, and returns an item corresponding to the first match it finds.",
      "aliases": { "fr-FR": "RECHERCHEX" },
      "localized": {
        "fr-FR": {
          "signature": "RECHERCHEX(valeur_cherchée, tableau_recherche, tableau_retour, [si_non_trouvé], [mode_correspondance], [mode_recherche])",
          "description": "Recherche dans une plage ou un tableau et renvoie l'élément correspondant à la première correspondance trouvée."
        }
      }
    }
  }
}
```

### Fields

| Field | Type | Meaning |
|-------|------|---------|
| `version` | integer | Schema version. Currently `1`. |
| `functions` | object | Map keyed by **canonical UPPERCASE** function name. |
| `functions[NAME].signature` | string? | Default (locale-agnostic) signature template. |
| `functions[NAME].description` | string? | Default (locale-agnostic) description. |
| `functions[NAME].aliases` | object? | Map of locale tag → localized display name. |
| `functions[NAME].localized` | object? | Map of locale tag → `{ signature?, description? }` overrides. |

- Keys of `functions` are **canonical, uppercase, English** function names
  (`XLOOKUP`, `SUM`, `VLOOKUP`), matching the engine's canonical name.
- Keys inside `aliases` and `localized` are BCP-47 locale tags
  (`fr-FR`, `de-DE`, ...). They are the **display locale** the host chose,
  independent of the numeric locale code passed to the engine's
  `functionMetadata` call.

## Merge semantics

A host resolves the displayed metadata for a function by calling the
engine, then merging a provider entry over it with the pure helper
`mergeFunctionMetadata` (the native Node package) or
`merge_function_metadata` (Python). The raw WASM C ABI intentionally does
not export a merge helper; browser hosts can apply the same precedence below
in their provider layer. Given `base` (the engine result), `entry`
(`functions[NAME]`, or absent), and a display `locale` tag, each field is
resolved by first-non-null precedence:

- **signature** — `entry.localized[locale].signature` → `entry.signature`
  → engine value (`NULL`).
- **description** — `entry.localized[locale].description` →
  `entry.description` → engine value (`NULL`).
- **display name** — `entry.aliases[locale]` → the canonical name.

When no provider entry exists for a function, the merge is the identity:
`base` is returned unchanged, so `signature` / `description` stay `NULL`
exactly as the engine reports them. This preserves the engine's behaviour
for hosts that inject nothing.

## Reference example

`docs/examples/function-metadata.example.json` is a small, schema-conformant
document (`XLOOKUP`, `SUM`, `VLOOKUP`) suitable as a starting point.
