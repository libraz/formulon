# Reference capture — win-2019-ja_JP (NOT a maintained oracle target)

**This is historical reference data, not a tested oracle target.** It
lives under `tests/oracle/reference/` precisely so the variant test
harness (which scans only `tests/oracle/variants/*`) never picks it up.
It will not be regenerated.

The goldens here were captured by mistake: the host was believed to be
running Microsoft 365 and was labelled `win-365-ja_JP`, but it was
actually running **Microsoft Office Professional Plus 2019 Retail**
(ja-JP). The perpetual / subscription distinction was missed because the
Click-to-Run engine version reads indistinguishably from Microsoft 365.
All goldens were therefore generated against the Office 2019 function
set — modern dynamic-array additions (XLOOKUP, ARRAYTOTEXT, VALUETOTEXT,
FILTER, SORT, UNIQUE, SEQUENCE, LET, LAMBDA, TEXTBEFORE, TEXTAFTER,
TEXTSPLIT, ...) resolve to `#NAME?` here because the binary set does not
ship them. They are kept only as a record of what that environment
reported.

## Reported values

- **Application.Version**: `16.0` (shared by every Office 2016+, not a
  product discriminator on its own)
- **Application.Build**: `19929.0` (Click-to-Run engine build; advances
  independently of the licensed product family — do NOT use as a proxy
  for Microsoft 365)
- **CalculationVersion**: `191029`
- **Excel locale**: `ja-JP`
- **OperatingSystem**: `Windows (32-bit) NT 10.00`
- **date1904**: `False`
- **iterative**: `False`

## Licensed product (the real discriminator)

From `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`:

```
ProductReleaseIds: ProjectPro2019Retail, ProPlus2019Retail, VisioPro2019Retail
Platform:          x86
CDNBaseUrl:        .../492350f6-3a01-4f97-b9c0-c7c6ddf67d60   # Office 2019 Retail channel
```

`ProductReleaseIds` is the field that determines the available function
set, regardless of how recent the Click-to-Run engine `Build` is.

## Capture date

- **generated_at**: `2026-05-14T20:37:32Z`

This data is frozen. There is no `win-2019-ja_JP` entry in
`tools/oracle/targets.yaml` and no regeneration path. For the canonical
Microsoft 365 (Windows, ja-JP) target use `win-365-ja_JP`.
