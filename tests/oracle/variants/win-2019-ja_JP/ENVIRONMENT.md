# Oracle Environment — win-2019-ja_JP

This variant targets a Windows host running **Microsoft Office
Professional Plus 2019 Retail** in Japanese locale. The host was
previously labelled `win-365-ja_JP` (the perpetual / subscription
distinction was missed because the Click-to-Run engine version reads
indistinguishably from Microsoft 365). All goldens under this directory
were generated against the Office 2019 function set — modern dynamic-
array additions (XLOOKUP, ARRAYTOTEXT, VALUETOTEXT, FILTER, SORT,
UNIQUE, SEQUENCE, LET, LAMBDA, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, ...)
resolve to `#NAME?` here because the binary set does not ship them.

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

## Generation

- **generated_at**: `2026-05-14T20:37:32Z`

Re-generate via the standard WSL bridge invocation; the driver writes
this file automatically from the live `Application` properties:

```
make oracle-gen TARGET=win-2019-ja_JP SUITE=<name>
```

For the canonical Microsoft 365 (Windows, ja-JP) target, use
`win-365-ja_JP` — which is currently `wanted` in `tools/oracle/
targets.yaml`, pending a real M365 install on the host.
