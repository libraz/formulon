# Security policy

## Reporting a vulnerability

Report privately by email to `libraz@libraz.net`, not through the public issue tracker.

Include the formula or workbook that triggers it, what happened, the affected version and binding (npm/WASM, Python, native addon or CLI), and a minimal reproduction if you have one. Expect an acknowledgement within a few days.

## Supported versions

Fixes land on the latest release only; older versions are not patched.

## What is in scope

Formulon evaluates formulas it did not write — that is the whole job — so the
parser and the evaluator are the interesting surface. In scope:

- A formula or workbook that causes a crash, a hang, unbounded recursion or
  unbounded memory growth, in the C++ core or through the WebAssembly, Python
  or CLI packaging.
- Anything reached through `FILTERXML`, which parses a string the caller did not
  write: entity expansion, external entity resolution, or resource exhaustion.
- Any path by which evaluating a formula touches the filesystem or the network.
  Evaluation is meant to be pure; a reachable side effect is a vulnerability
  even if it looks benign.
- Escapes from the WebAssembly sandbox, or a formula that reads memory belonging
  to another workbook in the same instance.

## What is not in scope

- Divergence from Excel. A cell that computes a different value than Excel does
  is a correctness bug, tracked against the oracle data; report it as a normal
  issue.
- Documented ceilings behaving as documented — iteration, depth and cell-count
  limits exist so a hostile workbook cannot exhaust the host. A ceiling that can
  be bypassed is in scope.
- Spreadsheet UI behaviour. Report it to [formulon-cell](https://github.com/libraz/formulon-cell), which has its own security policy.
- Findings that require an attacker to already control the process embedding the
  engine.
