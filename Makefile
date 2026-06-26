# Formulon convenience targets.
#
# Long-running commands (lint, test, full builds) should redirect output to a
# file and grep the file — never pipe directly into grep.

BUILD_DIR ?= build
GENERATOR ?=
CMAKE ?= cmake
CTEST ?= ctest
CLANG_FORMAT ?= clang-format
CLANG_TIDY ?= clang-tidy

# Node.js binary. Overridable so CI can dodge emsdk's bundled non-executable
# `node` shim (see release.yml's `make CMAKE=... NODE=...` pattern).
NODE ?= node

# Default test parallelism. Override with `make test CTEST_JOBS=N`. CTest
# infers a reasonable number when the variable is empty (`-j 0` lets it
# pick), but defaulting to the host CPU count keeps the fast loop near
# 2 minutes on a modern workstation rather than the 10-15 minute serial
# walk through the ~10k oracle parametric tests.
CTEST_JOBS ?= 0

SRC_DIRS := src tests
CPP_GLOB := $(shell find $(SRC_DIRS) -type f \( -name '*.cpp' -o -name '*.h' \) 2>/dev/null)

.PHONY: all build release test test-slow test-all format format-check lint clean \
        wasm wasm-debug wasm-capi test-wasm test-python size-check \
        npm-package npm-test npm-pack npm-check-dts \
        node-native node-package node-test \
        python-package python-test python-wheel \
        parity-test \
        oracle-setup oracle-setup-mac oracle-setup-wsl \
        oracle-gen oracle-gen-cf oracle-gen-workbook \
        oracle-verify oracle-contribute oracle-contribute-list \
        ironcalc-import ironcalc-verify \
        fuzz-parser fuzz-xlsx fuzz-eval bench coverage mutation \
        function-status behavior-status

all: build

build:
	$(CMAKE) -B $(BUILD_DIR) -DCMAKE_BUILD_TYPE=Debug
	$(CMAKE) --build $(BUILD_DIR) --parallel

release:
	$(CMAKE) -B build-release -DCMAKE_BUILD_TYPE=Release
	$(CMAKE) --build build-release --parallel

test:
	(cd $(BUILD_DIR) && $(CTEST) -LE "SLOW|LOAD" -j $(CTEST_JOBS) --output-on-failure --timeout 30)

test-slow:
	(cd $(BUILD_DIR) && $(CTEST) -LE "LOAD" -j $(CTEST_JOBS) --output-on-failure --timeout 120)

test-all:
	(cd $(BUILD_DIR) && $(CTEST) -j $(CTEST_JOBS) --output-on-failure --timeout 300)

format:
	@find $(SRC_DIRS) -type f \( -name '*.cpp' -o -name '*.h' \) -print0 \
	  | xargs -0 -r $(CLANG_FORMAT) -i

format-check:
	@find $(SRC_DIRS) -type f \( -name '*.cpp' -o -name '*.h' \) -print0 \
	  | xargs -0 -r $(CLANG_FORMAT) --dry-run --Werror

lint:
	@if [ ! -f $(BUILD_DIR)/compile_commands.json ]; then \
	  echo "compile_commands.json missing; run 'make build' first."; exit 1; \
	fi
	@find src -type f \( -name '*.cpp' -o -name '*.h' \) \
	  -exec $(CLANG_TIDY) -p $(BUILD_DIR) {} +

clean:
	rm -rf $(BUILD_DIR) build-release build-relwithdebinfo \
	       build-asan build-ubsan build-tsan build-coverage \
	       $(WASM_BUILD_DIR) $(WASM_DEBUG_BUILD_DIR) $(WASM_CAPI_BUILD_DIR)

# -- WASM build / smoke-test targets --------------------------------------
# `make wasm`        -> Release-mode formulon.{js,wasm} under build-wasm/.
# `make wasm-debug`  -> Debug-mode artifact (with assertions) under
#                       build-wasm-debug/.
# `make test-wasm`   -> Node-based smoke tests against build-wasm/formulon.js.
#
# All three short-circuit cleanly when the host lacks emscripten so the
# native CI path stays green without an Emscripten install.
EM_CMAKE := emcmake cmake
WASM_BUILD_DIR ?= build-wasm
WASM_DEBUG_BUILD_DIR ?= build-wasm-debug
WASM_CAPI_BUILD_DIR ?= build-wasm-capi

wasm:
	@if ! command -v emcmake >/dev/null 2>&1; then \
	  echo "wasm: emscripten toolchain not found in PATH"; \
	  echo "  Install: https://emscripten.org/docs/getting_started/downloads.html"; \
	  exit 1; \
	fi
	$(EM_CMAKE) -B $(WASM_BUILD_DIR) -DCMAKE_BUILD_TYPE=Release \
	  -DFM_BUILD_WASM=ON -DFM_BUILD_TESTING=OFF -DFM_BUILD_CLI=OFF
	$(CMAKE) --build $(WASM_BUILD_DIR) --parallel --target formulon_wasm
	@echo ""
	@echo "wasm artifacts:"
	@ls -la $(WASM_BUILD_DIR)/formulon.wasm $(WASM_BUILD_DIR)/formulon.js 2>/dev/null || \
	  echo "  (artifacts not found; check build log above)"

wasm-debug:
	@if ! command -v emcmake >/dev/null 2>&1; then \
	  echo "wasm-debug: emscripten toolchain not found in PATH"; \
	  exit 1; \
	fi
	$(EM_CMAKE) -B $(WASM_DEBUG_BUILD_DIR) -DCMAKE_BUILD_TYPE=Debug \
	  -DFM_BUILD_WASM=ON -DFM_BUILD_TESTING=OFF -DFM_BUILD_CLI=OFF
	$(CMAKE) --build $(WASM_DEBUG_BUILD_DIR) --parallel --target formulon_wasm

# capi WASM: standalone reactor build consumed by wasmtime-py. No JS
# glue, no pthread, exports the curated fm_* list from
# tools/wasm/capi_exports.txt. Drives the PyPI distribution.
wasm-capi:
	@if ! command -v emcmake >/dev/null 2>&1; then \
	  echo "wasm-capi: emscripten toolchain not found in PATH"; \
	  exit 1; \
	fi
	$(EM_CMAKE) -B $(WASM_CAPI_BUILD_DIR) -DCMAKE_BUILD_TYPE=Release \
	  -DFM_BUILD_WASM=ON -DFM_WASM_VARIANT=capi \
	  -DFM_BUILD_TESTING=OFF -DFM_BUILD_CLI=OFF
	$(CMAKE) --build $(WASM_CAPI_BUILD_DIR) --parallel --target formulon_wasm
	@echo ""
	@echo "wasm-capi artifact:"
	@ls -la $(WASM_CAPI_BUILD_DIR)/formulon_capi.wasm 2>/dev/null || \
	  echo "  (artifact not found; check build log above)"

test-wasm:
	@if [ ! -f $(WASM_BUILD_DIR)/formulon.js ]; then \
	  echo "test-wasm: $(WASM_BUILD_DIR)/formulon.js missing; run 'make wasm' first"; \
	  exit 1; \
	fi
	@if ! command -v $(NODE) >/dev/null 2>&1; then \
	  echo "test-wasm: '$(NODE)' not found"; \
	  exit 1; \
	fi
	$(NODE) tests/wasm/run.mjs

# -- npm packaging targets ------------------------------------------------
# `make npm-package` -> stage build-wasm/formulon.{js,wasm} + the
#                       hand-written .d.ts into packages/npm/dist/.
# `make npm-test`    -> run node:test smoke tests against the staged
#                       package (catches staging mistakes that running
#                       against build-wasm/ would mask).
# `make npm-pack`    -> produce a publishable .tgz under build-wasm/.
#
# None of these run automatically as part of `make build` / `make test`.
# Publishing is a manual, out-of-band step.
NPM_PKG_DIR := packages/npm

npm-package: wasm
	@if ! command -v $(NODE) >/dev/null 2>&1; then \
	  echo "npm-package: '$(NODE)' not found"; \
	  exit 1; \
	fi
	$(NODE) $(NPM_PKG_DIR)/scripts/stage.mjs \
	  --build-dir $(WASM_BUILD_DIR) \
	  --out-dir $(NPM_PKG_DIR)/dist

# `make npm-check-dts` -> fail if the staged dist/formulon.d.ts has
#                         drifted from the canonical src/wasm/formulon.d.ts.
# Needs no WASM build, so it is cheap to run in CI and locally.
npm-check-dts:
	@if ! command -v $(NODE) >/dev/null 2>&1; then \
	  echo "npm-check-dts: '$(NODE)' not found"; \
	  exit 1; \
	fi
	$(NODE) $(NPM_PKG_DIR)/scripts/check-dts.mjs

npm-test: npm-package npm-check-dts
	@if ! command -v $(NODE) >/dev/null 2>&1; then \
	  echo "npm-test: '$(NODE)' not found"; \
	  exit 1; \
	fi
	(cd $(NPM_PKG_DIR) && $(NODE) --test 'test/*.test.mjs')

npm-pack: npm-package
	@if ! command -v npm >/dev/null 2>&1; then \
	  echo "npm-pack: npm not found in PATH"; \
	  exit 1; \
	fi
	(cd $(NPM_PKG_DIR) && npm pack --pack-destination ../../$(WASM_BUILD_DIR)/)
	@echo ""
	@echo "npm tarball:"
	@ls -la $(WASM_BUILD_DIR)/libraz-formulon-*.tgz 2>/dev/null | tail -1 || \
	  echo "  (tarball not found; check log above)"

# -- Python packaging targets --------------------------------------------
# `make python-package` -> stage build-wasm-capi/formulon_capi.wasm into
#                          packages/python/formulon/_wasm/.
# `make python-test`    -> run the unittest smoke suite against the
#                          staged source tree (catches stage / packaging
#                          errors).
# `make python-wheel`   -> produce build-py/dist/formulon-*-py3-none-any.whl.
#
# The Python distribution is pure-Python with a bundled WebAssembly
# module. The WASM is built by `make wasm-capi` (Emscripten toolchain
# required); the wheel itself is py3-none-any and works on every OS
# and arch wasmtime supports.
PY_PKG_DIR := packages/python
PY_BUILD_DIR ?= build-py
PYTHON ?= $(shell python3 -c 'import sys; print(sys.executable)')

python-package: wasm-capi
	@if ! command -v $(PYTHON) >/dev/null 2>&1; then \
	  echo "python-package: python3 not found in PATH"; \
	  exit 1; \
	fi
	$(PYTHON) $(PY_PKG_DIR)/scripts/stage.py \
	  --build-dir $(WASM_CAPI_BUILD_DIR)

python-test: python-package
	@if ! $(PYTHON) -c "import wasmtime" >/dev/null 2>&1; then \
	  echo "python-test: wasmtime runtime missing; install with:"; \
	  echo "  $(PYTHON) -m pip install wasmtime"; \
	  exit 1; \
	fi
	@(cd $(PY_PKG_DIR) && $(PYTHON) -m unittest discover -v tests)

python-wheel: python-package
	@if ! $(PYTHON) -m pip --version >/dev/null 2>&1; then \
	  echo "python-wheel: pip not available for python3"; \
	  exit 1; \
	fi
	@if ! $(PYTHON) -c "import setuptools" >/dev/null 2>&1; then \
	  echo "python-wheel: Python package 'setuptools' not installed"; \
	  exit 1; \
	fi
	@if ! $(PYTHON) -c "import wheel" >/dev/null 2>&1; then \
	  echo "python-wheel: Python package 'wheel' not installed"; \
	  exit 1; \
	fi
	mkdir -p $(PY_BUILD_DIR)/dist
	(cd $(PY_PKG_DIR) && $(PYTHON) setup.py bdist_wheel \
	    --dist-dir ../../$(PY_BUILD_DIR)/dist/)
	@echo ""
	@echo "python wheel:"
	@ls -la $(PY_BUILD_DIR)/dist/formulon-*.whl 2>/dev/null | tail -1 || \
	  echo "  (wheel not found; check log above)"

# -- Node native (N-API) packaging targets ----------------------------
# `make node-native`  -> build formulon.node via FM_BUILD_NODE_ADDON=ON.
# `make node-package` -> stage build/bin/formulon.node + JS shim into
#                        packages/npm-native/dist/.
# `make node-test`    -> node:test smoke against the staged package.
NODE_NATIVE_PKG_DIR := packages/npm-native
NODE_NATIVE_BUILD_DIR ?= build

node-native:
	$(CMAKE) -B $(NODE_NATIVE_BUILD_DIR) -DCMAKE_BUILD_TYPE=Release -DFM_BUILD_NODE_ADDON=ON
	$(CMAKE) --build $(NODE_NATIVE_BUILD_DIR) --target formulon_node --parallel

node-package: node-native
	@if ! command -v $(NODE) >/dev/null 2>&1; then \
	  echo "node-package: '$(NODE)' not found"; exit 1; \
	fi
	$(NODE) $(NODE_NATIVE_PKG_DIR)/scripts/stage.mjs \
	  --build-dir $(NODE_NATIVE_BUILD_DIR) \
	  --out-dir $(NODE_NATIVE_PKG_DIR)/dist

node-test: node-package
	(cd $(NODE_NATIVE_PKG_DIR) && $(NODE) --test 'test/*.test.mjs')

# -- Cross-channel parity gate ------------------------------------------------
# `make parity-test` -> evaluate fixtures.json on every available channel
#                       (CLI / npm / Python wheel) and assert all channels
#                       agree at %.15g-canonicalized IEEE-754 bit
#                       granularity. Skip-aware: missing channels are
#                       reported but do not fail. Fails (exit 1) only when
#                       two or more channels actually disagreed.
#
# Intentionally has no make-level dependency on `npm-package`, `python-package`,
# or the native build: the runner itself decides which channels to exercise
# based on what is on disk. Driving the prerequisite builds is a CI-side
# concern.
parity-test:
	@command -v python3 >/dev/null 2>&1 || { \
	  echo "parity-test: python3 not found"; exit 1; }
	python3 tests/parity/run_parity.py

# Alias kept for backward compatibility with `make test-python`.
test-python: python-test

# Standalone .wasm size report. Reads the artifact built by `make wasm`
# and gates against the milestone size ceiling. See
# tools/bench/wasm_size_report.sh and CLAUDE.md "WASM Size Policy".
size-check:
	@sh tools/bench/wasm_size_report.sh $(WASM_BUILD_DIR)/formulon.wasm

# -- Oracle targets --------------------------------------------------------
# Drive Mac Excel 365 to generate golden JSON (oracle-gen), and verify
# Formulon's output against committed goldens (oracle-verify). The
# generator is macOS-only; the verifier runs on any platform that can
# build the test binary because it only reads committed JSON.
#
# Python tooling is managed by rye (https://rye.astral.sh/):
#   tools/oracle/pyproject.toml      project + deps
#   tools/oracle/requirements.lock   pinned versions (committed)
#   tools/oracle/.venv/              local venv (gitignored)
#
# See tools/oracle/README.md.
ORACLE_DIR := tools/oracle
ORACLE_VENV := $(ORACLE_DIR)/.venv
ORACLE_PY := $(ORACLE_VENV)/bin/python
ORACLE_GEN := $(ORACLE_PY) tools/oracle/oracle_gen.py

# All `oracle-gen*` targets below run an M365 sentinel at startup and
# abort with a clear error on Office 2019 / pre-M365 Excel. The check is
# implemented by OracleDriver.assert_m365_or_abort (formula + workbook
# tracks) and a parallel open-coded probe in cf_oracle_gen (CF track).

# Shared venv guard for the oracle-gen* targets. Each target prefixes
# this with `@$(call require_oracle_venv,<target-name>)` so the error
# message names the actual target the user invoked.
define require_oracle_venv
if [ ! -x "$(ORACLE_PY)" ]; then \
  echo "$(1): run 'make oracle-setup' first"; exit 1; \
fi
endef

oracle-setup:
	@uname_s=$$(uname -s); \
	if [ "$$uname_s" = "Darwin" ]; then \
	  $(MAKE) oracle-setup-mac; \
	elif [ "$$uname_s" = "Linux" ] && grep -qi microsoft /proc/version 2>/dev/null; then \
	  $(MAKE) oracle-setup-wsl; \
	else \
	  echo "oracle-setup: unsupported host $$uname_s"; \
	  echo "  supported: macOS (Darwin) for primary, WSL2 (Linux+Microsoft) for windows variant"; \
	  exit 1; \
	fi

oracle-setup-mac:
	@if ! command -v rye >/dev/null 2>&1; then \
	  echo "oracle-setup-mac: rye not found in PATH."; \
	  echo "  Install from https://rye.astral.sh/ (curl -sSf https://rye.astral.sh/get | bash)"; \
	  echo "  or: brew install rye"; \
	  exit 1; \
	fi
	@(cd $(ORACLE_DIR) && rye sync)
	@echo "oracle-setup-mac: venv ready at $(ORACLE_VENV)"
	@echo "Grant Automation permission: System Settings -> Privacy & Security"
	@echo "    -> Automation -> (your terminal) -> Microsoft Excel"
	@$(ORACLE_VENV)/bin/python tools/oracle/cli.py setup --target mac-365-ja_JP || true

oracle-setup-wsl:
	@if ! command -v rye >/dev/null 2>&1; then \
	  echo "oracle-setup-wsl: rye not found. Install: curl -sSf https://rye.astral.sh/get | bash"; \
	  exit 1; \
	fi
	@(cd $(ORACLE_DIR) && rye sync)
	@echo "oracle-setup-wsl: venv ready at $(ORACLE_VENV)"
	@echo ""
	@echo "Next steps for the Windows side (run in PowerShell, one-time):"
	@echo "  winget install Python.Python.3.12"
	@echo "  py -m pip install xlwings pywin32 pyyaml"
	@echo ""
	@echo "Then point the bridge at the Windows-side python.exe (preferred):"
	@echo "  export FORMULON_WIN_PYTHON=\"/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe\""
	@echo "  # the cli.py setup step below auto-discovers candidates and prints"
	@echo "  # the exact export line; copy that into your shell rc."
	@echo "  # (Editing tools/oracle/targets.yaml is only for private forks — the"
	@echo "  # env var keeps per-machine paths out of committed files.)"
	@$(ORACLE_VENV)/bin/python tools/oracle/cli.py setup --target win-365-ja_JP || true

oracle-gen:
	@$(call require_oracle_venv,oracle-gen)
	@$(ORACLE_GEN) $(if $(SUITE),--suite $(SUITE),) $(if $(TARGET),--target $(TARGET),)

# CF (conditional-formatting) track. macOS-only; see cf_oracle_gen.py for
# the supported rule subset.
oracle-gen-cf:
	@$(call require_oracle_venv,oracle-gen-cf)
	@$(ORACLE_PY) tools/oracle/cf_oracle_gen.py $(if $(SUITE),--suite $(SUITE),)

# Workbook track (pivot tables + print areas). Target is auto-detected
# from the host OS; pass TARGET=<name> to override.
oracle-gen-workbook:
	@$(call require_oracle_venv,oracle-gen-workbook)
	@$(ORACLE_PY) tools/oracle/cli.py workbook \
	  $(if $(TARGET),--target $(TARGET),) $(if $(SUITE),--suite $(SUITE),)

# oracle-verify shells to `ctest -L oracle`, which already selects the
# workbook oracle binary: formulon_workbook_oracle_tests carries the
# "oracle" CTest label, so the workbook track is verified alongside the
# formula and CF tracks with no extra selector.
oracle-verify:
	@if [ ! -f $(BUILD_DIR)/CMakeCache.txt ]; then \
	  echo "oracle-verify: run 'make build' first"; exit 1; \
	fi
	@$(CMAKE) --build $(BUILD_DIR) --target formulon_oracle_tests --parallel
	@$(CMAKE) --build $(BUILD_DIR) --target formulon_workbook_oracle_tests --parallel
	@# gtest_discover_tests caches the parameter list keyed by binary
	@# timestamp, but our goldens aren't a build input. Force
	@# rediscovery so newly regenerated golden*/*.golden.json files land
	@# in the ctest test list on the next run -- for both the formula
	@# track (golden/) and the workbook track (golden_wb/).
	@rm -f $(BUILD_DIR)/tests/oracle/formulon_oracle_tests*_tests.cmake
	@rm -f $(BUILD_DIR)/tests/oracle/formulon_workbook_oracle_tests*_tests.cmake
	@(cd $(BUILD_DIR) && $(CTEST) -L oracle --output-on-failure --timeout 60)

# One-command contributor onramp. Runs oracle-setup when the venv is
# missing, then dispatches to cli.py contribute, which probes Excel for
# version + locale, picks (or maps) the target, skips generation if the
# committed golden already covers the contributor's Excel build, and
# otherwise runs preflight + oracle_gen + push instructions.
#
# Override the auto-detected target with TARGET=<name> when needed.
oracle-contribute:
	@if [ ! -x "$(ORACLE_PY)" ]; then \
	  echo "oracle-contribute: venv missing -- bootstrapping via oracle-setup..."; \
	  $(MAKE) oracle-setup; \
	fi
	@$(ORACLE_PY) tools/oracle/cli.py contribute $(if $(TARGET),--target $(TARGET),)

# Lists every contribution target known to `tools/oracle/targets.yaml`
# along with the current host's runnable subset.
oracle-contribute-list:
	@if [ ! -x "$(ORACLE_PY)" ]; then \
	  echo "oracle-contribute-list: venv missing -- bootstrapping via oracle-setup..."; \
	  $(MAKE) oracle-setup; \
	fi
	@$(ORACLE_PY) tools/oracle/cli.py list

# -- IronCalc secondary oracle --------------------------------------------
# Imports xlsx fixtures vendored from IronCalc (dual MIT / Apache-2.0)
# into Formulon's golden JSON schema, then runs the secondary verifier
# registered under the `ironcalc` CTest label.
ironcalc-import:
	@if [ ! -x "$(ORACLE_PY)" ]; then \
	  echo "ironcalc-import: run 'make oracle-setup' first (or 'cd tools/oracle && rye sync')"; \
	  exit 1; \
	fi
	@$(ORACLE_PY) tools/oracle/ironcalc_import.py

ironcalc-verify:
	@if [ ! -f $(BUILD_DIR)/CMakeCache.txt ]; then \
	  echo "ironcalc-verify: run 'make build' first"; exit 1; \
	fi
	@$(CMAKE) --build $(BUILD_DIR) --target formulon_ironcalc_oracle_tests --parallel
	@# Force ctest to rediscover parameter list; the goldens aren't a
	@# build input so the cache would otherwise stick.
	@rm -f $(BUILD_DIR)/tests/oracle/formulon_ironcalc_oracle_tests*_tests.cmake
	@(cd $(BUILD_DIR) && $(CTEST) -L ironcalc --output-on-failure --timeout 60)

fuzz-parser:
	@echo "fuzz-parser: not yet implemented (planned for M8)"
	@exit 0

fuzz-xlsx:
	@echo "fuzz-xlsx: not yet implemented (planned for M8)"
	@exit 0

fuzz-eval:
	@echo "fuzz-eval: not yet implemented (planned for M8)"
	@exit 0

bench:
	@echo "bench: not yet implemented (planned for M9)"
	@exit 0

# Local coverage diagnostic. Builds with the gcov-instrumented preset,
# runs fast tests, and prints a per-area report plus the list of files
# the fast suite never touches. Targets in the report (util/value/eval
# 95% / functions 98% / io 90%) are aspirational; this is *not* a CI
# gate. Set FORMULON_COV_STRICT=1 for ad-hoc local enforcement.
# Driver: tools/dev/run_coverage.sh.
coverage:
	bash tools/dev/run_coverage.sh

# Mutation testing report. mull-runner-cxx mutates the C++ source under
# the configured filter, runs the test binary against each mutant, and
# reports the kill rate. Local diagnostic only. CI does NOT gate on
# mutation score; thresholds drift with refactors.
# Set FORMULON_MUT_STRICT=1 for ad-hoc local enforcement (~70% target).
# Driver: tools/dev/run_mutation.sh.
mutation:
	bash tools/dev/run_mutation.sh

# Function implementation coverage report. Scans src/eval/ for registered
# and lazy-dispatched function names and diffs them against the canonical
# catalog at tools/catalog/functions.txt. No build step required.
function-status:
	@python3 tools/catalog/status.py

# Behaviour-vocabulary report. Parallels function-status but operates on
# function sub-behaviours (TEXT format codes, DATEVALUE era strings, CP932
# CHAR / CODE ranges, SEARCH / FIND wildcards). See
# tools/catalog/behaviors.yaml.
behavior-status:
	@python3 tools/catalog/behaviors.py --report
