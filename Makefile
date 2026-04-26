# Formulon convenience targets (see backup/plans/06-build-packaging.md §6.2).
#
# Long-running commands (lint, test, full builds) should redirect output to a
# file and grep the file — never pipe directly into grep.

BUILD_DIR ?= build
GENERATOR ?=
CMAKE ?= cmake
CTEST ?= ctest
CLANG_FORMAT ?= clang-format
CLANG_TIDY ?= clang-tidy

SRC_DIRS := src tests
CPP_GLOB := $(shell find $(SRC_DIRS) -type f \( -name '*.cpp' -o -name '*.h' \) 2>/dev/null)

.PHONY: all build release test test-slow test-all fmt lint clean \
        wasm wasm-debug test-wasm test-python \
        oracle-setup oracle-setup-mac oracle-setup-wsl \
        oracle-gen oracle-verify \
        ironcalc-import ironcalc-verify \
        fuzz-parser fuzz-xlsx fuzz-eval bench coverage \
        function-status behavior-status

all: build

build:
	$(CMAKE) -B $(BUILD_DIR) -DCMAKE_BUILD_TYPE=Debug
	$(CMAKE) --build $(BUILD_DIR) --parallel

release:
	$(CMAKE) -B build-release -DCMAKE_BUILD_TYPE=Release
	$(CMAKE) --build build-release --parallel

test:
	(cd $(BUILD_DIR) && $(CTEST) -LE "SLOW|LOAD" --output-on-failure --timeout 30)

test-slow:
	(cd $(BUILD_DIR) && $(CTEST) -LE "LOAD" --output-on-failure --timeout 120)

test-all:
	(cd $(BUILD_DIR) && $(CTEST) --output-on-failure --timeout 300)

fmt:
	@if [ -n "$(CPP_GLOB)" ]; then \
	  $(CLANG_FORMAT) -i $(CPP_GLOB); \
	else \
	  echo "No C++ sources to format yet."; \
	fi

lint:
	@if [ ! -f $(BUILD_DIR)/compile_commands.json ]; then \
	  echo "compile_commands.json missing; run 'make build' first."; exit 1; \
	fi
	@find src -type f \( -name '*.cpp' -o -name '*.h' \) \
	  -exec $(CLANG_TIDY) -p $(BUILD_DIR) {} +

clean:
	rm -rf $(BUILD_DIR) build-release build-relwithdebinfo \
	       build-asan build-ubsan build-tsan build-coverage

# ---- Stubs wired in later milestones -------------------------------------
wasm:
	@echo "wasm: not yet implemented (planned for M11)"
	@exit 0

wasm-debug:
	@echo "wasm-debug: not yet implemented (planned for M11)"
	@exit 0

test-wasm:
	@echo "test-wasm: not yet implemented (planned for M11)"
	@exit 0

test-python:
	@echo "test-python: not yet implemented (planned for M12)"
	@exit 0

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
# See tools/oracle/README.md and backup/plans/07-oracle-testing.md.
ORACLE_DIR := tools/oracle
ORACLE_VENV := $(ORACLE_DIR)/.venv
ORACLE_GEN := $(ORACLE_VENV)/bin/python tools/oracle/oracle_gen.py

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
	@echo "Then edit tools/oracle/targets.yaml and set win_python under win-365-ja_JP."
	@$(ORACLE_VENV)/bin/python tools/oracle/cli.py setup --target win-365-ja_JP || true

oracle-gen:
	@if [ ! -x "$(ORACLE_VENV)/bin/python" ]; then \
	  echo "oracle-gen: run 'make oracle-setup' first"; \
	  exit 1; \
	fi
	@$(ORACLE_GEN) $(if $(SUITE),--suite $(SUITE),) $(if $(TARGET),--target $(TARGET),)

oracle-verify:
	@if [ ! -f $(BUILD_DIR)/CMakeCache.txt ]; then \
	  echo "oracle-verify: run 'make build' first"; exit 1; \
	fi
	@$(CMAKE) --build $(BUILD_DIR) --target formulon_oracle_tests --parallel
	@# gtest_discover_tests caches the parameter list keyed by binary
	@# timestamp, but our goldens aren't a build input. Force
	@# rediscovery so newly regenerated golden/*.golden.json files land
	@# in the ctest test list on the next run.
	@rm -f $(BUILD_DIR)/tests/oracle/formulon_oracle_tests*_tests.cmake
	@(cd $(BUILD_DIR) && $(CTEST) -L oracle --output-on-failure --timeout 60)

# -- IronCalc secondary oracle --------------------------------------------
# Imports xlsx fixtures vendored from IronCalc (dual MIT / Apache-2.0)
# into Formulon's golden JSON schema, then runs the secondary verifier
# registered under the `ironcalc` CTest label.
ironcalc-import:
	@if [ ! -x "$(ORACLE_VENV)/bin/python" ]; then \
	  echo "ironcalc-import: run 'make oracle-setup' first (or 'cd tools/oracle && rye sync')"; \
	  exit 1; \
	fi
	@$(ORACLE_VENV)/bin/python tools/oracle/ironcalc_import.py

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

coverage:
	@echo "coverage: not yet implemented (planned for M9)"
	@exit 0

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
