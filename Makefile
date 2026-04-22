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
        oracle-setup oracle-gen oracle-verify \
        fuzz-parser fuzz-xlsx fuzz-eval bench coverage

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

oracle-setup:
	@echo "oracle-setup: not yet implemented (planned for M7)"
	@exit 0

oracle-gen:
	@echo "oracle-gen: not yet implemented (planned for M7)"
	@exit 0

oracle-verify:
	@echo "oracle-verify: not yet implemented (planned for M7)"
	@exit 0

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
