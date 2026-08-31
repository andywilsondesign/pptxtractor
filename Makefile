# pptxtractor - build the CoreGraphics renderer.
# Needs the Xcode Command Line Tools (xcode-select --install).
BIN := bin/pdfrender
SRC := src/pdfrender.swift

all: $(BIN)

$(BIN): $(SRC)
	@mkdir -p bin
	swiftc -O -o $(BIN) $(SRC)
	@echo "built $(BIN)"

test: $(BIN)
	python3 tests/test_core.py

clean:
	rm -f $(BIN)

.PHONY: all test clean
