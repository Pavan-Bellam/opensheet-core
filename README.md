<p align="center">
  <img src="assets/banner.svg" alt="OpenSheet Core — Fast, memory-efficient spreadsheet I/O for Python, powered by Rust" width="100%">
</p>

<p align="center">
  <a href="https://github.com/0xNadr/opensheet-core/actions/workflows/ci.yml"><img src="https://github.com/0xNadr/opensheet-core/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://github.com/0xNadr/opensheet-core/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.9%E2%80%933.13-blue.svg" alt="Python 3.9–3.13"></a>
</p>

<p align="center">
  <a href="#features">Features</a> &nbsp;&bull;&nbsp;
  <a href="#benchmarks">Benchmarks</a> &nbsp;&bull;&nbsp;
  <a href="#installation">Installation</a> &nbsp;&bull;&nbsp;
  <a href="#quick-start">Quick Start</a> &nbsp;&bull;&nbsp;
  <a href="#api-reference">API</a> &nbsp;&bull;&nbsp;
  <a href="#roadmap">Roadmap</a> &nbsp;&bull;&nbsp;
  <a href="#contributing">Contributing</a>
</p>

---

## Why OpenSheet Core?

Existing Python spreadsheet libraries force you to choose between performance, memory efficiency, broad format support, and easy installation. OpenSheet Core eliminates that tradeoff with a native Rust core exposed through a clean Python API — installable with a single `pip install`.

## Features

- **Streaming XLSX reader** — row-by-row iteration without loading the entire file into memory
- **Streaming XLSX writer** — write millions of rows with constant memory usage
- **Typed cell extraction** — strings, numbers, booleans, and empty cells are returned as native Python types
- **Context manager support** — Pythonic `with` statement for safe resource management
- **Cross-platform** — tested on Linux, macOS, and Windows across Python 3.9–3.13
- **Zero Python dependencies** — single native extension, no dependency tree to manage

## Benchmarks

Benchmarked against [openpyxl](https://openpyxl.readthedocs.io/) on a 100,000-row dataset:

| Operation | OpenSheet Core | openpyxl | Speedup | Memory |
|-----------|---------------|----------|---------|--------|
| **Write** | ~0.7s | ~1.8s | **2.5x faster** | **~300x less** |
| **Read** | ~0.9s | ~2.4s | **2.7x faster** | Low & constant |

> Memory usage stays flat regardless of file size thanks to streaming architecture.

## Installation

### From source (requires Rust toolchain)

```bash
pip install maturin
git clone https://github.com/0xNadr/opensheet-core
cd opensheet-core
maturin develop --release
```

> Prebuilt wheels on PyPI are coming soon.

## Quick Start

### Reading an XLSX file

```python
from opensheet_core import read_xlsx

workbook = read_xlsx("report.xlsx")

for sheet_name, rows in workbook:
    print(f"Sheet: {sheet_name}")
    for row in rows:
        print(row)  # List of typed Python values
```

### Writing an XLSX file

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "Active"])
    writer.write_row(["Alice", 30, True])
    writer.write_row(["Bob", 25, False])
```

## API Reference

### `read_xlsx(path: str) -> list[tuple[str, list[list]]]`

Reads an XLSX file and returns a list of `(sheet_name, rows)` tuples. Each row is a list of typed Python values (`str`, `float`, `bool`, or `None`).

### `XlsxWriter(path: str)`

Streaming XLSX writer. Use as a context manager.

| Method | Description |
|--------|-------------|
| `add_sheet(name: str)` | Create a new worksheet |
| `write_row(values: list)` | Write a row of values to the current sheet |

## Architecture

```
┌──────────────────────────┐
│      Python API          │  ← opensheet_core (PyO3 bindings)
├──────────────────────────┤
│      Rust Core           │  ← Streaming parser & writer
│  ┌────────┐ ┌──────────┐ │
│  │ Reader │ │  Writer  │ │
│  │ (SAX)  │ │ (Stream) │ │
│  └────────┘ └──────────┘ │
├──────────────────────────┤
│  quick-xml  │    zip     │  ← Dependencies
└──────────────────────────┘
```

## Roadmap

- [x] XLSX reading with typed cell extraction
- [x] Streaming XLSX writing with low memory usage
- [x] Python bindings via PyO3
- [x] CI across Linux, macOS, Windows (Python 3.9–3.13)
- [x] Benchmarks vs openpyxl
- [ ] Prebuilt wheels on PyPI
- [ ] Formula writing support
- [ ] Merged cell metadata
- [ ] Basic cell styling
- [ ] Broader test corpus & fuzzing

## Project Status

**Working prototype** — functional reader and writer with 22 passing tests. The API may change before 1.0.

## Contributing

Contributions are welcome! Here are some great ways to get involved:

- Report bugs or real-world spreadsheet edge cases
- Submit representative sample files for testing
- Suggest benchmark scenarios
- Improve documentation
- Open PRs for roadmap items

## License

[MIT](LICENSE)

---

<p align="center">
  <sub>Built with Rust and PyO3 &nbsp;|&nbsp; Open digital infrastructure for the Python ecosystem</sub>
</p>
