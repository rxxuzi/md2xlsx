# md2xlsx

[![CI](https://github.com/rxxuzi/md2xlsx/actions/workflows/ci.yml/badge.svg)](https://github.com/rxxuzi/md2xlsx/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)

Write Markdown, get Excel. A small flavored-Markdown dialect for building
multi-sheet `.xlsx` workbooks from plain text: tables, formulas, colors,
dropdowns, and a "document mode" for the column-indented spec sheets common
in Japanese business.

[Docs](docs/README.md)

## Install

One file, one dependency (`openpyxl`), Python 3.8+.

With [uv](https://docs.astral.sh/uv/) there is nothing to install; run it
straight from GitHub:

```sh
uvx --from git+https://github.com/rxxuzi/md2xlsx md2xlsx input.md
```

Or grab the script and use pip:

```sh
curl -O https://raw.githubusercontent.com/rxxuzi/md2xlsx/main/md2xlsx.py
pip install openpyxl
```

Want a standalone `.exe`? See [docs/build.md](docs/build.md).

## Usage

```sh
python md2xlsx.py input.md            # writes input.xlsx
python md2xlsx.py input.md -o out.xlsx
cat input.md | python md2xlsx.py - -o out.xlsx
```

| Option | |
|---|---|
| `-o, --output FILE` | Output path (default: input with `.xlsx`) |
| `-d, --dump` | Also print each sheet as a text grid — check the layout without Excel |
| `-q, --quiet` | No summary line |
| `-v, --verbose` | Per-sheet summary: rows, columns, formulas, directives |
| `--version`, `-h` | |

Mistakes in the Markdown are reported rustc-style, with the line, a caret
under the offending token, and a one-line fix. Errors stop the conversion;
warnings don't.

```
error: unknown color `yelow`
 --> input.md:7:12
  |
7 | | Q1 | {bg:yelow}$1200 |
  |            ^^^^^
  |
  = help: named colors: red green blue yellow orange …; or #RRGGBB
```

With uv, the downloaded script runs without installing `openpyxl` yourself:

```sh
uv run --with openpyxl md2xlsx.py input.md
uv run md2xlsx.py input.md            # inside a clone of this repo
```

## A taste

```markdown
#[Sales]
@freeze(1)

| Quarter | Revenue | Margin | Status |
|---------|---------|--------|--------|
| Q1 | $1,200 | 30% | {dropdown:open,closed}closed |
| Q2 | $1,450 | 29% | {dropdown:open,closed}open |
| **Total** | $$=SUM(B2:B3)$$ | {bg:lightyellow}$$=AVERAGE(C2:C3)$$ | |

#[Spec]
@document

# 1. Login
## 1.1 Enter user ID
## 1.2 Enter password
```

That gives you a two-sheet workbook. `Sales` has a frozen, styled header
row, real currency and percent cells, live formulas, a dropdown per row and
one highlighted cell. `Spec` is a column-indented outline with a full grid
border, the way Excel specification documents are laid out in Japan.

The full dialect — sheets, directives, pipe and CSV tables, cell typing,
decorators, document mode, and the gotchas — is in
[docs/syntax.md](docs/syntax.md). [examples/sample.md](examples/sample.md)
uses every feature in one file.

## Use as a Claude skill

The repo doubles as a skill folder. Copy or clone it into your skills
directory and Claude will pick up `SKILL.md`.

## License

MIT
