# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`md2xlsx` converts a small flavored-Markdown dialect into multi-sheet `.xlsx` workbooks
(tables, formulas, currency/date/percent typing, colors, dropdowns, and a Japanese-style
column-indented "document mode"). The whole tool is one file, `md2xlsx.py`; the repo also
doubles as a Claude skill folder via `SKILL.md`.

## Commands

Only runtime dependency is `openpyxl`. Either install it with pip or let `uv` resolve
it from `pyproject.toml` (a `.venv/` is created and is gitignored).

```sh
# run
python md2xlsx.py input.md              # writes input.xlsx, prints "path: sheet, sheet"
python md2xlsx.py input.md -o out.xlsx -d   # -d dumps every sheet as a text grid (no Excel needed)
python md2xlsx.py - -o out.xlsx < input.md  # stdin needs -o; -q / -v / --version / -h as usual
uv run md2xlsx.py input.md              # same, uv installs openpyxl for you

# test / lint (pip: `pip install openpyxl pytest ruff` first)
python -m pytest tests/                                     # all tests
python -m pytest tests/test_md2xlsx.py::test_decorators     # one test
ruff check .                            # config in pyproject.toml (line-length 100)
uv run pytest                           # uv equivalents; pytest/ruff come from the dev group
uv run ruff check .

python md2xlsx.py examples/sample.md    # regenerate the committed examples/sample.xlsx
```

Tests import `md2xlsx` by inserting the repo root into `sys.path` — there is no package
dir and no install step needed.

## Architecture (`md2xlsx.py`)

Four layers, top to bottom of the file:

0. **Diagnostics** — dialect mistakes are reported rustc-style (`error:` / `warning:`,
   `--> file:line:col`, the source line, carets under the token, `= help:`). Low-level
   helpers raise `Bad(msg, token, help)` for hard errors; `Conv.guard()` catches them per
   line and records them in `Diags` with the current line index (`Sheet.at`), then keeps
   parsing so several errors surface at once. Warnings go through `Sheet.warn()`. Errors
   mean the workbook is not written (exit 1); warnings don't block. `Diags.render()` uses
   `width()` so carets line up under CJK text; `Paint` adds ANSI color only on a tty.
   When adding a rule: raise `Bad` if the output would be wrong or lossy, warn if it is
   merely a layout smell (table in a document sheet, `\|`, widths on a document sheet).

1. **Cell-level helpers** — `put(sh, col, raw, font)` is the single entry point for
   writing any cell (it takes the `Sheet` so it can warn). Its ordering is the important part:
   `decos()` strips `{key:value}` tags → `typed()` tries formula / `$` / `¥` / date /
   percent → else hyperlink → else `num(strip_md(text))` with `inline_font()`.
   Consequently a typed cell (formula, currency…) never gets inline bold/italic or a
   link, and a link cell never gets numeric typing. `apply_decos()` runs last; note
   openpyxl `Font` is immutable so `fg` rebuilds the font preserving other attributes.

2. **`Sheet`** — one per `#[name]`. `self.row` is 0-based and incremented *before* each
   write, so rows are 1-based in the workbook. `self.cols` tracks the widest row for the
   `@filter` range. Header styling for pipe tables is applied *retroactively*: a
   `|---|` separator line calls `head_row()` on the previous row; a pipe table with no
   separator has no header. CSV/TSV fenced tables always treat row 1 as header.
   `finish()` (called on sheet switch and at EOF) applies `@style` widths, then
   `auto_width` for the remaining columns (CJK chars count as 2 via `east_asian_width`),
   grid border in document mode, then the rest of `@style` (bg / bold / italic / align)
   cell by cell via `style_cell()` — which only fills in what a cell hasn't set itself,
   so header rows and `{bg:…}` / `{align:…}` decorators win — then freeze/filter.

3. **`Conv`** — line-driven state machine in `step()`. Priority order: inside a
   ```` ```table ```` fence → `#[sheet]` (immediately consumes following `@directive`
   lines via `read_dirs`) → leading directives with no sheet yet (creates `Sheet1`) →
   `body()`. In `body()`: fence open → `---` blank row → heading → any line containing
   `|` (not starting with `#`) is a pipe row → empty line ends the current table →
   `Sheet.text()` for everything else (plain text and `- ` / `1. ` list items).

Non-obvious dispatch rules worth knowing before changing syntax:

- In default mode a single `# heading` is **not** a heading; it falls through to plain
  text. Only `##`+ are headings. In `@document` mode every level is a heading and
  the level number is the column index (`#`→A, `##`→B…), bold for levels ≤ 2.
- Body text in `@document` mode: `Sheet.level` remembers the last heading level and
  `text()` writes plain lines and list items at `level + 1` (plus one column per two
  spaces of list indent). In default mode everything lands in column A. `- ` markers
  are dropped, `1. ` numbers are kept as text. This is what lets a document sheet have
  prose under a heading instead of forcing every sentence to be a heading.
- Directives are only recognized directly under `#[name]` or at the very top of the
  file; anywhere else `@foo` is plain text.
- Sheet names are truncated to 31 chars (Excel limit) silently.
- Dates are naive `datetime` on purpose (Excel has no tz); ruff `DTZ001` is ignored for
  this reason.

## Docs layout

`README.md` is a short landing page (install, usage, one teaser
example). The real documentation is in `docs/`, every page in an EN + JA pair
(`x.md` / `x.ja.md`): `syntax` (the dialect reference), `style` (default looks, `@style`,
decorators, and the precedence between them), `build` (PyInstaller / Nuitka /
`uv tool install` / CI), `development` (setup, tests, code layout). Keep pairs in sync
when editing one.

## When changing syntax

The syntax is documented in several places that must stay in sync: `docs/syntax.md`,
`docs/syntax.ja.md` (plus `docs/style.md` / `style.ja.md` for anything that affects how
cells look), `SKILL.md` (deliberately self-contained — an agent must be able to write a
valid file from that page alone, so keep its rules and templates complete, not just a
cheat sheet), and `examples/sample.md` (which is
meant to exercise every feature; its `.xlsx` output is committed and exempted from
`.gitignore`). The teaser in the root READMEs only needs touching if it uses the changed
feature. Add a test in `tests/test_md2xlsx.py` following the existing black-box pattern:
`build(tmp_path, md_text)` → `load_workbook` → assert on cell values/styles.
