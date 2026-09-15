# Development

[日本語版](development.ja.md)

## Setup

The only runtime dependency is `openpyxl`. Tests need `pytest`; linting
uses `ruff` (configured in `pyproject.toml`, line length 100).

With uv nothing needs installing up front; each command below resolves what
it needs into `.venv/` (gitignored) on first run. With pip:

```sh
pip install openpyxl pytest ruff
```

## Run

```sh
uv run md2xlsx.py examples/sample.md -o out.xlsx
python md2xlsx.py examples/sample.md -o out.xlsx -d   # -d prints the sheets as text
```

## Test

```sh
uv run pytest
uv run pytest tests/test_md2xlsx.py::test_decorators   # one test

python -m pytest tests/
```

Tests are black-box: write a Markdown string to a temp file, convert it,
`load_workbook` the result and assert on cell values and styles. Follow the
`build(tmp_path, text)` helper in `tests/test_md2xlsx.py` for new ones. The
tests import `md2xlsx` by putting the repo root on `sys.path`, so no install
step is needed.

## Lint

```sh
uv run ruff check .
ruff check .
```

`DTZ001` is ignored on purpose: Excel dates have no timezone, so naive
`datetime` is correct here.

## How the code is organized

Everything is in `md2xlsx.py`, roughly in three layers from top to bottom:

1. **Cell helpers.** `put()` is the single entry point for writing a cell.
   It strips `{key:value}` decorators, then tries the typed forms (formula,
   currency, date, percent), then hyperlinks, then falls back to text with
   inline bold/italic/code. That order is why a formula cell can't also be
   bold and a link cell can't be a number.
2. **`Sheet`.** One per `#[name]`. Tracks the current row, turns pipe and
   CSV rows into cells, and in `finish()` applies widths, the document-mode
   grid, freeze panes and the auto-filter.
3. **`Conv`.** A line-by-line state machine. Handles the ```` ```table ````
   fence, `#[sheet]` plus its directives, then dispatches each body line to
   heading / table row / blank / text.

## When changing the syntax

The syntax is described in several places that must stay in sync:

- `docs/syntax.md` and `docs/syntax.ja.md` — the reference
- `SKILL.md` — the cheat sheet Claude reads
- `examples/sample.md` — should exercise every feature; regenerate the
  committed `examples/sample.xlsx` with `python md2xlsx.py examples/sample.md`
- `tests/test_md2xlsx.py` — add a case
