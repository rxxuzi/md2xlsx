# Styling

[日本語版](style.ja.md)

How cells end up looking: what `md2xlsx` applies on its own, the three
ways you can change it, and what wins when they overlap. The syntax of each
construct is in the [syntax reference](syntax.md); this page is about
appearance.

## Defaults

Everything is Arial 11pt unless something below changes it.

| Element | Look |
|---|---|
| Table header (the row above `\|---\|`, or row 1 of a ```` ```table ```` fence) | Blue fill, white bold text, centered |
| Table body row | Thin light-gray bottom border, vertically centered |
| `##` heading | Bold 13pt, left-aligned |
| `###` and deeper | Bold 11pt, left-aligned |
| Hyperlink | Blue, underlined |
| `` `code` `` cell | Consolas, dark red |
| Document mode | Thin black grid over every used cell; `#` and `##` rows bold |
| Column width | Widest value + 3, clamped to 8–60; CJK characters count as 2 |

## Three ways to change it

### 1. Inline Markdown

`**bold**`, `*italic*` and `` `code` `` inside a cell set that cell's font.
Typed cells — formulas, `$` / `¥` amounts, percentages, dates — and cells
containing a link ignore inline markup; the value wins.

### 2. Cell decorators

`{bg:…}`, `{fg:…}` and `{align:…}` prefixed to a cell. (`{comment:…}` and
`{dropdown:…}` are not about looks; see the syntax reference.)

```markdown
| {bg:lightyellow}{fg:red}{align:center}Warning |
```

### 3. Column styles — `@style`

One directive under `#[name]` styles whole columns. Pairs of
`COLUMN=values` are comma-separated; the values are space-separated and
recognised by their shape, like a CSS shorthand, so order does not matter.

| Value | Effect |
|---|---|
| number | Column width |
| color name or `#RRGGBB` | Background fill |
| `bold` / `italic` | Font style, added to whatever the cell already has |
| `left` / `center` / `right` | Horizontal alignment |

```markdown
#[Tasks]
@style(A=6 center, B=30, D=14 lightgreen bold)
```

- Column letters are case-insensitive; `a=6` works.
- The width applies to the whole Excel column. Fill, font and alignment are
  applied cell by cell to the used rows only — including plain text and
  headings in that column — so the sheet stays clean below the data.
- Unlisted columns auto-size. Anything unrecognised in a pair is ignored;
  `@style(A=huge)` does nothing.
- A color in `@style` always means background. Font color is only
  available per cell with `{fg:…}`.

## Colors

`red green blue yellow orange purple pink gray grey white black lightblue
lightgreen lightred lightyellow`, or any `#RRGGBB`. The same list applies
everywhere a color is accepted.

## What wins

From strongest to weakest:

1. **The table header look.** Header rows are never touched by `@style`,
   and `{bg:…}` / `{fg:…}` / `{align:…}` on a header cell are overridden by
   the blue, white, centered header style. `{comment:…}` and `{dropdown:…}`
   survive.
2. **Cell decorators and inline Markdown.** A `{bg:red}` cell in a
   `B=yellow` column stays red; `{align:left}` in a `center` column stays
   left.
3. **Column style.** Fills in whatever the cell did not set itself.
4. **Defaults.**

Column `bold` / `italic` are *added* to a cell's font rather than replacing
it, so `*note*` in a `bold` column comes out bold italic.
