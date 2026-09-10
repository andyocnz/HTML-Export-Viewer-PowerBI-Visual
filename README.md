# HTML Export Viewer

Renders an HTML measure (a styled table, typically built by DAX) and lets the report user
export it to PDF, Excel or Word. Unlike most custom visuals, this one is not a drop-in "bind a
field and go" tool — the measure feeding it has to produce HTML in a specific shape, and if you
want the Excel export to contain live, recalculating formulas (not just static numbers), the
measure has to opt into that explicitly. This document is that contract.

Support: info@pbiwise.com

---

## 1. Basic use (required)

**Data binding**: the visual has one data role, "HTML content" (`content`, a Measure). Bind a
single DAX measure that returns a string of HTML. Only one measure - if it evaluates to blank
(e.g. no row selected), the visual shows a placeholder instead of the export buttons.

**Minimum shape the measure must return**:

```
"<html><head><meta charset=""UTF-8""></head>"
& "<body>"
& "<style>body{font-family:Arial;color:#000;} table{border-collapse:collapse;} td,th{border:1px solid #ccc;padding:4px;}</style>"
& "<table><tr><th>Column A</th><th>Column B</th></tr>"
& "<tr><td>Row 1</td><td>$100</td></tr>"
& "</table>"
& "</body></html>"
```

Notes on that shape:

- The outer `<html><head>...</head><body>...</body></html>` wrapper is expected, but you don't
  need to worry about whether your `<style>` block ends up inside `<head>` or `<body>` - the
  visual pulls out every `<style>` tag itself, wherever it is, before doing anything else, so
  either placement works.
- At least one `<table>` is required for the PDF and Word export buttons to do anything ("No
  table found in the content to export" otherwise). The Excel export also walks any `<p>`,
  `<div>` or heading (`<h1>`-`<h3>`) at the top level of `<body>`, laid out as one continuous
  sheet in document order.
- Cell values that look like `$1,234` (a leading `$`, digit groups, optional decimals) are
  recovered as real numbers in the Excel export, not text - so totals stay summable.
- A row gets `class="subtotal-row"` to make the Excel export SUM the numeric columns of the
  data rows above it automatically.

**PDF and Word export** don't need anything beyond a `<table>` - they render whatever tables are
present, using each cell's live computed CSS (background, color, bold, alignment, borders) so
the output matches what's on screen. Word is delivered as a WordprocessingML `.xml` file (Word
opens it via File > Open) rather than a real `.docx` - the Power BI download API's file-type
allowlist doesn't include `.docx`.

**Formatting pane options** (under "Export toolbar"): show/hide the button row, its corner
position, custom text for each of the three buttons, and the export file name.

---

## 2. Live Excel formulas (optional, not obvious - read this before asking why a number is static)

By default, every exported number is a static value - correct at export time, but it won't
recalculate if someone tweaks an input in the workbook. If you want specific cells to become
**live formulas** instead (e.g. "this cell = a base rate times an escalation factor, both of
which live in a small assumptions block on the sheet"), the measure has to declare two things.

### 2a. Declare the named values ("assumptions")

Add one hidden `<div>` anywhere in `<body>`:

```html
<div hidden data-role="export-assumptions"
     data-assumptions="[{&quot;ref&quot;:&quot;baseRent&quot;,&quot;label&quot;:&quot;Base rent (current)&quot;,&quot;value&quot;:200000,&quot;numFmt&quot;:&quot;$#,##0&quot;},
                        {&quot;ref&quot;:&quot;rentRate&quot;,&quot;label&quot;:&quot;Rent escalation rate&quot;,&quot;value&quot;:0.03,&quot;numFmt&quot;:&quot;0.0%&quot;}]">
</div>
```

`data-assumptions` is a JSON array. Each entry needs:

| field    | type   | meaning                                                              |
|----------|--------|-----------------------------------------------------------------------|
| `ref`    | string | the name formula templates will use to point at this value (letters/underscore/digits, must start with a letter or `_`) |
| `label`  | string | shown as the row label in the small "Assumptions" block the visual writes to the Excel sheet |
| `value`  | number | the actual value                                                     |
| `numFmt` | string | (optional) an Excel number format code, e.g. `"$#,##0"` or `"0.0%"`  |

Because this is an HTML attribute, every `"` inside the JSON must be written as `&quot;` - see
the example above. In DAX, build it with `SUBSTITUTE(json, """", "&quot;")`.

### 2b. Point a cell at those values with `data-formula`

Any `<td>` can carry:

```html
<td class="num" data-formula="{baseRent}*(1+{years}*{rentRate})" data-formula-vars="years:3">$218,000</td>
```

- `data-formula` is a template. `{baseRent}` and `{rentRate}` resolve against the `ref`s
  declared in the assumptions block above. `{years}` resolves against `data-formula-vars`
  instead - a comma-separated `key:number` list, for values that are specific to this one cell
  (not shared constants).
- The cell's own visible text (`$218,000`) is still required - it's what the *result* of the
  formula is set to, and it's what PDF/Word export and the on-screen visual use, since they
  don't evaluate formulas at all.
- If a template references a `{token}` that isn't found in either the assumptions map or
  `data-formula-vars`, or if the resolved formula contains anything outside cell references,
  numbers, `SUM`/`AVERAGE`/`MIN`/`MAX`/`ROUND`, and basic arithmetic/punctuation, the visual
  silently falls back to writing the cell's static value instead of a formula. This is a
  deliberate safety check (a resolved formula string lands directly in a file the user opens in
  Excel), not a bug - if your formula isn't showing up as live in the export, this is the first
  thing to check: is every `{token}` actually declared somewhere, and does the resolved formula
  only contain the allowed characters?

### 2c. Worked example (DAX)

```
VAR _AssumptionsJson =
    "[{""ref"":""baseRent"",""label"":""Base rent (current)"",""value"":" & _BaseRent & ",""numFmt"":""$#,##0""},"
    & "{""ref"":""rentRate"",""label"":""Rent escalation rate"",""value"":" & _RentRate & ",""numFmt"":""0.0%""}]"
VAR _AssumptionsAttr = SUBSTITUTE(_AssumptionsJson, """", "&quot;")
VAR _AssumptionsBlock =
    "<div hidden data-role=""export-assumptions"" data-assumptions=""" & _AssumptionsAttr & """></div>"

VAR _RentFormulaAttr =
    IF( _RowIsEscalating,
        " data-formula=""{baseRent}*(1+{years}*{rentRate})"" data-formula-vars=""years:" & _YearsSinceReview & """",
        ""
    )

RETURN
    ...
    & "<td class=""num""" & _RentFormulaAttr & ">" & FORMAT(_RentValue,"$#,##0") & "</td>"
    ...
    & _AssumptionsBlock
    & "</body></html>"
```

This pattern (constants in one small `_AssumptionsJson`/`_AssumptionsBlock` block, a per-row
`IF(...)` building a `data-formula` attribute only on rows that need it) is the template to
copy for any new use of this feature - the visual has no built-in notion of what the values
mean; that meaning lives entirely in the measure.

---

## 3. Troubleshooting

- **Blank visual, but buttons are enabled**: the measure is returning non-empty content, but
  something about its shape is off. Right-click the measure in the Data pane → "Show as a
  table" to see the raw HTML string directly, bypassing the visual entirely - confirms whether
  the problem is in the DAX or in how the visual is rendering it.
- **"No table found in the content to export"**: the HTML has no `<table>` element in `<body>`.
- **A cell that should be a live formula is static in the exported workbook**: see 2b above -
  almost always an unresolved `{token}` or a formula that didn't pass the safety check.
- **Word file opens as garbled text or a browser tab instead of Word**: expected if you
  double-click it in File Explorer - it's an XML file, not a `.docx`. Open it from inside Word
  via File > Open instead.
