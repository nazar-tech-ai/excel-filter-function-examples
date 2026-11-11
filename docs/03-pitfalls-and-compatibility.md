# Pitfalls & Compatibility

Practical notes from working with these formulas on real data. Read this before adapting the workbook to your own files.

## Contents

- [Error reference](#error-reference)
- [Text that looks equal but isn't](#text-that-looks-equal-but-isnt)
- [Text vs number](#text-vs-number)
- [Hard-coded ranges](#hard-coded-ranges)
- [Spill behaviour](#spill-behaviour)
- [Case sensitivity and wildcards](#case-sensitivity-and-wildcards)
- [Compatibility](#compatibility)
- [Best-practice checklist](#best-practice-checklist)

## Error reference

| Error | Typical cause in this workbook | Fix |
|---|---|---|
| `#SPILL!` | Something (a value, a merged cell, another spill) sits in the cells the result needs | Clear the blocked cells; select the error cell and use the warning icon to find the blocker |
| `#CALC!` | `FILTER` found **no** matching rows and has no `if_empty` argument | Add the third argument: `=FILTER(data, mask, "No record found")` |
| `#N/A` | `VLOOKUP`/`XMATCH` did not find the key; or `INDEX`/`MATCH` found no row | Intended in 01-P (it is tested with `ISNA`). Elsewhere, check spelling, spaces and data types |
| `#VALUE!` | `SEARCH` did not find the text; or array sizes in `FILTER` do not match | In a mask, wrap with `ISNUMBER(...)`. Check that the mask and the data cover the **same rows** |
| `#NAME?` | The function does not exist in your Excel version (for example `FILTER` in Excel 2019) | Use Excel 365 / 2021+ — see [Compatibility](#compatibility) |
| `#REF!` | A referenced range or sheet was deleted | Re-point the formula; do not delete sheets that others depend on |

## Text that looks equal but isn't

Two cells can look identical and still fail an exact comparison.

- **Trailing / leading spaces** — `"DEPT A "` is not `"DEPT A"`. The `Majlis_Name` column in `DM-DATA` deliberately keeps a trailing space, as exports often do. If you filter on such a column, either match the space or normalize with `TRIM`: `=FILTER(data, TRIM(col) = "DEPT A")`.
- **Non-breaking spaces** from web or PDF copy-paste (`CHAR(160)`) — `TRIM` does not remove them; use `SUBSTITUTE(col, CHAR(160), " ")` first.
- **Case** — `=`, `FILTER`, `XMATCH`, `SUMIFS` and `SEARCH` ignore case. Use `EXACT()` when case matters.

**Diagnose quickly:** `=LEN(A2)` versus the length you expect, or `=CODE(RIGHT(A2,1))`.

## Text vs number

`"20"` and `20` are different things.

```excel
="20" = 20              ' FALSE
=TEXT(DATE(2023,3,5),"dd")   ' "05"  (text, zero-padded)
```

This is why the day-of-month inputs in 11-P and 14-P are stored as **text**, and why `"5"` would not match `"05"`. Safer alternatives:

- Compare numerically: `DAY(dates) = 5`
- Convert the input: `TEXT(A3, "00")`

Numbers stored as text in the *data* cause the same trouble with `SUMIFS` (they are ignored by the sum range). A green triangle in the corner of a cell is Excel's hint.

## Hard-coded ranges

Most formulas here use fixed ranges such as `'SYS-DATA'!A8:H1298`. If you paste a longer export, the extra rows are silently ignored.

Options, from simplest to most robust:

1. **Over-size the range** — `A8:H5000`. Blank rows are harmless for `FILTER`, but `UNIQUE` will include one blank entry; wrap with `FILTER(UNIQUE(x), UNIQUE(x) <> "")` if that matters.
2. **Excel Tables** — select the data, press **Ctrl+T**, then write `FILTER(tblSys, tblSys[Location] = A3)`. Tables grow with the data and make formulas self-documenting.
3. **Dynamic ranges** — `INDEX`/`COUNTA` to build the last row. More complex; use Tables first.

Also keep ranges **the same height**. `FILTER` fails with `#VALUE!` when the data is rows 8–1298 but the mask covers rows 7–1298.

## Spill behaviour

- Enter the formula once, in the top-left cell. Do not copy it down.
- Do not type or paste into the spill area; clear it instead.
- Refer to the whole result with `A6#`. In a data-validation list, use `=$A$6#` so the dropdown follows the result.
- Spills cannot be created inside an Excel Table. Keep spill formulas outside Tables.
- A spilled result is **not** sorted — it follows source order. Wrap in `SORT()` when you need order.

## Case sensitivity and wildcards

| Function | Case-sensitive? | Wildcards? |
|---|---|---|
| `=` comparison, `FILTER`, `XMATCH` (default) | No | No (`XMATCH` supports them only in match mode 2) |
| `SEARCH` | No | Yes — `?`, `*`, escape with `~` |
| `FIND` | **Yes** | No |
| `SUMIFS` criteria | No | Yes — `*` and `?` work inside criteria text |

If the search text may contain `*`, `?` or `~`, escape them or use `FIND`.

## Compatibility

| Feature | Excel 365 / 2021+ | Excel 2019 | Excel 2016 and earlier |
|---|---|---|---|
| `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE` | ✅ | ❌ | ❌ |
| `XMATCH` | ✅ | ❌ | ❌ |
| `TEXTJOIN` | ✅ | ✅ | ❌ |
| `CHOOSECOLS`, `DROP` | ✅ (recent builds) | ❌ | ❌ |
| `FORMULATEXT`, `SUMIFS`, `SUBTOTAL` | ✅ | ✅ | ✅ |

**Opening the workbook in an older Excel** typically shows the last saved values, but the new functions appear as `_xlfn.…` names and return `#NAME?` once recalculated.

**Google Sheets** has `FILTER` and `UNIQUE`, but not `XMATCH` (use `MATCH`), and array-constant separators depend on locale. Results have not been verified in this repository.

**LibreOffice** adds `FILTER`, `UNIQUE` and `XMATCH` in version 24.8. Not verified here.

## Best-practice checklist

- [ ] Put **inputs** (criteria, dropdowns) in clearly labelled cells above the result, never inside the formula.
- [ ] Give `FILTER` an `if_empty` argument instead of wrapping in `IFERROR`, so real errors stay visible.
- [ ] Derive lengths from data (`LEN(prefix)`) instead of typing numbers.
- [ ] Keep one formula per spill; do not copy spill formulas down.
- [ ] Store source data as an Excel Table when it will keep growing.
- [ ] Use `FORMULATEXT` on documentation/teaching sheets so readers can see the logic.
- [ ] Check text/number types and stray spaces **before** blaming the formula.
- [ ] Test every lookup with a key you know is missing, and one you know is present.
