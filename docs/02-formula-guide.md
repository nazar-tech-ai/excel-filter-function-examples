# Formula Guide

A sheet-by-sheet explanation of every formula in `Filter_Function_Examples.xlsx`. Each section follows the same pattern:

- **Goal** — the business question being answered
- **Formula** — exactly what is in the workbook (cell addresses included)
- **How it works** — evaluation, step by step
- **Notes** — pitfalls, variations, and optional improvements

Data layout is described in [01-data-model.md](01-data-model.md). Common errors are covered in [03-pitfalls-and-compatibility.md](03-pitfalls-and-compatibility.md).

## Contents

- [Primer: how dynamic arrays work](#primer-how-dynamic-arrays-work)
- [01-P — Compare two lists](#01-p--compare-two-lists)
- [02-P — Filter by a list of values](#02-p--filter-by-a-list-of-values)
- [03-P — Unique lists with serial numbers](#03-p--unique-lists-with-serial-numbers)
- [04-P — Two-criteria lookup](#04-p--two-criteria-lookup)
- [05-P / 06-P — Filter by text prefix](#05-p--06-p--filter-by-text-prefix)
- [07-P — Category dropdown drives a filter](#07-p--category-dropdown-drives-a-filter)
- [08-P — Filter by a code inside the text](#08-p--filter-by-a-code-inside-the-text)
- [09-P — Reformat codes](#09-p--reformat-codes)
- [10-P — Unique combinations with totals](#10-p--unique-combinations-with-totals)
- [11-P / 14-P — Filter by day of month](#11-p--14-p--filter-by-day-of-month)
- [12-P — "Contains" search](#12-p--contains-search)
- [13-P — Return only selected columns](#13-p--return-only-selected-columns)
- [15-P — Search token in receipt numbers](#15-p--search-token-in-receipt-numbers)
- [Quick reference](#quick-reference)

---

## Primer: how dynamic arrays work

Before Excel 365, a formula returned **one value** into **one cell**. A dynamic-array formula can return a whole array; Excel writes it into as many cells as needed (the **spill range**) starting at the cell holding the formula.

Three ideas appear in nearly every sheet:

1. **Boolean masks.** A comparison such as `'SYS-DATA'!F8:F1298 = B2` produces an array of `TRUE`/`FALSE`, one per row. `FILTER` keeps the rows where the mask is `TRUE`.
2. **Arithmetic on booleans.** `TRUE` behaves as `1` and `FALSE` as `0` in arithmetic. So `A + B` means **OR**, `A * B` means **AND**.
3. **Spill references.** `G6#` means "the whole spill range that starts at G6". Data-validation lists and downstream formulas can point at it and grow automatically.

```text
FILTER(array, include, [if_empty])
        │       │          └─ value to show when nothing matches
        │       └─ boolean mask, same height (or width) as array
        └─ the data to return
```

---

## 01-P — Compare two lists

**Sheets:** `01-P (1)`, `01-P (2)` (the second is an identical copy for practice).
**Goal:** For each voucher in the accounting list, is it present in the receipt list?

**Layout:** `C5:C33` = receipt-system vouchers, `E5:E33` = accounting vouchers.

### Method A — `VLOOKUP` (works in every Excel version)

```excel
G5:  =ISNA(VLOOKUP(E5,$C$5:$C$33,1,0))            ' TRUE  = NOT in the receipt list
H5:  =NOT(ISNA(VLOOKUP(E5,$C$5:$C$33,1,0)))       ' TRUE  = present in the receipt list
```

**How it works**

1. `VLOOKUP(E5, $C$5:$C$33, 1, 0)` searches for `E5` in the first (and only) column of the range. The final `0` forces an **exact** match.
2. If found, it returns the matched value; if not, it returns the error `#N/A`.
3. `ISNA(...)` converts that outcome to a boolean: `TRUE` when the lookup failed.
4. `NOT(...)` flips it for the "present" column.

The range is anchored with `$` so it stays fixed when the formula is copied down; `E5` is relative so it advances row by row.

### Method B — `XMATCH` (one formula, spills)

```excel
I5:  =ISNA(XMATCH(E5:E33,C5:C33))        ' spills I5:I33
J5:  =ISNUMBER(XMATCH(E5:E33,C5:C33))    ' spills J5:J33
```

**How it works**

- `XMATCH(lookup, array)` returns the **position** of `lookup` in `array`, or `#N/A`. Its default match mode is exact, so no extra argument is needed.
- Passing the whole range `E5:E33` as the lookup makes `XMATCH` return an *array* of positions — one per voucher.
- `ISNA` → `TRUE` where missing. `ISNUMBER` → `TRUE` where found (a position is a number; `#N/A` is not).

### Using the result

The boolean columns are designed to feed `FILTER`:

```excel
=FILTER(E5:E33, I5:I33)        ' accounting vouchers NOT in the receipt list
=FILTER(E5:E33, J5:J33)        ' accounting vouchers that ARE in the receipt list
```

**Notes**

- Method B needs no helper column per row and no copy-down; it is also the building block of 02-P.
- To find the opposite gap (receipts with no voucher), swap the two ranges.
- Column `M` in the sheet stores the formulas as text for reference.

---

## 02-P — Filter by a list of values

**Goal:** Return every full record whose key appears in a given list — "instead of one or two values".

```excel
C5:   =FILTER('SYS-DATA'!A8:H1298, ISNUMBER(XMATCH('SYS-DATA'!C8:C1298, B5:B17)))
C20:  =FILTER('DM-DATA'!A2:I1292,  ISNUMBER(XMATCH('DM-DATA'!B2:B1292,  B20:B33)))
```

**How it works** (first formula)

1. `XMATCH('SYS-DATA'!C8:C1298, B5:B17)` looks up **every** reference from the accounting data in the short list `B5:B17`. Result: an array of 1,291 positions or `#N/A`.
2. `ISNUMBER(...)` turns that into a boolean mask: `TRUE` for references that are in the list.
3. `FILTER('SYS-DATA'!A8:H1298, mask)` returns the matching rows, **all eight columns**, spilled into `C5:J17`.

**Notes**

- Direction matters: the *big* column is the lookup value, the *short list* is the lookup array. Reversing them answers a different question.
- Compare with the classic approach of one `FILTER` per key joined with `+` — it does not scale beyond a handful of values; the `XMATCH` mask handles any list length.
- Result order follows the **data**, not the list.

---

## 03-P — Unique lists with serial numbers

**Goal:** Produce de-duplicated lists of receipt numbers from each system, numbered.

```excel
B3:  =UNIQUE('DM-DATA'!B2:B1292)      ' spills B3:B1286
D3:  =UNIQUE('SYS-DATA'!C8:C1298)     ' spills D3:D1286
A3:  =ROWS($B$3:B3)                   ' copied down: 1, 2, 3, …
```

**How it works**

- `UNIQUE` returns each distinct value once, in order of first appearance.
- `ROWS($B$3:B3)` counts the rows in an **expanding range**: the start is anchored (`$B$3`), the end moves down with the formula. Row 3 → 1, row 4 → 2, and so on. It is a robust serial-number idiom that survives row deletion, unlike typing numbers.

**Notes**

- `UNIQUE(array, [by_col], [exactly_once])`: set `exactly_once` to `TRUE` to list only values that occur a single time.
- The serial formula is copied down a fixed distance. To make it dynamic, use `=SEQUENCE(ROWS(B3#))`.

---

## 04-P — Two-criteria lookup

**Goal:** Total the amount for each *(main type, sub type)* pair, then look up one pair chosen from dropdowns.

### Step 1 — unique pairs

```excel
B4:  =UNIQUE('DM-DATA'!C1:D1292)        ' spills B4:C10 (header + 6 pairs)
```

`UNIQUE` on a **two-column** range returns unique *rows* — combinations of columns C and D. Including row 1 in the range pulls the header labels into `B4:C4` automatically.

### Step 2 — total per pair

```excel
D5:  =SUMIFS('DM-DATA'!H:H,'DM-DATA'!C:C,'04-P '!B5,'DM-DATA'!D:D,'04-P '!C5)
```

`SUMIFS(sum_range, criteria_range1, criterion1, criteria_range2, criterion2)` adds `Amount` (H) where Main type (C) equals `B5` **and** Sub type (D) equals `C5`. The formula is copied down beside the spill.

### Step 3 — pick one pair (two methods)

Inputs: `G5`/`H5` and `G10`/`H10` are dropdowns fed from the unique list.

```excel
' New method
I5:   =FILTER(D5:D12, B5:B12&C5:C12 = G5&H5, "not found")

' Old method
I10:  =INDEX(D5:D12, MATCH(G10&H10, B5:B12&C5:C12, 0))
```

**How it works**

- `B5:B12&C5:C12` concatenates the two columns row by row into a single **composite key** (`"WajibaZakat"`). `G5&H5` builds the same key from the dropdowns.
- *New method:* `FILTER` returns the amount whose key matches; the third argument supplies a friendly message when nothing matches.
- *Old method:* `MATCH(key, keys, 0)` finds the row number; `INDEX` returns the amount from that row. It needs the array evaluation that dynamic-array Excel provides automatically; in older Excel it required Ctrl+Shift+Enter.

**Notes**

- **Delimiter risk.** `"ab"&"c"` and `"a"&"bc"` are the same key. If your values can collide, join with a separator: `B5:B12&"|"&C5:C12 = G5&"|"&H5`.
- **Boolean alternative** (no concatenation): `=FILTER(D5:D12,(B5:B12=G5)*(C5:C12=H5))`. `*` is AND.
- **Optional improvement:** the per-row `SUMIFS` must be extended manually if the number of pairs grows. A single spilling formula avoids that: `=SUMIFS('DM-DATA'!H:H,'DM-DATA'!C:C,DROP(B4#,1,-1),'DM-DATA'!D:D,DROP(B4#,1,1))` (placed in `D5`). Verify in your Excel build before adopting.

---

## 05-P / 06-P — Filter by text prefix

**Goal:** From all accounting vouchers, extract those that start with a given prefix (`ITD`, `DEMBB`, `CB`). `06-P` is a copy of `05-P` for practice.

**Layout:** Row 2 holds one sample voucher per prefix. Row 1 derives the prefix from it; row 3 is a self-check; row 5 holds the spilling formulas.

```excel
' Derive the prefix (row 1) — length differs per prefix
C1:  =LEFT(C2,3)      ' "ITD"
D1:  =LEFT(D2,5)      ' "DEMBB"
E1:  =LEFT(E2,2)      ' "CB"

' Self-check (row 3) — must be TRUE
D3:  =LEFT(D2,5)=D1

' Extract (row 5)
B5:  =UNIQUE('SYS-DATA'!A8:A1298)                                               ' all vouchers, de-duplicated
D5:  =FILTER('SYS-DATA'!A8:A1298, LEFT('SYS-DATA'!A8:A1298,5)=D1)               ' every DEMBB voucher
G5:  =UNIQUE(FILTER('SYS-DATA'!A8:A1298, LEFT('SYS-DATA'!A8:A1298,5)=G1))       ' …de-duplicated
```

**How it works**

1. `LEFT(range, 5)` applied to a **range** returns an array — the first five characters of every voucher.
2. Comparing that array to `D1` (`"DEMBB"`) produces the boolean mask.
3. `FILTER` returns vouchers where the mask is `TRUE`; wrapping in `UNIQUE` removes repeats.

**Notes**

- The number in `LEFT(…, n)` must equal the prefix length (3, 5 or 2 here). Hard-coding it is the most common source of silent wrong results when a prefix changes.
- **Optional improvement:** derive the length from the prefix itself so it never goes out of sync:
  `=FILTER('SYS-DATA'!A8:A1298, LEFT('SYS-DATA'!A8:A1298, LEN(D1)) = D1)`
- `LEFT`-based comparison is **case-insensitive**.

---

## 07-P — Category dropdown drives a filter

**Goal:** Pick a category (`P-3`) from a dropdown and list all matching vouchers.

```excel
A2:  =UNIQUE('SYS-DATA'!F7:F1298)                               ' spills A2:A4: header, "For All", "Masjid"
B2:  (dropdown, source: $A$3:$A$9)                              ' the chosen category
C5:  =FILTER('SYS-DATA'!A7:H1298, 'SYS-DATA'!F7:F1298 = B2)     ' spills C5:J1292
```

**How it works**

- `UNIQUE` builds the list of categories that actually exist; a data-validation list on `B2` offers them.
- `FILTER` keeps rows whose `P-3` value equals the selection. Headers in `C4:J4` are typed once; the formula returns data rows only.

**Notes**

- The unique list is a spill, so new categories in the data appear in the dropdown source automatically — as long as the validation range (`$A$3:$A$9`) is wide enough. Pointing validation at the spill reference (`=$A$3#`) removes that limit.
- Selecting nothing makes the comparison match blank cells, not everything. Wrap in `IF(B2="", …)` if you want a prompt instead.

---

## 08-P — Filter by a code inside the text

**Goal:** Extract receipts by *channel* — the segment right after the first hyphen (`ER`, `MB`, `DM`).

Example receipt numbers: `DEMO8-ER-23-555`, `DEMO8-MB-ER23-877`, `DEMO8-DM-23-319`.

```excel
B5:  =FILTER('DM-DATA'!$B$2:$B$1292,
             MID('DM-DATA'!B2:B1292, SEARCH("-",'DM-DATA'!B2:B1292)+1, LEN(B2)) = B2)

F5:  =UNIQUE( same FILTER, using F2 )
```

`B2` holds the code (`ER`); `C2` → `MB`; `D2` → `DM`.

**How it works**

1. `SEARCH("-", receipts)` returns the position of the first hyphen in each receipt number (an array).
2. `+1` moves to the first character *after* it.
3. `MID(receipts, start, LEN(B2))` cuts out exactly as many characters as the code has — 2 for `ER`.
4. The cut-out segment is compared with `B2`; `FILTER` keeps the matches.

**Notes**

- Because the extraction length comes from `LEN(B2)`, the same formula works for codes of any length — a better pattern than the hard-coded `LEFT` used in 05-P.
- If a receipt has no hyphen, `SEARCH` returns `#VALUE!` and the whole `FILTER` fails. Guard with `IFERROR(SEARCH(…),0)` if your data is not clean.

---

## 09-P — Reformat codes

**Goal:** Turn `ABCDIT0000001` into `ABC-DIT-000-0001`.

```excel
' Classic: slice and join with &
D4:  =LEFT(B4,3)&"-"&MID(B4,4,3)&"-"&MID(B4,7,3)&"-"&RIGHT(B4,4)

' Modern: slice with array constants, join once
E4:  =TEXTJOIN("-",TRUE,MID(B4,{1,4,7,10},{3,3,3,4}))

' Show a formula as text
D1:  =FORMULATEXT(D4)
D2:  =FORMULATEXT(E4)
```

**How it works**

- Classic: four slices (3 + 3 + 3 + 4 = 13 characters) glued with hyphens.
- Modern: `MID` is given two **array constants**: start positions `{1,4,7,10}` and lengths `{3,3,3,4}`. It returns four pieces at once; `TEXTJOIN("-", TRUE, …)` joins them. `TRUE` tells `TEXTJOIN` to skip empty pieces.
- `FORMULATEXT(ref)` returns the formula in `ref` as a string — handy for documentation sheets like this one.

**Notes**

- The array-constant version is easier to change: edit two lists instead of rewriting a chain of `&`.
- `FORMULATEXT` returns `#N/A` if the referenced cell has no formula.

---

## 10-P — Unique combinations with totals

**Goal:** List each unique *(main type, sub type, purpose)* combination and its total amount.

**Layout:** Source columns `B:D` (rows 1–1316) hold the combinations as they appear in an extract.

```excel
F2:  =UNIQUE(B1:D1316)                                                              ' spills F2:H15
I3:  =SUMIFS('DM-DATA'!H:H,'DM-DATA'!C:C,F3,'DM-DATA'!D:D,G3,'DM-DATA'!E:E,H3)       ' copied down
I18: =SUM(I3:I17)                                                                    ' grand total
```

**How it works**

- `UNIQUE` on three columns returns unique **rows**.
- `SUMIFS` with three criteria pairs adds `Amount` only where **all three** match — each additional pair is an extra AND.

**Notes**

- A combination that exists in the helper columns but not in `DM-DATA` returns `0`; that is expected, not an error.
- The total range `I3:I17` is wider than the current list so that it keeps working if a few combinations are added.
- Same limitation as 04-P: the `SUMIFS` column is copied manually beside a spilling list.

---

## 11-P / 14-P — Filter by day of month

**Goal:** Show every receipt dated on a given day (11-P), or on either of two days (14-P), with a running total.

```excel
' 11-P: one day
A6:  =FILTER('DM-DATA'!A1:I1292, TEXT('DM-DATA'!A1:A1292,"dd") = '11-P '!A3)
B3:  =SUBTOTAL(9,H:H)

' 14-P: either of two days
A6:  =FILTER('DM-DATA'!A1:I1292,
             (TEXT('DM-DATA'!A1:A1292,"dd") = '14-P'!A3) + (TEXT('DM-DATA'!A1:A1292,"dd") = '14-P'!B3))
C3:  =SUBTOTAL(9,H:H)
```

Inputs are typed as **text**: `"20"` in 11-P; `"10"` and `"30"` in 14-P.

**How it works**

- `TEXT(date, "dd")` formats each date as its two-digit day: `"05"`, `"20"`. The result is *text*.
- The comparison is therefore text-to-text, which is why the inputs must be text too (see the warning below).
- 14-P adds two masks with `+`, which is **OR**: a row passes if it matches day A *or* day B.
- `SUBTOTAL(9, H:H)` sums the `Amount` column (H) of the spilled result. Function number `9` is SUM; `SUBTOTAL` also respects any AutoFilter applied to the sheet.

> **Warning — text vs number.** In Excel, `"20" = 20` is `FALSE`. If the day cell is stored as the number `20` (or `5` instead of `"05"`), the filter returns nothing. Format the input cells as Text, or compare numerically with `DAY('DM-DATA'!A1:A1292) = A3`, which removes the problem.

**Notes**

- The range starts at row 1, so the header row is part of the array; its text never equals a day, so it is excluded automatically.
- `DAY()` works across all dates in the month, but `TEXT(…,"dd")` is language-sensitive in some locales; `DAY()` is safer for shared files.

---

## 12-P — "Contains" search

**Goal:** Return every accounting record whose location **contains** the typed text.

```excel
A6:  =IFERROR(FILTER('SYS-DATA'!A8:H1298, ISNUMBER(SEARCH(A3,'SYS-DATA'!D8:D1298))), "No record found")
C3:  =FORMULATEXT(A6)
```

`A3` is a dropdown of locations, but it also accepts any partial text you type.

**How it works**

1. `SEARCH(A3, locations)` looks for `A3` *inside* each location. It returns the position (a number) if found and `#VALUE!` if not.
2. `ISNUMBER` converts that to `TRUE`/`FALSE`.
3. `FILTER` keeps the `TRUE` rows. If none match, `FILTER` returns `#CALC!`, which `IFERROR` replaces with a message.

**Notes**

- `SEARCH` is **case-insensitive** and supports wildcards: `?` (one character), `*` (any run), and `~` to escape them. Use `FIND` for case-sensitive matching (no wildcards).
- **An empty `A3` matches everything** (`SEARCH("", text)` returns 1). Guard with `IF(A3="","",…)` if that is undesirable.
- Prefer `FILTER`'s own third argument (`if_empty`) for the "no match" message when the only error you expect is an empty result — `IFERROR` also hides real mistakes.

---

## 13-P — Return only selected columns

**Goal:** Filter rows by account name, but return only some columns (account and amount).

**Layout:** Data `B5:E1295` — Voucher, Date, Account (`C-3`), Amount. `F3` selects the account (dropdown fed from `G6#`).

```excel
G6:  =UNIQUE(D6:D1295)                                                         ' list of accounts (dropdown source)

' Method 1 — filter rows, then filter columns
H5:  =FILTER(FILTER(B6:E1295, D6:D1295 = F3), {0,0,1,1})

' Method 2 — build a two-column array, then filter rows
K5:  =FILTER(CHOOSE({1,2}, D6:D1295, E6:E1295), D6:D1295 = F3)
```

**How it works**

- *Method 1.* The inner `FILTER` returns the matching rows (4 columns). The outer `FILTER` is given a **row** vector `{0,0,1,1}` as its include argument. When the include array is horizontal, `FILTER` filters **columns** — keeping the 3rd and 4th.
- *Method 2.* `CHOOSE({1,2}, D, E)` assembles a new two-column array from columns D and E. `FILTER` then filters its rows with the same mask.

**Notes**

- Method 2 reads better when the columns you want are not adjacent or must be reordered.
- Both return the account name *and* the amount; to return only the amount, use `FILTER(E6:E1295, D6:D1295 = F3)` directly.
- Newer Excel has `CHOOSECOLS(array, 3, 4)`, which replaces both tricks with one readable call.

---

## 15-P — Search token in receipt numbers

**Goal:** Show receipts whose number contains a chosen token (`ER-`, `DM-`, `MB-`).

```excel
A4:  =FILTER('DM-DATA'!A2:I1292, ISNUMBER(SEARCH(B2, 'DM-DATA'!B2:B1292)))
```

`B2` is a dropdown fed from `F1:H1` (the three tokens).

**How it works**

Identical to 12-P: `SEARCH` → position or error → `ISNUMBER` → boolean mask → `FILTER`.

**Notes**

- Compare with 08-P. `SEARCH` finds the token **anywhere** in the string, while the `MID`/`SEARCH("-")` pattern in 08-P inspects only the segment *after the first hyphen*. Use 08-P when the code's position is fixed and false positives matter; use this pattern when it is not.
- The trailing hyphen in the token (`DM-`) reduces false positives, for example `DM` inside another word.

---

## Quick reference

| I want to… | Pattern |
|---|---|
| Keep rows where a column equals a value | `=FILTER(data, col = value)` |
| …equals **A and** B | `=FILTER(data, (col1 = A) * (col2 = B))` |
| …equals **A or** B | `=FILTER(data, (col1 = A) + (col1 = B))` |
| …is in a list | `=FILTER(data, ISNUMBER(XMATCH(col, list)))` |
| …contains text | `=FILTER(data, ISNUMBER(SEARCH(text, col)))` |
| …starts with text | `=FILTER(data, LEFT(col, LEN(text)) = text)` |
| Show a friendly message when empty | `=FILTER(data, mask, "No record found")` |
| Remove duplicates | `=UNIQUE(range)` |
| Return only some columns | `=CHOOSECOLS(FILTER(data, mask), 3, 4)` |
| Sum with two conditions | `=SUMIFS(sum, rng1, c1, rng2, c2)` |
| Serial numbers that survive deletes | `=ROWS($B$3:B3)` |
| Show a formula as text | `=FORMULATEXT(cell)` |
