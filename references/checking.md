# Writing a checker for a generated workbook

Read this before writing a script that inspects the output. Every trap below has
actually caught a checker on this workbook, and each one reports a clean pass or a
false defect rather than an error — which is why they are worth knowing in advance
rather than discovering.

## Traps

- **Header text is inline, not shared.** The generated sheet writes `t="inlineStr"`
  with the text under `<is><t>`, while the template uses `xl/sharedStrings.xml`. A
  checker written for the shared-strings layout finds zero headers — and a zero-header
  run looks exactly like "no duplicates found". The generated file may have no
  `sharedStrings.xml` at all.
- **`metadata` is `xl/worksheets/sheet2.xml`, not `sheet1.xml`.** `sheet1` is
  `Instructions`. Resolve the sheet through `xl/workbook.xml` and its `.rels`, or find
  the sheet whose row 3 contains `druid`. Don't index by position.
- **Parse the XML; don't regex it.** Two shapes break naive patterns. In `<sheet>` the
  attribute order is `state="visible" name="metadata"`, so `<sheet name="...">` matches
  nothing. And empty cells are self-closing (`<c r="AC4" s="5"/>`), so a pattern like
  `<c\b[^>]*>.*?</c>` runs straight past them and attributes the *next* cell's formula
  to the empty one, which surfaces as a phantom cross-block reference.
- **Resolve columns by header text, never by remembered letter.** Column letters move
  with the spec: the same date entry column has landed on `AW`, `AM` and `AY` in
  different builds. A hardcoded letter silently checks the wrong column.
- **The builder prints a column count, not a header count, and the two differ.** Three
  field sets put the typed cell in a column with no row-3 header of its own, one per
  instance: `title`'s "Main title", `event`'s "Date (YYYY-MM-DD)", and
  `multipartSubject`'s "part 1 value". So `headers = columns - instances of those
  three`: a title plus two events is 3 short, three titles is 3 short, and a title plus
  two simple `subject` sets is 1 short, because the simple set's own value column does
  carry its header. Deriving the expected header count from the printed column count
  produces an off-by-a-few that reads as missing headers. `blocks.md` marks these
  columns with `—` in the header column.
- **Assert your checker found what it meant to inspect** — `len(headers) > 50` or
  similar. It is the one guard that turns every trap above from a green run into a
  visible failure. It only catches a total miss, though: a checker that found 85 of 87
  headers still looks fine, so compare against `blocks.md` when the count matters.

## False positives

Expect these, and don't report them as defects.

- **Twelve formulas legitimately reference the Main title entry cell** (`F4` in an
  unshifted build): the nine `adminMetadata` ones, which are the designed cross-block
  reference, plus three inside the title block itself —
  `title1.structuredValue2.value`, `title1.value` and `title1.status` — which read the
  cell their own block owns. A "reference too far from its own block" heuristic flags
  the first nine; exempting `adminMetadata` alone still leaves the other three looking
  wrong.
- **A cross-block check must compare each formula against the block that owns the
  referencing column**, not against a list of exemptions. The three title formulas
  above are violations only if you forget that `F` is inside the title block.
- **Row 1 is not a block-boundary source.** Every computed column carries its own
  `DO NOT EDIT - FILL DOWN` label, so deriving block ranges from row-1 labels treats
  each column as its own block — one checker produced 133 phantom cross-block
  references that way. Take the ranges from the builder's own printed block list.
- **Inline-list dropdowns are not ranges.** `"publication,creation"` has no sheet
  reference, so an "is the range absolute" check must skip any `formula1` without a
  `!`.
- **In Excel COM, `Validation.Type` does not throw on a cell with no validation** — it
  returns `xlValidateInputOnly`, so every column looks like a dropdown. Count
  validations from the `dataValidations` XML instead. Relatedly,
  `SpecialCells(xlCellTypeAllValidation)` raises "No cells were found" on a row with
  none, and `xlCellTypeAllFormatConditions` is a different constant — neither is a
  finding about the workbook.
