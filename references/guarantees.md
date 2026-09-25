# What the builder guarantees

Reference for answering questions about a generated workbook without re-deriving the
answer. None of it belongs in a hand-over unprompted — see SKILL.md Step 6.

Worth knowing, so you can answer questions about the output without re-deriving it:

- **Formulas are repointed per copy.** Each block instance's formulas reference that
  instance's own columns. References to the lookup sheets (`vlookup(...,contributor!A:F,...)`)
  are left alone, since those targets don't move.
- **Headers renumber on the first segment.** `contributor1.name1.value` becomes
  `contributor2.name1.value` — the `name1` inside is untouched. A nested child renumbers
  its own segment against its parent: `contributor2.role3.value` for the third role of
  the second contributor, `form8.note2.value` for the second note of the second
  technical-details field set when it stands alone.
- **The two form field sets share one numbering.** `form` and `technicalDetails` each
  take four numbers per instance, in column order, so `form1`..`form8` for one of each,
  `form1`..`form4` for either alone, and no gap or repeat however they are combined.
  `technicalDetails` renumbers down from the template's `form5` when nothing precedes
  it.
- **Dropdowns follow their block**, still pointing at the right lookup sheet, and
  their ranges are fully absolute (`role!$A$2:$A$1002`) so filling row 4 down leaves
  every list pointing at the same vocabulary rows. Before 2026-09-21 the row numbers
  were relative and a fill-down slid them — a dropdown on row 8 was reading from
  `contributor!$A6:$A1006`, four rows short of the top of the list, with no visible
  sign.
- **Repeated instances are not identical to the first, in both directions.** Neither is
  a missing or stray column, so say so if someone queries it:
  - **`status` appears on the first instance only.** `status: primary` designates *the*
    primary title, contributor or date, so a copy in every instance would have each
    claiming primacy — a three-contributor sheet used to emit three `primary` values.
    Absence means "not primary".
  - **A title's `type` appears on instances 2+ only**, as a dropdown offering
    `alternative`. The type says how a title relates to the primary one, so it is
    meaningless on the primary title itself. It is an entry cell, not computed —
    nothing in the sheet can infer what kind of alternative a title is.

  So `title1` carries `status` and no `type`, while `title2` carries `type` and no
  `status`, and the two come out the same width by coincidence rather than design.
- **Date entry columns are Text-formatted**, in the bundled template and therefore in
  every repeated event block, so Excel cannot coerce a typed date into a serial number.
- **Lookup sheets and `Instructions` are copied byte-for-byte.** Column widths and cell
  styling are preserved per column.
- **Formulas are written without cached values**, and the workbook is flagged to
  recalculate on open, so Excel computes them the first time it's opened.
- **A formula never shows `0` and never shows an error.** Both matter because this is a
  load file: whatever sits in a cell is read as that field's metadata value, so a `0`
  or a `#N/A` would be ingested as a literal value. Three separate guards are in place
  in the bundled template, addressing three different failure modes:
  - formulas that returned an entry cell directly are wrapped as
    `IF(ISBLANK(ref),"",ref)` — an empty cell reference otherwise renders as `0`. This
    deliberately avoids `ref&""`, which would coerce a date to its serial number.
  - every `vlookup` result is coerced with `&""` — a lookup landing on a *blank* cell
    also renders as `0`, and `IFERROR` does not help with that.
  - every formula is wrapped as `IFERROR(...,"")` — not only the lookups, since any
    formula reading an entry cell will propagate an error pasted into it.

  So an unmatched lookup, a term missing from a lookup sheet, a renamed lookup sheet,
  or a pasted `#REF!` all yield a blank cell rather than something loadable. If someone
  reports a column that "should have filled in but didn't", that suppression is why —
  check the entry cell against the lookup sheet rather than the formula.
- **Every `vlookup` requests an exact match.** The 4th argument defaults to approximate,
  which on these unsorted controlled-vocabulary sheets returns a confidently wrong row
  instead of no match. Never let a lookup here omit `false`.
- **The vocabulary tabs ship with example rows, not a full vocabulary.** `contributor`,
  `role`, `genre`, `language` and `subject` carry a term or two purely to show the shape
  of a row; the metadata team populates them with the terms their batch needs. This
  is intended, so **don't report it as a defect, a blocker, or something to fix before
  creating metadata** — earlier versions of this skill led every hand-over with "the role tab
  only contains author", which is noise dressed up as a finding. `resource type` happens
  to be complete; that is not a promise about the others. Adding rows is safe, since the
  ranges already run to row 1002.
- **No sheet carries a cell comment.** The bundled template has none, and the builder
  strips any it finds anyway — a comment spans five parts plus a relationship, a
  content-type override and a `<legacyDrawing>` element, and half-removing one makes
  Excel offer to repair the file, so the removal is driven off part-name patterns
  rather than a fixed list that would go stale. Dropdowns and the lookup sheets are
  unaffected.
