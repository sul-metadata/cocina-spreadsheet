---
name: cocina-spreadsheet
description: Build a Cocina metadata spreadsheet for SDR or DOR, sized to the material being described — repeating or omitting blocks of header columns such as titles, contributors, forms, events, subjects or geographic data, with every formula, dropdown and lookup sheet carried across intact. Use for requests like "a spreadsheet with 3 contributors and 2 subjects", "a Cocina template with the formulas", or "drop the geographic columns" — anywhere a working workbook to enter metadata into is wanted, rather than a flat list of header names.
---

# Cocina metadata spreadsheet

The bundled template at `assets/cocina_spreadsheet.xlsx` is a working metadata
workbook, not just a header row. Its `metadata` sheet has three header rows (block
label, entry label, Cocina header), a row of 79 formulas that derive types, URIs and
authority codes from what the metadata creator types, dropdowns on the entry cells, and six
lookup sheets the formulas `vlookup` into.

Metadata creators need different shapes of it: three contributors instead of one, two
subjects, no geographic data. Doing that by hand is where it goes wrong — duplicating
a block means renumbering its headers, repointing every formula in the copy at the
copy's own columns rather than the original's, and carrying the dropdowns along.
`scripts/build_spreadsheet.py` does all of that; your job is to get the block counts
right and to check the result.

## Workflow

### Step 0 — Confirm there is a Python interpreter

Do this first, before asking anything about blocks. It takes one command, and if no
interpreter exists you want to say so straight away rather than after the user has
worked through their block counts with you.

Try these in order and stop at the first that prints a version:

```
python --version
python3 --version
wsl -d <distro> -- python3 --version
```

**Read the output, not the exit status.** On Windows `python` is often a 0-byte
Microsoft Store stub: the command exists and exits cleanly to a status check, but
prints "Python was not found". Treat that as no Python.

Nothing else needs installing — the builder is standard library only. If no
interpreter is found, say so before asking about blocks, and see
`references/python-setup.md` for what to offer them per platform.

### Step 1 — Find out what shape they need

Ask for the blocks they want and how many of each. Blocks they don't mention are left
out of the workbook entirely, so it's worth confirming omissions rather than assuming
them.

**Never add a block they didn't ask for — suggest it instead.** If a block looks like
something the batch will want, say so and let them decide, rather than putting it in and
noting it afterwards. The reasoning that makes adding it tempting — an unused block
loads as nothing, while a missing one means a rebuild — is real, and it is an argument
for *raising* the block in Step 2, not for widening the spec yourself. What they asked
for is the spec; a column they never agreed to is a column someone has to be told to
ignore.

**Suggest, don't insist, and don't invent a requirement.** `form` is the usual candidate
— it holds resource type, genre, extent and media type, which many batches want — so
naming it when it's missing is useful. But it is not required: say "worth considering",
not "the load needs it". The only requirement to state is `title`, which the builder
itself rejects a spec without.

Two blocks are not the caller's choice. `title` is required — every object needs one,
so ask how many titles rather than whether they want any, and the builder rejects a
spec without it. `adminMetadata` is added automatically and placed last whether or not
anyone mentions it, because it records which agency created the metadata and every
SDR record carries that. Don't ask about it, and don't list it as an omission; just
mention it's there if the block list you show would otherwise look incomplete.

Be careful what "required" means here — it is the *block* that is required, not the
title text. You are building an empty workbook for someone to fill in later, so never
ask for actual title wording at this stage; the metadata creator types that into the sheet in
Excel. The same goes for every other block: counts are the only thing being decided.

Valid block types, and how many times each may appear:

People ask in descriptive terms, not block names, so the middle column is what you
match their words against — "repository" and "shelf locator" both mean `access`, and a
request for "genre" or "extent" means `form`:

| type | what it holds | repeats |
|---|---|---|
| `title` | main title, subtitle, part number/name, nonsorting article | yes — **required**, always ask how many |
| `contributor` | name, type, name and identifier URIs, authority codes; carries one or more nested `role` blocks | yes |
| `form` | resource type, form, extent, genre, reformatting quality, digital origin, media type; carries one or more nested `note` blocks | yes |
| `event` | labelled "Origin info": date, end-date-if-range, approximate flag, place, **publisher** | yes |
| `language` | language, code, URI, authority code | yes |
| `note` | note, type, display label | yes |
| `identifier` | identifier, type, display label | yes |
| `subject` | one heading per subject, with type, URI and authority | yes |
| `multipartSubject` | a two-part structured heading, with type, URI and authority for the whole and for each part | yes |
| `relatedResource` | type, display label, title, PURL, other URL, abstract | yes |
| `geographic` | MIME type, DC resource type, point coordinates, bounding box | yes |
| `access` | **repository/library** (with URI and authority code) and **shelf locator** | **once only** — its headers carry no instance number, so a second copy would collide |
| `adminMetadata` | cataloguing agency and language of cataloguing | **automatic** — added for you, always last |

Two blocks nest a repeatable child, given as a list with one number per parent
instance:

- **Roles inside contributors.** "Two contributors, the first with three roles" is
  `contributor` count 2 with `"roles": [3, 1]`.
- **Form notes inside forms.** "Two forms, two notes on the second" is `form` count 2
  with `"notes": [1, 2]`.

If they don't mention roles or form notes, give each parent one. Note the two senses of
"note": a form note is the `notes` list inside a `form` block, while a general note on
the object is the standalone `note` block. Ask which they mean if it isn't clear from
context — "a note about the file format" is a form note, "a general note" is the `note`
block.

**The two subject field sets are alternatives, and a workbook carries only one.**

- Plain **"subject"**, "subjects", "simple subject", "single subject", and also
  **"keywords"** or **"topics"** — what people call them when not thinking in Cocina
  terms → `subject`: four columns per instance, one heading typed into the value column
  with type, URI and authority derived from it.
- **"multipart subject"**, "complex subject" or "structured subject" →
  `multipartSubject`: thirteen columns per instance, a heading in two parts, each part
  with its own type, URI and authority.

They cannot both appear. Both number their headers from `subject1`, and the simple set's
four headers are the same strings as four inside the multipart set, so a sheet with both
would carry `subject1.value` twice. The builder refuses a spec naming both rather than
picking one.

**If someone asks for both, ask which they want — don't choose for them.** A batch
described as "subjects, some of them multipart" is one question, not a judgement call:
offer `subject` and `multipartSubject` as the two options and say what each gives them.
Someone who asks only for "subjects" gets the simple set without being asked.

**That question is a fifth question, asked in its own `AskUserQuestion` call before
the four in Step 2.** The tool accepts at most four questions per call, so it cannot
join them; it goes first because the answer changes the field-set list those four
confirm. Head it `Subjects` and ask:

> Should the subjects be single or multipart?

with the two options `Single subjects` and `Multipart subjects`, each described by what
it gives them — "One heading per subject, with type, URI and authority derived from it;
4 columns each" and "A two-part heading per subject, with type, URI and authority for
the whole and for each part; 13 columns each". Keep the count they asked for in both.
Once they answer, show the field-set list with the chosen kind and make the Step 2 call
as usual. Ask this only when the request implies both kinds; it is the one exception to
Step 2's single round trip.

`references/blocks.md` lists every header, entry label and formula in each block. Read
it when someone asks what a block covers, or to check whether a field they want is
already inside a block they've asked for — `form` in particular bundles eight distinct
form entries, so people asking for "genre" and "extent" separately need one `form`
block, not two.

### Step 2 — Confirm before building

Show the parsed request back as a list and get agreement, because a misread count
means a workbook that has to be thrown away rather than edited:

**Four things must be settled before you build**: the block list, the folder, the
filename, and whether they have an object list to pre-fill from. None is optional, and
none is worth its own round trip.

**Ask them as clickable options, not as prose to reply to.** Use `AskUserQuestion` with
all of them in a single call, so answering is a few clicks rather than a typed message.
Every question the tool shows also carries an "Other" choice for free text, which is
how someone names a folder or a filename you didn't offer — so offering two options
costs them nothing and saves everyone else the typing.

**Call them field sets, not blocks, in anything the user reads.** `block` is the
builder's word: the spec key, `--list-blocks`, `references/blocks.md` and the column
ranges in this file all keep it. What the user sees says *field set* — "Build these
field sets", "Field sets to include", "2 access field sets". The term is also the more
accurate one for them: each is a group of related fields, which is what they get.

Show the parsed list in the message *before* the `AskUserQuestion` call, as plain text.
Never put it inside a question's text or an option label: it runs to a dozen lines, so
a question carrying it renders as a wall of text above two buttons, and it reappears
every time they look at that question. Then ask:

```
Field sets to include:
 - 1 title
 - 2 contributors (3 roles on the first, 1 on the second)
 - 2 subjects
 - access information
 - administrative metadata (always included)
Leaving out: form/genre, origin info, language, note, identifier, related resource, geographic
```

| question | options |
|---|---|
| Build these field sets, or choose Other to describe changes? | `Build as listed` · one concrete amendment, e.g. `Add origin info` |
| Which folder should it go in? | `Desktop` · `Documents` |
| What should the file be called? | `cocina_spreadsheet.xlsx` · `project_workbook.xlsx` |
| Pre-fill the rows from a list of objects? | `Build it empty` · `I'll upload a list` |

The `header` chip on each question is user-visible as well, so it follows the same
vocabulary: `Field sets`, `Folder`, `Filename`, `Prefill`.

**Labels stay short, but never leave them to speak for themselves.** They are buttons;
the person reading them may not know this template, so each one needs a description
saying *what happens if they pick it*, in concrete terms:

- `Desktop` → "C:\Users\arcadia\Desktop — the file will be sitting there when you
  switch windows."
- `cocina_spreadsheet.xlsx` → "The workbook is saved under this name, in whichever
  folder you pick above."
- `Build it empty` → "Formulas on row 4, ready to fill down. You type the druids and
  titles yourself."
- `I'll upload a list` → "Upload a .csv, .tsv or .xlsx with druid, source ID and title
  in the first three columns. Each object gets a row with its formulas already filled
  down, so no identifier is typed by hand, plus one empty row below them to fill down
  from if you add more objects."
- `Build as listed` → restate the shape in one line — "1 title, 2 contributors with one
  role each, 3 subjects, access information, admin metadata; about 110 columns."

**Describe the consequence, don't recommend the choice.** "Formulas on row 4, ready to
fill down" is context; "probably what you want" is advice. The distinction matters most
on the filename, where there is nothing to recommend — but don't fill that gap with a
full path. All four questions are shown at once, so the folder is still unanswered
while they are reading the filename: a description promising
`C:\Users\arcadia\Documents\cocina_spreadsheet.xlsx` names a folder they may never
pick, and in testing it named one they hadn't. Describe the name; let the folder
question own the folder.

Put the same care into the question text. "Which folder?" assumes they know why they're
being asked; "Which folder should it go in?" with the paths in the descriptions tells
them what the answer decides. If they pick the amendment option or `I'll upload a list`,
the next message is theirs: the correction, or the uploaded file.

**Folder and filename are separate questions.** Each has two real options, so neither
needs padding: `Desktop` or `Documents` for the folder, `cocina_spreadsheet.xlsx` or
`project_workbook.xlsx` for the name. Keeping them apart means someone changing only
the name doesn't have to restate the folder, and vice versa; "Other" on either one
takes anything else.

**Every option must be a complete answer they can click.** Never write an option that
stands in for typing — "Something else (I'll type it)", "Name it myself", "I want
changes". Selecting one records that label and moves straight on, so they get no prompt
and you get no answer; it reads as a text field and isn't one. The tool already adds an
**"Other"** choice to every question, and that is the only thing that collects free
text.

This bites hardest on the block question, where the alternative to "yes" is an
open-ended correction. Don't render that as `I want changes`. Offer **`Build as listed`
plus the one concrete amendment you would otherwise have recommended** — `Add origin
info`, `Add a contributor` — so both options are directly actionable, and put the
free-text route in the *question text* instead: "Build these field sets, or choose
Other to describe changes?". Naming Other in the question is not the duplicate the
rule above forbids; it is a pointer to the real control rather than a decoy beside it.

If no concrete amendment is worth offering, the second option can be the block you most
expect them to drop — anything real. What it cannot be is a label that means "let me
type".

Then wait for the answers. Building first and asking afterwards wastes the build, since
a rebuild is a new file rather than an edit.

If `AskUserQuestion` isn't available, ask the same things in one plain-text message
with the defaults named — "Documents or Desktop, or name another folder; default
`cocina_spreadsheet.xlsx`". Falling back to prose is fine; skipping the questions is
not.

A block you think they need goes in that list as a question, not in the spec — "you
haven't mentioned form/genre, which is where resource type lives; add it?" — on its own
line after `Leaving out`, so the omission and the suggestion are visible together.

If you can't get an answer before building — a one-shot request, or a context where
asking isn't possible — build exactly what was asked and put the recommendation in the
reply alongside the file. Not being able to ask is a reason to be more conservative
about the spec, not less: an unrequested block in a delivered workbook is a decision
they never got to make, whereas a recommendation they can act on costs them one rebuild
of an empty sheet.

The same applies to every one of those questions. Unanswerable is not the same as
skippable: take the default filename and folder, build empty, and carry the offers into
the hand-over. The cost of asking after the fact is one rebuild of an empty sheet; the
cost of not asking is someone typing 40 druids by hand.

**Carry them as one sentence, not a paragraph each.** Deferred questions are the only
part of the hand-over that scales with how little you were told, so left loose they
crowd out what the reply is actually for — a one-shot run has hit 239 words with the
file itself described in three lines. Compress the lot into a single closing sentence:

```
I chose the filename and folder; say the word and I'll rebuild with your own, or
with the rows pre-filled from a list of druids, source IDs and titles.
```

One sentence covers all three. If a deferred question needs more room than that, it is
a caveat about the file rather than an unasked question, and belongs with the other
caveats. The 200-word ceiling is not waived for the can't-ask path — it is the case
that most needs it.

**The filename and the folder are both theirs to choose.** They are the ones who have to
find the file afterwards, and someone who can't locate the file they just asked for has
been handed nothing.

- **Offer `Desktop` and `Documents` as the two choices rather than asking for a path.**
  Two familiar places, each one click, both trivial to find afterwards. Typing a path is
  work — open a file manager, copy the location, get the separators right — and it is
  work most people don't need to do to answer this question. The question tool's "Other"
  choice already covers anyone who wants somewhere specific, so there is no need to
  invite it in the label.
- `Documents` is the default if they don't pick. Resolve either name to the real path
  under their profile (`C:\Users\<them>\Documents`, `C:\Users\<them>\Desktop`) and show
  the full path in the `File at:` line at the end, so they see where it actually went.
- If they name a folder, use it verbatim. Create it if it doesn't exist, and say so.
- **Don't choose a temp or scratch directory, and never the skill directory.** A path
  under `AppData\Local\Temp` or a session scratchpad is not somewhere a person keeps a
  work file; it may also be cleaned up under them. The skill directory is for the skill.

  **If they explicitly name such a path, use it anyway** and add one line to the
  hand-over: "that folder is under `AppData\Local\Temp`, which Windows may clear out, so
  move the file somewhere permanent before creating metadata." An explicit instruction outranks
  this default — silently relocating their file is worse than saving it where they said,
  because they will go looking where they asked for it. The rule above governs what you
  choose when they haven't; it is not a veto over what they ask for.
- **Filename: offer `cocina_spreadsheet.xlsx` and `project_workbook.xlsx`, with no
  advice about which to pick.** Naming is theirs. Don't explain when each one fits,
  don't suggest a name from the batch they described, and don't comment on their choice
  once they've made it — someone who ends up with several of these may want
  `sierra_club_photos.xlsx`, and "Other" is where that goes. Each option still needs a
  description, but the useful one is the resulting full path, not a reason to prefer it.
  If they give a name without `.xlsx`, or with a different extension, replace it with
  `.xlsx`.
- Write Windows paths with backslashes here, the same form Step 6 shows in the `File at:`
  line, so the path you confirm and the path they end up with read identically.

If asking isn't possible, save to the working directory rather than a temp path, and
name the folder in the reply so they can move it — **unless the working directory is
unsuitable**, which is common enough to check. A `\\wsl.localhost\...` UNC path is one
Step 6 forbids handing back, and a source repository is not where a metadata work file
belongs. In either case use `Documents` instead and say that is what you did.

One thing to raise at this point if it applies, since it isn't obvious from the output:
**repeating `form`**. The block holds `form1`..`form8`, so a second instance runs
`form9`..`form16` — its seven entries are `form9`..`form15` and its note is `form16`,
with the extra notes of one form sharing that number (`form16.note1.value`,
`form16.note2.value`).
That is the correct Cocina encoding, but the numbers jump, and someone scanning the
headers may read it as a mistake.

### Step 3 — Pre-fill the rows from an uploaded list

**Always ask whether they have an object list; the asking is not optional, only the
list is.** The question belongs in the Step 2 confirmation so it costs no extra round
trip. Ask every time, including when nothing in the request suggests they have one —
the alternative is someone pasting identifiers in by hand, and a druid transposed by
hand is a broken PURL that nothing downstream will catch. If the answer is no, build
the empty workbook.

**Ask them to upload the file, not to type a path.** "Upload a .csv, .tsv or .xlsx"
is the request; they attach it to the conversation the way they would any other file.
Asking for a path makes them go and find one — open a file manager, copy a location,
get the separators right — for no benefit, and a mistyped path is a failed build
rather than a pre-filled sheet. If they happen to give you a path anyway, use it; the
builder does not care which way the file arrived.

Whichever way it arrives, the file is on disk by the time you build, so pass its path
to `--data`:

```bash
python3 scripts/build_spreadsheet.py --spec /tmp/spec.json --data objects.csv --output out.xlsx
```

How the file is read, so you can set expectations:

- **The first three columns are druid, source ID and title, in that order.** Columns are
  taken positionally, not by name — anything after the third is ignored.
- **A heading row is skipped** if any of its first three cells looks like a heading
  (`druid`, `source_id`, `title`, `id`, `label`, `purl`). A file with headings and the
  same file without them produce identical output, so there is no need to ask which
  they have.
- **Blank rows are dropped**; a blank cell within a row is fine and just leaves that
  entry empty.
- The title goes in the title block's **entry** cell — the one column in that block with
  no Cocina header of its own in row 3, labelled "Main title". That placement is
  deliberate: it is the cell the formulas read to decide whether the title is simple or
  structured, so a title written anywhere else would not drive `title1.value`.

Each object gets a row, and **the formulas and dropdowns are filled down across those
rows plus one spare row past the last object** — so `purl` derives per row from that
row's own druid, the whole `adminMetadata` run populates immediately, and every lookup
dropdown is live on each row rather than only the first.

**The spare row is there to fill down from.** It carries the formulas, dropdowns and
formatting but no druid, source id or title, so dragging it down copies only the
formulas and dropdowns. Filling down from the last object's row instead would copy that
object's druid, source id and title into every row it touches, and the entered metadata
would have to be cleared out again. Below the spare row the plain blank rows continue as
in an empty workbook.

The builder reports how many rows it populated, which heading row it skipped, and warns
about rows with no title. Read that back to the user — a count that doesn't match their
expectation usually means a stray heading or a blank line in the file.

### Step 4 — Write the spec and build

**Always write a fresh spec from what they asked for. Never build from a `spec.json`
you found on disk.** A spec lying around in the output directory came from some earlier
batch; it is not a record of this request, and building from it produces a workbook that
looks deliberate and is silently the wrong shape. This has happened: a run given "same
shape as before" went looking, found a sibling run's `spec.json`, and built from the
file rather than the sentence.

"Same shape as before", "same as last time" and "like the photographs one" refer to the
conversation, not the filesystem. If the earlier shape is in the conversation, reuse it
from there. If it isn't, reconstruct it from what they describe now and list the blocks
back in the reply, so a misreading is visible to them rather than buried in a column
count. A found spec may be read as a hypothesis to check field-by-field against their
words — never as a shortcut past reading them — and if you do lean on one, name the file
in the reply so they can catch it.

**Write the spec to a scratch location — not the skill directory, and not the directory
the workbook goes to.** A spec left beside the deliverable is exactly what the next run
trips over, and the person creating metadata has no use for it. Same for any checker scripts: the
delivery directory should contain the workbook and nothing else.

The spec itself is small:

```json
{
  "blocks": [
    {"type": "title", "count": 1},
    {"type": "contributor", "count": 2, "roles": [3, 1]},
    {"type": "form", "count": 2, "notes": [1, 2]},
    {"type": "subject", "count": 2},
    {"type": "access"},
    {"type": "adminMetadata"}
  ]
}
```

Block order in the list is the column order in the output, except `adminMetadata`,
which is always moved last. `count` defaults to 1. `data_rows` sets how many blank
fill-down rows sit below the formula row and defaults to matching the template (998).

```bash
python3 scripts/build_spreadsheet.py --spec /tmp/spec.json --output cocina_spreadsheet.xlsx
```

Use whichever interpreter Step 0 found:

- **Windows with Python**: `python scripts\build_spreadsheet.py --spec ... --output ...`
  Works directly; there is no need for WSL.
- **macOS or Linux**: `python3 ...` as above.
- **Windows with no Python at all**: borrow an interpreter from WSL. This is a fallback
  for that situation, not a requirement of the skill:
  ```
  wsl -d <distro> -- python3 /mnt/c/path/to/scripts/build_spreadsheet.py --spec ... --output ...
  ```
  Issue that from PowerShell, not Git Bash. Git Bash rewrites `/mnt/c/...` arguments on
  the way to `wsl.exe` — the path becomes something like
  `C:/Program Files/Git/mnt/c/...` and then resolves against the WSL working directory
  instead of failing cleanly, so the error reads like a broken script rather than a
  mangled path. `MSYS_NO_PATHCONV=1` also works if you are stuck in Git Bash.

  More generally, **anything passed inline through a shell to `wsl.exe` is liable to be
  chewed up before WSL sees it** — `$variables` and heredocs inside `bash -lc '...'`
  from Git Bash, and PowerShell here-strings or `\$`-escapes carrying inline Python or
  bash. The failures look like bugs in the thing you are running rather than quoting
  damage, which wastes real time. Write anything beyond a single flat command to a
  `.sh` or `.py` file and invoke that file; inside it, quoting behaves normally.

Paths in the spec and `--output` are resolved by the interpreter you invoke, so if you
go through WSL they need to be paths that WSL can see (`/mnt/c/...`). The bundled
template is found relative to the script itself, so the working directory does not
matter.

`--list-blocks` prints the block table read live from the template, which is a quick
way to confirm the template is intact before building.

The script refuses to build rather than guess when something is wrong: a missing
`title` block, an unknown block type, a repeat count on `access` or `adminMetadata`, a
`roles` or `notes` list that doesn't match its parent's count, a `roles` list on a block
other than `contributor` (or `notes` on a block other than `form`), or a template whose
layout no longer matches what the block ranges were derived from. Read the error and fix the spec — these messages are
specific, and working around one by hand-editing the output is how silently wrong
workbooks get shipped.

### Step 5 — Check what came out

The script prints the emitted blocks with their new column ranges and a sample
renumbered header. Read it and confirm the counts match what was agreed, then sanity
check the headers renumbered as expected — `contributor2.role1.value` rather than a
second `contributor1.role1.value`. Duplicate headers are the failure mode that
matters, because the load process keys on them.

If you write a script to check the output, read `references/checking.md` first. It
lists five ways a checker on this workbook reports a clean pass or a false defect
rather than an error — inline strings, `sheet2.xml`, self-closing cells, shifting
column letters — and the five false positives not worth reporting. Each has caught a
real checker here.

### Step 6 — Hand it over

**Keep the reply short and professional.** These are working metadata creators, not
readers of a manual: the file, the shape, and anything that differs from what they asked
for. Aim for well under 200 words.

**Say "creating metadata", not "cataloguing".** That is the house vocabulary, in the
reply as well as in this file: "before creating metadata", "once the metadata is
entered", "the person creating the metadata" — never "before cataloguing starts",
"after cataloguing", or "the cataloguer". The one exception is the `adminMetadata`
block, whose field genuinely holds the *cataloguing agency* and whose derived value
reads "original cataloging agency"; name that field as it is.

**Say "field set", not "block".** Same rule as Step 2, and it applies to the hand-over
in full: "1 title, 2 contributors…" needs no noun at all, but where one is needed it is
*field set*. `block` stays inside the spec, the builder output and this file.

**Name which subject kind in the shape line**, not just further down: "1 multipart
subject" or "3 subjects", never a bare "1 subject" for the structured set. The two are
alternatives and a rebuild is a new file, so which one they got is exactly the thing
worth catching on the first read rather than in a later paragraph.

**Professional means a colleague's register, not a service voice.** Plain declarative
sentences about the work. No greeting, no sign-off, no "Great question", no "Happy to
help", no "I hope this helps", no congratulating them on the batch, no exclamation
marks, no emoji. Don't announce what you are about to do or narrate your own diligence
("I carefully verified…", "As requested…") — the reply is the result, not a report on
the effort. Don't apologise for the template's design or for constraints you didn't
choose; state them. Where you made a judgement call, say what you decided and why in
one sentence, without hedging it into vagueness.

Concretely:

- No section headings, and no bold lead-ins on every bullet.
- One line per point. If a point needs a paragraph of justification, it belongs in an
  answer to a follow-up question, not in the hand-over.
- Say what you built, not how the builder works. Nothing from "What the builder
  guarantees" below goes in unprompted — it is there so you can answer questions, not
  so you can pre-empt them.
- Raise a caveat only if it applies to *this* file. A pre-filled workbook doesn't need
  the fill-down instruction.
- **No export or download instructions.** Don't tell them to select the `metadata` tab,
  don't tell them which CSV format to use, don't describe getting the file out of
  Sheets. That comes long after the hand-over, the workbook's own `Instructions` tab
  carries it, and the README repeats it — putting it in the reply pushes a step they
  cannot take yet ahead of the one they can. Answer it if they ask.
- **Don't say the SDR load requires resource type, or that any block is required by it.**
  `title` is required by the builder, and that is the only requirement to state.
- **Never lead with the state of the vocabulary tabs.** They hold example rows by
  design and the team fills them in; saying "the role tab only has author" reads as a
  fault report about their own template. If the batch obviously needs terms they'll
  have to add — a photograph batch needing `photographer` — one clause is the most it
  is worth, and only alongside something they actually have to act on.
- Don't explain the same thing twice, and don't close with a summary of what you just
  said.

A hand-over that covers everything:

```
Built: example_batch.xlsx — 110 columns.

1 title, 2 contributors (1 role each), origin info, 3 subjects, access
information, admin metadata. Left out: form/genre, language, note, identifier,
related resource, geographic.

You asked for 2 access field sets; access can't repeat, so there's one. Its
headers carry no instance number, and a second copy would collide.

Worth considering: form/genre, which holds resource type, genre, extent and
media type. I left it out since you didn't ask — say the word and I'll rebuild.

Formulas are on row 4 — fill down to your 40 rows.

File at: `C:\Users\arcadia\Documents\example_batch.xlsx`
```

That is the ceiling, not the target. Cut anything that doesn't apply — except the
`File at:` line, which always stays. The filename, counts and caveats above are an
invented scenario chosen to exercise every part of the shape; take the wording, never
the content. In particular the "you asked for 2 access field sets" line only belongs in
a reply where they actually did.

**The fill-down number is how many objects they said they have**, not the row count in
the file. "40 rows" above is the batch in that invented scenario. A workbook with no
object list carries the template's 998 blank rows whatever the batch size, so "fill down
to your 998 rows" tells someone with 15 maps to fill down 983 rows they will never use —
and it reads as a fact about their batch, which it isn't. If they never said a number,
drop it: "Formulas are on row 4 — fill down as far as you need." If their list was
pre-filled, the line goes altogether; the rows already have their formulas, and the
spare row past the last object is worth one clause instead: "row 44 is empty and set up
— fill down from it, not from row 43, if you add more objects."

**Always end with the workbook's full path, labelled `File at:`, in a code span.** Not a
markdown link — a plain path:

```
File at: `C:\Users\arcadia\Documents\example_batch.xlsx`
```

**Don't make it a link.** Measured in the Claude Code desktop app on 2026-09-22, a
link to the workbook cannot work: a file outside the session's working directory is
unreachable by every mechanism, including the attachment tool, and the save location is
normally `Documents` or `Desktop`. A link that looks live and isn't is worse than plain
text. `references/file-delivery.md` has the full results if you need to re-check them.

So the code span is doing real work: it keeps the path byte-accurate whatever
punctuation it contains, and it makes it selectable for pasting into Explorer or a file
manager. The Instructions tab and README describe that route.

Use the **absolute path with backslashes**, the native form, even when the file sits
under the working directory. Naming the file on the first line and its path on the last
is not repetition — the first says what was built, the last says where it is.

The label is `File at:` rather than `Open:` or `Download:` because it describes what the
line is: a location on their disk. Nothing is being fetched, and nothing opens by
itself — a label that promises either sets up the disappointment.

Never hand back a `\\wsl.localhost\...` UNC prefix or a `/mnt/c/...` WSL path. The
second is real inside WSL and meaningless to their file manager; translate it to the
`C:\...` form of the same file.

If a future client does render working file links, this is the paragraph to revisit —
re-test it rather than assuming, since the failure mode is a link that looks fine.

Two things are worth a line when they apply:

- **A rebuild is a new file**, not something mergeable into a copy that already has
  metadata in it. Say it when a count is uncertain or you are suggesting a field set.
- **A fresh workbook looks emptier than people expect**, because the computed columns
  derive from what is typed. Worth a clause if they might read it as a broken file.

Answer these if asked, but never volunteer them:

- **Exporting.** A CSV holds one sheet and takes the active one. The workbook opens on
  `metadata` at A1, so the risk is exporting after a detour through `Instructions`
  rather than on the first try; in Excel the format must be `CSV UTF-8 (Comma delimited)`,
  since plain CSV writes `é` as the cp1252 byte `E9` and corrupts every diacritic
  silently. The Sheets route sidesteps both. This is real and it matters — it is just
  not hand-over material, because it belongs to a day that hasn't happened yet.
- **Dates** are Text-formatted, so `1923`, `circa 1923`, `192-` and open ranges survive
  verbatim, and Text does not reorder `06/04/1923`. The workbook has been checked in
  Google Sheets as well as Excel; if a date converts there anyway, the fix is
  Format → Number → Plain text.
- **Either tool works.** Creating metadata in Sheets or in Excel is supported, so don't
  write instructions that fit only one.

## Answering questions about the output

`references/guarantees.md` records what the builder guarantees and why: per-copy
formula repointing, header renumbering, absolute dropdown ranges, `status` on first
instances only, Text-formatted dates, the three guards that stop a cell ever showing
`0` or an error, and the vocabulary tabs being example rows by design. Read it when
someone queries the output; don't recite it unasked.
