# Block reference

Every block in the bundled template's `metadata` sheet, with the headers
(row 3) and entry labels (row 2) it contributes. Read this when you need to
tell someone what a block covers, or to check whether a field they asked
for already lives inside a block they've requested.

Columns are positions **in the template**, not in the generated file — once
blocks repeat or drop out, everything shifts. Cells marked *entry* are where
the cataloguer types; the rest are computed by formula or left for manual entry.

| block | template cols | repeats | headers |
|---|---|---|---|
| `title` | D–P | yes | 13 |
| `contributor` | Q–W | yes | 7 |
| `contributor > role` | X–AA | nested | 4 |
| `form` | AB–AS | yes | 18 |
| `form > note` | AT–AU | nested | 2 |
| `event` | AV–BN | yes | 19 |
| `language` | BO–BR | yes | 4 |
| `note` | BS–BU | yes | 3 |
| `identifier` | BV–BX | yes | 3 |
| `access` | BY–CD | no | 6 |
| `subject` | CE–CQ | yes | 13 |
| `relatedResource` | CR–CX | yes | 7 |
| `geographic` | CY–DT | yes | 22 |
| `adminMetadata` | DU–EC | no | 9 |


## `title`  (D–P)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| D | Nonsorting article | `title1.structuredValue1.value` | *entry* |
| E |  | `title1.structuredValue1.type` | `=IFERROR(if(isblank(D4),"","nonsorting characters"),"")` |
| F | Main title | — | *entry* |
| G |  | `title1.structuredValue2.value` | `=IFERROR(if(not(and(isblank(D4),isblank(J4),isblank(L4),isblank(N4))),IF(ISBLANK(F4),"",F4),""),"")` |
| H |  | `title1.structuredValue2.type` | `=IFERROR(if(G4="","","main title"),"")` |
| I |  | `title1.value` | `=IFERROR(if(and(isblank(D4),isblank(J4),isblank(L4),isblank(N4)),IF(ISBLANK(F4),"",F4),""),"")` |
| J | Subtitle | `title1.structuredValue3.value` | *entry* |
| K |  | `title1.structuredValue3.type` | `=IFERROR(if(isblank(J4),"","subtitle"),"")` |
| L | Part number | `title1.structuredValue4.value` | *entry* |
| M |  | `title1.structuredValue4.type` | `=IFERROR(if(isblank(L4),"","part number"),"")` |
| N | Part name | `title1.structuredValue5.value` | *entry* |
| O |  | `title1.structuredValue5.type` | `=IFERROR(if(isblank(N4),"","part name"),"")` |
| P |  | `title1.status` | `=IFERROR(if(isblank(F4),"","primary"),"")` |


## `contributor`  (Q–W)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| Q | Name | `contributor1.name1.value` | *entry* |
| R | Status | `contributor1.status` | `=IFERROR(if(isblank(Q4),"","primary"),"")` |
| S |  | `contributor1.type` | `=IFERROR(if(isblank(Q4),"",vlookup(Q4,contributor!A:F,2,false)&""),"")` |
| T |  | `contributor1.name1.uri` | `=IFERROR(if(isblank(Q4),"",vlookup(Q4,contributor!A:F,3,false)&""),"")` |
| U |  | `contributor1.name1.source.code` | `=IFERROR(if(isblank(Q4),"",vlookup(Q4,contributor!A:F,4,false)&""),"")` |
| V |  | `contributor1.identifier1.uri` | `=IFERROR(if(isblank(Q4),"",vlookup(Q4,contributor!A:F,5,false)&""),"")` |
| W |  | `contributor1.identifier1.type` | `=IFERROR(if(isblank(Q4),"",vlookup(Q4,contributor!A:F,6,false)&""),"")` |


## `contributor > role`  (X–AA)

Nested: repeats inside its parent block, one or more per parent
instance. The `role` number counts within the parent, while the
leading segment follows the parent's own numbering.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| X | Role | `contributor1.role1.value` | *entry* |
| Y |  | `contributor1.role1.code` | `=IFERROR(if(isblank(X4),"",vlookup(X4,role!A:D,2,false)&""),"")` |
| Z |  | `contributor1.role1.uri` | `=IFERROR(if(isblank(X4),"",vlookup(X4,role!A:D,3,false)&""),"")` |
| AA |  | `contributor1.role1.source.code` | `=IFERROR(if(isblank(X4),"",vlookup(X4,role!A:D,4,false)&""),"")` |


## `form`  (AB–AS)

Repeats. Holds 8 numbered entries (`form1`..`form8`), so instance 2
starts at `form9`.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| AB | General resource type | `form1.value` | *entry* |
| AC |  | `form1.type` | `=IFERROR(if(isblank(AB4),"","resource type"),"")` |
| AD |  | `form1.uri` | `=IFERROR(if(isblank(AB4),"",vlookup(AB4,'resource type'!A:C,2,false)&""),"")` |
| AE |  | `form1.source.value` | `=IFERROR(if(isblank(AB4),"",vlookup(AB4,'resource type'!A:C,3,false)&""),"")` |
| AF | Form | `form2.value` | *entry* |
| AG |  | `form2.type` | `=IFERROR(if(isblank(AF4),"","form"),"")` |
| AH | Extent | `form3.value` | *entry* |
| AI |  | `form3.type` | `=IFERROR(if(isblank(AH4),"","extent"),"")` |
| AJ | Genre | `form4.value` | *entry* |
| AK |  | `form4.type` | `=IFERROR(if(isblank(AJ4),"","genre"),"")` |
| AL |  | `form4.uri` | `=IFERROR(if(isblank(AJ4),"",vlookup(AJ4,genre!A:C,2,false)&""),"")` |
| AM |  | `form4.source.code` | `=IFERROR(if(isblank(AJ4),"",vlookup(AJ4,genre!A:C,3,false)&""),"")` |
| AN | Reformatting quality | `form5.value` | *entry* |
| AO |  | `form5.type` | `=IFERROR(if(isblank(AN4),"","reformatting quality"),"")` |
| AP | Digital origin | `form6.value` | *entry* |
| AQ |  | `form6.type` | `=IFERROR(if(isblank(AP4),"","digital origin"),"")` |
| AR | Internet media type | `form7.value` | *entry* |
| AS |  | `form7.type` | `=IFERROR(if(isblank(AR4),"","media type"),"")` |


## `form > note`  (AT–AU)

Nested: repeats inside its parent block, one or more per parent
instance. The `note` number counts within the parent, while the
leading segment follows the parent's own numbering.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| AT | Form note | `form8.note1.value` | *entry* |
| AU | Display label | `form8.note1.displayLabel` | *entry* |


## `event`  (AV–BN)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| AV | Publication or creation | `event1.type` | *entry* |
| AW | Date (YYYY-MM-DD) | — | *entry* |
| AX |  | `event1.date1.value` | `=IFERROR(if(isblank(BD4),IF(ISBLANK(AW4),"",AW4),""),"")` |
| AY |  | `event1.date1.type` | `=IFERROR(if(isblank(AW4),"",IF(ISBLANK(AV4),"",AV4)),"")` |
| AZ |  | `event1.date1.status` | `=IFERROR(IF(isblank(AW4),"","primary"),"")` |
| BA |  | `event1.date1.encoding.code` | `=IFERROR(if(isblank(AW4),"","w3cdtf"),"")` |
| BB |  | `event1.date1.structuredValue1.value` | `=IFERROR(if(isblank(BD4),"",IF(ISBLANK(AW4),"",AW4)),"")` |
| BC |  | `event1.date1.structuredValue1.type` | `=IFERROR(if(isblank(BD4),"","start"),"")` |
| BD | End date if range | `event1.date1.structuredValue2.value` | *entry* |
| BE |  | `event1.date1.structuredValue2.type` | `=IFERROR(if(isblank(BD4),"","end"),"")` |
| BF | Approximate date? | `event1.date1.qualifier` | *entry* |
| BG | Publication/creation place | `event1.location1.value` | *entry* |
| BH | Publication/creation place URI | `event1.location1.uri` | *entry* |
| BI | Publisher | `event1.contributor1.name1.value` | *entry* |
| BJ |  | `event1.contributor1.type` | `=IFERROR(if(isblank(BI4),"","organization"),"")` |
| BK |  | `event1.contributor1.role1.value` | `=IFERROR(if(isblank(BI4),"","publisher"),"")` |
| BL |  | `event1.contributor1.role1.code` | `=IFERROR(if(isblank(BI4),"","pbl"),"")` |
| BM |  | `event1.contributor1.role1.uri` | `=IFERROR(if(isblank(BI4),"","http://id.loc.gov/vocabulary/relators/pbl"),"")` |
| BN |  | `event1.contributor1.role1.source.code` | `=IFERROR(if(isblank(BI4),"","marcrelator"),"")` |


## `language`  (BO–BR)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| BO | Language | `language1.value` | *entry* |
| BP |  | `language1.code` | `=IFERROR(if(isblank(BO4),"",vlookup(BO4,language!A:D,2,false)&""),"")` |
| BQ |  | `language1.uri` | `=IFERROR(if(isblank(BO4),"",vlookup(BO4,language!A:D,3,false)&""),"")` |
| BR |  | `language1.source.code` | `=IFERROR(if(isblank(BO4),"",vlookup(BO4,language!A:D,4,false)&""),"")` |


## `note`  (BS–BU)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| BS | Note | `note1.value` | *entry* |
| BT | Type | `note1.type` | *entry* |
| BU | Display label | `note1.displayLabel` | *entry* |


## `identifier`  (BV–BX)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| BV | Identifier | `identifier1.value` | *entry* |
| BW | Type | `identifier1.type` | *entry* |
| BX | Display label | `identifier1.displayLabel` | *entry* |


## `access`  (BY–CD)

Appears at most once. Its headers carry no instance number, so a
second copy would collide with the first.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| BY | Repository/Library | `access.accessContact1.value` | *entry* |
| BZ | URI | `access.accessContact1.uri` | *entry* |
| CA | Authority code | `access.accessContact1.source.code` | *entry* |
| CB |  | `access.accessContact1.type` | `=IFERROR(if(isblank(BY4),"","repository"),"")` |
| CC | Shelf locator | `access.physicalLocation1.value` | *entry* |
| CD |  | `access.physicalLocation1.type` | `=IFERROR(if(isblank(CC4),"","shelf locator"),"")` |


## `subject`  (CE–CQ)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| CE | Subject #1 part 1 value | — | *entry* |
| CF |  | `subject1.value` | `=IFERROR(if(AND(not(isblank(CE4)),isblank(CN4)),IF(ISBLANK(CE4),"",CE4),""),"")` |
| CG |  | `subject1.type` | `=IFERROR(if(and(not(isblank(CE4)),isblank(CN4)),vlookup(CE4,subject!A:D,2,false)&"",""),"")` |
| CH |  | `subject1.uri` | `=IFERROR(if(and(not(isblank(CE4)),isblank(CN4)),vlookup(CE4,subject!A:D,3,false)&"",""),"")` |
| CI |  | `subject1.source.code` | `=IFERROR(if(and(not(isblank(CE4)),isblank(CN4)),vlookup(CE4,subject!A:D,4,false)&"",""),"")` |
| CJ |  | `subject1.structuredValue1.value` | `=IFERROR(if(not(isblank(CN4)),IF(ISBLANK(CE4),"",CE4),""),"")` |
| CK |  | `subject1.structuredValue1.type` | `=IFERROR(if(and(not(isblank(CE4)),not(isblank(CN4))),vlookup(CE4,subject!A:D,2,false)&"",""),"")` |
| CL |  | `subject1.structuredValue1.uri` | `=IFERROR(if(and(not(isblank(CE4)),not(isblank(CN4))),vlookup(CE4,subject!A:D,3,false)&"",""),"")` |
| CM |  | `subject1.structuredValue1.source.code` | `=IFERROR(if(and(not(isblank(CE4)),not(isblank(CN4))),vlookup(CE4,subject!A:D,4,false)&"",""),"")` |
| CN | Subject #1 part 2 value | `subject1.structuredValue2.value` | *entry* |
| CO |  | `subject1.structuredValue2.type` | `=IFERROR(if(isblank(CN4),"",vlookup(CN4,subject!A:D,2,false)&""),"")` |
| CP |  | `subject1.structuredValue2.uri` | `=IFERROR(if(isblank(CN4),"",vlookup(CN4,subject!A:D,3,false)&""),"")` |
| CQ |  | `subject1.structuredValue2.source.code` | `=IFERROR(if(isblank(CN4),"",vlookup(CN4,subject!A:D,4,false)&""),"")` |


## `relatedResource`  (CR–CX)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| CR | Type | `relatedResource1.type` | *entry* |
| CS | Display label | `relatedResource1.displayLabel` | *entry* |
| CT | Title | `relatedResource1.title1.value` | *entry* |
| CU | PURL | `relatedResource1.purl` | *entry* |
| CV | Other URL | `relatedResource1.access.url1.value` | *entry* |
| CW | Abstract | `relatedResource1.note1.value` | *entry* |
| CX |  | `relatedResource1.note1.type` | `=IFERROR(if(isblank(CW4),"","abstract"),"")` |


## `geographic`  (CY–DT)

Repeats; each instance advances the number by one.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| CY | MIME type | `geographic1.form1.value` | *entry* |
| CZ |  | `geographic1.form1.type` | `=IFERROR(if(isblank(CY4),"","media type"),"")` |
| DA |  | `geographic1.form1.source.value` | `=IFERROR(if(isblank(CY4),"","IANA media type terms"),"")` |
| DB | Dublin Core resource type | `geographic1.form2.value` | *entry* |
| DC |  | `geographic1.form2.type` | `=IFERROR(if(isblank(DB4),"","media type"),"")` |
| DD |  | `geographic1.form2.source.value` | `=IFERROR(if(isblank(DB4),"","DCMI Type Vocabulary"),"")` |
| DE | Latitude | `geographic1.subject1.structuredValue1.value` | *entry* |
| DF |  | `geographic1.subject1.structuredValue1.type` | `=IFERROR(if(isblank(DE4),"","latitude"),"")` |
| DG | Longitude | `geographic1.subject1.structuredValue2.value` | *entry* |
| DH |  | `geographic1.subject1.structuredValue2.type` | `=IFERROR(if(isblank(DG4),"","longitude"),"")` |
| DI |  | `geographic1.subject1.type` | `=IFERROR(if(isblank(DE4),"","point coordinates"),"")` |
| DJ |  | `geographic1.subject1.encoding.value` | `=IFERROR(if(isblank(DE4),"","decimal"),"")` |
| DK | West | `geographic1.subject2.structuredValue1.value` | *entry* |
| DL |  | `geographic1.subject2.structuredValue1.type` | `=IFERROR(if(isblank(DK4),"","west"),"")` |
| DM | South | `geographic1.subject2.structuredValue2.value` | *entry* |
| DN |  | `geographic1.subject2.structuredValue2.type` | `=IFERROR(if(isblank(DM4),"","south"),"")` |
| DO | East | `geographic1.subject2.structuredValue3.value` | *entry* |
| DP |  | `geographic1.subject2.structuredValue3.type` | `=IFERROR(if(isblank(DO4),"","east"),"")` |
| DQ | North | `geographic1.subject2.structuredValue4.value` | *entry* |
| DR |  | `geographic1.subject2.structuredValue4.type` | `=IFERROR(if(isblank(DQ4),"","north"),"")` |
| DS |  | `geographic1.subject2.type` | `=IFERROR(if(isblank(DK4),"","bounding box coordinates"),"")` |
| DT |  | `geographic1.subject2.encoding.value` | `=IFERROR(if(isblank(DK4),"","decimal"),"")` |


## `adminMetadata`  (DU–EC)

Appears at most once. Its headers carry no instance number, so a
second copy would collide with the first.

| col | entry label (row 2) | header (row 3) | formula |
|---|---|---|---|
| DU | Language of cataloging | `adminMetadata.language1.value` | `=IFERROR(if(isblank(F4),"","English"),"")` |
| DV |  | `adminMetadata.language1.code` | `=IFERROR(if(isblank(F4),"","eng"),"")` |
| DW |  | `adminMetadata.language1.source.code` | `=IFERROR(if(isblank(F4),"","iso639-2b"),"")` |
| DX |  | `adminMetadata.language1.uri` | `=IFERROR(if(isblank(F4),"","http://id.loc.gov/vocabulary/iso639-2/eng"),"")` |
| DY |  | `adminMetadata.contributor1.name1.code` | `=IFERROR(if(ISBLANK(F4),"","CSt"),"")` |
| DZ |  | `adminMetadata.contributor1.name1.uri` | `=IFERROR(if(isblank(F4),"","http://id.loc.gov/vocabulary/organizations/cst"),"")` |
| EA |  | `adminMetadata.contributor1.name1.source.code` | `=IFERROR(if(ISBLANK(F4),"","marcorg"),"")` |
| EB |  | `adminMetadata.contributor1.type` | `=IFERROR(if(isblank(F4),"","organization"),"")` |
| EC |  | `adminMetadata.contributor1.role1.value` | `=IFERROR(if(isblank(F4),"","original cataloging agency"),"")` |


## Fixed leading columns (A–C)

| col | header | formula |
|---|---|---|
| A | `druid` | *entry* |
| B | `source_id` | *entry* |
| C | `purl` | `=IFERROR(if(isblank(A4),"",concatenate("https://purl.stanford.edu/",A4)),"")` |

These always lead the sheet and are never repeated or dropped.


## Lookup sheets

The workbook ships six lookup sheets that the `vlookup` formulas read:
`contributor`, `role`, `resource type`, `genre`, `language`, `subject`.
They are copied into the output untouched, and the dropdowns on the entry
cells point at them. Adding rows to a lookup sheet afterwards is safe —
the ranges already run to row 1002.

