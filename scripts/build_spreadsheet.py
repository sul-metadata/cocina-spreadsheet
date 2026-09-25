#!/usr/bin/env python3
"""Build a Cocina metadata spreadsheet by repeating/omitting header blocks.

Reads the bundled template workbook, rebuilds its `metadata` sheet from a
block spec, and copies every other sheet through untouched. Uses only the
standard library (no openpyxl).

Usage:
    python3 build_spreadsheet.py --spec spec.json --output out.xlsx
    python3 build_spreadsheet.py --list-blocks
"""

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = "{%s}" % MAIN
ET.register_namespace("", MAIN)

# Excel's own extensions. A workbook saved by Excel carries validations whose
# list comes from another sheet in an <extLst> under these namespaces instead of
# in the plain <dataValidations> block, and tags rows with x14ac:dyDescent.
# Both are read here; neither is written back, since the output declares only
# the main namespace.
X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
XM = "{http://schemas.microsoft.com/office/excel/2006/main}"

# Parts that are not comments but still must not be copied into the output.
# calcChain.xml caches the order Excel evaluated formulas in; the generated
# sheet has different columns, so a copied chain names cells that no longer
# exist and Excel offers to repair the file. force_recalc() sets
# fullCalcOnLoad, which makes the chain redundant anyway.
DROP_PARTS = ("xl/calcChain.xml",)
DROP_REL_HINTS = ("calcChain",)

SHEET_PART = "xl/worksheets/sheet2.xml"
SHEET_RELS = "xl/worksheets/_rels/sheet2.xml.rels"

# Cell comments are dropped from every sheet, not carried into the output. A
# comment spans five parts plus a relationship, a content-type override and a
# <legacyDrawing> element, and a half-removed one makes Excel offer to repair
# the file - so the removal is driven off these patterns rather than a list of
# specific filenames, which would go stale the moment the template changes.
COMMENT_PART_PATTERNS = (
    r"^xl/comments\d+\.xml$",
    r"^xl/threadedComments/threadedComment\d+\.xml$",
    r"^xl/drawings/vmlDrawing\d+\.vml$",
    r"^xl/documenttasks/documenttask\d+\.xml$",
    r"^xl/persons/person\.xml$",
)
COMMENT_REL_HINTS = ("comments", "threadedComment", "vmlDrawing",
                     "documenttask", "/person")


def is_comment_part(name):
    return any(re.match(p, name) for p in COMMENT_PART_PATTERNS)


def is_dropped_part(name):
    return is_comment_part(name) or name in DROP_PARTS


def plain_attrs(el):
    """An element's attributes with namespaced ones removed.

    ElementTree hands back a namespaced attribute as "{uri}local". The sheet is
    assembled as text, so writing that key out produces an attribute name no
    parser accepts - which is exactly what happened when Excel added
    x14ac:dyDescent to every row. The output declares only the main namespace,
    so the honest thing is to drop what it cannot name; dyDescent is a
    rendering hint Excel recomputes.
    """
    return {k: v for k, v in el.attrib.items() if not k.startswith("{")}


def strip_ns_attrs(el):
    """Same, applied in place to an element and its descendants."""
    for node in el.iter():
        for k in [k for k in node.attrib if k.startswith("{")]:
            del node.attrib[k]
    return el


def home_the_cursor(sheet_views):
    """Open the metadata sheet at A1, wherever the template was left.

    Excel stores the cursor and scroll position of the sheet it was saved on,
    so a template last edited in column EJ hands every generated workbook to
    the cataloguer scrolled past the end of the headers with a far-right cell
    selected. Normalising here rather than in the template means the next save
    cannot re-pin it.

    A frozen pane is left alone: its selections are per-pane, and this template
    has none, so rewriting them would be a guess.
    """
    for view in sheet_views:
        if view.find(NS + "pane") is not None:
            continue
        view.attrib.pop("topLeftCell", None)
        for sel in view.findall(NS + "selection"):
            view.remove(sel)
        sel = ET.SubElement(view, NS + "selection")
        sel.set("activeCell", "A1")
        sel.set("sqref", "A1")
    return sheet_views

HEADER_ROWS = (1, 2, 3)
FORMULA_ROW = 4
BLANK_ROW = 5


# --------------------------------------------------------------------------
# column helpers
# --------------------------------------------------------------------------

def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(idx):
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------
# block table
#
# Ranges are source-template column letters. They are validated against
# FINGERPRINT at load time, so a template revision fails loudly instead of
# silently shifting a block by one column.
# --------------------------------------------------------------------------

BLOCK_DEFS = [
    dict(key="title", first="D", last="P", repeatable=True, number_seg=0,
         required=True, first_only=("title1.status",),
         repeat_only=("title1.type",)),
    dict(key="contributor", first="Q", last="W", repeatable=True, number_seg=0,
         first_only=("contributor1.status",),
         child=dict(key="role", first="X", last="AA", repeatable=True, number_seg=1)),
    # The template's form entries form1..form8 are split into two field sets
    # that may be chosen separately or together: descriptive form (resource
    # type, form, extent, genre) and technical details (reformatting quality,
    # digital origin, media type, and a nested note repeatable within its
    # field set the way a role is within a contributor). Both belong to the
    # "form" numbering group, so their headers number consecutively across
    # whichever instances are emitted - see build_plan().
    dict(key="form", first="AB", last="AM", repeatable=True, number_seg=0,
         number_group="form"),
    dict(key="technicalDetails", first="AN", last="AS", repeatable=True,
         number_seg=0, number_group="form",
         child=dict(key="note", first="AT", last="AU", repeatable=True,
                    number_seg=1)),
    dict(key="event", first="AV", last="BN", repeatable=True, number_seg=0,
         first_only=("event1.date1.status",)),
    dict(key="language", first="BO", last="BR", repeatable=True, number_seg=0),
    dict(key="note", first="BS", last="BU", repeatable=True, number_seg=0),
    dict(key="identifier", first="BV", last="BX", repeatable=True, number_seg=0),
    dict(key="access", first="BY", last="CD", repeatable=False, number_seg=None),
    # Two subject field sets, and a workbook may carry only one of them. Both
    # number their headers from `subject1`, so a sheet holding both would have
    # `subject1.value` twice - CF in the multipart set, EF in the simple one.
    # EXCLUSIVE_GROUPS below is what enforces it.
    dict(key="multipartSubject", first="CE", last="CQ", repeatable=True,
         number_seg=0),
    dict(key="subject", first="EF", last="EI", repeatable=True, number_seg=0),
    dict(key="relatedResource", first="CR", last="CX", repeatable=True, number_seg=0),
    dict(key="geographic", first="CY", last="DT", repeatable=True, number_seg=0),
    dict(key="adminMetadata", first="DU", last="EC", repeatable=False,
         number_seg=None, always=True),
]

BLOCK_BY_KEY = {b["key"]: b for b in BLOCK_DEFS}

# Field sets that cannot share a workbook. The builder cannot choose between
# them - the person asking has to - so a spec naming both is an error rather
# than a silent preference for one.
EXCLUSIVE_GROUPS = [("subject", "multipartSubject")]


def CHILD_FIELD(child):
    """The spec key naming how many of a nested child each parent gets.

    Derived from the child's own name so the two cannot drift: a `role` child
    is counted by "roles", a `note` child by "notes".
    """
    return child["key"] + "s"


ALL_CHILD_FIELDS = tuple(CHILD_FIELD(b["child"]) for b in BLOCK_DEFS
                         if "child" in b)

PREFIX_FIRST, PREFIX_LAST = "A", "C"

# Spoken/legacy names the caller may use for a block key.
ALIASES = {
    "titles": "title", "contributors": "contributor", "creator": "contributor",
    "creators": "contributor", "forms": "form", "form/genre": "form",
    "genre": "form", "resource type": "form", "extent": "form",
    "technical details": "technicalDetails", "technical": "technicalDetails",
    "technicaldetails": "technicalDetails",
    "technical metadata": "technicalDetails",
    "reformatting quality": "technicalDetails",
    "digital origin": "technicalDetails", "media type": "technicalDetails",
    "internet media type": "technicalDetails",
    "form note": "technicalDetails", "form notes": "technicalDetails",
    "events": "event",
    "origin info": "event", "origininfo": "event", "publication": "event",
    "languages": "language", "notes": "note", "identifiers": "identifier",
    "access information": "access",
    # Plain "subject" means the simple set; the structured one has to be asked
    # for by name.
    "subjects": "subject", "simple subject": "subject",
    "simple subjects": "subject", "simplesubject": "subject",
    "single subject": "subject", "single subjects": "subject",
    "singlesubject": "subject",
    # What a cataloger calls them when not thinking in Cocina terms
    "keyword": "subject", "keywords": "subject",
    "topic": "subject", "topics": "subject",
    "multipart subject": "multipartSubject",
    "multipart subjects": "multipartSubject",
    "multipartsubject": "multipartSubject",
    "complex subject": "multipartSubject",
    "complex subjects": "multipartSubject",
    "complexsubject": "multipartSubject",
    "structured subject": "multipartSubject",
    "structured subjects": "multipartSubject",
    "related resource": "relatedResource", "relatedresource": "relatedResource",
    "related resources": "relatedResource", "geographic data": "geographic",
    "geo": "geographic", "administrative metadata": "adminMetadata",
    "adminmetadata": "adminMetadata", "admin": "adminMetadata",
    "roles": "role",
}

# Sentinel row-3 headers. Mismatch means the bundled template changed shape
# and the block ranges above need rechecking before any output is trusted.
FINGERPRINT = {
    "A": "druid", "C": "purl",
    "D": "title1.structuredValue1.value", "P": "title1.status",
    "Q": "contributor1.name1.value", "W": "contributor1.identifier1.type",
    "X": "contributor1.role1.value", "AA": "contributor1.role1.source.code",
    "AB": "form1.value", "AM": "form4.source.code",
    "AN": "form5.value", "AS": "form7.type",
    "AT": "form8.note1.value", "AU": "form8.note1.displayLabel",
    "AV": "event1.type", "BN": "event1.contributor1.role1.source.code",
    "BO": "language1.value", "BR": "language1.source.code",
    "BS": "note1.value", "BU": "note1.displayLabel",
    "BV": "identifier1.value", "BX": "identifier1.displayLabel",
    "BY": "access.accessContact1.value", "CD": "access.physicalLocation1.type",
    "CF": "subject1.value", "CQ": "subject1.structuredValue2.source.code",
    "EF": "subject1.value", "EI": "subject1.source.code",
    "CR": "relatedResource1.type", "CX": "relatedResource1.note1.type",
    "CY": "geographic1.form1.value", "DT": "geographic1.subject2.encoding.value",
    "DU": "adminMetadata.language1.value",
    "EC": "adminMetadata.contributor1.role1.value",
    # appended past the end of the sheet rather than inserted into the title
    # block, so that adding it renumbered nothing; emitted on title instances 2+
    "EE": "title1.type",
}

# adminMetadata's formulas all fire off the title block's "Main title" entry
# cell (column F) - the sheet's only cross-block reference. That is why the
# title block is always emitted: it is where the cataloguer types the title in
# Excel, and without that cell the adminMetadata columns could never fill in.
ADMIN_TRIGGER_SRC = "F"


class BuildError(Exception):
    pass


# --------------------------------------------------------------------------
# optional seed data: druid, source id, title
# --------------------------------------------------------------------------

# Words that mark a leading row as column headings rather than an object.
# Columns are taken positionally - first druid, second source id, third title -
# so the file's own headings are only used to recognise and skip that row.
HEADER_WORDS = {"druid", "sourceid", "source", "sourceidentifier", "id",
                "title", "label", "purl"}


def _looks_like_header(row):
    for cell in row[:3]:
        word = re.sub(r"[^a-z]", "", (cell or "").lower())
        if word in HEADER_WORDS:
            return True
    return False


def read_seed_data(path):
    """Read (druid, source_id, title) triples from .csv, .tsv or .xlsx.

    Only the first three columns are read, in that order. A leading heading
    row is skipped if it looks like one; anything beyond three columns is
    ignored, as are wholly blank rows.
    """
    ext = os.path.splitext(path)[1].lower()
    rows = []
    # .txt used to be accepted here and was never documented: a prose file
    # dropped in got sniffed for a delimiter and its first line became an
    # object, druid and all. Refusing it is the less surprising failure.
    if ext in (".csv", ".tsv"):
        import csv
        with open(path, newline="", encoding="utf-8-sig") as fh:
            sample = fh.read(8192)
            fh.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            except csv.Error:
                dialect = csv.excel_tab if ext == ".tsv" else csv.excel
            for raw in csv.reader(fh, dialect):
                rows.append([(c or "").strip() for c in raw])
    elif ext == ".xlsx":
        rows = _read_xlsx_rows(path)
    else:
        raise BuildError(
            "Don't know how to read %r. Provide the data as .csv, .tsv or "
            ".xlsx with druid, source id and title in the first three columns."
            % os.path.basename(path))

    rows = [r for r in rows if any((c or "").strip() for c in r)]
    if not rows:
        raise BuildError("%s has no data rows." % os.path.basename(path))
    skipped = None
    if _looks_like_header(rows[0]):
        skipped = rows[0][:3]
        rows = rows[1:]
    if not rows:
        raise BuildError(
            "%s has a heading row but no data under it." % os.path.basename(path))

    out = []
    for r in rows:
        vals = [(r[i].strip() if i < len(r) and r[i] else "") for i in range(3)]
        out.append(tuple(vals))
    return out, skipped


def _read_xlsx_rows(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
                  for si in ET.fromstring(
                      z.read("xl/sharedStrings.xml")).findall(NS + "si")]
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"',
                           z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))
    m = re.search(r'<sheet[^>]*?r:id="([^"]+)"', wb)
    if not m:
        raise BuildError("%s has no worksheets." % os.path.basename(path))
    target = rels[m.group(1)].lstrip("/")
    part = target if target.startswith("xl/") else "xl/" + target

    grid = {}
    for row in ET.fromstring(z.read(part)).find(NS + "sheetData"):
        r = int(row.get("r"))
        for c in row.findall(NS + "c"):
            col = col_to_idx(re.match(r"([A-Z]+)", c.get("r")).group(1))
            if col > 3:
                continue
            v, istag = c.find(NS + "v"), c.find(NS + "is")
            if istag is not None:
                txt = "".join(x.text or "" for x in istag.iter(NS + "t"))
            elif c.get("t") == "s" and v is not None:
                txt = shared[int(v.text)]
            elif c.find(NS + "f") is not None and v is not None:
                txt = v.text          # a formula's cached result
            else:
                txt = v.text if v is not None else ""
            grid.setdefault(r, {})[col] = (txt or "").strip()
    return [[grid[r].get(i, "") for i in (1, 2, 3)] for r in sorted(grid)]


def shift_formula_row(formula, new_row):
    """Rewrite a row-4 formula for a row further down, as filling down would.

    Only bare references are moved: quoted strings and sheet-qualified lookup
    ranges such as `role!$A$2:$A$1002` must stay exactly as they are, or the
    lookups would walk down the vocabulary sheets alongside the data.
    """
    out = []
    i, n = 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(formula[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            j = i + 1
            while j < n and formula[j] != "'":
                j += 1
            k = j + 1
            if k < n and formula[k] == "!":
                k += 1
                while k < n and (formula[k].isalnum() or formula[k] in "$:."):
                    k += 1
            out.append(formula[i:k])
            i = k
            continue
        if ch.isalpha() or ch == "_" or ch == "$":
            j = i
            if formula[j] == "$":
                j += 1
            while j < n and (formula[j].isalnum() or formula[j] in "_.$"):
                j += 1
            tok = formula[i:j]
            if j < n and formula[j] == "!":
                k = j + 1
                while k < n and (formula[k].isalnum() or formula[k] in "$:."):
                    k += 1
                out.append(formula[i:k])
                i = k
                continue
            m = REF_RE.match(tok)
            if m and m.group(4) == str(FORMULA_ROW) and not m.group(3):
                out.append("%s%s%d" % (m.group(1), m.group(2), new_row))
            else:
                out.append(tok)
            i = j
            continue
        out.append(ch)
        i += 1
    return "".join(out)


# --------------------------------------------------------------------------
# XML escaping
# --------------------------------------------------------------------------

def esc_text(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def esc_attr(t):
    return esc_text(t).replace('"', "&quot;")


# --------------------------------------------------------------------------
# template parsing
# --------------------------------------------------------------------------

class Template(object):
    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.shared = self._shared_strings()
        self._parse_sheet()
        self._check_fingerprint()

    def _shared_strings(self):
        out = []
        if "xl/sharedStrings.xml" not in self.zf.namelist():
            return out
        root = ET.fromstring(self.zf.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            out.append("".join(t.text or "" for t in si.iter(NS + "t")))
        return out

    def _cell_text(self, c):
        t = c.get("t")
        v = c.find(NS + "v")
        istag = c.find(NS + "is")
        if t == "s" and v is not None:
            return self.shared[int(v.text)]
        if istag is not None:
            return "".join(x.text or "" for x in istag.iter(NS + "t"))
        if v is not None:
            return v.text
        return None

    def _parse_sheet(self):
        raw = self.zf.read(SHEET_PART)
        root = ET.fromstring(raw)
        self.sheet_root = root

        # verbatim leading elements
        self.pre_xml = b""
        for tag in ("sheetPr", "sheetViews", "sheetFormatPr"):
            el = root.find(NS + tag)
            if el is not None:
                if tag == "sheetViews":
                    el = home_the_cursor(el)
                self.pre_xml += ET.tostring(strip_ns_attrs(el), encoding="utf-8")

        el = root.find(NS + "drawing")
        self.drawing_xml = (ET.tostring(strip_ns_attrs(el), encoding="utf-8")
                            if el is not None else b"")

        # column widths, expanded per column index
        self.widths = {}
        cols = root.find(NS + "cols")
        if cols is not None:
            for c in cols:
                lo, hi = int(c.get("min")), int(c.get("max"))
                attrs = {k: v for k, v in plain_attrs(c).items()
                         if k not in ("min", "max")}
                for i in range(lo, hi + 1):
                    self.widths[i] = attrs

        # cells by row -> col letter
        self.rows = {}
        sd = root.find(NS + "sheetData")
        self.max_row = 0
        for row in sd.findall(NS + "row"):
            r = int(row.get("r"))
            self.max_row = max(self.max_row, r)
            if r not in HEADER_ROWS and r != FORMULA_ROW and r != BLANK_ROW:
                continue
            d = {}
            for c in row.findall(NS + "c"):
                ref = c.get("r")
                col = re.match(r"([A-Z]+)", ref).group(1)
                f = c.find(NS + "f")
                d[col] = dict(
                    style=c.get("s"),
                    text=self._cell_text(c),
                    formula=f.text if f is not None else None,
                )
            self.rows[r] = d

        self.row_attrs = {}
        for row in sd.findall(NS + "row"):
            r = int(row.get("r"))
            if r in HEADER_ROWS or r == FORMULA_ROW or r == BLANK_ROW:
                self.row_attrs[r] = {
                    k: v for k, v in plain_attrs(row).items()
                    if k not in ("r", "spans")
                }

        # Data validations, keyed by the single column each one covers. They
        # live in two places: the plain <dataValidations> block, and - once
        # Excel has saved the file - an <extLst> for any list sourced from
        # another sheet. Reading only the first cost six vocabulary dropdowns
        # (contributor, role, resource type, genre, language, subject) the
        # moment the template was opened and saved.
        self.validations = {}
        self.dv_from_ext = set()

        def record(col, attrs, formula):
            self.validations[col] = (attrs, formula)

        dv = root.find(NS + "dataValidations")
        if dv is not None:
            for d in dv:
                f1 = d.find(NS + "formula1")
                attrs = {k: v for k, v in plain_attrs(d).items() if k != "sqref"}
                for ref in (d.get("sqref") or "").split():
                    col = re.match(r"([A-Z]+)", ref).group(1)
                    record(col, attrs, f1.text if f1 is not None else None)

        for d in root.iter(X14 + "dataValidation"):
            attrs = {k: v for k, v in plain_attrs(d).items() if k != "sqref"}
            f1 = d.find(X14 + "formula1")
            formula = None
            if f1 is not None:
                inner = f1.find(XM + "f")
                formula = (inner.text if inner is not None else f1.text)
            sq = d.find(XM + "sqref")
            for ref in ((sq.text if sq is not None else "") or "").split():
                col = re.match(r"([A-Z]+)", ref).group(1)
                # The plain block wins if a column somehow appears in both.
                if col not in self.validations:
                    record(col, attrs, formula)
                    self.dv_from_ext.add(col)

    def header(self, col):
        return self.rows.get(3, {}).get(col, {}).get("text")

    def column_of(self, header):
        """The column carrying a given row-3 header, or None.

        Needed for source columns that sit outside their block's range.
        """
        for col, cell in self.rows.get(3, {}).items():
            if (cell.get("text") or "").strip() == header:
                return col
        return None

    def _check_fingerprint(self):
        bad = []
        for col, expect in FINGERPRINT.items():
            got = self.header(col)
            if (got or "").strip() != expect:
                bad.append("  %s: expected %r, found %r" % (col, expect, got))
        if bad:
            raise BuildError(
                "The template's metadata sheet does not match the expected layout:\n"
                + "\n".join(bad)
                + "\n\nThe column ranges in BLOCK_DEFS were derived from the "
                  "original template. Recheck them against the new template before "
                  "building, otherwise field sets will be copied from the wrong "
                  "columns."
            )

    def numbering(self, block):
        """The (lowest, count) of numbers a block uses in its numbering token.

        Most blocks use one number (`note1`), so each instance steps by one.
        A form field set uses several - `form` holds form1..form4 and
        `technicalDetails` form5..form8 - so an instance consumes four numbers
        and the next one has to start past all of them. The lowest number
        matters for technicalDetails, whose template headers start at form5
        but which must renumber from form1 when it is the only form field set.
        """
        seg = block["number_seg"]
        if seg is None:
            return None
        cols = block_columns(block)
        child = block.get("child")
        if child:
            # A nested child's columns still carry the parent's numbering in
            # segment 0 - form8.note1.value sits outside the parent's own
            # column range but is still a form8 - so they count towards how
            # many numbers an instance consumes.
            cols = cols + block_columns(child)
        nums = []
        for col in cols:
            h = self.header(col)
            if not h:
                continue
            parts = h.split(".")
            if len(parts) <= seg:
                continue
            m = re.match(r"^([A-Za-z]+)(\d+)", parts[seg])
            if m:
                nums.append(int(m.group(2)))
        if not nums:
            return (1, 1)
        return (min(nums), max(nums) - min(nums) + 1)


def block_columns(block):
    return [idx_to_col(i)
            for i in range(col_to_idx(block["first"]), col_to_idx(block["last"]) + 1)]


def instance_columns(tpl, block, number):
    """The source columns a given instance of a block emits.

    Instances are not identical, in both directions:

    * `first_only` columns appear on the first instance only. `status: primary`
      designates *the* primary title, contributor or date, so copying it into
      every instance makes each claim primacy - a 3-contributor sheet produced
      three `primary` values, ambiguous rather than merely redundant. Absence
      means "not primary".
    * `repeat_only` columns appear on instances 2+ only. A title's `type`
      ("alternative") describes how a title relates to the primary one, so it
      is meaningless on the primary title itself and required on the others.

    Both are matched on the template's own header text, never a column letter,
    since letters shift whenever a column is added or removed. A `repeat_only`
    source column may also live outside the block's own range - it is appended
    at the far right of the template, which avoids renumbering the sheet - so
    it is resolved by header lookup rather than by position.
    """
    skip = set(block.get("first_only", ()))
    cols = []
    for col in block_columns(block):
        if number > 1 and skip and (tpl.header(col) or "").strip() in skip:
            continue
        cols.append(col)
    if number > 1:
        for header in block.get("repeat_only", ()):
            col = tpl.column_of(header)
            if col is None:
                raise BuildError(
                    "The %s field set expects a %r column for repeated "
                    "instances, but the template has no such header. It is appended past "
                    "the end of the sheet; if the template was rebuilt it may "
                    "have been dropped." % (block["key"], header))
            cols.append(col)
    return cols


# --------------------------------------------------------------------------
# renumbering
# --------------------------------------------------------------------------

def bump_seg(seg, delta):
    """Shift the number on a leading name token by `delta`.

    build_plan() works out each instance's delta. For most blocks it is
    whole block-widths - the second note is note1 + 1 - but a form field set
    steps past every number the form field sets before it used, and
    technicalDetails can shift down (form5 -> form1) when it stands alone.
    """
    m = re.match(r"^([A-Za-z]+)(\d+)(.*)$", seg)
    if not m:
        return seg
    new = int(m.group(2)) + delta
    return "%s%d%s" % (m.group(1), new, m.group(3))


def renumber_header(header, block, delta, parent_delta=None):
    if header is None or block["number_seg"] is None:
        return header
    parts = header.split(".")
    seg = block["number_seg"]
    if seg >= len(parts):
        return header
    parts[seg] = bump_seg(parts[seg], delta)
    if parent_delta is not None and seg > 0:
        parts[0] = bump_seg(parts[0], parent_delta)
    return ".".join(parts)


FILL_DOWN = "DO NOT EDIT - FILL DOWN"


def renumber_label(text, number, parent_number=None, number_missing=False):
    """Renumber a row 1 block label or a row 2 field label.

    A label already carrying `#1` always gets the instance number substituted.
    Whether a label *without* a number gains one differs by row, which is why
    the caller passes `number_missing`:

    * Row 1 holds block labels, and every label in a repeatable block is
      numbered so a reader can tell which instance a run of columns belongs to.
      The template only bothered to number the blocks it expected to repeat, so
      `Form/Genre`, `Origin info`, `Geographic data` and the geographic block's
      interior labels (`Point coordinates (decimal)`) all get their number
      added here. Only blocks that cannot repeat stay plain - `Access
      information` and `Administrative metadata`, where a `#1` would imply an
      impossible `#2`. The caller decides via `number_missing`.
    * Row 2 holds field labels - `Note`, `Type`, `Display label` - which name
      the field rather than the instance, so nothing is added. The one row-2
      label the template does number, `Subject #1 part 1 value`, still
      renumbers because the number is already there.

    The `DO NOT EDIT - FILL DOWN` marker is a note to the cataloguer, not a
    block label, so it is never touched.
    """
    if text is None or text.strip() == FILL_DOWN:
        return text
    hashes = re.findall(r"#\d+", text)
    if not hashes:
        return "%s #%d" % (text, number) if number_missing else text
    if parent_number is not None and len(hashes) >= 2:
        # e.g. "Contributor #1 Role #1" - first number is the parent's
        done = {"n": 0}

        def sub(m):
            done["n"] += 1
            return "#%d" % (parent_number if done["n"] == 1 else number)
        return re.sub(r"#\d+", sub, text)
    return re.sub(r"#\d+", "#%d" % number, text)


# --------------------------------------------------------------------------
# formula rewriting
# --------------------------------------------------------------------------

REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")


def rewrite_formula(formula, colmap, where):
    """Remap bare cell references onto the block instance's new columns.

    Skips string literals and sheet-qualified references such as
    `contributor!A:F`, which point at the lookup sheets and must not move.
    """
    out = []
    i, n = 0, len(formula)
    while i < n:
        ch = formula[i]

        if ch == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(formula[i:j + 1])
            i = j + 1
            continue

        if ch == "'":
            j = i + 1
            while j < n and formula[j] != "'":
                j += 1
            k = j + 1
            if k < n and formula[k] == "!":
                k += 1
                while k < n and (formula[k].isalnum() or formula[k] in "$:."):
                    k += 1
            out.append(formula[i:k])
            i = k
            continue

        if ch.isalpha() or ch == "_" or ch == "$":
            j = i
            if formula[j] == "$":
                j += 1
            while j < n and (formula[j].isalnum() or formula[j] in "_.$"):
                j += 1
            tok = formula[i:j]

            if j < n and formula[j] == "!":
                k = j + 1
                while k < n and (formula[k].isalnum() or formula[k] in "$:."):
                    k += 1
                out.append(formula[i:k])
                i = k
                continue

            m = REF_RE.match(tok)
            if m:
                src = m.group(2).upper()
                if src not in colmap:
                    raise BuildError(
                        "Formula in %s references column %s, which is not part of "
                        "that field set instance. The column ranges in BLOCK_DEFS "
                        "may be wrong.\n"
                        "  formula: %s" % (where, src, formula)
                    )
                out.append("%s%s%s%s" % (m.group(1), colmap[src], m.group(3), m.group(4)))
            else:
                out.append(tok)
            i = j
            continue

        out.append(ch)
        i += 1

    return "".join(out)


# --------------------------------------------------------------------------
# spec -> column plan
# --------------------------------------------------------------------------

class Instance(object):
    """One emitted block instance and the columns it occupies."""

    def __init__(self, block, number, parent_number=None,
                 delta=0, parent_delta=None):
        self.block = block
        self.number = number            # instance number, used in labels
        self.parent_number = parent_number
        self.delta = delta              # header renumbering offset
        self.parent_delta = parent_delta
        self.parent_inst = None   # set for nested instances (role, form note)
        self.cols = []            # (src_col, out_col)

    def renumber(self, header):
        return renumber_header(header, self.block, self.delta,
                               self.parent_delta)


def normalize_spec(spec):
    blocks = spec.get("blocks")
    if not isinstance(blocks, list) or not blocks:
        raise BuildError("The spec needs a non-empty \"blocks\" list.")

    plan = []
    seen = {}
    for entry in blocks:
        if isinstance(entry, str):
            entry = {"type": entry}
        if not isinstance(entry, dict) or "type" not in entry:
            raise BuildError("Each entry in \"blocks\" needs a \"type\": %r" % (entry,))

        raw = str(entry["type"]).strip()
        key = raw if raw in BLOCK_BY_KEY else ALIASES.get(raw.lower())
        if key is None:
            raise BuildError(
                "Unknown field set type %r. Valid types: %s"
                % (raw, ", ".join(b["key"] for b in BLOCK_DEFS))
            )

        block = BLOCK_BY_KEY[key]
        count = entry.get("count", 1)
        try:
            count = int(count)
        except (TypeError, ValueError):
            raise BuildError("Count for %s must be a whole number, got %r" % (key, count))
        if count < 1:
            raise BuildError(
                "Count for %s must be at least 1. To leave a field set out, omit "
                "it from the spec entirely." % key
            )
        if not block["repeatable"] and count != 1:
            raise BuildError(
                "The %s field set cannot repeat - it appears at most once, and "
                "its headers carry no instance number." % key
            )
        if key in seen:
            raise BuildError(
                "%s is listed twice. Give it a single entry with \"count\": %d instead."
                % (key, seen[key] + count)
            )
        seen[key] = count

        # A block may nest a repeatable child: roles inside a contributor,
        # notes inside a form. The spec names the child counts by its plural -
        # "roles", "notes" - one number per parent instance.
        child = block.get("child")
        child_field = CHILD_FIELD(child) if child else None
        counts = entry.get(child_field) if child_field else None

        for other in ALL_CHILD_FIELDS:
            if other != child_field and other in entry:
                owner = next(b["key"] for b in BLOCK_DEFS
                             if "child" in b and CHILD_FIELD(b["child"]) == other)
                raise BuildError(
                    "%r only applies to the %s field set, not %s."
                    % (other, owner, key))

        if child:
            if counts is None:
                counts = [1] * count
            if isinstance(counts, int):
                counts = [counts] * count
            if len(counts) != count:
                raise BuildError(
                    "\"%s\" for %s has %d entries but count is %d. Give one %s "
                    "count per %s."
                    % (child_field, key, len(counts), count, child["key"], key)
                )
            counts = [int(c) for c in counts]
            if any(c < 1 for c in counts):
                raise BuildError("Every %s needs at least 1 %s."
                                 % (key, child["key"]))

        plan.append((block, count, counts))

    for group in EXCLUSIVE_GROUPS:
        both = [k for k in group if k in seen]
        if len(both) > 1:
            raise BuildError(
                "A workbook can carry %s or %s, but not both - they both number "
                "their headers from \"%s1\", so a sheet with both would repeat "
                "every one of those headers. Ask which is wanted and build that "
                "one: %s for a single heading per subject, %s for a structured "
                "heading in two parts."
                % (both[0], both[1], group[0], group[0], group[1])
            )

    for block in BLOCK_DEFS:
        if block.get("required") and block["key"] not in seen:
            raise BuildError(
                "The %s field set is required and must appear in the spec. Every "
                "object needs a title, and the adminMetadata formulas key off the "
                "title field set's \"Main title\" entry cell, so the sheet does not "
                "work without it. This is about the columns, not the content - the "
                "cells ship empty and the title is typed in when the metadata is "
                "created."
                % block["key"]
            )

    # adminMetadata records who catalogued the object, which every SDR record
    # carries, so it goes in whether or not the caller thought to ask for it.
    for block in BLOCK_DEFS:
        if block.get("always") and block["key"] not in seen:
            plan.append((block, 1, None))

    # It is record provenance, so it belongs at the far right regardless of
    # where the caller listed it.
    plan.sort(key=lambda t: 1 if t[0]["key"] == "adminMetadata" else 0)
    return plan


def build_plan(tpl, spec):
    plan = normalize_spec(spec)
    instances = []
    out_idx = 1

    # fixed prefix: druid / source_id / purl
    prefix = Instance(dict(key="_prefix", number_seg=None, first=PREFIX_FIRST,
                           last=PREFIX_LAST, repeatable=False), 1)
    for i in range(col_to_idx(PREFIX_FIRST), col_to_idx(PREFIX_LAST) + 1):
        prefix.cols.append((idx_to_col(i), idx_to_col(out_idx)))
        out_idx += 1
    instances.append(prefix)

    # Numbers already used per numbering group. A block is its own group
    # unless it names one; the two form field sets share "form", so whichever
    # instances are emitted number form1, form2, ... consecutively in column
    # order, with no gap where a field set was left out.
    used = {}
    for block, count, child_counts in plan:
        lo, width = tpl.numbering(block) or (1, 1)
        group = block.get("number_group", block["key"])
        child = block.get("child")
        for n in range(1, count + 1):
            delta = used.get(group, 0) - (lo - 1)
            used[group] = used.get(group, 0) + width
            inst = Instance(block, n, delta=delta)
            for col in instance_columns(tpl, block, n):
                inst.cols.append((col, idx_to_col(out_idx)))
                out_idx += 1
            instances.append(inst)

            if child:
                child_width = (tpl.numbering(child) or (1, 1))[1]
                for m in range(1, child_counts[n - 1] + 1):
                    sub = Instance(child, m, parent_number=n,
                                   delta=(m - 1) * child_width,
                                   parent_delta=delta)
                    sub.parent_inst = inst
                    for col in instance_columns(tpl, child, m):
                        sub.cols.append((col, idx_to_col(out_idx)))
                        out_idx += 1
                    instances.append(sub)

    total = out_idx - 1
    if total > 16384:
        raise BuildError(
            "That spec needs %d columns; Excel's limit is 16384. Reduce the counts."
            % total
        )
    return instances, total


# --------------------------------------------------------------------------
# sheet generation
# --------------------------------------------------------------------------

def generate_sheet(tpl, instances, total_cols, data_rows, seed=None):
    # per-instance column maps, used for both formulas and renumbering
    for inst in instances:
        inst.map = {src: out for src, out in inst.cols}

    # adminMetadata keys off the first title instance's "Main title" cell;
    # normalize_spec guarantees a title block exists.
    title_inst = next(i for i in instances
                      if i.block["key"] == "title" and i.number == 1)
    admin_trigger = title_inst.map[ADMIN_TRIGGER_SRC]

    # A nested instance absorbs its parent's columns, so a formula spanning the
    # two - a role reading its contributor's name cell, say - still resolves.
    # Keyed off the recorded parent instance rather than searching by block
    # name, which only worked while role/contributor was the single case.
    for inst in instances:
        if inst.parent_inst is not None:
            merged = dict(inst.parent_inst.map)
            merged.update(inst.map)
            inst.map = merged

    parts = []
    parts.append(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    parts.append(
        ('<worksheet xmlns="%s" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
         % MAIN).encode("utf-8"))
    parts.append(tpl.pre_xml)

    # --- cols
    cols_xml = []
    for inst in instances:
        for src, out in inst.cols:
            attrs = tpl.widths.get(col_to_idx(src))
            if not attrs:
                continue
            a = " ".join('%s="%s"' % (k, esc_attr(v)) for k, v in sorted(attrs.items()))
            i = col_to_idx(out)
            cols_xml.append('<col min="%d" max="%d" %s/>' % (i, i, a))
    if cols_xml:
        parts.append(("<cols>%s</cols>" % "".join(cols_xml)).encode("utf-8"))

    # --- sheetData
    parts.append(b"<sheetData>")

    def row_open(r):
        a = "".join(' %s="%s"' % (k, esc_attr(v))
                    for k, v in sorted(tpl.row_attrs.get(r, {}).items()))
        return '<row r="%d" spans="1:%d"%s>' % (r, total_cols, a)

    def cell(out_col, r, style, text=None, formula=None):
        ref = "%s%d" % (out_col, r)
        s = ' s="%s"' % style if style else ""
        if formula is not None:
            return '<c r="%s"%s><f>%s</f></c>' % (ref, s, esc_text(formula))
        if text is None or text == "":
            return '<c r="%s"%s/>' % (ref, s)
        return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                % (ref, s, esc_text(text)))

    # header rows
    for r in HEADER_ROWS:
        src_row = tpl.rows.get(r, {})
        buf = [row_open(r)]
        for inst in instances:
            block = inst.block
            for src, out in inst.cols:
                c = src_row.get(src)
                if c is None:
                    continue
                text = c["text"]
                if r == 3:
                    text = inst.renumber(text)
                else:
                    # Every row 1 label in a repeatable block is numbered,
                    # including ones interior to the block such as the
                    # geographic block's `Point coordinates (decimal)`, so each
                    # label says which instance it belongs to. Blocks that
                    # cannot repeat stay plain - numbering `Access information`
                    # would imply a #2 that cannot exist.
                    text = renumber_label(
                        text, inst.number, inst.parent_number,
                        number_missing=(r == 1
                                        and block.get("repeatable", False)))
                buf.append(cell(out, r, c["style"], text=text))
        buf.append("</row>")
        parts.append("".join(buf).encode("utf-8"))

    # where seeded values go: druid and source_id are the fixed leading
    # columns; the title goes in the first title block's entry cell, which is
    # the one column in that block with no header of its own in row 3
    seed_cols = None
    if seed:
        prefix = next(i for i in instances if i.block["key"] == "_prefix")
        entry = [out for src, out in title_inst.cols
                 if not (tpl.header(src) or "").strip()]
        if len(entry) != 1:
            raise BuildError(
                "Expected exactly one title column with no row-3 header to put "
                "the title in; found %d (%s). The template's title field set has "
                "changed shape." % (len(entry), ", ".join(entry) or "none"))
        seed_cols = {"druid": prefix.map["A"],
                     "source_id": prefix.map["B"],
                     "title": entry[0]}

    # formula row, plus one row per seeded object with the formulas filled down,
    # plus one spare row past the last object: it carries the formulas and
    # dropdowns but no druid, source id or title, so it is the row to fill down
    # from when more objects are added. Filling down from the last object's row
    # would drag that object's druid, source id and title along with the
    # formulas. Without a seed there is nothing to be past, and row 4 is
    # already that row.
    src_row = tpl.rows.get(FORMULA_ROW, {})
    blank_src = tpl.rows.get(BLANK_ROW, {})
    n_seeded = len(seed) if seed else 0
    spare = 1 if n_seeded else 0
    last_data_row = FORMULA_ROW + max(n_seeded - 1, 0) + spare

    for r in range(FORMULA_ROW, last_data_row + 1):
        idx = r - FORMULA_ROW
        record = seed[idx] if idx < n_seeded else None
        buf = [row_open(FORMULA_ROW).replace('r="%d"' % FORMULA_ROW,
                                             'r="%d"' % r, 1)]
        for inst in instances:
            colmap = dict(inst.map)
            if inst.block["key"] == "adminMetadata":
                colmap[ADMIN_TRIGGER_SRC] = admin_trigger
            for src, out in inst.cols:
                c = src_row.get(src)
                value = None
                if record and seed_cols:
                    if out == seed_cols["druid"]:
                        value = record[0]
                    elif out == seed_cols["source_id"]:
                        value = record[1]
                    elif out == seed_cols["title"]:
                        value = record[2]
                if value:
                    # These three are entry columns, and the template carries
                    # no cell for them at all - they rely on the column
                    # default - so a seeded value has to create the cell,
                    # borrowing a style from the row below if there is one.
                    style = ((c or {}).get("style")
                             or (blank_src.get(src) or {}).get("style"))
                    buf.append(cell(out, r, style, text=value))
                    continue
                if c is None:
                    continue
                if c["formula"]:
                    where = "%s instance %s (column %s)" % (
                        inst.block["key"], inst.number, src)
                    f = rewrite_formula(c["formula"], colmap, where)
                    if r != FORMULA_ROW:
                        f = shift_formula_row(f, r)
                    buf.append(cell(out, r, c["style"], formula=f))
                else:
                    buf.append(cell(out, r, c["style"],
                                    text=c["text"] if r == FORMULA_ROW else None))
        buf.append("</row>")
        parts.append("".join(buf).encode("utf-8"))

    # blank fill-down rows below whatever was populated
    if blank_src and data_rows > 0:
        blanks = [(out, blank_src[src]["style"])
                  for inst in instances for src, out in inst.cols
                  if src in blank_src]
        first_blank = last_data_row + 1
        for r in range(first_blank, first_blank + data_rows):
            buf = [row_open(BLANK_ROW).replace('r="%d"' % BLANK_ROW,
                                               'r="%d"' % r, 1)]
            for out, style in blanks:
                buf.append('<c r="%s%d"%s/>'
                           % (out, r, ' s="%s"' % style if style else ""))
            buf.append("</row>")
            parts.append("".join(buf).encode("utf-8"))

    parts.append(b"</sheetData>")

    # --- dataValidations, spanning every populated row so the dropdowns are
    # available on each one rather than only the first
    dvs = []
    for inst in instances:
        for src, out in inst.cols:
            if src not in tpl.validations:
                continue
            attrs, f1 = tpl.validations[src]
            a = " ".join('%s="%s"' % (k, esc_attr(v)) for k, v in sorted(attrs.items()))
            body = "<formula1>%s</formula1>" % esc_text(f1) if f1 is not None else ""
            sqref = ("%s%d" % (out, FORMULA_ROW) if last_data_row == FORMULA_ROW
                     else "%s%d:%s%d" % (out, FORMULA_ROW, out, last_data_row))
            dvs.append('<dataValidation %s sqref="%s">%s</dataValidation>'
                       % (a, sqref, body))
    if dvs:
        parts.append(('<dataValidations count="%d">%s</dataValidations>'
                      % (len(dvs), "".join(dvs))).encode("utf-8"))

    parts.append(tpl.drawing_xml)
    parts.append(b"</worksheet>")
    return b"".join(parts)


# --------------------------------------------------------------------------
# workbook assembly
# --------------------------------------------------------------------------

def strip_rels(raw):
    """Drop relationships to parts the output does not carry."""
    txt = raw.decode("utf-8")
    kept = [m.group(0) for m in re.finditer(r"<Relationship\b[^>]*/>", txt)
            if not any(h.lower() in m.group(0).lower()
                       for h in COMMENT_REL_HINTS + DROP_REL_HINTS)]
    head = txt[:txt.index(">", txt.index("<Relationships")) + 1]
    return (head + "".join(kept) + "</Relationships>").encode("utf-8")


def strip_legacy_drawing(raw):
    """Remove <legacyDrawing>, which points at a comment's VML anchor part."""
    txt = raw.decode("utf-8")
    txt = re.sub(r"<legacyDrawing\b[^>]*/>", "", txt)
    txt = re.sub(r"<legacyDrawing\b[^>]*>.*?</legacyDrawing>", "", txt, flags=re.S)
    return txt.encode("utf-8")


def strip_content_types(raw, dropped):
    """Remove overrides for the parts we dropped.

    Attribute order in [Content_Types].xml is not guaranteed, so match
    PartName wherever it sits inside the tag.
    """
    txt = raw.decode("utf-8")
    for part in dropped:
        txt = re.sub(
            r'<Override\b[^>]*PartName="/%s"[^>]*/>' % re.escape(part), "", txt)
    return txt.encode("utf-8")


def force_recalc(raw):
    """Formulas are written without cached values, so ask Excel to compute
    them when the file opens."""
    txt = raw.decode("utf-8")
    if "<calcPr" in txt:
        txt = re.sub(r"<calcPr[^>]*/>", '<calcPr fullCalcOnLoad="1"/>', txt)
    else:
        txt = txt.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    return txt.encode("utf-8")


def write_workbook(tpl, sheet_xml, out_path):
    src = tpl.zf
    dropped = [n for n in src.namelist() if is_dropped_part(n)]
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            name = item.filename
            if name in dropped:
                continue
            if name == SHEET_PART:
                out.writestr(name, sheet_xml)
            elif name.endswith(".rels"):
                out.writestr(name, strip_rels(src.read(name)))
            elif re.match(r"^xl/worksheets/sheet\d+\.xml$", name):
                out.writestr(name, strip_legacy_drawing(src.read(name)))
            elif name == "[Content_Types].xml":
                out.writestr(name, strip_content_types(src.read(name), dropped))
            elif name == "xl/workbook.xml":
                out.writestr(name, force_recalc(src.read(name)))
            else:
                out.writestr(item, src.read(name))


# --------------------------------------------------------------------------
# reporting
# --------------------------------------------------------------------------

def summarize(tpl, instances):
    lines = []
    for inst in instances:
        key = inst.block["key"]
        if key == "_prefix":
            label = "druid / source_id / purl"
        elif inst.parent_inst is not None:
            label = "%s %d %s %d" % (inst.parent_inst.block["key"],
                                     inst.parent_number, key, inst.number)
        else:
            label = "%s %d" % (key, inst.number)
        first, last = inst.cols[0][1], inst.cols[-1][1]
        h = [tpl.header(s) for s, _ in inst.cols]
        sample = next((x for x in h if x), "")
        if sample:
            sample = inst.renumber(sample)
        lines.append("  %-26s %s:%s  (%d cols)%s"
                     % (label, first, last, len(inst.cols),
                        "  e.g. " + sample if sample else ""))
    return "\n".join(lines)


def list_blocks(tpl):
    print("Field sets available in the bundled template:\n")
    print("  %-16s %-9s %-8s %s" % ("type", "columns", "repeats", "first header"))
    print("  " + "-" * 62)
    for b in BLOCK_DEFS:
        rep = "yes" if b["repeatable"] else "no"
        if b.get("required"):
            rep += " (req)"
        elif b.get("always"):
            rep += " (auto)"
        print("  %-16s %-9s %-8s %s"
              % (b["key"], "%s-%s" % (b["first"], b["last"]), rep,
                 tpl.header(b["first"]) or "(entry column)"))
        if "child" in b:
            c = b["child"]
            print("  %-16s %-9s %-8s %s"
                  % ("  " + c["key"], "%s-%s" % (c["first"], c["last"]),
                     "nested", tpl.header(c["first"]) or ""))
    print("\n  form and technicalDetails each hold four form entries and number them")
    print("  consecutively across every instance of either, in column order.")
    print("  Access and adminMetadata appear at most once.")
    print("  Title is required: adminMetadata's formulas key off its Main title cell.")
    for group in EXCLUSIVE_GROUPS:
        print("  %s and %s are alternatives: a workbook carries one or the other."
              % (group[0], group[1]))
    print("  adminMetadata is added automatically and always sits last.")


# --------------------------------------------------------------------------

def main(argv=None):
    here = os.path.dirname(os.path.abspath(__file__))
    default_tpl = os.path.join(here, "..", "assets", "cocina_spreadsheet.xlsx")

    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--spec", help="JSON file describing the blocks to emit")
    ap.add_argument("--output", help="path for the generated .xlsx")
    ap.add_argument("--template", default=os.path.normpath(default_tpl))
    ap.add_argument("--data-rows", type=int, default=None,
                    help="blank fill-down rows below the populated rows "
                         "(default: match the template)")
    ap.add_argument("--data", default=None,
                    help="optional .csv/.tsv/.xlsx whose first three columns "
                         "are druid, source id and title; one row per object. "
                         "Each row is written out with the formulas and "
                         "dropdowns filled down to match.")
    ap.add_argument("--list-blocks", action="store_true")
    args = ap.parse_args(argv)

    if not os.path.exists(args.template):
        raise BuildError("Template not found: %s" % args.template)
    tpl = Template(args.template)

    if args.list_blocks:
        list_blocks(tpl)
        return 0

    if not args.spec or not args.output:
        ap.error("--spec and --output are both required (or use --list-blocks)")

    # explicit utf-8: the Windows default is the ANSI codepage, which would
    # misread a spec containing non-ASCII text
    with open(args.spec, encoding="utf-8") as fh:
        spec = json.load(fh)

    data_rows = args.data_rows
    if data_rows is None:
        data_rows = spec.get("data_rows")
    if data_rows is None:
        data_rows = max(0, tpl.max_row - BLANK_ROW + 1)

    seed, skipped_header = None, None
    data_path = args.data or spec.get("data")
    if data_path:
        if not os.path.exists(data_path):
            raise BuildError("Data file not found: %s" % data_path)
        seed, skipped_header = read_seed_data(data_path)

    instances, total = build_plan(tpl, spec)
    sheet_xml = generate_sheet(tpl, instances, total, data_rows, seed)
    write_workbook(tpl, sheet_xml, args.output)

    print("Wrote %s" % args.output)
    if seed:
        last = FORMULA_ROW + len(seed) - 1
        spare_row = last + 1
        print("  %d columns (A:%s), %d header rows, %d populated rows "
              "(%d-%d), 1 spare formatted row (%d), %d blank rows below"
              % (total, idx_to_col(total), len(HEADER_ROWS), len(seed),
                 FORMULA_ROW, last, spare_row, data_rows))
        if skipped_header:
            print("  skipped the data file's heading row: %s"
                  % " | ".join(skipped_header))
        # Name the source id, not just the sheet row: the sheet row is the one
        # number that cannot be looked up in the data file, so reporting it
        # alone leaves the reader to undo the header offset by hand.
        blank = [(n, rec) for n, rec in enumerate(seed, FORMULA_ROW)
                 if not rec[2]]
        if blank:
            print("  note: %d row(s) have no title:" % len(blank))
            for n, rec in blank[:8]:
                ident = rec[1] or rec[0] or "(no druid or source id either)"
                print("    row %d  %s" % (n, ident))
            if len(blank) > 8:
                print("    ... and %d more" % (len(blank) - 8))
        print("  formulas and dropdowns filled down through row %d, one past "
              "the last object\n" % spare_row)
    else:
        print("  %d columns (A:%s), %d header rows + 1 formula row + "
              "%d blank rows\n"
              % (total, idx_to_col(total), len(HEADER_ROWS), data_rows))
    print(summarize(tpl, instances))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BuildError as exc:
        sys.stderr.write("error: %s\n" % exc)
        sys.exit(1)
