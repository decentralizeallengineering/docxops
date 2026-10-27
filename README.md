# docxops

Word (`.docx`) structure operations for LLM agents, including a real table of contents.

Each operation is a plain Python function that returns a JSON-serializable report and
never raises for user errors:

```json
{"ok": true, "changed": true, "warnings": [], "details": {"...": "..."}}
```

On failure you get `ok: false` with `details.error`. Each operation's arguments have a JSON
Schema (at most 600 characters) in `docxops.SCHEMAS[name]`, and the arguments are checked
against it on every call. It runs on Python 3.12+ on Windows and Linux. It never uses the
network, subprocesses or Office automation. Dependencies: `python-docx`, `docxcompose`,
`jsonschema`.

```python
import docxops

docxops.add_toc("report.docx")                         # levels="1-3", title="Contents"
docxops.call("outline", {"path": "report.docx"})       # dispatch by name (tool-use)
tools = docxops.tool_definitions()                     # [{name, description, input_schema}]
```

Operations that change the file write it back to `path`. Pass `out` to write a new file
instead. Saves are atomic: the file is written to a temp file, which then replaces the
target.

## Operations

| name | what it does |
|---|---|
| `add_toc` | Adds a real `TOC \o "1-3" \h \z \u` field with a pre-filled result (details below). |
| `outline` | Headings tree, counts (paragraphs, tables, images, sections, comments), section page setup, styles used, field counts, `has_toc`. |
| `add_heading` | Adds a heading at the end of the document, or at the end of the `after_heading` section. |
| `set_heading_numbering` | Links a multilevel list to the Heading 1..N styles (`decimal`: 1 / 1.1 / 1.1.1, or `outline`: I. / A. / 1.). `none` removes it. Running it again replaces the list instead of adding a second one. |
| `add_paragraph_after` | Adds paragraph(s) directly after a heading, or at the end of its section (`where`). |
| `replace_text` | Literal find/replace, including across runs (details below). |
| `add_table` | Builds a table from `rows`. Row 1 becomes a header: bold, repeats on each page, optional `header_fill`. Uses the table style named in `style`. Numeric cells are right-aligned. |
| `table_from_csv` | Same as `add_table`, but reads the rows from a CSV file (delimiter, encoding and `max_rows` can be set). |
| `add_caption` | Adds `Table {SEQ}: text` or `Figure {SEQ}: text`. Table captions go above, figure captions below. Every caption is renumbered in document order. An existing caption next to the target is replaced. The label and number get a `_Ref` bookmark. |
| `cross_reference` | Replaces a placeholder with a `REF` or `PAGEREF` field pointing to a caption (`"Table 2"` or caption text) or a heading. `display` can be `text`, `page` or `above_below`. |
| `header_footer` | Sets header and footer text. `{page}` and `{pages}` become `PAGE` / `NUMPAGES` fields. `page_numbers="page_x_of_y"` adds "Page X of Y". Applies to every section, or to one with `section` (1-based). |
| `section_break` | Starts a new section before a heading, with its own orientation and start type. `isolate=true` also closes the section after that heading's content. |
| `page_setup` | Paper size (A3/A4/A5/Letter/Legal), orientation, and margins in mm, cm, in or pt. |
| `compose` | Appends documents using docxcompose, which keeps their styles, numbering and images. Separator: page break, section break or none. |
| `styles_apply` | Rule-based restyling to built-in styles. Rule conditions: `regex`, `current_style`, `bold`, `max_chars`. `dry_run` and `clear_direct` are available. |
| `comments_list` | Lists comments with id, author, initials, date, text, the commented text, `reply_to`, and `resolved` (when the file records it). |

Targets are matched like this. Headings: exact text first (case- and
whitespace-insensitive), then substring, with `occurrence` for duplicates. Tables and
figures: 1-based index, and negative numbers count from the end. Sections: 1-based. When a
match is ambiguous, the first one is used and a warning says so.

## How `add_toc` works

* The field is a complex field, `begin` / `TOC \o "1-3" \h \z \u` / `separate` / … / `end`,
  inside a Table of Contents content control. That is the same structure Word writes.
* The field result is pre-filled with one paragraph per heading in range. Each paragraph
  uses the `TOC n` style (created if missing), has a right-aligned dot-leader tab at the
  text width, and holds a `w:hyperlink w:anchor="_Toc…"`. A reader sees a clickable TOC
  even if fields are never updated.
* Every heading gets a `_Toc…` bookmark. If a heading already has one, it is reused.
* Page numbers are nested `PAGEREF _Toc… \h` fields with an **empty** result. They are
  never guessed. `w:updateFields` is set in `settings.xml` so Word fills them in on open.
* Headings are found the way the `\o` and `\u` switches find them: by outline level,
  either set directly on the paragraph or inherited through the style chain (built-in
  `heading N` styles included). The title paragraph uses `TOC Heading` (outline level 9),
  so it never lists itself. If heading numbering is on, the title is explicitly
  unnumbered.
* Running it again replaces the old TOC in place. It recognises its own content control,
  any Word TOC content control, and a bare `TOC` field (along with a `TOC Heading` title
  just before it). Tables of figures (`TOC \c …`) are left alone.

## How `replace_text` stays safe

Each paragraph is read as a sequence of `w:t` pieces with their runs. When a match spans
several runs, the replacement goes into the run where the match starts. Text that was not
matched stays in its original runs with its original formatting, and runs left empty are
removed. Some matches are **reported in `details.skipped` instead of changed**:

* a match that crosses a tab or line break;
* a match that crosses a field boundary (for example a caption number), a drawing or
  another object;
* a match that runs between a hyperlink, a tracked change or a content control and the
  text outside it.

A replacement that merged runs with different formatting is counted in
`mixed_formatting`, and a warning is added.

## Known limits

* **Page numbers stay blank until a field update.** Word updates them on open because of
  `w:updateFields`. Word also shows its usual "update fields?" prompt when that flag is set.
  LibreOffice ignores the flag and shows the TOC without page numbers.
* **TOC entries leave out heading numbers.** If headings are numbered, the pre-filled
  entries show the heading text without the "2.1" label. Word adds the labels when it
  updates the field.
* **`replace_text` scope.** `scope="all"` covers the body, headers and footers. Footnotes
  and endnotes are not searched.
* **Tested against python-docx and LibreOffice only.** The output was rendered in
  LibreOffice (PDF) during development. It has not been opened in Microsoft Word here.

## Development

```bash
pip install -e .[test]
pytest
```

The tests create their own fixtures with python-docx. Licensed under Apache-2.0 (see
`LICENSE`).
