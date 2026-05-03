# Specification: generate_research_doc.py

## Task
Write a Python script using `python-docx` that generates a clean, minimal Word document formatted specifically for research reports and summaries.

## Script name
`generate_research_doc.py`

## Command line arguments
- `--title` — report title (string)
- `--content` — path to a JSON file containing the document content
- `--output-dir` — output directory path (default: `~/.openclaw/workspace/Research Reports`)
- `--toc` — optional flag to include a table of contents

## JSON content file structure
```json
{
  "title": "Report Title",
  "subtitle": "Optional subtitle",
  "author": "Optional author name",
  "date": "auto",
  "abstract": "Optional short summary paragraph shown at the top",
  "sections": [
    {
      "heading": "Section Title",
      "level": 1,
      "content": "Body text here"
    },
    {
      "heading": "Subsection",
      "level": 2,
      "content": "More text"
    }
  ],
  "references": [
    "Author et al. (2024). Title. Journal."
  ]
}
```

## Styling requirements
- Clean, minimal aesthetic — no cover page
- Title at the top, large and bold
- Optional subtitle, author, and auto-generated date beneath the title, in a smaller muted style
- Optional abstract section beneath the title block, italicised, with a horizontal rule separating it from the main body
- Optional table of contents after the title block
- Heading 1 and Heading 2 styles, clearly distinct
- Body text in Calibri, 12pt
- Consistent paragraph spacing throughout
- Optional references section at the end, formatted as a numbered list
- Page numbers in the footer

## Additional requirements
- Output filename should be auto-generated from the title with a timestamp suffix, e.g. `Mitochondrial_Dysfunction_Summary_2026-05-03.docx`
- Print the full output file path to stdout on success so OpenClaw can capture it
- Handle errors gracefully with clear error messages
- Include a `--preview` flag that prints a plain text outline of the document structure without writing a file

## Environment
Python 3, Ubuntu/Linux, `python-docx` installed via pip.
