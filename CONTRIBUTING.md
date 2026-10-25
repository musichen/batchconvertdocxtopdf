# Contributing to batchconvertdocxtopdf

## Setup

You need Python 3, Pandoc, and pdflatex (see Requirements in the [README](README.md)). The script checks both at startup and exits with an install link if either is missing.

## Test manually

```bash
mkdir -p in out
# put one small .docx file in in/, then:
python batchconvertdocx2pdf.py in out
```

Confirm a PDF with the same base name appears in `out/`. There is no automated test suite yet.

## Pull requests

Create a branch from `main`. Keep the CLI interface (two positional folder arguments) stable unless the PR explains the change, and say which file you converted in the PR.
