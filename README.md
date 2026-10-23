# batchconvertdocxtopdf

Batch convert DOCX files to PDF with a small Python CLI tool, using Pandoc and pdflatex.

## Requirements

- Python 3
- [Pandoc](https://pandoc.org/installing.html)
- pdflatex, from a LaTeX distribution such as [TeX Live](https://www.tug.org/texlive/) or [MiKTeX](https://miktex.org/). On macOS: `brew install mactex`

The script checks for both `pandoc` and `pdflatex` at startup and exits with an install link if either one is missing.

## Usage

Navigate to any folder, then run:

```bash
python batchconvertdocx2pdf.py 'path_to_folder_with_docx_INPUT' 'path_to_folder_where_to_put_PDF'
```

- The first argument is the folder containing the `.docx` files to convert.
- The second argument is the folder where the generated PDFs are written.
- Every `.docx` file in the input folder is converted. The PDF keeps the same base name, for example `report.docx` becomes `report.pdf`.

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
