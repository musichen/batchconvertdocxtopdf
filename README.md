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

## Troubleshooting

- **`Pandoc is not installed`**: install Pandoc from https://pandoc.org/installing.html and run the script again. The script checks `pandoc -v` at startup and exits early without it.
- **`pdflatex is not installed`**: install a LaTeX distribution (TeX Live, MiKTeX, or `brew install mactex` on macOS). The script checks `pdflatex -v` the same way.
- **No PDFs produced**: confirm the input folder path is correct, the files end in `.docx` (lowercase), and the output folder path is writable. Files with other extensions are skipped silently.

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
