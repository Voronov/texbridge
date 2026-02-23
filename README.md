# texbridge

LaTeX to DOCX converter with Ukrainian university formatting (ДСТУ 3008:2015) and lossless round-trip support.

## What it does

Drop a `.zip` with your LaTeX project into the folder and run `make docx`. You get a properly formatted Word document with:

- All equations rendered as native Word formulas
- Figures numbered as **Рисунок 1.1 — Назва** per ДСТУ 3008:2015
- Cross-references (рис. 1.1, табл. 2.1) resolved automatically
- Times New Roman 14pt, 1.5 spacing, correct margins (30/10/20/20 mm)
- Page numbers top-right

The reverse also works — `make latex` converts the DOCX back to LaTeX with automatic restoration of equations, labels, image paths, and document structure.

## Quick start

```bash
# 1. Install dependencies
brew install pandoc pandoc-crossref
pip3 install python-docx

# 2. Place your LaTeX project as a .zip file
ls *.zip
# → DIPLOM.zip

# 3. Convert to Word
make docx
# → DIPLOM.docx

# 4. Convert back to LaTeX
make latex
# → build/latex/DIPLOM.tex
```

## Requirements

- [pandoc](https://pandoc.org/) >= 3.0
- [pandoc-crossref](https://github.com/lierdakil/pandoc-crossref)
- Python 3 with [python-docx](https://python-docx.readthedocs.io/)
- (optional) XeLaTeX for `make pdf`

## Make targets

| Command | Description |
|---------|-------------|
| `make docx` | LaTeX → DOCX with ДСТУ formatting |
| `make latex` | DOCX → LaTeX with auto-fixes |
| `make pdf` | LaTeX → PDF (requires xelatex) |
| `make check` | Full round-trip with quality report |
| `make clean` | Remove build artifacts |
| `make help` | Show all targets |

## Project structure

```
├── *.zip                  # Your LaTeX project (auto-detected)
├── Makefile               # Build system
├── pandoc-crossref.yaml   # Ukrainian figure/table numbering config
├── templates/
│   └── reference.docx     # Word template with ДСТУ 3008:2015 styles
└── scripts/
    ├── preprocess_tex.py   # Fixes LaTeX math for pandoc compatibility
    ├── fix_roundtrip.py    # Post-processor for DOCX → LaTeX round-trip
    ├── create_reference_docx.py  # Generates the reference template
    └── extract_text.py     # Text extractor for diff comparison
```

## ZIP file format

The `.zip` should contain a LaTeX project with `main.tex` at the root:

```
PROJECT.zip
├── main.tex           # Entry point (uses \input{sections/...})
├── sections/
│   ├── titlepage.tex
│   ├── vstup.tex
│   ├── section_1_1.tex
│   └── ...
├── image1.png
├── image2.jpg
└── ...
```

The output name is derived from the ZIP filename: `PROJECT.zip` → `PROJECT.docx`.

## ДСТУ 3008:2015 formatting

The generated DOCX follows Ukrainian university standards:

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman, 14pt |
| Line spacing | 1.5 |
| Margins | left 30mm, right 10mm, top/bottom 20mm |
| Paragraph indent | 1.25 cm |
| Headings | Bold, 14pt (Heading 1: centered + CAPS) |
| Figure captions | `Рисунок X.Y — Назва` (below, centered) |
| Table captions | `Таблиця X.Y — Назва` (above) |
| Page numbers | Top right, Arabic |

## Round-trip quality

Verified on a 31,000+ word thesis with 50 equations, 20 labels, and 11 images:

| Metric | Result |
|--------|--------|
| Equations preserved | 50/50 (100%) |
| Labels restored | 20/20 (100%) |
| Images mapped | 11/11 (100%) |
| Words lost | 0 |

The `make check` target runs the full round-trip and prints a comparison report.

## How it works

### LaTeX → DOCX

1. **Extract** ZIP into `src/`
2. **Pre-process** LaTeX to fix pandoc-incompatible math (Cyrillic in `\text{}`, escaped underscores, duplicate labels)
3. **Convert** with pandoc + pandoc-crossref using ДСТУ reference template
4. **Copy** result to project root

### DOCX → LaTeX

1. **Convert** DOCX to raw LaTeX with pandoc
2. **Post-process** with 11 automatic fixes:
   - Restore `\begin{equation}` from `\[...\]`
   - Restore `$...$` from `\(...\)`
   - Convert `\hyperref[label]{N}` → `\ref{label}`
   - Strip baked-in section numbers
   - Fix section hierarchy
   - Restore en-dashes
   - Fix list formatting
   - Restore `\section*` for unnumbered sections
   - Restore `\label{}` by matching captions
   - Map image paths back to originals via MD5 checksums
   - Re-add preamble with all packages
