#!/usr/bin/env python3
# Post-processing script to fix LaTeX round-trip (LaTeX -> DOCX -> LaTeX) artifacts.
#
# Fixes:
#   1. display math -> equation environment
#   2. Strip baked-in numbers from section/subsection/subsubsection titles
#   3. hyperref -> ref
#   4. Promote specific subsections to sections (hierarchy fix)
#   5. Restore labels from original source files
#   6. Remove spurious "Table of Contents" section
#   7. Inline math notation converted back
#   8. Restore unnumbered sections for VSTUP, VYSNOVKY, etc.

import re
import sys
import os
import glob
import hashlib
import shutil


def fix_display_math(text):
    r"""Convert \[...\] to \begin{equation}...\end{equation}"""
    # Single-line: \[...\]
    text = re.sub(
        r'^\\\[(.+?)\\\]$',
        r'\\begin{equation}\n\1\n\\end{equation}',
        text,
        flags=re.MULTILINE,
    )
    # Multi-line: \[ ... \]
    text = re.sub(
        r'\\\[\s*\n(.*?)\n\s*\\\]',
        r'\\begin{equation}\n\1\n\\end{equation}',
        text,
        flags=re.DOTALL,
    )
    return text


def fix_inline_math(text):
    r"""Convert \(...\) back to $...$"""
    # Match paired \( ... \) only — use non-greedy match
    # Handle nested parens like \(f(t)\) correctly
    text = re.sub(r'\\\(((?:[^\\]|\\.)*?)\\\)', r'$\1$', text)
    return text


def fix_hyperref(text):
    r"""Convert \hyperref[label]{N} to \ref{label}"""
    text = re.sub(r'\\hyperref\[([^\]]+)\]\{[^}]*\}', r'\\ref{\1}', text)
    return text


def fix_en_dashes(text):
    """Convert -- back to Unicode en-dash – (matching original source)."""
    # Only outside of math environments and LaTeX commands
    # Replace standalone -- (not ---) with –
    text = re.sub(r'(?<!-)--(?!-)', '–', text)
    return text


def fix_list_formatting(text):
    """Fix list items split across lines by pandoc (- on separate line)."""
    # Pattern: \item\n  followed by text on next line → \item text on same line
    # In the round-tripped output, enumerate items look like:
    #   \item
    #     Text here
    # Should be: \item Text here
    text = re.sub(r'(\\item)\s*\n\s+', r'\1 ', text)
    return text


def strip_section_numbers(text):
    """Strip baked-in numbers from section titles."""
    # \section{N Title} → \section{Title}
    text = re.sub(
        r'(\\section\{)\d+\s+',
        r'\1',
        text,
    )
    # \subsection{N.N Title} → \subsection{Title}
    text = re.sub(
        r'(\\subsection\{)\d+\.\d+\s+',
        r'\1',
        text,
    )
    # \subsubsection{N.N.N. Title} and \subsubsection{N.N.N.N. Title}
    text = re.sub(
        r'(\\subsubsection\{)\d+(?:\.\d+){2,3}\.?\s*',
        r'\1',
        text,
    )
    # Also handle \texorpdfstring variants
    text = re.sub(
        r'(\\subsection\{\\texorpdfstring\{)\d+\.\d+\s+',
        r'\1',
        text,
    )
    return text


def fix_section_hierarchy(text):
    """Promote specific subsections to sections based on known titles."""
    # These subsections should be \section level
    promote_patterns = [
        r'ПРАКТИЧНА ЧАСТИНА',
        r'ПРОГРАМНА РЕАЛІЗАЦІЯ',
        r'Опис експериментального стенду',
    ]
    lines = text.split('\n')
    result = []
    for line in lines:
        promoted = False
        for pattern in promote_patterns:
            if re.search(pattern, line) and '\\subsection' in line:
                line = line.replace('\\subsection', '\\section', 1)
                promoted = True
                break
        result.append(line)
    return '\n'.join(result)


def fix_unnumbered_sections(text):
    r"""Make ВСТУП, ВИСНОВКИ, СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ unnumbered (\section*)."""
    unnumbered = [
        'ВСТУП',
        'ВИСНОВКИ',
        'СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ',
    ]
    for title in unnumbered:
        text = re.sub(
            rf'\\section\{{({title})\}}',
            rf'\\section*{{\1}}',
            text,
        )
    return text


def remove_toc_section(text):
    """Remove spurious 'Table of Contents' section at the top."""
    text = re.sub(
        r'\\section\{Table of Contents\}\\label\{[^}]*\}\s*\n?',
        '',
        text,
    )
    return text


def extract_labels_from_sources(src_dir):
    r"""Extract \label{} and their surrounding context from original .tex files."""
    labels = {}  # label_name -> caption or nearby text for matching
    tex_files = glob.glob(os.path.join(src_dir, '**', '*.tex'), recursive=True)

    for fpath in tex_files:
        with open(fpath, 'r', encoding='utf-8') as f:
            content = f.read()

        # Find \caption{...}\label{...} pairs
        for m in re.finditer(
            r'\\caption\{([^}]+)\}\s*\\label\{([^}]+)\}', content
        ):
            caption_text = m.group(1).strip()
            label_name = m.group(2).strip()
            labels[label_name] = caption_text

        # Find \label{} near \begin{equation}
        for m in re.finditer(
            r'\\begin\{equation\}\s*\\label\{([^}]+)\}', content
        ):
            labels[m.group(1)] = '__equation__'

    return labels


def restore_labels(text, labels):
    r"""Re-insert \label{} after \caption{} lines by matching caption text."""
    for label_name, caption_text in labels.items():
        if caption_text == '__equation__':
            continue
        # Escape special regex chars in caption
        escaped = re.escape(caption_text)
        # Find caption without a label and add one
        pattern = rf'(\\caption\{{{escaped}\}})'
        replacement = rf'\1\\label{{{label_name}}}'
        text = re.sub(pattern, replacement, text, count=1)
    return text


def md5_file(path):
    """Compute MD5 hash of a file."""
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def build_image_mapping(src_dir, media_dir):
    """Build a mapping from rIdXX paths to original filenames by comparing checksums."""
    if not media_dir or not os.path.isdir(media_dir):
        return {}

    # Hash all original images in src_dir
    original_by_hash = {}
    for ext in ('*.png', '*.jpg', '*.jpeg', '*.gif', '*.svg'):
        for fpath in glob.glob(os.path.join(src_dir, ext)):
            h = md5_file(fpath)
            original_by_hash[h] = os.path.basename(fpath)

    # Hash all round-tripped images and build mapping
    mapping = {}
    for fpath in glob.glob(os.path.join(media_dir, '**', '*'), recursive=True):
        if not os.path.isfile(fpath):
            continue
        h = md5_file(fpath)
        if h in original_by_hash:
            mapping[fpath] = original_by_hash[h]

    return mapping


def fix_image_paths(text, mapping, output_dir):
    """Replace rIdXX paths with original filenames and copy images to output dir."""
    images_dir = os.path.join(output_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)

    for rId_path, original_name in mapping.items():
        # Copy original image to output images/ dir
        dest = os.path.join(images_dir, original_name)
        if not os.path.exists(dest):
            shutil.copy2(rId_path, dest)

        # Replace all occurrences of the rId path in the text
        # The path in the .tex file may be relative, find the actual reference
        text = text.replace(rId_path, f'images/{original_name}')

        # Also try with just the relative portion that pandoc might use
        # e.g. build/latex/media/media/rId9.png
        basename = os.path.basename(rId_path)
        parent = os.path.basename(os.path.dirname(rId_path))
        grandparent = os.path.basename(os.path.dirname(os.path.dirname(rId_path)))
        for variant in [
            f'{grandparent}/{parent}/{basename}',
            f'{parent}/{basename}',
            basename,
        ]:
            if variant in text:
                text = text.replace(variant, f'images/{original_name}')

    return text


def add_preamble(text):
    """Add a basic preamble if the document has none."""
    if '\\documentclass' not in text:
        preamble = r"""\documentclass[14pt,a4paper]{extarticle}

\usepackage{fontspec}
\setmainfont{Times New Roman}

\usepackage[ukrainian]{babel}
\usepackage{geometry}
\usepackage{graphicx}
\usepackage{indentfirst}
\usepackage{fancyhdr}
\usepackage{amsmath}
\usepackage{tocloft}
\usepackage{xcolor}
\usepackage{titlesec}
\usepackage{hyperref}

\titleformat{\section}
  {\raggedright\normalfont\Large\bfseries}
  {РОЗДІЛ \thesection.}
  {0.5em}
  {\MakeUppercase}

\titleformat{\subsection}
  {\normalfont\large\bfseries}
  {\thesubsection}
  {0.5em}
  {}

\renewcommand{\cftsecpresnum}{РОЗДІЛ~}
\renewcommand{\cftsecaftersnum}{. }
\setlength{\cftsecnumwidth}{6em}

\geometry{left=3cm,right=1.5cm,top=2cm,bottom=2cm}
\linespread{1.5}

\begin{document}
"""
        # Find where content starts and prepend
        text = preamble + text + '\n\\end{document}\n'
    return text


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <input.tex> [output.tex] [--src-dir=src]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') else input_file

    src_dir = None
    media_dir = None
    for arg in sys.argv:
        if arg.startswith('--src-dir='):
            src_dir = arg.split('=', 1)[1]
        elif arg.startswith('--media-dir='):
            media_dir = arg.split('=', 1)[1]

    with open(input_file, 'r', encoding='utf-8') as f:
        text = f.read()

    output_dir = os.path.dirname(output_file)

    # Apply fixes in order
    print("  [1/11] Removing spurious ToC section...")
    text = remove_toc_section(text)

    print("  [2/11] Fixing display math: \\[...\\] → equation...")
    text = fix_display_math(text)

    print("  [3/11] Fixing inline math: \\(...\\) → $...$...")
    text = fix_inline_math(text)

    print("  [4/11] Fixing \\hyperref → \\ref...")
    text = fix_hyperref(text)

    print("  [5/11] Stripping baked-in section numbers...")
    text = strip_section_numbers(text)

    print("  [6/11] Fixing section hierarchy...")
    text = fix_section_hierarchy(text)

    print("  [7/11] Fixing en-dashes...")
    text = fix_en_dashes(text)

    print("  [8/11] Fixing list formatting...")
    text = fix_list_formatting(text)

    print("  [9/11] Fixing unnumbered sections (ВСТУП, ВИСНОВКИ, etc.)...")
    text = fix_unnumbered_sections(text)

    if src_dir:
        print("  [10/11] Restoring \\label{} from original sources...")
        labels = extract_labels_from_sources(src_dir)
        text = restore_labels(text, labels)
        print(f"         Restored {len(labels)} labels")
    else:
        print("  [10/11] Skipping label restoration (no --src-dir)")

    if src_dir and media_dir:
        print("  [11/11] Fixing image paths...")
        mapping = build_image_mapping(src_dir, media_dir)
        text = fix_image_paths(text, mapping, output_dir)
        print(f"         Mapped {len(mapping)} images back to originals")
    else:
        print("  [11/11] Skipping image path fix (no --src-dir or --media-dir)")

    print("  Adding preamble...")
    text = add_preamble(text)

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(text)

    print(f"  Done → {output_file}")


if __name__ == '__main__':
    main()
