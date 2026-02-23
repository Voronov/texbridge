#!/usr/bin/env python3
# Pre-process LaTeX files to make math compatible with pandoc's texmath converter.
#
# Fixes:
#   1. \text{Ukrainian} inside math → \mathrm{Ukrainian}  (texmath handles \mathrm)
#   2. V_{q\_rms} → V_{q,rms}  (escaped underscore in subscripts)
#   3. V_{перенапруга} → V_{\text{перенапруга}} (ensure Cyrillic subscripts wrapped)
#   4. \begin{figure}[H] → \begin{figure}[htbp]  (pandoc doesn't handle [H])
#   5. Combine split .tex files into one for pandoc

import re
import sys
import os


def fix_math_text_commands(content):
    r"""Replace \text{Cyrillic} with pandoc-safe alternatives inside math."""
    # \text{Ukrainian text} → \textrm{Ukrainian text} inside math
    # texmath handles \textrm better than \text for non-ASCII
    def replace_text_in_math(match):
        inner = match.group(1)
        # If it contains Cyrillic, use a space-separated plain text approach
        if re.search(r'[а-яА-ЯіІїЇєЄґҐ]', inner):
            # Replace with just the text outside of math context
            return f'\\textrm{{{inner}}}'
        return match.group(0)

    content = re.sub(r'\\text\{([^}]+)\}', replace_text_in_math, content)
    return content


def fix_escaped_underscores_in_math(content):
    r"""Fix V_{q\_rms} → V_{q\text{\_}rms} patterns in math."""
    # Replace \_ inside subscripts/superscripts with a comma or hyphen
    # This pattern: stuff_{...\_...} → stuff_{...-...}
    def fix_subscript(match):
        return match.group(0).replace('\\_', '{\\text{-}}')

    content = re.sub(r'_\{[^}]*\\_[^}]*\}', fix_subscript, content)
    return content


def fix_figure_placement(content):
    r"""Fix \begin{figure}[H] → \begin{figure}[htbp] and remove stray H."""
    content = re.sub(r'\\begin\{figure\}\{?[Hh]!?\}?', r'\\begin{figure}[htbp]', content)
    content = re.sub(r'\\begin\{figure\}\[H!?\]', r'\\begin{figure}[htbp]', content)
    return content


def fix_cyrillic_subscripts(content):
    """Wrap bare Cyrillic subscripts: V_{перенапруга} → V_{\\text{перенапруга}}"""
    def wrap_cyrillic_sub(match):
        prefix = match.group(1)
        inner = match.group(2)
        # If the inner content is purely Cyrillic (no existing \text wrapper)
        if re.match(r'^[а-яА-ЯіІїЇєЄґҐ_]+$', inner) and '\\text' not in inner:
            return f'{prefix}{{\\textrm{{{inner}}}}}'
        return match.group(0)

    content = re.sub(r'([_^])\{([^}]+)\}', wrap_cyrillic_sub, content)
    return content


def convert_lstlisting_for_pandoc(content, listing_counter):
    r"""Convert \begin{lstlisting}[caption={...}] to captioned \begin{verbatim} for pandoc.

    Args:
        content: LaTeX file content
        listing_counter: dict {chapter_num: count} shared across files
    """
    # Detect chapter number from subsubsection numbering (e.g. \subsubsection*{3.1.1. ...})
    chapter_match = re.search(r'\\subsubsection\*?\{(\d+)\.', content)
    chapter_num = int(chapter_match.group(1)) if chapter_match else 0

    def replace_listing(m):
        caption = m.group(1)
        code = m.group(2)
        # Clean LaTeX escapes for pandoc
        caption = caption.replace(r'\_', '_').replace(r'\#', '#')
        caption = caption.replace(r'\&', '&').replace(r'\%', '%')
        # Increment counter
        if chapter_num not in listing_counter:
            listing_counter[chapter_num] = 0
        listing_counter[chapter_num] += 1
        num = f"{chapter_num}.{listing_counter[chapter_num]}"
        cap_line = f"\\begin{{center}}\n\\textbf{{Лістинг {num} --- {caption}}}\n\\end{{center}}\n"
        return f"{cap_line}\\begin{{verbatim}}{code}\\end{{verbatim}}"

    pattern = re.compile(
        r'\\begin\{lstlisting\}\[caption=\{(.+?)\}\]\s*(.*?)\\end\{lstlisting\}',
        re.DOTALL
    )
    return pattern.sub(replace_listing, content)


def strip_titleformat(content):
    r"""Remove \titleformat commands — they're for xelatex only, pandoc misinterprets them."""
    content = re.sub(
        r'\\titleformat\s*\{[^}]*\}\s*\{[^}]*\}\s*\{[^}]*\}\s*\{[^}]*\}\s*\{[^}]*\}',
        '',
        content
    )
    return content


def strip_titlepage_content(content):
    """Strip titlepage environment and TOC commands — added back by fix_docx.py."""
    # Remove entire \begin{titlepage}...\end{titlepage}
    content = re.sub(
        r'\\begin\{titlepage\}.*?\\end\{titlepage\}',
        '',
        content,
        flags=re.DOTALL
    )
    # Remove \tableofcontents and related tocloft commands
    content = re.sub(r'\\tableofcontents', '', content)
    content = re.sub(r'\\renewcommand\{\\cfttoctitlefont\}.*', '', content)
    content = re.sub(r'\\renewcommand\{\\cftaftertoctitle\}.*', '', content)
    content = re.sub(r'\\renewcommand\{\\contentsname\}.*', '', content)
    content = re.sub(r'\\setcounter\{page\}\{.*?\}', '', content)
    content = re.sub(r'\\thispagestyle\{empty\}', '', content)
    content = re.sub(r'\\newpage', '', content)
    return content


def fix_duplicate_labels(content):
    """Remove duplicate labels, keeping only the first occurrence."""
    seen = set()
    lines = content.split('\n')
    result = []
    for line in lines:
        m = re.search(r'\\label\{([^}]+)\}', line)
        if m:
            label = m.group(1)
            if label in seen:
                # Comment out duplicate label
                line = line.replace(f'\\label{{{label}}}', f'% duplicate: \\label{{{label}}}')
            else:
                seen.add(label)
        result.append(line)
    return '\n'.join(result)


def preprocess_file(input_path, output_path, listing_counter=None):
    """Pre-process a single .tex file."""
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Strip titlepage content (will be added by fix_docx.py post-processor)
    if os.path.basename(input_path) == 'titlepage.tex':
        content = strip_titlepage_content(content)

    content = fix_figure_placement(content)
    content = fix_escaped_underscores_in_math(content)
    content = fix_math_text_commands(content)
    content = fix_cyrillic_subscripts(content)
    content = fix_duplicate_labels(content)
    content = strip_titleformat(content)
    if listing_counter is not None:
        content = convert_lstlisting_for_pandoc(content, listing_counter)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)


def preprocess_project(src_dir, out_dir):
    """Pre-process all .tex files in the project."""
    os.makedirs(out_dir, exist_ok=True)
    sections_out = os.path.join(out_dir, 'sections')
    os.makedirs(sections_out, exist_ok=True)

    listing_counter = {}  # {chapter_num: count}, shared across files
    processed = 0
    for root, dirs, files in os.walk(src_dir):
        for fname in sorted(files):  # sorted for consistent listing numbering
            if not fname.endswith('.tex'):
                continue
            src_path = os.path.join(root, fname)
            # Preserve relative path structure
            rel = os.path.relpath(src_path, src_dir)
            dst_path = os.path.join(out_dir, rel)
            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
            preprocess_file(src_path, dst_path, listing_counter)
            processed += 1

    # Symlink images from src to preprocessed dir
    for f in os.listdir(src_dir):
        ext = os.path.splitext(f)[1].lower()
        if ext in ('.png', '.jpg', '.jpeg', '.gif', '.svg'):
            src_img = os.path.join(src_dir, f)
            dst_img = os.path.join(out_dir, f)
            if not os.path.exists(dst_img):
                os.symlink(os.path.abspath(src_img), dst_img)

    return processed


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f"Usage: {sys.argv[0]} <src_dir> <output_dir>")
        sys.exit(1)

    src_dir = sys.argv[1]
    out_dir = sys.argv[2]

    count = preprocess_project(src_dir, out_dir)
    print(f"  Pre-processed {count} .tex files → {out_dir}/")
