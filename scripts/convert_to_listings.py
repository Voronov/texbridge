#!/usr/bin/env python3
"""Convert \\begin{verbatim} blocks to \\begin{lstlisting}[caption={...}] in LaTeX source files.

Criteria:
- Block must have >= 5 non-empty lines
- Block must not be ASCII art (box-drawing, file trees)
- Caption extracted from preceding \\subsubsection* title
"""

import re
import os
import sys
import glob

MIN_LINES = 5


def is_ascii_art(code):
    """Check if content is ASCII art (diagrams, file trees)."""
    lines = [l for l in code.strip().split('\n') if l.strip()]
    if not lines:
        return False
    art_pattern = re.compile(r'[+|├└┌┐┘┤┬┴┼─│]')
    art_lines = sum(1 for l in lines if art_pattern.search(l))
    return art_lines >= len(lines) * 0.3


def get_caption(text_before):
    """Extract caption from preceding subsubsection or textbf."""
    # Most recent \subsubsection*{X.Y.Z. Title}
    matches = re.findall(r'\\subsubsection\*?\{[\d.]*\s*(.+?)\}', text_before)
    if matches:
        return matches[-1].strip()

    # Most recent \textbf{Title.}
    matches = re.findall(r'\\textbf\{(.+?)\}', text_before[-500:])
    if matches:
        return matches[-1].strip().rstrip('.')

    return 'Фрагмент програмного коду'


def escape_caption(text):
    """Escape LaTeX special characters for use in lstlisting caption."""
    text = text.replace('_', r'\_')
    text = text.replace('#', r'\#')
    text = text.replace('&', r'\&')
    text = text.replace('%', r'\%')
    return text


def process_file(filepath):
    """Convert verbatim blocks to lstlisting in a single .tex file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    pattern = re.compile(r'\\begin\{verbatim\}(.*?)\\end\{verbatim\}', re.DOTALL)

    parts = []
    last_end = 0
    converted = 0

    for m in pattern.finditer(content):
        code = m.group(1)
        lines = [l for l in code.strip().split('\n') if l.strip()]

        parts.append(content[last_end:m.start()])

        if len(lines) < MIN_LINES or is_ascii_art(code):
            parts.append(m.group(0))
        else:
            caption = get_caption(content[:m.start()])
            caption = escape_caption(caption)
            parts.append(f'\\begin{{lstlisting}}[caption={{{caption}}}]')
            parts.append(code.rstrip())
            parts.append('\n\\end{lstlisting}')
            converted += 1

        last_end = m.end()

    parts.append(content[last_end:])

    if converted > 0:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(''.join(parts))

    return converted


def main(src_dir):
    """Process all .tex files in the source directory."""
    total = 0
    for filepath in sorted(glob.glob(os.path.join(src_dir, '**/*.tex'), recursive=True)):
        count = process_file(filepath)
        if count > 0:
            print(f'  {os.path.basename(filepath)}: {count} listings')
            total += count
    print(f'  Total: {total} verbatim blocks converted to lstlisting')


if __name__ == '__main__':
    src_dir = sys.argv[1] if len(sys.argv) > 1 else 'src'
    main(src_dir)
