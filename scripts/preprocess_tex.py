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


def preprocess_file(input_path, output_path):
    """Pre-process a single .tex file."""
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    content = fix_figure_placement(content)
    content = fix_escaped_underscores_in_math(content)
    content = fix_math_text_commands(content)
    content = fix_cyrillic_subscripts(content)
    content = fix_duplicate_labels(content)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)


def preprocess_project(src_dir, out_dir):
    """Pre-process all .tex files in the project."""
    os.makedirs(out_dir, exist_ok=True)
    sections_out = os.path.join(out_dir, 'sections')
    os.makedirs(sections_out, exist_ok=True)

    processed = 0
    for root, dirs, files in os.walk(src_dir):
        for fname in files:
            if not fname.endswith('.tex'):
                continue
            src_path = os.path.join(root, fname)
            # Preserve relative path structure
            rel = os.path.relpath(src_path, src_dir)
            dst_path = os.path.join(out_dir, rel)
            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
            preprocess_file(src_path, dst_path)
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
