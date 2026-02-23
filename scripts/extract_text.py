#!/usr/bin/env python3
"""Extract plain text content from LaTeX for diffing.
Strips all LaTeX commands, keeps only readable text and math."""

import re
import sys


def extract_text(tex):
    # Remove comments
    tex = re.sub(r'(?<!\\)%.*$', '', tex, flags=re.MULTILINE)

    # Remove preamble (everything before \begin{document})
    m = re.search(r'\\begin\{document\}', tex)
    if m:
        tex = tex[m.end():]
    m = re.search(r'\\end\{document\}', tex)
    if m:
        tex = tex[:m.start()]

    # Normalize equation environments to a marker
    tex = re.sub(r'\\begin\{equation\}', '[EQ]', tex)
    tex = re.sub(r'\\end\{equation\}', '[/EQ]', tex)

    # Keep math content but mark it
    tex = re.sub(r'\$([^$]+)\$', r'$\1$', tex)

    # Remove figure/table environments but keep captions
    tex = re.sub(r'\\begin\{(figure|table|longtable|tabular)\}[^}]*\}?\s*', '', tex)
    tex = re.sub(r'\\end\{(figure|table|longtable|tabular)\}', '', tex)

    # Remove longtable boilerplate
    tex = re.sub(r'\\toprule.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\midrule.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\bottomrule.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\endhead.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\endfirsthead.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\endlastfoot.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\tabularnewline.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\begin\{minipage\}.*$', '', tex, flags=re.MULTILINE)
    tex = re.sub(r'\\end\{minipage\}', '', tex)

    # Extract section titles
    tex = re.sub(r'\\section\*?\{([^}]+)\}', r'\n## \1\n', tex)
    tex = re.sub(r'\\subsection\*?\{([^}]+)\}', r'\n### \1\n', tex)
    tex = re.sub(r'\\subsubsection\*?\{([^}]+)\}', r'\n#### \1\n', tex)

    # Extract caption text
    tex = re.sub(r'\\caption\{([^}]+)\}', r'[Caption: \1]', tex)

    # Extract text from formatting commands
    tex = re.sub(r'\\textbf\{([^}]+)\}', r'\1', tex)
    tex = re.sub(r'\\textit\{([^}]+)\}', r'\1', tex)
    tex = re.sub(r'\\texttt\{([^}]+)\}', r'\1', tex)
    tex = re.sub(r'\\emph\{([^}]+)\}', r'\1', tex)

    # Remove labels, refs (keep ref target)
    tex = re.sub(r'\\label\{[^}]+\}', '', tex)
    tex = re.sub(r'\\ref\{([^}]+)\}', r'[ref:\1]', tex)
    tex = re.sub(r'\\hyperref\[([^\]]+)\]\{[^}]*\}', r'[ref:\1]', tex)

    # Remove includegraphics
    tex = re.sub(r'\\includegraphics\[[^\]]*\]\{[^}]+\}', '[IMAGE]', tex)

    # Remove remaining LaTeX commands but keep their text arguments
    tex = re.sub(r'\\(?:centering|hline|newpage|noindent|clearpage|pagebreak)\b', '', tex)
    tex = re.sub(r'\\(?:begin|end)\{[^}]+\}', '', tex)
    tex = re.sub(r'\\item\b', '- ', tex)
    tex = re.sub(r'\\(?:hfill|vfill|vspace|hspace)\{[^}]*\}', '', tex)
    tex = re.sub(r'\\(?:def|let|renewcommand|newcommand)[^{]*\{[^}]*\}', '', tex)

    # Remove remaining backslash commands without args
    tex = re.sub(r'\\[a-zA-Z]+\b(?!\{)', '', tex)

    # Clean up
    tex = re.sub(r'\{|\}', '', tex)
    tex = re.sub(r'&', ' | ', tex)
    tex = re.sub(r'\\\\', '', tex)
    tex = re.sub(r'~', ' ', tex)
    tex = re.sub(r'\n{3,}', '\n\n', tex)
    tex = re.sub(r'[ \t]+', ' ', tex)

    # Strip each line
    lines = [line.strip() for line in tex.split('\n')]
    # Remove empty lines that are consecutive
    result = []
    prev_empty = False
    for line in lines:
        if not line:
            if not prev_empty:
                result.append('')
            prev_empty = True
        else:
            result.append(line)
            prev_empty = False

    return '\n'.join(result).strip()


if __name__ == '__main__':
    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        tex = f.read()
    print(extract_text(tex))
