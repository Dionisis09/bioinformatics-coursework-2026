from __future__ import annotations

from html import escape
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "report"
OUTPUT_HTML = REPORT_DIR / "viopliroforiki_ergasia_clean.html"

REPORT_FILES = [
    REPORT_DIR / "question1_tools.md",
    REPORT_DIR / "question1_blast_alignment.md",
    REPORT_DIR / "question2_cytochrome_c.md",
    REPORT_DIR / "question3_identification.md",
    REPORT_DIR / "question4_spike.md",
    REPORT_DIR / "question5_dp.md",
    REPORT_DIR / "question6_theory.md",
]


CSS = """
:root {
  --bg: #f3ecdf;
  --paper: #fffdf8;
  --ink: #243041;
  --muted: #6b7280;
  --accent: #9a6b43;
  --line: #dccdb8;
  --code: #f4ede2;
}

* { box-sizing: border-box; }

body {
  margin: 0;
  background:
    radial-gradient(circle at top left, #f8f1e3 0%, #f3ecdf 42%, #efe5d5 100%);
  color: var(--ink);
  font-family: "Georgia", "Times New Roman", serif;
}

main {
  max-width: 920px;
  margin: 0 auto;
  padding: 32px 20px 64px;
}

.cover,
section {
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 18px;
  box-shadow: 0 12px 32px rgba(36, 48, 65, 0.08);
}

.cover {
  margin-bottom: 24px;
  padding: 36px 32px;
  text-align: center;
}

.cover h1 {
  margin: 0 0 10px;
  color: var(--accent);
  font-size: 2.4rem;
}

.cover p {
  margin: 0;
  color: var(--muted);
  font-style: italic;
}

section {
  margin-bottom: 24px;
  padding: 28px 32px;
  page-break-inside: avoid;
}

h1, h2, h3 {
  color: var(--accent);
  line-height: 1.15;
}

h1 {
  font-size: 2rem;
  margin: 0 0 18px;
}

h2 {
  font-size: 1.45rem;
  margin: 28px 0 12px;
}

h3 {
  font-size: 1.15rem;
  margin: 22px 0 10px;
}

p, li {
  font-size: 1rem;
  line-height: 1.7;
}

ul, ol {
  margin: 10px 0 16px;
  padding-left: 26px;
}

code {
  background: var(--code);
  border-radius: 6px;
  padding: 2px 6px;
  font-family: "Consolas", "Courier New", monospace;
  font-size: 0.95em;
}

pre {
  background: var(--code);
  border-radius: 12px;
  padding: 16px 18px;
  overflow-x: auto;
  white-space: pre-wrap;
  word-break: break-word;
}

pre code {
  background: transparent;
  padding: 0;
}

@media print {
  body { background: white; }
  main { max-width: none; padding: 0; }
  .cover, section {
    box-shadow: none;
    border-radius: 0;
    border: none;
  }
  .cover { page-break-after: always; }
  section { page-break-after: always; }
  section:last-child { page-break-after: auto; }
}
"""


def render_inline(text: str) -> str:
    parts: list[str] = []
    segments = text.split("`")
    for index, segment in enumerate(segments):
        if index % 2 == 0:
            parts.append(escape(segment))
        else:
            parts.append(f"<code>{escape(segment)}</code>")
    return "".join(parts)


def markdown_to_html(text: str) -> str:
    lines = text.splitlines()
    output: list[str] = []
    in_code = False
    in_ul = False
    in_ol = False
    code_lines: list[str] = []

    def close_lists() -> None:
        nonlocal in_ul, in_ol
        if in_ul:
            output.append("</ul>")
            in_ul = False
        if in_ol:
            output.append("</ol>")
            in_ol = False

    def flush_code() -> None:
        nonlocal in_code, code_lines
        if in_code:
            code_text = "\n".join(code_lines)
            output.append(f"<pre><code>{escape(code_text)}</code></pre>")
            in_code = False
            code_lines = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("```"):
            close_lists()
            if in_code:
                flush_code()
            else:
                in_code = True
                code_lines = []
            continue

        if in_code:
            code_lines.append(line)
            continue

        if not stripped:
            close_lists()
            continue

        if stripped.startswith("# "):
            close_lists()
            output.append(f"<h1>{render_inline(stripped[2:])}</h1>")
            continue

        if stripped.startswith("## "):
            close_lists()
            output.append(f"<h2>{render_inline(stripped[3:])}</h2>")
            continue

        if stripped.startswith("### "):
            close_lists()
            output.append(f"<h3>{render_inline(stripped[4:])}</h3>")
            continue

        if stripped.startswith("- "):
            if in_ol:
                output.append("</ol>")
                in_ol = False
            if not in_ul:
                output.append("<ul>")
                in_ul = True
            output.append(f"<li>{render_inline(stripped[2:])}</li>")
            continue

        if len(stripped) > 2 and stripped[0].isdigit() and ". " in stripped:
            if in_ul:
                output.append("</ul>")
                in_ul = False
            if not in_ol:
                output.append("<ol>")
                in_ol = True
            item_text = stripped.split(". ", 1)[1]
            output.append(f"<li>{render_inline(item_text)}</li>")
            continue

        close_lists()
        output.append(f"<p>{render_inline(stripped)}</p>")

    close_lists()
    flush_code()
    return "\n".join(output)


def build_document() -> str:
    sections: list[str] = []
    for path in REPORT_FILES:
        content = path.read_text(encoding="utf-8")
        sections.append(f"<section>{markdown_to_html(content)}</section>")

    return f"""<!doctype html>
<html lang="el">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Βιοπληροφορική - Εργασία 1</title>
  <style>{CSS}</style>
</head>
<body>
  <main>
    <div class="cover">
      <h1>Βιοπληροφορική - Εργασία 1</h1>
      <p>Καθαρή έκδοση αναφοράς για άνοιγμα σε browser ή εξαγωγή σε PDF</p>
    </div>
    {''.join(sections)}
  </main>
</body>
</html>
"""


def main() -> None:
    html = build_document()
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(OUTPUT_HTML)


if __name__ == "__main__":
    main()
