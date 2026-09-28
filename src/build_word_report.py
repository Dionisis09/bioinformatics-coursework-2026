from __future__ import annotations

import csv
import textwrap
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "report"
RESULTS_DIR = ROOT / "results"
ASSETS_DIR = RESULTS_DIR / "presentation_assets"
OUTPUT_DOCX = REPORT_DIR / "viopliroforiki_ergasia_parousiasi.docx"

EMAIL = "up1093449@ac.upatras.gr"


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def read_lines(path: Path, max_lines: int | None = None) -> list[str]:
    lines = read_text(path).splitlines()
    if max_lines is not None:
        return lines[:max_lines]
    return lines


def make_dirs() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)


def load_font(size: int, mono: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if mono:
        candidates = [
            "C:/Windows/Fonts/consola.ttf",
            "C:/Windows/Fonts/lucon.ttf",
            "DejaVuSansMono.ttf",
        ]
    else:
        candidates = [
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "DejaVuSans.ttf",
        ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def wrap_terminal_text(text: str, width: int = 118) -> list[str]:
    wrapped: list[str] = []
    for raw_line in text.splitlines():
        if not raw_line:
            wrapped.append("")
            continue
        expanded = raw_line.expandtabs(4)
        chunks = textwrap.wrap(
            expanded,
            width=width,
            replace_whitespace=False,
            drop_whitespace=False,
            break_long_words=True,
            break_on_hyphens=False,
        )
        wrapped.extend(chunks or [""])
    return wrapped


def create_terminal_image(path: Path, title: str, command: str, body: str) -> Path:
    mono = load_font(26, mono=True)
    title_font = load_font(24, mono=False)
    body_lines = wrap_terminal_text(body, width=110)
    content_lines = [f"PS {ROOT}> {command}", ""] + body_lines

    left_pad = 34
    top_pad = 30
    line_height = 34
    header_height = 52
    max_chars = max((len(line) for line in content_lines), default=40)
    width = max(1200, min(1850, 100 + max_chars * 14))
    height = header_height + top_pad * 2 + max(8, len(content_lines)) * line_height

    image = Image.new("RGB", (width, height), "#0f172a")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((0, 0, width - 1, height - 1), radius=18, fill="#0f172a", outline="#1e293b", width=2)
    draw.rounded_rectangle((0, 0, width - 1, header_height), radius=18, fill="#111827", outline="#111827")
    draw.rectangle((0, header_height - 18, width - 1, header_height), fill="#111827")

    for idx, color in enumerate(["#ef4444", "#f59e0b", "#10b981"]):
        x = 28 + idx * 24
        draw.ellipse((x, 16, x + 14, 30), fill=color)

    draw.text((96, 12), title, font=title_font, fill="#e5e7eb")

    y = header_height + top_pad
    for line in content_lines:
        fill = "#93c5fd" if line.startswith("PS ") else "#e2e8f0"
        draw.text((left_pad, y), line, font=mono, fill=fill)
        y += line_height

    image.save(path)
    return path


def create_note_image(path: Path, title: str, body: str) -> Path:
    font = load_font(28, mono=False)
    mono = load_font(24, mono=True)
    lines = wrap_terminal_text(body, width=95)
    line_height = 34
    width = 1500
    height = 110 + len(lines) * line_height

    image = Image.new("RGB", (width, height), "#fffdf6")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((0, 0, width - 1, height - 1), radius=24, fill="#fffdf6", outline="#d6d3d1", width=2)
    draw.text((38, 26), title, font=font, fill="#9a3412")
    y = 76
    for line in lines:
        draw.text((40, y), line, font=mono, fill="#1f2937")
        y += line_height

    image.save(path)
    return path


def style_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.font.size = Pt(11)

    if "CodeBlock" not in doc.styles:
        style = doc.styles.add_style("CodeBlock", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Consolas"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
        style.font.size = Pt(9.5)

    for style_name, size in [("Title", 24), ("Heading 1", 17), ("Heading 2", 14), ("Heading 3", 12)]:
        style = doc.styles[style_name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        style.font.size = Pt(size)
        style.font.bold = True


def shade_paragraph(paragraph, fill: str = "F3F4F6") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    p_pr.append(shd)


def add_code_block(doc: Document, code: str) -> None:
    paragraph = doc.add_paragraph(style="CodeBlock")
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.left_indent = Cm(0.3)
    paragraph.paragraph_format.right_indent = Cm(0.3)
    shade_paragraph(paragraph)
    run = paragraph.add_run(code)
    run.font.name = "Consolas"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")


def add_bullets(doc: Document, items: list[str]) -> None:
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def add_image(doc: Document, path: Path, width_inches: float) -> None:
    if not path.exists():
        return
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(str(path), width=Inches(width_inches))


def add_caption(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(90, 90, 90)


def add_metric_table(doc: Document, headers: list[str], rows: list[list[str]], widths: list[float] | None = None) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True

    hdr_cells = table.rows[0].cells
    for idx, header in enumerate(headers):
        hdr_cells[idx].text = header
        hdr_cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for paragraph in hdr_cells[idx].paragraphs:
            for run in paragraph.runs:
                run.bold = True

    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cells[idx].text = value
            cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if widths:
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = Inches(width)


def add_section_title(doc: Document, title: str, subtitle: str | None = None) -> None:
    doc.add_heading(title, level=1)
    if subtitle:
        p = doc.add_paragraph()
        run = p.add_run(subtitle)
        run.italic = True
        run.font.color.rgb = RGBColor(90, 90, 90)


def parse_tsv_rows(path: Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        rows = []
        for row in reader:
            rows.append(
                [
                    row["record_id"],
                    row["identity_percent"],
                    row["aligned_positions"],
                    row["description"],
                ]
            )
        return rows


def q1_hit_rows(summary_text: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in summary_text.splitlines():
        stripped = line.strip()
        if stripped[:2].isdigit() and ". accession=" in stripped:
            parts = stripped.split("|")
            accession = parts[0].split("accession=")[-1].strip()
            evalue = parts[1].split("e-value=")[-1].strip() if len(parts) > 1 else ""
            desc = parts[2].strip() if len(parts) > 2 else stripped
            rows.append([accession, evalue, desc[:90]])
    return rows[:10]


def extract_alignment_excerpt(path: Path, start_marker: str, length: int = 18) -> str:
    lines = read_lines(path)
    for idx, line in enumerate(lines):
        if start_marker in line:
            return "\n".join(lines[idx : idx + length])
    return "\n".join(lines[:length])


def build_assets() -> dict[str, Path]:
    assets: dict[str, Path] = {}

    q1_summary = read_text(RESULTS_DIR / "question1" / "hemoglobin_beta_summary.txt")
    q1_msa_excerpt = extract_alignment_excerpt(
        RESULTS_DIR / "question1" / "msa" / "muscle_run.aln-clustalw.clw",
        "XP_018891709",
        length=18,
    )
    q2_summary = read_text(RESULTS_DIR / "question2" / "summary.txt")
    q3_summary = read_text(RESULTS_DIR / "question3" / "summary.txt")
    q4_seq_summary = read_text(RESULTS_DIR / "question4" / "sequence_summary.txt")
    q4_needle_excerpt = extract_alignment_excerpt(
        RESULTS_DIR / "question4" / "needle_run.aln.txt",
        "Length:",
        length=28,
    )
    q4_structure_summary = read_text(RESULTS_DIR / "question4" / "rcsb_structure_summary.txt")
    q5_global = read_text(RESULTS_DIR / "question5" / "global_alignment.txt")
    q5_local = read_text(RESULTS_DIR / "question5" / "local_alignment.txt")
    q6_algo = """Preprocess(T):
    build suffix tree ST of T$
    DFS(ST.root)

DFS(v):
    if v is leaf:
        v.minPos = suffix_start(v)
    else:
        v.minPos = +inf
        for each child u of v:
            DFS(u)
            v.minPos = min(v.minPos, u.minPos)

OccursBeforeK(P, k):
    locus = traverse_suffix_tree(ST, P)
    if locus does not exist:
        return False
    return locus.minPos < k
"""

    assets["q1_blast"] = create_terminal_image(
        ASSETS_DIR / "q1_blast_terminal.png",
        "Question 1(ii) - BLASTP summary",
        r".\.venv\Scripts\python.exe src\q1_blastp_collect.py",
        q1_summary,
    )
    assets["q1_msa"] = create_terminal_image(
        ASSETS_DIR / "q1_msa_terminal.png",
        "Question 1(iii) - MUSCLE alignment excerpt",
        r".\.venv\Scripts\python.exe tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
        q1_msa_excerpt,
    )
    assets["q2_terminal"] = create_terminal_image(
        ASSETS_DIR / "q2_terminal.png",
        "Question 2 - Cytochrome c analysis",
        r".\.venv\Scripts\python.exe src\q2_cytochrome_c.py",
        q2_summary,
    )
    assets["q3_terminal"] = create_terminal_image(
        ASSETS_DIR / "q3_terminal.png",
        "Question 3 - Short sequence identification",
        r".\.venv\Scripts\python.exe src\q3_identify_protein.py",
        q3_summary,
    )
    assets["q4_sequence"] = create_terminal_image(
        ASSETS_DIR / "q4_sequence_terminal.png",
        "Question 4(i) - Spike sequence prep",
        r".\.venv\Scripts\python.exe src\q4_spike_alignment.py",
        q4_seq_summary,
    )
    assets["q4_needle"] = create_terminal_image(
        ASSETS_DIR / "q4_needle_terminal.png",
        "Question 4(i) - EMBOSS Needle excerpt",
        r".\.venv\Scripts\python.exe tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run",
        q4_needle_excerpt,
    )
    assets["q4_structure"] = create_terminal_image(
        ASSETS_DIR / "q4_structure_terminal.png",
        "Question 4(ii) - RCSB structure alignment",
        r".\.venv\Scripts\python.exe src\q4_structure_alignment.py",
        q4_structure_summary,
    )
    assets["q5_global"] = create_terminal_image(
        ASSETS_DIR / "q5_global_terminal.png",
        "Question 5 - Global alignment",
        r".\.venv\Scripts\python.exe src\q5_alignment_dp.py",
        q5_global,
    )
    assets["q5_local"] = create_terminal_image(
        ASSETS_DIR / "q5_local_terminal.png",
        "Question 5 - Local alignment",
        r"type results\question5\local_alignment.txt",
        q5_local,
    )
    assets["q6_algo"] = create_note_image(
        ASSETS_DIR / "q6_algorithm_note.png",
        "Question 6(a) - Core algorithm idea",
        q6_algo,
    )
    assets["q6_theory"] = create_note_image(
        ASSETS_DIR / "q6_theory_note.png",
        "Question 6(b) - Main indexing ideas",
        "Bloom filters: very small memory, fast membership, allow false positives.\n"
        "Suffix arrays: exact text indexing and efficient substring search.\n"
        "Burrows-Wheeler transform: strong compression and FM-index style search.\n"
        "MBWT combines compact indexing ideas for large sequence collections.",
    )
    return assets


def build_cover(doc: Document) -> None:
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Ξ’ΞΉΞΏΟ€Ξ»Ξ·ΟΞΏΟ†ΞΏΟΞΉΞΊΞ® - Ξ•ΟΞ³Ξ±ΟƒΞ―Ξ± 1")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(31, 41, 55)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Ξ‘Ξ½Ξ±Ξ»Ο…Ο„ΞΉΞΊΞ® Ο€Ξ±ΟΞΏΟ…ΟƒΞ―Ξ±ΟƒΞ· Ξ²Ξ·ΞΌΞ¬Ο„Ο‰Ξ½, ΞµΞ½Ο„ΞΏΞ»ΟΞ½ ΞΊΞ±ΞΉ Ξ±Ο€ΞΏΟ„ΞµΞ»ΞµΟƒΞΌΞ¬Ο„Ο‰Ξ½")
    run.italic = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(107, 114, 128)

    doc.add_paragraph("")
    for line in [
        "ΞΞ¬ΞΈΞ·ΞΌΞ±: Ξ’ΞΉΞΏΟ€Ξ»Ξ·ΟΞΏΟ†ΞΏΟΞΉΞΊΞ®",
        "ΞΞ³Ξ³ΟΞ±Ο†ΞΏ: ΞΏΟΞ³Ξ±Ξ½Ο‰ΞΌΞ­Ξ½Ξ· Ο€Ξ±ΟΞΏΟ…ΟƒΞ―Ξ±ΟƒΞ· Ξ±Ξ½Ξ¬ ΞµΟΟΟ„Ξ·ΞΌΞ±",
        f"Email Ο€ΞΏΟ… Ο‡ΟΞ·ΟƒΞΉΞΌΞΏΟ€ΞΏΞΉΞ®ΞΈΞ·ΞΊΞµ ΟƒΟ„Ξ± web services: {EMAIL}",
        f"Ξ—ΞΌΞµΟΞΏΞΌΞ·Ξ½Ξ―Ξ± ΟƒΟΞ½ΞΈΞµΟƒΞ·Ο‚: {date.today().strftime('%d/%m/%Y')}",
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(line)

    doc.add_paragraph("")
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.add_run(
        "Ξ£Ο„ΞΏ Ο€Ξ±ΟΟΞ½ Word ΞΊΞ¬ΞΈΞµ ΞµΟΟΟ„Ξ·ΞΌΞ± Ο€Ξ±ΟΞΏΟ…ΟƒΞΉΞ¬Ξ¶ΞµΟ„Ξ±ΞΉ ΞΎΞµΟ‡Ο‰ΟΞΉΟƒΟ„Ξ¬. "
        "Ξ“ΞΉΞ± ΞΊΞ¬ΞΈΞµ ΞΌΞ­ΟΞΏΟ‚ Ο†Ξ±Ξ―Ξ½ΞµΟ„Ξ±ΞΉ ΟƒΟΞ½Ο„ΞΏΞΌΞ± Ο„ΞΉ Ξ¶Ξ·Ο„Ξ¬ Ξ· ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ·, Ο€ΞΏΞΉΞ± ΞµΞ½Ο„ΞΏΞ»Ξ® Ξ® Ο€ΞΏΞΉΞΏ workflow "
        "Ο‡ΟΞ·ΟƒΞΉΞΌΞΏΟ€ΞΏΞΉΞ®ΞΈΞ·ΞΊΞµ, ΞΊΞ±ΞΉ Ο„ΞΉ ΞµΞΌΟ†Ξ±Ξ½Ξ―ΟƒΟ„Ξ·ΞΊΞµ Ο‰Ο‚ Ξ±Ο€ΞΏΟ„Ξ­Ξ»ΞµΟƒΞΌΞ± ΞΌΞµ Ο€Ξ―Ξ½Ξ±ΞΊΞµΟ‚ Ξ® screenshots."
    )
    doc.add_page_break()


def build_question_1(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ξ•ΟΟΟ„Ξ·ΞΌΞ± 1", "AlphaFold 3, ESMFold, BLASTP ΞΊΞ±ΞΉ multiple sequence alignment")

    doc.add_heading("1(i) Ξ£ΟΞ½Ο„ΞΏΞΌΞ· Ξ±Ξ½Ξ±Ο†ΞΏΟΞ¬ Ξ³ΞΉΞ± AlphaFold 3 ΞΊΞ±ΞΉ ESMFold", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ ΟƒΟΞ½Ο„ΞΏΞΌΞ· Ξ²ΞΉΞ²Ξ»ΞΉΞΏΞ³ΟΞ±Ο†ΞΉΞΊΞ® Ξ±Ξ½Ξ±Ο†ΞΏΟΞ¬ 1-3 ΟƒΞµΞ»Ξ―Ξ΄Ο‰Ξ½ Ξ³ΞΉΞ± Ο„Ξ· Ξ²Ξ±ΟƒΞΉΞΊΞ® Ξ»ΞµΞΉΟ„ΞΏΟ…ΟΞ³ΞΉΞΊΟΟ„Ξ·Ο„Ξ± "
        "Ο„Ο‰Ξ½ ΞµΟΞ³Ξ±Ξ»ΞµΞ―Ο‰Ξ½ AlphaFold 3 ΞΊΞ±ΞΉ ESMFold."
    )
    add_bullets(
        doc,
        [
            "Ξ¤ΞΏ AlphaFold 3 ΞµΟ€ΞµΞΊΟ„ΞµΞ―Ξ½ΞµΞΉ Ο„Ξ·Ξ½ Ο€ΟΟΞ²Ξ»ΞµΟΞ· Ξ΄ΞΏΞΌΞ®Ο‚ ΟƒΞµ Ο€ΟΟ‰Ο„ΞµΞΞ½ΞµΟ‚, DNA, RNA, ligands ΞΊΞ±ΞΉ ΟƒΟΞΌΟ€Ξ»ΞΏΞΊΞ± Ξ±Ξ»Ξ»Ξ·Ξ»ΞµΟ€ΞΉΞ΄ΟΞ¬ΟƒΞµΟ‰Ξ½.",
            "Ξ¤ΞΏ ESMFold Ο€ΟΞΏΞ²Ξ»Ξ­Ο€ΞµΞΉ Ξ³ΟΞ®Ξ³ΞΏΟΞ± Ο€ΟΟ‰Ο„ΞµΟΞ½ΞΉΞΊΞ® Ξ΄ΞΏΞΌΞ® Ξ±Ο€Ο Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―Ξ±, Ξ±ΞΎΞΉΞΏΟ€ΞΏΞΉΟΞ½Ο„Ξ±Ο‚ protein language models.",
            "Ξ£Ο…ΞΌΟ€Ξ­ΟΞ±ΟƒΞΌΞ±: Ο„ΞΏ AlphaFold 3 ΞµΞ―Ξ½Ξ±ΞΉ Ο€ΞΉΞΏ Ξ³ΞµΞ½ΞΉΞΊΟ Ο‰Ο‚ Ο€ΟΞΏΟ‚ Ο„Ξ± Ξ²ΞΉΞΏΞΌΟΟΞΉΞ±, ΞµΞ½Ο Ο„ΞΏ ESMFold ΞΎΞµΟ‡Ο‰ΟΞ―Ξ¶ΞµΞΉ Ξ³ΞΉΞ± Ο„Ξ±Ο‡ΟΟ„Ξ·Ο„Ξ± ΞΊΞ±ΞΉ Ξ±Ο€Ξ»Ο pipeline.",
        ],
    )
    doc.add_paragraph(
        "Ξ¤ΞΏ Ξ±Ξ½Ξ±Ξ»Ο…Ο„ΞΉΞΊΟ ΞΊΞµΞ―ΞΌΞµΞ½ΞΏ Ο„Ξ·Ο‚ ΞΈΞµΟ‰ΟΞ·Ο„ΞΉΞΊΞ®Ο‚ Ξ±Ξ½Ξ±Ο†ΞΏΟΞ¬Ο‚ Ξ²ΟΞ―ΟƒΞΊΞµΟ„Ξ±ΞΉ ΟƒΟ„ΞΏ Ξ±ΟΟ‡ΞµΞ―ΞΏ "
        "`report/question1_tools.md` ΞΊΞ±ΞΉ ΞµΞ½ΟƒΟ‰ΞΌΞ±Ο„ΟΞ½ΞµΟ„Ξ±ΞΉ ΟƒΟ…Ξ½ΞΏΟ€Ο„ΞΉΞΊΞ¬ ΞµΞ΄Ο."
    )

    doc.add_heading("1(ii) BLASTP Ξ³ΞΉΞ± Hemoglobin subunit beta", level=2)
    doc.add_paragraph("Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ½Ξ± ΞµΞΊΟ„ΞµΞ»ΞµΟƒΟ„ΞµΞ― BLASTP Ξ³ΞΉΞ± Ο„Ξ·Ξ½ Ξ±Ξ½ΞΈΟΟΟ€ΞΉΞ½Ξ· hemoglobin subunit beta (NP_000509), Ξ½Ξ± ΞµΟ€ΞΉΞ»ΞµΞ³ΞΏΟΞ½ 10 Ο€ΟΟ‰Ο„ΞµΞΞ½ΞµΟ‚ Ξ±Ο€Ο Ο„Ξ± Ξ±Ο€ΞΏΟ„ΞµΞ»Ξ­ΟƒΞΌΞ±Ο„Ξ± ΞΊΞ±ΞΉ Ξ½Ξ± Ξ±Ο€ΞΏΞΈΞ·ΞΊΞµΟ…Ο„ΞΏΟΞ½ ΟƒΞµ FASTA.")
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q1_blastp_collect.py")
    add_image(doc, assets["q1_blast"], 6.5)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 1. Ξ•ΞΊΟ„Ξ­Ξ»ΞµΟƒΞ· Ο„ΞΏΟ… script Ξ³ΞΉΞ± BLASTP ΞΊΞ±ΞΉ ΞµΟ€ΞΉΞ»ΞΏΞ³Ξ® Ο„Ο‰Ξ½ 10 Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡ΞΉΟΞ½.")

    q1_rows = q1_hit_rows(read_text(RESULTS_DIR / "question1" / "hemoglobin_beta_summary.txt"))
    add_metric_table(
        doc,
        ["Accession", "E-value", "Ξ ΞµΟΞΉΞ³ΟΞ±Ο†Ξ® hit"],
        q1_rows,
        widths=[1.2, 1.1, 4.7],
    )

    doc.add_paragraph(
        "Ξ¤Ξ± Ξ±ΟΟ‡ΞµΞ―Ξ± Ο€ΞΏΟ… Ο€ΟΞΏΞ­ΞΊΟ…ΟΞ±Ξ½ ΞµΞ―Ξ½Ξ±ΞΉ Ο„Ξ± `hemoglobin_beta_blast.xml`, "
        "`hemoglobin_beta_top10.fasta` ΞΊΞ±ΞΉ `hemoglobin_beta_summary.txt`."
    )

    doc.add_heading("1(iii) Multiple sequence alignment ΞΌΞµ MAFFT, MUSCLE ΞΊΞ±ΞΉ T-Coffee", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ½Ξ± Ξ³Ξ―Ξ½ΞµΞΉ Ο€ΞΏΞ»Ξ»Ξ±Ο€Ξ»Ξ® ΟƒΟ„ΞΏΞ―Ο‡ΞΉΟƒΞ· Ο„Ο‰Ξ½ 10 FASTA Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡ΞΉΟΞ½ ΟƒΟ„Ξ± ΞµΟΞ³Ξ±Ξ»ΞµΞ―Ξ± MAFFT, MUSCLE ΞΊΞ±ΞΉ T-Coffee ΞΊΞ±ΞΉ Ξ½Ξ± ΟƒΟ…Ξ³ΞΊΟΞΉΞΈΞΏΟΞ½ Ο„Ξ± Ξ±Ο€ΞΏΟ„ΞµΞ»Ξ­ΟƒΞΌΞ±Ο„Ξ¬ Ο„ΞΏΟ…Ο‚."
    )
    add_code_block(
        doc,
        "\n".join(
            [
                r".\.venv\Scripts\python.exe tools\ebi_clients\mafft.py --email up1093449@ac.upatras.gr --stype protein --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\mafft_test",
                r".\.venv\Scripts\python.exe tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
                r".\.venv\Scripts\python.exe tools\ebi_clients\tcoffee.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\tcoffee_run",
            ]
        ),
    )
    add_image(doc, assets["q1_msa"], 6.7)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 2. Ξ‘Ο€ΟΟƒΟ€Ξ±ΟƒΞΌΞ± alignment ΟƒΞµ ΞΌΞΏΟΟ†Ξ® ClustalW Ξ±Ο€Ο Ο„ΞΏ MUSCLE.")
    add_bullets(
        doc,
        [
            "ΞΞ±ΞΉ Ο„Ξ± Ο„ΟΞ―Ξ± ΞµΟΞ³Ξ±Ξ»ΞµΞ―Ξ± Ξ­Ξ΄Ο‰ΟƒΞ±Ξ½ ΞΏΟ…ΟƒΞΉΞ±ΟƒΟ„ΞΉΞΊΞ¬ Ο„ΞΏ Ξ―Ξ΄ΞΉΞΏ Ξ²ΞΉΞΏΞ»ΞΏΞ³ΞΉΞΊΟ ΟƒΟ…ΞΌΟ€Ξ­ΟΞ±ΟƒΞΌΞ±.",
            "Ξ¤ΞΏ alignment ΞµΞ―Ο‡Ξµ ΞΌΞ®ΞΊΞΏΟ‚ 147 ΞΈΞ­ΟƒΞµΞΉΟ‚, ΞΌΞµ 137 Ο€Ξ»Ξ®ΟΟ‰Ο‚ ΟƒΟ…Ξ½Ο„Ξ·ΟΞ·ΞΌΞ­Ξ½ΞµΟ‚ ΟƒΟ„Ξ®Ξ»ΞµΟ‚.",
            "Ξ— Ο€ΞΏΞ»Ο ΞΌΞΉΞΊΟΞ® Ξ±Ο€ΟΞΊΞ»ΞΉΟƒΞ· Ξ΄ΞµΞ―Ο‡Ξ½ΞµΞΉ ΟΟ„ΞΉ ΞΏΞΉ ΞµΟ€ΞΉΞ»ΞµΞ³ΞΌΞ­Ξ½ΞµΟ‚ Ξ²-ΟƒΟ†Ξ±ΞΉΟΞ―Ξ½ΞµΟ‚ ΞµΞ―Ξ½Ξ±ΞΉ ΞµΞΎΞ±ΞΉΟΞµΟ„ΞΉΞΊΞ¬ ΟƒΟ…Ξ³Ξ³ΞµΞ½ΞµΞ―Ο‚.",
        ],
    )
    doc.add_page_break()


def build_question_2(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ξ•ΟΟΟ„Ξ·ΞΌΞ± 2", "Cytochrome c ΟƒΟ„ΞΏΞ½ Ξ¬Ξ½ΞΈΟΟ‰Ο€ΞΏ ΞΊΞ±ΞΉ ΞΏΞΌΟΞ»ΞΏΞ³ΞµΟ‚ Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―ΞµΟ‚ ΟƒΟ„Ξ± Mammalia")

    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ½Ξ± ΞµΞ½Ο„ΞΏΟ€ΞΉΟƒΟ„ΞµΞ― Ξ· Ξ±Ξ½ΞΈΟΟΟ€ΞΉΞ½Ξ· Cytochrome c, Ξ½Ξ± Ξ²ΟΞµΞΈΞµΞ― Ο„ΞΏ UniProt accession, "
        "Ξ½Ξ± ΞµΞΊΟ„ΞµΞ»ΞµΟƒΟ„ΞµΞ― BLASTP Ο€ΞµΟΞΉΞΏΟΞΉΟƒΞΌΞ­Ξ½ΞΏ ΟƒΟ„Ξ± Mammalia ΞΊΞ±ΞΉ Ξ½Ξ± Ο…Ο€ΞΏΞ»ΞΏΞ³ΞΉΟƒΟ„ΞµΞ― Ο„ΞΏ identity percentage "
        "ΞΊΞ¬ΞΈΞµ ΞΏΞΌΟΞ»ΞΏΞ³Ξ·Ο‚ Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―Ξ±Ο‚ Ο‰Ο‚ Ο€ΟΞΏΟ‚ Ο„ΞΏΞ½ Ξ¬Ξ½ΞΈΟΟ‰Ο€ΞΏ."
    )

    doc.add_heading("Ξ’Ξ®ΞΌΞ± 1: UniProt accession ΞΊΞ±ΞΉ screenshot", level=2)
    add_bullets(
        doc,
        [
            "Ξ ΟΟ‰Ο„ΞµΞΞ½Ξ·: Cytochrome c",
            "UniProt accession: P99999",
            "Ξ¤ΞΏ screenshot Ξ±Ο€Ο Ο„ΞΏ UniProt Ξ±Ο€ΞΏΞΈΞ·ΞΊΞµΟΟ„Ξ·ΞΊΞµ ΟƒΟ„ΞΏ `results/question2/uniprot_p99999.png`.",
        ],
    )
    add_image(doc, RESULTS_DIR / "question2" / "uniprot_p99999.png", 6.3)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 3. Ξ— ΞΊΞ±Ο„Ξ±Ο‡ΟΟΞΉΟƒΞ· P99999 ΟƒΟ„ΞΏ UniProt Ξ³ΞΉΞ± Ο„Ξ·Ξ½ Ξ±Ξ½ΞΈΟΟΟ€ΞΉΞ½Ξ· Cytochrome c.")

    doc.add_heading("Ξ’Ξ®ΞΌΞ± 2: BLASTP ΞΊΞ±ΞΉ Ο…Ο€ΞΏΞ»ΞΏΞ³ΞΉΟƒΞΌΟΟ‚ identity", level=2)
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q2_cytochrome_c.py")
    add_image(doc, assets["q2_terminal"], 6.6)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 4. Ξ£ΟΞ½ΞΏΟΞ· Ο„ΞΏΟ… script Ο€ΞΏΟ… Ξ­ΞΊΞ±Ξ½Ξµ BLASTP ΟƒΟ„Ξ± Mammalia ΞΊΞ±ΞΉ Ο…Ο€ΞΏΞ»ΟΞ³ΞΉΟƒΞµ Ο„Ξ± identity percentages.")

    rows = parse_tsv_rows(RESULTS_DIR / "question2" / "identity_table.tsv")
    add_metric_table(
        doc,
        ["Accession", "Identity %", "Aligned positions", "Ξ ΞµΟΞΉΞ³ΟΞ±Ο†Ξ®"],
        rows,
        widths=[1.2, 1.0, 1.0, 4.0],
    )
    add_bullets(
        doc,
        [
            "Ξ— Ο€ΞΉΞΏ ΞΊΞΏΞ½Ο„ΞΉΞ½Ξ® Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―Ξ± ΟƒΟ„ΞΏ retained set Ξ®Ο„Ξ±Ξ½ Ξ· `XP_055149593.1` ΞΌΞµ 100.00% identity.",
            "ΞΞΉ Ο…Ο€ΟΞ»ΞΏΞΉΟ€ΞµΟ‚ ΞΏΞΌΟΞ»ΞΏΞ³ΞµΟ‚ Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―ΞµΟ‚ ΞµΞ―Ο‡Ξ±Ξ½ 98.10%-99.05% identity, Ξ¬ΟΞ± Ξ· Cytochrome c ΞµΞ―Ξ½Ξ±ΞΉ ΞµΞΎΞ±ΞΉΟΞµΟ„ΞΉΞΊΞ¬ ΟƒΟ…Ξ½Ο„Ξ·ΟΞ·ΞΌΞ­Ξ½Ξ· ΟƒΟ„Ξ± ΞΈΞ·Ξ»Ξ±ΟƒΟ„ΞΉΞΊΞ¬.",
            "Ξ£Ο„Ξ·Ξ½ Ξ±Ξ½Ξ±Ο†ΞΏΟΞ¬ ΟƒΟ‡ΞΏΞ»ΞΉΞ¬Ξ¶ΞΏΞ½Ο„Ξ±ΞΉ ΞΊΞ±ΞΉ ΞΏΞΉ Ο€Ξ»Ξ®ΟΟ‰Ο‚ ΟƒΟ…Ξ½Ο„Ξ·ΟΞ·ΞΌΞ­Ξ½ΞµΟ‚ Ο€ΞµΟΞΉΞΏΟ‡Ξ­Ο‚ Ο„Ξ·Ο‚ ΟƒΟ„ΞΏΞ―Ο‡ΞΉΟƒΞ·Ο‚.",
        ],
    )
    doc.add_page_break()


def build_question_3(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ερώτημα 3", "Ταυτοποίηση μικρής πρωτεϊνικής αλληλουχίας με BLASTP")

    doc.add_paragraph(
        "Η εκφώνηση δίνει την αλληλουχία `GIVEQCCTSICSLYQLENYCN` και ζητά να "
        "ταυτοποιηθεί η πρωτεΐνη, να αναφερθεί ο οργανισμός του καλύτερου hit "
        "και να επιβεβαιωθεί σύντομα η λειτουργία της από τη βιβλιογραφία."
    )
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q3_identify_protein.py")
    add_image(doc, assets["q3_terminal"], 6.6)
    add_caption(doc, "Εικόνα 5. Σύνοψη BLASTP για την αλληλουχία του ερωτήματος 3.")
    add_bullets(
        doc,
        [
            "Τα κορυφαία BLAST hits ήταν insulin / insulin A chain, με 100% identity σε μήκος 21 αμινοξέων.",
            "Αν ακολουθήσουμε αυστηρά το πρώτο BLAST hit, ο οργανισμός είναι `Sus scrofa` και όχι `Homo sapiens`.",
            "Υπάρχουν όμως και ακριβή ανθρώπινα hits, άρα η ασφαλέστερη ερμηνεία είναι ότι η αλληλουχία ανήκει σε εξαιρετικά συντηρημένη insulin A chain θηλαστικών.",
            "Η γραπτή αναφορά εξηγεί γιατί το πολύ μικρό μήκος της αλληλουχίας δεν επιτρέπει μοναδική ανθρώπινη ταυτοποίηση, παρότι η αναγνώριση της πρωτεΐνης είναι πολύ ισχυρή.",
        ],
    )
    doc.add_page_break()


def build_question_4(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ξ•ΟΟΟ„Ξ·ΞΌΞ± 4", "Ξ£ΟΞ³ΞΊΟΞΉΟƒΞ· spike Ο€ΟΟ‰Ο„ΞµΟΞ½ΟΞ½ SARS-CoV-2 ΞΊΞ±ΞΉ SARS-CoV")

    doc.add_heading("4(i) Pairwise sequence alignment ΞΌΞµ EMBOSS Needle", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ½Ξ± Ξ±Ξ½Ξ±ΞΊΟ„Ξ·ΞΈΞΏΟΞ½ ΞΏΞΉ spike Ο€ΟΟ‰Ο„ΞµΞΞ½ΞµΟ‚ Ο„Ο‰Ξ½ SARS-CoV-2 ΞΊΞ±ΞΉ SARS-CoV "
        "ΞΊΞ±ΞΉ Ξ½Ξ± ΟƒΟ…Ξ³ΞΊΟΞΉΞΈΞΏΟΞ½ ΞΌΞµ global alignment ΞΌΞ­ΟƒΟ‰ Ο„ΞΏΟ… EMBOSS Needle."
    )
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q4_spike_alignment.py")
    add_image(doc, assets["q4_sequence"], 6.2)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 6. Ξ ΟΞΏΞµΟ„ΞΏΞΉΞΌΞ±ΟƒΞ―Ξ± Ο„Ο‰Ξ½ FASTA Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡ΞΉΟΞ½ Ξ³ΞΉΞ± Ο„ΞΉΟ‚ Ξ΄ΟΞΏ spike Ο€ΟΟ‰Ο„ΞµΞΞ½ΞµΟ‚.")

    add_code_block(
        doc,
        r".\.venv\Scripts\python.exe tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run"
    )
    add_image(doc, assets["q4_needle"], 6.7)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 7. Ξ‘Ο€ΟΟƒΟ€Ξ±ΟƒΞΌΞ± Ο„ΞΏΟ… EMBOSS Needle alignment ΞΊΞ±ΞΉ Ο„Ο‰Ξ½ Ξ²Ξ±ΟƒΞΉΞΊΟΞ½ metrics.")
    add_metric_table(
        doc,
        ["Metric", "Ξ¤ΞΉΞΌΞ®"],
        [
            ["Alignment length", "1277"],
            ["Identity", "975/1277 (76.4%)"],
            ["Similarity", "1111/1277 (87.0%)"],
            ["Gaps", "26/1277 (2.0%)"],
            ["Score", "5230.0"],
        ],
        widths=[2.2, 3.8],
    )

    doc.add_heading("4(ii) Ξ”ΞΏΞΌΞΉΞΊΞ® ΟƒΟΞ³ΞΊΟΞΉΟƒΞ· ΞΌΞµ SWISS-MODEL ΞΊΞ±ΞΉ RCSB", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ΄ΞΏΞΌΞΉΞΊΞ® ΟƒΟΞ³ΞΊΟΞΉΟƒΞ· Ο„Ο‰Ξ½ Ξ΄ΟΞΏ spike Ο€ΟΟ‰Ο„ΞµΟΞ½ΟΞ½. "
        "Ξ§ΟΞ·ΟƒΞΉΞΌΞΏΟ€ΞΏΞΉΞ®ΞΈΞ·ΞΊΞ±Ξ½ Ο„Ξ± ΞΌΞΏΞ½Ο„Ξ­Ξ»Ξ± Ο„ΞΏΟ… SWISS-MODEL Repository ΞΊΞ±ΞΉ ΟƒΟ„Ξ· ΟƒΟ…Ξ½Ξ­Ο‡ΞµΞΉΞ± Ξ­Ξ³ΞΉΞ½Ξµ ΞµΟ…ΞΈΟ…Ξ³ΟΞ¬ΞΌΞΌΞΉΟƒΞ· ΞΌΞµ Ο„ΞΏ RCSB."
    )
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q4_structure_alignment.py")
    add_image(doc, RESULTS_DIR / "question4" / "swissmodel_p0dtc2.png", 5.9)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 8. SWISS-MODEL page Ξ³ΞΉΞ± Ο„ΞΏ SARS-CoV-2 spike (P0DTC2).")
    add_image(doc, RESULTS_DIR / "question4" / "swissmodel_p59594.png", 5.9)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 9. SWISS-MODEL page Ξ³ΞΉΞ± Ο„ΞΏ SARS-CoV spike (P59594).")
    add_image(doc, assets["q4_structure"], 6.2)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 10. Ξ£ΟΞ½ΞΏΟΞ· Ο„ΞΏΟ… RCSB structure alignment.")
    add_image(doc, RESULTS_DIR / "question4" / "rcsb_alignment_summary.png", 5.5)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 11. ΞΟ€Ο„ΞΉΞΊΟ summary Ο„Ο‰Ξ½ Ξ΄ΞΏΞΌΞΉΞΊΟΞ½ metrics Ξ±Ο€Ο Ο„ΞΏ RCSB.")
    add_metric_table(
        doc,
        ["Metric", "Ξ¤ΞΉΞΌΞ®"],
        [
            ["TM-score", "0.95"],
            ["RMSD", "1.8"],
            ["Sequence identity", "0.75"],
            ["Aligned residue pairs", "1078"],
            ["Coverage", "94% / 97%"],
        ],
        widths=[2.3, 3.2],
    )
    add_bullets(
        doc,
        [
            "Ξ Ξ±ΟΟΟ„ΞΉ Ξ· sequence identity Ξ΄ΞµΞ½ ΞµΞ―Ξ½Ξ±ΞΉ Ο€Ξ»Ξ®ΟΞ·Ο‚, Ο„ΞΏ Ο€ΞΏΞ»Ο Ο…ΟΞ·Ξ»Ο TM-score Ξ΄ΞµΞ―Ο‡Ξ½ΞµΞΉ ΞµΞΎΞ±ΞΉΟΞµΟ„ΞΉΞΊΞ¬ Ο€Ξ±ΟΟΞΌΞΏΞΉΞΏ ΟƒΟ…Ξ½ΞΏΞ»ΞΉΞΊΟ fold.",
            "Ξ†ΟΞ± ΞΏΞΉ Ξ΄ΟΞΏ spike Ο€ΟΟ‰Ο„ΞµΞΞ½ΞµΟ‚ ΞµΞ―Ξ½Ξ±ΞΉ ΞµΞΎΞµΞ»ΞΉΞΊΟ„ΞΉΞΊΞ¬ ΟƒΟ…Ξ³Ξ³ΞµΞ½ΞµΞ―Ο‚ ΞΊΞ±ΞΉ Ξ΄ΞΏΞΌΞΉΞΊΞ¬ Ο€ΞΏΞ»Ο ΞΊΞΏΞ½Ο„ΞΉΞ½Ξ­Ο‚.",
        ],
    )
    doc.add_page_break()


def build_question_5(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ξ•ΟΟΟ„Ξ·ΞΌΞ± 5", "Dynamic programming Ξ³ΞΉΞ± global ΞΊΞ±ΞΉ local alignment")

    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ½Ξ± ΟƒΟ…ΞΌΟ€Ξ»Ξ·ΟΟ‰ΞΈΞµΞ― ΞΏ Ο€Ξ―Ξ½Ξ±ΞΊΞ±Ο‚ Ξ΄Ο…Ξ½Ξ±ΞΌΞΉΞΊΞΏΟ Ο€ΟΞΏΞ³ΟΞ±ΞΌΞΌΞ±Ο„ΞΉΟƒΞΌΞΏΟ Ξ³ΞΉΞ± global "
        "ΞΊΞ±ΞΉ local alignment ΟƒΟ„ΞΉΟ‚ Ξ±Ξ»Ξ»Ξ·Ξ»ΞΏΟ…Ο‡Ξ―ΞµΟ‚ `ACTTGGGTG` ΞΊΞ±ΞΉ `GTGTGAATT`, "
        "ΞΌΞµ score `+1` Ξ³ΞΉΞ± match ΞΊΞ±ΞΉ `-1` Ξ³ΞΉΞ± mismatch Ξ® gap."
    )
    add_code_block(doc, r".\.venv\Scripts\python.exe src\q5_alignment_dp.py")

    doc.add_heading("Global alignment", level=2)
    add_image(doc, assets["q5_global"], 6.6)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 12. Ξ Ξ―Ξ½Ξ±ΞΊΞ±Ο‚ Ξ΄Ο…Ξ½Ξ±ΞΌΞΉΞΊΞΏΟ Ο€ΟΞΏΞ³ΟΞ±ΞΌΞΌΞ±Ο„ΞΉΟƒΞΌΞΏΟ ΞΊΞ±ΞΉ traceback Ξ³ΞΉΞ± Needleman-Wunsch.")
    add_bullets(doc, ["Ξ’Ξ­Ξ»Ο„ΞΉΟƒΟ„ΞΏ global score: -2."])

    doc.add_heading("Local alignment", level=2)
    add_image(doc, assets["q5_local"], 6.6)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 13. Ξ Ξ―Ξ½Ξ±ΞΊΞ±Ο‚ Ξ΄Ο…Ξ½Ξ±ΞΌΞΉΞΊΞΏΟ Ο€ΟΞΏΞ³ΟΞ±ΞΌΞΌΞ±Ο„ΞΉΟƒΞΌΞΏΟ ΞΊΞ±ΞΉ traceback Ξ³ΞΉΞ± Smith-Waterman.")
    add_bullets(doc, ["Ξ’Ξ­Ξ»Ο„ΞΉΟƒΟ„ΞΏ local score: 3.", "Ξ¤ΞΏ ΞΊΞ±Ξ»ΟΟ„ΞµΟΞΏ local alignment Ξ±Ξ½Ο„ΞΉΟƒΟ„ΞΏΞΉΟ‡ΞµΞ― ΟƒΟ„ΞΏ motif `GTG`."])
    doc.add_page_break()


def build_question_6(doc: Document, assets: dict[str, Path]) -> None:
    add_section_title(doc, "Ξ•ΟΟΟ„Ξ·ΞΌΞ± 6", "Suffix trees, Bloom filters, suffix arrays ΞΊΞ±ΞΉ Burrows-Wheeler")

    doc.add_heading("6(Ξ±) Ξ‘Ξ»Ξ³ΟΟΞΉΞΈΞΌΞΏΟ‚ ΞΌΞµ suffix tree", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ Ξ±Ξ»Ξ³ΟΟΞΉΞΈΞΌΞΏ Ο€ΞΏΟ…, ΞΌΞµΟ„Ξ¬ Ξ±Ο€Ο `O(n)` Ο€ΟΞΏΞµΟ€ΞµΞΎΞµΟΞ³Ξ±ΟƒΞ―Ξ±, Ξ±Ο€Ξ±Ξ½Ο„Ξ¬ ΟƒΞµ `O(m)` Ξ±Ξ½ "
        "Ξ­Ξ½Ξ± pattern `P` ΞµΞΌΟ†Ξ±Ξ½Ξ―Ξ¶ΞµΟ„Ξ±ΞΉ ΟƒΟ„ΞΏ `T` Ο€ΟΞΉΞ½ Ξ±Ο€Ο ΞΈΞ­ΟƒΞ· `k`."
    )
    add_image(doc, assets["q6_algo"], 6.4)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 14. Ξ Ο€Ο…ΟΞ®Ξ½Ξ±Ο‚ Ο„ΞΏΟ… Ξ±Ξ»Ξ³ΞΏΟΞ―ΞΈΞΌΞΏΟ… ΞΌΞµ `minPos` ΟƒΞµ ΞΊΞ¬ΞΈΞµ ΞΊΟΞΌΞ²ΞΏ Ο„ΞΏΟ… suffix tree.")
    add_bullets(
        doc,
        [
            "Ξ£Ο„ΞΏ suffix tree Ξ±Ο€ΞΏΞΈΞ·ΞΊΞµΟΞµΟ„Ξ±ΞΉ ΟƒΞµ ΞΊΞ¬ΞΈΞµ ΞΊΟΞΌΞ²ΞΏ Ξ· ΞΌΞΉΞΊΟΟΟ„ΞµΟΞ· ΞΈΞ­ΟƒΞ· Ξ­Ξ½Ξ±ΟΞΎΞ·Ο‚ suffix ΟƒΟ„ΞΏ Ο…Ο€ΞΏΞ΄Ξ­Ξ½Ο„ΟΞΏ Ο„ΞΏΟ….",
            "ΞΞµΟ„Ξ¬ Ο„ΞΏ traversal Ο„ΞΏΟ… pattern, Ξ±ΟΞΊΞµΞ― Ξ½Ξ± ΞµΞ»ΞµΞ³Ο‡ΞΈΞµΞ― Ξ±Ξ½ `minPos(locus) < k`.",
            "Ξ— Ο€ΟΞΏΞµΟ€ΞµΞΎΞµΟΞ³Ξ±ΟƒΞ―Ξ± ΞµΞ―Ξ½Ξ±ΞΉ `O(n)` ΞΊΞ±ΞΉ Ο„ΞΏ query `O(m)`.",
        ],
    )

    doc.add_heading("6(Ξ²) Ξ£ΟΞ½Ο„ΞΏΞΌΞ· Ξ΄ΞΉΞµΟΞµΟΞ½Ξ·ΟƒΞ·", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ· Ξ¶Ξ·Ο„Ξ¬ ΞΌΞΉΞΊΟΞ® Ξ΄ΞΉΞµΟΞµΟΞ½Ξ·ΟƒΞ· Ξ³ΞΉΞ± Ο„ΞΏΞ½ ΟΟΞ»ΞΏ Ο„Ο‰Ξ½ Bloom filters, Ο„Ο‰Ξ½ suffix arrays "
        "ΞΊΞ±ΞΉ Ο„ΞΏΟ… Burrows-Wheeler transform ΟƒΟ„Ξ· Ξ²ΞΉΞΏΟ€Ξ»Ξ·ΟΞΏΟ†ΞΏΟΞΉΞΊΞ®."
    )
    add_image(doc, assets["q6_theory"], 6.4)
    add_caption(doc, "Ξ•ΞΉΞΊΟΞ½Ξ± 15. Ξ£Ο…Ξ½ΞΏΟ€Ο„ΞΉΞΊΞ® ΟƒΟΞ³ΞΊΟΞΉΟƒΞ· Ο„Ο‰Ξ½ Ξ²Ξ±ΟƒΞΉΞΊΟΞ½ ΞΉΞ΄ΞµΟΞ½ Ξ΄ΞµΞΉΞΊΟ„ΞΏΞ΄ΟΟ„Ξ·ΟƒΞ·Ο‚.")
    add_bullets(
        doc,
        [
            "Ξ¤Ξ± Bloom filters ΞµΞΎΞΏΞΉΞΊΞΏΞ½ΞΏΞΌΞΏΟΞ½ ΞΌΞ½Ξ®ΞΌΞ· ΞΊΞ±ΞΉ ΞµΞ―Ξ½Ξ±ΞΉ ΞΉΞ΄Ξ±Ξ½ΞΉΞΊΞ¬ Ξ³ΞΉΞ± Ξ³ΟΞ®Ξ³ΞΏΟΞ± membership queries ΞΌΞµ Ο€ΞΉΞΈΞ±Ξ½ΟΟ„Ξ·Ο„Ξ± false positives.",
            "Ξ¤Ξ± suffix arrays Ο…Ο€ΞΏΟƒΟ„Ξ·ΟΞ―Ξ¶ΞΏΟ…Ξ½ Ξ±ΞΊΟΞΉΞ²Ξ® substring search.",
            "Ξ Burrows-Wheeler transform ΞΏΞ΄Ξ·Ξ³ΞµΞ― ΟƒΞµ Ο€ΞΏΞ»Ο ΟƒΟ…ΞΌΟ€Ξ±Ξ³ΞµΞ―Ο‚ Ξ΄ΞΏΞΌΞ­Ο‚ Ξ±Ξ½Ξ±Ξ¶Ξ®Ο„Ξ·ΟƒΞ·Ο‚, ΟΟ€Ο‰Ο‚ Ο„Ξ± FM-indexes.",
        ],
    )

    doc.add_heading("Ξ¤ΞµΞ»ΞΉΞΊΟ ΟƒΟ…ΞΌΟ€Ξ­ΟΞ±ΟƒΞΌΞ±", level=2)
    doc.add_paragraph(
        "Ξ— ΞµΟΞ³Ξ±ΟƒΞ―Ξ± ΞΏΞ»ΞΏΞΊΞ»Ξ·ΟΟΞΈΞ·ΞΊΞµ Ξ±Ξ½Ξ¬ ΞµΟΟΟ„Ξ·ΞΌΞ± Ο‡Ο‰ΟΞ―Ο‚ Ξ½Ξ± Ο€ΟΞΏΟƒΟ„ΞµΞΈΞΏΟΞ½ Ο€ΞµΟΞΉΟ„Ο„Ξ­Ο‚ Ξ±Ξ½Ξ±Ξ»ΟΟƒΞµΞΉΟ‚ ΞµΞΊΟ„ΟΟ‚ ΞµΞΊΟ†ΟΞ½Ξ·ΟƒΞ·Ο‚. "
        "Ξ£Ο„Ξ± Ο€ΟΞ±ΞΊΟ„ΞΉΞΊΞ¬ ΞΌΞ­ΟΞ· Ο†Ξ±Ξ―Ξ½ΞΏΞ½Ο„Ξ±ΞΉ ΞΏΞΉ ΞµΞ½Ο„ΞΏΞ»Ξ­Ο‚, Ο„Ξ± Ο€Ξ±ΟΞ±Ξ³ΟΞΌΞµΞ½Ξ± Ξ±ΟΟ‡ΞµΞ―Ξ± ΞΊΞ±ΞΉ Ο„Ξ± Ξ²Ξ±ΟƒΞΉΞΊΞ¬ Ξ±Ο€ΞΏΟ„ΞµΞ»Ξ­ΟƒΞΌΞ±Ο„Ξ±, "
        "ΞµΞ½Ο ΟƒΟ„Ξ± ΞΈΞµΟ‰ΟΞ·Ο„ΞΉΞΊΞ¬ ΞΌΞ­ΟΞ· Ο€Ξ±ΟΞΏΟ…ΟƒΞΉΞ¬Ξ¶ΞΏΞ½Ο„Ξ±ΞΉ Ο„Ξ± ΞΏΟ…ΟƒΞΉΟΞ΄Ξ· Ξ²Ξ®ΞΌΞ±Ο„Ξ± ΞΊΞ±ΞΉ Ο„ΞΏ ΟƒΟ…ΞΌΟ€Ξ­ΟΞ±ΟƒΞΌΞ±."
    )


def build_document() -> Path:
    make_dirs()
    assets = build_assets()

    doc = Document()
    style_document(doc)
    build_cover(doc)
    build_question_1(doc, assets)
    build_question_2(doc, assets)
    build_question_3(doc, assets)
    build_question_4(doc, assets)
    build_question_5(doc, assets)
    build_question_6(doc, assets)
    doc.save(OUTPUT_DOCX)
    return OUTPUT_DOCX


if __name__ == "__main__":
    output = build_document()
    print(f"Saved report to: {output}")

