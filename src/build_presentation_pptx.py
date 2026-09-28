from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt

from build_presentation_docx_greek_utf8 import BASE_DIR, SCREENSHOTS_DIR


REPORT_DIR = BASE_DIR / "report"
OUTPUT_PATH = REPORT_DIR / "viopliroforiki_parousiasi_ellinika.pptx"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

COLOR_BG = RGBColor(248, 245, 239)
COLOR_PANEL = RGBColor(255, 252, 247)
COLOR_TITLE = RGBColor(108, 62, 42)
COLOR_TEXT = RGBColor(48, 50, 56)
COLOR_MUTED = RGBColor(95, 102, 110)
COLOR_ACCENT = RGBColor(181, 114, 68)
COLOR_SOFT = RGBColor(231, 223, 210)
COLOR_CMD = RGBColor(31, 35, 40)
COLOR_NOTE = RGBColor(241, 232, 214)


def resolve_image(path_like: str) -> Path:
    path = Path(path_like)
    if path.is_absolute():
        return path
    screenshot_candidate = SCREENSHOTS_DIR / path_like
    if screenshot_candidate.exists():
        return screenshot_candidate
    return BASE_DIR / path_like


def set_slide_bg(slide) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG


def add_header_band(slide, title: str, subtitle: str | None = None) -> None:
    band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.85))
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(243, 233, 216)
    band.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.14), Inches(8.8), Inches(0.34))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = "Calibri"
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.48), Inches(0.48), Inches(7.5), Inches(0.22))
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = subtitle
        run.font.name = "Calibri"
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_MUTED


def add_footer(slide, label: str) -> None:
    box = slide.shapes.add_textbox(Inches(11.75), Inches(7.05), Inches(1.0), Inches(0.2))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = label
    run.font.name = "Calibri"
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_MUTED


def add_text_panel(slide, left: float, top: float, width: float, height: float, title: str | None = None) -> object:
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    panel.fill.solid()
    panel.fill.fore_color.rgb = COLOR_PANEL
    panel.line.color.rgb = COLOR_SOFT
    panel.line.width = Pt(1)
    if title:
        title_box = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.08), width - Inches(0.3), Inches(0.22))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title
        run.font.name = "Calibri"
        run.font.size = Pt(12.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_TITLE
    return panel


def add_bullets_box(slide, left: float, top: float, width: float, height: float, lines: list[str], title: str | None = None, font_size: float = 15) -> None:
    add_text_panel(slide, left, top, width, height, title)
    text_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.34), width - Inches(0.35), height - Inches(0.42))
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = 0
        p.bullet = True
        p.space_after = Pt(6)
        run = p.add_run()
        run.text = line
        run.font.name = "Calibri"
        run.font.size = Pt(font_size)
        run.font.color.rgb = COLOR_TEXT


def add_paragraph_box(slide, left: float, top: float, width: float, height: float, paragraphs: list[str], title: str | None = None, font_size: float = 14) -> None:
    add_text_panel(slide, left, top, width, height, title)
    text_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.34), width - Inches(0.35), height - Inches(0.42))
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    first = True
    for line in paragraphs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(7)
        run = p.add_run()
        run.text = line
        run.font.name = "Calibri"
        run.font.size = Pt(font_size)
        run.font.color.rgb = COLOR_TEXT


def add_command_box(slide, left: float, top: float, width: float, commands: list[str], title: str = "Εντολή που τρέχουμε") -> None:
    height = Inches(0.55 + 0.34 * max(1, len(commands)))
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(26, 29, 34)
    panel.line.fill.background()

    title_box = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.08), width - Inches(0.3), Inches(0.18))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(238, 214, 176)

    code_box = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.28), width - Inches(0.3), height - Inches(0.3))
    tf = code_box.text_frame
    tf.word_wrap = True
    first = True
    for cmd in commands:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(4)
        run = p.add_run()
        run.text = cmd
        run.font.name = "Consolas"
        run.font.size = Pt(10.2)
        run.font.color.rgb = RGBColor(245, 245, 245)


def add_note_box(slide, left: float, top: float, width: float, height: float, notes: list[str]) -> None:
    add_text_panel(slide, left, top, width, height, "Τι τονίζω προφορικά")
    box = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.34), width - Inches(0.28), height - Inches(0.42))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for idx, note in enumerate(notes):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.bullet = True
        p.space_after = Pt(5)
        run = p.add_run()
        run.text = note
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_TEXT


def add_image_with_caption(slide, image_path: Path, left: float, top: float, width: float, height: float, caption: str) -> None:
    add_text_panel(slide, left, top, width, height)
    cap = slide.shapes.add_textbox(left + Inches(0.16), top + Inches(0.08), width - Inches(0.3), Inches(0.28))
    tf = cap.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = caption
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.italic = True
    run.font.color.rgb = COLOR_MUTED
    if image_path.exists():
        slide.shapes.add_picture(str(image_path), left + Inches(0.14), top + Inches(0.38), width=width - Inches(0.28), height=height - Inches(0.5))
    else:
        missing = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.8), width - Inches(0.35), Inches(0.3))
        p = missing.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = f"Λείπει εικόνα: {image_path.name}"
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = COLOR_ACCENT


def create_title_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide)

    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.65), Inches(12.0), Inches(5.9))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_PANEL
    shape.line.color.rgb = COLOR_SOFT
    shape.line.width = Pt(1.2)

    title = slide.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(10.8), Inches(1.4))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Βιοπληροφορική - Εργασία 1"
    run.font.name = "Calibri"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

    sub = slide.shapes.add_textbox(Inches(1.25), Inches(2.2), Inches(10.3), Inches(1.3))
    tf = sub.text_frame
    for idx, line in enumerate(
        [
            "Παρουσίαση της μεθοδολογίας, της εκτέλεσης του κώδικα",
            "και των βασικών αποτελεσμάτων της εργασίας.",
        ]
    ):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = line
        run.font.name = "Calibri"
        run.font.size = Pt(18 if idx == 0 else 16)
        run.font.color.rgb = COLOR_TEXT

    bullet_box = slide.shapes.add_textbox(Inches(1.3), Inches(3.5), Inches(10.0), Inches(1.6))
    tf = bullet_box.text_frame
    tf.word_wrap = True
    for idx, text in enumerate(
        [
            "BLASTP και ομόλογες πρωτεΐνες",
            "Multiple sequence alignment",
            "Σύγκριση sequence και δομής spike πρωτεϊνών",
            "Δυναμικός προγραμματισμός και indexing",
        ]
    ):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.bullet = False
        run = p.add_run()
        run.text = f"• {text}"
        run.font.name = "Calibri"
        run.font.size = Pt(15)
        run.font.color.rgb = COLOR_TEXT

    add_footer(slide, "1")


def create_content_slide(prs: Presentation, title: str, subtitle: str, bullets: list[str], notes: list[str], commands: list[str] | None = None, images: list[tuple[str, str]] | None = None, footer: str = "") -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide)
    add_header_band(slide, title, subtitle)

    add_bullets_box(slide, Inches(0.45), Inches(1.1), Inches(5.0), Inches(3.2), bullets, "Κύρια σημεία")

    if commands:
        add_command_box(slide, Inches(0.45), Inches(4.45), Inches(5.0), commands)
        note_top = Inches(5.95)
        note_h = Inches(1.1)
    else:
        note_top = Inches(4.5)
        note_h = Inches(2.55)

    add_note_box(slide, Inches(0.45), note_top, Inches(5.0), note_h, notes)

    images = images or []
    if len(images) == 1:
        img, cap = images[0]
        add_image_with_caption(slide, resolve_image(img), Inches(5.75), Inches(1.1), Inches(7.05), Inches(5.95), cap)
    elif len(images) == 2:
        add_image_with_caption(slide, resolve_image(images[0][0]), Inches(5.75), Inches(1.1), Inches(3.4), Inches(5.95), images[0][1])
        add_image_with_caption(slide, resolve_image(images[1][0]), Inches(9.38), Inches(1.1), Inches(3.4), Inches(5.95), images[1][1])

    add_footer(slide, footer)


def build_presentation() -> Path:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    create_title_slide(prs)

    create_content_slide(
        prs,
        "Εισαγωγή και στόχοι",
        "Τι καλύπτει συνολικά η εργασία",
        [
            "Αναζήτηση ομόλογων πρωτεϊνών με BLASTP",
            "Multiple sequence alignment με MAFFT, MUSCLE και T-Coffee",
            "Pairwise και structural alignment για spike πρωτεΐνες",
            "Global/local alignment με δυναμικό προγραμματισμό",
            "Θεωρητική προσέγγιση suffix tree και indexing",
        ],
        [
            "Η εργασία λειτουργεί σαν αλυσίδα βημάτων όπου κάθε output τροφοδοτεί το επόμενο.",
            "Το ζητούμενο δεν είναι μόνο να τρέξω commands αλλά να ερμηνεύσω βιολογικά τι σημαίνουν τα αποτελέσματα.",
        ],
        footer="2",
    )

    create_content_slide(
        prs,
        "Ερώτημα 1(i)",
        "AlphaFold 3 και ESMFold",
        [
            "Σύντομη θεωρητική σύγκριση δύο σύγχρονων εργαλείων πρόβλεψης δομής.",
            "AlphaFold 3: πιο γενικό, κατάλληλο για ευρύτερα βιομοριακά συστήματα.",
            "ESMFold: πολύ γρήγορο όταν θέλουμε δομή από πρωτεϊνική αλληλουχία.",
            "Συμπέρασμα: η επιλογή εργαλείου εξαρτάται από το βιολογικό πρόβλημα.",
        ],
        [
            "Στη βιοπληροφορική δεν υπάρχει πάντα ένα μοναδικά καλύτερο εργαλείο.",
            "Η πρόβλεψη δομής είναι κρίσιμη όταν δεν υπάρχουν πειραματικά δεδομένα.",
        ],
        footer="3",
    )

    create_content_slide(
        prs,
        "Ερώτημα 1(ii)",
        "BLASTP για hemoglobin beta",
        [
            "Query: ανθρώπινη hemoglobin subunit beta με accession NP_000509.",
            "BLASTP για εντοπισμό σχετικών πρωτεϊνικών αλληλουχιών.",
            "Επιλογή 10 κατάλληλων hits και αποθήκευση σε FASTA.",
            "Το dataset αυτό χρησιμοποιείται ως είσοδος στο επόμενο βήμα του MSA.",
        ],
        [
            "Τα πρώτα hits είναι από συγγενικά πρωτεύοντα, κάτι βιολογικά αναμενόμενο.",
            "Το BLASTP εδώ λειτουργεί ως φίλτρο για να φτιάξω λογικό εξελικτικό dataset.",
        ],
        commands=[r"python src\q1_blastp_collect.py"],
        images=[
            ("q1 ii results.png", "Εκτέλεση του script για BLASTP της ανθρώπινης hemoglobin beta."),
            ("q1 ii results top10 file.png", r"Output file: results\question1\hemoglobin_beta_top10.fasta"),
        ],
        footer="4",
    )

    create_content_slide(
        prs,
        "Ερώτημα 1(iii)",
        "Multiple Sequence Alignment με MAFFT, MUSCLE και T-Coffee",
        [
            "Χρήση τριών διαφορετικών MSA εργαλείων πάνω στο ίδιο FASTA dataset.",
            "Και τα τρία jobs ολοκληρώθηκαν επιτυχώς και έδωσαν ουσιαστικά το ίδιο alignment.",
            "Έγινε και επιπλέον MAFFT run με maxiterate=2 για έλεγχο παραμέτρων.",
            "Το εναλλακτικό alignment βγήκε ίδιο, άρα η στοίχιση είναι σταθερή.",
        ],
        [
            "Οι αστερίσκοι στο alignment δείχνουν πλήρως συντηρημένες θέσεις.",
            "Η μη αλλαγή του alignment με διαφορετική παράμετρο ενισχύει την αξιοπιστία του αποτελέσματος.",
        ],
        commands=[
            r"python tools\ebi_clients\mafft.py --email up1093449@ac.upatras.gr --stype protein --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\mafft_test",
            r"python tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
            r"python tools\ebi_clients\tcoffee.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\tcoffee_run",
        ],
        images=[
            ("q1(iii) 1 terminal.png", "Εκτέλεση των jobs για MAFFT, MUSCLE και T-Coffee."),
            ("q1(iii) 2 result file.png", r"Alignment output: results\question1\msa\muscle_run.aln-clustalw.clw"),
        ],
        footer="5",
    )

    create_content_slide(
        prs,
        "Ερώτημα 2",
        "Cytochrome c στα Mammalia",
        [
            "Αφετηρία: UniProt accession P99999 για την ανθρώπινη Cytochrome c.",
            "BLASTP με περιορισμό στα θηλαστικά και υπολογισμός identity percentages.",
            "Τα identities κυμαίνονται περίπου από 98% έως 100%.",
            "Συμπέρασμα: η πρωτεΐνη είναι έντονα συντηρημένη στα Mammalia.",
        ],
        [
            "Η Cytochrome c είναι κλασικό παράδειγμα εξελικτικά συντηρημένης πρωτεΐνης.",
            "Η υψηλή συντήρηση σημαίνει ότι μεγάλες μεταβολές πιθανότατα δεν είναι λειτουργικά ανεκτές.",
        ],
        commands=[r"python src\q2_cytochrome_c.py"],
        images=[
            ("q2 3.png", "Source: UniProt entry P99999 (CYC_HUMAN)"),
            ("q2 1.png", "Εκτέλεση του script για BLASTP στα Mammalia."),
        ],
        footer="6",
    )

    create_content_slide(
        prs,
        "Ερώτημα 2",
        "Cytochrome c - τελικά αποτελέσματα",
        [
            "Το summary συγκεντρώνει τα κορυφαία mammalian hits και τα identity percentages.",
            "Οι πιο κοντινές αλληλουχίες έχουν σχεδόν πλήρη ταύτιση με την ανθρώπινη.",
            "Αυτό συνδέεται με τον βασικό ρόλο της Cytochrome c σε θεμελιώδεις κυτταρικές λειτουργίες.",
        ],
        [
            "Δεν αρκεί να δω απλώς υψηλό identity· πρέπει να το ερμηνεύσω λειτουργικά.",
            "Οι conserved περιοχές υποδηλώνουν κρίσιμα κατάλοιπα για τη λειτουργία της πρωτεΐνης.",
        ],
        images=[
            ("q2 2.png", r"Output file: results\question2\summary.txt"),
        ],
        footer="7",
    )

    create_content_slide(
        prs,
        "Ερώτημα 3",
        "Ταυτοποίηση άγνωστης αλληλουχίας",
        [
            "Η αλληλουχία GIVEQCCTSICSLYQLENYCN αναζητήθηκε με BLASTP.",
            "Η πρωτεΐνη ταυτοποιήθηκε ως Insulin A chain.",
            "Το πρώτο hit είναι από Sus scrofa, αλλά υπάρχουν και ακριβή ανθρώπινα hits.",
            "Άρα η πρωτεΐνη αναγνωρίζεται σωστά, όχι όμως μοναδικά ως ανθρώπινη.",
        ],
        [
            "Το πρώτο BLAST hit δεν σημαίνει πάντα μοναδική ταυτοποίηση οργανισμού.",
            "Το μικρό μήκος της αλληλουχίας εξηγεί γιατί υπάρχουν πολλά 100% matches.",
        ],
        commands=[r"python src\q3_identify_protein.py"],
        images=[
            ("q3 1.png", "Εκτέλεση του script για BLASTP της άγνωστης αλληλουχίας."),
            ("q3 3 summary.png", r"Summary file: results\question3\summary.txt"),
        ],
        footer="8",
    )

    create_content_slide(
        prs,
        "Ερώτημα 3",
        "Ερμηνεία των κορυφαίων BLAST hits",
        [
            "Το πρώτο hit ανήκει σε Sus scrofa με πολύ χαμηλό E-value.",
            "Στα κορυφαία αποτελέσματα εμφανίζονται και πολλαπλά human exact matches.",
            "Η γνωστή λειτουργία της insulin είναι η ρύθμιση της γλυκόζης στο αίμα.",
        ],
        [
            "Η σωστή ερμηνεία είναι συνδυασμός sequence identity, E-value και βιολογικού context.",
            "Αυτό δείχνει ότι δεν κοιτάω μηχανικά μόνο μία γραμμή αποτελέσματος.",
        ],
        images=[
            ("q3 2.png", "Συνέχεια των top BLAST hits στο terminal."),
            ("q3 4 summary .png", r"Summary file: results\question3\summary.txt (συνέχεια)"),
        ],
        footer="9",
    )

    create_content_slide(
        prs,
        "Ερώτημα 4(i)",
        "Pairwise alignment spike πρωτεϊνών με EMBOSS Needle",
        [
            "Ανάκτηση των spike πρωτεϊνών SARS-CoV-2 και SARS-CoV.",
            "Global pairwise alignment με EMBOSS Needle.",
            "Identity 76.4%, similarity 87.0%, gaps 2.0%.",
            "Οι αλληλουχίες είναι συγγενικές αλλά όχι ίδιες.",
        ],
        [
            "Το identity μετρά ακριβή ίδια κατάλοιπα, ενώ το similarity και βιοχημικά παρόμοιες αντικαταστάσεις.",
            "Αυτό βοηθά να εξηγήσω γιατί δύο πρωτεΐνες μπορεί να παραμένουν λειτουργικά συγγενείς παρά τις μεταλλάξεις.",
        ],
        commands=[
            r"python src\q4_spike_alignment.py",
            r"python tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run",
        ],
        images=[
            ("q4 i 1 .png", "Προετοιμασία των αλληλουχιών και εκτέλεση του EMBOSS Needle."),
            ("q4 i txt 2.png", r"Output file: results\question4\needle_run.aln.txt (summary metrics)"),
        ],
        footer="10",
    )

    create_content_slide(
        prs,
        "Ερώτημα 4(i)",
        "Αναλυτικό sequence alignment",
        [
            "Το detailed alignment δείχνει πού ακριβώς υπάρχουν matches, conservative substitutions και gaps.",
            "Μικρές αλληλουχιακές διαφορές μπορεί να έχουν σημαντικό βιολογικό αντίκτυπο στη spike πρωτεΐνη.",
        ],
        [
            "Το screenshot αυτό με βοηθά να δείξω ότι είδα όχι μόνο τα συνοπτικά metrics αλλά και το ίδιο το alignment.",
            "Άρα η ερμηνεία μου δεν βασίζεται μόνο σε ένα ποσοστό, αλλά και στη δομή της στοίχισης.",
        ],
        images=[
            ("q4 i txt 3.png", "Αναλυτικό pairwise alignment από το EMBOSS Needle."),
        ],
        footer="11",
    )

    create_content_slide(
        prs,
        "Ερώτημα 4(ii)",
        "Δομική σύγκριση με SWISS-MODEL και RCSB",
        [
            "Παρουσιάζονται και τα δύο SWISS-MODEL μοντέλα: P0DTC2 και P59594.",
            "Έτσι τεκμηριώνεται πλήρως ότι εξετάστηκαν και οι δύο δομές πριν από το alignment.",
            "Στη συνέχεια έγινε RCSB structure alignment πάνω στα αντίστοιχα μοντέλα.",
        ],
        [
            "Η παρουσία των δύο SWISS-MODEL screenshots καλύπτει καλύτερα την εκφώνηση.",
            "Δεν βασίζομαι μόνο στο τελικό metric αλλά δείχνω και τα ίδια τα δομικά μοντέλα που συγκρίθηκαν.",
        ],
        commands=[r"python src\q4_structure_alignment.py"],
        images=[
            (r"results\question4\swissmodel_p0dtc2.png", "SWISS-MODEL καταχώριση για το SARS-CoV-2 spike (P0DTC2)."),
            (r"results\question4\swissmodel_p59594.png", "SWISS-MODEL καταχώριση για το SARS-CoV spike (P59594)."),
        ],
        footer="12",
    )

    create_content_slide(
        prs,
        "Ερώτημα 4(ii)",
        "RCSB structural alignment metrics",
        [
            "TM-score 0.95 και RMSD 1.8 δείχνουν πολύ κοντινές δομές.",
            "Sequence identity 0.75 με coverage 94% / 97%.",
            "Παρότι η αλληλουχία διαφέρει, η συνολική 3D αναδίπλωση παραμένει εξαιρετικά όμοια.",
        ],
        [
            "Sequence similarity και structural similarity δεν είναι το ίδιο πράγμα.",
            "Το υψηλό TM-score δείχνει διατήρηση της συνολικής δομής παρά τις sequence διαφορές.",
        ],
        images=[
            ("q4 ii 1 terminal.png", "Εκτέλεση του structure alignment script."),
            ("q4 ii 2 png.png", r"Output image: results\question4\rcsb_alignment_summary.png"),
        ],
        footer="13",
    )

    create_content_slide(
        prs,
        "Ερώτημα 5",
        "Global και Local alignment με δυναμικό προγραμματισμό",
        [
            "Needleman-Wunsch για global alignment και Smith-Waterman για local alignment.",
            "Global score: -2 όταν προσπαθούμε να στοιχίσουμε ολόκληρες τις αλληλουχίες.",
            "Local score: 3 και καλύτερο κοινό motif το GTG.",
        ],
        [
            "Η global alignment τιμωρεί όλες τις διαφορές στο πλήρες μήκος.",
            "Η local alignment βρίσκει το καλύτερο κοινό υποτμήμα ακόμα κι όταν οι συνολικές αλληλουχίες διαφέρουν αρκετά.",
        ],
        commands=[r"python src\q5_alignment_dp.py"],
        images=[
            ("q5 1.png", "Εκτέλεση του script για Needleman-Wunsch και Smith-Waterman."),
            ("q5 2 txt global.png", r"Output file: results\question5\global_alignment.txt"),
        ],
        footer="14",
    )

    create_content_slide(
        prs,
        "Ερώτημα 5",
        "Τοπική στοίχιση και κοινό motif",
        [
            "Η local alignment απομονώνει το καλύτερο κοινό τμήμα των δύο αλληλουχιών.",
            "Εδώ το κοινό motif GTG εξηγεί γιατί η local score είναι θετική και μεγαλύτερη από τη global.",
        ],
        [
            "Αυτό το slide με βοηθά να εξηγήσω πρακτικά τη διαφορά global vs local alignment.",
        ],
        images=[
            ("q5 3 txt local.png", r"Output file: results\question5\local_alignment.txt"),
        ],
        footer="15",
    )

    create_content_slide(
        prs,
        "Ερώτημα 6(α)",
        "Θεωρητική λύση με suffix tree",
        [
            "Κατασκευή suffix tree του T$ με preprocessing O(n).",
            "Αποθήκευση της μικρότερης θέσης εμφάνισης minPos σε κάθε κόμβο.",
            "Απάντηση query μήκους m σε χρόνο O(m).",
        ],
        [
            "Η πληροφορία minPos συμπυκνώνει αν υπάρχει εμφάνιση πριν από ένα όριο k χωρίς νέο πλήρες scanning.",
            "Άρα πληρώνω το κόστος μία φορά στην προεπεξεργασία και κερδίζω στα επαναλαμβανόμενα queries.",
        ],
        images=[
            ("q6 1.png", r"Theory file: report\question6_theory.md"),
        ],
        footer="16",
    )

    create_content_slide(
        prs,
        "Ερώτημα 6(β)",
        "Bloom filters, Suffix Arrays και Burrows-Wheeler",
        [
            "Bloom filters: μικρός χώρος και membership queries, αλλά όχι πλήρες exact indexing.",
            "Suffix arrays: ακριβές full-text indexing, αλλά βαρύτερα σε χώρο.",
            "BWT / FM-index: ακριβές indexing με πολύ καλύτερη συμπίεση.",
        ],
        [
            "Οι τεχνικές αυτές εκφράζουν διαφορετικούς συμβιβασμούς ανάμεσα σε χώρο, ταχύτητα και ακρίβεια.",
            "Αυτό είναι κρίσιμο στη βιοπληροφορική επειδή τα datasets είναι τεράστια.",
        ],
        images=[
            ("q6 2.png", r"Theory file: report\question6_theory.md"),
        ],
        footer="17",
    )

    create_content_slide(
        prs,
        "Συνολικά συμπεράσματα",
        "Τι δείχνει συνολικά η εργασία",
        [
            "Συνδέει sequence analysis, structure comparison και indexing.",
            "Συνδυάζει πρακτικά εργαλεία και θεωρητικούς αλγορίθμους.",
            "Τα αποτελέσματα είναι βιολογικά λογικά και υπολογιστικά συνεπή.",
            "Δείχνει ότι η καλή βιοπληροφορική θέλει και σωστή εκτέλεση και σωστή ερμηνεία.",
        ],
        [
            "Μπορώ να κλείσω λέγοντας ότι δεν αντιμετώπισα τα ερωτήματα αποσπασματικά αλλά ως ενιαία ροή ανάλυσης.",
            "Αυτό βοηθά να φανεί ότι έχω καταλάβει όχι μόνο τα εργαλεία, αλλά και τη λογική που τα συνδέει.",
        ],
        footer="18",
    )

    prs.save(OUTPUT_PATH)
    return OUTPUT_PATH


if __name__ == "__main__":
    path = build_presentation()
    print(path)
