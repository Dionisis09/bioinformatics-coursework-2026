from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


BASE_DIR = Path(r"C:\Users\dioni\Desktop\general\viopliroforiki ergasia 1")
REPORT_DIR = BASE_DIR / "report"
SCREENSHOTS_DIR = Path(r"C:\Users\dioni\Pictures\vioplitoforiki screenshots")
OUTPUT_PATH = REPORT_DIR / "viopliroforiki_parousiasi_unicode.docx"


SECTIONS = [
    {
        "title": "Βιοπληροφορική - Εργασία 1",
        "intro": [
            "Παρουσίαση αποτελεσμάτων με βάση την εκτέλεση του κώδικα και τα παραγόμενα outputs.",
            "Σε κάθε ενότητα φαίνεται η εντολή που τρέχουμε, όπου υπάρχει εκτέλεση, και ακολουθούν τα αντίστοιχα screenshots.",
        ],
    },
    {
        "title": "Ερώτημα 1(i): AlphaFold 3 και ESMFold",
        "bullets": [
            "Στο πρώτο υποερώτημα έγινε σύντομη θεωρητική σύγκριση των εργαλείων AlphaFold 3 και ESMFold.",
            "Το AlphaFold 3 είναι πιο γενικό, επειδή υποστηρίζει πρόβλεψη δομής και αλληλεπιδράσεων διαφορετικών βιομορίων.",
            "Το ESMFold δίνει πολύ γρήγορη πρόβλεψη πρωτεϊνικής δομής απευθείας από την αλληλουχία.",
        ],
    },
    {
        "title": "Ερώτημα 1(ii): BLASTP συλλογή 10 hits",
        "commands": [
            r"python src\q1_blastp_collect.py",
        ],
        "images": [
            ("q1 ii results.png", "Εκτέλεση του script για BLASTP της ανθρώπινης hemoglobin beta."),
            ("q1 ii results top10 file.png", r"Output file: results\question1\hemoglobin_beta_top10.fasta"),
        ],
        "bullets": [
            "Χρησιμοποίησα την ανθρώπινη hemoglobin beta με accession NP_000509.",
            "Το script έκανε BLASTP, επέλεξε 10 σχετικά πρωτεϊνικά hits και τα αποθήκευσε σε αρχείο FASTA.",
            "Το αρχείο αυτό χρησιμοποιήθηκε στο επόμενο βήμα για το multiple sequence alignment.",
        ],
    },
    {
        "title": "Ερώτημα 1(iii): Multiple Sequence Alignment",
        "commands": [
            r"python tools\ebi_clients\mafft.py --email up1093449@ac.upatras.gr --stype protein --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\mafft_test",
            r"python tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
            r"python tools\ebi_clients\tcoffee.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\tcoffee_run",
        ],
        "images": [
            ("q1(iii) 1 terminal.png", "Εκτέλεση των jobs για MAFFT, MUSCLE και T-Coffee."),
            ("q1(iii) 2 result file.png", r"Output file: results\question1\msa\muscle_run.aln-clustalw.clw"),
        ],
        "bullets": [
            "Έκανα multiple sequence alignment των 10 αλληλουχιών με τρία διαφορετικά εργαλεία.",
            "Τα αποτελέσματα ήταν πολύ παρόμοια και έδειξαν υψηλή συντήρηση της hemoglobin beta.",
            "Το alignment του MUSCLE χρησιμοποιήθηκε ως αντιπροσωπευτικό παράδειγμα στο output.",
        ],
    },
    {
        "title": "Ερώτημα 2: Cytochrome c στα Mammalia",
        "source": "Source: UniProt entry P99999 (CYC_HUMAN)",
        "commands": [
            r"python src\q2_cytochrome_c.py",
        ],
        "images": [
            ("q2 3.png", "Σελίδα UniProt για την ανθρώπινη Cytochrome c."),
            ("q2 1.png", "Εκτέλεση του script για BLASTP στα Mammalia."),
            ("q2 2.png", r"Output file: results\question2\summary.txt"),
        ],
        "bullets": [
            "Η ανθρώπινη Cytochrome c ανακτήθηκε από το UniProt με accession P99999.",
            "Έγινε BLASTP στα Mammalia και υπολογίστηκαν identity percentages ως προς την ανθρώπινη αλληλουχία.",
            "Οι τιμές ήταν περίπου 98% έως 100%, άρα η πρωτεΐνη είναι εξαιρετικά συντηρημένη στα θηλαστικά.",
        ],
    },
    {
        "title": "Ερώτημα 3: Ταυτοποίηση άγνωστης αλληλουχίας",
        "commands": [
            r"python src\q3_identify_protein.py",
        ],
        "images": [
            ("q3 1.png", "Εκτέλεση του script για BLASTP της άγνωστης αλληλουχίας."),
            ("q3 2.png", "Συνέχεια των top BLAST hits στο terminal."),
            ("q3 3 summary.png", r"Output file: results\question3\summary.txt (μέρος 1)"),
            ("q3 4 summary .png", r"Output file: results\question3\summary.txt (μέρος 2)"),
        ],
        "bullets": [
            "Η αλληλουχία GIVEQCCTSICSLYQLENYCN ταυτοποιήθηκε ως Insulin A chain.",
            "Το πρώτο BLAST hit ανήκε σε Sus scrofa, αλλά υπήρχαν και ακριβή ανθρώπινα hits.",
            "Άρα η ταυτοποίηση της πρωτεΐνης είναι ασφαλής, αλλά η αλληλουχία δεν προσδιορίζει μοναδικά έναν μόνο οργανισμό λόγω μικρού μήκους και υψηλής συντήρησης.",
        ],
    },
    {
        "title": "Ερώτημα 4(i): Pairwise alignment spike πρωτεϊνών",
        "commands": [
            r"python src\q4_spike_alignment.py",
            r"python tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run",
        ],
        "images": [
            ("q4 i 1 .png", "Προετοιμασία των αλληλουχιών και εκτέλεση του EMBOSS Needle."),
            ("q4 i txt 2.png", r"Output file: results\question4\needle_run.aln.txt (summary metrics)"),
            ("q4 i txt 3.png", "Detailed pairwise alignment από το EMBOSS Needle."),
        ],
        "bullets": [
            "Συγκρίθηκαν οι spike πρωτεΐνες του SARS-CoV-2 και του SARS-CoV.",
            "Το αποτέλεσμα έδωσε identity 76.4%, similarity 87.0% και gaps 2.0%.",
            "Άρα οι δύο αλληλουχίες είναι συγγενικές αλλά όχι ίδιες.",
        ],
    },
    {
        "title": "Ερώτημα 4(ii): Δομική σύγκριση spike πρωτεϊνών",
        "commands": [
            r"python src\q4_structure_alignment.py",
        ],
        "images": [
            ("q4 ii 1 terminal.png", "Εκτέλεση του structure alignment script."),
            ("q4 ii 2 png.png", r"Output image: results\question4\rcsb_alignment_summary.png"),
        ],
        "bullets": [
            "Έγινε δομική σύγκριση των δύο spike μοντέλων μέσω RCSB structure alignment.",
            "Το αποτέλεσμα έδωσε TM-score 0.95 και RMSD 1.8.",
            "Άρα οι δύο δομές είναι πολύ κοντινές στον τρισδιάστατο χώρο.",
        ],
    },
    {
        "title": "Ερώτημα 5: Global και Local alignment",
        "commands": [
            r"python src\q5_alignment_dp.py",
        ],
        "images": [
            ("q5 1.png", "Εκτέλεση του script για Needleman-Wunsch και Smith-Waterman."),
            ("q5 2 txt global.png", r"Output file: results\question5\global_alignment.txt"),
            ("q5 3 txt local.png", r"Output file: results\question5\local_alignment.txt"),
        ],
        "bullets": [
            "Υλοποιήθηκαν global alignment με Needleman-Wunsch και local alignment με Smith-Waterman.",
            "Η ολική στοίχιση έδωσε optimal score -2.",
            "Η τοπική στοίχιση έδωσε optimal score 3 και ανέδειξε το κοινό μοτίβο GTG.",
        ],
    },
    {
        "title": "Ερώτημα 6(α): Suffix tree",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 1.png", "Θεωρητική λύση με suffix tree και αποθήκευση minPos."),
        ],
        "bullets": [
            "Η βασική ιδέα είναι η κατασκευή suffix tree του T$.",
            "Σε κάθε κόμβο αποθηκεύεται η ελάχιστη θέση εμφάνισης minPos.",
            "Έτσι πετυχαίνουμε preprocessing O(n) και query time O(m).",
        ],
    },
    {
        "title": "Ερώτημα 6(β): Bloom filters, Suffix Arrays και BWT",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 2.png", "Σύγκριση Bloom filters, suffix arrays και Burrows-Wheeler based indexing."),
        ],
        "bullets": [
            "Τα Bloom filters προσφέρουν πολύ γρήγορα membership queries με πολύ μικρό χώρο.",
            "Τα suffix arrays δίνουν ακριβές full-text indexing.",
            "Οι BWT/FM-index προσεγγίσεις κρατούν ακριβές indexing με καλύτερη συμπίεση και είναι πολύ χρήσιμες σε μεγάλα βιολογικά δεδομένα.",
        ],
    },
    {
        "title": "Συμπεράσματα",
        "bullets": [
            "Η εργασία συνδύασε πρακτική ανάλυση αλληλουχιών, σύγκριση δομών και θεωρητικούς αλγορίθμους indexing.",
            "Χρησιμοποιήθηκαν BLASTP, MAFFT, MUSCLE, T-Coffee, EMBOSS Needle, RCSB structure alignment και δυναμικός προγραμματισμός.",
            "Τα αποτελέσματα ήταν υπολογιστικά και βιολογικά συνεπή και δείχνουν τη σύνδεση sequence analysis, structural comparison και indexing στη βιοπληροφορική.",
        ],
    },
]


def set_normal_style(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)


def add_intro(document: Document) -> None:
    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_paragraph.add_run(SECTIONS[0]["title"])
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(22)

    for line in SECTIONS[0]["intro"]:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run(line)

    document.add_page_break()


def add_commands(document: Document, commands: list[str]) -> None:
    paragraph = document.add_paragraph()
    paragraph.add_run("Εντολή που τρέχουμε:").bold = True

    for command in commands:
        command_paragraph = document.add_paragraph()
        run = command_paragraph.add_run(command)
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)


def add_images(document: Document, images: list[tuple[str, str]]) -> None:
    for image_name, caption in images:
        image_path = SCREENSHOTS_DIR / image_name
        caption_paragraph = document.add_paragraph()
        caption_run = caption_paragraph.add_run(caption)
        caption_run.italic = True

        if image_path.exists():
            document.add_picture(str(image_path), width=Inches(6.0))
        else:
            missing = document.add_paragraph()
            missing.add_run(f"Λείπει screenshot: {image_name}").bold = True


def add_bullets(document: Document, bullets: list[str]) -> None:
    for bullet in bullets:
        document.add_paragraph(bullet, style="List Bullet")


def build_document() -> Path:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    set_normal_style(document)
    add_intro(document)

    content_sections = SECTIONS[1:]
    for index, content in enumerate(content_sections):
        document.add_heading(content["title"], level=1)

        source = content.get("source")
        if source:
            source_paragraph = document.add_paragraph()
            source_run = source_paragraph.add_run(source)
            source_run.italic = True

        commands = content.get("commands", [])
        if commands:
            add_commands(document, commands)

        images = content.get("images", [])
        if images:
            add_images(document, images)

        bullets = content.get("bullets", [])
        if bullets:
            add_bullets(document, bullets)

        if index != len(content_sections) - 1:
            document.add_page_break()

    document.save(OUTPUT_PATH)
    return OUTPUT_PATH


if __name__ == "__main__":
    path = build_document()
    print(path)
