from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


BASE_DIR = Path(r"C:\Users\dioni\Desktop\general\viopliroforiki ergasia 1")
REPORT_DIR = BASE_DIR / "report"
SCREENSHOTS_DIR = Path(r"C:\Users\dioni\Pictures\vioplitoforiki screenshots")
OUTPUT_PATH = REPORT_DIR / "viopliroforiki_parousiasi_ascii.docx"


SECTIONS = [
    {
        "title": "Viopliroforiki - Ergasia 1",
        "intro": [
            "Parousiasi apotelesmaton me vasi tin ektelesi tou kodika kai ta paragomena outputs.",
            "Se kathe enotita fainetai i entoli pou trexoume, opou yparxei ektelesi, kai akolouthoun ta antistoixa screenshots.",
        ],
    },
    {
        "title": "Erotima 1(i): AlphaFold 3 kai ESMFold",
        "bullets": [
            "Sto proto ypoerotima egine syntomi theoritiki sygkrisi ton ergaleion AlphaFold 3 kai ESMFold.",
            "To AlphaFold 3 einai pio geniko, epeidi ypostirizei provlepsi domis kai allilepidraseon diaforetikon biomorion.",
            "To ESMFold dinei poly grigori provlepsi proteinikis domis apefthias apo tin allilouxia.",
        ],
    },
    {
        "title": "Erotima 1(ii): BLASTP syllogi 10 hits",
        "commands": [
            r"python src\q1_blastp_collect.py",
        ],
        "images": [
            ("q1 ii results.png", "Ektelesi tou script gia BLASTP tis anthropinis hemoglobin beta."),
            ("q1 ii results top10 file.png", r"Output file: results\question1\hemoglobin_beta_top10.fasta"),
        ],
        "bullets": [
            "Xrisimopoiisa tin anthropini hemoglobin beta me accession NP_000509.",
            "To script ekane BLASTP, epelekse 10 sxetika proteinika hits kai ta apothikefse se arxeio FASTA.",
            "To arxeio auto xrisimopoiithike sto epomeno vima gia to multiple sequence alignment.",
        ],
    },
    {
        "title": "Erotima 1(iii): Multiple Sequence Alignment",
        "commands": [
            r"python tools\ebi_clients\mafft.py --email up1093449@ac.upatras.gr --stype protein --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\mafft_test",
            r"python tools\ebi_clients\muscle.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\muscle_run",
            r"python tools\ebi_clients\tcoffee.py --email up1093449@ac.upatras.gr --sequence results\question1\hemoglobin_beta_top10.fasta --outfile results\question1\msa\tcoffee_run",
        ],
        "images": [
            ("q1(iii) 1 terminal.png", "Ektelesi ton jobs gia MAFFT, MUSCLE kai T-Coffee."),
            ("q1(iii) 2 result file.png", r"Output file: results\question1\msa\muscle_run.aln-clustalw.clw"),
        ],
        "bullets": [
            "Ekana multiple sequence alignment ton 10 allilouxion me tria diaforetika ergaleia.",
            "Ta apotelesmata itan poly paromoia kai edeixan ypsili syntirisi tis hemoglobin beta.",
            "To alignment tou MUSCLE xrisimopoiithike os antiprosopftiko paradeigma sto output.",
        ],
    },
    {
        "title": "Erotima 2: Cytochrome c sta Mammalia",
        "source": "Source: UniProt entry P99999 (CYC_HUMAN)",
        "commands": [
            r"python src\q2_cytochrome_c.py",
        ],
        "images": [
            ("q2 3.png", "Selida UniProt gia tin anthropini Cytochrome c."),
            ("q2 1.png", "Ektelesi tou script gia BLASTP sta Mammalia."),
            ("q2 2.png", r"Output file: results\question2\summary.txt"),
        ],
        "bullets": [
            "I anthropini Cytochrome c anaktithike apo to UniProt me accession P99999.",
            "Egine BLASTP sta Mammalia kai ypologistikan identity percentages os pros tin anthropini allilouxia.",
            "Oi times itan peripou 98% eos 100%, ara i proteini einai exairetika syntirimeni sta thilastika.",
        ],
    },
    {
        "title": "Erotima 3: Tautopoiisi agnostis allilouxias",
        "commands": [
            r"python src\q3_identify_protein.py",
        ],
        "images": [
            ("q3 1.png", "Ektelesi tou script gia BLASTP tis agnostis allilouxias."),
            ("q3 2.png", "Synecheia ton top BLAST hits sto terminal."),
            ("q3 3 summary.png", r"Output file: results\question3\summary.txt (meros 1)"),
            ("q3 4 summary .png", r"Output file: results\question3\summary.txt (meros 2)"),
        ],
        "bullets": [
            "I allilouxia GIVEQCCTSICSLYQLENYCN tautopoiithike os Insulin A chain.",
            "To proto BLAST hit anike se Sus scrofa, alla ypirxan kai akrivi anthropina hits.",
            "Ara i tautopoiisi tis proteinis einai asfalis, alla i allilouxia den prosdiorizei monadika enan mono organismo logo mikrou mikous kai ypsilis syntirisis.",
        ],
    },
    {
        "title": "Erotima 4(i): Pairwise alignment spike proteinon",
        "commands": [
            r"python src\q4_spike_alignment.py",
            r"python tools\ebi_clients\emboss_needle.py --email up1093449@ac.upatras.gr --stype protein --asequence results\question4\sars_cov_2_spike.fasta --bsequence results\question4\sars_cov_spike.fasta --outfile results\question4\needle_run",
        ],
        "images": [
            ("q4 i 1 .png", "Proetoimasia ton allilouxion kai ektelesi tou EMBOSS Needle."),
            ("q4 i txt 2.png", r"Output file: results\question4\needle_run.aln.txt (summary metrics)"),
            ("q4 i txt 3.png", "Detailed pairwise alignment apo to EMBOSS Needle."),
        ],
        "bullets": [
            "Synkrithikan oi spike proteinis tou SARS-CoV-2 kai tou SARS-CoV.",
            "To apotelesma edose identity 76.4%, similarity 87.0% kai gaps 2.0%.",
            "Ara oi dyo allilouxies einai syngenikes alla oxi idies.",
        ],
    },
    {
        "title": "Erotima 4(ii): Domiki sygkrisi spike proteinon",
        "commands": [
            r"python src\q4_structure_alignment.py",
        ],
        "images": [
            ("q4 ii 1 terminal.png", "Ektelesi tou structure alignment script."),
            ("q4 ii 2 png.png", r"Output image: results\question4\rcsb_alignment_summary.png"),
        ],
        "bullets": [
            "Egine domiki sygkrisi ton dyo spike montelon meso RCSB structure alignment.",
            "To apotelesma edose TM-score 0.95 kai RMSD 1.8.",
            "Ara oi dyo domes einai poly kontines ston tridastato xoro.",
        ],
    },
    {
        "title": "Erotima 5: Global kai Local alignment",
        "commands": [
            r"python src\q5_alignment_dp.py",
        ],
        "images": [
            ("q5 1.png", "Ektelesi tou script gia Needleman-Wunsch kai Smith-Waterman."),
            ("q5 2 txt global.png", r"Output file: results\question5\global_alignment.txt"),
            ("q5 3 txt local.png", r"Output file: results\question5\local_alignment.txt"),
        ],
        "bullets": [
            "Ylopoiithikan global alignment me Needleman-Wunsch kai local alignment me Smith-Waterman.",
            "I oliki stoixisi edose optimal score -2.",
            "I topiki stoixisi edose optimal score 3 kai anedeixe to koino motibo GTG.",
        ],
    },
    {
        "title": "Erotima 6(a): Suffix tree",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 1.png", "Theoritiki lysi me suffix tree kai apothikefsi minPos."),
        ],
        "bullets": [
            "I vasiki idea einai i kataskevi suffix tree tou T$.",
            "Se kathe komvo apothikevetai i elaxisti thesi emfanisis minPos.",
            "Etsi petyxainoume preprocessing O(n) kai query time O(m).",
        ],
    },
    {
        "title": "Erotima 6(b): Bloom filters, Suffix Arrays kai BWT",
        "source": r"Theory file: report\question6_theory.md",
        "images": [
            ("q6 2.png", "Sygkrisi Bloom filters, suffix arrays kai Burrows-Wheeler based indexing."),
        ],
        "bullets": [
            "Ta Bloom filters prosferoun poly grigora membership queries me poly mikro xoro.",
            "Ta suffix arrays dinoun akrives full-text indexing.",
            "Oi BWT/FM-index proseggiseis kratoun akrives indexing me kalyteri sympiesi kai einai poly xrisimes se megala viologika dedomena.",
        ],
    },
    {
        "title": "Symperasmata",
        "bullets": [
            "I ergasia syndyase praktiki analysi allilouxion, sygkrisi domon kai theoritikous algorithmous indexing.",
            "Xrisimopoiithikan BLASTP, MAFFT, MUSCLE, T-Coffee, EMBOSS Needle, RCSB structure alignment kai dynamikos programmatismos.",
            "Ta apotelesmata itan ypologistika kai viologika syneti kai deixnoun ti syndesi sequence analysis, structural comparison kai indexing sti viopliroforiki.",
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
    paragraph.add_run("Entoli pou trexoume:").bold = True

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
            missing.add_run(f"Leipei screenshot: {image_name}").bold = True


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
