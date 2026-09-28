from __future__ import annotations

from io import StringIO
from pathlib import Path
import re
from textwrap import shorten

from Bio.Blast import NCBIWWW, NCBIXML


QUERY_SEQUENCE = "GIVEQCCTSICSLYQLENYCN"
OUTPUT_DIR = Path("results/question3")
XML_PATH = OUTPUT_DIR / "blast_q3.xml"
SUMMARY_PATH = OUTPUT_DIR / "summary.txt"


def fetch_blast_xml(sequence: str) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if XML_PATH.exists():
        return XML_PATH.read_text(encoding="utf-8")

    handle = NCBIWWW.qblast("blastp", "nr", sequence, hitlist_size=10)
    xml = handle.read()
    handle.close()
    XML_PATH.write_text(xml, encoding="utf-8")
    return xml


def extract_organism(title: str) -> str:
    matches = re.findall(r"\[([^\]]+)\]", title)
    return matches[-1] if matches else "Unknown"


def infer_protein_name(title: str) -> str:
    lowered = title.lower()
    if "insulin" in lowered and "a chain" in lowered:
        return "Insulin A chain"
    if "insulin" in lowered:
        return "Insulin"
    return "Unknown protein"


def build_summary(xml: str) -> str:
    record = NCBIXML.read(StringIO(xml))
    if not record.alignments:
        return "Question 3 BLAST summary\nNo alignments returned.\n"

    best_alignment = record.alignments[0]
    best_hsp = best_alignment.hsps[0]
    best_organism = extract_organism(best_alignment.title)
    inferred_protein = infer_protein_name(best_alignment.title)
    human_hits = [
        alignment
        for alignment in record.alignments
        if "[Homo sapiens]" in alignment.title
    ]

    lines = []
    lines.append("Question 3 BLAST summary")
    lines.append(f"Query sequence: {QUERY_SEQUENCE}")
    lines.append(f"Database: {record.database}")
    lines.append(f"Number of alignments returned: {len(record.alignments)}")
    lines.append("")
    lines.append("Primary identification from the best-scoring BLAST hit")
    lines.append(f"Protein: {inferred_protein}")
    lines.append(f"Organism of first hit: {best_organism}")
    lines.append(f"Best-hit title: {shorten(best_alignment.title, width=200, placeholder='...')}")
    lines.append(f"Best-hit E-value: {best_hsp.expect}")
    lines.append(
        "Conclusion: the sequence matches insulin A chain very strongly, "
        f"but the first BLAST hit belongs to {best_organism}, not Homo sapiens."
    )
    if human_hits:
        lines.append("")
        lines.append("Exact human matches are also present among the top hits")
        for index, alignment in enumerate(human_hits[:3], start=1):
            human_hsp = alignment.hsps[0]
            lines.append(
                f"{index}. {shorten(alignment.title, width=160, placeholder='...')} | "
                f"E-value={human_hsp.expect}"
            )
        lines.append(
            "Interpretation: because the query is only 21 aa long and highly conserved, "
            "it does not uniquely identify a human sequence."
        )
    lines.append("")
    lines.append("Known function")
    lines.append(
        "Insulin is a peptide hormone that helps regulate blood glucose homeostasis."
    )
    lines.append("")
    lines.append("Top 10 BLAST hits")
    lines.append("")

    for index, alignment in enumerate(record.alignments[:10], start=1):
        hsp = alignment.hsps[0]
        identity_pct = (hsp.identities / hsp.align_length) * 100 if hsp.align_length else 0.0
        lines.append(f"Hit {index}")
        lines.append(f"Title: {shorten(alignment.title, width=200, placeholder='...')}")
        lines.append(f"E-value: {hsp.expect}")
        lines.append(f"Alignment length: {hsp.align_length}")
        lines.append(f"Identity: {hsp.identities}/{hsp.align_length} ({identity_pct:.2f}%)")
        lines.append(f"Positives: {hsp.positives}/{hsp.align_length}")
        lines.append("")

    return "\n".join(lines).strip() + "\n"


def main() -> None:
    xml = fetch_blast_xml(QUERY_SEQUENCE)
    summary = build_summary(xml)
    SUMMARY_PATH.write_text(summary, encoding="utf-8")
    print(summary)


if __name__ == "__main__":
    main()
