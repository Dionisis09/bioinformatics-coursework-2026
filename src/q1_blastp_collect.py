from __future__ import annotations

from io import StringIO
from pathlib import Path
import re

from Bio import Entrez, SeqIO
from Bio.Blast import NCBIWWW, NCBIXML
from Bio.SeqRecord import SeqRecord


QUERY_ACCESSION = "NP_000509"
OUTPUT_DIR = Path("results/question1")
QUERY_FASTA_PATH = OUTPUT_DIR / "hemoglobin_beta_query.fasta"
BLAST_XML_PATH = OUTPUT_DIR / "hemoglobin_beta_blast.xml"
HITS_FASTA_PATH = OUTPUT_DIR / "hemoglobin_beta_top10.fasta"
SUMMARY_PATH = OUTPUT_DIR / "hemoglobin_beta_summary.txt"

Entrez.email = "up1093449@ac.upatras.gr"


def fetch_query() -> SeqRecord:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if QUERY_FASTA_PATH.exists():
        return SeqIO.read(QUERY_FASTA_PATH, "fasta")

    with Entrez.efetch(
        db="protein",
        id=QUERY_ACCESSION,
        rettype="fasta",
        retmode="text",
    ) as handle:
        record = SeqIO.read(handle, "fasta")

    SeqIO.write(record, QUERY_FASTA_PATH, "fasta")
    return record


def fetch_blast_xml(sequence: str, force_refresh: bool = False) -> str:
    if BLAST_XML_PATH.exists() and not force_refresh:
        return BLAST_XML_PATH.read_text(encoding="utf-8")

    handle = NCBIWWW.qblast("blastp", "nr", sequence, hitlist_size=100)
    xml = handle.read()
    handle.close()
    BLAST_XML_PATH.write_text(xml, encoding="utf-8")
    return xml


def select_accessions(xml: str) -> list[tuple[str, str, float]]:
    record = NCBIXML.read(StringIO(xml))
    selected: list[tuple[str, str, float]] = []
    seen_species: set[str] = set()

    for alignment in record.alignments:
        if not alignment.hsps:
            continue
        title = alignment.title.lower()
        if QUERY_ACCESSION.lower() in title:
            continue
        if "synthetic construct" in title:
            continue
        if "partial" in title:
            continue
        if title.startswith("pdb|"):
            continue
        if "[homo sapiens]" in title:
            continue

        species_match = re.findall(r"\[([^\]]+)\]", alignment.title)
        species = species_match[-1].strip() if species_match else "unknown"
        if species in seen_species:
            continue

        selected.append((alignment.accession, alignment.title, alignment.hsps[0].expect))
        seen_species.add(species)
        if len(selected) >= 10:
            break

    return selected


def fetch_hit_fastas(accessions: list[str]) -> list[SeqRecord]:
    with Entrez.efetch(
        db="protein",
        id=",".join(accessions),
        rettype="fasta",
        retmode="text",
    ) as handle:
        records = list(SeqIO.parse(handle, "fasta"))

    SeqIO.write(records, HITS_FASTA_PATH, "fasta")
    return records


def write_summary(query_record: SeqRecord, selected: list[tuple[str, str, float]]) -> str:
    lines = []
    lines.append("Question 1(ii) BLASTP collection")
    lines.append(f"Query accession: {QUERY_ACCESSION}")
    lines.append(f"Query description: {query_record.description}")
    lines.append("")
    lines.append("Selected 10 protein hits")
    for index, (accession, title, evalue) in enumerate(selected, start=1):
        lines.append(f"{index}. accession={accession} | e-value={evalue} | {title}")
    summary = "\n".join(lines).strip() + "\n"
    SUMMARY_PATH.write_text(summary, encoding="utf-8")
    return summary


def main() -> None:
    query_record = fetch_query()
    xml = fetch_blast_xml(str(query_record.seq))
    selected = select_accessions(xml)
    if len(selected) < 10:
        xml = fetch_blast_xml(str(query_record.seq), force_refresh=True)
        selected = select_accessions(xml)
    if len(selected) < 10:
        raise RuntimeError(f"Expected 10 filtered hits, found only {len(selected)}")
    fetch_hit_fastas([accession for accession, _, _ in selected])
    summary = write_summary(query_record, selected)
    print(summary)


if __name__ == "__main__":
    main()
