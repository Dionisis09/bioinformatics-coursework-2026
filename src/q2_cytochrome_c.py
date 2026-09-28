from __future__ import annotations

from dataclasses import dataclass
from io import StringIO
from pathlib import Path
from textwrap import shorten

import requests
from Bio import Entrez, SeqIO
from Bio.Blast import NCBIWWW, NCBIXML
from Bio.Align import PairwiseAligner
from Bio.SeqRecord import SeqRecord


UNIPROT_ACCESSION = "P99999"
UNIPROT_FASTA_URL = f"https://rest.uniprot.org/uniprotkb/{UNIPROT_ACCESSION}.fasta"
OUTPUT_DIR = Path("results/question2")
QUERY_FASTA_PATH = OUTPUT_DIR / "human_cytochrome_c.fasta"
BLAST_XML_PATH = OUTPUT_DIR / "blast_mammalia.xml"
HITS_FASTA_PATH = OUTPUT_DIR / "mammalia_hits.fasta"
IDENTITY_TABLE_PATH = OUTPUT_DIR / "identity_table.tsv"
SUMMARY_PATH = OUTPUT_DIR / "summary.txt"

Entrez.email = "up1093449@ac.upatras.gr"


@dataclass
class HitResult:
    accession: str
    title: str
    evalue: float
    identities: int
    align_length: int


def fetch_uniprot_fasta() -> SeqRecord:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if QUERY_FASTA_PATH.exists():
        return SeqIO.read(QUERY_FASTA_PATH, "fasta")

    response = requests.get(UNIPROT_FASTA_URL, timeout=60)
    response.raise_for_status()
    QUERY_FASTA_PATH.write_text(response.text, encoding="utf-8")
    return SeqIO.read(StringIO(response.text), "fasta")


def fetch_blast_xml(sequence: str) -> str:
    if BLAST_XML_PATH.exists():
        return BLAST_XML_PATH.read_text(encoding="utf-8")

    handle = NCBIWWW.qblast(
        "blastp",
        "nr",
        sequence,
        entrez_query="Mammalia[Organism]",
        hitlist_size=12,
    )
    xml = handle.read()
    handle.close()
    BLAST_XML_PATH.write_text(xml, encoding="utf-8")
    return xml


def parse_hits(xml: str, query_record: SeqRecord) -> list[HitResult]:
    record = NCBIXML.read(StringIO(xml))
    hits: list[HitResult] = []
    query_name = "cytochrome c"

    for alignment in record.alignments:
        if not alignment.hsps:
            continue

        hsp = alignment.hsps[0]
        title_lower = alignment.title.lower()

        if query_name not in title_lower:
            continue
        if "homo sapiens" in title_lower:
            continue

        hits.append(
            HitResult(
                accession=alignment.accession,
                title=alignment.title,
                evalue=hsp.expect,
                identities=hsp.identities,
                align_length=hsp.align_length,
            )
        )

        if len(hits) >= 10:
            break

    return hits


def fetch_hit_sequences(accessions: list[str]) -> list[SeqRecord]:
    if HITS_FASTA_PATH.exists():
        return list(SeqIO.parse(HITS_FASTA_PATH, "fasta"))

    if not accessions:
        return []

    with Entrez.efetch(
        db="protein",
        id=",".join(accessions),
        rettype="fasta",
        retmode="text",
    ) as handle:
        records = list(SeqIO.parse(handle, "fasta"))

    SeqIO.write(records, HITS_FASTA_PATH, "fasta")
    return records


def compute_identity(query_record: SeqRecord, hit_records: list[SeqRecord]) -> list[tuple[str, float, int, str]]:
    aligner = PairwiseAligner()
    aligner.mode = "global"
    aligner.match_score = 1
    aligner.mismatch_score = 0
    aligner.open_gap_score = -1
    aligner.extend_gap_score = -1

    query_seq = str(query_record.seq)
    results: list[tuple[str, float, int, str]] = []

    for record in hit_records:
        alignment = aligner.align(query_seq, str(record.seq))[0]
        aligned_query, aligned_hit = str(alignment[0]), str(alignment[1])
        matches = sum(
            1
            for a, b in zip(aligned_query, aligned_hit)
            if a == b and a != "-" and b != "-"
        )
        aligned_positions = sum(
            1
            for a, b in zip(aligned_query, aligned_hit)
            if a != "-" and b != "-"
        )
        identity = (matches / aligned_positions) * 100 if aligned_positions else 0.0
        results.append((record.id, identity, aligned_positions, record.description))

    return sorted(results, key=lambda row: row[1], reverse=True)


def write_outputs(query_record: SeqRecord, hits: list[HitResult], identity_rows: list[tuple[str, float, int, str]]) -> str:
    table_lines = ["record_id\tidentity_percent\taligned_positions\tdescription"]
    for record_id, identity, aligned_positions, description in identity_rows:
        table_lines.append(
            f"{record_id}\t{identity:.2f}\t{aligned_positions}\t{description}"
        )
    IDENTITY_TABLE_PATH.write_text("\n".join(table_lines) + "\n", encoding="utf-8")

    lines = []
    lines.append("Question 2 summary")
    lines.append(f"Human protein: Cytochrome c")
    lines.append(f"UniProt accession: {UNIPROT_ACCESSION}")
    lines.append(f"Query FASTA header: {query_record.description}")
    lines.append("")
    lines.append("Top retained mammalian BLAST hits")
    for index, hit in enumerate(hits, start=1):
        lines.append(
            f"{index}. {shorten(hit.title, width=200, placeholder='...')} | accession={hit.accession} | e-value={hit.evalue}"
        )
    lines.append("")
    lines.append("Identity percentages against the human sequence")
    for record_id, identity, aligned_positions, description in identity_rows:
        lines.append(
            f"- {record_id}: {identity:.2f}% over {aligned_positions} aligned positions | {shorten(description, width=160, placeholder='...')}"
        )

    summary = "\n".join(lines).strip() + "\n"
    SUMMARY_PATH.write_text(summary, encoding="utf-8")
    return summary


def main() -> None:
    query_record = fetch_uniprot_fasta()
    xml = fetch_blast_xml(str(query_record.seq))
    hits = parse_hits(xml, query_record)
    hit_records = fetch_hit_sequences([hit.accession for hit in hits])
    identity_rows = compute_identity(query_record, hit_records)
    summary = write_outputs(query_record, hits, identity_rows)
    print(summary)


if __name__ == "__main__":
    main()
