from __future__ import annotations

from pathlib import Path

import requests
from Bio import Entrez, SeqIO


NCBI_NUCCORE_ID = "NC_045512"
UNIPROT_SARS_COV_SPIKE = "P59594"
OUTPUT_DIR = Path("results/question4")
SARS2_FASTA = OUTPUT_DIR / "sars_cov_2_spike.fasta"
SARS1_FASTA = OUTPUT_DIR / "sars_cov_spike.fasta"
SUMMARY_PATH = OUTPUT_DIR / "sequence_summary.txt"

Entrez.email = "up1093449@ac.upatras.gr"


def fetch_sars_cov_2_spike() -> str:
    if SARS2_FASTA.exists():
        return SARS2_FASTA.read_text(encoding="utf-8")

    with Entrez.efetch(
        db="nucleotide",
        id=NCBI_NUCCORE_ID,
        rettype="gb",
        retmode="text",
    ) as handle:
        record = SeqIO.read(handle, "genbank")

    spike_translation = None
    for feature in record.features:
        if feature.type != "CDS":
            continue
        gene = feature.qualifiers.get("gene", [""])[0].lower()
        product = feature.qualifiers.get("product", [""])[0].lower()
        if gene == "s" or "spike" in product:
            spike_translation = feature.qualifiers.get("translation", [None])[0]
            break

    if spike_translation is None:
        raise RuntimeError("Could not locate SARS-CoV-2 spike CDS translation.")

    fasta = (
        f">SARS-CoV-2_spike|{NCBI_NUCCORE_ID}\n"
        + "\n".join(spike_translation[i : i + 70] for i in range(0, len(spike_translation), 70))
        + "\n"
    )
    SARS2_FASTA.write_text(fasta, encoding="utf-8")
    return fasta


def fetch_sars_cov_spike() -> str:
    if SARS1_FASTA.exists():
        return SARS1_FASTA.read_text(encoding="utf-8")

    url = f"https://rest.uniprot.org/uniprotkb/{UNIPROT_SARS_COV_SPIKE}.fasta"
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    SARS1_FASTA.write_text(response.text, encoding="utf-8")
    return response.text


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    sars2 = fetch_sars_cov_2_spike()
    sars1 = fetch_sars_cov_spike()

    sars2_record = SeqIO.read(SARS2_FASTA, "fasta")
    sars1_record = SeqIO.read(SARS1_FASTA, "fasta")

    summary = "\n".join(
        [
            "Question 4 sequence preparation",
            f"SARS-CoV-2 spike length: {len(sars2_record.seq)} aa",
            f"SARS-CoV spike length: {len(sars1_record.seq)} aa",
            f"SARS-CoV-2 source: NCBI nuccore {NCBI_NUCCORE_ID}",
            f"SARS-CoV source: UniProt {UNIPROT_SARS_COV_SPIKE}",
            "",
            "Files:",
            f"- {SARS2_FASTA}",
            f"- {SARS1_FASTA}",
            "",
        ]
    )
    SUMMARY_PATH.write_text(summary, encoding="utf-8")
    print(summary)


if __name__ == "__main__":
    main()
