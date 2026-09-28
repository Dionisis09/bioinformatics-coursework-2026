from __future__ import annotations

import json
import time
from pathlib import Path

import requests


OUTPUT_DIR = Path("results/question4")
SARS2_PDB = OUTPUT_DIR / "sars2_swissmodel.pdb"
SARS1_PDB = OUTPUT_DIR / "sars1_swissmodel.pdb"
RCSB_JSON = OUTPUT_DIR / "rcsb_structure_alignment.json"
RCSB_SUMMARY = OUTPUT_DIR / "rcsb_structure_summary.txt"


def submit_job() -> str:
    query = {
        "context": {
            "mode": "pairwise",
            "method": {"name": "tm-align"},
            "structures": [
                {"format": "pdb"},
                {"format": "pdb"},
            ],
        }
    }

    with open(SARS2_PDB, "rb") as f1, open(SARS1_PDB, "rb") as f2:
        files = [
            ("files", ("sars2_swissmodel.pdb", f1, "chemical/x-pdb")),
            ("files", ("sars1_swissmodel.pdb", f2, "chemical/x-pdb")),
        ]
        response = requests.post(
            "https://alignment.rcsb.org/api/v1/structures/submit",
            params={"query": json.dumps(query)},
            files=files,
            timeout=120,
        )
    response.raise_for_status()
    return response.text.strip()


def poll_job(uuid: str) -> dict:
    url = "https://alignment.rcsb.org/api/v1/structures/results"
    for _ in range(90):
        response = requests.get(url, params={"uuid": uuid}, timeout=120)
        if response.status_code == 404:
            time.sleep(5)
            continue
        response.raise_for_status()
        data = response.json()
        status = data.get("info", {}).get("status")
        if status == "COMPLETE":
            return data
        if status == "ERROR":
            raise RuntimeError(data.get("info", {}).get("message", "Unknown RCSB alignment error"))
        time.sleep(5)
    raise TimeoutError("Timed out while waiting for RCSB structure alignment job.")


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    uuid = submit_job()
    result = poll_job(uuid)
    RCSB_JSON.write_text(json.dumps(result, indent=2), encoding="utf-8")

    summary = result["results"][0]["summary"]
    scores = {item["type"]: item["value"] for item in summary["scores"]}
    text = "\n".join(
        [
            "Question 4(ii) RCSB structure alignment",
            f"Job UUID: {uuid}",
            f"TM-score: {scores.get('TM-score')}",
            f"RMSD: {scores.get('RMSD')}",
            f"Sequence identity: {scores.get('sequence-identity')}",
            f"Aligned residue pairs: {summary.get('n_aln_residue_pairs')}",
            f"Sequence alignment length: {summary.get('seq_aln_len')}",
            f"Alignment coverage structure 1: {summary.get('aln_coverage')[0]}%",
            f"Alignment coverage structure 2: {summary.get('aln_coverage')[1]}%",
            f"Modeled residues structure 1: {summary.get('n_modeled_residues')[0]}",
            f"Modeled residues structure 2: {summary.get('n_modeled_residues')[1]}",
            "",
        ]
    )
    RCSB_SUMMARY.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
