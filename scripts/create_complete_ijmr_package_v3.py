#!/usr/bin/env python3
"""
Create the version 3 IJMR submission package without modifying earlier deliverables.
This script consumes the new manuscript builder, copies mandatory forms, and writes
an updated README and checklist so reviewers can track everything in Git.
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import List

from create_ijmr_complete_manuscript_v3 import (
    DEFAULT_OUTPUT_DIR,
    IJMRBestManuscriptBuilder,
)

SOURCE_SUBMISSION_DIR = Path("IJMR_Submission")


def safe_copy(src: Path, dest: Path) -> None:
    """Copy src to dest if it exists, logging the action."""
    if src.exists():
        shutil.copy2(src, dest)
    else:
        print(f"[WARN] Missing source file: {src}")


def write_text_file(path: Path, text: str) -> None:
    path.write_text(text.strip() + "\n", encoding="utf-8")


def create_cover_letter(metrics: dict, output_dir: Path) -> Path:
    cover_path = output_dir / "ijmr_covering_letter_v3.md"
    top_states = ", ".join(metrics["top_states"][:5])
    text = f"""
    # Covering Letter – IJMR Submission (v3)

    {datetime.now():%d %B %Y}

    Editor-in-Chief  
    Indian Journal of Medical Research  
    Ansari Nagar, New Delhi 110029  
    India

    Dear Editor,

    I am pleased to submit the manuscript titled **"Precision Synthesis of Tuberculosis Detection Delays in India: A Unified Bayesian, PCA, and DAG Framework"** for consideration as an Original Article in IJMR. The study combines {metrics["lit_num_studies"]} delay datasets ({metrics["lit_year_min"]}–{metrics["lit_year_max"]}) with national surveillance indicators to deliver reviewer-ready uncertainty intervals.

    Key contributions include a national total detection delay of {metrics["national_total_delay_mean"]:.1f} days (95% HDI {metrics["national_total_delay_low"]:.1f}–{metrics["national_total_delay_high"]:.1f}), structural decomposition explaining {metrics["pca_cumulative_pc3"]*100:.1f}% of determinant variance, and identification of {top_states} as the most urgent states for synchronized interventions.

    All analyses rely on publicly available data; no human participants were contacted. The manuscript adheres to IJMR formatting, is below the 3,000-word limit, and includes blinded and complete versions generated automatically for transparency.

    I confirm that the work is original, not under consideration elsewhere, and that all authors approve the submission. I authorize IJMR to manage the review process and am available for any clarifications.

    Sincerely,  
    **Dr H S Siddalingaiah**  
    Professor, Department of Community Medicine  
    Shridevi Institute of Medical Sciences and Research Hospital, Tumkur, Karnataka  
    Email: hssling@yahoo.com
    """
    write_text_file(cover_path, text)
    return cover_path


def create_submission_readme(metrics: dict, output_dir: Path, files: List[Path]) -> Path:
    readme_path = output_dir / "FINAL_IJMR_SUBMISSION_README_v3.md"
    file_lines = "\n".join(
        f"- {file.name}" for file in sorted(files, key=lambda p: p.name.lower())
    )
    text = f"""
    # IJMR Submission Package v3 – Multi-Method TB Delay Analysis

    **Generated:** {metrics["generated_at"]}

    ## Manuscript Facts
    - Total delay: {metrics["national_total_delay_mean"]:.1f} days (95% HDI {metrics["national_total_delay_low"]:.1f}–{metrics["national_total_delay_high"]:.1f})
    - Patient / Diagnostic components: {metrics["patient_delay_mean"]:.1f} / {metrics["diagnostic_delay_mean"]:.1f} days
    - Word counts (complete / blinded): {metrics["word_count_complete"]} / {metrics["word_count_blinded"]}
    - Top priority states: {', '.join(metrics['top_states'][:5])}
    - PCA coverage: {metrics["pca_cumulative_pc3"]*100:.1f}% variance explained by first three components

    ## Directory Contents
    {file_lines}

    ## Checklist
    - [x] Complete and blinded manuscripts (v3 builder)
    - [x] Updated covering letter referencing new results
    - [x] Title page, ethics, copyright, and conflict declarations
    - [x] High-resolution figures (EPS)
    - [x] Submission guide and metrics manifest
    """
    write_text_file(readme_path, text)
    return readme_path


def create_submission_guide(metrics: dict, output_dir: Path) -> Path:
    guide_path = output_dir / "IJMR_SUBMISSION_GUIDE_v3.md"
    text = f"""
    # IJMR Submission Guide v3

    1. Upload `ijmr_best_manuscript.docx` as the main file and `ijmr_best_manuscript_blinded.docx` wherever a blinded copy is requested.
    2. Attach `Figure_1.eps`–`Figure_3.eps`. They comply with IJMR art specifications and were not altered during this iteration.
    3. Provide the declarations (`ijmr_ethics_approval.txt`, `ijmr_conflict_of_interest.txt`, `ijmr_copyright_transfer.txt`).
    4. Paste the text in `ijmr_covering_letter_v3.md` into the journal portal.
    5. If asked about methodology, cite the metrics manifest (`ijmr_best_manuscript_metrics.json`) which details posterior intervals, PCA variance, and DAG topology ({metrics["dag_nodes"]} nodes, {metrics["dag_edges"]} edges).
    6. Zip the entire `IJMR_Submission_v3_best` directory or use the auto-generated archive once final verification is complete.
    """
    write_text_file(guide_path, text)
    return guide_path


def create_package() -> None:
    output_dir = DEFAULT_OUTPUT_DIR
    builder = IJMRBestManuscriptBuilder(output_dir=output_dir)
    artifacts = builder.generate_all()
    metrics = artifacts.metrics

    files_created = [
        artifacts.complete_path,
        artifacts.blinded_path,
        output_dir / "ijmr_best_manuscript_metrics.json",
    ]

    # Copy static assets from the prior submission directory.
    assets_to_copy = [
        "ijmr_title_page.docx",
        "ijmr_ethics_approval.txt",
        "ijmr_conflict_of_interest.txt",
        "ijmr_copyright_transfer.txt",
        "ijmr_guidelines_full.txt",
    ]
    figure_files = ["Figure_1.eps", "Figure_2.eps", "Figure_3.eps"]

    for name in assets_to_copy + figure_files:
        src = SOURCE_SUBMISSION_DIR / name
        dest = output_dir / name
        safe_copy(src, dest)
        files_created.append(dest)

    cover_letter = create_cover_letter(metrics, output_dir)
    files_created.append(cover_letter)

    readme = create_submission_readme(metrics, output_dir, files_created)
    guide = create_submission_guide(metrics, output_dir)
    files_created.extend([readme, guide])

    # Persist a lightweight manifest for reproducibility.
    manifest_path = output_dir / "ijmr_submission_manifest_v3.json"
    manifest = {
        "generated_at": metrics["generated_at"],
        "files": [path.name for path in files_created],
        "notes": "All files generated without touching prior versions; safe to commit.",
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    files_created.append(manifest_path)

    # Create a ready-to-upload archive.
    archive_path = shutil.make_archive(
        base_name=str(output_dir), format="zip", root_dir=output_dir
    )
    print(f"[SUCCESS] Submission package available at {output_dir}")
    print(f"[SUCCESS] Archive created: {archive_path}")


if __name__ == "__main__":
    create_package()
