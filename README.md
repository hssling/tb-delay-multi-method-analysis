# TB Delay Multi-Method Analysis

This repository hosts the IJMR submission package for the India TB detection delay study that combines Bayesian meta-analysis, principal component analysis, and directed acyclic graphs to prioritize states and intervention levers.

## Contents
- `manuscripts/ijmr_best_manuscript_v13_imrad.docx`: final IMRAD manuscript with in-text cross-references and numbered tables/figures.
- `scripts/`: analysis and manuscript-generation pipeline (data prep, meta-analysis, PCA, DAG, figure and manuscript builders).
- `submission/`: cover letter, declarations, title page, submission manifest, and checklist.
- `figures/`: Figures 1–4 in PNG and EPS formats (Bayesian dashboard, PCA diagnostics, DAG, state coverage map).
- `supporting/ijmr_best_manuscript_metrics.json`: posterior and PCA metrics referenced in the manuscript.
- `environment.yml`, `requirements.txt`: dependencies for reproducing analyses and document builds.
- `LICENSE`: license for this repository.

## Usage
- Cite this repo link in the manuscript for access to figures and submission artefacts.
- Use the EPS figures for journal upload; PNGs are included for quick review.
- The manifest (`submission/ijmr_submission_manifest_v4.json`) lists the submission files and formats.
- Run `scripts/run_all.py` (or `scripts/run_all_v2.py`) to execute the end-to-end pipeline; individual numbered scripts in `scripts/` correspond to specific stages (ingest, clean, meta-analysis, PCA, DAG, figures, manuscript assembly).

## Authors
- H S Siddalingaiah (hssling@yahoo.com)

## Reproducibility notes
- Manuscript references start in the Introduction; tables/figures are cross-referenced in text.
- Bayesian models were fit in NumPyro/PyMC; see metrics JSON for posterior summaries.

## Repository link for manuscript
- GitHub: https://github.com/hssling/tb-delay-multi-method-analysis
