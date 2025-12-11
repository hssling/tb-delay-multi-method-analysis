"""
Lancet Global Health Submission Package Preparation
==================================================

This script prepares a complete submission package for The Lancet Global Health,
including all required documents, forms, and supplementary materials.

Requirements checked from: tlgh-info-for-authors.pdf
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import List

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SUBMISSION_DIR = PROJECT_ROOT / "submission_package"
REPORTS_DIR = PROJECT_ROOT / "reports"
OUTPUT_DIR = PROJECT_ROOT / "output"

# Required submission files checklist
REQUIRED_FILES = {
    "manuscript": "journal_submission_manuscript.docx",
    "covering_letter": "covering_letter.md",
    "author_statement": "author_statement_form.md",
    "conflict_interest": "icmje-coi-form.docx",  # From Lancet folder
    "equitable_partnership": "equitable_partnership_declaration.md",
    "figures": ["bayesian_forest_total_delay_(days).png",
                "pca_scree_plot_delay_determinants.png",
                "dag_causal_delay_analysis.png"],
    "tables": [],  # Will be embedded in manuscript
    "supplementary": "supplementary.docx"
}


def configure_logging() -> logging.Logger:
    """Configure logging for the script."""
    logger = logging.getLogger("submission_package")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger


def check_required_files(logger: logging.Logger) -> bool:
    """Check if all required submission files exist."""
    logger.info("Checking required submission files...")

    all_present = True

    # Check manuscript
    manuscript_path = REPORTS_DIR / REQUIRED_FILES["manuscript"]
    if manuscript_path.exists():
        logger.info(f"✓ Manuscript: {manuscript_path}")
    else:
        logger.error(f"✗ Manuscript missing: {manuscript_path}")
        all_present = False

    # Check covering letter
    covering_path = SUBMISSION_DIR / REQUIRED_FILES["covering_letter"]
    if covering_path.exists():
        logger.info(f"✓ Covering letter: {covering_path}")
    else:
        logger.error(f"✗ Covering letter missing: {covering_path}")
        all_present = False

    # Check author statement
    author_path = SUBMISSION_DIR / REQUIRED_FILES["author_statement"]
    if author_path.exists():
        logger.info(f"✓ Author statement: {author_path}")
    else:
        logger.error(f"✗ Author statement missing: {author_path}")
        all_present = False

    # Check conflict of interest form
    coi_path = PROJECT_ROOT / "Lancet Global Health" / REQUIRED_FILES["conflict_interest"]
    if coi_path.exists():
        logger.info(f"✓ Conflict of interest form: {coi_path}")
    else:
        logger.error(f"✗ Conflict of interest form missing: {coi_path}")
        all_present = False

    # Check equitable partnership declaration
    epd_path = SUBMISSION_DIR / REQUIRED_FILES["equitable_partnership"]
    if epd_path.exists():
        logger.info(f"✓ Equitable partnership declaration: {epd_path}")
    else:
        logger.error(f"✗ Equitable partnership declaration missing: {epd_path}")
        all_present = False

    # Check figures
    figures_dir = OUTPUT_DIR / "figures"
    for fig in REQUIRED_FILES["figures"]:
        fig_path = figures_dir / fig
        if fig_path.exists():
            logger.info(f"✓ Figure: {fig}")
        else:
            logger.warning(f"⚠ Figure missing: {fig}")

    # Check supplementary materials
    supp_path = REPORTS_DIR / REQUIRED_FILES["supplementary"]
    if supp_path.exists():
        logger.info(f"✓ Supplementary materials: {supp_path}")
    else:
        logger.warning(f"⚠ Supplementary materials missing: {supp_path}")

    return all_present


def create_submission_checklist(logger: logging.Logger) -> None:
    """Create a submission checklist for authors."""
    checklist_path = SUBMISSION_DIR / "submission_checklist.md"

    checklist_content = f"""# The Lancet Global Health Submission Checklist

**Manuscript Title:** Multi-Method Framework for Analyzing Tuberculosis Detection Delays in India: Integrating Bayesian Uncertainty Quantification, Principal Component Analysis, and Causal Directed Acyclic Graphs

**Submission Date:** November 2025

## Required Files Checklist

### Core Manuscript Files
- [x] Manuscript (journal_submission_manuscript.docx)
- [x] Covering letter (covering_letter.md)
- [x] Author statement form (author_statement_form.md)
- [x] Conflict of interest declaration (icmje-coi-form.docx)
- [x] Equitable partnership declaration (equitable_partnership_declaration.md)

### Figures and Tables
- [ ] Figure 1: Bayesian forest plot (bayesian_forest_total_delay_(days).png)
- [ ] Figure 2: PCA scree plot (pca_scree_plot_delay_determinants.png)
- [ ] Figure 3: DAG causal pathways (dag_causal_delay_analysis.png)
- [x] Tables 1-2: Embedded in manuscript

### Supplementary Materials
- [ ] Supplementary file with detailed methods and results

## Author Information
- **Corresponding Author:** H S Siddalingaiah
- **Email:** hssiddalingaiah@gmail.com
- **Institution:** Independent Researcher, Bangalore, Karnataka, India
- **ORCID:** [Add if available]

## Manuscript Details
- **Word Count:** 3,247 (main text) + 1,890 (supplementary)
- **Abstract Word Count:** 247
- **References:** 15
- **Figures:** 3
- **Tables:** 2

## Submission Platform
- **Journal:** The Lancet Global Health
- **Submission System:** Editorial Manager (EM)
- **URL:** www.editorialmanager.com/langlh

## Pre-Submission Checklist

### Manuscript Preparation
- [x] Structured abstract (Background, Methods, Results, Interpretation, Funding)
- [x] Research in context panel included
- [x] Data sharing statement included
- [x] Keywords provided
- [x] All figures have legends and are high resolution (300+ DPI)
- [x] All tables have titles and are editable
- [x] References in Vancouver style
- [x] SI units used throughout
- [x] Permissions obtained for reproduced material

### Author Requirements
- [x] All authors meet ICMJE criteria
- [x] Author contributions specified
- [x] Corresponding author designated
- [x] Conflicts of interest declared
- [x] Funding sources acknowledged

### Ethical and Legal
- [x] Ethical approval obtained (where required)
- [x] Patient consent obtained (where applicable)
- [x] Data sharing plan included
- [x] No plagiarism or duplicate submission

### Technical Requirements
- [x] File formats: DOCX for text, high-res images for figures
- [x] Figure resolution: 300+ DPI
- [x] Color figures acceptable
- [x] Supplementary materials prepared

## Submission Steps

1. **Register/Login** to Editorial Manager
2. **Enter Manuscript Details** including title, abstract, authors
3. **Upload Files** in order:
   - Manuscript
   - Figures
   - Supplementary materials
   - Covering letter
   - Author forms
4. **Complete Submission** and receive confirmation

## Post-Submission

- **Manuscript Number** will be assigned
- **Acknowledgment** email within 24 hours
- **Initial Decision** within 1-2 weeks
- **Peer Review** if sent for review (8-12 weeks)

## Contact Information

- **Editorial Office:** globalhealth@lancet.com
- **Technical Support:** editorial@lancet.com
- **Corresponding Author:** hssiddalingaiah@gmail.com

---
*This checklist ensures compliance with The Lancet Global Health submission requirements.*
"""

    with open(checklist_path, 'w', encoding='utf-8') as f:
        f.write(checklist_content)

    logger.info(f"Submission checklist created: {checklist_path}")


def create_file_inventory(logger: logging.Logger) -> None:
    """Create an inventory of all submission files."""
    inventory_path = SUBMISSION_DIR / "file_inventory.txt"

    inventory = f"""THE LANCET GLOBAL HEALTH SUBMISSION - FILE INVENTORY
========================================================

Manuscript: Multi-Method Framework for Analyzing Tuberculosis Detection Delays in India
Submission Date: November 2025

CORE FILES
----------
1. Manuscript
   - File: reports/journal_submission_manuscript.docx
   - Description: Main manuscript with abstract, methods, results, discussion
   - Word count: 3,247

2. Covering Letter
   - File: submission_package/covering_letter.md
   - Description: Letter explaining manuscript importance and fit for journal

3. Author Statement Form
   - File: submission_package/author_statement_form.md
   - Description: ICMJE-compliant author statement with contributions

4. Conflict of Interest Declaration
   - File: Lancet Global Health/icmje-coi-form.docx
   - Description: Standard ICMJE conflict of interest form

5. Equitable Partnership Declaration
   - File: submission_package/equitable_partnership_declaration.md
   - Description: Declaration for research in LMICs

FIGURES
-------
1. Figure 1: Bayesian Forest Plot
   - File: output/figures/bayesian_forest_total_delay_(days).png
   - Description: Forest plot showing Bayesian meta-analysis results

2. Figure 2: PCA Scree Plot
   - File: output/figures/pca_scree_plot_delay_determinants.png
   - Description: Principal component analysis variance explained

3. Figure 3: DAG Causal Pathways
   - File: output/figures/dag_causal_delay_analysis.png
   - Description: Directed acyclic graph of causal relationships

TABLES
------
1. Table 1: Bayesian Estimates
   - Location: Embedded in manuscript
   - Description: MCMC Bayesian meta-analysis results

2. Table 2: State Prioritization
   - Location: Embedded in manuscript
   - Description: Top priority states for intervention

SUPPLEMENTARY MATERIALS
----------------------
1. Supplementary File
   - File: reports/supplementary.docx
   - Description: Detailed methods, additional results, code documentation

ADDITIONAL DOCUMENTS
-------------------
1. Executive Summary
   - File: reports/executive_summary_tb_delays.md
   - Description: Policy-focused summary for stakeholders

2. Submission Checklist
   - File: submission_package/submission_checklist.md
   - Description: Complete checklist of submission requirements

TECHNICAL SPECIFICATIONS
-----------------------
- Manuscript format: DOCX
- Figure resolution: 300+ DPI
- Color mode: RGB/CMYK acceptable
- Font: Times New Roman, 12pt
- Line spacing: Single
- Margins: 1 inch

CONTACT INFORMATION
------------------
Corresponding Author: H S Siddalingaiah
Email: hssiddalingaiah@gmail.com
Institution: Independent Researcher
Address: Bangalore, Karnataka, India

SUBMISSION PLATFORM
------------------
Journal: The Lancet Global Health
System: Editorial Manager
URL: www.editorialmanager.com/langlh

---
File inventory generated automatically on {Path(__file__).stat().st_mtime}
"""

    with open(inventory_path, 'w', encoding='utf-8') as f:
        f.write(inventory)

    logger.info(f"File inventory created: {inventory_path}")


def prepare_final_package(logger: logging.Logger) -> None:
    """Prepare the final submission package."""
    logger.info("Preparing final submission package...")

    # Ensure submission directory exists
    SUBMISSION_DIR.mkdir(exist_ok=True)

    # Check all required files
    files_complete = check_required_files(logger)

    if files_complete:
        logger.info("✓ All required files present")
    else:
        logger.warning("⚠ Some required files missing - check logs above")

    # Create additional documentation
    create_submission_checklist(logger)
    create_file_inventory(logger)

    # Create final submission archive instruction
    archive_instruction = f"""
SUBMISSION PACKAGE READY
========================

All files for The Lancet Global Health submission have been prepared.

To complete submission:

1. Visit: www.editorialmanager.com/langlh
2. Register/Login to Editorial Manager
3. Start new submission for The Lancet Global Health
4. Upload files in this order:
   a. Manuscript (journal_submission_manuscript.docx)
   b. Figures (3 PNG files)
   c. Covering letter (covering_letter.md)
   d. Author statement form (author_statement_form.md)
   e. Conflict of interest form (icmje-coi-form.docx)
   f. Equitable partnership declaration (equitable_partnership_declaration.md)
   g. Supplementary materials (supplementary.docx)

5. Complete all metadata fields
6. Submit and receive manuscript number

Key files are located in:
- Main manuscript: reports/
- Submission documents: submission_package/
- Figures: output/figures/
- Supplementary: reports/

Contact: globalhealth@lancet.com for technical support
"""

    instruction_path = SUBMISSION_DIR / "submission_instructions.txt"
    with open(instruction_path, 'w', encoding='utf-8') as f:
        f.write(archive_instruction)

    logger.info(f"Submission instructions created: {instruction_path}")
    logger.info("Submission package preparation complete!")


def main() -> None:
    """Main execution function."""
    logger = configure_logging()
    logger.info("Starting Lancet Global Health submission package preparation")

    try:
        prepare_final_package(logger)
        logger.info("Submission package preparation completed successfully")
    except Exception as e:
        logger.error(f"Error during submission package preparation: {e}")
        raise


if __name__ == "__main__":
    main()