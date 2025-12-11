# Data Guidance (Not Tracked)

Raw and intermediate data are **not** committed to this repository. To rerun analyses, place the following inputs under `data/` (keep them local and out of version control):

- **Study extracts**: CSV/JSON of delay metrics for the five Indian studies (PMIDs 37198982, 37459028, 38948543, 39775521, 41045529) with patient/diagnostic/total delays and sample sizes.
- **Programme indicators**: notification, incidence, GeneXpert coverage, bacteriological confirmation, and private-first-contact proportions (e.g., from India TB Report).
- **Structural determinants**: poverty, literacy, household crowding (e.g., NFHS-5, Census 2011).
- **Any preprocessing outputs** you generate (cleaned merges, posterior draws) should remain local.

Recommended layout (example):
```
data/
  sources/
    india_tb_report_2024.xlsx
    nfhs5_india.xlsx
    census2011_summary.xlsx
  studies/
    delay_extracts.csv
  intermediate/
    cleaned_merged.csv
    posterior_samples.parquet
```

Do not commit raw data or intermediate artifacts; only commit regenerated public outputs (manuscripts, figures, metrics JSON).
