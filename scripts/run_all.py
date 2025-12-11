"""Master runner for the TB delay automation pipeline."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
from typing import List

PROJECT_ROOT = Path(__file__).resolve().parent
SCRIPTS = [
    "scripts/00_lit_search.py",
    "scripts/01_ingest_sources.py",
    "scripts/02_clean_merge.py",
    "scripts/03_extract_delay_from_lit.py",
    "scripts/04_meta_analysis_delays.py",
    "scripts/05_state_proxy_model.py",
    "scripts/06_visualizations.py",
    "scripts/07_generate_manuscript.py",
]


def run_script(script_path: Path, extra_args: List[str], stop_on_error: bool) -> None:
    cmd = [sys.executable, str(script_path), *extra_args]
    print(f"\n[run_all] Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    if result.returncode != 0 and stop_on_error:
        raise RuntimeError(f"Script {script_path} failed with exit code {result.returncode}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the full TB delay pipeline")
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="Stop the pipeline if any script exits with a non-zero code.",
    )
    parser.add_argument(
        "--lit-dry-run",
        action="store_true",
        help="Pass --dry-run to the literature search step.",
    )
    args = parser.parse_args()

    for script in SCRIPTS:
        script_path = PROJECT_ROOT / script
        extra = ["--dry-run"] if ("00_lit_search.py" in script and args.lit_dry_run) else []
        try:
            run_script(script_path, extra, args.stop_on_error)
        except RuntimeError as exc:
            print(f"[run_all] {exc}")
            break


if __name__ == "__main__":
    main()
