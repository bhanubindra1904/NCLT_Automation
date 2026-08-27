import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
OUTPUT_ROOT = PROJECT_DIR / "outputs"
PIPELINE_SCRIPTS = [
    "scraper.py",
    "audit.py",
    "cin_enricher.py",
]

def run_pipeline():
    run_timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_dir = OUTPUT_ROOT / run_timestamp
    output_dir.mkdir(parents=True, exist_ok=False)

    child_environment = os.environ.copy()
    child_environment["NCLT_RUN_TIMESTAMP"] = run_timestamp
    child_environment["NCLT_OUTPUT_DIR"] = str(output_dir)

    print("[INIT] INITIALIZING END-TO-END NCLT PIPELINE...\n")
    print(f"[OUTPUT] Run folder: {output_dir}\n")

    for script in PIPELINE_SCRIPTS:
        script_path = PROJECT_DIR / script
        print("==================================================")
        print(f"[RUN] EXECUTING: {script}")
        print("==================================================")
        
        try:
            subprocess.run(
                [sys.executable, str(script_path)],
                check=True,
                cwd=str(PROJECT_DIR),
                env=child_environment,
            )
            print(f"\n[OK] {script} completed successfully.\n")
            
        except subprocess.CalledProcessError as e:
            print(f"\n[FAIL] FATAL ERROR: {script} crashed or failed.")
            print("[STOP] Halting the pipeline to prevent data corruption.")
            break
        except FileNotFoundError:
            print(f"\n[FAIL] ERROR: Could not find '{script}' in this folder.")
            print("[STOP] Halting the pipeline.")
            break

    print(f"[OUTPUT] Files saved in: {output_dir}")
    print("[DONE] PIPELINE EXECUTION FINISHED!")

if __name__ == "__main__":
    run_pipeline()
