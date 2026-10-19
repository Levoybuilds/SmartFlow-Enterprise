import os
import runpy
import time
from pathlib import Path

PROJECT_ROOT = Path(
    os.environ.get("SMARTFLOW_PROJECT_ROOT", Path(__file__).resolve().parent)
).expanduser().resolve()

PIPELINE = [
    ("File Discovery", "app.ingestion.scanner"),
    ("Data Profiling", "app.ingestion.profiler"),
    ("Business Validation", "app.validation.validator"),
    ("Data Quality Report", "app.reporting.quality_report"),
    ("Transformation Engine", "app.transformation.test_transformation"),
    ("Data Mart Generation", "app.analytics.data_mart"),
    ("KPI Engine", "app.analytics.test_kpi_engine"),
    ("Executive Report", "app.reporting.executive_report"),
]


def run_step(step_number, total_steps, name, module):
    print("\\n" + "-" * 72)
    print(f"[{step_number}/{total_steps}] {name}")
    print("-" * 72)
    started = time.time()
    try:
        # Run stages in-process. In a frozen app, sys.executable points back
        # to SmartFlow.exe; spawning it for each stage causes recursive starts.
        runpy.run_module(module, run_name="__main__")
    except SystemExit as exc:
        if exc.code not in (None, 0):
            print(f"\\n✗ FAILED: {name} ({module}); exit={exc.code}")
            return False
    except Exception as exc:
        print(f"\\n✗ FAILED: {name} ({module}): {type(exc).__name__}: {exc}")
        raise
    print(f"\\n✓ COMPLETED: {name}")
    print(f"  Time: {time.time() - started:.2f}s")
    return True


def main(input_file=None):
    if input_file:
        selected = Path(input_file).expanduser().resolve()
        if not selected.is_file():
            raise FileNotFoundError(f"Selected input file does not exist: {selected}")
        if selected.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
            raise ValueError("Choose an .xlsx, .xls, or .csv file.")
        if selected.name.startswith("~$"):
            raise ValueError("This looks like a temporary Excel lock file.")
        os.environ["SMARTFLOW_INPUT_FILE"] = str(selected)

    print("\\n" + "=" * 72)
    print("                 SMARTFLOW ENTERPRISE")
    print("              AUTOMATED ANALYTICS ENGINE")
    print("=" * 72)
    print("Starting SmartFlow Enterprise pipeline...")
    print(f"Project: {PROJECT_ROOT}")
    if input_file:
        print(f"Input file: {Path(input_file).name}")

    started = time.time()
    old_cwd = Path.cwd()
    try:
        os.chdir(PROJECT_ROOT)
        for index, (name, module) in enumerate(PIPELINE, start=1):
            if not run_step(index, len(PIPELINE), name, module):
                print("\\nSMARTFLOW PIPELINE FAILED")
                return 1
    finally:
        os.chdir(old_cwd)

    print("\\n" + "=" * 72)
    print("              SMARTFLOW PIPELINE COMPLETE")
    print("=" * 72)
    print(f"Total execution time: {time.time() - started:.2f} seconds")
    print("\\nGenerated outputs:")
    output_dir = PROJECT_ROOT / "output"
    if output_dir.exists():
        for file in sorted(output_dir.rglob("*")):
            if file.is_file():
                size_mb = file.stat().st_size / (1024 * 1024)
                print(f"  ✓ {file.relative_to(PROJECT_ROOT)} ({size_mb:.2f} MB)")
    print("\\nClient-ready analytics package generated.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
