from pathlib import Path

import pandas as pd

from app.ingestion.scanner import discover_input_files, get_project_root
from app.ingestion.profiler import load_excel_file
from app.transformation.sales_transformer import run_transformation


PROJECT_ROOT = get_project_root()

ANALYTICAL_DIR = (
    PROJECT_ROOT
    / "output"
    / "analytical"
)


def save_dataset(
    dataframe: pd.DataFrame,
    filename: str,
) -> Path:
    """Save an analytical dataset as CSV."""

    ANALYTICAL_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        ANALYTICAL_DIR
        / filename
    )

    dataframe.to_csv(
        output_path,
        index=False,
    )

    return output_path


def create_data_mart(
    transformations: dict,
) -> dict:
    """
    Persist SmartFlow analytical datasets.

    Returns a dictionary containing generated paths.
    """

    paths = {}

    paths["transactions"] = save_dataset(
        transformations["transactions"],
        "transaction_fact.csv",
    )

    paths["customers"] = save_dataset(
        transformations["customer_summary"],
        "customer_summary.csv",
    )

    paths["products"] = save_dataset(
        transformations["product_summary"],
        "product_summary.csv",
    )

    paths["countries"] = save_dataset(
        transformations["country_summary"],
        "country_summary.csv",
    )

    paths["monthly_sales"] = save_dataset(
        transformations["monthly_sales"],
        "monthly_sales.csv",
    )

    return paths


def print_data_mart_summary(
    transformations: dict,
    paths: dict,
) -> None:
    """Display data-mart generation summary."""

    print()
    print("=" * 70)
    print("SMARTFLOW ANALYTICAL DATA MART")
    print("=" * 70)

    for name, path in paths.items():

        print(
            f"{name:<20}"
            f" {path.name:<30}"
            f" {path.stat().st_size / 1024 / 1024:.2f} MB"
        )

    print()
    print("DATASET ROW COUNTS")
    print("-" * 70)

    print(
        f"Transactions         : "
        f"{len(transformations['transactions']):,}"
    )

    print(
        f"Customers            : "
        f"{len(transformations['customer_summary']):,}"
    )

    print(
        f"Products             : "
        f"{len(transformations['product_summary']):,}"
    )

    print(
        f"Countries            : "
        f"{len(transformations['country_summary']):,}"
    )

    print(
        f"Monthly records      : "
        f"{len(transformations['monthly_sales']):,}"
    )

    print("=" * 70)


def main():

    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No input files found."
        )

    input_file = files[0]

    print()
    print("=" * 70)
    print("SMARTFLOW DATA MART GENERATOR")
    print("=" * 70)

    dataframe = load_excel_file(
        input_file
    )

    print("Running transformation engine...")

    transformations = run_transformation(
        dataframe
    )

    print("Persisting analytical datasets...")

    paths = create_data_mart(
        transformations
    )

    print_data_mart_summary(
        transformations,
        paths,
    )

    print()
    print("Data mart generation complete.")


if __name__ == "__main__":
    main()