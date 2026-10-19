
from pathlib import Path

import pandas as pd

from app.ingestion.schema_mapper import (
    map_columns,
    map_columns_with_review,
)


def load_excel_file(
    file_path: Path,
    review_mapping: bool = False,
) -> pd.DataFrame:
    """Load an Excel or CSV file and map its columns."""

    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {file_path}"
        )

    print(f"\nLoading: {file_path.name}")
    print("This may take a little while for large workbooks...")

    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        dataframe = pd.read_csv(file_path)
    elif suffix in {".xlsx", ".xls"}:
        dataframe = pd.read_excel(file_path)
    else:
        raise ValueError(
            f"Unsupported file type: {suffix}"
        )

    if dataframe.empty:
        raise ValueError(
            "The input file contains no records."
        )

    if review_mapping:
        mapped, report = map_columns_with_review(dataframe)
    else:
        mapped, report = map_columns(dataframe)

    print("\nSMARTFLOW COLUMN MAPPING")
    print("-" * 60)

    for canonical, source in report["matched"].items():
        print(f"{source} -> {canonical}")

    if report["missing_required_fields"]:
        print(
            "Missing required fields:",
            ", ".join(report["missing_required_fields"]),
        )

    if report["missing_optional_fields"]:
        print(
            "Fields not supplied:",
            ", ".join(report["missing_optional_fields"]),
        )

    if report["unmapped_source_columns"]:
        print(
            "Unmapped source columns:",
            ", ".join(
                str(column)
                for column in report["unmapped_source_columns"]
            ),
        )

    return mapped



def profile_dataframe(df: pd.DataFrame) -> dict:
    """Generate a high-level data profile."""

    return {
        "rows": len(df),
        "columns": len(df.columns),
        "column_names": list(df.columns),
        "duplicate_rows": int(df.duplicated().sum()),
        "missing_values": {
            column: int(df[column].isna().sum())
            for column in df.columns
        },
        "data_types": {
            column: str(df[column].dtype)
            for column in df.columns
        },
    }


def print_profile(profile: dict) -> None:
    """Print the data profile."""

    print()
    print("=" * 60)
    print("SMARTFLOW DATA PROFILE")
    print("=" * 60)

    print(f"Rows              : {profile['rows']:,}")
    print(f"Columns           : {profile['columns']}")
    print(f"Duplicate rows    : {profile['duplicate_rows']:,}")

    print()
    print("COLUMNS")
    print("-" * 60)

    for column in profile["column_names"]:
        print(
            f"{column:<20}"
            f"Type: {profile['data_types'][column]:<12}"
            f"Missing: {profile['missing_values'][column]:,}"
        )

    print("=" * 60)


def profile_file(file_path: Path) -> dict:
    dataframe = load_excel_file(file_path)
    profile = profile_dataframe(dataframe)
    print_profile(profile)
    return profile


if __name__ == "__main__":
    from app.ingestion.scanner import discover_input_files

    files = discover_input_files()

    if not files:
        raise FileNotFoundError("No supported input files found.")

    profile_file(files[0])
