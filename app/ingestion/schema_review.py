
from app.ingestion.scanner import discover_input_files
from app.ingestion.profiler import load_excel_file


def main():
    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No input files found. Select an Excel or CSV file."
        )

    input_file = files[0]

    print("=" * 70)
    print("SMARTFLOW INTERACTIVE SCHEMA REVIEW")
    print("=" * 70)
    print(f"Input file: {input_file}")

    mapped_df = load_excel_file(
        input_file,
        review_mapping=True,
    )

    print()
    print("=" * 70)
    print("MAPPING REVIEW COMPLETED")
    print("=" * 70)
    print("Canonical columns:")
    print(", ".join(str(column) for column in mapped_df.columns))
    print()
    print(
        "This command previews the mapped dataset in memory. "
        "It does not overwrite your original file."
    )
    print(
        "Review completion does not yet guarantee that every "
        "downstream transformation supports this dataset."
    )


if __name__ == "__main__":
    main()
