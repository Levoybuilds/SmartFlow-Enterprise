
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from app.ingestion.scanner import discover_input_files, get_project_root
from app.ingestion.profiler import load_excel_file
from app.validation.validator import run_validation


PROJECT_ROOT = get_project_root()
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "SmartFlow_Data_Quality_Report.xlsx"


# ============================================================
# 1. EXECUTIVE SUMMARY
# ============================================================

def create_summary_dataframe(df, validation_results):
    """Create executive summary metrics safely."""

    def missing_count(column):
        if column not in df.columns:
            return None
        return int(df[column].isna().sum())

    def count_if(column, condition):
        if column not in df.columns:
            return None

        values = condition(df[column])
        return int(values.sum())

    # Cancellation detection is specific to the original retail
    # dataset. Do not assume that every dataset uses C-prefixed IDs.
    cancellation_count = None

    if "InvoiceNo" in df.columns:
        invoice_ids = df["InvoiceNo"].astype(str).str.upper()

        is_sample_order_format = invoice_ids.str.startswith("ORD-").any()

        if not is_sample_order_format:
            cancellation_count = int(
                invoice_ids.str.startswith("C").sum()
            )

    metrics = [
        ("Total Records", len(df)),
        ("Total Columns", len(df.columns)),
        ("Duplicate Records", int(df.duplicated().sum())),
        ("Missing Customer IDs", missing_count("CustomerID")),
        ("Missing Descriptions", missing_count("Description")),
        (
            "Negative Quantity Records",
            count_if(
                "Quantity",
                lambda s: pd.to_numeric(s, errors="coerce") < 0,
            ),
        ),
        (
            "Invalid Price Records",
            count_if(
                "UnitPrice",
                lambda s: (
                    pd.to_numeric(s, errors="coerce").isna()
                    | (pd.to_numeric(s, errors="coerce") <= 0)
                ),
            ),
        ),
        ("Cancellation Records", cancellation_count),
    ]

    summary = pd.DataFrame(
        metrics,
        columns=["Metric", "Value"],
    )

    warning_count = sum(
        result.severity == "WARNING"
        for result in validation_results
    )

    critical_count = sum(
        result.severity == "CRITICAL"
        for result in validation_results
    )

    if critical_count:
        status = "CRITICAL"
    elif warning_count:
        status = "REVIEW REQUIRED"
    else:
        status = "PASS"

    summary.loc[len(summary)] = [
        "Overall Data Status",
        status,
    ]

    return summary


# ============================================================
# 2. VALIDATION RESULTS
# ============================================================

def create_validation_dataframe(validation_results):
    """Convert validation results into a report-ready dataframe."""

    rows = []

    for result in validation_results:
        rows.append({
            "Rule": result.rule,
            "Severity": result.severity,
            "Count": result.count,
            "Description": result.description,
        })

    return pd.DataFrame(
        rows,
        columns=[
            "Rule",
            "Severity",
            "Count",
            "Description",
        ],
    )


# ============================================================
# 3. COLUMN PROFILE
# ============================================================

def create_column_profile(df):
    """Profile every column that exists in the input dataset."""

    rows = []

    for column in df.columns:
        missing_values = int(df[column].isna().sum())

        missing_percentage = (
            round(df[column].isna().mean() * 100, 2)
            if len(df) > 0
            else 0.0
        )

        rows.append({
            "Column": column,
            "Data Type": str(df[column].dtype),
            "Total Records": len(df),
            "Missing Values": missing_values,
            "Missing Percentage": missing_percentage,
            "Unique Values": int(
                df[column].nunique(dropna=True)
            ),
        })

    return pd.DataFrame(
        rows,
        columns=[
            "Column",
            "Data Type",
            "Total Records",
            "Missing Values",
            "Missing Percentage",
            "Unique Values",
        ],
    )


# ============================================================
# 4. DATA QUALITY METRICS
# ============================================================

def create_quality_metrics(df):
    """Calculate quality percentages for available fields."""

    metrics = []

    def add_completeness(column, label):
        if column in df.columns:
            percentage = (
                round(df[column].notna().mean() * 100, 2)
                if len(df) > 0
                else 100.0
            )

            metrics.append((
                f"Completeness - {label}",
                percentage,
            ))

    add_completeness("CustomerID", "CustomerID")
    add_completeness("Description", "Description")

    if "UnitPrice" in df.columns:
        prices = pd.to_numeric(
            df["UnitPrice"],
            errors="coerce",
        )

        positive_price_rate = (
            round((prices > 0).mean() * 100, 2)
            if len(df) > 0
            else 0.0
        )

        metrics.append((
            "Positive Price Rate",
            positive_price_rate,
        ))

    if "Quantity" in df.columns:
        quantities = pd.to_numeric(
            df["Quantity"],
            errors="coerce",
        )

        positive_quantity_rate = (
            round((quantities > 0).mean() * 100, 2)
            if len(df) > 0
            else 0.0
        )

        metrics.append((
            "Positive Quantity Rate",
            positive_quantity_rate,
        ))

    duplicate_free_rate = (
        round((1 - df.duplicated().mean()) * 100, 2)
        if len(df) > 0
        else 100.0
    )

    metrics.append((
        "Duplicate-Free Rate",
        duplicate_free_rate,
    ))

    return pd.DataFrame(
        metrics,
        columns=["Quality Metric", "Percentage"],
    )


# ============================================================
# 5. EXCEPTION SAMPLE
# ============================================================

def create_exception_sample(df):
    """Return up to 1,000 rows containing detectable exceptions."""

    conditions = pd.Series(
        False,
        index=df.index,
        dtype=bool,
    )

    if "CustomerID" in df.columns:
        conditions |= df["CustomerID"].isna()

    if "Description" in df.columns:
        conditions |= df["Description"].isna()

    if "Quantity" in df.columns:
        quantities = pd.to_numeric(
            df["Quantity"],
            errors="coerce",
        )

        conditions |= (
            quantities.isna()
            | (quantities < 0)
        )

    if "UnitPrice" in df.columns:
        prices = pd.to_numeric(
            df["UnitPrice"],
            errors="coerce",
        )

        conditions |= (
            prices.isna()
            | (prices <= 0)
        )

    return df.loc[conditions].head(1000).copy()


# ============================================================
# 6. EXCEL FORMATTING
# ============================================================

def format_workbook(file_path):
    """Apply consistent formatting to every report worksheet."""

    workbook = load_workbook(file_path)

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    title_font = Font(
        bold=True,
        size=16,
        color="1F1F1F",
    )

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Format the header row.
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        worksheet.row_dimensions[1].height = 30

        # Set readable column widths.
        for column_cells in worksheet.columns:
            if not column_cells:
                continue

            column_letter = get_column_letter(
                column_cells[0].column
            )

            max_length = 0

            for cell in column_cells:
                value = (
                    str(cell.value)
                    if cell.value is not None
                    else ""
                )

                max_length = max(
                    max_length,
                    len(value),
                )

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 12),
                50,
            )

        # Make long text easier to read.
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

    # Highlight the summary title cell.
    if "Executive Summary" in workbook.sheetnames:
        summary_sheet = workbook["Executive Summary"]
        summary_sheet["A1"].font = title_font

    workbook.save(file_path)
    workbook.close()


# ============================================================
# 7. GENERATE EXCEL REPORT
# ============================================================

def generate_report(df, validation_results):
    """Generate the complete SmartFlow data quality workbook."""

    if df.empty:
        raise ValueError(
            "The input dataset contains no records. "
            "Cannot generate a meaningful quality report."
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    summary = create_summary_dataframe(
        df,
        validation_results,
    )

    validation = create_validation_dataframe(
        validation_results,
    )

    column_profile = create_column_profile(df)

    quality_metrics = create_quality_metrics(df)

    exceptions = create_exception_sample(df)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="Executive Summary",
            index=False,
        )

        validation.to_excel(
            writer,
            sheet_name="Validation Results",
            index=False,
        )

        column_profile.to_excel(
            writer,
            sheet_name="Column Profile",
            index=False,
        )

        quality_metrics.to_excel(
            writer,
            sheet_name="Data Quality",
            index=False,
        )

        exceptions.to_excel(
            writer,
            sheet_name="Exceptions",
            index=False,
        )

    format_workbook(OUTPUT_FILE)

    return OUTPUT_FILE


# ============================================================
# 8. MAIN EXECUTION
# ============================================================

def main():
    """Load the selected dataset, validate it, and create the report."""

    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No supported input files found. "
            "Select an Excel or CSV file first."
        )

    input_file = files[0]

    print()
    print("=" * 60)
    print("SMARTFLOW DATA QUALITY REPORT")
    print("=" * 60)
    print(f"Input: {input_file}")
    print()

    df = load_excel_file(input_file)

    if df.empty:
        raise ValueError(
            "The selected input file contains no records."
        )

    print()
    print("Running validation...")

    validation_results = run_validation(df)

    print()
    print("Generating Excel report...")

    output_file = generate_report(
        df,
        validation_results,
    )

    print()
    print("=" * 60)
    print("REPORT GENERATED SUCCESSFULLY")
    print("=" * 60)
    print()
    print(f"Output: {output_file}")
    print(f"Records processed: {len(df):,}")
    print(f"Columns processed: {len(df.columns)}")
    print()


if __name__ == "__main__":
    main()
