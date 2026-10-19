
from pathlib import Path

import pandas as pd

from app.ingestion.scanner import discover_input_files
from app.ingestion.profiler import load_excel_file
from app.validation.business_rules import (
    REQUIRED_COLUMNS,
    OPTIONAL_COLUMNS,
    ValidationResult,
)


def validate_schema(df: pd.DataFrame) -> list[ValidationResult]:
    """Check required and optional SmartFlow fields."""

    results = []
    actual_columns = set(df.columns)

    missing_required = REQUIRED_COLUMNS - actual_columns
    missing_optional = OPTIONAL_COLUMNS - actual_columns

    if missing_required:
        results.append(
            ValidationResult(
                rule="Required columns",
                severity="CRITICAL",
                count=len(missing_required),
                description=(
                    "Required columns are missing: "
                    + ", ".join(sorted(missing_required))
                ),
            )
        )
    else:
        results.append(
            ValidationResult(
                rule="Required columns",
                severity="PASS",
                count=0,
                description="All required columns are present.",
            )
        )

    if missing_optional:
        results.append(
            ValidationResult(
                rule="Optional columns",
                severity="WARNING",
                count=len(missing_optional),
                description=(
                    "Optional columns not supplied: "
                    + ", ".join(sorted(missing_optional))
                    + ". Related analyses may be unavailable."
                ),
            )
        )
    else:
        results.append(
            ValidationResult(
                rule="Optional columns",
                severity="PASS",
                count=0,
                description="All configured optional columns are present.",
            )
        )

    return results


def _field_result(
    df: pd.DataFrame,
    column: str,
    rule: str,
    check,
    warning_description: str,
    pass_description: str,
    severity: str = "WARNING",
) -> ValidationResult:
    """Safely validate a field, including absent or null values."""

    if column not in df.columns:
        return ValidationResult(
            rule=rule,
            severity="WARNING",
            count=1,
            description=f"Cannot evaluate: column '{column}' is absent.",
        )

    count = int(check(df[column]).sum())

    return ValidationResult(
        rule=rule,
        severity=severity if count else "PASS",
        count=count,
        description=(
            warning_description.format(count=count)
            if count
            else pass_description
        ),
    )


def validate_duplicates(df: pd.DataFrame) -> list[ValidationResult]:
    count = int(df.duplicated().sum())

    return [
        ValidationResult(
            rule="Duplicate transactions",
            severity="WARNING" if count else "PASS",
            count=count,
            description=(
                f"{count:,} completely duplicated rows detected."
                if count
                else "No completely duplicated rows detected."
            ),
        )
    ]


def validate_missing_customer_ids(
    df: pd.DataFrame,
) -> list[ValidationResult]:
    return [
        _field_result(
            df,
            "CustomerID",
            "Missing CustomerID",
            lambda s: s.isna(),
            "{count:,} transactions do not have a CustomerID.",
            "All transactions have a CustomerID.",
        )
    ]


def validate_missing_descriptions(
    df: pd.DataFrame,
) -> list[ValidationResult]:
    return [
        _field_result(
            df,
            "Description",
            "Missing Description",
            lambda s: s.isna() | s.astype(str).str.strip().eq(""),
            "{count:,} transactions do not have a product description.",
            "All transactions have product descriptions.",
        )
    ]


def validate_negative_quantities(
    df: pd.DataFrame,
) -> list[ValidationResult]:
    return [
        _field_result(
            df,
            "Quantity",
            "Negative quantities",
            lambda s: pd.to_numeric(s, errors="coerce") < 0,
            "{count:,} transactions contain negative quantities.",
            "No negative quantities detected.",
            severity="INFO",
        )
    ]


def validate_prices(df: pd.DataFrame) -> list[ValidationResult]:
    return [
        _field_result(
            df,
            "UnitPrice",
            "Invalid UnitPrice",
            lambda s: pd.to_numeric(s, errors="coerce").isna()
            | (pd.to_numeric(s, errors="coerce") <= 0),
            "{count:,} transactions contain invalid, zero, or negative prices.",
            "All UnitPrice values are positive and numeric.",
        )
    ]


def validate_cancellations(
    df: pd.DataFrame,
) -> list[ValidationResult]:
    if "InvoiceNo" not in df.columns:
        return [
            ValidationResult(
                rule="Cancellation transactions",
                severity="WARNING",
                count=1,
                description="Cannot evaluate: InvoiceNo is absent.",
            )
        ]

    # This convention is specific to the original retail dataset.
    # Do not assume other clients use C-prefixed cancellation IDs.
    if df["InvoiceNo"].astype(str).str.upper().str.startswith("ORD-").any():
        return [
            ValidationResult(
                rule="Cancellation transactions",
                severity="INFO",
                count=0,
                description=(
                    "Cancellation status was not inferred: "
                    "this order-ID format has no configured cancellation rule."
                ),
            )
        ]

    count = int(
        df["InvoiceNo"].astype(str).str.upper().str.startswith("C").sum()
    )

    return [
        ValidationResult(
            rule="Cancellation transactions",
            severity="INFO" if count else "PASS",
            count=count,
            description=(
                f"{count:,} cancellation transaction rows detected."
                if count
                else "No cancellation invoices detected."
            ),
        )
    ]


def run_validation(df: pd.DataFrame) -> list[ValidationResult]:
    results = validate_schema(df)

    # Do not run field-specific rules when required fields are absent.
    if REQUIRED_COLUMNS - set(df.columns):
        return results + validate_duplicates(df)

    results.extend(validate_duplicates(df))
    results.extend(validate_missing_customer_ids(df))
    results.extend(validate_missing_descriptions(df))
    results.extend(validate_negative_quantities(df))
    results.extend(validate_prices(df))
    results.extend(validate_cancellations(df))

    return results


def print_validation_report(
    results: list[ValidationResult],
) -> None:
    print()
    print("=" * 70)
    print("SMARTFLOW BUSINESS VALIDATION")
    print("=" * 70)

    for result in results:
        icon = {
            "PASS": "[OK]",
            "INFO": "[INFO]",
            "WARNING": "[WARN]",
            "CRITICAL": "[FAIL]",
        }.get(result.severity, "[INFO]")

        print()
        print(f"{icon} {result.rule} [{result.severity}]")
        print(f"   Count : {result.count:,}")
        print(f"   {result.description}")

    print()
    print("=" * 70)
    print("VALIDATION COMPLETE")
    print("=" * 70)


def main():
    files = discover_input_files()

    if not files:
        raise FileNotFoundError("No supported input files found.")

    dataframe = load_excel_file(files[0])
    results = run_validation(dataframe)
    print_validation_report(results)

    if any(result.severity == "CRITICAL" for result in results):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
