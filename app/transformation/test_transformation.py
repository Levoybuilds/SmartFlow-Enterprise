from app.ingestion.scanner import discover_input_files
from app.ingestion.profiler import load_excel_file

from app.transformation.sales_transformer import (
    run_transformation,
    print_business_metrics,
)


def main():

    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No input file found."
        )

    input_file = files[0]

    print()
    print("=" * 60)
    print("SMARTFLOW TRANSFORMATION ENGINE")
    print("=" * 60)

    df = load_excel_file(input_file)

    print("Preparing transactions...")

    result = run_transformation(df)

    print_business_metrics(
        result["metrics"]
    )

    transactions = result["transactions"]

    print()
    print("TRANSACTION CLASSIFICATION")
    print("-" * 60)

    print(
        f"Positive valid sales     : "
        f"{transactions['IsValidSale'].sum():,}"
    )

    print(
        f"Cancellation rows        : "
        f"{transactions['IsCancellation'].sum():,}"
    )

    print(
        f"Return rows              : "
        f"{transactions['IsReturn'].sum():,}"
    )
    print(
        f"Zero-value adjustments   : "
        f"{transactions['IsZeroValueAdjustment'].sum():,}"
    )

    print(
        f"Negative quantity rows   : "
        f"{(transactions['Quantity'] < 0).sum():,}"
    )

    print(
        f"Zero/negative price rows : "
        f"{(transactions['UnitPrice'] <= 0).sum():,}"
    )

    print()
    print("ANALYTICAL DATASETS")
    print("-" * 60)

    print(
        f"Transactions     : "
        f"{len(result['transactions']):,}"
    )

    print(
        f"Customers         : "
        f"{len(result['customer_summary']):,}"
    )

    print(
        f"Products          : "
        f"{len(result['product_summary']):,}"
    )

    print(
        f"Countries         : "
        f"{len(result['country_summary']):,}"
    )

    print(
        f"Months            : "
        f"{len(result['monthly_sales']):,}"
    )

    print()
    print("Transformation complete.")


if __name__ == "__main__":
    main()
