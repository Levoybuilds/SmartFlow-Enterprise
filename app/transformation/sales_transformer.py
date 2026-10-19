from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class BusinessMetrics:
    """Canonical business metrics used throughout SmartFlow."""

    gross_sales: float
    return_value: float
    net_sales: float
    order_count: int
    cancelled_order_count: int
    customer_count: int
    product_count: int


def prepare_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare the raw retail dataset for downstream analytics.

    Important:
    The original DataFrame is never modified.
    """

    data = df.copy()

    # Standardize column names.
    data.columns = [
        str(column).strip()
        for column in data.columns
    ]

    # Normalize text fields.
    data["InvoiceNo"] = (
        data["InvoiceNo"]
        .astype("string")
        .str.strip()
    )

    data["StockCode"] = (
        data["StockCode"]
        .astype("string")
        .str.strip()
    )

    data["Description"] = (
        data["Description"]
        .astype("string")
        .str.strip()
    )

    data["Country"] = (
        data["Country"]
        .astype("string")
        .str.strip()
    )

    # Ensure correct date type.
    data["InvoiceDate"] = pd.to_datetime(
        data["InvoiceDate"],
        errors="coerce",
    )

    # Numeric conversion.
    data["Quantity"] = pd.to_numeric(
        data["Quantity"],
        errors="coerce",
    )

    data["UnitPrice"] = pd.to_numeric(
        data["UnitPrice"],
        errors="coerce",
    )

    # Customer IDs should be treated as identifiers,
    # not numerical measures.
    data["CustomerID"] = (
        data["CustomerID"]
        .astype("Int64")
    )

    # Business transaction value.
    data["LineValue"] = (
        data["Quantity"] * data["UnitPrice"]
    )

    # ---------------------------------------------------------
    # BUSINESS TRANSACTION CLASSIFICATION
    # ---------------------------------------------------------

    data["IsCancellation"] = (
        data["InvoiceNo"]
        .str.upper()
        .str.startswith("C")
        .fillna(False)
    )

    data["IsReturn"] = (
    (data["Quantity"] < 0)
    & (~data["IsCancellation"])
    & (data["UnitPrice"] > 0)
    )

    data["IsZeroValueAdjustment"] = (
        (data["Quantity"] < 0)
        & (~data["IsCancellation"])
        & (data["UnitPrice"] <= 0)
    )

    data["IsValidSale"] = (
        (data["Quantity"] > 0)
        & (data["UnitPrice"] > 0)
        & (~data["IsCancellation"])
    )

    # Explicit transaction type makes the business logic
    # transparent and reusable across the entire system.

    data["TransactionType"] = "OTHER"

    data.loc[
        data["IsValidSale"],
        "TransactionType"
    ] = "SALE"

    data.loc[
        data["IsReturn"],
        "TransactionType"
    ] = "RETURN"

    data.loc[
        data["IsCancellation"],
        "TransactionType"
    ] = "CANCELLATION"
    data.loc[
        data["IsZeroValueAdjustment"],
        "TransactionType"
    ] = "ZERO_VALUE_ADJUSTMENT"



    # Transaction date components.
    data["OrderDate"] = data["InvoiceDate"].dt.date

    data["Year"] = data["InvoiceDate"].dt.year

    data["Month"] = data["InvoiceDate"].dt.month

    data["YearMonth"] = (
        data["InvoiceDate"]
        .dt.to_period("M")
        .astype("string")
    )

    return data


def calculate_business_metrics(
    transactions: pd.DataFrame,
) -> BusinessMetrics:
    """Calculate the single canonical set of business metrics."""

    sales = transactions[
        transactions["TransactionType"] == "SALE"
    ]

    returns = transactions[
        transactions["TransactionType"] == "RETURN"
    ]

    cancellations = transactions[
        transactions["TransactionType"] == "CANCELLATION"
    ]

    gross_sales = float(
        sales["LineValue"].sum()
    )

    return_value = float(
        returns["LineValue"].sum()
    )

    net_sales = (
        gross_sales
        + return_value
    )

    order_count = int(
        sales["InvoiceNo"].nunique()
    )

    cancelled_order_count = int(
        cancellations["InvoiceNo"].nunique()
    )

    customer_count = int(
        sales["CustomerID"].nunique()
    )

    product_count = int(
        sales["StockCode"].nunique()
    )

    return BusinessMetrics(
        gross_sales=gross_sales,
        return_value=return_value,
        net_sales=net_sales,
        order_count=order_count,
        cancelled_order_count=cancelled_order_count,
        customer_count=customer_count,
        product_count=product_count,
    )


def create_customer_summary(
    transactions: pd.DataFrame,
) -> pd.DataFrame:
    """Create customer-level sales analytics."""

    sales = transactions[
        transactions["IsValidSale"]
        & transactions["CustomerID"].notna()
    ].copy()

    summary = (
        sales.groupby("CustomerID")
        .agg(
            Revenue=("LineValue", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Units=("Quantity", "sum"),
            Products=("StockCode", "nunique"),
        )
        .reset_index()
    )

    summary["Average_Order_Value"] = (
        summary["Revenue"] / summary["Orders"]
    )

    summary = summary.sort_values(
        "Revenue",
        ascending=False,
    )

    return summary


def create_product_summary(
    transactions: pd.DataFrame,
) -> pd.DataFrame:
    """Create product-level sales analytics."""

    sales = transactions[
        transactions["IsValidSale"]
    ].copy()

    summary = (
        sales.groupby(
            [
                "StockCode",
                "Description",
            ],
            dropna=False,
        )
        .agg(
            Revenue=("LineValue", "sum"),
            Units=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
        )
        .reset_index()
    )

    summary = summary.sort_values(
        "Revenue",
        ascending=False,
    )

    return summary


def create_country_summary(
    transactions: pd.DataFrame,
) -> pd.DataFrame:
    """Create country-level sales analytics."""

    sales = transactions[
        transactions["IsValidSale"]
    ].copy()

    summary = (
        sales.groupby("Country")
        .agg(
            Revenue=("LineValue", "sum"),
            Units=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
        )
        .reset_index()
    )

    summary = summary.sort_values(
        "Revenue",
        ascending=False,
    )

    return summary


def create_monthly_sales(
    transactions: pd.DataFrame,
) -> pd.DataFrame:
    """Create monthly sales trend data."""

    sales = transactions[
        transactions["IsValidSale"]
    ].copy()

    summary = (
        sales.groupby("YearMonth")
        .agg(
            Revenue=("LineValue", "sum"),
            Units=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
        )
        .reset_index()
    )

    summary["YearMonth"] = summary[
        "YearMonth"
    ].astype(str)

    return summary.sort_values("YearMonth")


def print_business_metrics(
    metrics: BusinessMetrics,
) -> None:
    """Print canonical business metrics."""

    print()
    print("=" * 60)
    print("CANONICAL BUSINESS METRICS")
    print("=" * 60)

    print(
        f"Gross Sales              : "
        f"{metrics.gross_sales:,.2f}"
    )

    print(
        f"Return Value             : "
        f"{metrics.return_value:,.2f}"
    )

    print(
        f"Net Sales                : "
        f"{metrics.net_sales:,.2f}"
    )

    print(
        f"Valid Orders             : "
        f"{metrics.order_count:,}"
    )

    print(
        f"Cancelled Orders         : "
        f"{metrics.cancelled_order_count:,}"
    )

    print(
        f"Customers                : "
        f"{metrics.customer_count:,}"
    )

    print(
        f"Products                 : "
        f"{metrics.product_count:,}"
    )

    print("=" * 60)


def run_transformation(
    df: pd.DataFrame,
):
    """
    Execute the canonical transformation layer.
    """

    transactions = prepare_transactions(df)

    metrics = calculate_business_metrics(
        transactions
    )

    customer_summary = create_customer_summary(
        transactions
    )

    product_summary = create_product_summary(
        transactions
    )

    country_summary = create_country_summary(
        transactions
    )

    monthly_sales = create_monthly_sales(
        transactions
    )

    return {
        "transactions": transactions,
        "metrics": metrics,
        "customer_summary": customer_summary,
        "product_summary": product_summary,
        "country_summary": country_summary,
        "monthly_sales": monthly_sales,
    }