from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class ExecutiveKPIs:
    gross_sales: float
    net_sales: float
    return_value: float
    orders: int
    customers: int
    products: int
    units_sold: int
    average_order_value: float
    cancellation_rate: float
    customer_revenue: float


def calculate_executive_kpis(
    transactions: pd.DataFrame,
    metrics,
) -> ExecutiveKPIs:
    """Calculate the canonical executive KPI set."""

    sales = transactions[
        transactions["TransactionType"] == "SALE"
    ]

    units_sold = int(
        sales["Quantity"].sum()
    )

    average_order_value = (
        metrics.gross_sales
        / metrics.order_count
        if metrics.order_count > 0
        else 0.0
    )

    cancellation_rate = (
        metrics.cancelled_order_count
        / (
            metrics.order_count
            + metrics.cancelled_order_count
        )
        * 100
        if (
            metrics.order_count
            + metrics.cancelled_order_count
        ) > 0
        else 0.0
    )

    customer_revenue = float(
        sales[
            sales["CustomerID"].notna()
        ]["LineValue"].sum()
    )

    return ExecutiveKPIs(
        gross_sales=metrics.gross_sales,
        net_sales=metrics.net_sales,
        return_value=metrics.return_value,
        orders=metrics.order_count,
        customers=metrics.customer_count,
        products=metrics.product_count,
        units_sold=units_sold,
        average_order_value=average_order_value,
        cancellation_rate=cancellation_rate,
        customer_revenue=customer_revenue,
    )


def print_kpis(
    kpis: ExecutiveKPIs,
) -> None:

    print()
    print("=" * 70)
    print("SMARTFLOW EXECUTIVE KPI ENGINE")
    print("=" * 70)

    print(
        f"Gross Sales          : "
        f"£{kpis.gross_sales:,.2f}"
    )

    print(
        f"Net Sales            : "
        f"£{kpis.net_sales:,.2f}"
    )

    print(
        f"Return Value         : "
        f"£{kpis.return_value:,.2f}"
    )

    print(
        f"Orders               : "
        f"{kpis.orders:,}"
    )

    print(
        f"Customers            : "
        f"{kpis.customers:,}"
    )

    print(
        f"Products             : "
        f"{kpis.products:,}"
    )

    print(
        f"Units Sold           : "
        f"{kpis.units_sold:,}"
    )

    print(
        f"Average Order Value  : "
        f"£{kpis.average_order_value:,.2f}"
    )

    print(
        f"Cancellation Rate    : "
        f"{kpis.cancellation_rate:.2f}%"
    )

    print(
        f"Customer Revenue     : "
        f"£{kpis.customer_revenue:,.2f}"
    )

    print("=" * 70)