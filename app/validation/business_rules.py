
from dataclasses import dataclass


@dataclass(frozen=True)
class ValidationResult:
    rule: str
    severity: str
    count: int
    description: str


# Minimum fields needed for the current sales analytics pipeline.
REQUIRED_COLUMNS = {
    "InvoiceNo",
    "Quantity",
    "InvoiceDate",
    "UnitPrice",
}

# Useful fields that are not mandatory for basic sales calculations.
OPTIONAL_COLUMNS = {
    "StockCode",
    "Description",
    "CustomerID",
    "Country",
}
