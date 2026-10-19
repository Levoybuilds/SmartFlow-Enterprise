
from __future__ import annotations

from difflib import SequenceMatcher
from typing import Any

import pandas as pd


# Canonical field -> known client column names.
# Keep canonical field names stable for downstream modules.
COLUMN_ALIASES = {
    "InvoiceNo": [
        "InvoiceNo", "Invoice No", "Order ID", "OrderID",
        "Order Ref", "Order Reference", "Order Number",
        "OrderNo", "Transaction ID",
    ],
    "StockCode": [
        "StockCode", "Stock Code", "SKU", "Product ID",
        "Product Code", "Item Code",
    ],
    "Description": [
        "Description", "Product", "Product Name",
        "Item", "Item Name", "Merchandise",
    ],
    "Quantity": [
        "Quantity", "Qty", "Units", "Units Sold",
        "Quantity Sold", "Units Purchased", "Merchandise Count",
    ],
    "InvoiceDate": [
        "InvoiceDate", "Invoice Date", "Order Date",
        "OrderDate", "Purchase Date", "Purchase Timestamp",
        "Transaction Date", "Date",
    ],
    "UnitPrice": [
        "UnitPrice", "Unit Price", "Price Per Unit",
        "Rate Per Item", "Price", "Unit Cost",
    ],
    "SalesAmount": [
        "SalesAmount", "Sales Amount", "Amount",
        "Total Amount", "Order Value", "Line Total",
        "Total Sales", "Revenue", "Net Amount",
    ],
    "CustomerID": [
        "CustomerID", "Customer ID", "Customer Number",
        "Buyer Number", "Client ID", "Buyer ID",
    ],
    "Country": [
        "Country", "Market", "Nation",
    ],
    "Region": [
        "Region", "Sales Region", "Territory",
    ],
    "SalesChannel": [
        "Sales Channel", "Channel", "Order Channel",
    ],
}

# Core fields for the current transaction model.
# SalesAmount is an alternative revenue input, not a UnitPrice alias.
REQUIRED_FIELDS = [
    "InvoiceNo",
    "InvoiceDate",
]

OPTIONAL_FIELDS = [
    "StockCode",
    "Description",
    "Quantity",
    "UnitPrice",
    "SalesAmount",
    "CustomerID",
    "Country",
    "Region",
    "SalesChannel",
]


def normalize_name(name: Any) -> str:
    """Normalize a column name for comparison."""
    return "".join(
        character.lower()
        for character in str(name).strip()
        if character.isalnum()
    )


def _build_alias_lookup():
    """Create a normalized alias lookup."""
    lookup = {}

    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in [canonical, *aliases]:
            normalized = normalize_name(alias)
            lookup.setdefault(normalized, set()).add(canonical)

    return lookup


ALIAS_LOOKUP = _build_alias_lookup()


def _profile_column(series: pd.Series) -> dict:
    """Collect lightweight evidence about a column's contents."""
    sample = series.dropna().head(20)

    date_ratio = 0.0
    numeric_ratio = 0.0

    if len(sample):
        try:
            parsed_dates = pd.to_datetime(
                sample,
                errors="coerce",
            )
            date_ratio = float(parsed_dates.notna().mean())
        except (ValueError, TypeError):
            pass

        try:
            parsed_numbers = pd.to_numeric(
                sample,
                errors="coerce",
            )
            numeric_ratio = float(parsed_numbers.notna().mean())
        except (ValueError, TypeError):
            pass

    return {
        "dtype": str(series.dtype),
        "sample_values": [
            str(value)[:80]
            for value in sample.head(5).tolist()
        ],
        "missing_count": int(series.isna().sum()),
        "unique_count": int(series.nunique(dropna=True)),
        "date_like_ratio": round(date_ratio, 2),
        "numeric_like_ratio": round(numeric_ratio, 2),
    }


def _suggest_for_canonical(
    canonical: str,
    columns: list,
) -> dict | None:
    """Suggest one source column for a canonical field."""

    aliases = [
        normalize_name(alias)
        for alias in [canonical, *COLUMN_ALIASES[canonical]]
    ]

    # First preference: exact normalized alias.
    exact_matches = [
        column
        for column in columns
        if normalize_name(column) in aliases
    ]

    if len(exact_matches) == 1:
        return {
            "canonical": canonical,
            "source": exact_matches[0],
            "confidence": 1.0,
            "method": "exact_alias",
        }

    # Multiple exact matches are ambiguous; don't guess.
    if len(exact_matches) > 1:
        return {
            "canonical": canonical,
            "source": None,
            "candidates": exact_matches,
            "confidence": 0.0,
            "method": "ambiguous_exact_alias",
        }

    # Fuzzy matching proposes candidates but does not auto-map them.
    scores = []

    for column in columns:
        source_name = normalize_name(column)

        score = max(
            SequenceMatcher(
                None,
                source_name,
                alias,
            ).ratio()
            for alias in aliases
        )

        scores.append((score, column))

    scores.sort(key=lambda item: item[0], reverse=True)

    if not scores or scores[0][0] < 0.78:
        return None

    best_score, best_column = scores[0]

    # If the top two scores are close, mark the match ambiguous.
    if (
        len(scores) > 1
        and best_score - scores[1][0] < 0.04
    ):
        return {
            "canonical": canonical,
            "source": None,
            "candidates": [
                scores[0][1],
                scores[1][1],
            ],
            "confidence": round(best_score, 2),
            "method": "ambiguous_fuzzy",
        }

    return {
        "canonical": canonical,
        "source": best_column,
        "confidence": round(best_score, 2),
        "method": "fuzzy_suggestion",
    }


def suggest_column_mappings(df: pd.DataFrame) -> list[dict]:
    """Suggest canonical mappings without modifying the DataFrame."""

    columns = list(df.columns)
    suggestions = []

    for canonical in COLUMN_ALIASES:
        suggestion = _suggest_for_canonical(
            canonical,
            columns,
        )

        if suggestion is not None:
            suggestions.append(suggestion)

    # A single source column must not be assigned to multiple targets.
    source_counts = {}

    for suggestion in suggestions:
        source = suggestion.get("source")
        if source is not None:
            source_counts[source] = source_counts.get(source, 0) + 1

    for suggestion in suggestions:
        source = suggestion.get("source")

        if source is not None and source_counts[source] > 1:
            suggestion["source"] = None
            suggestion["method"] = "ambiguous_multiple_targets"
            suggestion["candidates"] = [source]
            suggestion["confidence"] = 0.0

    return suggestions


def review_column_mappings(
    df: pd.DataFrame,
) -> tuple[dict[str, str], dict]:
    """Interactively review suggested source-to-canonical mappings."""

    suggestions = suggest_column_mappings(df)
    accepted = {}
    claimed_sources = set()

    print()
    print("=" * 70)
    print("SMARTFLOW SCHEMA MAPPING REVIEW")
    print("=" * 70)

    print("\nDetected source columns:")

    for index, column in enumerate(df.columns, start=1):
        profile = _profile_column(df[column])

        print(
            f"{index}. {column!r} | "
            f"type={profile['dtype']} | "
            f"sample={profile['sample_values']}"
        )

    by_canonical = {
        item["canonical"]: item
        for item in suggestions
    }

    for canonical in COLUMN_ALIASES:
        suggestion = by_canonical.get(canonical)

        if not suggestion:
            continue

        source = suggestion.get("source")
        method = suggestion.get("method")
        confidence = suggestion.get("confidence", 0.0)

        print()
        print(f"Target field: {canonical}")

        if source is None:
            candidates = suggestion.get("candidates", [])
            print("  Mapping is ambiguous or needs manual selection.")

            if candidates:
                for index, candidate in enumerate(
                    candidates,
                    start=1,
                ):
                    print(f"  {index}. {candidate!r}")

                answer = input(
                    "  Enter candidate number, or press Enter to skip: "
                ).strip()

                if answer.isdigit():
                    candidate_index = int(answer) - 1

                    if 0 <= candidate_index < len(candidates):
                        chosen = candidates[candidate_index]

                        if chosen not in claimed_sources:
                            accepted[canonical] = chosen
                            claimed_sources.add(chosen)

            continue

        if source in claimed_sources:
            continue

        if method == "exact_alias":
            print(
                f"  Suggested: {source!r} -> {canonical} "
                "(exact alias)"
            )
            default = "Y"
        else:
            print(
                f"  Suggested: {source!r} -> {canonical} "
                f"(fuzzy score {confidence:.2f})"
            )
            default = "N"

        # Show evidence for the proposed mapping.
        profile = _profile_column(df[source])

        print(f"  Sample values: {profile['sample_values']}")
        answer = input(
            f"  Accept mapping? [y/N] (default {default}): "
        ).strip().lower()

        if not answer:
            answer = default.lower()

        if answer in {"y", "yes"}:
            accepted[canonical] = source
            claimed_sources.add(source)

    # Allow manual mapping for targets that have not been assigned.
    print()
    print("Manual mapping for remaining fields")
    print("Enter a source column name, or press Enter to skip.")

    for canonical in COLUMN_ALIASES:
        if canonical in accepted:
            continue

        available_columns = [
            column
            for column in df.columns
            if column not in claimed_sources
        ]

        if not available_columns:
            break

        print(f"\nTarget: {canonical}")
        answer = input(
            f"Source column for {canonical}: "
        ).strip()

        if answer and answer in available_columns:
            accepted[canonical] = answer
            claimed_sources.add(answer)
        elif answer:
            print(
                "Column not found or already assigned. "
                "This mapping was skipped."
            )

    report = _build_mapping_report(
        df,
        accepted,
        suggestions,
    )

    return accepted, report


def _build_mapping_report(
    df: pd.DataFrame,
    mapping: dict[str, str],
    suggestions: list[dict],
) -> dict:
    """Build a report for the selected mappings."""

    source_columns = list(df.columns)
    mapped_sources = set(mapping.values())

    missing_required = [
        field
        for field in REQUIRED_FIELDS
        if field not in mapping
    ]

    missing_optional = [
        field
        for field in OPTIONAL_FIELDS
        if field not in mapping
    ]

    return {
        "matched": dict(mapping),
        "missing_required_fields": missing_required,
        "missing_optional_fields": missing_optional,
        "unmapped_source_columns": [
            column
            for column in source_columns
            if column not in mapped_sources
        ],
        "suggestions": suggestions,
    }


def apply_column_mapping(
    df: pd.DataFrame,
    mapping: dict[str, str],
) -> tuple[pd.DataFrame, dict]:
    """Rename accepted source columns to canonical names."""

    source_to_target = {
        source: canonical
        for canonical, source in mapping.items()
    }

    mapped_df = df.rename(
        columns=source_to_target,
    ).copy()

    report = _build_mapping_report(
        df,
        mapping,
        suggest_column_mappings(df),
    )

    return mapped_df, report


def map_columns(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """
    Backward-compatible automatic mapper.

    Exact aliases are mapped automatically. Fuzzy suggestions are
    reported but not applied without review.
    """

    suggestions = suggest_column_mappings(df)
    mapping = {}

    for suggestion in suggestions:
        if suggestion.get("method") != "exact_alias":
            continue

        source = suggestion.get("source")
        canonical = suggestion["canonical"]

        if source is not None:
            mapping[canonical] = source

    return apply_column_mapping(df, mapping)


def map_columns_with_review(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """Review suggestions, then return the mapped DataFrame and report."""

    mapping, _ = review_column_mappings(df)

    return apply_column_mapping(
        df,
        mapping,
    )
