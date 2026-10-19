from pathlib import Path

import pandas as pd
import xlsxwriter

from app.ingestion.scanner import discover_input_files, get_project_root
from app.ingestion.profiler import load_excel_file
from app.transformation.sales_transformer import run_transformation
from app.analytics.kpi_engine import calculate_executive_kpis


# ============================================================
# PATHS
# ============================================================

PROJECT_ROOT = get_project_root()

OUTPUT_DIR = PROJECT_ROOT / "output"

REPORT_PATH = (
    OUTPUT_DIR
    / "SmartFlow_Executive_Report.xlsx"
)


# ============================================================
# ENTERPRISE COLOR SYSTEM
# ============================================================

COLORS = {
    "navy": "#172554",
    "blue": "#2563EB",
    "blue_light": "#DBEAFE",

    "teal": "#0F766E",
    "teal_light": "#CCFBF1",

    "purple": "#7C3AED",
    "purple_light": "#EDE9FE",

    "orange": "#EA580C",
    "orange_light": "#FFEDD5",

    "green": "#059669",
    "green_light": "#D1FAE5",

    "red": "#DC2626",
    "red_light": "#FEE2E2",

    "amber": "#D97706",
    "amber_light": "#FEF3C7",

    "slate": "#475569",
    "gray": "#64748B",
    "gray_light": "#F1F5F9",

    "border": "#E2E8F0",
    "white": "#FFFFFF",
    "background": "#F8FAFC",
    "black": "#0F172A",
}


# ============================================================
# FORMATTING
# ============================================================

def create_formats(workbook):

    formats = {}

    formats["title"] = workbook.add_format({
        "bold": True,
        "font_size": 24,
        "font_color": COLORS["white"],
        "bg_color": COLORS["navy"],
        "align": "left",
        "valign": "vcenter",
    })

    formats["subtitle"] = workbook.add_format({
        "font_size": 10,
        "font_color": "#CBD5E1",
        "bg_color": COLORS["navy"],
        "align": "left",
        "valign": "vcenter",
    })

    formats["section"] = workbook.add_format({
        "bold": True,
        "font_size": 12,
        "font_color": COLORS["navy"],
        "bg_color": COLORS["white"],
        "bottom": 2,
        "bottom_color": COLORS["blue"],
        "align": "left",
    })

    formats["card_label_blue"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["blue"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_blue"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["blue"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "£#,##0.00",
    })

    formats["card_label_teal"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["teal"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_teal"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["teal"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    formats["card_label_purple"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["purple"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_purple"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["purple"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    formats["card_label_orange"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["orange"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_orange"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["orange"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    formats["card_label_green"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["green"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_green"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["green"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "£#,##0.00",
    })

    formats["card_label_red"] = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": COLORS["white"],
        "bg_color": COLORS["red"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["card_value_red"] = workbook.add_format({
        "bold": True,
        "font_size": 19,
        "font_color": COLORS["red"],
        "bg_color": COLORS["white"],
        "border": 1,
        "border_color": COLORS["border"],
        "align": "center",
        "valign": "vcenter",
        "num_format": "0.00%",
    })

    formats["currency"] = workbook.add_format({
        "num_format": "£#,##0.00",
    })

    formats["number"] = workbook.add_format({
        "num_format": "#,##0",
    })

    formats["percent"] = workbook.add_format({
        "num_format": "0.00%",
    })

    formats["table_header"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["white"],
        "bg_color": COLORS["navy"],
        "border": 1,
        "border_color": COLORS["navy"],
        "align": "center",
        "valign": "vcenter",
    })

    formats["body"] = workbook.add_format({
        "font_color": COLORS["black"],
        "border": 1,
        "border_color": COLORS["border"],
    })

    formats["body_alt"] = workbook.add_format({
        "font_color": COLORS["black"],
        "bg_color": "#F8FAFC",
        "border": 1,
        "border_color": COLORS["border"],
    })

    formats["alert_red"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["red"],
        "bg_color": COLORS["red_light"],
        "border": 1,
        "border_color": "#FCA5A5",
        "align": "center",
        "valign": "vcenter",
    })

    formats["alert_amber"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["amber"],
        "bg_color": COLORS["amber_light"],
        "border": 1,
        "border_color": "#FCD34D",
        "align": "center",
        "valign": "vcenter",
    })

    formats["alert_green"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["green"],
        "bg_color": COLORS["green_light"],
        "border": 1,
        "border_color": "#6EE7B7",
        "align": "center",
        "valign": "vcenter",
    })

    formats["snapshot_label"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["gray"],
        "font_size": 9,
    })

    formats["snapshot_value"] = workbook.add_format({
        "bold": True,
        "font_color": COLORS["black"],
        "font_size": 11,
    })

    formats["small"] = workbook.add_format({
        "font_size": 9,
        "font_color": COLORS["gray"],
    })

    return formats


# ============================================================
# HELPERS
# ============================================================

def write_card(
    worksheet,
    formats,
    cell_range,
    label,
    value,
    label_format,
    value_format,
):
    """
    Write a two-row KPI card.

    Example:
        ("B6", "C6")

    produces:

        B6:C6  -> label
        B7:C7  -> value
    """

    from xlsxwriter.utility import xl_cell_to_rowcol

    start_cell, end_cell = cell_range

    # Convert Excel A1 references to zero-based
    # row / column coordinates.
    first_row, first_col = xl_cell_to_rowcol(
        start_cell
    )

    last_row, last_col = xl_cell_to_rowcol(
        end_cell
    )

    # Label row
    worksheet.merge_range(
        first_row,
        first_col,
        last_row,
        last_col,
        label,
        label_format,
    )

    # Value row
    worksheet.merge_range(
        first_row + 1,
        first_col,
        last_row + 1,
        last_col,
        value,
        value_format,
    )


def prepare_data(transformations):

    monthly = transformations[
        "monthly_sales"
    ].copy()

    countries = transformations[
        "country_summary"
    ].copy()

    products = transformations[
        "product_summary"
    ].copy()

    customers = transformations[
        "customer_summary"
    ].copy()

    for dataframe in [
        monthly,
        countries,
        products,
        customers,
    ]:

        if "Revenue" in dataframe.columns:
            dataframe["Revenue"] = (
                dataframe["Revenue"]
                .round(2)
            )

        if "UnitPrice" in dataframe.columns:
            dataframe["UnitPrice"] = (
                dataframe["UnitPrice"]
                .round(2)
            )

    return (
        monthly,
        countries,
        products,
        customers,
    )


def calculate_quality_metrics(dataframe):

    duplicate_rows = int(
        dataframe.duplicated().sum()
    )

    missing_customer = int(
        dataframe["CustomerID"]
        .isna()
        .sum()
    )

    missing_description = int(
        dataframe["Description"]
        .isna()
        .sum()
    )

    invalid_price = int(
        (dataframe["UnitPrice"] <= 0)
        .sum()
    )

    negative_quantity = int(
        (dataframe["Quantity"] < 0)
        .sum()
    )

    cancellation_rows = int(
        dataframe["InvoiceNo"]
        .astype(str)
        .str.startswith("C")
        .sum()
    )

    zero_value_adjustments = int(
        (
            (dataframe["Quantity"] < 0)
            &
            (
                ~dataframe["InvoiceNo"]
                .astype(str)
                .str.startswith("C")
            )
            &
            (dataframe["UnitPrice"] <= 0)
        ).sum()
    )

    return {
        "duplicate_rows": duplicate_rows,
        "missing_customer": missing_customer,
        "missing_description": missing_description,
        "invalid_price": invalid_price,
        "negative_quantity": negative_quantity,
        "cancellation_rows": cancellation_rows,
        "zero_value_adjustments":
            zero_value_adjustments,
    }


def calculate_health_score(quality):

    score = 100

    total_rows = max(
        quality.get("total_rows", 1),
        1,
    )

    duplicate_rate = (
        quality["duplicate_rows"]
        / total_rows
    )

    missing_customer_rate = (
        quality["missing_customer"]
        / total_rows
    )

    missing_description_rate = (
        quality["missing_description"]
        / total_rows
    )

    invalid_price_rate = (
        quality["invalid_price"]
        / total_rows
    )

    score -= min(
        duplicate_rate * 100 * 20,
        20,
    )

    score -= min(
        missing_customer_rate * 100 * 20,
        20,
    )

    score -= min(
        missing_description_rate * 100 * 15,
        15,
    )

    score -= min(
        invalid_price_rate * 100 * 20,
        20,
    )

    return max(
        0,
        min(
            100,
            round(score),
        ),
    )


# ============================================================
# DASHBOARD
# ============================================================

def build_dashboard(
    workbook,
    formats,
    kpis,
    monthly,
    countries,
    products,
    quality,
):

    sheet = workbook.add_worksheet(
        "Executive Dashboard"
    )

    sheet.hide_gridlines(2)

    sheet.set_tab_color(
        COLORS["blue"]
    )

    # --------------------------------------------------------
    # PAGE
    # --------------------------------------------------------

    sheet.set_column("A:A", 2)
    sheet.set_column("B:B", 16)
    sheet.set_column("C:C", 3)
    sheet.set_column("D:D", 16)
    sheet.set_column("E:E", 3)
    sheet.set_column("F:F", 16)
    sheet.set_column("G:G", 3)
    sheet.set_column("H:H", 16)
    sheet.set_column("I:I", 3)
    sheet.set_column("J:J", 16)
    sheet.set_column("K:K", 3)
    sheet.set_column("L:L", 16)

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    sheet.merge_range(
        "B2:L3",
        "SMARTFLOW ENTERPRISE",
        formats["title"],
    )

    sheet.merge_range(
        "B4:L4",
        "EXECUTIVE SALES INTELLIGENCE  •  AUTOMATED ANALYTICS",
        formats["subtitle"],
    )

    sheet.set_row(1, 28)
    sheet.set_row(2, 28)
    sheet.set_row(3, 18)

    # --------------------------------------------------------
    # KPI CARDS
    # --------------------------------------------------------

    write_card(
        sheet,
        formats,
        ("B6", "C6"),
        "NET SALES",
        kpis.net_sales,
        formats["card_label_blue"],
        formats["card_value_blue"],
    )

    write_card(
        sheet,
        formats,
        ("D6", "E6"),
        "ORDERS",
        kpis.orders,
        formats["card_label_teal"],
        formats["card_value_teal"],
    )

    write_card(
        sheet,
        formats,
        ("F6", "G6"),
        "CUSTOMERS",
        kpis.customers,
        formats["card_label_purple"],
        formats["card_value_purple"],
    )

    write_card(
        sheet,
        formats,
        ("H6", "I6"),
        "PRODUCTS",
        kpis.products,
        formats["card_label_orange"],
        formats["card_value_orange"],
    )

    write_card(
        sheet,
        formats,
        ("J6", "K6"),
        "AVERAGE ORDER VALUE",
        kpis.average_order_value,
        formats["card_label_green"],
        formats["card_value_green"],
    )

    write_card(
        sheet,
        formats,
        ("L6", "L6"),
        "CANCEL %",
        kpis.cancellation_rate / 100,
        formats["card_label_red"],
        formats["card_value_red"],
    )

    sheet.set_row(5, 22)
    sheet.set_row(6, 34)

    # --------------------------------------------------------
    # REVENUE SECTION
    # --------------------------------------------------------

    sheet.merge_range(
        "B10:H10",
        "REVENUE PERFORMANCE",
        formats["section"],
    )

    revenue_chart = workbook.add_chart({
        "type": "line",
    })

    revenue_chart.add_series({
        "name": "Revenue",
        "categories": [
            "Monthly Trends",
            1,
            0,
            len(monthly),
            0,
        ],
        "values": [
            "Monthly Trends",
            1,
            1,
            len(monthly),
            1,
        ],
        "line": {
            "color": COLORS["blue"],
            "width": 3,
        },
        "marker": {
            "type": "circle",
            "size": 6,
            "border": {
                "color": COLORS["blue"],
            },
            "fill": {
                "color": COLORS["white"],
            },
        },
    })

    revenue_chart.set_title({
        "name": "Monthly Revenue",
        "name_font": {
            "size": 11,
            "bold": True,
        },
    })

    revenue_chart.set_y_axis({
        "num_format": "£#,##0",
        "major_gridlines": {
            "visible": False,
        },
    })

    revenue_chart.set_x_axis({
        "label_position": "low",
    })

    revenue_chart.set_legend({
        "none": True,
    })

    revenue_chart.set_chartarea({
        "border": {
            "none": True,
        },
        "fill": {
            "color": COLORS["white"],
        },
    })

    revenue_chart.set_plotarea({
        "border": {
            "none": True,
        },
        "fill": {
            "color": COLORS["white"],
        },
    })

    revenue_chart.set_size({
        "width": 620,
        "height": 300,
    })

    sheet.insert_chart(
        "B11",
        revenue_chart,
    )

    # --------------------------------------------------------
    # BUSINESS SNAPSHOT
    # --------------------------------------------------------

    sheet.merge_range(
        "I10:L10",
        "BUSINESS SNAPSHOT",
        formats["section"],
    )

    max_month = monthly.loc[
        monthly["Revenue"].idxmax()
    ]

    top_country = countries.iloc[0]

    top_product = products.iloc[0]

    snapshot_rows = [
        (
            "Highest Revenue Month",
            str(max_month["YearMonth"]),
        ),
        (
            "Peak Monthly Revenue",
            f"£{max_month['Revenue']:,.2f}",
        ),
        (
            "Top Market",
            str(top_country["Country"]),
        ),
        (
            "Top Market Revenue",
            f"£{top_country['Revenue']:,.2f}",
        ),
        (
            "Top Product",
            str(
                top_product.get(
                    "Description",
                    top_product.iloc[0],
                )
            )[:32],
        ),
        (
            "Units Sold",
            f"{kpis.units_sold:,}",
        ),
    ]

    row = 11

    for label, value in snapshot_rows:

        sheet.write(
            row,
            8,
            label,
            formats["snapshot_label"],
        )

        sheet.merge_range(
            row,
            9,
            row,
            11,
            value,
            formats["snapshot_value"],
        )

        row += 2

    # --------------------------------------------------------
    # TOP MARKETS
    # --------------------------------------------------------

    sheet.merge_range(
        "B29:F29",
        "TOP MARKETS",
        formats["section"],
    )

    top_countries = countries.head(6)

    country_chart = workbook.add_chart({
        "type": "bar",
    })

    country_chart.add_series({
        "name": "Revenue",
        "categories": [
            "Market Analysis",
            1,
            0,
            len(top_countries),
            0,
        ],
        "values": [
            "Market Analysis",
            1,
            1,
            len(top_countries),
            1,
        ],
        "fill": {
            "color": COLORS["teal"],
        },
        "border": {
            "none": True,
        },
    })

    country_chart.set_title({
        "name": "Top Markets by Revenue",
    })

    country_chart.set_x_axis({
        "num_format": "£#,##0",
        "major_gridlines": {
            "visible": False,
        },
    })

    country_chart.set_legend({
        "none": True,
    })

    country_chart.set_chartarea({
        "border": {
            "none": True,
        },
        "fill": {
            "color": COLORS["white"],
        },
    })

    country_chart.set_plotarea({
        "border": {
            "none": True,
        },
        "fill": {
            "color": COLORS["white"],
        },
    })

    country_chart.set_size({
        "width": 450,
        "height": 280,
    })

    sheet.insert_chart(
        "B30",
        country_chart,
    )

    # --------------------------------------------------------
    # TOP PRODUCTS
    # --------------------------------------------------------

    sheet.merge_range(
        "G29:L29",
        "TOP PRODUCTS",
        formats["section"],
    )

    top_products = products.head(8)

    headers = [
        "RANK",
        "PRODUCT",
        "REVENUE",
    ]

    sheet.write_row(
        "G30",
        headers,
        formats["table_header"],
    )

    for index, (_, product) in enumerate(
        top_products.iterrows(),
        start=1,
    ):

        excel_row = 30 + index

        description = str(
            product.get(
                "Description",
                "Unknown",
            )
        )

        description = description[:38]

        sheet.write(
            excel_row,
            6,
            index,
            formats["body"],
        )

        sheet.write(
            excel_row,
            7,
            description,
            formats["body"],
        )

        sheet.write(
            excel_row,
            8,
            product["Revenue"],
            formats["currency"],
        )

    sheet.set_column(
        "H:H",
        30,
    )

    # --------------------------------------------------------
    # DATA QUALITY
    # --------------------------------------------------------

    sheet.merge_range(
        "B49:L49",
        "DATA QUALITY COMMAND CENTER",
        formats["section"],
    )

    quality_cards = [
        (
            "DUPLICATE RECORDS",
            quality["duplicate_rows"],
            formats["alert_red"],
        ),
        (
            "MISSING CUSTOMER ID",
            quality["missing_customer"],
            formats["alert_amber"],
        ),
        (
            "MISSING DESCRIPTION",
            quality["missing_description"],
            formats["alert_amber"],
        ),
        (
            "INVALID PRICES",
            quality["invalid_price"],
            formats["alert_red"],
        ),
        (
            "ZERO-VALUE ADJUSTMENTS",
            quality["zero_value_adjustments"],
            formats["alert_amber"],
        ),
    ]

    columns = [
        "B",
        "D",
        "F",
        "H",
        "J",
    ]

    for column, (
        label,
        value,
        alert_format,
    ) in zip(
        columns,
        quality_cards,
    ):

        sheet.merge_range(
            f"{column}51:{column}51",
            label,
            alert_format,
        )

        sheet.merge_range(
            f"{column}52:{column}53",
            f"{value:,}",
            alert_format,
        )

    sheet.merge_range(
        "L51:L51",
        "HEALTH",
        formats["alert_green"],
    )

    health = calculate_health_score(
        {
            **quality,
            "total_rows":
                quality["total_rows"],
        }
    )

    health_format = (
        formats["alert_green"]
        if health >= 85
        else formats["alert_amber"]
        if health >= 70
        else formats["alert_red"]
    )

    sheet.merge_range(
        "L52:L53",
        f"{health}/100",
        health_format,
    )

    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------

    sheet.merge_range(
        "B56:L56",
        "SmartFlow Enterprise  •  Automated Business Intelligence  •  Generated from source data",
        formats["small"],
    )


# ============================================================
# DETAIL SHEETS
# ============================================================

def build_detail_sheet(
    workbook,
    name,
    dataframe,
    tab_color,
    formats,
):

    worksheet = workbook.add_worksheet(
        name
    )

    worksheet.hide_gridlines(2)

    worksheet.set_tab_color(
        tab_color
    )

    for column_number, column_name in enumerate(
        dataframe.columns
    ):

        worksheet.write(
            0,
            column_number,
            column_name,
            formats["table_header"],
        )

    for row_number, row in enumerate(
        dataframe.itertuples(
            index=False
        ),
        start=1,
    ):

        for column_number, value in enumerate(
            row
        ):

            if pd.isna(value):
                value = ""

            worksheet.write(
                row_number,
                column_number,
                value,
                formats["body_alt"]
                if row_number % 2 == 0
                else formats["body"],
            )

    worksheet.freeze_panes(
        1,
        0,
    )

    worksheet.autofilter(
        0,
        0,
        len(dataframe),
        len(dataframe.columns) - 1,
    )

    for column_number, column_name in enumerate(
        dataframe.columns
    ):

        width = max(
            len(str(column_name)) + 3,
            14,
        )

        worksheet.set_column(
            column_number,
            column_number,
            min(width, 40),
        )


# ============================================================
# DATA QUALITY SHEET
# ============================================================

def build_quality_sheet(
    workbook,
    quality,
    formats,
):

    worksheet = workbook.add_worksheet(
        "Data Quality"
    )

    worksheet.hide_gridlines(2)

    worksheet.set_tab_color(
        COLORS["red"]
    )

    worksheet.merge_range(
        "A1:F2",
        "SMARTFLOW DATA QUALITY",
        formats["title"],
    )

    worksheet.merge_range(
        "A3:F3",
        "Automated quality assessment of the source dataset",
        formats["subtitle"],
    )

    rows = [
        (
            "Duplicate Records",
            quality["duplicate_rows"],
            "Potential duplicate rows",
        ),
        (
            "Missing CustomerID",
            quality["missing_customer"],
            "Transactions without customer identifier",
        ),
        (
            "Missing Description",
            quality["missing_description"],
            "Transactions without product description",
        ),
        (
            "Invalid UnitPrice",
            quality["invalid_price"],
            "Records with zero or negative UnitPrice",
        ),
        (
            "Negative Quantity",
            quality["negative_quantity"],
            "Transactions containing negative quantity",
        ),
        (
            "Cancellation Rows",
            quality["cancellation_rows"],
            "Rows identified as cancellations",
        ),
        (
            "Zero-Value Adjustments",
            quality["zero_value_adjustments"],
            "Negative quantity with zero/invalid price",
        ),
    ]

    worksheet.write_row(
        "A5",
        [
            "Check",
            "Count",
            "Description",
        ],
        formats["table_header"],
    )

    for row_number, (
        check,
        count,
        description,
    ) in enumerate(
        rows,
        start=5,
    ):

        worksheet.write(
            row_number,
            0,
            check,
            formats["body"],
        )

        worksheet.write(
            row_number,
            1,
            count,
            formats["number"],
        )

        worksheet.write(
            row_number,
            2,
            description,
            formats["body"],
        )

    worksheet.set_column(
        "A:A",
        28,
    )

    worksheet.set_column(
        "B:B",
        18,
    )

    worksheet.set_column(
        "C:C",
        65,
    )

    health = calculate_health_score(
        {
            **quality,
            "total_rows":
                quality["total_rows"],
        }
    )

    worksheet.merge_range(
        "A15:C16",
        f"DATA HEALTH SCORE\n{health}/100",
        formats["alert_green"]
        if health >= 85
        else formats["alert_amber"]
        if health >= 70
        else formats["alert_red"],
    )


# ============================================================
# METHODOLOGY
# ============================================================

def build_methodology(
    workbook,
    formats,
):

    worksheet = workbook.add_worksheet(
        "Methodology"
    )

    worksheet.hide_gridlines(2)

    worksheet.set_tab_color(
        COLORS["slate"]
    )

    worksheet.merge_range(
        "A1:F2",
        "SMARTFLOW METHODOLOGY",
        formats["title"],
    )

    worksheet.merge_range(
        "A3:F3",
        "Definitions used by the automated analytics engine",
        formats["subtitle"],
    )

    definitions = [
        (
            "Gross Sales",
            "Valid positive sales based on Quantity > 0 and UnitPrice > 0.",
        ),
        (
            "Net Sales",
            "Gross sales adjusted by validated monetary returns.",
        ),
        (
            "Orders",
            "Unique invoice numbers associated with valid sales.",
        ),
        (
            "Cancelled Orders",
            "Unique invoice numbers identified as cancellation transactions.",
        ),
        (
            "Customers",
            "Unique CustomerID values associated with valid transactions.",
        ),
        (
            "Products",
            "Unique StockCode values represented in analytical sales data.",
        ),
        (
            "Average Order Value",
            "Gross sales divided by valid orders.",
        ),
        (
            "Cancellation Rate",
            "Cancelled orders divided by sales plus cancelled orders.",
        ),
        (
            "Zero-Value Adjustment",
            "Negative quantity record with zero or invalid UnitPrice that is not a cancellation.",
        ),
    ]

    worksheet.write_row(
        "A5",
        ["Metric", "Definition"],
        formats["table_header"],
    )

    for row, (
        metric,
        definition,
    ) in enumerate(
        definitions,
        start=5,
    ):

        worksheet.write(
            row,
            0,
            metric,
            formats["body"],
        )

        worksheet.write(
            row,
            1,
            definition,
            formats["body"],
        )

    worksheet.set_column(
        "A:A",
        28,
    )

    worksheet.set_column(
        "B:B",
        100,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print("SMARTFLOW ENTERPRISE REPORT ENGINE V2")
    print("=" * 70)

    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No input Excel file found."
        )

    input_file = files[0]

    print()
    print(
        f"Source: {input_file.name}"
    )

    dataframe = load_excel_file(
        input_file
    )

    print(
        "Running transformation engine..."
    )

    transformations = run_transformation(
        dataframe
    )

    print(
        "Calculating executive KPIs..."
    )

    kpis = calculate_executive_kpis(
        transformations["transactions"],
        transformations["metrics"],
    )

    print(
        "Preparing analytical datasets..."
    )

    (
        monthly,
        countries,
        products,
        customers,
    ) = prepare_data(
        transformations
    )

    print(
        "Running data-quality intelligence..."
    )

    quality = calculate_quality_metrics(
        dataframe
    )

    quality["total_rows"] = len(
        dataframe
    )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "Building professional Excel dashboard..."
    )

    workbook = xlsxwriter.Workbook(
        REPORT_PATH
    )

    formats = create_formats(
        workbook
    )

    # Dashboard
    build_dashboard(
        workbook,
        formats,
        kpis,
        monthly,
        countries,
        products,
        quality,
    )

    # Detail sheets
    build_detail_sheet(
        workbook,
        "Monthly Trends",
        monthly,
        COLORS["blue"],
        formats,
    )

    build_detail_sheet(
        workbook,
        "Market Analysis",
        countries,
        COLORS["teal"],
        formats,
    )

    build_detail_sheet(
        workbook,
        "Product Performance",
        products,
        COLORS["orange"],
        formats,
    )

    build_detail_sheet(
        workbook,
        "Customer Intelligence",
        customers,
        COLORS["purple"],
        formats,
    )

    build_quality_sheet(
        workbook,
        quality,
        formats,
    )

    build_methodology(
        workbook,
        formats,
    )

    workbook.close()

    print()
    print("=" * 70)
    print("SMARTFLOW V2 REPORT GENERATED")
    print("=" * 70)
    print()
    print(
        f"Output: {REPORT_PATH}"
    )
    print()
    print(
        "Dashboard: Executive Dashboard"
    )
    print(
        "Sheets: 7"
    )
    print()


if __name__ == "__main__":
    main()