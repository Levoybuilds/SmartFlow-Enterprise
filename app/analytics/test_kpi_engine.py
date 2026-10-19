from app.ingestion.scanner import discover_input_files
from app.ingestion.profiler import load_excel_file

from app.transformation.sales_transformer import (
    run_transformation,
)

from app.analytics.kpi_engine import (
    calculate_executive_kpis,
    print_kpis,
)


def main():

    files = discover_input_files()

    if not files:
        raise FileNotFoundError(
            "No input file found."
        )

    dataframe = load_excel_file(
        files[0]
    )

    transformations = run_transformation(
        dataframe
    )

    kpis = calculate_executive_kpis(
        transformations["transactions"],
        transformations["metrics"],
    )

    print_kpis(kpis)


if __name__ == "__main__":
    main()