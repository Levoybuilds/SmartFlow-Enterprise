
import os
from pathlib import Path


SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


def get_project_root() -> Path:
    """Return the configured SmartFlow project root directory."""
    configured_root = os.environ.get("SMARTFLOW_PROJECT_ROOT", "").strip()

    if configured_root:
        return Path(configured_root).expanduser().resolve()

    return Path(__file__).resolve().parents[2]


def get_input_directory() -> Path:
    """Return the configured input directory."""
    return get_project_root() / "data" / "input"


def discover_input_files() -> list[Path]:
    """
    Return the selected input file when configured.
    Otherwise, discover supported files in data/input.
    """
    selected_file = os.environ.get("SMARTFLOW_INPUT_FILE", "").strip()

    if selected_file:
        file_path = Path(selected_file).expanduser().resolve()

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Selected input file does not exist: {file_path}"
            )

        if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported input format: {file_path.suffix}. "
                f"Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
            )

        if file_path.name.startswith("~$"):
            raise ValueError(
                "The selected file appears to be a temporary Excel lock file."
            )

        return [file_path]

    input_directory = get_input_directory()

    if not input_directory.exists():
        raise FileNotFoundError(
            f"Input directory does not exist: {input_directory}"
        )

    files = [
        file
        for file in input_directory.iterdir()
        if file.is_file()
        and file.suffix.lower() in SUPPORTED_EXTENSIONS
        and not file.name.startswith("~$")
    ]

    return sorted(files)


def display_discovered_files(files: list[Path]) -> None:
    """Display discovered Excel/CSV files."""

    print()
    print("=" * 60)
    print("SMARTFLOW FILE DISCOVERY")
    print("=" * 60)

    if not files:
        print("No supported input files found.")
        print("Supported formats: .xlsx, .xls, .csv")
        return

    print(f"Files discovered: {len(files)}")
    print()

    for number, file in enumerate(files, start=1):
        size_mb = file.stat().st_size / (1024 * 1024)
        print(f"{number}. {file.name} ({size_mb:.2f} MB)")

    print("=" * 60)


if __name__ == "__main__":
    discovered_files = discover_input_files()
    display_discovered_files(discovered_files)
