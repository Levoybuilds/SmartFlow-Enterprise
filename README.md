
# SmartFlow Enterprise

**One-Click Business Analytics and Excel Automation**

SmartFlow Enterprise is a Python-based data analytics application designed to automate data profiling, data quality validation, transformation, KPI calculations, and business report generation from structured business data.

## Download the Windows Demo

**[Download SmartFlow Enterprise v1.0.0](https://github.com/Levoybuilds/SmartFlow-Enterprise/releases/tag/v1.0.0)**

Want to try SmartFlow without setting up Python?

1. Open the [v1.0.0 Release](https://github.com/Levoybuilds/SmartFlow-Enterprise/releases/tag/v1.0.0).
2. Download `SmartFlow.zip` from the **Assets** section.
3. Extract the complete ZIP folder.
4. Open the extracted folder and run `SmartFlow.exe`.
5. Select a supported Excel or CSV file and click **Run Analysis**.

**Important:** Keep the `_internal` folder alongside `SmartFlow.exe`. Do not move or distribute the EXE by itself.

- [Browse the source code](https://github.com/Levoybuilds/SmartFlow-Enterprise)
- [View release notes](https://github.com/Levoybuilds/SmartFlow-Enterprise/releases/tag/v1.0.0)

## Key Features

- **Data ingestion:** Work with supported Excel and CSV input files.
- **Data profiling:** Inspect dataset structure, columns, missing values, and data types.
- **Data quality validation:** Identify missing values, duplicates, invalid prices, and unusual quantities.
- **Schema mapping:** Map incoming column names to canonical business fields.
- **Data transformation:** Prepare structured data for analysis.
- **Analytics and KPIs:** Generate analytical datasets and business metrics.
- **Automated reporting:** Produce Excel-based data quality and executive reports.
- **Desktop launcher:** Start the analysis workflow through a desktop interface.

## Technology Stack

- Python
- Pandas
- Excel and CSV processing libraries
- Tkinter
- PyInstaller

## Project Structure

```text
SmartFlow-Enterprise/
├── app/
│   ├── analytics/
│   ├── ingestion/
│   ├── reporting/
│   ├── transformation/
│   └── validation/
├── config/
├── tests/
├── ui/
├── SmartFlow_Launcher.py
├── run.py
├── requirements.txt
└── README.md
```

## Getting Started

### 1. Requirements

- Python installed on your system
- Git
- Windows is recommended for the desktop launcher

### 2. Clone the repository

```bash
git clone https://github.com/Levoybuilds/SmartFlow-Enterprise.git
cd SmartFlow-Enterprise
```

### 3. Create a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 4. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 5. Run SmartFlow

```powershell
python SmartFlow_Launcher.py
```

You can also run the pipeline directly:

```powershell
python run.py
```

Follow the application's input-file selection or configuration process to provide a supported dataset.

## Outputs

Depending on the workflow and configuration, SmartFlow can generate:

- Data quality reports
- Executive Excel reports
- Transaction-level analytical data
- Customer and product summaries
- Monthly sales analytics

Generated files and local datasets are excluded from this public repository by default.

## Project Status

SmartFlow Enterprise is an evolving portfolio project. Features, supported input formats, and setup instructions may change as development continues.

## Author

**Levoy**

Building practical software, automation, and digital solutions for real business problems.
