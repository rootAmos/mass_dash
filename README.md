# AstroM Mass Tools

This repository contains the MassProps workbook and a Streamlit dashboard for reviewing CSV weight exports from `MassExports/`.

## Setup

Create a virtual environment from the repository root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

The editable install reads dependencies from `pyproject.toml` and installs Streamlit, Pandas, and Plotly.

## Run the Dashboard

```powershell
streamlit run streamlit_app.py
```

The app loads every `*.csv` file in `MassExports/`, treats each file as an export iteration, and plots:

- total mass and center of gravity metrics for the selected iteration
- ATA or component weight breakdowns
- total mass trend across selected export iterations
- source rows for inspection

## Export Format

CSV exports should include these columns:

```text
ata, component, total_mass_kg, mx_kgm, my_kgm, mz_kgm
```

Additional columns such as `version`, `status`, `qty`, `unit_mass_kg`, `x_m`, `y_m`, `z_m`, and `notes` are displayed when present.

## Git Ignore

Workbook files are ignored with `*.xlsm`, including temporary Excel lock files matching `~$*.xlsm`. CSV exports remain trackable so dashboard inputs can be versioned.
