# Sales Data Cleaning and Visualization

This project cleans a sample of inconsistent sales data, prints a summary,
creates `sales_analysis.png`, and writes a detailed first-person
`sales_analysis_report.doc` in the project folder. All input data is defined in
`cleanandvisualize.py`; no external data files or services are needed.

## Requirements

- Python 3.15.2 (recorded in `.python-version`)
- pip
- LibreOffice (required to convert the report into the legacy `.doc` format)

`requirements.txt` pins the runtime dependencies used to generate the chart and
Word report. Install LibreOffice for your operating system and make sure its
`soffice` command is available on `PATH`.
`requirements-dev.txt` also installs Ruff for linting and formatting checks.

## Setup

Run these commands from the project folder.

### Windows PowerShell

```powershell
py -3.15 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### macOS or Linux

```sh
python3.15 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

To install the optional development checker as well, use
`requirements-dev.txt` in the same install command.

## Run

### Windows PowerShell

```powershell
.\.venv\Scripts\python.exe cleanandvisualize.py
```

### macOS or Linux

```sh
.venv/bin/python cleanandvisualize.py
```

The script uses a non-interactive plotting backend, so chart generation works
without a graphical window. It opens the Word report using the system's default
document viewer. Later runs reuse the generated files while they remain current;
they don't reopen the report by default. Pass `--force` to regenerate the files
or `--open-report` to open the existing report.

## Expected sample results

- Input rows: 10
- Retained rows: 8
- Total sales: $1,555.00
- Top product: Widget C ($550.00)
- Outputs: `sales_analysis.png`, `sales_analysis_report.doc`

## Validate

With the development dependencies installed:

```sh
python -m ruff check cleanandvisualize.py
python -m ruff format --check cleanandvisualize.py
python -m compileall -q cleanandvisualize.py
```
