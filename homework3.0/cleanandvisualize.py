"""Clean sample sales data and generate a chart and detailed Word report."""

from __future__ import annotations

import argparse
import math
import os
import pathlib
import platform
import shutil
import subprocess
import tempfile
import textwrap
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:
    import pandas as pd

OUTPUT_PATH = pathlib.Path(__file__).with_name("sales_analysis.png")
DOC_REPORT_PATH = pathlib.Path(__file__).with_name("sales_analysis_report.doc")


def _find_libreoffice() -> str | None:
    """Find the LibreOffice executable used to convert DOCX reports to DOC."""
    for executable in ("soffice", "libreoffice"):
        if found := shutil.which(executable):
            return found

    if platform.system() == "Windows":
        program_files = pathlib.Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
        candidates = [
            program_files / "LibreOffice" / "program" / "soffice.exe",
            pathlib.Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        ]
    elif platform.system() == "Darwin":
        candidates = [
            pathlib.Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        ]
    else:
        candidates = []

    return next((str(path) for path in candidates if path.is_file()), None)


def create_sample_data() -> pd.DataFrame:
    """Return the intentionally inconsistent sales data used in this example."""
    import pandas as pd

    return pd.DataFrame(
        {
            "product": [
                "Widget A",
                "Widget B",
                "widget a",
                "Widget C",
                "Widget B",
                "Widget A",
                " Widget C",
                "Widget D",
                None,
                "Widget A",
            ],
            "sales": [
                "150",
                "200",
                "175",
                "300",
                "200",
                "180",
                "250",
                "abc",
                "100",
                "-50",
            ],
            "date": [
                "2025-01-01",
                "2025-01-01",
                "2025-01-02",
                "2025-01-02",
                "2025-01-03",
                "2025-01-03",
                "2025-01-04",
                "2025-01-04",
                "2025-01-05",
                "2025-01-05",
            ],
            "region": [
                "North",
                "South",
                "north",
                "East",
                "South",
                "West",
                "east",
                "North",
                "South",
                "West",
            ],
        }
    )


def _standardize_sales_data(data: pd.DataFrame) -> pd.DataFrame:
    """Standardize values and exclude rows that cannot be analyzed."""
    import pandas as pd

    cleaned = data.copy()
    cleaned["product"] = cleaned["product"].str.strip().str.title().fillna("Unknown")
    cleaned["region"] = cleaned["region"].str.strip().str.title().fillna("Unknown")
    cleaned["sales"] = pd.to_numeric(cleaned["sales"], errors="coerce")
    cleaned["date"] = pd.to_datetime(cleaned["date"], errors="coerce")

    cleaned = cleaned.dropna(subset=["sales", "date"])
    return cleaned.loc[cleaned["sales"] >= 0].reset_index(drop=True)


def clean_sales_data(data: pd.DataFrame) -> pd.DataFrame:
    """Standardize fields, remove invalid rows, and discard exact duplicates."""
    cleaned = _standardize_sales_data(data)
    return cleaned.drop_duplicates().reset_index(drop=True)


def analyze_sales(data: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """Calculate sales totals by product and by date."""
    sales_by_product = (
        data.groupby("product")["sales"].sum().sort_values(ascending=False)
    )
    daily_sales = data.groupby("date")["sales"].sum().sort_index()
    return sales_by_product, daily_sales


def print_summary(data: pd.DataFrame, sales_by_product: pd.Series) -> None:
    """Print a concise summary of the cleaned data."""
    print("=" * 52)
    print("SALES ANALYSIS SUMMARY")
    print("=" * 52)
    print(f"Rows after cleaning: {len(data)}")
    print(f"Total sales:         ${data['sales'].sum():,.2f}")
    print("\nSales by product:")
    print(sales_by_product.to_string())
    print("\nCleaned data preview:")
    print(data.head().to_string(index=False))


def create_doc_report(
    raw_data: pd.DataFrame,
    cleaned_data: pd.DataFrame,
    sales_by_product: pd.Series,
    daily_sales: pd.Series,
    output_path: pathlib.Path = DOC_REPORT_PATH,
    open_after_export: bool = True,
) -> pathlib.Path:
    """Create a legacy .doc report using LibreOffice and optionally open it."""
    import pandas as pd
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
    from docx.styles.style import ParagraphStyle

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal_style = cast(ParagraphStyle, document.styles["Normal"])
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(39, 57, 75)
    for style_name in ("Title", "Heading 1", "Heading 2"):
        style = cast(ParagraphStyle, document.styles[style_name])
        style.font.name = "Aptos Display"
        style.font.color.rgb = RGBColor(24, 50, 75)
    cast(ParagraphStyle, document.styles["Heading 1"]).font.size = Pt(16)
    cast(ParagraphStyle, document.styles["Heading 2"]).font.size = Pt(12)

    total_sales = float(cleaned_data["sales"].sum())
    transaction_count = len(cleaned_data)
    average_sale = float(cleaned_data["sales"].mean()) if transaction_count else 0
    median_sale = float(cleaned_data["sales"].median()) if transaction_count else 0
    top_product = (
        str(sales_by_product.index[0]) if not sales_by_product.empty else "N/A"
    )
    top_product_sales = (
        float(sales_by_product.iloc[0]) if not sales_by_product.empty else 0
    )
    top_product_share = top_product_sales / total_sales if total_sales else 0
    daily_counts = cleaned_data.groupby("date")["sales"].count()
    region_sales = (
        cleaned_data.groupby("region")["sales"].sum().sort_values(ascending=False)
    )
    region_counts = cleaned_data.groupby("region")["sales"].count()
    peak_date = daily_sales.idxmax() if not daily_sales.empty else None
    peak_sales = float(daily_sales.max()) if not daily_sales.empty else 0
    peak_date_text = (
        pd.Timestamp(str(peak_date)).strftime("%b %d, %Y")
        if peak_date is not None
        else "N/A"
    )

    def display_value(value: object) -> str:
        if (
            value is None
            or value is pd.NA
            or value is pd.NaT
            or (isinstance(value, float) and math.isnan(value))
        ):
            return "(missing)"
        return str(value)

    def set_cell_shading(cell, fill: str) -> None:
        properties = cell._tc.get_or_add_tcPr()
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), fill)
        properties.append(shading)

    def add_table(
        headers: list[str],
        rows: list[list[str]],
    ) -> None:
        table = document.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        table.autofit = True
        for cell, heading in zip(table.rows[0].cells, headers):
            cell.text = heading
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_shading(cell, "18324B")
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(8)
        for row_values in rows:
            cells = table.add_row().cells
            for cell, value in zip(cells, row_values):
                cell.text = value
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(2)
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
        document.add_paragraph()

    title = document.add_paragraph(style="Title")
    title.add_run("Sales Data Cleaning & Performance Report")
    subtitle = document.add_paragraph()
    subtitle_run = subtitle.add_run(
        "A personal review of the original records, my cleaning decisions, and results"
    )
    subtitle_run.italic = True
    subtitle_run.font.color.rgb = RGBColor(101, 120, 138)
    document.add_paragraph(
        f"Prepared from {len(raw_data)} supplied records | "
        f"{transaction_count} usable transactions | USD {total_sales:,.2f} total sales"
    )

    document.add_heading("My overall findings", level=1)
    document.add_paragraph(
        f"I started with {len(raw_data)} sales records and kept every source value "
        f"available for comparison. After standardizing names, converting sales and "
        f"dates, and checking each row, I retained {transaction_count} valid, unique "
        f"transactions and excluded {len(raw_data) - transaction_count}. The cleaned "
        f"records total USD {total_sales:,.2f} in sales."
    )
    document.add_paragraph(
        f"{top_product} was the leading product, with "
        f"USD {top_product_sales:,.2f}, or {top_product_share:.1%} of all retained "
        f"sales. The average transaction was USD {average_sale:,.2f}, while the "
        f"median was USD {median_sale:,.2f}. The busiest day was {peak_date_text}, "
        f"when sales reached USD {peak_sales:,.2f}. These results describe only "
        "this sample; I would not treat them as a forecast."
    )

    invalid_sales_count = int(
        pd.to_numeric(raw_data["sales"], errors="coerce").isna().sum()
    )
    invalid_date_count = int(
        pd.to_datetime(raw_data["date"], errors="coerce").isna().sum()
    )
    negative_sales_count = int(
        pd.to_numeric(raw_data["sales"], errors="coerce").lt(0).sum()
    )
    missing_product_count = int(raw_data["product"].isna().sum())
    missing_region_count = int(raw_data["region"].isna().sum())
    standardized = _standardize_sales_data(raw_data)
    duplicate_count = len(standardized) - len(standardized.drop_duplicates())

    document.add_heading("What I cleaned, how, and why", level=1)
    cleaning_points = [
        (
            "Product names: I removed extra spaces and standardized capitalization "
            "so differently typed versions of the same item are grouped together. "
            f"I replaced {missing_product_count} missing product label(s) with "
            "'Unknown' instead of discarding an otherwise valid sale."
        ),
        (
            "Region names: I trimmed spaces and standardized capitalization, "
            f"including entries such as 'east' to 'East'. I found "
            f"{missing_region_count} missing region label(s)."
        ),
        (
            "Sales amounts: I converted text values into numbers. "
            f"{invalid_sales_count} value(s) could not be interpreted as sales, "
            f"and {negative_sales_count} negative amount(s) were excluded."
        ),
        (
            "Dates: I converted date text into actual dates for reliable daily "
            f"grouping. {invalid_date_count} date value(s) were invalid or missing."
        ),
        (
            "Duplicates: I checked for exact duplicate transactions after "
            f"standardizing the fields and found {duplicate_count} duplicate(s)."
        ),
    ]
    for point in cleaning_points:
        document.add_paragraph(point, style="List Bullet")
    document.add_paragraph(
        "I excluded negative sales because this exercise assumes nonnegative "
        "transactions. In real business data, a negative amount could be a valid "
        "refund, so I would confirm the accounting rules before applying this "
        "assumption."
    )

    document.add_heading("Original records and cleaning audit", level=1)
    document.add_paragraph(
        "I kept the source values below and paired each record with its normalized "
        "values or the reason I excluded it."
    )
    seen_signatures: set[tuple[str, float, str, str]] = set()
    audit_rows: list[list[str]] = []
    for row_number, (_, row) in enumerate(raw_data.iterrows(), start=1):
        raw_sales = pd.to_numeric(pd.Series([row["sales"]]), errors="coerce").iloc[0]
        raw_date = pd.to_datetime(pd.Series([row["date"]]), errors="coerce").iloc[0]
        product = (
            "Unknown"
            if pd.isna(row["product"])
            else str(row["product"]).strip().title()
        )
        region = (
            "Unknown" if pd.isna(row["region"]) else str(row["region"]).strip().title()
        )
        original_values = "\n".join(
            (
                f"Product: {display_value(row['product'])}",
                f"Sales: {display_value(row['sales'])}",
                f"Date: {display_value(row['date'])}",
                f"Region: {display_value(row['region'])}",
            )
        )
        if pd.isna(raw_sales):
            outcome = "Excluded: sales are missing or invalid."
        elif pd.isna(raw_date):
            outcome = "Excluded: date is missing or invalid."
        elif raw_sales < 0:
            outcome = "Excluded: negative sales are outside this exercise's scope."
        else:
            signature = (
                product,
                float(raw_sales),
                str(pd.Timestamp(raw_date)),
                region,
            )
            if signature in seen_signatures:
                outcome = "Excluded: duplicate after standardization."
            else:
                seen_signatures.add(signature)
                outcome = "\n".join(
                    (
                        "Included in analysis.",
                        f"Product: {product}",
                        f"Sales: USD {float(raw_sales):,.2f}",
                        f"Date: {pd.Timestamp(raw_date):%Y-%m-%d}",
                        f"Region: {region}",
                    )
                )
        audit_rows.append([f"{row_number:02}", original_values, outcome])
    add_table(
        ["Source row", "Original values", "Cleaned result / decision"], audit_rows
    )

    document.add_heading("Detailed sales results", level=1)
    document.add_heading("Sales by product", level=2)
    product_rows = [
        [
            str(product),
            f"USD {float(amount):,.2f}",
            f"{float(amount) / total_sales:.1%}" if total_sales else "0.0%",
            str(int(cleaned_data.loc[cleaned_data["product"] == product].shape[0])),
        ]
        for product, amount in sales_by_product.items()
    ]
    add_table(
        ["Product", "Sales total", "Share of sales", "Transactions"], product_rows
    )
    document.add_paragraph(
        "I compare both total value and share here. Widget C leads this sample, "
        "while Widget A is close behind; the Unknown category keeps the valid "
        "sale whose product label was missing."
    )

    document.add_heading("Daily sales", level=2)
    daily_rows = [
        [
            pd.Timestamp(str(date)).strftime("%b %d, %Y"),
            f"USD {float(amount):,.2f}",
            str(int(daily_counts.get(date, 0))),
        ]
        for date, amount in daily_sales.items()
    ]
    add_table(["Date", "Sales total", "Transactions"], daily_rows)
    if not daily_sales.empty:
        document.add_paragraph(
            f"I see the daily high on {peak_date_text} at USD {peak_sales:,.2f}. "
            "Daily totals can shift when a day has more transactions or larger "
            "individual sales, so I read this together with the transaction "
            "distribution rather than treating one day as a lasting trend."
        )

    document.add_heading("Sales by region", level=2)
    region_rows = [
        [
            str(region),
            f"USD {float(amount):,.2f}",
            str(int(region_counts.get(region, 0))),
            f"{float(amount) / total_sales:.1%}" if total_sales else "0.0%",
        ]
        for region, amount in region_sales.items()
    ]
    add_table(["Region", "Sales total", "Transactions", "Share of sales"], region_rows)
    document.add_paragraph(
        "I compare regional revenue and transaction counts together. That helps "
        "show whether a larger regional total comes from more transactions, "
        "higher-value transactions, or both."
    )

    document.add_heading("Transaction values and data limits", level=1)
    document.add_paragraph(
        f"The mean transaction was USD {average_sale:,.2f} and the median was "
        f"USD {median_sale:,.2f}. Because these values are close, the center of "
        "the sample is fairly consistent; however, only "
        f"{transaction_count} valid transactions are available, so this is a "
        "small dataset. The histogram shows how those amounts are distributed, "
        "not how future sales will behave."
    )
    document.add_paragraph(
        "The invalid 'abc' amount and the negative amount do not contribute to "
        "the totals. I preserved their original values in the audit so the "
        "exclusions are transparent. Before applying these rules to a larger "
        "dataset, I would review invalid entries with their source and confirm "
        "how refunds and missing product names should be handled."
    )

    if OUTPUT_PATH.exists():
        document.add_heading("Sales dashboard", level=1)
        document.add_picture(str(OUTPUT_PATH), width=Inches(6.7))
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    converter = _find_libreoffice()
    if converter is None:
        raise RuntimeError(
            "LibreOffice is required to create a real legacy .doc report. "
            "Install LibreOffice and ensure 'soffice' is available on PATH."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="sales-report-") as temp_folder:
        temp_dir = pathlib.Path(temp_folder)
        docx_path = temp_dir / f"{output_path.stem}.docx"
        converted_dir = temp_dir / "converted"
        profile_dir = temp_dir / "libreoffice-profile"
        converted_dir.mkdir()
        profile_dir.mkdir()
        document.save(str(docx_path))
        result = subprocess.run(
            [
                converter,
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--headless",
                "--convert-to",
                "doc:MS Word 97",
                "--outdir",
                str(converted_dir),
                str(docx_path),
            ],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        converted_path = converted_dir / output_path.name
        if result.returncode != 0 or not converted_path.is_file():
            diagnostic = (result.stderr or result.stdout).strip()
            raise RuntimeError(
                f"LibreOffice could not create the .doc report. {diagnostic}"
            )
        os.replace(converted_path, output_path)

    if open_after_export:
        _open_report(output_path)
    return output_path


def _open_report(output_path: pathlib.Path) -> None:
    """Open a report with the operating system's default application."""
    resolved_path = str(output_path.resolve())
    try:
        operating_system = platform.system()
        if operating_system == "Windows":
            os.startfile(resolved_path)
        else:
            opener = "open" if operating_system == "Darwin" else "xdg-open"
            subprocess.Popen(
                [opener, resolved_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
    except OSError as error:
        raise RuntimeError(
            f"Word report exists at {resolved_path}, but could not be opened."
        ) from error


def _outputs_are_current() -> bool:
    """Check whether generated outputs are newer than their source script."""
    source_time = pathlib.Path(__file__).stat().st_mtime_ns
    return all(
        path.is_file() and path.stat().st_mtime_ns >= source_time
        for path in (OUTPUT_PATH, DOC_REPORT_PATH)
    )


def create_visualization(
    data: pd.DataFrame,
    sales_by_product: pd.Series,
    daily_sales: pd.Series,
    output_path: pathlib.Path = OUTPUT_PATH,
) -> None:
    """Create a compact, labeled sales dashboard and save an optimized PNG."""
    import matplotlib

    matplotlib.use("Agg")

    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker
    import pandas as pd
    from matplotlib.patches import Rectangle

    background = "#F4F7FB"
    ink = "#18324B"
    muted = "#65788A"
    blue = "#2878A5"
    teal = "#2A9D8F"
    gold = "#E8A23A"
    grid = "#E2E9F0"
    total_sales = float(data["sales"].sum())
    transaction_count = len(data)
    average_sale = float(data["sales"].mean()) if transaction_count else 0
    top_product = (
        str(sales_by_product.index[0]) if not sales_by_product.empty else "N/A"
    )
    top_product_sales = (
        float(sales_by_product.iloc[0]) if not sales_by_product.empty else 0
    )
    top_product_share = top_product_sales / total_sales if total_sales else 0
    region_sales = data.groupby("region")["sales"].sum().sort_values()
    region_counts = data.groupby("region")["sales"].count()
    product_counts = data.groupby("product")["sales"].count()
    median_sale = float(data["sales"].median()) if transaction_count else 0
    peak_date = daily_sales.idxmax() if not daily_sales.empty else None
    peak_sales = float(daily_sales.max()) if not daily_sales.empty else 0
    start_date = (
        data["date"].min().strftime("%b %d, %Y") if transaction_count else "N/A"
    )
    end_date = data["date"].max().strftime("%b %d, %Y") if transaction_count else "N/A"

    with plt.rc_context(
        {
            "font.family": "DejaVu Sans",
            "font.size": 9,
            "axes.titleweight": "bold",
            "axes.labelcolor": muted,
            "xtick.color": muted,
            "ytick.color": muted,
            "text.color": ink,
            "figure.facecolor": background,
            "axes.facecolor": "white",
            "savefig.facecolor": background,
        }
    ):
        fig = plt.figure(figsize=(12, 10.4), layout="constrained")
        grid_spec = fig.add_gridspec(
            nrows=4,
            ncols=2,
            height_ratios=(0.62, 1.35, 1.25, 1.2),
            hspace=0.2,
            wspace=0.24,
        )
        header_ax = fig.add_subplot(grid_spec[0, :])
        product_ax = fig.add_subplot(grid_spec[1, 0])
        daily_ax = fig.add_subplot(grid_spec[1, 1])
        region_ax = fig.add_subplot(grid_spec[2, 0])
        distribution_ax = fig.add_subplot(grid_spec[2, 1])
        summary_ax = fig.add_subplot(grid_spec[3, :])
        header_ax.set_axis_off()
        header_ax.text(
            0.01,
            0.92,
            "SALES PERFORMANCE",
            fontsize=19,
            fontweight="bold",
            color=ink,
            transform=header_ax.transAxes,
            va="top",
        )
        header_ax.text(
            0.012,
            0.64,
            f"Executive dashboard  |  {start_date} – {end_date}",
            fontsize=9,
            color=muted,
            transform=header_ax.transAxes,
            va="top",
        )

        kpis = [
            ("TOTAL SALES", f"${total_sales:,.2f}", blue),
            ("TRANSACTIONS", f"{transaction_count:,}", teal),
            ("AVERAGE SALE", f"${average_sale:,.2f}", gold),
            ("TOP PRODUCT", f"{top_product}  ·  {top_product_share:.1%}", "#7656A5"),
        ]
        for index, (label, value, accent) in enumerate(kpis):
            x_position = 0.012 + index * 0.249
            header_ax.add_patch(
                Rectangle(
                    (x_position, 0.04),
                    0.235,
                    0.39,
                    transform=header_ax.transAxes,
                    facecolor="white",
                    edgecolor="#E4EBF2",
                    linewidth=0.8,
                    clip_on=False,
                )
            )
            header_ax.add_patch(
                Rectangle(
                    (x_position, 0.04),
                    0.006,
                    0.39,
                    transform=header_ax.transAxes,
                    facecolor=accent,
                    edgecolor="none",
                    clip_on=False,
                )
            )
            header_ax.text(
                x_position + 0.018,
                0.34,
                label,
                fontsize=7,
                fontweight="bold",
                color=muted,
                transform=header_ax.transAxes,
                va="top",
            )
            header_ax.text(
                x_position + 0.018,
                0.2,
                value,
                fontsize=12 if index != 3 else 10,
                fontweight="bold",
                color=ink,
                transform=header_ax.transAxes,
                va="top",
            )

        product_values = sales_by_product.sort_values(ascending=True)
        product_labels = product_values.index.astype(str).tolist()
        product_widths = product_values.to_numpy(dtype=float)
        product_max = float(product_values.max()) if not product_values.empty else 0
        product_bars = product_ax.barh(
            product_labels,
            product_widths,
            color=[
                teal if product == top_product else blue for product in product_labels
            ],
            height=0.58,
            zorder=3,
        )
        product_ax.set_title("PRODUCT CONTRIBUTION", loc="left", fontsize=9, pad=12)
        product_ax.set_xlabel("Sales", labelpad=5)
        product_ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
        product_ax.set_xlim(0, max(product_max * 1.58, 1))
        product_ax.xaxis.grid(True, color=grid, linewidth=0.8, zorder=0)
        product_ax.yaxis.grid(False)
        product_ax.set_axisbelow(True)
        for bar, product in zip(product_bars, product_labels):
            value = float(bar.get_width())
            share = value / total_sales if total_sales else 0
            count = int(product_counts.get(product, 0))
            product_ax.text(
                value + float(product_values.max()) * 0.025,
                bar.get_y() + bar.get_height() / 2,
                f"${value:,.0f}  ·  {share:.1%}  ·  {count} txn",
                va="center",
                fontsize=8,
                color=ink,
            )

        daily_ax.set_title("DAILY SALES TREND", loc="left", fontsize=9, pad=12)
        if not daily_sales.empty:
            dates = pd.to_datetime(daily_sales.index)
            daily_values = daily_sales.to_numpy(dtype=float)
            daily_ax.plot(
                dates,
                daily_values,
                color=teal,
                linewidth=2.4,
                marker="o",
                markersize=5,
                markerfacecolor="white",
                markeredgewidth=1.8,
                zorder=3,
            )
            daily_ax.fill_between(dates, daily_values, color=teal, alpha=0.1)
            for date, value in zip(dates, daily_values):
                daily_ax.annotate(
                    f"${value:,.0f}",
                    (date, value),
                    xytext=(0, 9),
                    textcoords="offset points",
                    ha="center",
                    fontsize=8,
                    fontweight="bold",
                    color=ink,
                )
            daily_ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
            daily_ax.xaxis.set_major_locator(
                mdates.AutoDateLocator(minticks=3, maxticks=6)
            )
            daily_ax.set_ylim(0, max(daily_values) * 1.22)
        daily_ax.set_ylabel("Sales")
        daily_ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
        daily_ax.yaxis.grid(True, color=grid, linewidth=0.8, zorder=0)
        daily_ax.xaxis.grid(False)
        daily_ax.set_axisbelow(True)
        daily_ax.tick_params(axis="x", rotation=0)

        region_bars = region_ax.barh(
            region_sales.index.astype(str),
            region_sales.to_numpy(dtype=float),
            color=blue,
            height=0.58,
            zorder=3,
        )
        region_ax.set_title("REGIONAL PERFORMANCE", loc="left", fontsize=9, pad=12)
        region_ax.set_xlabel("Sales", labelpad=5)
        region_ax.xaxis.set_major_formatter(mticker.StrMethodFormatter("${x:,.0f}"))
        region_max = float(region_sales.max()) if not region_sales.empty else 0
        region_ax.set_xlim(0, max(region_max * 1.58, 1))
        region_ax.xaxis.grid(True, color=grid, linewidth=0.8, zorder=0)
        region_ax.yaxis.grid(False)
        region_ax.set_axisbelow(True)
        for bar, region in zip(region_bars, region_sales.index):
            value = float(bar.get_width())
            region_ax.text(
                value + float(region_sales.max()) * 0.025,
                bar.get_y() + bar.get_height() / 2,
                f"${value:,.0f}  ·  {int(region_counts[region])} txn",
                va="center",
                fontsize=8,
                color=ink,
            )

        distribution_ax.set_title(
            "TRANSACTION VALUE DISTRIBUTION", loc="left", fontsize=9, pad=12
        )
        distribution_ax.hist(
            data["sales"],
            bins=min(5, max(1, transaction_count)),
            color=gold,
            edgecolor="white",
            linewidth=1.2,
            zorder=3,
        )
        distribution_ax.axvline(
            average_sale,
            color=blue,
            linestyle="--",
            linewidth=1.8,
            label=f"Mean  ${average_sale:,.2f}",
        )
        distribution_ax.axvline(
            median_sale,
            color="#7656A5",
            linestyle=":",
            linewidth=1.8,
            label=f"Median  ${median_sale:,.2f}",
        )
        distribution_ax.set_xlabel("Individual transaction amount")
        distribution_ax.set_ylabel("Transactions")
        distribution_ax.xaxis.set_major_formatter(
            mticker.StrMethodFormatter("${x:,.0f}")
        )
        distribution_ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        distribution_ax.yaxis.grid(True, color=grid, linewidth=0.8, zorder=0)
        distribution_ax.xaxis.grid(False)
        distribution_ax.set_axisbelow(True)
        distribution_ax.legend(frameon=False, loc="upper right", fontsize=8)

        for axis in (product_ax, daily_ax, region_ax, distribution_ax):
            axis.spines[["top", "right", "left"]].set_visible(False)
            axis.spines["bottom"].set_color(grid)
            axis.tick_params(axis="both", length=0, pad=6)

        summary_ax.set_axis_off()
        summary_ax.add_patch(
            Rectangle(
                (0, 0),
                1,
                1,
                transform=summary_ax.transAxes,
                facecolor="white",
                edgecolor="#E4EBF2",
                linewidth=0.8,
                clip_on=False,
            )
        )
        summary_ax.text(
            0.018,
            0.94,
            "SALES SUMMARY",
            fontsize=9,
            fontweight="bold",
            color=ink,
            transform=summary_ax.transAxes,
            va="top",
        )
        peak_date_text = (
            pd.Timestamp(str(peak_date)).strftime("%b %d, %Y")
            if peak_date is not None
            else "N/A"
        )
        leading_region = (
            str(region_sales.index[-1]) if not region_sales.empty else "N/A"
        )
        leading_region_sales = (
            float(region_sales.iloc[-1]) if not region_sales.empty else 0
        )
        overview_text = textwrap.fill(
            (
                f"Sales totaled USD {total_sales:,.2f} across {transaction_count} "
                f"transactions from {start_date} to {end_date}. {top_product} "
                f"was the best-selling product (USD {top_product_sales:,.0f}; "
                f"{top_product_share:.1%} of sales), {leading_region} led regions "
                f"(USD {leading_region_sales:,.0f}), and the busiest day was "
                f"{peak_date_text} (USD {peak_sales:,.0f}). A typical sale was "
                f"USD {average_sale:,.2f} on average (median USD {median_sale:,.2f})."
            ),
            width=145,
        )
        product_text = (
            "Bars compare product sales; labels show sales share and transaction count."
        )
        daily_text = "The line shows combined sales for each date."
        region_text = "Bars compare regional sales; labels show transaction counts."
        distribution_text = (
            "Bars count sales by value range; dashed and dotted lines mark average "
            "and median."
        )
        summary_ax.text(
            0.018,
            0.76,
            overview_text,
            fontsize=7.6,
            color=ink,
            transform=summary_ax.transAxes,
            va="top",
        )
        summary_sections = (
            (0.018, 0.48, "PRODUCT CONTRIBUTION", product_text),
            (0.515, 0.48, "DAILY SALES TREND", daily_text),
            (0.018, 0.22, "REGIONAL PERFORMANCE", region_text),
            (0.515, 0.22, "TRANSACTION VALUE DISTRIBUTION", distribution_text),
        )
        for x_position, y_position, label, details in summary_sections:
            summary_ax.text(
                x_position,
                y_position,
                label,
                fontsize=7,
                fontweight="bold",
                color=muted,
                transform=summary_ax.transAxes,
                va="top",
            )
            summary_ax.text(
                x_position,
                y_position - 0.07,
                textwrap.fill(details, width=75),
                fontsize=7,
                color=ink,
                transform=summary_ax.transAxes,
                va="top",
            )

        fig.savefig(
            str(output_path),
            dpi=150,
            bbox_inches=None,
            facecolor=background,
            pil_kwargs={"optimize": True},
        )
        plt.close(fig)


def main(force: bool = False, open_report: bool = False) -> None:
    """Run the sales cleaning, analysis, and reporting workflow."""
    if not force and _outputs_are_current():
        print("Sales analysis outputs are current; use --force to rebuild them.")
        print(f"Chart: {OUTPUT_PATH}")
        print(f"Report: {DOC_REPORT_PATH}")
        if open_report:
            _open_report(DOC_REPORT_PATH)
        return
    raw_data = create_sample_data()
    cleaned_data = clean_sales_data(raw_data)
    sales_by_product, daily_sales = analyze_sales(cleaned_data)
    print_summary(cleaned_data, sales_by_product)
    create_visualization(cleaned_data, sales_by_product, daily_sales)
    print(f"\nVisualization saved to: {OUTPUT_PATH}")
    report_path = create_doc_report(
        raw_data,
        cleaned_data,
        sales_by_product,
        daily_sales,
        open_after_export=True,
    )
    print(f"Detailed Word report opened: {report_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--force",
        action="store_true",
        help="regenerate the chart and report even when outputs are current",
    )
    parser.add_argument(
        "--open-report",
        action="store_true",
        help="open the existing report when outputs are already current",
    )
    arguments = parser.parse_args()
    main(force=arguments.force, open_report=arguments.open_report)
