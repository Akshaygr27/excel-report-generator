from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REQUIRED_COLUMNS = ["Date", "Product", "Category", "Quantity", "Unit Price"]


def parse_date_input(date_str: str, label: str) -> pd.Timestamp | None:
    """Parse an optional user date string (e.g. '01-01-2026' or '2026-01-31')."""
    cleaned = date_str.strip()
    if not cleaned:
        return None
    parsed = pd.to_datetime(cleaned, format="mixed", dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        print(f"Warning: Could not parse {label} '{cleaned}'. Ignoring that bound.")
        return None
    return parsed


def load_and_combine_excel_files(folder_path: Path) -> tuple[pd.DataFrame, list[str]]:
    """
    Read all .xlsx / .xls files in folder_path (skipping temporary ~$ files)
    and combine them into a single DataFrame.
    """
    excel_files = sorted(
        [
            f
            for f in folder_path.iterdir()
            if f.is_file()
            and f.suffix.lower() in {".xlsx", ".xls"}
            and not f.name.startswith("~$")
        ]
    )

    frames: list[pd.DataFrame] = []
    skipped_files: list[str] = []

    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, engine="openpyxl" if file_path.suffix.lower() == ".xlsx" else None)
            df.columns = [str(col).strip() for col in df.columns]

            missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
            if missing_cols:
                skipped_files.append(f"{file_path.name} (missing columns: {', '.join(missing_cols)})")
                continue

            df = df[REQUIRED_COLUMNS].copy()
            df["Source File"] = file_path.name
            frames.append(df)
        except Exception as exc:
            skipped_files.append(f"{file_path.name} ({exc})")

    if not frames:
        return pd.DataFrame(columns=REQUIRED_COLUMNS + ["Source File"]), skipped_files

    combined = pd.concat(frames, ignore_index=True)
    return combined, skipped_files


def clean_and_prepare_data(
    raw_df: pd.DataFrame,
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Clean text, coerce dates and numbers, filter by optional date range,
    calculate Total Sales, and separate valid rows from invalid rows.
    """
    if raw_df.empty:
        empty_valid = pd.DataFrame(
            columns=["Date", "Product", "Category", "Quantity", "Unit Price", "Total Sales", "Source File"]
        )
        empty_invalid = pd.DataFrame(columns=list(raw_df.columns) + ["Issue"])
        return empty_valid, empty_invalid

    df = raw_df.copy()

    # Clean text fields
    for col in ["Product", "Category"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
        df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    df["Product"] = df["Product"].str.title()
    df["Category"] = df["Category"].str.title()

    # Parse dates and numeric columns safely
    parsed_dates = pd.to_datetime(df["Date"], format="mixed", dayfirst=True, errors="coerce")
    qty_num = pd.to_numeric(df["Quantity"], errors="coerce")
    price_num = pd.to_numeric(df["Unit Price"], errors="coerce")

    valid_mask = pd.Series(True, index=df.index)
    issues: list[list[str]] = [[] for _ in range(len(df))]

    for idx in df.index:
        row_issues: list[str] = []

        if pd.isna(parsed_dates.loc[idx]):
            row_issues.append("Invalid or missing Date")
        else:
            dt = parsed_dates.loc[idx]
            if start_date is not None and dt < start_date:
                row_issues.append(f"Date before start ({start_date.strftime('%Y-%m-%d')})")
            if end_date is not None and dt > end_date:
                row_issues.append(f"Date after end ({end_date.strftime('%Y-%m-%d')})")

        if pd.isna(df.loc[idx, "Product"]):
            row_issues.append("Missing Product")

        if pd.isna(df.loc[idx, "Category"]):
            row_issues.append("Missing Category")

        if pd.isna(qty_num.loc[idx]) or qty_num.loc[idx] <= 0:
            row_issues.append("Invalid or non-positive Quantity")

        if pd.isna(price_num.loc[idx]) or price_num.loc[idx] < 0:
            row_issues.append("Invalid or negative Unit Price")

        if row_issues:
            valid_mask.loc[idx] = False
            issues[idx] = row_issues

    invalid_df = df[~valid_mask].copy()
    invalid_df["Issue"] = ["; ".join(issues[i]) for i, keep in enumerate(valid_mask) if not keep]

    valid_df = df[valid_mask].copy()
    valid_df["Date"] = parsed_dates[valid_mask].dt.strftime("%Y-%m-%d")
    valid_df["Quantity"] = qty_num[valid_mask].astype(int)
    valid_df["Unit Price"] = price_num[valid_mask].astype(float)
    valid_df["Total Sales"] = valid_df["Quantity"] * valid_df["Unit Price"]

    ordered_cols = ["Date", "Product", "Category", "Quantity", "Unit Price", "Total Sales", "Source File"]
    valid_df = valid_df[ordered_cols].sort_values(by=["Date", "Product"]).reset_index(drop=True)

    return valid_df, invalid_df


def build_summaries(valid_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Build Product Summary, Category Summary, and Overall Summary DataFrames."""
    if valid_df.empty:
        product_summary = pd.DataFrame(
            columns=["Product", "Category", "Total Quantity Sold", "Total Revenue", "Transactions"]
        )
        category_summary = pd.DataFrame(
            columns=["Category", "Total Quantity Sold", "Total Revenue", "Transactions"]
        )
        overall_summary = pd.DataFrame(
            [
                {"Metric": "Total Transactions", "Value": 0},
                {"Metric": "Total Quantity Sold", "Value": 0},
                {"Metric": "Total Revenue", "Value": 0.0},
                {"Metric": "Best-Selling Product (by Quantity)", "Value": "N/A"},
                {"Metric": "Top Revenue Category", "Value": "N/A"},
            ]
        )
        return product_summary, category_summary, overall_summary

    product_summary = (
        valid_df.groupby(["Product", "Category"], as_index=False)
        .agg(
            **{
                "Total Quantity Sold": ("Quantity", "sum"),
                "Total Revenue": ("Total Sales", "sum"),
                "Transactions": ("Total Sales", "count"),
            }
        )
        .sort_values(by=["Total Quantity Sold", "Total Revenue"], ascending=[False, False])
        .reset_index(drop=True)
    )

    category_summary = (
        valid_df.groupby("Category", as_index=False)
        .agg(
            **{
                "Total Quantity Sold": ("Quantity", "sum"),
                "Total Revenue": ("Total Sales", "sum"),
                "Transactions": ("Total Sales", "count"),
            }
        )
        .sort_values(by="Total Revenue", ascending=False)
        .reset_index(drop=True)
    )

    best_product_row = product_summary.iloc[0]
    best_product_text = f"{best_product_row['Product']} ({int(best_product_row['Total Quantity Sold'])} units)"

    top_category_row = category_summary.iloc[0]
    top_category_text = f"{top_category_row['Category']} ({top_category_row['Total Revenue']:,.2f})"

    overall_summary = pd.DataFrame(
        [
            {"Metric": "Total Transactions", "Value": int(len(valid_df))},
            {"Metric": "Total Quantity Sold", "Value": int(valid_df["Quantity"].sum())},
            {"Metric": "Total Revenue", "Value": float(valid_df["Total Sales"].sum())},
            {"Metric": "Best-Selling Product (by Quantity)", "Value": best_product_text},
            {"Metric": "Top Revenue Category", "Value": top_category_text},
        ]
    )

    return product_summary, category_summary, overall_summary


def style_workbook(output_path: Path) -> None:
    """Apply clean header styling and auto-fit column widths using openpyxl."""
    wb = load_workbook(output_path)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                text_len = len(str(val)) if val is not None else 0
                if text_len > max_len:
                    max_len = text_len
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 45)

    wb.save(output_path)


def generate_consolidated_report(
    folder_str: str,
    start_date_str: str = "",
    end_date_str: str = "",
    output_filename: str = "consolidated_sales_report.xlsx",
) -> None:
    """Orchestrate reading, cleaning, summarizing, and exporting the Excel report."""
    folder_path = Path(folder_str.strip().strip('"').strip("'")).expanduser()
    if not folder_path.exists() or not folder_path.is_dir():
        print(f"Error: '{folder_path}' is not a valid folder.")
        return

    start_date = parse_date_input(start_date_str, "start date")
    end_date = parse_date_input(end_date_str, "end date")
    if start_date and end_date and start_date > end_date:
        print("Error: Start date cannot be after end date.")
        return

    raw_df, skipped_files = load_and_combine_excel_files(folder_path)
    if skipped_files:
        print("\n--- Skipped Files ---")
        for info in skipped_files:
            print(f"- {info}")

    valid_df, invalid_df = clean_and_prepare_data(raw_df, start_date, end_date)
    product_summary, category_summary, overall_summary = build_summaries(valid_df)

    output_path = folder_path / output_filename
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        valid_df.to_excel(writer, sheet_name="All Transactions", index=False)
        product_summary.to_excel(writer, sheet_name="Product Summary", index=False)
        category_summary.to_excel(writer, sheet_name="Category Summary", index=False)
        overall_summary.to_excel(writer, sheet_name="Overall Summary", index=False)

    style_workbook(output_path)

    print("\n--- Report Generated Successfully ---")
    print(f"Output file        : {output_path}")
    print(f"Valid transactions : {len(valid_df)}")
    print(f"Invalid/filtered   : {len(invalid_df)}")
    if not product_summary.empty:
        top = product_summary.iloc[0]
        print(f"Best-selling item  : {top['Product']} ({int(top['Total Quantity Sold'])} units)")


if __name__ == "__main__":
    sales_folder = input("Enter path to monthly sales folder (e.g. ./sales): ")
    start_dt = input("Optional start date (DD-MM-YYYY, press Enter for none): ")
    end_dt = input("Optional end date (DD-MM-YYYY, press Enter for none): ")
    generate_consolidated_report(sales_folder, start_dt, end_dt)