from __future__ import annotations

import argparse
import html
import json
import logging
import re
import unicodedata
from datetime import date, datetime, timedelta
from numbers import Number
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


PROJECT = Path(__file__).resolve().parent
DEFAULT_CONFIG = PROJECT / "config.json"
LOG_PATH = PROJECT / "automation.log"


def normalize(value) -> str:
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^A-Z0-9]", "", text.upper())


def parse_excel_dates(values: pd.Series) -> pd.Series:
    """Read Excel serial dates, Python dates, and common date strings."""
    parsed = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")
    numeric = values.map(lambda value: isinstance(value, Number) and not isinstance(value, bool))

    if numeric.any():
        parsed.loc[numeric] = pd.to_datetime(
            pd.to_numeric(values.loc[numeric], errors="coerce"),
            unit="D", origin="1899-12-30", errors="coerce",
        )

    other = values.notna() & ~numeric
    if other.any():
        text = values.loc[other].astype(str).str.strip()
        serial = text.str.fullmatch(r"\d{5}(?:\.\d+)?")
        if serial.any():
            indexes = text.index[serial]
            parsed.loc[indexes] = pd.to_datetime(
                pd.to_numeric(text.loc[indexes], errors="coerce"),
                unit="D", origin="1899-12-30", errors="coerce",
            )
        date_text = text.loc[~serial]
        if not date_text.empty:
            parsed.loc[date_text.index] = pd.to_datetime(
                date_text, format="mixed", dayfirst=True, errors="coerce"
            )
    return parsed.dt.normalize()


def resolve_relative(path_value: str, base: Path = PROJECT) -> Path:
    path = Path(path_value).expanduser()
    return path if path.is_absolute() else (base / path).resolve()


def eligible_dates(as_of: date, ages: list[int]) -> set[date]:
    # Inclusive age: record date is day 1, so age 6 means as_of minus 5 days.
    return {as_of - timedelta(days=age - 1) for age in ages}


def resolve_sheet_name(sheet_names, requested: str) -> str:
    by_normalized = {normalize(name): name for name in sheet_names}
    actual = by_normalized.get(normalize(requested))
    if not actual:
        raise ValueError(f"Sheet '{requested}' not found. Available: {sheet_names}")
    return actual


def read_source(config: dict, as_of: date):
    master = resolve_relative(config["master_file"])
    if not master.is_file():
        raise FileNotFoundError(f"Master workbook not found: {master}")

    ages = config.get("age_days_inclusive", [6, 7])
    dates = pd.to_datetime(sorted(eligible_dates(as_of, ages)))
    category_data = {}
    category_headers = {}

    with pd.ExcelFile(master, engine="openpyxl") as book:
        for category in config["categories"]:
            sheet = resolve_sheet_name(book.sheet_names, category["sheet"])
            frame = pd.read_excel(book, sheet_name=sheet, dtype=object)
            frame.columns = [str(column).strip() for column in frame.columns]

            store_column = category["store_column"]
            date_column = category["date_column"]
            for required in (store_column, date_column):
                if required not in frame.columns:
                    raise ValueError(
                        f"Column '{required}' is missing from sheet '{sheet}'. "
                        f"Found: {list(frame.columns)}"
                    )

            frame = frame.dropna(how="all").copy()
            frame["__STORE_KEY"] = frame[store_column].map(normalize)
            frame["__DATE_NORMALIZED"] = parse_excel_dates(frame[date_column])
            frame = frame[frame["__STORE_KEY"] != ""]
            frame = frame[frame["__DATE_NORMALIZED"].isin(dates)]

            type_column = category.get("type_column")
            type_values = category.get("type_values", [])
            if type_column and type_values:
                if type_column not in frame.columns:
                    raise ValueError(
                        f"Type column '{type_column}' is missing from sheet '{sheet}'."
                    )
                allowed = {normalize(value) for value in type_values}
                frame = frame[frame[type_column].map(normalize).isin(allowed)]

            category_data[category["key"]] = frame
            category_headers[category["key"]] = [
                column for column in frame.columns if not column.startswith("__")
            ]

    return category_data, category_headers


def read_contacts(config: dict) -> dict:
    path = resolve_relative(config["contacts_file"])
    if not path.is_file():
        raise FileNotFoundError(f"Contacts workbook not found: {path}")
    contact_cfg = config["contacts"]
    frame = pd.read_excel(
        path,
        sheet_name=contact_cfg.get("sheet_name") or 0,
        dtype=object,
        engine="openpyxl",
    )
    frame.columns = [str(column).strip() for column in frame.columns]
    store_col = contact_cfg["store_column"]
    email_col = contact_cfg["email_column"]
    manager_col = contact_cfg.get("manager_column")
    for col in (store_col, email_col):
        if col not in frame.columns:
            raise ValueError(f"Contacts column '{col}' not found. Found: {list(frame.columns)}")

    contacts = {}
    for _, row in frame.iterrows():
        key = normalize(row[store_col])
        if key:
            contacts[key] = {
                "store_name": str(row[store_col]).strip(),
                "email": str(row[email_col]).strip() if pd.notna(row[email_col]) else "",
                "manager": (
                    str(row[manager_col]).strip()
                    if manager_col in frame.columns and pd.notna(row[manager_col])
                    else ""
                ),
            }
    return contacts


def safe_filename(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", value).strip(" .")


def store_output_path(store_root: Path, store_name: str, category: dict) -> Path:
    subfolder = category.get("output_subfolder", category["key"])
    filename = category.get("output_filename", "{store}_{category}.xlsx").format(
        store=safe_filename(store_name), category=safe_filename(category["key"])
    )
    return store_root / safe_filename(store_name) / subfolder / filename


def validate_template(template_path: Path, expected_headers: list[str]) -> None:
    if not template_path.is_file():
        raise FileNotFoundError(f"Workbook template not found: {template_path}")
    workbook = load_workbook(template_path, read_only=True, data_only=False)
    try:
        sheet = workbook.active
        actual = [cell.value for cell in sheet[1]]
        while actual and actual[-1] is None:
            actual.pop()
        if [normalize(value) for value in actual] != [normalize(value) for value in expected_headers]:
            raise ValueError(
                f"Header mismatch in {template_path.name}. "
                f"Template={actual}; source={expected_headers}"
            )
    finally:
        workbook.close()


def write_store_workbook(output_path: Path, template_path: Path, headers: list[str], rows):
    workbook = load_workbook(template_path)
    sheet = workbook.active
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)
    for _, row in rows.iterrows():
        values = []
        for header in headers:
            value = row[header]
            if pd.isna(value):
                value = None
            elif hasattr(value, "item"):
                value = value.item()
            values.append(value)
        sheet.append(values)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()


def build_html_table(store_name: str, counts: dict[str, int], categories: list[dict]) -> str:
    heading_cells = "".join(f"<th>{html.escape(item['label'])}</th>" for item in categories)
    count_cells = "".join(
        f"<td style='text-align:center'>{counts.get(item['key'], 0)}</td>"
        for item in categories
    )
    return (
        "<table border='1' cellpadding='6' cellspacing='0' "
        "style='border-collapse:collapse;font-family:Arial,sans-serif'>"
        f"<thead><tr><th>Store</th>{heading_cells}</tr></thead>"
        f"<tbody><tr><td>{html.escape(store_name)}</td>{count_cells}</tr></tbody></table>"
    )


def run(config_path: Path, dry_run: bool, as_of: date):
    config = json.loads(config_path.read_text(encoding="utf-8"))
    logging.basicConfig(
        filename=LOG_PATH,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )

    data, headers = read_source(config, as_of)
    contacts = read_contacts(config)
    store_root = resolve_relative(config["stores_root"])
    categories = config["categories"]
    template_by_key = {
        item["key"]: resolve_relative(item["workbook_template"])
        for item in categories
    }
    for category in categories:
        validate_template(template_by_key[category["key"]], headers[category["key"]])

    plan = []
    for store_key, contact in contacts.items():
        store_name = contact["store_name"]
        store_folder = store_root / safe_filename(store_name)
        rows_by_category = {
            item["key"]: data[item["key"]][
                data[item["key"]]["__STORE_KEY"] == store_key
            ].copy()
            for item in categories
        }
        counts = {key: len(rows) for key, rows in rows_by_category.items()}
        plan.append((store_name, store_folder, contact, rows_by_category, counts))

    if dry_run:
        dates = ", ".join(
            item.strftime("%d/%m/%Y")
            for item in sorted(eligible_dates(as_of, config.get("age_days_inclusive", [6, 7])))
        )
        print(f"As-of date: {as_of:%d/%m/%Y}; eligible dates: {dates}")
        for store_name, _, _, _, counts in plan:
            print(f"{store_name}: total={sum(counts.values())}; {counts}")
        print("Dry run only: no workbooks changed; Outlook not opened.")
        return

    email_template_path = resolve_relative(config["email_template"])
    email_template = email_template_path.read_text(encoding="utf-8")
    if "{{TABLE}}" not in email_template:
        raise ValueError("Email HTML template must include {{TABLE}}.")

    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Install pywin32 on Windows to prepare Outlook drafts.") from exc
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Write only the eligible records and retain each category's formatted header template.
    for store_name, store_folder, contact, rows_by_category, counts in plan:
        for category in categories:
            key = category["key"]
            output_path = store_output_path(store_root, store_name, category)
            rows = rows_by_category[key].drop(
                columns=["__STORE_KEY", "__DATE_NORMALIZED"], errors="ignore"
            )
            write_store_workbook(output_path, template_by_key[key], headers[key], rows)

        total = sum(counts.values())
        if total == 0:
            logging.info("No eligible records; no email draft: %s", store_name)
            continue
        if not contact["email"]:
            logging.error("Missing email address; drafts skipped: %s", store_name)
            continue

        body = email_template.replace("{{TABLE}}", build_html_table(store_name, counts, categories))
        body = body.replace("{{STORE_NAME}}", html.escape(store_name))
        body = body.replace("{{MANAGER}}", html.escape(contact["manager"]))
        mail = outlook.CreateItem(0)
        mail.To = contact["email"]
        mail.Subject = config.get("email_subject", "Pending items — {store}").format(
            store=store_name
        )
        sender = config.get("sender_on_behalf", "").strip()
        if sender:
            mail.SentOnBehalfOfName = sender
        mail.HTMLBody = body
        for category in categories:
            mail.Attachments.Add(str(store_output_path(store_root, store_name, category)))
        mail.Display()
        logging.info("Draft displayed for review: %s", store_name)

    print("Preparation complete. Outlook drafts were displayed; this script never sends email.")


def main():
    parser = argparse.ArgumentParser(description="Prepare category workbooks and Outlook drafts.")
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--dry-run", action="store_true", help="Preview eligible stores and counts only.")
    modes.add_argument("--prepare", action="store_true", help="Write workbooks and display Outlook drafts.")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument(
        "--as-of",
        type=lambda value: datetime.strptime(value, "%Y-%m-%d").date(),
        default=date.today(),
        help="Reference date for age filtering (YYYY-MM-DD; defaults to today).",
    )
    args = parser.parse_args()
    run(args.config.resolve(), args.dry_run, args.as_of)


if __name__ == "__main__":
    main()
