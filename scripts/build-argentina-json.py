
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path.cwd()
SHEET_NAME = "local"
OUTPUT_FILE = ROOT / "argentina.json"

REQUIRED_COLUMNS = {
    "código de jurisdicción": "jurisdiction_code",
    "jurisdicción": "jurisdiction",
    "código de departamento": "department_code",
    "departamento": "department",
    "código de aglomerado": "agglomeration_code",
    "aglomerado": "agglomeration",
    "código de localidad": "locality_id",
    "localidad": "locality",
    "tipo de localidad": "locality_type",
    "código de gobierno local": "local_government_code",
    "gobierno local": "local_government",
    "categoría de gobierno local": "local_government_category",
}


def normalize_header(value):
    if value is None:
        return ""
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def cell_value(cell, preserve_code=False):
    value = cell.value

    if value is None:
        return None

    if isinstance(value, str):
        value = value.strip()
        return value if value else None

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            return None

        if isinstance(value, float) and value.is_integer():
            value = int(value)

        if preserve_code and isinstance(value, int):
            number_format = cell.number_format or ""

            # Preserve leading zeros when the Excel cell uses a zero mask.
            if re.fullmatch(r"0+", number_format):
                return str(value).zfill(len(number_format))

        return str(value)

    return str(value).strip()


def find_workbook():
    candidates = []

    for path in ROOT.rglob("*.xlsx"):
        if any(part.startswith(".") for part in path.relative_to(ROOT).parts):
            continue
        if path.name.startswith("~$"):
            continue

        try:
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=True,
            )
            has_sheet = SHEET_NAME in workbook.sheetnames
            workbook.close()

            if has_sheet:
                candidates.append(path)
        except Exception as exc:
            print(f"Skipping unreadable workbook {path}: {exc}")

    if not candidates:
        raise FileNotFoundError(
            f"No Excel workbook containing worksheet '{SHEET_NAME}' was found."
        )

    if len(candidates) > 1:
        names = "\n".join(str(path) for path in candidates)
        raise RuntimeError(
            "Multiple workbooks contain the worksheet 'local'. "
            "Keep only the intended workbook or configure its path:\n" + names
        )

    return candidates[0]


def main():
    workbook_path = find_workbook()
    print(f"Source workbook: {workbook_path}")

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:
        worksheet = workbook[SHEET_NAME]
        rows = worksheet.iter_rows()
        header_cells = next(rows, None)

        if not header_cells:
            raise ValueError("The worksheet is empty.")

        headers = {}
        for index, cell in enumerate(header_cells):
            normalized = normalize_header(cell.value)
            if normalized:
                headers[normalized] = index

        missing = [
            name for name in REQUIRED_COLUMNS
            if name not in headers
        ]

        if missing:
            raise ValueError(
                "Missing required Excel columns:\n- "
                + "\n- ".join(missing)
                + "\nFound columns: "
                + ", ".join(headers.keys())
            )

        records = []
        seen_ids = set()
        skipped_rows = 0

        for row_number, cells in enumerate(rows, start=2):
            def get_value(source_header):
                index = headers[source_header]
                if index >= len(cells):
                    return None

                return cell_value(
                    cells[index],
                    preserve_code=(
                        "código" in source_header
                    ),
                )

            locality_id = get_value("código de localidad")
            locality_name = get_value("localidad")

            # Ignore blank rows and source/update-note footer rows.
            if not locality_id or not locality_name:
                skipped_rows += 1
                continue

            record = {}

            for source_header, output_key in REQUIRED_COLUMNS.items():
                value = get_value(source_header)

                if output_key == "locality_id":
                    value = locality_id
                elif output_key == "locality":
                    value = locality_name

                record[output_key] = value

            # Preserve the locality ID as a string.
            record["locality_id"] = str(locality_id)

            if record["locality_id"] in seen_ids:
                raise ValueError(
                    f"Duplicate locality ID '{record['locality_id']}' "
                    f"found at Excel row {row_number}. "
                    "Resolve duplicate IDs before generating the JSON."
                )

            seen_ids.add(record["locality_id"])
            records.append(record)

    finally:
        workbook.close()

    if not records:
        raise ValueError("No valid locality records were found.")

    # Stable ordering for reproducible builds.
    records.sort(
        key=lambda item: (
            item["jurisdiction_code"] or "",
            item["department_code"] or "",
            item["locality_id"],
        )
    )

    temporary_file = OUTPUT_FILE.with_suffix(".json.tmp")

    with temporary_file.open("w", encoding="utf-8", newline="\n") as file:
        json.dump(
            records,
            file,
            ensure_ascii=False,
            indent=2,
        )
        file.write("\n")

    temporary_file.replace(OUTPUT_FILE)

    print(f"Valid localities: {len(records)}")
    print(f"Skipped blank/footer rows: {skipped_rows}")
    print(f"Unique locality IDs: {len(seen_ids)}")
    print(f"Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)
