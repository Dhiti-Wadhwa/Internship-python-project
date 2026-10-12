from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "members.xlsx"
SCIENTISTS_FILE = BASE_DIR / "scientists.xlsx"
OUTPUT_FILE = BASE_DIR / "birthday_greetings.xlsx"


def load_excel(path):
    if not path.exists():
        raise FileNotFoundError(f"Could not find {path.name}. Place it beside main.py.")
    return pd.read_excel(path)


def clean_members(df):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    required = {"name", "email", "date_of_birth"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"members.xlsx is missing columns: {', '.join(sorted(missing))}")
    df["name"] = df["name"].astype("string").str.strip()
    df["email"] = df["email"].astype("string").str.strip()
    df["date_of_birth"] = pd.to_datetime(df["date_of_birth"], errors="coerce")
    df = df[df["name"].notna() & df["name"].ne("") & df["date_of_birth"].notna()].copy()
    df["birthday_key"] = df["date_of_birth"].dt.strftime("%m-%d")
    return df


def parse_scientist_date(value, birth_year=None):
    """Return month-day as MM-DD, supporting Excel serials and historical years."""
    if pd.isna(value):
        return None
    # Excel serial numbers, e.g. 44930, represent a date with a placeholder year.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(value))
            return dt.strftime("%m-%d")
        except (OverflowError, ValueError):
            return None
    text = str(value).strip()
    # First try common readable date formats, including historical dates like 04-Jan-1643.
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%Y-%m-%d", "%d-%b", "%m/%d/%Y", "%m/%d/%y"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%m-%d")
        except ValueError:
            pass
    # Fall back to pandas for standard modern date strings.
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%m-%d")
    except (ValueError, TypeError, OverflowError):
        pass
    return None


def clean_scientists(df):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    required = {"scientist", "date_of_birth"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"scientists.xlsx is missing columns: {', '.join(sorted(missing))}")
    df["scientist"] = df["scientist"].astype("string").str.strip()
    if "birth_year" not in df.columns:
        # Preserve the year from parseable dates where no separate column exists.
        df["birth_year"] = pd.NA
    years = pd.to_numeric(df["birth_year"], errors="coerce")
    df["birth_year"] = years.astype("Int64")
    df["birthday_key"] = [
        parse_scientist_date(value, year)
        for value, year in zip(df["date_of_birth"], df["birth_year"])
    ]
    df = df[df["scientist"].notna() & df["scientist"].ne("") & df["birthday_key"].notna()].copy()
    # Store a readable text date because Python/Pandas cannot represent some old dates
    # (for example Newton's 1643 birth date) as standard Excel/Pandas datetimes.
    readable_dates = []
    for value, year in zip(df["date_of_birth"], df["birth_year"]):
        key = parse_scientist_date(value, year)
        if key is None:
            readable_dates.append(str(value))
            continue
        month, day = map(int, key.split("-"))
        actual_year = int(year) if pd.notna(year) else 2023
        readable_dates.append(datetime(2023, month, day).strftime("%d-%b-") + str(actual_year))
    df["date_of_birth"] = readable_dates
    return df


def save_clean_scientists(scientists):
    saved = scientists[["scientist", "date_of_birth", "birth_year"]].copy()
    saved.to_excel(SCIENTISTS_FILE, index=False)
    workbook = load_workbook(SCIENTISTS_FILE)
    worksheet = workbook.active
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    worksheet.column_dimensions["A"].width = 32
    worksheet.column_dimensions["B"].width = 20
    worksheet.column_dimensions["C"].width = 14
    workbook.save(SCIENTISTS_FILE)


def build_greetings(members, scientists):
    matches = scientists.groupby("birthday_key")["scientist"].apply(
        lambda names: ", ".join(sorted(set(names.astype(str))))
    ).to_dict()
    result = members.copy()
    result["matching_scientist"] = result["birthday_key"].map(matches).fillna("")
    result["greeting"] = result.apply(
        lambda row: (
            f"Happy Birthday, {row['name']}! Did you know you share your birthday with "
            f"{row['matching_scientist']}? Wishing you a wonderful year filled with curiosity and discovery!"
            if row["matching_scientist"] else
            f"Happy Birthday, {row['name']}! Wishing you a wonderful year filled with curiosity and discovery!"
        ), axis=1
    )
    result["match_found"] = result["matching_scientist"].ne("")
    result["date_of_birth"] = result["date_of_birth"].dt.strftime("%Y-%m-%d")
    return result[["name", "email", "date_of_birth", "matching_scientist", "match_found", "greeting"]].sort_values("name")


def main():
    members = clean_members(load_excel(INPUT_FILE))
    scientists = clean_scientists(load_excel(SCIENTISTS_FILE))
    save_clean_scientists(scientists)
    output = build_greetings(members, scientists)
    output.to_excel(OUTPUT_FILE, index=False)
    print("Birthday greeting generation complete.")
    print(f"Valid member records processed: {len(output)}")
    print(f"Members matched with a scientist: {int(output['match_found'].sum())}")
    print(f"Output saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
