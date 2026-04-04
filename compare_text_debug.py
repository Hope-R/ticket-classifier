import pandas as pd


# ============================
# CONFIG - UPDATE FILE PATHS
# ============================

XLSX_FILE = r"PATH_TO_YOUR_XLSX_FILE"
CSV_FILE = r"PATH_TO_YOUR_CSV_FILE"

TICKETS_TO_CHECK = [
    "INC37879772",
    "INC37879243",
    "INC37878403",
    "INC37877563",
]


# ============================
# LOAD FILES
# ============================

xlsx_df = pd.read_excel(XLSX_FILE, dtype=str)
csv_df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")


# ============================
# HELPER FUNCTIONS
# ============================

def find_number_column(df):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if col_clean in ["number", "incident number", "ticket number"]:
            return col
    return None


def find_text_column(df, candidates):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if col_clean in [c.lower() for c in candidates]:
            return col
    return None


def get_row(df, ticket):
    number_col = find_number_column(df)

    if number_col is None:
        print("❌ Could not find ticket number column. Available columns:")
        print(df.columns.tolist())
        return None

    row = df[df[number_col].astype(str).str.strip() == ticket]

    if row.empty:
        return None

    return row.iloc[0]


def compare_strings(label, s1, s2):
    s1 = "" if pd.isna(s1) else str(s1)
    s2 = "" if pd.isna(s2) else str(s2)

    print(f"\n--- {label} ---")
    print("Length XLSX:", len(s1))
    print("Length CSV :", len(s2))

    if s1 == s2:
        print("✅ EXACT MATCH")
        return

    print("❌ DIFFERENT")

    print("\nXLSX repr:")
    print(repr(s1[:500]))

    print("\nCSV repr:")
    print(repr(s2[:500]))

    min_len = min(len(s1), len(s2))
    found_diff = False

    for i in range(min_len):
        if s1[i] != s2[i]:
            print(f"\n🔍 First difference at position {i}:")
            print(f"XLSX char: {repr(s1[i])} (ord={ord(s1[i])})")
            print(f"CSV  char: {repr(s2[i])} (ord={ord(s2[i])})")

            start = max(0, i - 40)
            end = min(min_len, i + 80)

            print("\nXLSX context:")
            print(repr(s1[start:end]))

            print("\nCSV context:")
            print(repr(s2[start:end]))

            found_diff = True
            break

    if not found_diff:
        if len(s1) != len(s2):
            print("\n🔍 No differing character found in shared length, but lengths differ.")
            print(f"Extra XLSX tail: {repr(s1[min_len:min_len+200])}")
            print(f"Extra CSV tail : {repr(s2[min_len:min_len+200])}")


# ============================
# IDENTIFY TEXT COLUMNS
# ============================

short_desc_col_xlsx = find_text_column(
    xlsx_df,
    ["Short description", "Short Description", "short_description"]
)
short_desc_col_csv = find_text_column(
    csv_df,
    ["Short description", "Short Description", "short_description"]
)

description_col_xlsx = find_text_column(
    xlsx_df,
    ["Description", "description"]
)
description_col_csv = find_text_column(
    csv_df,
    ["Description", "description"]
)

print("Detected columns:")
print("XLSX number column      :", find_number_column(xlsx_df))
print("CSV number column       :", find_number_column(csv_df))
print("XLSX short desc column  :", short_desc_col_xlsx)
print("CSV short desc column   :", short_desc_col_csv)
print("XLSX description column :", description_col_xlsx)
print("CSV description column  :", description_col_csv)


# ============================
# MAIN LOOP
# ============================

for ticket in TICKETS_TO_CHECK:
    print("\n======================================")
    print(f"Ticket: {ticket}")
    print("======================================")

    x_row = get_row(xlsx_df, ticket)
    c_row = get_row(csv_df, ticket)

    if x_row is None or c_row is None:
        print("⚠️ Ticket missing in one of the files")
        continue

    x_short = x_row.get(short_desc_col_xlsx, "") if short_desc_col_xlsx else ""
    c_short = c_row.get(short_desc_col_csv, "") if short_desc_col_csv else ""

    x_desc = x_row.get(description_col_xlsx, "") if description_col_xlsx else ""
    c_desc = c_row.get(description_col_csv, "") if description_col_csv else ""

    compare_strings("Short Description", x_short, c_short)
    compare_strings("Description", x_desc, c_desc)