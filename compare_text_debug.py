import pandas as pd


# ============================
# CONFIG - UPDATE FILE PATHS
# ============================

XLSX_FILE = "PATH_TO_YOUR_XLSX_FILE"
CSV_FILE = "PATH_TO_YOUR_CSV_FILE"

TICKETS_TO_CHECK = [
    "INC37879772",
    "INC37879243",
    "INC37878403",
    "INC37877563",
    "UR1407327"
]


# ============================
# LOAD FILES
# ============================

xlsx_df = pd.read_excel(XLSX_FILE, dtype=str)
csv_df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")


# ============================
# HELPER FUNCTIONS
# ============================

def get_row(df, ticket):
    row = df[df["Number"] == ticket]
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

    # show repr
    print("\nXLSX repr:")
    print(repr(s1[:300]))

    print("\nCSV repr:")
    print(repr(s2[:300]))

    # find first difference
    min_len = min(len(s1), len(s2))

    for i in range(min_len):
        if s1[i] != s2[i]:
            print(f"\n🔍 First difference at position {i}:")
            print(f"XLSX char: {repr(s1[i])} (ord={ord(s1[i])})")
            print(f"CSV  char: {repr(s2[i])} (ord={ord(s2[i])})")
            break


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

    compare_strings("Short Description",
                    x_row.get("Short description"),
                    c_row.get("Short description"))

    compare_strings("Description",
                    x_row.get("Description"),
                    c_row.get("Description"))