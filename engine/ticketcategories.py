import os
import re
import pandas as pd


def process_tickets(run_month):

    # ==================================================
    # 1️⃣ BASE PATHS
    # ==================================================

    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CONFIG_DIR = os.path.join(BASE_DIR, "config")
    RAW_DIR = os.path.join(BASE_DIR, "raw_data")
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")

    MONTH_RAW_DIR = os.path.join(RAW_DIR, run_month)

    # ==================================================
    # 2️⃣ VALIDATE INPUT MONTH FOLDER
    # ==================================================

    if not os.path.exists(MONTH_RAW_DIR):
        print(f"❌ Month folder not found: {MONTH_RAW_DIR}")
        return None

    def resolve_input_file(folder_path, base_name):
        """
        Allow either .xlsx or .csv for raw input files.
        Fail if none found or more than one found.
        """
        candidates = []

        xlsx_file = os.path.join(folder_path, f"{base_name}.xlsx")
        csv_file = os.path.join(folder_path, f"{base_name}.csv")

        if os.path.exists(xlsx_file):
            candidates.append(xlsx_file)

        if os.path.exists(csv_file):
            candidates.append(csv_file)

        if len(candidates) == 0:
            print(f"❌ Missing file for '{base_name}' in {folder_path}.")
            print(f"   Expected one of: {base_name}.xlsx or {base_name}.csv")
            return None

        if len(candidates) > 1:
            print(f"❌ Duplicate input files found for '{base_name}' in {folder_path}.")
            print("   Please keep only one of the following:")
            for c in candidates:
                print(f"   - {os.path.basename(c)}")
            return None

        return candidates[0]

    incident_file = resolve_input_file(MONTH_RAW_DIR, "incident")
    ur_file = resolve_input_file(MONTH_RAW_DIR, "ur")
    task_file = resolve_input_file(MONTH_RAW_DIR, "task")

    if not all([incident_file, ur_file, task_file]):
        return None

    # ==================================================
    # 2️⃣A CSV WARNING (WARNING-ONLY, NO PROMPT)
    # ==================================================

    if any(f.lower().endswith(".csv") for f in [incident_file, ur_file, task_file]):
        print("\n⚠️ Input format guidance:")
        print(".xlsx is the preferred raw input format.")
        print(".csv is supported, but depending on the source export, it may introduce formatting differences.")
        print("For the most consistent results, use .xlsx whenever available.\n")

    # ==================================================
    # 3️⃣ LOAD CONFIG FILES
    # ==================================================

    RULES_FILE = os.path.join(CONFIG_DIR, "business_rules.xlsx")

    assignment_groups_df = pd.read_excel(
        RULES_FILE,
        sheet_name="assignment_groups"
    )

    keyword_rules_df = pd.read_excel(
        RULES_FILE,
        sheet_name="keyword_rules"
    )

    phrase_rules_df = pd.read_excel(
        RULES_FILE,
        sheet_name="phrase_rules"
    )

    template_noise_df = pd.read_excel(
        RULES_FILE,
        sheet_name="template_noise"
    )

    category_rules_df = pd.read_excel(
        RULES_FILE,
        sheet_name="category_check"
    )

    service_rules_df = pd.read_excel(
        RULES_FILE,
        sheet_name="service_check"
    )

    column_mapping_df = pd.read_excel(
        RULES_FILE,
        sheet_name="column_mapping"
    )

    final_category_consolidation_df = pd.read_excel(
        RULES_FILE,
        sheet_name="final_category_consolidation"
    )

    # ==================================================
    # 4️⃣ COLUMN STANDARDIZATION
    # ==================================================

    column_mapping_df.columns = column_mapping_df.columns.str.strip()

    column_mapping_df["Standard Column"] = (
        column_mapping_df["Standard Column"].astype(str).str.strip()
    )

    column_mapping_df["Possible Column Name"] = (
        column_mapping_df["Possible Column Name"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    column_aliases = (
        column_mapping_df
        .groupby("Standard Column")["Possible Column Name"]
        .apply(list)
        .to_dict()
    )

    required_columns = [
        "Number", "Caller", "Email", "Contact type", "Opened",
        "Short description", "Description", "Assignment group",
        "Priority", "Category", "Service", "Resolved", "Close notes"
    ]

    def standardize_columns(df):
        df.columns = df.columns.str.strip()
        lower_cols = {col.lower(): col for col in df.columns}

        rename_dict = {}

        for standard_col, aliases in column_aliases.items():
            for alias in aliases:
                if alias in lower_cols:
                    rename_dict[lower_cols[alias]] = standard_col

        df = df.rename(columns=rename_dict)

        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        return df.loc[:, ~df.columns.duplicated()]

    # ==================================================
    # 5️⃣ LOAD RAW FILES (CANONICAL INGESTION)
    # ==================================================

    def canonicalize_scalar(value):
        """
        Canonical standard:
        - Treat blank-like values consistently
        - Preserve text content
        - Strip leading/trailing spaces
        """
        if pd.isna(value):
            return ""

        if isinstance(value, str):
            value = value.strip()

            if value.lower() in {"nan", "none", "null"}:
                return ""

            return value

        return value

    def load_raw_file(file_path):
        """
        Canonical ingestion layer:
        - Read .xlsx and .csv in a controlled way
        - Avoid pandas guessing types differently across formats
        - Standardize blanks/strings before business logic runs
        """
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".xlsx":
            df = pd.read_excel(
                file_path,
                dtype=str,
                keep_default_na=False
            )
        elif ext == ".csv":
            df = pd.read_csv(
                file_path,
                dtype=str,
                keep_default_na=False
            )
        else:
            raise ValueError(f"Unsupported file format: {file_path}")

        df = df.apply(lambda col: col.map(canonicalize_scalar))

        return df

    incident_df = standardize_columns(load_raw_file(incident_file))
    ur_df = standardize_columns(load_raw_file(ur_file))
    task_df = standardize_columns(load_raw_file(task_file))

    # ==================================================
    # 6️⃣ BUSINESS FILTERS
    # ==================================================

    incident_df = incident_df[
        incident_df["Contact type"].astype(str).str.strip() != "Alert"
    ]

    ur_df = ur_df[
        (
            ur_df["Primary Ticket"].isna() |
            (ur_df["Primary Ticket"].astype(str).str.strip() == "")
        )
        &
        (
            ur_df["Category"].notna() &
            (ur_df["Category"].astype(str).str.strip() != "")
        )
    ]

    valid_groups = (
        assignment_groups_df.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    task_df = task_df[
        task_df["Assignment group"].astype(str).str.strip().isin(valid_groups)
    ]

    # ==================================================
    # 7️⃣ ADD METADATA
    # ==================================================

    for df, ttype in [(incident_df, "Inc"), (ur_df, "UR"), (task_df, "Task")]:
        df["Ticket Type"] = ttype

    # ==================================================
    # 8️⃣ COMBINE
    # ==================================================

    end_user_tickets = pd.concat(
        [incident_df, ur_df, task_df],
        ignore_index=True
    )

    # ==================================================
    # 9️⃣ TEXT NORMALIZATION
    # ==================================================

    def normalize_matching_text(text):
        """
        For fields used in keyword/phrase matching.
        Lowercase + normalize whitespace.
        """
        if pd.isna(text):
            return ""

        text = str(text)
        text = text.strip()
        text = re.sub(r"\s+", " ", text)
        text = text.lower()
        return text

    def normalize_exact_match_text(text):
        """
        For fields used in exact business-rule matching.
        Preserve case, just clean whitespace.
        """
        if pd.isna(text):
            return ""

        text = str(text)
        text = text.strip()
        text = re.sub(r"\s+", " ", text)
        return text

    if "Short description" in end_user_tickets.columns:
        end_user_tickets["Short description"] = (
            end_user_tickets["Short description"].apply(normalize_matching_text)
        )

    if "Description" in end_user_tickets.columns:
        end_user_tickets["Description"] = (
            end_user_tickets["Description"].apply(normalize_matching_text)
        )

    if "Service" in end_user_tickets.columns:
        end_user_tickets["Service"] = (
            end_user_tickets["Service"].apply(normalize_exact_match_text)
        )

    if "Category" in end_user_tickets.columns:
        end_user_tickets["Category"] = (
            end_user_tickets["Category"].apply(normalize_exact_match_text)
        )

    # ==================================================
    # 🔟 DERIVE MONTH FROM OPENED COLUMN
    # ==================================================

    end_user_tickets["Opened"] = pd.to_datetime(
        end_user_tickets["Opened"],
        errors="coerce"
    )

    end_user_tickets["Month"] = end_user_tickets["Opened"].dt.strftime("%b-%y")

    month_values = (
        end_user_tickets["Month"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    if not month_values:
        print("❌ Could not derive Month from 'Opened' column. Please check input data.")
        return None

    # STRICT VALIDATION: only one month allowed in input
    if len(month_values) > 1:
        print(f"❌ Multiple months detected in input data: {', '.join(sorted(month_values))}")
        print(f"❌ Selected month folder: {run_month}")
        print("❌ Please ensure only one month of ticket data is placed in the selected folder.")
        return None

    derived_month = month_values[0]

    # CONSISTENCY CHECK: derived month must match selected folder
    if derived_month != run_month:
        print("❌ Mismatch detected.")
        print(f"❌ Selected month folder: {run_month}")
        print(f"❌ Derived month from 'Opened': {derived_month}")
        print("❌ Please check the raw files and selected month folder.")
        return None

    # ==================================================
    # 1️⃣1️⃣ PREPARE RULE LOGIC
    # ==================================================

    keyword_rules_df["Category"] = (
        keyword_rules_df["Category"].astype(str).str.strip()
    )
    keyword_rules_df["Keyword"] = (
        keyword_rules_df["Keyword"].astype(str).str.lower().str.strip()
    )

    category_keywords = {}
    category_priority = {}

    for _, row in keyword_rules_df.iterrows():
        category = row["Category"]
        keyword = row["Keyword"]
        priority = row["Priority"]

        category_keywords.setdefault(category, []).append(keyword)
        category_priority[category] = priority

    phrase_rules_df["Category"] = (
        phrase_rules_df["Category"].astype(str).str.strip()
    )
    phrase_rules_df["Phrase"] = (
        phrase_rules_df["Phrase"].astype(str).str.lower().str.strip()
    )

    category_phrases = {}

    for _, row in phrase_rules_df.iterrows():
        category = row["Category"]
        phrase = row["Phrase"]
        weight = row["Weight"]

        category_phrases.setdefault(category, []).append((phrase, weight))

    # ==================================================
    # 1️⃣2️⃣ LOAD CATEGORY & SERVICE RULES
    # ==================================================

    category_rules_df["Category"] = (
        category_rules_df["Category"].astype(str).str.lower().str.strip()
    )
    category_rules_df["Mapped Category"] = (
        category_rules_df["Mapped Category"].astype(str).str.strip()
    )

    category_rules = dict(
        zip(category_rules_df["Category"], category_rules_df["Mapped Category"])
    )

    service_rules_df["Service"] = (
        service_rules_df["Service"].astype(str).str.strip()
    )
    service_rules_df["Category"] = (
        service_rules_df["Category"].astype(str).str.strip()
    )

    service_category_map = dict(
        zip(service_rules_df["Service"], service_rules_df["Category"])
    )

    # ==================================================
    # 1️⃣3️⃣ LOAD FINAL CATEGORY CONSOLIDATION RULES
    # ==================================================

    final_category_consolidation_df.columns = (
        final_category_consolidation_df.columns.str.strip()
    )

    final_category_consolidation_df["Old Category Name"] = (
        final_category_consolidation_df["Old Category Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    final_category_consolidation_df["Final Mapped Category"] = (
        final_category_consolidation_df["Final Mapped Category"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    final_category_consolidation_df = final_category_consolidation_df[
        (final_category_consolidation_df["Old Category Name"] != "") &
        (final_category_consolidation_df["Final Mapped Category"] != "")
    ]

    final_category_map = dict(
        zip(
            final_category_consolidation_df["Old Category Name"],
            final_category_consolidation_df["Final Mapped Category"]
        )
    )

    # ==================================================
    # 1️⃣4️⃣ TEMPLATE SANITIZATION
    # ==================================================

    template_noise_df["Phrase"] = (
        template_noise_df["Phrase"]
        .astype(str)
        .str.strip()
    )

    TEMPLATE_PHRASES = (
        template_noise_df["Phrase"]
        .dropna()
        .tolist()
    )

    def clean_description(text):
        cleaned = text

        for phrase in TEMPLATE_PHRASES:
            pattern = re.compile(re.escape(phrase), re.IGNORECASE)
            cleaned = pattern.sub("", cleaned)

        return cleaned

    end_user_tickets["Description_clean"] = (
        end_user_tickets["Description"].apply(clean_description)
    )

    # ==================================================
    # 1️⃣5️⃣ WEIGHTED SCORING ENGINE
    # ==================================================

    MIN_SCORE_THRESHOLD = 1

    def determine_category(row):

        category_value = str(row.get("Category", "")).lower().strip()
        service_value = str(row.get("Service", "")).strip()

        # -----------------------------
        # 1️⃣ CATEGORY OVERRIDE
        # -----------------------------
        if category_value in category_rules:
            return category_rules[category_value]

        # -----------------------------
        # 2️⃣ SERVICE MAPPING
        # -----------------------------
        if service_value:
            if service_value in service_category_map:
                return service_category_map[service_value]
            else:
                return service_value

        short_desc = str(row.get("Short description", "")).lower().strip()
        description = str(row.get("Description_clean", "")).lower().strip()

        # -----------------------------
        # 3️⃣ MICROSOFT TEAMS RULE
        # -----------------------------
        if "Microsoft Teams" in category_keywords:
            for keyword in category_keywords["Microsoft Teams"]:
                if keyword in short_desc:
                    return "Microsoft Teams"

        combined_text = short_desc + " " + description
        scores = {}

        all_categories = sorted(
            set(category_keywords.keys()) | set(category_phrases.keys())
        )

        # -----------------------------
        # 4️⃣ WEIGHTED SCORING ENGINE
        # -----------------------------
        for category in all_categories:

            if category == "Microsoft Teams":
                continue

            phrase_score = 0
            keyword_score = 0

            for phrase, weight in category_phrases.get(category, []):
                if phrase in combined_text:
                    phrase_score += weight

            for keyword in category_keywords.get(category, []):
                pattern = r'\b' + re.escape(keyword) + r'\b'
                if re.search(pattern, combined_text):
                    keyword_score += 1

            total_score = phrase_score + keyword_score

            if total_score > 0:
                scores[category] = {
                    "total": total_score,
                    "phrase": phrase_score,
                    "priority": category_priority.get(category, 999)
                }

        if not scores:
            return "Others"

        sorted_categories = sorted(
            scores.items(),
            key=lambda x: (
                -x[1]["total"],
                -x[1]["phrase"],
                x[1]["priority"],
                x[0]
            )
        )

        best_category, best_data = sorted_categories[0]

        if best_data["total"] >= MIN_SCORE_THRESHOLD:
            return best_category
        else:
            return "Others"

    end_user_tickets["Ticket Category"] = (
        end_user_tickets.apply(determine_category, axis=1)
    )

    # ==================================================
    # 1️⃣6️⃣ FINAL CATEGORY CONSOLIDATION
    # ==================================================

    end_user_tickets["Ticket Category"] = (
        end_user_tickets["Ticket Category"]
        .astype(str)
        .str.strip()
        .replace(final_category_map)
    )

    # ==================================================
    # 1️⃣7️⃣ FINAL OUTPUT PREPARATION
    # ==================================================

    final_columns = required_columns + [
        "Ticket Type", "Month", "Ticket Category"
    ]

    end_user_tickets = end_user_tickets[final_columns].copy()

    def remove_garbage_rows(df):
        # Remove ServiceNow/Excel export artifact row
        df = df[
            ~df["Number"].astype(str).str.contains(
                "Export stopped", case=False, na=False
            )
        ].copy()

        # Remove fully blank rows
        df = df.dropna(how="all")

        # Remove structurally empty rows
        df = df[
            ~(
                df["Number"].isna() &
                df["Opened"].isna() &
                (df["Short description"].astype(str).str.strip() == "") &
                (df["Description"].astype(str).str.strip() == "")
            )
        ].copy()

        return df

    def sanitize_for_excel(df):
        """
        Make dataframe safe for Excel export:
        1. Truncate text cells to Excel's 32767-character limit
        2. Neutralize text starting with =, +, -, @ so Excel won't treat it as formula
        """
        text_columns = df.select_dtypes(include=["object"]).columns.tolist()

        for col in text_columns:
            df[col] = df[col].apply(sanitize_excel_cell)

        return df

    def sanitize_excel_cell(value):
        if pd.isna(value):
            return value

        if not isinstance(value, str):
            return value

        text = value

        # Excel cell limit
        if len(text) > 32767:
            text = text[:32767]

        # Prevent Excel from interpreting as formula
        if text.startswith(("=", "+", "-", "@")):
            text = "'" + text

        return text

    end_user_tickets = remove_garbage_rows(end_user_tickets)

    output_file = os.path.join(
        OUTPUT_DIR,
        f"end_user_ticket_data_{run_month.lower()}.xlsx"
    )

    master_file = os.path.join(
        OUTPUT_DIR,
        "master_ticket_data.xlsx"
    )

    # ==================================================
    # 1️⃣8️⃣ SAVE MONTHLY OUTPUT
    # ==================================================

    monthly_output_df = sanitize_for_excel(end_user_tickets.copy())
    monthly_output_df.to_excel(output_file, index=False)

    # ==================================================
    # 1️⃣9️⃣ REBUILD MASTER FILE FROM MONTHLY OUTPUTS
    # ==================================================

    monthly_files = []

    for file_name in os.listdir(OUTPUT_DIR):
        if (
            file_name.startswith("end_user_ticket_data_")
            and file_name.endswith(".xlsx")
        ):
            monthly_files.append(os.path.join(OUTPUT_DIR, file_name))

    monthly_files = sorted(monthly_files)

    master_parts = []

    for monthly_file in monthly_files:
        df = pd.read_excel(monthly_file)
        df = remove_garbage_rows(df)
        master_parts.append(df)

    if master_parts:
        master_df = pd.concat(master_parts, ignore_index=True)
    else:
        master_df = pd.DataFrame(columns=end_user_tickets.columns)

    # Remove duplicates across monthly files using ticket number
    if "Number" in master_df.columns:
        master_df = master_df.drop_duplicates(subset=["Number"], keep="last")

    master_df = remove_garbage_rows(master_df)

    master_output_df = sanitize_for_excel(master_df.copy())
    master_output_df.to_excel(master_file, index=False)

    print("✅ Weighted scoring processing complete.")
    print("✅ Final category consolidation applied.")
    print("✅ Month derived from 'Opened' column.")
    print("✅ Strict month validation passed.")
    print("✅ Excel-safe export sanitization applied.")
    print("✅ Raw input supports .xlsx and .csv.")
    print("✅ Canonical ingestion applied.")
    print("✅ Monthly output saved to:", output_file)
    print("✅ Master file rebuilt from monthly outputs:", master_file)

    return end_user_tickets