import os
import re
import pandas as pd


def process_tickets(run_month):

    # ==================================================
    # 0️⃣ DEBUG TICKETS TO TRACE
    # ==================================================

    DEBUG_TICKETS = {
        "INC37879772",
        "INC37879243",
        "INC37878403",
        "INC37878275",
        "INC37877856",
        "INC37877563",
        "INC37877432",
        "INC37877282",
        "INC37875074",
        "INC37874948",
        "INC37874852",
        "INC37874816",
        "INC37874299",
        "INC37873154",
        "INC37873081",
        "INC37872408",
        "INC37872391",
        "INC37872248",
        "INC37870950",
        "INC37870579",
        "INC37870545",
        "INC37870539",
        "INC37870288",
        "INC37869740",
        "INC37869602",
        "INC37868772",
        "INC37868735",
        "INC37868571",
        "INC37868333",
        "INC37867710",
        "INC37867616",
        "INC37861931",
        "INC37860178",
        "INC37859509",
        "INC37859119",
        "INC37859042",
        "INC37858940",
        "INC37858895",
        "INC37858741",
        "INC37858113",
        "INC37857995",
        "UR1400768",
        "UR1404927",
        "UR1404967",
        "UR1406682",
        "UR1407327",
        "UR1408879",
        "UR1415130",
    }

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
        .groupby("Standard Column", sort=True)["Possible Column Name"]
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

    def normalize_invisible_chars(text):
        """
        Normalize invisible / special whitespace characters that often
        differ between Excel and CSV exports.
        """
        if pd.isna(text):
            return ""

        text = str(text)

        replacements = {
            "\u00A0": " ",   # non-breaking space
            "\u2007": " ",   # figure space
            "\u202F": " ",   # narrow no-break space
            "\u200B": "",    # zero-width space
            "\u200C": "",    # zero-width non-joiner
            "\u200D": "",    # zero-width joiner
            "\u2060": "",    # word joiner
            "\uFEFF": "",    # BOM / zero-width no-break space
        }

        for bad, good in replacements.items():
            text = text.replace(bad, good)

        return text

    def canonicalize_scalar(value):
        """
        Canonical standard:
        - Treat blank-like values consistently
        - Preserve text content
        - Strip leading/trailing spaces
        - Normalize hidden/invisible characters
        """
        if pd.isna(value):
            return ""

        if isinstance(value, str):
            value = normalize_invisible_chars(value)
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
    # 9️⃣ TEXT NORMALIZATION / CANONICAL HELPERS
    # ==================================================

    def preserve_display_text(text):
        """
        Preserve business-facing/output text.
        - Keep original case
        - Keep line breaks
        - Keep internal spacing structure
        - Normalize invisible chars and line endings only
        - Strip only outer whitespace
        """
        if pd.isna(text):
            return ""

        text = normalize_invisible_chars(text)
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.strip()

        if text.lower() in {"nan", "none", "null"}:
            return ""

        return text

    def normalize_matching_text(text):
        """
        For fields used in keyword/phrase matching.
        Lowercase + normalize whitespace + normalize invisible chars.
        """
        if pd.isna(text):
            return ""

        text = normalize_invisible_chars(text)
        text = str(text).strip()
        text = re.sub(r"\s+", " ", text)
        text = text.lower()
        return text

    def normalize_canonical_key(text):
        """
        For controlled exact-match keys where we want
        in-memory canonical matching without changing
        business-facing values.
        """
        if pd.isna(text):
            return ""

        text = normalize_invisible_chars(text)
        text = str(text).strip()
        text = re.sub(r"\s+", " ", text)
        text = text.lower()
        return text

    # Preserve original business-facing columns
    if "Short description" in end_user_tickets.columns:
        end_user_tickets["Short description"] = (
            end_user_tickets["Short description"].apply(preserve_display_text)
        )
        end_user_tickets["Short description_canonical"] = (
            end_user_tickets["Short description"].apply(normalize_matching_text)
        )
    else:
        end_user_tickets["Short description"] = ""
        end_user_tickets["Short description_canonical"] = ""

    if "Description" in end_user_tickets.columns:
        end_user_tickets["Description"] = (
            end_user_tickets["Description"].apply(preserve_display_text)
        )
        end_user_tickets["Description_canonical"] = (
            end_user_tickets["Description"].apply(normalize_matching_text)
        )
    else:
        end_user_tickets["Description"] = ""
        end_user_tickets["Description_canonical"] = ""

    if "Close notes" in end_user_tickets.columns:
        end_user_tickets["Close notes"] = (
            end_user_tickets["Close notes"].apply(preserve_display_text)
        )
    else:
        end_user_tickets["Close notes"] = ""

    if "Service" in end_user_tickets.columns:
        end_user_tickets["Service"] = (
            end_user_tickets["Service"].apply(preserve_display_text)
        )
        end_user_tickets["Service_canonical"] = (
            end_user_tickets["Service"].apply(normalize_canonical_key)
        )
    else:
        end_user_tickets["Service"] = ""
        end_user_tickets["Service_canonical"] = ""

    if "Category" in end_user_tickets.columns:
        end_user_tickets["Category"] = (
            end_user_tickets["Category"].apply(preserve_display_text)
        )
        end_user_tickets["Category_canonical"] = (
            end_user_tickets["Category"].apply(normalize_canonical_key)
        )
    else:
        end_user_tickets["Category"] = ""
        end_user_tickets["Category_canonical"] = ""

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

    if len(month_values) > 1:
        print(f"❌ Multiple months detected in input data: {', '.join(sorted(month_values))}")
        print(f"❌ Selected month folder: {run_month}")
        print("❌ Please ensure only one month of ticket data is placed in the selected folder.")
        return None

    derived_month = month_values[0]

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

    keyword_rules_df["Priority"] = pd.to_numeric(
        keyword_rules_df["Priority"],
        errors="coerce"
    ).fillna(999)

    keyword_rules_df = keyword_rules_df.sort_values(
        by=["Category", "Keyword", "Priority"],
        kind="stable"
    ).reset_index(drop=True)

    category_keywords = {}
    category_priority = {}

    for _, row in keyword_rules_df.iterrows():
        category = row["Category"]
        keyword = row["Keyword"]
        priority = int(row["Priority"])

        category_keywords.setdefault(category, []).append(keyword)
        category_priority[category] = min(
            priority,
            category_priority.get(category, 999)
        )

    for category in category_keywords:
        category_keywords[category] = sorted(set(category_keywords[category]))

    phrase_rules_df["Category"] = (
        phrase_rules_df["Category"].astype(str).str.strip()
    )
    phrase_rules_df["Phrase"] = (
        phrase_rules_df["Phrase"].astype(str).str.lower().str.strip()
    )
    phrase_rules_df["Weight"] = pd.to_numeric(
        phrase_rules_df["Weight"],
        errors="coerce"
    ).fillna(0)

    phrase_rules_df = phrase_rules_df.sort_values(
        by=["Category", "Phrase", "Weight"],
        kind="stable"
    ).reset_index(drop=True)

    category_phrases = {}

    for _, row in phrase_rules_df.iterrows():
        category = row["Category"]
        phrase = row["Phrase"]
        weight = row["Weight"]

        category_phrases.setdefault(category, []).append((phrase, weight))

    for category in category_phrases:
        category_phrases[category] = sorted(
            category_phrases[category],
            key=lambda x: (x[0], x[1])
        )

    # ==================================================
    # 1️⃣2️⃣ LOAD CATEGORY & SERVICE RULES
    # ==================================================

    category_rules_df["Category"] = (
        category_rules_df["Category"].astype(str).str.lower().str.strip()
    )
    category_rules_df["Mapped Category"] = (
        category_rules_df["Mapped Category"].astype(str).str.strip()
    )

    category_rules_df = category_rules_df.sort_values(
        by=["Category", "Mapped Category"],
        kind="stable"
    ).reset_index(drop=True)

    category_rules = dict(
        zip(category_rules_df["Category"], category_rules_df["Mapped Category"])
    )

    service_rules_df["Service"] = (
        service_rules_df["Service"]
        .astype(str)
        .apply(normalize_canonical_key)
    )
    service_rules_df["Category"] = (
        service_rules_df["Category"].astype(str).str.strip()
    )

    service_rules_df = service_rules_df.sort_values(
        by=["Service", "Category"],
        kind="stable"
    ).reset_index(drop=True)

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
    ].copy()

    final_category_consolidation_df["Old Category Name_canonical"] = (
        final_category_consolidation_df["Old Category Name"].apply(normalize_canonical_key)
    )

    final_category_consolidation_df = final_category_consolidation_df.sort_values(
        by=["Old Category Name_canonical", "Final Mapped Category"],
        kind="stable"
    ).reset_index(drop=True)

    final_category_map = dict(
        zip(
            final_category_consolidation_df["Old Category Name_canonical"],
            final_category_consolidation_df["Final Mapped Category"]
        )
    )

    # ==================================================
    # 1️⃣4️⃣ TEMPLATE SANITIZATION
    # ==================================================

    template_noise_df["Phrase"] = (
        template_noise_df["Phrase"]
        .astype(str)
        .apply(normalize_invisible_chars)
        .str.strip()
    )

    template_noise_df = template_noise_df.sort_values(
        by=["Phrase"],
        kind="stable"
    ).reset_index(drop=True)

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
        end_user_tickets["Description_canonical"].apply(clean_description)
    )
        # ==================================================
    # 1️⃣5️⃣ WEIGHTED SCORING ENGINE + DEBUG TRACE
    # ==================================================

    MIN_SCORE_THRESHOLD = 1

    def determine_category_with_trace(row):

        trace = {
            "Decision Path": "",
            "Category Override Match": "",
            "Service Rule Match": "",
            "Service Fallback Used": "No",
            "Microsoft Teams Match": "No",
            "Top Candidate 1": "",
            "Top Candidate 1 Score": "",
            "Top Candidate 1 Phrase Score": "",
            "Top Candidate 1 Priority": "",
            "Top Candidate 2": "",
            "Top Candidate 2 Score": "",
            "Top Candidate 2 Phrase Score": "",
            "Top Candidate 2 Priority": "",
            "Top Candidate 3": "",
            "Top Candidate 3 Score": "",
            "Top Candidate 3 Phrase Score": "",
            "Top Candidate 3 Priority": "",
            "All Candidate Scores": ""
        }

        category_value = str(row.get("Category_canonical", "")).strip()
        service_value = str(row.get("Service_canonical", "")).strip()

        # -----------------------------
        # 1️⃣ CATEGORY OVERRIDE
        # -----------------------------
        if category_value in category_rules:
            winning_category = category_rules[category_value]
            trace["Decision Path"] = "category_override"
            trace["Category Override Match"] = winning_category
            return winning_category, trace

        # -----------------------------
        # 2️⃣ SERVICE MAPPING / FALLBACK
        # -----------------------------
        if service_value:
            if service_value in service_category_map:
                winning_category = service_category_map[service_value]
                trace["Decision Path"] = "service_mapping"
                trace["Service Rule Match"] = winning_category
                return winning_category, trace
            else:
                winning_category = str(row.get("Service", "")).strip()
                trace["Decision Path"] = "service_fallback"
                trace["Service Fallback Used"] = "Yes"
                return winning_category, trace

        short_desc = str(row.get("Short description_canonical", "")).strip()
        description = str(row.get("Description_clean", "")).strip()

        # -----------------------------
        # 3️⃣ MICROSOFT TEAMS RULE
        # -----------------------------
        if "Microsoft Teams" in category_keywords:
            for keyword in sorted(category_keywords["Microsoft Teams"]):
                if keyword in short_desc:
                    trace["Decision Path"] = "microsoft_teams"
                    trace["Microsoft Teams Match"] = "Yes"
                    return "Microsoft Teams", trace

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
                pattern = r"\b" + re.escape(keyword) + r"\b"
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
            trace["Decision Path"] = "weighted_scoring_no_match"
            return "Others", trace

        sorted_categories = sorted(
            scores.items(),
            key=lambda x: (
                -x[1]["total"],
                -x[1]["phrase"],
                x[1]["priority"],
                x[0]
            )
        )

        trace["Decision Path"] = "weighted_scoring"

        if len(sorted_categories) >= 1:
            trace["Top Candidate 1"] = sorted_categories[0][0]
            trace["Top Candidate 1 Score"] = sorted_categories[0][1]["total"]
            trace["Top Candidate 1 Phrase Score"] = sorted_categories[0][1]["phrase"]
            trace["Top Candidate 1 Priority"] = sorted_categories[0][1]["priority"]

        if len(sorted_categories) >= 2:
            trace["Top Candidate 2"] = sorted_categories[1][0]
            trace["Top Candidate 2 Score"] = sorted_categories[1][1]["total"]
            trace["Top Candidate 2 Phrase Score"] = sorted_categories[1][1]["phrase"]
            trace["Top Candidate 2 Priority"] = sorted_categories[1][1]["priority"]

        if len(sorted_categories) >= 3:
            trace["Top Candidate 3"] = sorted_categories[2][0]
            trace["Top Candidate 3 Score"] = sorted_categories[2][1]["total"]
            trace["Top Candidate 3 Phrase Score"] = sorted_categories[2][1]["phrase"]
            trace["Top Candidate 3 Priority"] = sorted_categories[2][1]["priority"]

        trace["All Candidate Scores"] = " | ".join(
            [
                f"{cat}: total={data['total']}, phrase={data['phrase']}, priority={data['priority']}"
                for cat, data in sorted_categories
            ]
        )

        best_category, best_data = sorted_categories[0]

        if best_data["total"] >= MIN_SCORE_THRESHOLD:
            return best_category, trace

        return "Others", trace

    category_trace_results = end_user_tickets.apply(
        determine_category_with_trace,
        axis=1
    )

    end_user_tickets["Ticket Category"] = category_trace_results.apply(lambda x: x[0])
    debug_trace_df = pd.DataFrame(category_trace_results.apply(lambda x: x[1]).tolist())

    end_user_tickets = pd.concat(
        [end_user_tickets.reset_index(drop=True), debug_trace_df.reset_index(drop=True)],
        axis=1
    )

    # ==================================================
    # 1️⃣6️⃣ FINAL CATEGORY CONSOLIDATION
    # ==================================================

    end_user_tickets["Ticket Category"] = (
        end_user_tickets["Ticket Category"]
        .astype(str)
        .str.strip()
    )

    end_user_tickets["Ticket Category"] = (
        end_user_tickets["Ticket Category"]
        .apply(lambda x: final_category_map.get(normalize_canonical_key(x), x))
    )

    # ==================================================
    # 1️⃣7️⃣ DEBUG FILE PREPARATION
    # ==================================================

    debug_columns = [
        "Number",
        "Ticket Type",
        "Opened",
        "Month",
        "Caller",
        "Email",
        "Contact type",
        "Assignment group",
        "Priority",
        "Category",
        "Category_canonical",
        "Service",
        "Service_canonical",
        "Short description",
        "Short description_canonical",
        "Description_clean",
        "Ticket Category",
        "Decision Path",
        "Category Override Match",
        "Service Rule Match",
        "Service Fallback Used",
        "Microsoft Teams Match",
        "Top Candidate 1",
        "Top Candidate 1 Score",
        "Top Candidate 1 Phrase Score",
        "Top Candidate 1 Priority",
        "Top Candidate 2",
        "Top Candidate 2 Score",
        "Top Candidate 2 Phrase Score",
        "Top Candidate 2 Priority",
        "Top Candidate 3",
        "Top Candidate 3 Score",
        "Top Candidate 3 Phrase Score",
        "Top Candidate 3 Priority",
        "All Candidate Scores"
    ]

    debug_output_df = end_user_tickets[
        end_user_tickets["Number"].astype(str).isin(DEBUG_TICKETS)
    ].copy()

    debug_output_df = debug_output_df[debug_columns].copy()

    debug_file = os.path.join(
        OUTPUT_DIR,
        f"debug_ticket_trace_{run_month.lower()}.xlsx"
    )

    # ==================================================
    # 1️⃣8️⃣ FINAL OUTPUT PREPARATION
    # ==================================================

    final_columns = required_columns + [
        "Ticket Type", "Month", "Ticket Category"
    ]

    final_output_df = end_user_tickets[final_columns].copy()

    def remove_garbage_rows(df):
        df = df[
            ~df["Number"].astype(str).str.contains(
                "Export stopped", case=False, na=False
            )
        ].copy()

        df = df.dropna(how="all")

        df = df[
            ~(
                df["Number"].isna() &
                df["Opened"].isna() &
                (df["Short description"].astype(str).str.strip() == "") &
                (df["Description"].astype(str).str.strip() == "")
            )
        ].copy()

        return df

    def sanitize_excel_cell(value):
        if pd.isna(value):
            return value

        if not isinstance(value, str):
            return value

        text = value

        if len(text) > 32767:
            text = text[:32767]

        if text.startswith(("=", "+", "-", "@")):
            text = "'" + text

        return text

    def sanitize_for_excel(df):
        text_columns = df.select_dtypes(include=["object"]).columns.tolist()

        for col in text_columns:
            df[col] = df[col].apply(sanitize_excel_cell)

        return df

    final_output_df = remove_garbage_rows(final_output_df)
    debug_output_df = remove_garbage_rows(debug_output_df)

    output_file = os.path.join(
        OUTPUT_DIR,
        f"end_user_ticket_data_{run_month.lower()}.xlsx"
    )

    master_file = os.path.join(
        OUTPUT_DIR,
        "master_ticket_data.xlsx"
    )

    # ==================================================
    # 1️⃣9️⃣ SAVE MONTHLY OUTPUT + DEBUG OUTPUT
    # ==================================================

    monthly_output_df = sanitize_for_excel(final_output_df.copy())
    monthly_output_df.to_excel(output_file, index=False)

    debug_output_export_df = sanitize_for_excel(debug_output_df.copy())
    debug_output_export_df.to_excel(debug_file, index=False)

    # ==================================================
    # 2️⃣0️⃣ REBUILD MASTER FILE FROM MONTHLY OUTPUTS
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
        master_df = pd.DataFrame(columns=final_output_df.columns)

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
    print("✅ Original text preserved in output.")
    print("✅ Canonical helper fields used for matching logic.")
    print("✅ Debug trace file created for selected tickets.")
    print("✅ Monthly output saved to:", output_file)
    print("✅ Debug output saved to:", debug_file)
    print("✅ Master file rebuilt from monthly outputs:", master_file)

    return final_output_df