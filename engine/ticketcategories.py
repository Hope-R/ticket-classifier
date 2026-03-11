import os
import pandas as pd
import re


def process_tickets():

    # ==================================================
    # 1️⃣ BASE PATHS
    # ==================================================

    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    CONFIG_DIR = os.path.join(BASE_DIR, "config")
    RAW_DIR = os.path.join(BASE_DIR, "raw_data")
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")

    # ==================================================
    # 2️⃣ LOAD CONFIG FILES
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

    # ==================================================
    # 3️⃣ COLUMN STANDARDIZATION
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
    # 4️⃣ LOAD RAW FILES
    # ==================================================

    incident_df = standardize_columns(
        pd.read_excel(os.path.join(RAW_DIR, "incident_jan_26.xlsx"))
    )

    ur_df = standardize_columns(
        pd.read_excel(os.path.join(RAW_DIR, "ur_jan_26.xlsx"))
    )

    task_df = standardize_columns(
        pd.read_excel(os.path.join(RAW_DIR, "task_jan_26.xlsx"))
    )

    # ==================================================
    # 5️⃣ BUSINESS FILTERS
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
    # 6️⃣ ADD METADATA
    # ==================================================

    for df, ttype in [(incident_df, "Inc"), (ur_df, "UR"), (task_df, "Task")]:
        df["Ticket Type"] = ttype
        df["Month"] = "Jan-26"

    # ==================================================
    # 7️⃣ COMBINE
    # ==================================================

    end_user_tickets = pd.concat(
        [incident_df, ur_df, task_df],
        ignore_index=True
    )

    # ==================================================
    # 8️⃣ CLEAN TEXT FIELDS
    # ==================================================

    for col in ["Short description", "Description", "Service"]:
        if col in end_user_tickets.columns:
            end_user_tickets[col] = (
                end_user_tickets[col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    # ==================================================
    # 9️⃣ PREPARE RULE LOGIC
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
    # 🔟 LOAD CATEGORY & SERVICE RULES
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
    # 1️⃣1️⃣ TEMPLATE SANITIZATION
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
    # 1️⃣2️⃣ WEIGHTED SCORING ENGINE
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

        # --------------------------------------------------
        # STABILITY FIX:
        # Evaluate categories from BOTH keyword rules
        # and phrase rules, but in a stable sorted order
        # --------------------------------------------------

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
    # 1️⃣3️⃣ FINAL OUTPUT
    # ==================================================

    final_columns = required_columns + [
        "Ticket Type", "Month", "Ticket Category"
    ]

    end_user_tickets = end_user_tickets[final_columns]

    output_file = os.path.join(
        OUTPUT_DIR,
        "end_user_ticket_data_jan-26.xlsx"
    )

    end_user_tickets.to_excel(output_file, index=False)

    print("✅ Weighted scoring processing complete.")
    print("Saved to:", output_file)

    return end_user_tickets
