# Ticket Categorization – Code Flow

## Actual Execution Flow (Derived from Code)

0. RUN MONTH NORMALIZATION
   - Normalize input month to format: Jan-26
   - Handles variations like Jan26, jan-26, etc.

1–2. PATH SETUP & INPUT VALIDATION
   - Build base directories: config/, raw_data/<run_month>/, output/
   - Resolve input files: incident, ur, task
   - Validate:
       • Only one file per type
       • Supports .xlsx or .csv
       • Fail if missing or duplicates

3. LOAD BUSINESS RULES
   - Load all sheets from business_rules.xlsx:
       • assignment_groups
       • keyword_rules
       • phrase_rules
       • template_noise
       • category_check
       • service_check
       • column_mapping
       • final_category_consolidation
   - Uses safe Excel reader (openpyxl)

4. COLUMN MAPPING SETUP
   - Build alias dictionary from column_mapping sheet
   - Define:
       • required_columns
       • logic_required_columns
   - Prepare for column standardization

5. CANONICAL INGESTION (RAW FILE LOAD)
   - Load incident / ur / task files (.xlsx or .csv)
   - Force dtype=str for consistency
   - Apply:
       • Invisible character normalization
       • Null normalization (nan, none, null → "")
       • Whitespace cleanup
   - Apply column standardization AFTER ingestion

6. BUSINESS FILTERS
   - Incidents:
       • Exclude Contact type = Alert
   - UR:
       • Keep Primary Ticket blank
       • AND Category not blank
   - Tasks:
       • Filter using assignment_groups list

7. METADATA ADDITION
   - Add column:
       • Ticket Type = Inc / UR / Task

8. COMBINE DATASETS
   - Merge all filtered datasets using concat

9. TEXT NORMALIZATION (CANONICAL LAYER)
   - Create dual representation:
       • Original text → for output
       • Canonical text → for matching
   - Generate:
       • Short description_canonical
       • Description_canonical
       • Service_canonical
       • Category_canonical
   - Preserve original text for output

10. MONTH DERIVATION & VALIDATION
    - Convert Opened to datetime
    - Derive Month = %b-%y
    - Validate:
        • Only ONE month present
        • Must match run_month

11–13. RULE PREPARATION
    - Keyword rules:
        • Build category_keywords
        • Build category_priority
    - Phrase rules:
        • Build category_phrases with weights
    - Category override:
        • category_rules (category_check)
    - Service mapping:
        • service_category_map (service_check)
    - Final consolidation:
        • final_category_map
    - Ensure deterministic sorting

14. TEMPLATE SANITIZATION
    - Remove template_noise phrases from Description
    - Create:
        • Description_clean

15. CATEGORIZATION ENGINE (CORE LOGIC)
    Apply in strict order:

    1) Category override (category_check)
    2) Service mapping (service_check)
    3) Service fallback (use Service as category)
    4) Microsoft Teams special rule (Short description only)
    5) Weighted scoring:
        • Phrase matches (weighted)
        • Keyword matches (weight = 1)
        • Combined scoring
    6) Tie-breaking:
        • Highest total score
        • Highest phrase score
        • Lowest priority
        • Alphabetical order
    7) Others fallback if no match

16. FINAL CATEGORY CONSOLIDATION
    - Map category variants to standardized final categories
    - Uses canonical matching

17. FINAL OUTPUT PREPARATION
    - Select final columns:
        • required_columns + Ticket Type + Month + Ticket Category
    - Remove garbage rows:
        • Empty rows
        • Export artifacts
    - Apply Excel sanitization:
        • Trim long text (>32767 chars)
        • Prevent formula injection (=, +, -, @)

18. SAVE MONTHLY OUTPUT
    - Save file:
        • output/end_user_ticket_data_<month>.xlsx

19. REBUILD MASTER FILE
    - Read all monthly output files
    - Combine into master dataset
    - Deduplicate using Ticket Number
    - Clean data again
    - Save:
        • output/master_ticket_data.xlsx