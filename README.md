# Ticket Classifier

Python automation for categorizing IT support tickets using keyword and phrase-based scoring rules.

## Overview

This project processes ServiceNow ticket exports and automatically categorizes tickets using configurable business rules.

The classification engine supports:

- Keyword matching
- Phrase matching with weights
- Category overrides
- Service mapping rules

## Project Structure

ticket-classifier
│
├── main.py
├── config/
│   └── business_rules_template.xlsx
├── raw_data/
└── output/

## Setup

1. Copy the template:

business_rules_template.xlsx → business_rules.xlsx

2. Place the real business rules in the **config** folder.

3. Place ticket export files in:

raw_data/

4. Run the script:

python main.py

## Security

Sensitive data such as ticket exports and real business rules are excluded using `.gitignore`.
