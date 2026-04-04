# main.py

from engine.ticketcategories import process_tickets


def normalize_run_month_text(value):
    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    value = value.replace("_", "-").replace("/", "-")

    # Remove internal spaces like "Jan - 26"
    value = "".join(value.split())

    if len(value) == 6 and "-" not in value:
        # Example: Jan26 -> Jan-26
        value = f"{value[:3]}-{value[3:]}"

    if len(value) >= 6:
        value = value[:3].title() + value[3:]

    return value


def main():
    print("🚀 Starting Ticket Categorization Automation...")

    run_month = input("📅 Enter month to process (e.g. Jan-26): ").strip()
    run_month = normalize_run_month_text(run_month)

    if not run_month:
        print("❌ No month entered. Exiting...")
        return

    result = process_tickets(run_month)

    if result is None:
        print("❌ Ticket Categorization did not complete.")
        return

    print("✅ Ticket Categorization Completed Successfully!")


if __name__ == "__main__":
    main()