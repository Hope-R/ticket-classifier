from engine.ticketcategories import process_tickets


def main():
    print("🚀 Starting Ticket Categorization Automation...")

    # Ask user for month
    run_month = input("📅 Enter month to process (e.g. Jan-26): ").strip()

    if not run_month:
        print("❌ No month entered. Exiting...")
        return

    process_tickets(run_month)

    print("✅ Ticket Categorization Completed Successfully!")


if __name__ == "__main__":
    main()