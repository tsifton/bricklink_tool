from config import load_google_sheet
from orders import load_orders  # ...existing code...
from wanted_lists import parse_wanted_lists
from build_logic import determine_buildable
from sheets import (
    update_summary,
    update_inventory_sheet,
    update_leftovers,
    update_orders_sheet
)
import merge_orders

def main():
    """
    Main entry point for the Minifig Profit Tool.
    Loads data, computes buildable quantities, and updates Google Sheets.
    """

    # Merge order files
    merge_orders.merge_xml()
    merge_orders.merge_csv()

    # Load inventory (list[OrderItem]) and orders (list[Order])
    inv_list, orders_list = load_orders()

    # Load or create the main Google Sheet
    sheet = load_google_sheet()

    # Update the Inventory worksheet with the current inventory (pre-build)
    update_inventory_sheet(sheet, inv_list)

    # Use object-based wanted lists for build logic
    wanted_lists = parse_wanted_lists()

    summary_rows = []
    # For each wanted list, determine how many builds can be made and the cost
    for wl in wanted_lists:
        count, cost, updated_inventory_list = determine_buildable(wl, inv_list)
        if count:
            inv_list = updated_inventory_list
        avg_cost = round(cost / count, 2) if count else 0.0
        summary_rows.append([wl.title, count, avg_cost, "", "", "", "", "", "", "", ""])

    # Update the Summary worksheet with build results and formulas
    update_summary(sheet, summary_rows)
    # Update the Leftover Inventory worksheet with remaining inventory (post-build)
    update_leftovers(sheet, inv_list)
    # Update the Orders worksheet with all order and item rows
    update_orders_sheet(sheet, orders_list)

if __name__ == "__main__":
    # Run the main function if this script is executed directly
    main()