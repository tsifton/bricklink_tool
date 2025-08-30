from config import get_config_value, load_google_sheet
from orders import load_orders
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

    # Load or create the main Google Sheet
    sheet = load_google_sheet()

    # Merge order files
    merge_orders.merge_xml()
    merge_orders.merge_csv()

    # Load inventory and order rows from BrickLink order files
    inventory, order_rows = load_orders(return_rows=True)
    
    # Retrieve configuration values (shipping fee and materials cost) from the Config worksheet
    shipping_fee = get_config_value(sheet, "Shipping Fee", "B1")
    materials_cost = get_config_value(sheet, "Materials Cost", "B2")
    
    # Update the Inventory worksheet with the current inventory
    update_inventory_sheet(sheet, inventory)

    # Parse all wanted lists from XML files
    wanted_lists = parse_wanted_lists()

    summary_rows = []
    # For each wanted list, determine how many builds can be made and the cost
    for title, items in wanted_lists.items():
        count, cost, updated_inventory = determine_buildable(items, inventory)
        if count:
            # Update inventory only if at least one build was possible
            inventory = updated_inventory
        # Calculate average cost per build (avoid division by zero)
        avg_cost = round(cost / count, 2) if count else 0.0
        # Prepare the summary row (placeholders for formulas)
        summary_rows.append([title, count, avg_cost, "", "", "", "", "", "", "", ""])

    # Update the Summary worksheet with build results and formulas
    update_summary(sheet, summary_rows)
    # Update the Leftover Inventory worksheet with remaining inventory
    update_leftovers(sheet, inventory)
    # Update the Orders worksheet with all order and item rows
    update_orders_sheet(sheet, order_rows)

if __name__ == "__main__":
    # Run the main function if this script is executed directly
    main()
