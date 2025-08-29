import gspread
from config import get_or_create_worksheet, LEFTOVERS_TAB_NAME

def update_summary(sheet, summary_rows):
    """
    Updates the 'Summary' worksheet with summary_rows data.
    Preserves existing prices, sets formulas for profit, margin, markup, and formats columns.
    """
    ws = get_or_create_worksheet(sheet, "Summary")
    existing = ws.get_all_records()
    # Map existing minifig IDs to their prices
    existing_prices = {row['Minifig ID']: row.get('Price') for row in existing if 'Minifig ID' in row}
    # Write header row
    ws.update(values=[[
        "Minifig ID", "Buildable", "Avg Cost", "Price", "Profit", "Margin", "Markup",
        "75%", "100%", "125%", "150%"
    ]], range_name="A1")
    values = []
    for i, row in enumerate(summary_rows, start=2):
        title = row[0]
        # Preserve existing price if available, otherwise set default formula
        existing_price = existing_prices.get(title)
        row[3] = existing_price if existing_price is not None and str(existing_price).strip() else "=14.99"
        values.append(row)
    # Write summary data rows
    ws.update(values=values, range_name="A2", value_input_option="USER_ENTERED")
    # Prepare formula cells for profit, margin, markup, and various markups
    profit_cells = [gspread.Cell(i+2, 5, f"=ROUND((D{i+2} * 0.85) - C{i+2} - Config!$B$1 - Config!$B$2, 2)") for i in range(len(summary_rows))]
    margin_cells = [gspread.Cell(i+2, 6, f"=IF(D{i+2}=0, \"\", ROUND(E{i+2} / D{i+2}, 2))") for i in range(len(summary_rows))]
    markup_cells = [gspread.Cell(i+2, 7, f"=IF(C{i+2}=0, \"\", ROUND(E{i+2} / C{i+2}, 2))") for i in range(len(summary_rows))]
    markup_75_cells = [gspread.Cell(i+2, 8, f"=CEILING(((D{i+2} * 0.85) - (Config!$B$1 + Config!$B$2)) / 1.75, 0.25)") for i in range(len(summary_rows))]
    markup_100_cells = [gspread.Cell(i+2, 9, f"=CEILING(((D{i+2} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.0, 0.25)") for i in range(len(summary_rows))]
    markup_125_cells = [gspread.Cell(i+2, 10, f"=CEILING(((D{i+2} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.25, 0.25)") for i in range(len(summary_rows))]
    markup_150_cells = [gspread.Cell(i+2, 11, f"=CEILING(((D{i+2} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.5, 0.25)") for i in range(len(summary_rows))]
    # Batch update all formula cells
    ws.update_cells(
        profit_cells + margin_cells + markup_cells +
        markup_75_cells + markup_100_cells +
        markup_125_cells + markup_150_cells,
        value_input_option="USER_ENTERED"
    )
    try:
        # Format margin and markup columns as percent, markups as currency
        end_row = len(summary_rows) + 1
        ws.format(f"F2:G{end_row}", {"numberFormat": {"type": "PERCENT", "pattern": "##0.00%"}})
        ws.format(f"H2:K{end_row}", {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}})
    except Exception:
        pass  # Ignore formatting errors

def update_inventory_sheet(sheet, inventory):
    """
    Updates the 'Inventory' worksheet with inventory data.
    Only includes items with qty > 0.
    """
    ws = get_or_create_worksheet(sheet, "Inventory")
    ws.clear()
    # Write header row
    ws.update(values=[["Item ID", "Description", "Color", "Qty", "Total Cost", "Unit Cost"]], range_name="A1")
    # Prepare inventory rows
    rows = [
        [iid, item['description'], item['color_name'], item['qty'],
         round(item['total_cost'], 2), round(item['unit_cost'], 2)]
        for (iid, color_id), item in inventory.items() if item['qty'] > 0
    ]
    # Write inventory data rows
    if rows:
        ws.update(values=rows, range_name="A2")

def update_leftovers(sheet, inventory):
    """
    Updates the leftovers worksheet with items that have leftover quantity.
    Only includes items with qty > 0.
    """
    ws = get_or_create_worksheet(sheet, LEFTOVERS_TAB_NAME)
    ws.clear()
    # Write header row
    ws.update(values=[["Item ID", "Description", "Color", "Leftover Qty"]], range_name="A1")
    # Prepare leftover rows
    rows = [
        [iid, data['description'], data['color_name'], data['qty']]
        for (iid, _), data in inventory.items() if data['qty'] > 0
    ]
    # Write leftover data rows
    if rows:
        ws.update(values=rows, range_name="A2")

def update_orders_sheet(sheet, order_rows):
    """
    Updates the 'Orders' worksheet with order_rows data.
    Formats currency columns for 'Each' and 'Total'.
    """
    if not order_rows:
        return
    ws = get_or_create_worksheet(sheet, "Orders")
    ws.clear()
    # Define header columns
    headers = [
        "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg 1",
        "Order Total", "Base Grand Total", "Total Lots", "Total Items",
        "Tracking No", "Condition", "Item Number", "Item Description",
        "Color", "Qty", "Each", "Total"
    ]
    # Prepare order data rows
    values = [headers] + [[row.get(col, '') for col in headers] for row in order_rows]
    ws.update(values=values, range_name="A1")
    try:
        # Format 'Each' and 'Total' columns as currency
        each_index = headers.index("Each")
        each_letter = chr(ord('A') + each_index)
        total_index = headers.index("Total")
        total_letter = chr(ord('A') + total_index)
        last_row = len(values)
        ws.format(f"{each_letter}2:{each_letter}{last_row}", {
            "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
        })
        ws.format(f"{total_letter}2:{total_letter}{last_row}", {
            "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
        })
    except Exception:
        pass  # Ignore formatting errors