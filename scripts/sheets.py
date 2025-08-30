import gspread
from config import get_or_create_worksheet, LEFTOVERS_TAB_NAME
from colors import get_color_name

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

def _aggregate_inventory(items):
    """
    Aggregate a list[OrderItem] into a map keyed by (item_id, color_key)
    where color_key is None for sets/minifigs and color_id for parts.
    """
    from collections import defaultdict
    agg = defaultdict(lambda: {
        'qty': 0,
        'total_cost': 0.0,
        'unit_cost': 0.0,
        'description': '',
        'color_id': None,
        'color_name': None,
        'item_type': None,
    })
    for it in items or []:
        key = (it.item_id, None if it.item_type in ('S', 'M') else it.color_id)
        entry = agg[key]
        line_qty = int(getattr(it, 'qty', 0) or 0)
        line_total_cost = float(getattr(it, 'unit_cost', 0.0) or 0.0) * line_qty
        new_qty = entry['qty'] + line_qty
        new_total_cost = entry['total_cost'] + line_total_cost
        agg[key].update({
            'qty': new_qty,
            'total_cost': new_total_cost,
            'unit_cost': (new_total_cost / new_qty) if new_qty else 0.0,
            'description': getattr(it, 'description', '') or entry['description'],
            'color_id': it.color_id if it.item_type == 'P' else None,
            'color_name': get_color_name(it.color_id) if it.item_type == 'P' else None,
            'item_type': it.item_type,
        })
    return agg

def update_inventory_sheet(sheet, items):
    """
    Updates the 'Inventory' worksheet from a list[OrderItem].
    Only includes items with qty > 0.
    """
    ws = get_or_create_worksheet(sheet, "Inventory")
    ws.clear()
    # Write header row
    ws.update(values=[["Item ID", "Description", "Color", "Qty", "Total Cost", "Unit Cost"]], range_name="A1")
    # Aggregate list to map
    inventory = _aggregate_inventory(items)
    # Prepare inventory rows
    rows = [
        [iid, data['description'], data['color_name'], data['qty'],
         round(data['total_cost'], 2), round(data['unit_cost'], 2)]
        for (iid, _), data in inventory.items() if data['qty'] > 0
    ]
    if rows:
        ws.update(values=rows, range_name="A2")

def update_leftovers(sheet, items):
    """
    Updates the leftovers worksheet from a list[OrderItem].
    Only includes items with qty > 0.
    """
    ws = get_or_create_worksheet(sheet, LEFTOVERS_TAB_NAME)
    ws.clear()
    # Write header row
    ws.update(values=[["Item ID", "Description", "Color", "Qty", "Total Cost", "Unit Cost"]], range_name="A1")
    # Aggregate list to map
    inventory = _aggregate_inventory(items)
    # Prepare inventory rows
    rows = [
        [iid, data['description'], data['color_name'], data['qty'],
         round(data['total_cost'], 2), round(data['unit_cost'], 2)]
        for (iid, _), data in inventory.items() if data['qty'] > 0
    ]
    if rows:
        ws.update(values=rows, range_name="A2")

def read_orders_sheet_edits(sheet):
    """
    Reads the current Orders worksheet to capture any user edits.
    Returns a dictionary keyed by (Order ID, Item Number) with ALL fields.
    
    Handles the order structure where only the first row of each order 
    contains the Order ID, and subsequent item rows have empty Order ID
    until the next order starts.
    """
    try:
        ws = get_or_create_worksheet(sheet, "Orders")
        records = ws.get_all_records()
        
        edits = {}
        current_order_id = ""
        
        for record in records:
            # Get the Order ID from this row
            row_order_id = record.get("Order ID", "")
            item_number = record.get("Item Number", "")
            
            # If this row has an Order ID, update our current order context
            if row_order_id:
                current_order_id = row_order_id
                order_id = row_order_id
            else:
                # This row doesn't have an Order ID, so it belongs to the current order
                order_id = current_order_id
            
            # Skip rows that don't belong to any order
            if not order_id:
                continue
            
            # For order header rows (no Item Number), use just Order ID
            # For item rows, use Order ID + Item Number
            key = (order_id, item_number) if item_number else (order_id, "")
            
            # Store all fields for this row, but ensure Order ID is set correctly
            # for item rows that had empty Order ID in the sheet
            record_copy = record.copy()
            if not row_order_id and order_id:
                record_copy["Order ID"] = order_id
            
            edits[key] = record_copy
                
        return edits
    except Exception:
        # If there's any error reading the sheet, return empty dict
        return {}

def update_orders_sheet(sheet, orders):
    """
    Updates the 'Orders' worksheet from a list[Order] objects.
    Preserves user edits to ALL fields and formats currency columns.
    """
    if not orders:
        return

    existing_edits = read_orders_sheet_edits(sheet)

    ws = get_or_create_worksheet(sheet, "Orders")
    ws.clear()
    headers = [
        "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg",
        "Subtotal", "Order Total", "Total Lots", "Total Items",
        "Tracking #", "Condition", "Item #", "Description",
        "Color", "Qty", "Each", "Total"
    ]

    data_rows = []
    for order in orders:
        # Header row for the order
        header_dict = {
            "Order ID": order.order_id,
            "Seller": order.seller,
            "Order Date": order.order_date,
            "Shipping": order.shipping,
            "Add Chrg": order.add_chrg_1,
            "Subtotal": order.order_total,
            "Order Total": order.base_grand_total,
            "Total Lots": order.total_lots,
            "Total Items": order.total_items,
            "Tracking #": order.tracking_no,
            "Condition": "",
            "Item #": "",
            "Description": "",
            "Color": "",
            "Qty": "",
            "Each": "",
            "Total": "",
        }
        # Apply edits for header row
        key_header = (order.order_id, "")
        if key_header in existing_edits:
            for field, value in existing_edits[key_header].items():
                if field in header_dict and str(value).strip():
                    header_dict[field] = value
        data_rows.append([header_dict.get(col, '') for col in headers])

        # Item rows
        for item in order.items:
            color_name = get_color_name(item.color_id)
            item_dict = {
                "Order ID": order.order_id,  # will be blanked in display
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Add Chrg": "",
                "Subtotal": "",
                "Order Total": "",
                "Total Lots": "",
                "Total Items": "",
                "Tracking #": "",
                "Condition": item.condition,
                "Item #": item.item_id,
                "Description": item.description or "",
                "Color": color_name or item.item_type,
                "Qty": item.qty,
                "Each": item.price,
                "Total": item.qty * item.price
            }
            # Apply edits for item row
            key_item = (order.order_id, item.item_id)
            if key_item in existing_edits:
                for field, value in existing_edits[key_item].items():
                    if field in item_dict and str(value).strip():
                        item_dict[field] = value
            row_data = [item_dict.get(col, '') for col in headers]
            # Clear Order ID for display on item rows
            row_data[headers.index("Order ID")] = ""
            data_rows.append(row_data)

    values = [headers] + data_rows
    ws.update(values=values, range_name="A1")
    try:
        # Format monetary columns as currency
        currency_cols = ["Shipping", "Add Chrg 1", "Order Total", "Base Grand Total", "Each", "Total"]
        last_row = len(values)
        for col in currency_cols:
            if col in headers:
                col_idx = headers.index(col)
                col_letter = chr(ord('A') + col_idx)
                ws.format(f"{col_letter}2:{col_letter}{last_row}", {
                    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
                })
    except Exception:
        pass  # Ignore formatting errors