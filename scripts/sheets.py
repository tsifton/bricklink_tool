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

        # Prefer cleaned description (CSV minus seller note)
        desc = getattr(it, 'clean_description', None) or getattr(it, 'description', '') or entry['description']
        # For parts, remove leading color name and any whitespace after it
        cn = get_color_name(it.color_id) if it.item_type == 'P' else None
        if it.item_type == 'P' and cn:
            if desc.startswith(cn):
                desc = desc[len(cn):].lstrip()
            elif desc.lower().startswith(cn.lower()):
                desc = desc[len(cn):].lstrip()

        agg[key].update({
            'qty': new_qty,
            'total_cost': new_total_cost,
            'unit_cost': (new_total_cost / new_qty) if new_qty else 0.0,
            'description': desc,
            'color_id': it.color_id if it.item_type == 'P' else None,
            'color_name': cn if it.item_type == 'P' else None,
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
    Returns a dictionary keyed by (Order ID, Item #) with ALL fields.

    Handles structure where only the first item row of each order shows Order ID
    and subsequent item rows have an empty Order ID (we carry forward context).
    """
    try:
        ws = get_or_create_worksheet(sheet, "Orders")
        records = ws.get_all_records()

        edits = {}
        current_order_id = ""

        for record in records:
            # Get Order ID from this row (or carry forward)
            row_order_id = record.get("Order ID", "")
            item_number = record.get("Item #", "") or record.get("Item Number", "")  # fallback to legacy name

            if row_order_id:
                current_order_id = row_order_id
                order_id = row_order_id
            else:
                order_id = current_order_id

            if not order_id:
                continue

            # Every row is an item row in the new structure
            key = (order_id, item_number)

            record_copy = record.copy()
            if not row_order_id and order_id:
                record_copy["Order ID"] = order_id

            edits[key] = record_copy

        return edits
    except Exception:
        return {}

def _format_currency_cols(ws, headers, cols, last_row):
    try:
        for col in cols:
            if col in headers:
                col_idx = headers.index(col)
                col_letter = chr(ord('A') + col_idx)
                ws.format(f"{col_letter}2:{col_letter}{last_row}", {
                    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
                })
    except Exception:
        pass  # Ignore formatting errors

def update_orders_sheet(sheet, orders):
    """
    Updates the 'Orders' worksheet from a list[Order] objects.
    First item row of each order includes order-level fields for readability.
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
        for idx, item in enumerate(order.items):
            color_name = get_color_name(item.color_id)
            # Strip color prefix for Orders sheet description as well
            desc = item.description or ""
            if item.item_type == 'P' and color_name:
                if desc.startswith(color_name):
                    desc = desc[len(color_name):].lstrip()
                elif desc.lower().startswith(color_name.lower()):
                    desc = desc[len(color_name):].lstrip()

            # Populate order-level fields on first item row; blank on subsequent rows
            item_dict = {
                "Order ID": order.order_id if idx == 0 else "",
                "Seller": order.seller if idx == 0 else "",
                "Order Date": order.order_date if idx == 0 else "",
                "Shipping": order.shipping if idx == 0 else "",
                "Add Chrg": order.add_chrg_1 if idx == 0 else "",
                "Subtotal": order.order_total if idx == 0 else "",
                "Order Total": order.base_grand_total if idx == 0 else "",
                "Total Lots": order.total_lots if idx == 0 else "",
                "Total Items": order.total_items if idx == 0 else "",
                "Tracking #": order.tracking_no if idx == 0 else "",
                "Condition": item.condition,
                "Item #": item.item_id,
                "Description": desc,
                "Color": color_name or item.item_type,
                "Qty": item.qty,
                "Each": item.price,
                "Total": item.qty * item.price
            }

            # Apply any user edits for this item row
            key_item = (order.order_id, item.item_id)
            if key_item in existing_edits:
                user_record = existing_edits[key_item]
                for field in headers:
                    val = user_record.get(field, '')
                    if str(val).strip():
                        item_dict[field] = val
                # Ensure non-first rows keep order fields blank for readability
                if idx > 0:
                    for f in ["Order ID", "Seller", "Order Date", "Shipping", "Add Chrg",
                              "Subtotal", "Order Total", "Total Lots", "Total Items", "Tracking #"]:
                        item_dict[f] = ""

            row_data = [item_dict.get(col, '') for col in headers]
            data_rows.append(row_data)

    values = [headers] + data_rows
    ws.update(values=values, range_name="A1")
    last_row = len(values)
    _format_currency_cols(ws, headers, ["Shipping", "Add Chrg", "Subtotal", "Order Total", "Each", "Total"], last_row)