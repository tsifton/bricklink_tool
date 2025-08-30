import gspread
from collections import defaultdict
from typing import List, Dict, Any, Optional
from config import get_or_create_worksheet, LEFTOVERS_TAB_NAME

# Constants
INVENTORY_HEADERS = ["Item ID", "Description", "Color", "Qty", "Total Cost", "Unit Cost"]
SUMMARY_HEADERS = [
    "Minifig ID", "Buildable", "Avg Cost", "Price", "Profit", "Margin", "Markup",
    "75%", "100%", "125%", "150%"
]
ORDERS_HEADERS = [
    "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg",
    "Subtotal", "Order Total", "Total Lots", "Total Items",
    "Tracking #", "Condition", "Item #", "Description",
    "Color", "Qty", "Each", "Total"
]
ORDER_LEVEL_FIELDS = [
    "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg",
    "Subtotal", "Order Total", "Total Lots", "Total Items", "Tracking #"
]

def update_summary(sheet, summary_rows: List[List[Any]]) -> None:
    """Updates the 'Summary' worksheet with summary data and formulas."""
    ws = get_or_create_worksheet(sheet, "Summary")
    existing = ws.get_all_records()
    
    # Map existing minifig IDs to their prices
    existing_prices = {
        row['Minifig ID']: row.get('Price') 
        for row in existing if 'Minifig ID' in row
    }
    
    # Write headers and preserve existing prices
    ws.update(values=[SUMMARY_HEADERS], range_name="A1")
    
    for i, row in enumerate(summary_rows, start=2):
        title = row[0]
        existing_price = existing_prices.get(title)
        row[3] = (existing_price if existing_price is not None 
                 and str(existing_price).strip() else "=14.99")
    
    ws.update(values=summary_rows, range_name="A2", value_input_option="USER_ENTERED")
    
    # Generate formula cells
    row_count = len(summary_rows)
    formula_cells = []
    
    for i in range(row_count):
        row_num = i + 2
        formula_cells.extend([
            gspread.Cell(row_num, 5, f"=ROUND((D{row_num} * 0.85) - C{row_num} - Config!$B$1 - Config!$B$2, 2)"),
            gspread.Cell(row_num, 6, f"=IF(D{row_num}=0, \"\", ROUND(E{row_num} / D{row_num}, 2))"),
            gspread.Cell(row_num, 7, f"=IF(C{row_num}=0, \"\", ROUND(E{row_num} / C{row_num}, 2))"),
            gspread.Cell(row_num, 8, f"=CEILING(((D{row_num} * 0.85) - (Config!$B$1 + Config!$B$2)) / 1.75, 0.25)"),
            gspread.Cell(row_num, 9, f"=CEILING(((D{row_num} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.0, 0.25)"),
            gspread.Cell(row_num, 10, f"=CEILING(((D{row_num} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.25, 0.25)"),
            gspread.Cell(row_num, 11, f"=CEILING(((D{row_num} * 0.85) - (Config!$B$1 + Config!$B$2)) / 2.5, 0.25)"),
        ])
    
    ws.update_cells(formula_cells, value_input_option="USER_ENTERED")
    
    # Format columns
    try:
        end_row = row_count + 1
        ws.format(f"F2:G{end_row}", {"numberFormat": {"type": "PERCENT", "pattern": "##0.00%"}})
        ws.format(f"H2:K{end_row}", {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}})
    except Exception:
        pass

def _strip_color_prefix(description: str, color_name: Optional[str]) -> str:
    """Remove color name prefix from description if present."""
    if not color_name or not description:
        return description
        
    if description.startswith(color_name):
        return description[len(color_name):].lstrip()
    elif description.lower().startswith(color_name.lower()):
        return description[len(color_name):].lstrip()
    return description

def _aggregate_inventory(items) -> Dict[tuple, Dict[str, Any]]:
    """Aggregate OrderItems by (item_id, color_key)."""
    agg = defaultdict(lambda: {
        'qty': 0, 'total_cost': 0.0, 'unit_cost': 0.0, 'description': '',
        'color_id': None, 'color_name': None, 'item_type': None,
    })
    
    for item in items or []:
        key = (item.item_id, None if item.item_type in ('S', 'M') else item.color_id)
        entry = agg[key]
        
        qty = int(getattr(item, 'qty', 0) or 0)
        cost = float(getattr(item, 'unit_cost', 0.0) or 0.0) * qty
        new_qty = entry['qty'] + qty
        new_total = entry['total_cost'] + cost

        # Get description and strip color prefix for parts
        desc = getattr(item, 'clean_description', None) or getattr(item, 'description', '') or entry['description']
        color_name = getattr(item, 'color_name', None)
        
        if item.item_type == 'P' and color_name and color_name != item.item_type:
            desc = _strip_color_prefix(desc, color_name)

        entry.update({
            'qty': new_qty,
            'total_cost': new_total,
            'unit_cost': new_total / new_qty if new_qty else 0.0,
            'description': desc,
            'color_id': item.color_id if item.item_type == 'P' else None,
            'color_name': color_name,
            'item_type': item.item_type,
        })
    return agg

def _update_inventory_worksheet(sheet, tab_name: str, items) -> None:
    """Generic function to update inventory-style worksheets."""
    ws = get_or_create_worksheet(sheet, tab_name)
    ws.clear()
    ws.update(values=[INVENTORY_HEADERS], range_name="A1")
    
    inventory = _aggregate_inventory(items)
    rows = [
        [item_id, data['description'], data['color_name'], data['qty'],
         round(data['total_cost'], 2), round(data['unit_cost'], 2)]
        for (item_id, _), data in inventory.items() if data['qty'] > 0
    ]
    
    if rows:
        ws.update(values=rows, range_name="A2")

def update_inventory_sheet(sheet, items) -> None:
    """Updates the 'Inventory' worksheet."""
    _update_inventory_worksheet(sheet, "Inventory", items)

def update_leftovers(sheet, items) -> None:
    """Updates the leftovers worksheet."""
    _update_inventory_worksheet(sheet, LEFTOVERS_TAB_NAME, items)

def read_orders_sheet_edits(sheet) -> Dict[tuple, Dict[str, Any]]:
    """Read Orders worksheet to capture user edits."""
    try:
        ws = get_or_create_worksheet(sheet, "Orders")
        records = ws.get_all_records()
        edits = {}
        current_order_id = ""

        for record in records:
            row_order_id = record.get("Order ID", "")
            item_number = record.get("Item #", "") or record.get("Item Number", "")

            order_id = row_order_id if row_order_id else current_order_id
            if row_order_id:
                current_order_id = row_order_id

            if not order_id:
                continue

            key = (order_id, item_number)
            record_copy = record.copy()
            if not row_order_id and order_id:
                record_copy["Order ID"] = order_id

            edits[key] = record_copy

        return edits
    except Exception:
        return {}

def _format_currency_columns(ws, headers: List[str], columns: List[str], last_row: int) -> None:
    """Format specified columns as currency."""
    try:
        for col in columns:
            if col in headers:
                col_idx = headers.index(col)
                col_letter = chr(ord('A') + col_idx)
                ws.format(f"{col_letter}2:{col_letter}{last_row}", {
                    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
                })
    except Exception:
        pass

def update_orders_sheet(sheet, orders) -> None:
    """Updates the 'Orders' worksheet from Order objects."""
    if not orders:
        return

    existing_edits = read_orders_sheet_edits(sheet)
    ws = get_or_create_worksheet(sheet, "Orders")
    ws.clear()

    data_rows = []
    for order in orders:
        for idx, item in enumerate(order.items):
            # Process description
            desc = item.description or ""
            color_name = getattr(item, 'color_name', None)
            if item.item_type == 'P' and color_name and color_name != item.item_type:
                desc = _strip_color_prefix(desc, color_name)

            # Build item dictionary with order-level fields on first row only
            is_first_item = idx == 0
            item_dict = {
                "Order ID": order.order_id if is_first_item else "",
                "Seller": order.seller if is_first_item else "",
                "Order Date": order.order_date if is_first_item else "",
                "Shipping": order.shipping if is_first_item else "",
                "Add Chrg": order.add_chrg_1 if is_first_item else "",
                "Subtotal": order.order_total if is_first_item else "",
                "Order Total": order.base_grand_total if is_first_item else "",
                "Total Lots": order.total_lots if is_first_item else "",
                "Total Items": order.total_items if is_first_item else "",
                "Tracking #": order.tracking_no if is_first_item else "",
                "Condition": item.condition,
                "Item #": item.item_id,
                "Description": desc,
                "Color": getattr(item, 'color_name', item.item_type),
                "Qty": item.qty,
                "Each": item.price,
                "Total": item.qty * item.price
            }

            # Apply user edits
            key_item = (order.order_id, item.item_id)
            if key_item in existing_edits:
                user_record = existing_edits[key_item]
                for field in ORDERS_HEADERS:
                    val = user_record.get(field, '')
                    if str(val).strip():
                        item_dict[field] = val
                
                # Keep order fields blank for non-first rows
                if not is_first_item:
                    for field in ORDER_LEVEL_FIELDS:
                        item_dict[field] = ""

            data_rows.append([item_dict.get(col, '') for col in ORDERS_HEADERS])

    values = [ORDERS_HEADERS] + data_rows
    ws.update(values=values, range_name="A1")
    _format_currency_columns(ws, ORDERS_HEADERS, 
                           ["Shipping", "Add Chrg", "Subtotal", "Order Total", "Each", "Total"], 
                           len(values))