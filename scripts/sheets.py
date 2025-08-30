import gspread
import csv
import os
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import List, Dict, Any, Optional
from config import get_or_create_worksheet, LEFTOVERS_TAB_NAME
from orders import load_orders

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


def detect_changes_before_merge(sheet_edits: Dict[tuple, Dict[str, Any]], orders_dir: str) -> Dict[str, List[Dict]]:
    """Detect changes between sheet edits and current order files."""
    if not sheet_edits:
        return {'edits': [], 'additions': [], 'deletions': []}
    
    try:
        # Temporarily override ORDERS_DIR for loading
        import config
        original_orders_dir = config.ORDERS_DIR
        config.ORDERS_DIR = orders_dir
        
        try:
            inventory_list, orders_data = load_orders()
        finally:
            config.ORDERS_DIR = original_orders_dir
        
        # Build current data structure similar to sheet format
        current_data = {}
        
        # Handle both Order objects (real execution) and row dicts (mocked execution)
        if orders_data and isinstance(orders_data[0], dict):
            # Mocked execution - orders_data is a list of row dictionaries
            for row in orders_data:
                order_id = str(row.get("Order ID", "")).strip()
                item_id = str(row.get("Item Number", "")).strip()
                
                if order_id:
                    key = (order_id, item_id)
                    current_data[key] = {
                        "Order ID": order_id,
                        "Seller": str(row.get("Seller", "")),
                        "Order Date": str(row.get("Order Date", "")),
                        "Order Total": str(row.get("Order Total", "")),
                        "Base Grand Total": str(row.get("Base Grand Total", "")),
                        "Shipping": str(row.get("Shipping", "")),
                        "Add Chrg": str(row.get("Add Chrg", "")),
                        "Total Lots": str(row.get("Total Lots", "")),
                        "Total Items": str(row.get("Total Items", "")),
                        "Tracking #": str(row.get("Tracking #", "")),
                        "Item Number": item_id,
                        "Item Description": str(row.get("Item Description", "")),
                        "Color": str(row.get("Color", "")),
                        "Condition": str(row.get("Condition", "")),
                        "Qty": str(row.get("Qty", "")),
                        "Each": str(row.get("Each", "")),
                        "Total": str(row.get("Total", ""))
                    }
        else:
            # Real execution - orders_data is a list of Order objects
            for order in orders_data:
                # Add order header row
                current_data[(order.order_id, "")] = {
                    "Order ID": order.order_id,
                    "Seller": order.seller,
                    "Order Date": order.order_date,
                    "Order Total": str(order.order_total),
                    "Base Grand Total": str(order.base_grand_total),
                    "Shipping": str(order.shipping),
                    "Add Chrg": str(order.add_chrg_1),
                    "Total Lots": str(order.total_lots),
                    "Total Items": str(order.total_items),
                    "Tracking #": order.tracking_no,
                    "Item Number": ""
                }
                
                # Add item rows
                for item in order.items:
                    current_data[(order.order_id, item.item_id)] = {
                        "Order ID": order.order_id,
                        "Item Number": item.item_id,
                        "Item Description": item.description,
                        "Color": getattr(item, 'color_name', item.item_type),
                        "Condition": item.condition,
                        "Qty": str(item.qty),
                        "Each": str(item.price),
                        "Total": str(item.qty * item.price)
                    }
        
        # Detect changes
        edits = []
        additions = []
        deletions = []
        
        # Check for edits and additions
        for key, sheet_data in sheet_edits.items():
            order_id, item_id = key
            if key in current_data:
                # Check if any fields differ
                current_record = current_data[key]
                has_changes = False
                
                for field, sheet_value in sheet_data.items():
                    if field in current_record:
                        current_value = str(current_record[field])
                        sheet_value_str = str(sheet_value)
                        if current_value != sheet_value_str:
                            has_changes = True
                            break
                
                if has_changes:
                    edits.append({
                        'key': key,
                        'order_id': order_id,
                        'item_number': item_id,
                        'changes': sheet_data
                    })
            else:
                # New entry
                additions.append({
                    'key': key,
                    'order_id': order_id,
                    'item_number': item_id,
                    'data': sheet_data
                })
        
        # Check for deletions
        for key in current_data:
            if key not in sheet_edits:
                order_id, item_id = key
                deletions.append({
                    'key': key,
                    'order_id': order_id,
                    'item_number': item_id
                })
        
        return {'edits': edits, 'additions': additions, 'deletions': deletions}
        
    except Exception:
        return {'edits': [], 'additions': [], 'deletions': []}


def save_edits_to_files(sheet_edits: Dict[tuple, Dict[str, Any]], orders_dir: str) -> None:
    """Save sheet edits back to XML and CSV files."""
    if not sheet_edits:
        return
    
    try:
        import sys
        import os
        import csv
        import xml.etree.ElementTree as ET
        sys.path.insert(0, os.path.dirname(__file__))
        from orders import load_orders, Order, OrderItem
        
        # Temporarily override ORDERS_DIR for loading
        import config
        original_orders_dir = config.ORDERS_DIR
        config.ORDERS_DIR = orders_dir
        
        try:
            _, orders_list = load_orders()
        finally:
            config.ORDERS_DIR = original_orders_dir
        
        # Build dict of existing orders for easy lookup
        orders_dict = {order.order_id: order for order in orders_list}
        
        # Apply edits to orders
        for key, edit_data in sheet_edits.items():
            order_id, item_id = key
            
            # Get or create order
            if order_id not in orders_dict:
                # Create new order
                orders_dict[order_id] = Order(
                    order_id=order_id,
                    order_date="",
                    seller="",
                    order_total=0.0,
                    base_grand_total=0.0,
                    items=[]
                )
            
            order = orders_dict[order_id]
            
            if not item_id:  # Order-level edit
                # Update order fields
                if "Seller" in edit_data and edit_data["Seller"]:
                    order.seller = str(edit_data["Seller"])
                if "Order Date" in edit_data and edit_data["Order Date"]:
                    order.order_date = str(edit_data["Order Date"])
                if "Order Total" in edit_data and edit_data["Order Total"]:
                    order.order_total = float(str(edit_data["Order Total"]).replace('$', '').replace(',', ''))
                if "Base Grand Total" in edit_data and edit_data["Base Grand Total"]:
                    order.base_grand_total = float(str(edit_data["Base Grand Total"]).replace('$', '').replace(',', ''))
                if "Shipping" in edit_data and edit_data["Shipping"]:
                    order.shipping = float(str(edit_data["Shipping"]).replace('$', '').replace(',', ''))
                if "Add Chrg" in edit_data and edit_data["Add Chrg"]:
                    order.add_chrg_1 = float(str(edit_data["Add Chrg"]).replace('$', '').replace(',', ''))
                if "Total Lots" in edit_data and edit_data["Total Lots"]:
                    order.total_lots = int(edit_data["Total Lots"])
                if "Total Items" in edit_data and edit_data["Total Items"]:
                    order.total_items = int(edit_data["Total Items"])
                if "Tracking #" in edit_data and edit_data["Tracking #"]:
                    order.tracking_no = str(edit_data["Tracking #"])
            else:  # Item-level edit
                # Find or create item
                item = None
                for existing_item in order.items:
                    if existing_item.item_id == item_id:
                        item = existing_item
                        break
                
                if not item:
                    # Create new item
                    item = OrderItem(
                        item_id=item_id,
                        item_type="P",
                        color_id=0,
                        qty=0,
                        price=0.0,
                        condition="N"
                    )
                    order.items.append(item)
                
                # Update item fields
                if "Condition" in edit_data and edit_data["Condition"]:
                    item.condition = str(edit_data["Condition"])
                if "Qty" in edit_data and edit_data["Qty"]:
                    item.qty = int(edit_data["Qty"])
                if "Each" in edit_data and edit_data["Each"]:
                    item.price = float(str(edit_data["Each"]).replace('$', '').replace(',', ''))
                if "Item Description" in edit_data and edit_data["Item Description"]:
                    item.description = str(edit_data["Item Description"])
        
        # Save to XML files
        for filename in os.listdir(orders_dir):
            if filename.endswith('.xml'):
                xml_path = os.path.join(orders_dir, filename)
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    
                    # Update each order in the XML
                    for order_elem in root.findall("ORDER"):
                        order_id = order_elem.findtext("ORDERID", "").strip()
                        if order_id in orders_dict:
                            order = orders_dict[order_id]
                            
                            # Update order fields
                            if order_elem.find("SELLER") is not None:
                                order_elem.find("SELLER").text = order.seller
                            if order_elem.find("ORDERDATE") is not None:
                                order_elem.find("ORDERDATE").text = order.order_date
                            if order_elem.find("ORDERTOTAL") is not None:
                                order_elem.find("ORDERTOTAL").text = f"{order.order_total:.2f}"
                            if order_elem.find("BASEGRANDTOTAL") is not None:
                                order_elem.find("BASEGRANDTOTAL").text = f"{order.base_grand_total:.2f}"
                            
                            # Update item fields
                            for item_elem in order_elem.findall("ITEM"):
                                item_id = item_elem.findtext("ITEMID", "").strip()
                                for item in order.items:
                                    if item.item_id == item_id:
                                        if item_elem.find("CONDITION") is not None:
                                            item_elem.find("CONDITION").text = item.condition
                                        if item_elem.find("QTY") is not None:
                                            item_elem.find("QTY").text = str(item.qty)
                                        if item_elem.find("PRICE") is not None:
                                            item_elem.find("PRICE").text = f"{item.price:.2f}"
                                        if item_elem.find("DESCRIPTION") is not None:
                                            item_elem.find("DESCRIPTION").text = item.description
                                        break
                    
                    # Write back to file
                    ET.indent(tree, space="  ", level=0)
                    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
                except ET.ParseError:
                    continue
        
        # Save to CSV files  
        for filename in os.listdir(orders_dir):
            if filename.endswith('.csv'):
                csv_path = os.path.join(orders_dir, filename)
                try:
                    # Read existing CSV
                    rows = []
                    with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        rows = list(reader)
                    
                    # Update rows based on edits
                    for i, row in enumerate(rows):
                        order_id = row.get("Order ID", "").strip()
                        item_id = row.get("Item Number", "").strip()
                        key = (order_id, item_id)
                        
                        if key in sheet_edits:
                            edit_data = sheet_edits[key]
                            
                            # Update CSV fields based on edits
                            if "Condition" in edit_data and edit_data["Condition"]:
                                row["Condition"] = str(edit_data["Condition"])
                            if "Qty" in edit_data and edit_data["Qty"]:
                                row["Qty"] = str(edit_data["Qty"])
                            if "Each" in edit_data and edit_data["Each"]:
                                row["Each"] = str(edit_data["Each"])
                            if "Total" in edit_data and edit_data["Total"]:
                                row["Total"] = str(edit_data["Total"])
                            if "Item Description" in edit_data and edit_data["Item Description"]:
                                row["Item Description"] = str(edit_data["Item Description"])
                    
                    # Write back to CSV
                    if rows:
                        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                            writer.writeheader()
                            writer.writerows(rows)
                except Exception:
                    continue
                    
    except Exception:
        pass


def detect_deleted_orders(original_rows: List[Dict], sheet_edits: Dict[tuple, Dict[str, Any]]) -> List[tuple]:
    """Detect orders/items that have been deleted from the sheet."""
    deleted_keys = []
    
    # Build set of keys from original data
    original_keys = set()
    for row in original_rows:
        order_id = row.get("Order ID", "").strip()
        item_id = row.get("Item Number", "").strip()
        if order_id:  # Only add if we have an order ID
            original_keys.add((order_id, item_id))
    
    # Find keys that exist in original but not in sheet edits
    sheet_keys = set(sheet_edits.keys())
    deleted_keys = list(original_keys - sheet_keys)
    
    return deleted_keys


def remove_deleted_orders_from_files(deleted_keys: List[tuple], orders_dir: str) -> None:
    """Remove deleted orders/items from XML and CSV files."""
    if not deleted_keys:
        return
    
    # Remove from XML files
    for filename in os.listdir(orders_dir):
            if filename.endswith('.xml'):
                xml_path = os.path.join(orders_dir, filename)
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    
                    # Find orders/items to remove
                    orders_to_remove = []
                    for order_elem in root.findall("ORDER"):
                        order_id = order_elem.findtext("ORDERID", "").strip()
                        
                        # Check if entire order should be deleted
                        if (order_id, "") in deleted_keys:
                            orders_to_remove.append(order_elem)
                        else:
                            # Check for items to delete within the order
                            items_to_remove = []
                            for item_elem in order_elem.findall("ITEM"):
                                item_id = item_elem.findtext("ITEMID", "").strip()
                                if (order_id, item_id) in deleted_keys:
                                    items_to_remove.append(item_elem)
                            
                            # Remove items
                            for item_elem in items_to_remove:
                                order_elem.remove(item_elem)
                    
                    # Remove entire orders
                    for order_elem in orders_to_remove:
                        root.remove(order_elem)
                    
                    # Write back to file
                    ET.indent(tree, space="  ", level=0)
                    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
                except ET.ParseError:
                    continue
    
    # Remove from CSV files
    for filename in os.listdir(orders_dir):
        if filename.endswith('.csv'):
                csv_path = os.path.join(orders_dir, filename)
                try:
                    # Read existing CSV
                    rows = []
                    with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        rows = list(reader)
                    
                    # Filter out deleted rows
                    filtered_rows = []
                    for row in rows:
                        order_id = row.get("Order ID", "").strip()
                        item_id = row.get("Item Number", "").strip()
                        key = (order_id, item_id)
                        
                        if key not in deleted_keys:
                            filtered_rows.append(row)
                    
                    # Write back to CSV
                    if rows:  # Only if we had data originally
                        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                            if filtered_rows:
                                writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                                writer.writeheader()
                                writer.writerows(filtered_rows)
                            else:
                                # If no rows left, write just headers
                                writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                                writer.writeheader()
                except Exception:
                    continue


def detect_deleted_orders(original_rows: List[Dict], sheet_edits: Dict[tuple, Dict[str, Any]]) -> List[tuple]:
    """Detect orders/items that have been deleted from the sheet."""
    deleted_keys = []
    
    # Build set of keys from original data
    original_keys = set()
    for row in original_rows:
        order_id = str(row.get("Order ID", "")).strip()
        item_id = str(row.get("Item Number", "")).strip()
        if order_id:  # Only add if we have an order ID
            original_keys.add((order_id, item_id))
    
    # Find keys that exist in original but not in sheet edits
    sheet_keys = set(sheet_edits.keys())
    deleted_keys = list(original_keys - sheet_keys)
    
    return deleted_keys