import gspread
import os
import csv
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import List, Dict, Any, Optional, Tuple
from config import get_or_create_worksheet, LEFTOVERS_TAB_NAME

# Constants
INVENTORY_HEADERS = ["Item ID", "Description", "Color", "Qty", "Total Cost", "Unit Cost"]
SUMMARY_HEADERS = [
    "Minifig ID", "Buildable", "Avg Cost", "Price", "Profit", "Margin", "Markup",
    "75%", "100%", "125%", "150%"
]
ORDERS_HEADERS = [
    "Order ID","Seller","Order Date","Shipping","Add Chrg 1","Order Total","Base Grand Total","Total Lots","Total Items","Tracking No","Condition","Item Number","Item Description","Color","Qty","Each","Total"
]
ORDER_LEVEL_FIELDS = [
    "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg 1",
    "Order Total", "Base Grand Total", "Total Lots", "Total Items", "Tracking No"
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
            item_number = record.get("Item Number", "")

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
                "Add Chrg 1": order.add_chrg_1 if is_first_item else "",
                "Order Total": order.order_total if is_first_item else "",
                "Base Grand Total": order.base_grand_total if is_first_item else "",
                "Total Lots": order.total_lots if is_first_item else "",
                "Total Items": order.total_items if is_first_item else "",
                "Tracking No": order.tracking_no if is_first_item else "",
                "Condition": item.condition,
                "Item Number": item.item_id,
                "Item Description": desc,
                "Color": getattr(item, 'color_name', item.item_type),
                "Qty": item.qty,
                "Each": item.price,
                "Total": item.qty * item.price
            }

            # Apply user edits
            key_item = (order.order_id, item.item_id)
            key_order = (order.order_id, "")
            
            # Apply order-level edits (for first item row only)
            if is_first_item and key_order in existing_edits:
                user_record = existing_edits[key_order]
                for field in ORDERS_HEADERS:
                    val = user_record.get(field, '')
                    if str(val).strip():
                        item_dict[field] = val
            
            # Apply item-level edits
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
    _format_currency_columns(ws, ORDERS_HEADERS, ["Shipping", "Add Chrg 1", "Order Total", "Base Grand Total", "Each", "Total"], len(values))

    
def detect_changes_before_merge(sheet_edits: Optional[Dict[tuple, Dict[str, Any]]], orders_dir: str) -> Dict[str, List[Dict[str, Any]]]:
    """Detect changes between sheet edits and existing order files before merging."""
    if not sheet_edits:
        return {'edits': [], 'additions': [], 'deletions': []}
    
    # Load existing order data from files
    existing_orders = {}
    
    if os.path.exists(orders_dir):
        # Check for XML files first
        xml_files = ['orders.xml'] if os.path.exists(os.path.join(orders_dir, 'orders.xml')) else [
            f for f in os.listdir(orders_dir) if f.endswith('.xml')
        ]
        
        for filename in xml_files:
            filepath = os.path.join(orders_dir, filename)
            try:
                tree = ET.parse(filepath)
                for order_elem in tree.getroot().findall("ORDER"):
                    order_id = (order_elem.findtext("ORDERID") or "").strip()
                    if not order_id:
                        continue
                    
                    # Store order header info
                    existing_orders[(order_id, "")] = {
                        "Order ID": order_id,
                        "Seller": (order_elem.findtext("SELLER") or "").strip(),
                        "Order Date": (order_elem.findtext("ORDERDATE") or "").strip(),
                        "Order Total": (order_elem.findtext("ORDERTOTAL") or "").strip(),
                        "Base Grand Total": (order_elem.findtext("BASEGRANDTOTAL") or "").strip(),
                        "Item Number": ""
                    }
                    
                    # Store item info
                    for item_elem in order_elem.findall("ITEM"):
                        item_id = (item_elem.findtext("ITEMID") or "").strip()
                        if item_id:
                            existing_orders[(order_id, item_id)] = {
                                "Order ID": order_id,
                                "Item Number": item_id,
                                "Item Description": (item_elem.findtext("DESCRIPTION") or "").strip(),
                                "Color": (item_elem.findtext("COLOR") or "").strip(),
                                "Condition": (item_elem.findtext("CONDITION") or "").strip(),
                                "Qty": (item_elem.findtext("QTY") or "").strip(),
                                "Each": (item_elem.findtext("PRICE") or "").strip(),
                                "Total": ""
                            }
            except (ET.ParseError, Exception):
                continue
    
    changes = {'edits': [], 'additions': [], 'deletions': []}
    
    # Check for edits and additions
    for key, sheet_data in sheet_edits.items():
        order_id, item_number = key
        if key in existing_orders:
            # Compare with existing data to detect edits - only check meaningful fields
            existing_data = existing_orders[key]
            has_changes = False
            
            # Define which fields to check for changes
            fields_to_check = ["Seller", "Order Date", "Order Total", "Base Grand Total"] if not item_number else [
                "Item Description", "Condition", "Qty", "Each"  # Exclude Color since XML stores IDs, sheets show names
            ]
            
            for field in fields_to_check:
                sheet_val = str(sheet_data.get(field, "")).strip()
                existing_val = str(existing_data.get(field, "")).strip()
                if sheet_val and sheet_val != existing_val:
                    has_changes = True
                    break
            
            if has_changes:
                changes['edits'].append({
                    'key': key,
                    'order_id': order_id,
                    'item_number': item_number,
                    'changes': sheet_data
                })
        else:
            # Not in existing files = addition
            changes['additions'].append({
                'key': key,
                'order_id': order_id,
                'item_number': item_number,
                'data': sheet_data
            })
    
    # Check for deletions
    for key in existing_orders:
        if key not in sheet_edits:
            order_id, item_number = key
            changes['deletions'].append({
                'key': key,
                'order_id': order_id,
                'item_number': item_number
            })
    
    return changes


def save_edits_to_files(sheet_edits: Dict[tuple, Dict[str, Any]], orders_dir: str) -> None:
    """Save edited data back to order files."""
    if not sheet_edits or not os.path.exists(orders_dir):
        return
    
    # Update XML files
    xml_files = ['orders.xml'] if os.path.exists(os.path.join(orders_dir, 'orders.xml')) else [
        f for f in os.listdir(orders_dir) if f.endswith('.xml')
    ]
    
    for filename in xml_files:
        filepath = os.path.join(orders_dir, filename)
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            
            for order_elem in root.findall("ORDER"):
                order_id = (order_elem.findtext("ORDERID") or "").strip()
                if not order_id:
                    continue
                
                # Update order-level fields
                order_key = (order_id, "")
                if order_key in sheet_edits:
                    edits = sheet_edits[order_key]
                    if "Seller" in edits and edits["Seller"].strip():
                        order_elem.find("SELLER").text = edits["Seller"]
                    if "Order Date" in edits and edits["Order Date"].strip():
                        order_elem.find("ORDERDATE").text = edits["Order Date"]
                    if "Order Total" in edits and edits["Order Total"].strip():
                        order_elem.find("ORDERTOTAL").text = edits["Order Total"]
                    if "Base Grand Total" in edits and edits["Base Grand Total"].strip():
                        order_elem.find("BASEGRANDTOTAL").text = edits["Base Grand Total"]
                
                # Update item-level fields
                for item_elem in order_elem.findall("ITEM"):
                    item_id = (item_elem.findtext("ITEMID") or "").strip()
                    item_key = (order_id, item_id)
                    
                    if item_key in sheet_edits:
                        edits = sheet_edits[item_key]
                        if "Condition" in edits and edits["Condition"].strip():
                            item_elem.find("CONDITION").text = edits["Condition"]
                        if "Qty" in edits and edits["Qty"].strip():
                            item_elem.find("QTY").text = edits["Qty"]
                        if "Each" in edits and edits["Each"].strip():
                            item_elem.find("PRICE").text = edits["Each"]
                        if "Item Description" in edits and edits["Item Description"].strip():
                            item_elem.find("DESCRIPTION").text = edits["Item Description"]
            
            # Write back to file
            ET.indent(tree, space="  ", level=0)
            tree.write(filepath, encoding='utf-8', xml_declaration=True)
            
        except (ET.ParseError, Exception):
            continue
    
    # Update CSV files
    csv_files = ['orders.csv'] if os.path.exists(os.path.join(orders_dir, 'orders.csv')) else [
        f for f in os.listdir(orders_dir) if f.endswith('.csv')
    ]
    
    for filename in csv_files:
        filepath = os.path.join(orders_dir, filename)
        try:
            # Read existing CSV
            rows = []
            with open(filepath, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                fieldnames = reader.fieldnames
                for row in reader:
                    order_id = (row.get("Order ID") or "").strip()
                    item_number = (row.get("Item Number") or "").strip()
                    key = (order_id, item_number)
                    
                    # Apply edits if they exist
                    if key in sheet_edits:
                        edits = sheet_edits[key]
                        for field, value in edits.items():
                            if field in row and str(value).strip():
                                row[field] = value
                    
                    rows.append(row)
            
            # Write back to file
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
                
        except Exception:
            continue


def detect_deleted_orders(original_rows: List[Dict[str, Any]], sheet_edits: Dict[tuple, Dict[str, Any]]) -> List[Tuple[str, str]]:
    """Detect orders/items that were deleted from the sheet."""
    deleted_keys = []
    
    for row in original_rows:
        order_id = (row.get("Order ID") or "").strip()
        item_number = (row.get("Item Number") or "").strip()
        key = (order_id, item_number)
        
        if key not in sheet_edits:
            deleted_keys.append(key)
    
    return deleted_keys


def remove_deleted_orders_from_files(deleted_keys: List[Tuple[str, str]], orders_dir: str) -> None:
    """Remove deleted orders/items from order files."""
    if not deleted_keys or not os.path.exists(orders_dir):
        return
    
    # Process XML files
    xml_files = ['orders.xml'] if os.path.exists(os.path.join(orders_dir, 'orders.xml')) else [
        f for f in os.listdir(orders_dir) if f.endswith('.xml')
    ]
    
    for filename in xml_files:
        filepath = os.path.join(orders_dir, filename)
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            
            # Remove orders and items based on deleted keys
            orders_to_remove = []
            for order_elem in root.findall("ORDER"):
                order_id = (order_elem.findtext("ORDERID") or "").strip()
                
                # Check if entire order should be deleted
                if (order_id, "") in deleted_keys:
                    orders_to_remove.append(order_elem)
                    continue
                
                # Remove specific items
                items_to_remove = []
                for item_elem in order_elem.findall("ITEM"):
                    item_id = (item_elem.findtext("ITEMID") or "").strip()
                    if (order_id, item_id) in deleted_keys:
                        items_to_remove.append(item_elem)
                
                for item_elem in items_to_remove:
                    order_elem.remove(item_elem)
            
            # Remove entire orders
            for order_elem in orders_to_remove:
                root.remove(order_elem)
            
            # Write back to file
            ET.indent(tree, space="  ", level=0)
            tree.write(filepath, encoding='utf-8', xml_declaration=True)
            
        except (ET.ParseError, Exception):
            continue
    
    # Process CSV files
    csv_files = ['orders.csv'] if os.path.exists(os.path.join(orders_dir, 'orders.csv')) else [
        f for f in os.listdir(orders_dir) if f.endswith('.csv')
    ]
    
    for filename in csv_files:
        filepath = os.path.join(orders_dir, filename)
        try:
            # Read existing CSV
            remaining_rows = []
            with open(filepath, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                fieldnames = reader.fieldnames
                
                for row in reader:
                    order_id = (row.get("Order ID") or "").strip()
                    item_number = (row.get("Item Number") or "").strip()
                    key = (order_id, item_number)
                    
                    # Keep row if it's not in deleted keys
                    if key not in deleted_keys:
                        remaining_rows.append(row)
            
            # Write back to file
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(remaining_rows)
                
        except Exception:
            continue