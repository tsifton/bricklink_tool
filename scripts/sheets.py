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
    """Detect changes between sheet edits and existing order files."""
    if not sheet_edits:
        return {'edits': [], 'additions': [], 'deletions': []}
    
    changes = {'edits': [], 'additions': [], 'deletions': []}
    
    # Load existing orders from files for comparison
    existing_orders = _load_orders_from_files(orders_dir)
    
    # Compare sheet edits against existing file data
    for key, sheet_row in sheet_edits.items():
        order_id, item_id = key
        
        # Check if this order/item exists in files
        if order_id in existing_orders:
            if item_id == "":
                # Order header row - compare order-level data
                existing_order = existing_orders[order_id]
                if _has_order_changes(sheet_row, existing_order):
                    changes['edits'].append({'key': key, 'data': sheet_row, 'type': 'order'})
            else:
                # Item row - compare item-level data  
                if item_id in existing_orders[order_id].get('items', {}):
                    existing_item = existing_orders[order_id]['items'][item_id]
                    if _has_item_changes(sheet_row, existing_item):
                        changes['edits'].append({'key': key, 'data': sheet_row, 'type': 'item'})
                else:
                    # New item
                    changes['additions'].append({'key': key, 'data': sheet_row, 'type': 'item'})
        else:
            # New order
            changes['additions'].append({'key': key, 'data': sheet_row, 'type': 'order' if item_id == "" else 'item'})
    
    # TODO: Detect deletions by comparing existing orders against sheet edits
    
    return changes


def _load_orders_from_files(orders_dir: str) -> Dict[str, Dict[str, Any]]:
    """Load existing orders from XML/CSV files for comparison."""
    orders = {}
    
    if not os.path.exists(orders_dir):
        return orders
    
    # Check for merged XML file first
    xml_file = os.path.join(orders_dir, 'orders.xml')
    if os.path.exists(xml_file):
        orders.update(_load_orders_from_xml(xml_file))
    else:
        # Load from individual XML files
        for filename in os.listdir(orders_dir):
            if filename.endswith('.xml'):
                filepath = os.path.join(orders_dir, filename)
                orders.update(_load_orders_from_xml(filepath))
    
    return orders


def _load_orders_from_xml(xml_path: str) -> Dict[str, Dict[str, Any]]:
    """Load orders from a single XML file."""
    orders = {}
    
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        for order_elem in root.findall("ORDER"):
            order_id = (order_elem.findtext("ORDERID") or "").strip()
            if not order_id:
                continue
                
            order_data = {
                'order_id': order_id,
                'seller': (order_elem.findtext("SELLER") or "").strip(),
                'order_date': (order_elem.findtext("ORDERDATE") or "").strip(),
                'order_total': float(order_elem.findtext("ORDERTOTAL") or "0"),
                'base_grand_total': float(order_elem.findtext("BASEGRANDTOTAL") or "0"),
                'items': {}
            }
            
            for item_elem in order_elem.findall("ITEM"):
                item_id = (item_elem.findtext("ITEMID") or "").strip()
                if item_id:
                    order_data['items'][item_id] = {
                        'item_id': item_id,
                        'qty': int(item_elem.findtext("QTY") or "0"),
                        'price': float(item_elem.findtext("PRICE") or "0"),
                        'condition': (item_elem.findtext("CONDITION") or "").strip(),
                        'description': (item_elem.findtext("DESCRIPTION") or "").strip(),
                    }
            
            orders[order_id] = order_data
            
    except (ET.ParseError, ValueError, Exception):
        pass
        
    return orders


def _has_order_changes(sheet_row: Dict[str, Any], existing_order: Dict[str, Any]) -> bool:
    """Check if order-level data has changes."""
    # Compare key order fields
    order_fields = {
        'Seller': 'seller',
        'Order Date': 'order_date', 
        'Order Total': 'order_total',
        'Base Grand Total': 'base_grand_total',
    }
    
    for sheet_field, file_field in order_fields.items():
        sheet_val = str(sheet_row.get(sheet_field, '')).strip()
        file_val = str(existing_order.get(file_field, '')).strip()
        
        # Handle numeric fields
        if sheet_field in ['Order Total', 'Base Grand Total']:
            try:
                sheet_val = str(float(sheet_val)) if sheet_val else '0.0'
                file_val = str(float(file_val)) if file_val else '0.0'
            except ValueError:
                pass
        
        if sheet_val != file_val:
            return True
    
    return False


def _has_item_changes(sheet_row: Dict[str, Any], existing_item: Dict[str, Any]) -> bool:
    """Check if item-level data has changes."""
    # Compare key item fields  
    item_fields = {
        'Qty': 'qty',
        'Each': 'price',
        'Condition': 'condition',
    }
    
    # Handle description field (could be "Description" or "Item Description")
    description_field = 'Item Description' if 'Item Description' in sheet_row else 'Description'
    if description_field in sheet_row:
        item_fields[description_field] = 'description'
    
    for sheet_field, file_field in item_fields.items():
        sheet_val = str(sheet_row.get(sheet_field, '')).strip()
        file_val = str(existing_item.get(file_field, '')).strip()
        
        # Handle numeric fields
        if sheet_field in ['Qty', 'Each']:
            try:
                if sheet_field == 'Qty':
                    sheet_val = str(int(float(sheet_val))) if sheet_val else '0'
                    file_val = str(int(float(file_val))) if file_val else '0'
                else:  # Each (price)
                    sheet_val = str(float(sheet_val)) if sheet_val else '0.0'
                    file_val = str(float(file_val)) if file_val else '0.0'
            except ValueError:
                pass
        
        if sheet_val != file_val:
            return True
    
    return False


def save_edits_to_files(sheet_edits: Dict[tuple, Dict[str, Any]], orders_dir: str) -> None:
    """Save sheet edits back to XML/CSV files."""
    if not sheet_edits or not os.path.exists(orders_dir):
        return
    
    # Group edits by order ID
    orders_to_save = defaultdict(lambda: {'order_data': {}, 'items': {}})
    
    for key, row_data in sheet_edits.items():
        order_id, item_id = key
        
        if item_id == "":
            # Order header row
            orders_to_save[order_id]['order_data'] = row_data
        else:
            # Item row
            orders_to_save[order_id]['items'][item_id] = row_data
    
    # Save to XML files
    xml_file = os.path.join(orders_dir, 'orders.xml')
    _save_orders_to_xml(orders_to_save, xml_file)
    
    # Also save to CSV if CSV file exists
    csv_file = os.path.join(orders_dir, 'orders.csv')
    if os.path.exists(csv_file):
        _save_orders_to_csv(orders_to_save, csv_file)


def _save_orders_to_xml(orders_data: Dict[str, Dict[str, Any]], xml_path: str) -> None:
    """Save orders data to XML file, preserving minimal structure."""
    # Load existing XML if it exists to preserve other orders
    existing_orders = {}
    if os.path.exists(xml_path):
        existing_orders = _load_orders_from_xml(xml_path)
    
    # Update with new edits
    for order_id, order_info in orders_data.items():
        order_data = order_info['order_data']
        items_data = order_info['items']
        
        if order_id not in existing_orders:
            existing_orders[order_id] = {
                'order_id': order_id,
                'seller': '',
                'order_date': '',
                'order_total': 0.0,
                'base_grand_total': 0.0,
                'items': {}
            }
        
        # Update order-level data if provided
        if order_data:
            existing_orders[order_id]['seller'] = order_data.get('Seller', existing_orders[order_id]['seller'])
            existing_orders[order_id]['order_date'] = order_data.get('Order Date', existing_orders[order_id]['order_date'])
            
            try:
                existing_orders[order_id]['order_total'] = float(order_data.get('Order Total', existing_orders[order_id]['order_total']))
            except ValueError:
                pass
            
            try:
                base_grand_total = order_data.get('Base Grand Total') or order_data.get('Order Total')
                if base_grand_total:
                    existing_orders[order_id]['base_grand_total'] = float(base_grand_total)
            except ValueError:
                pass
        
        # Update item data
        for item_id, item_row in items_data.items():
            if item_id not in existing_orders[order_id]['items']:
                existing_orders[order_id]['items'][item_id] = {
                    'item_id': item_id,
                    'qty': 0,
                    'price': 0.0,
                    'condition': 'N',
                    'description': '',
                }
            
            item_data = existing_orders[order_id]['items'][item_id]
            
            try:
                item_data['qty'] = int(float(item_row.get('Qty', item_data['qty'])))
            except ValueError:
                pass
            
            try:
                item_data['price'] = float(item_row.get('Each', item_data['price']))
            except ValueError:
                pass
                
            item_data['condition'] = item_row.get('Condition', item_data['condition'])
            # Handle both "Description" and "Item Description" field names
            description = item_row.get('Item Description') or item_row.get('Description')
            if description:
                item_data['description'] = description
    
    # Write XML
    root = ET.Element("ORDERS")
    
    for order_id, order_data in existing_orders.items():
        order_elem = ET.SubElement(root, "ORDER")
        
        ET.SubElement(order_elem, "ORDERID").text = order_id
        ET.SubElement(order_elem, "ORDERDATE").text = order_data['order_date']
        ET.SubElement(order_elem, "SELLER").text = order_data['seller']
        ET.SubElement(order_elem, "ORDERTOTAL").text = f"{order_data['order_total']:.2f}"
        ET.SubElement(order_elem, "BASEGRANDTOTAL").text = f"{order_data['base_grand_total']:.2f}"
        
        for item_id, item_data in order_data['items'].items():
            item_elem = ET.SubElement(order_elem, "ITEM")
            
            ET.SubElement(item_elem, "ITEMID").text = item_id
            ET.SubElement(item_elem, "ITEMTYPE").text = "P"  # Default to part
            ET.SubElement(item_elem, "COLOR").text = "4"     # Default color
            ET.SubElement(item_elem, "QTY").text = str(item_data['qty'])
            ET.SubElement(item_elem, "PRICE").text = f"{item_data['price']:.2f}"
            ET.SubElement(item_elem, "CONDITION").text = item_data['condition']
            ET.SubElement(item_elem, "DESCRIPTION").text = item_data['description']
    
    # Write to file
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)


def _save_orders_to_csv(orders_data: Dict[str, Dict[str, Any]], csv_path: str) -> None:
    """Save orders data to CSV file in BrickLink export format."""
    # Load existing CSV data if it exists
    existing_data = []
    if os.path.exists(csv_path):
        try:
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                existing_data = list(reader)
        except Exception:
            existing_data = []
    
    # Update existing data with edits
    updated_data = []
    current_order_id = ""
    
    # Process each existing row
    for row in existing_data:
        row_order_id = row.get('Order ID', '').strip()
        item_number = row.get('Item Number', '').strip()
        
        # Track current order ID (same logic as read_orders_sheet_edits)
        if row_order_id:
            current_order_id = row_order_id
        
        order_id = row_order_id if row_order_id else current_order_id
        
        # Check if we have edits for this specific key
        has_edits = False
        for edit_order_id, order_info in orders_data.items():
            if edit_order_id == order_id:
                if item_number == '':
                    # Order header row
                    has_edits = bool(order_info.get('order_data'))
                else:
                    # Item row
                    has_edits = item_number in order_info.get('items', {})
                break
        
        if has_edits:
            # This row has edits, update it
            order_info = orders_data.get(order_id, {})
            
            if item_number == '':
                # Order header row
                order_data = order_info.get('order_data', {})
                if order_data:
                    if 'Seller' in order_data:
                        row['Seller'] = order_data['Seller']
                    if 'Order Date' in order_data:
                        row['Order Date'] = order_data['Order Date']
                    if 'Order Total' in order_data:
                        row['Order Total'] = order_data['Order Total']
                    if 'Base Grand Total' in order_data:
                        row['Base Grand Total'] = order_data['Base Grand Total']
            else:
                # Item row
                items_data = order_info.get('items', {})
                item_data = items_data.get(item_number, {})
                if item_data:
                    if 'Qty' in item_data:
                        row['Qty'] = item_data['Qty']
                    if 'Each' in item_data:
                        row['Each'] = item_data['Each']
                    if 'Condition' in item_data:
                        row['Condition'] = item_data['Condition']
                    description = item_data.get('Item Description') or item_data.get('Description')
                    if description:
                        row['Item Description'] = description
        
        updated_data.append(row)
    
    # Add any completely new orders/items
    for order_id, order_info in orders_data.items():
        # Check if this is a completely new order
        existing_order_ids = {row.get('Order ID', '').strip() for row in existing_data}
        
        if order_id not in existing_order_ids:
            # Add new order header row
            order_data = order_info.get('order_data', {})
            header_row = {
                'Order ID': order_id,
                'Seller': order_data.get('Seller', ''),
                'Order Date': order_data.get('Order Date', ''),
                'Order Total': order_data.get('Order Total', ''),
                'Base Grand Total': order_data.get('Base Grand Total', ''),
                'Item Number': '',
                'Item Description': '',
                'Qty': '',
                'Each': '',
                'Condition': ''
            }
            updated_data.append(header_row)
            
            # Add item rows
            for item_id, item_data in order_info.get('items', {}).items():
                item_row = {
                    'Order ID': '',  # Empty for item rows
                    'Seller': '',
                    'Order Date': '',
                    'Order Total': '',
                    'Base Grand Total': '',
                    'Item Number': item_id,
                    'Item Description': item_data.get('Item Description', item_data.get('Description', '')),
                    'Qty': item_data.get('Qty', ''),
                    'Each': item_data.get('Each', ''),
                    'Condition': item_data.get('Condition', '')
                }
                updated_data.append(item_row)
    
    # Write updated CSV file
    if updated_data:
        fieldnames = updated_data[0].keys()
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(updated_data)


def detect_deleted_orders(original_rows: List[Dict], sheet_edits: Dict[tuple, Dict[str, Any]]) -> List[tuple]:
    """Detect deleted orders/items by comparing original data with sheet edits."""
    if not original_rows or not sheet_edits:
        return []
    
    # Extract keys from original data
    original_keys = set()
    current_order_id = ""
    
    for row in original_rows:
        row_order_id = row.get("Order ID", "")
        item_number = row.get("Item #", "") or row.get("Item Number", "")
        
        order_id = row_order_id if row_order_id else current_order_id
        if row_order_id:
            current_order_id = row_order_id
            
        if order_id:
            original_keys.add((order_id, item_number))
    
    # Find keys that exist in original but not in sheet edits
    sheet_keys = set(sheet_edits.keys())
    deleted_keys = original_keys - sheet_keys
    
    return list(deleted_keys)


def remove_deleted_orders_from_files(deleted_keys: List[tuple], orders_dir: str) -> None:
    """Remove deleted orders/items from XML/CSV files."""
    if not deleted_keys or not os.path.exists(orders_dir):
        return
    
    # Group deletions by order ID
    orders_to_modify = defaultdict(set)
    orders_to_delete = set()
    
    for order_id, item_id in deleted_keys:
        if item_id == "":
            # Entire order deletion
            orders_to_delete.add(order_id)
        else:
            # Item deletion
            orders_to_modify[order_id].add(item_id)
    
    # Remove from XML file
    xml_file = os.path.join(orders_dir, 'orders.xml')
    if os.path.exists(xml_file):
        _remove_from_xml_file(xml_file, orders_to_delete, orders_to_modify)
    
    # Remove from CSV file
    csv_file = os.path.join(orders_dir, 'orders.csv')
    if os.path.exists(csv_file):
        _remove_from_csv_file(csv_file, orders_to_delete, orders_to_modify)


def _remove_from_xml_file(xml_path: str, orders_to_delete: set, orders_to_modify: Dict[str, set]) -> None:
    """Remove specified orders/items from XML file."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Remove entire orders
        for order_elem in root.findall("ORDER"):
            order_id = (order_elem.findtext("ORDERID") or "").strip()
            if order_id in orders_to_delete:
                root.remove(order_elem)
                continue
            
            # Remove specific items from orders
            if order_id in orders_to_modify:
                items_to_remove = orders_to_modify[order_id]
                for item_elem in order_elem.findall("ITEM"):
                    item_id = (item_elem.findtext("ITEMID") or "").strip()
                    if item_id in items_to_remove:
                        order_elem.remove(item_elem)
        
        # Write back to file
        ET.indent(tree, space="  ", level=0)
        tree.write(xml_path, encoding="utf-8", xml_declaration=True)
        
    except (ET.ParseError, Exception):
        pass


def _remove_from_csv_file(csv_path: str, orders_to_delete: set, orders_to_modify: Dict[str, set]) -> None:
    """Remove specified orders/items from CSV file."""
    try:
        # Load existing CSV data
        existing_data = []
        with open(csv_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            existing_data = list(reader)
        
        # Filter out deleted rows
        filtered_data = []
        current_order_id = ""
        
        for row in existing_data:
            row_order_id = row.get('Order ID', '').strip()
            item_number = row.get('Item Number', '').strip()
            
            # Track current order ID
            if row_order_id:
                current_order_id = row_order_id
            
            order_id = row_order_id if row_order_id else current_order_id
            
            # Skip entire orders marked for deletion
            if order_id in orders_to_delete:
                continue
            
            # Skip specific items marked for deletion
            if order_id in orders_to_modify and item_number in orders_to_modify[order_id]:
                continue
            
            filtered_data.append(row)
        
        # Write filtered data back to file
        if filtered_data:
            fieldnames = filtered_data[0].keys()
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(filtered_data)
        else:
            # Write empty file with headers
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Order ID', 'Seller', 'Order Date', 'Item Number', 'Item Description', 'Qty', 'Each', 'Condition'])
                writer.writeheader()
        
    except Exception:
        pass