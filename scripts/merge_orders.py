"""
Merges multiple orders XML/CSV into canonical orders.xml/orders.csv 
(sorted by date, unique by Order ID).
"""

import os
import csv
import xml.etree.ElementTree as ET
from datetime import datetime

# Import ORDERS_DIR from config if available, otherwise use default
try:
    from config import ORDERS_DIR
except ImportError:
    ORDERS_DIR = "orders"


def parse_order_date(date_str):
    """
    Parse order date string and return datetime object for sorting.
    Handles various date formats from BrickLink.
    """
    if not date_str:
        return datetime.min
    
    # Try common BrickLink date formats
    formats = [
        "%Y-%m-%dT%H:%M:%S.%fZ",  # 2024-08-15T10:30:00.000Z
        "%Y-%m-%dT%H:%M:%SZ",     # 2024-08-15T10:30:00Z
        "%Y-%m-%d %H:%M:%S",      # 2024-08-15 10:30:00
        "%Y-%m-%d",               # 2024-08-15
        "%m/%d/%Y %H:%M:%S",      # 08/15/2024 10:30:00
        "%m/%d/%Y %H:%M",         # 08/15/2024 10:30
        "%m/%d/%Y",               # 08/15/2024
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    # If all formats fail, return min date to put at end when sorted descending
    return datetime.min


def merge_xml():
    """
    Merge all XML order files in ORDERS_DIR into a single orders.xml file.
    Orders are deduplicated by Order ID and sorted by date (newest first).
    """
    if not os.path.exists(ORDERS_DIR):
        return
        
    orders_by_id = {}  # Dict to store orders keyed by order_id for deduplication
    
    # Process all XML files
    for filename in os.listdir(ORDERS_DIR):
        if not filename.endswith('.xml') or filename == 'orders.xml':
            continue
            
        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            
            for order_elem in root.findall('ORDER'):
                order_id = order_elem.findtext('ORDERID', '').strip()
                order_date = order_elem.findtext('ORDERDATE', '').strip()
                
                if not order_id:
                    continue
                
                parsed_date = parse_order_date(order_date)
                
                # If we already have this order, keep the newer one
                if order_id in orders_by_id:
                    existing_date = orders_by_id[order_id]['parsed_date']
                    if parsed_date <= existing_date:
                        continue
                
                # Store the order element and parsed date for sorting
                orders_by_id[order_id] = {
                    'element': order_elem,
                    'parsed_date': parsed_date
                }
                
        except ET.ParseError as e:
            print(f"Warning: Could not parse XML file {filename}: {e}")
            continue
    
    if not orders_by_id:
        return
    
    # Sort orders by date (newest first)
    sorted_orders = sorted(
        orders_by_id.values(), 
        key=lambda x: x['parsed_date'], 
        reverse=True
    )
    
    # Create new XML structure
    root = ET.Element('ORDERS')
    for order_data in sorted_orders:
        root.append(order_data['element'])
    
    # Write merged XML file
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)  # Format with indentation
    output_path = os.path.join(ORDERS_DIR, 'orders.xml')
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    
    print(f"Merged {len(sorted_orders)} unique orders into orders.xml")


def merge_csv():
    """
    Merge all CSV order files in ORDERS_DIR into a single orders.csv file.
    Orders are deduplicated by Order ID and sorted by date (newest first).
    Item rows (order line items) are preserved under their corresponding order.
    """
    if not os.path.exists(ORDERS_DIR):
        return

    headers = None

    # orders_map: order_id -> { 'date': datetime, 'header': dict, 'items': Ordered insertion dict-like }
    orders_map = {}

    for filename in os.listdir(ORDERS_DIR):
        if not filename.endswith('.csv') or filename == 'orders.csv':
            continue

        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            with open(filepath, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                if headers is None:
                    headers = reader.fieldnames

                current_order_id = None
                current_order_date = None

                for row in reader:
                    # Skip rows that are entirely empty
                    if not any((v or '').strip() for v in row.values()):
                        continue

                    raw_order_id = (row.get('Order ID') or '').strip()

                    if raw_order_id:
                        # This is an order-level header row
                        current_order_id = raw_order_id
                        current_order_date = parse_order_date((row.get('Order Date') or '').strip())

                        existing = orders_map.get(current_order_id)
                        if existing is None or current_order_date > existing['date']:
                            # Keep latest header, preserve any already-collected items
                            items = existing['items'] if existing else {}
                            orders_map[current_order_id] = {
                                'date': current_order_date,
                                'header': row,
                                'items': items
                            }
                        else:
                            # Older header encountered; keep existing header and items
                            # Still maintain current_order_id/date for following item rows
                            pass

                    else:
                        # This is an item row; associate with the most recent order header
                        if not current_order_id:
                            # Orphaned item row; cannot associate
                            continue

                        entry = orders_map.setdefault(
                            current_order_id,
                            {'date': current_order_date or datetime.min, 'header': None, 'items': {}}
                        )

                        # Prefer Inv ID for deduplication; fallback to composite key
                        inv_id = (row.get('Inv ID') or '').strip()
                        if inv_id:
                            item_key = inv_id
                        else:
                            item_key = "|".join([
                                (row.get('Item Type') or '').strip(),
                                (row.get('Item Number') or '').strip(),
                                (row.get('Item Description') or '').strip(),
                                (row.get('Sub-Condition') or '').strip(),
                                (row.get('Qty') or '').strip(),
                                (row.get('Each') or '').strip(),
                                (row.get('Total') or '').strip(),
                                (row.get('Weight') or '').strip(),
                                (row.get('Batch') or '').strip(),
                                (row.get('Batch Date') or '').strip(),
                                (row.get('Condition') or '').strip(),
                            ])

                        existing_item = entry['items'].get(item_key)
                        item_date = current_order_date or entry['date'] or datetime.min

                        # Keep the item from the newer version of the order
                        if not existing_item or item_date >= existing_item['date']:
                            entry['items'][item_key] = {'row': row, 'date': item_date}

        except Exception as e:
            print(f"Warning: Could not parse CSV file {filename}: {e}")
            continue

    if not orders_map or not headers:
        return

    # Build final rows: order header followed by its items, orders sorted by date desc then by Order ID
    final_rows = []
    sorted_order_ids = sorted(
        orders_map.keys(),
        key=lambda oid: (orders_map[oid]['date'], oid),
        reverse=True
    )

    total_items = 0
    for oid in sorted_order_ids:
        entry = orders_map[oid]
        if not entry.get('header'):
            # Skip orders that have items but no header row
            continue

        # Clean header to include only keys in headers
        header_row = {k: entry['header'].get(k, '') for k in headers}
        final_rows.append(header_row)

        # Append items in insertion order
        for item in entry['items'].values():
            row = item['row']
            row_clean = {k: row.get(k, '') for k in headers}
            final_rows.append(row_clean)
            total_items += 1

    # Write merged CSV file
    output_path = os.path.join(ORDERS_DIR, 'orders.csv')
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(final_rows)

    print(f"Merged {len(sorted_order_ids)} orders with {total_items} items into orders.csv")