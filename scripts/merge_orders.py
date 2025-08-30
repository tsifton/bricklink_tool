"""
Merges multiple orders XML/CSV into canonical orders.xml/orders.csv 
(sorted by date, unique by Order ID).
"""

import os
import csv
import xml.etree.ElementTree as ET
from datetime import datetime
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from orders import Order, OrderItem  # use shared classes

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

    orders_by_id: Dict[str, Order] = {}

    # Process all XML files, except the merged output
    for filename in os.listdir(ORDERS_DIR):
        if not filename.endswith('.xml') or filename == 'orders.xml':
            continue

        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            for order_elem in root.findall('ORDER'):
                order = Order.from_xml_element(order_elem)
                if not order.order_id:
                    continue
                existing = orders_by_id.get(order.order_id)
                if not existing or parse_order_date(order.order_date) > parse_order_date(existing.order_date):
                    orders_by_id[order.order_id] = order
        except ET.ParseError as e:
            print(f"Warning: Could not parse XML file {filename}: {e}")
            continue

    if not orders_by_id:
        return

    # Sort orders by datetime desc
    sorted_orders = sorted(
        orders_by_id.values(),
        key=lambda o: parse_order_date(o.order_date),
        reverse=True
    )

    # Build XML tree
    root = ET.Element('ORDERS')
    for order in sorted_orders:
        root.append(order.to_xml_element())

    # Write merged XML file
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    output_path = os.path.join(ORDERS_DIR, 'orders.xml')
    tree.write(output_path, encoding='utf-8', xml_declaration=True)

    print(f"Merged {len(sorted_orders)} unique orders into orders.xml")


@dataclass
class CsvOrder:
    order_id: str
    date: datetime
    header: Optional[Dict[str, Any]] = None
    items: Dict[str, Dict[str, Any]] = field(default_factory=dict)  # key -> {'row': row, 'date': date}


def merge_csv():
    """
    Merge all CSV order files in ORDERS_DIR into a single orders.csv file.
    Orders are deduplicated by Order ID and sorted by date (newest first).
    Item rows (order line items) are preserved under their corresponding order.
    """
    if not os.path.exists(ORDERS_DIR):
        return

    headers = None
    orders_map: Dict[str, CsvOrder] = {}

    for filename in os.listdir(ORDERS_DIR):
        if not filename.endswith('.csv') or filename == 'orders.csv':
            continue

        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            with open(filepath, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                if headers is None:
                    headers = reader.fieldnames

                current_order_id: Optional[str] = None
                current_order_date: Optional[datetime] = None

                for row in reader:
                    # Skip empty rows
                    if not any((v or '').strip() for v in row.values()):
                        continue

                    raw_order_id = (row.get('Order ID') or '').strip()
                    if raw_order_id:
                        # Treat rows with Order ID as a header for that order (BL export compatible)
                        current_order_id = raw_order_id
                        current_order_date = parse_order_date((row.get('Order Date') or '').strip())
                        existing = orders_map.get(current_order_id)
                        if existing is None or (current_order_date or datetime.min) > existing.date:
                            # Keep the latest header; retain any existing items
                            items = existing.items if existing else {}
                            orders_map[current_order_id] = CsvOrder(
                                order_id=current_order_id,
                                date=current_order_date or datetime.min,
                                header=row,
                                items=items
                            )
                        # else keep existing newer header
                    else:
                        # Item line: associate with the current order
                        if not current_order_id:
                            continue
                        entry = orders_map.setdefault(
                            current_order_id,
                            CsvOrder(order_id=current_order_id, date=current_order_date or datetime.min)
                        )
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
                        existing_item = entry.items.get(item_key)
                        item_date = current_order_date or entry.date
                        if not existing_item or item_date >= existing_item['date']:
                            entry.items[item_key] = {'row': row, 'date': item_date}

        except Exception as e:
            print(f"Warning: Could not parse CSV file {filename}: {e}")
            continue

    if not orders_map or not headers:
        return

    # Sort orders by date desc then by Order ID for stability
    sorted_order_ids = sorted(orders_map.keys(), key=lambda oid: (orders_map[oid].date, oid), reverse=True)

    final_rows: List[Dict[str, Any]] = []
    total_items = 0
    for oid in sorted_order_ids:
        entry = orders_map[oid]
        if not entry.header:
            # Skip orders without a header row
            continue
        # Normalize header row to headers
        header_row = {k: entry.header.get(k, '') for k in headers}
        final_rows.append(header_row)
        # Append items in insertion order
        for it in entry.items.values():
            row = it['row']
            row_clean = {k: row.get(k, '') for k in headers}
            final_rows.append(row_clean)
            total_items += 1

    # Write merged CSV
    output_path = os.path.join(ORDERS_DIR, 'orders.csv')
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(final_rows)

    print(f"Merged {len(sorted_order_ids)} orders with {total_items} items into orders.csv")