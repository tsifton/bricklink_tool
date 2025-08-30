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
                    'parsed_date': parsed_date,
                    'order_date': order_date
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
    """
    if not os.path.exists(ORDERS_DIR):
        return
        
    all_rows = []
    headers = None
    orders_by_id = {}  # Track order dates for sorting
    
    # Process all CSV files
    for filename in os.listdir(ORDERS_DIR):
        if not filename.endswith('.csv') or filename == 'orders.csv':
            continue
            
        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            with open(filepath, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                if headers is None:
                    headers = reader.fieldnames
                
                for row in reader:
                    order_id = row.get('Order ID', '').strip()
                    order_date = row.get('Order Date', '').strip()
                    
                    if not order_id:
                        continue
                        
                    parsed_date = parse_order_date(order_date)
                    
                    # Create a unique key for this row (order_id + item details for uniqueness)
                    item_key = f"{order_id}_{row.get('Item Number', '')}_{row.get('Color ID', '')}_{row.get('Qty', '')}"
                    
                    # Track order dates and store rows
                    if order_id not in orders_by_id:
                        orders_by_id[order_id] = parsed_date
                    else:
                        # Keep the newer date for this order
                        if parsed_date > orders_by_id[order_id]:
                            orders_by_id[order_id] = parsed_date
                    
                    # Add row with sorting info
                    row['_parsed_date'] = parsed_date
                    row['_item_key'] = item_key
                    all_rows.append(row)
                    
        except Exception as e:
            print(f"Warning: Could not parse CSV file {filename}: {e}")
            continue
    
    if not all_rows:
        return
    
    # Remove duplicates - keep rows from orders with newer dates
    unique_rows = {}
    for row in all_rows:
        order_id = row.get('Order ID', '').strip()
        item_key = row['_item_key']
        
        if item_key not in unique_rows:
            unique_rows[item_key] = row
        else:
            # Keep the row from the order with the newer date
            existing_date = unique_rows[item_key]['_parsed_date']
            current_date = row['_parsed_date']
            if current_date > existing_date:
                unique_rows[item_key] = row
    
    # Sort by order date (newest first), then by order ID for consistency
    final_rows = sorted(
        unique_rows.values(),
        key=lambda x: (x['_parsed_date'], x.get('Order ID', '')),
        reverse=True
    )
    
    # Remove temporary sorting fields
    for row in final_rows:
        row.pop('_parsed_date', None)
        row.pop('_item_key', None)
    
    # Write merged CSV file
    if headers and final_rows:
        output_path = os.path.join(ORDERS_DIR, 'orders.csv')
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(final_rows)
        
        print(f"Merged {len(final_rows)} unique order items into orders.csv")