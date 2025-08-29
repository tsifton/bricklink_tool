import os
import csv
import xml.etree.ElementTree as ET
from collections import defaultdict
from colors import get_color_name
from config import ORDERS_DIR


def load_orders(return_rows=False):
    """
    Load BrickLink orders from XML and CSV.
    Returns inventory and (optionally) rows for Google Sheets.
    """
    # Inventory is a dict keyed by (item_id, color_id) or (item_id, None)
    # Each value is a dict with qty, total_cost, unit_cost, description, color_id, color_name
    inventory = defaultdict(lambda: {
        'qty': 0,
        'total_cost': 0.0,
        'unit_cost': 0.0,
        'description': '',
        'color_id': None,
        'color_name': None,
    })
    all_rows = []  # List of rows for Google Sheets
    csv_descriptions = {}  # Maps item_id to description from CSV

    # Check if merged files exist - if so, use only merged files to avoid duplication
    merged_csv_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.csv'))
    merged_xml_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.xml'))

    # --- Load CSV descriptions (full BrickLink descriptions) ---
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".csv"):
            continue
        # Skip individual CSV files if merged CSV exists
        if merged_csv_exists and fn != 'orders.csv':
            continue
        
        with open(os.path.join(ORDERS_DIR, fn), newline='', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                iid = row['Item Number'].strip()
                desc = row.get('Item Description', '').strip()
                # Only use the first description found for each item_id
                if iid and desc and iid not in csv_descriptions:
                    csv_descriptions[iid] = desc

    # --- Load XML orders ---
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".xml"):
            continue
        # Skip individual XML files if merged XML exists
        if merged_xml_exists and fn != 'orders.xml':
            continue
        tree = ET.parse(os.path.join(ORDERS_DIR, fn))
        root = tree.getroot()
        for order in root.findall("ORDER"):
            # Extract order-level information
            order_id = order.findtext("ORDERID", "").strip()
            order_date = order.findtext("ORDERDATE", "").strip()
            seller = order.findtext("SELLER", "").strip()
            order_total = float(order.findtext("ORDERTOTAL", "0") or 0)
            base_total = float(order.findtext("BASEGRANDTOTAL", "0") or 0)
            total_fees = base_total - order_total

            # Add a header row for this order to the output rows
            all_rows.append({
                "Order ID": order_id,
                "Order Date": order_date,
                "Seller": seller,
                "Shipping": "",
                "Add Chrg 1": "",
                "Order Total": order_total,
                "Base Grand Total": base_total,
                "Total Lots": "",
                "Total Items": "",
                "Tracking No": ""
            })

            items = []
            for item in order.findall("ITEM"):
                # Extract item-level information
                item_id = item.findtext("ITEMID", "").strip()
                item_type = item.findtext("ITEMTYPE", "").strip()
                color_id = int(item.findtext("COLOR", "0") or 0)
                color_name = get_color_name(color_id)
                qty = int(item.findtext("QTY", "0") or 0)
                price = float(item.findtext("PRICE", "0") or 0)
                seller_desc = item.findtext("DESCRIPTION", "").strip()
                total = price * qty
                # Use CSV description for Google Sheets (orders tab)
                csv_desc = csv_descriptions.get(item_id, "")
                part_row = {
                    "Item Number": item_id,
                    "Item Description": csv_desc,
                    "Color": color_name or item_type,
                    "Color ID": color_id,
                    "Item Type": item_type,
                    "Condition": item.findtext("CONDITION", "").strip(),
                    "Qty": qty,
                    "Each": price,
                    "Total": total
                }
                items.append((part_row, seller_desc, color_id))

            # Distribute fees across items and update inventory
            for part_row, seller_desc, color_id in items:
                # Calculate proportional share of fees for this item
                share = (part_row["Total"] / order_total) if order_total else 0
                fee_share = total_fees * share
                total_with_fees = part_row["Total"] + fee_share
                # Inventory key: sets/minifigs ignore color, parts use color
                key = (part_row["Item Number"], None) if part_row["Item Type"] in ("S", "M") \
                    else (part_row["Item Number"], part_row.get("Color ID", 0))
                prev = inventory[key]
                new_qty = prev['qty'] + part_row["Qty"]
                new_total_cost = prev['total_cost'] + total_with_fees
                unit_cost = (new_total_cost / new_qty) if new_qty else 0
                # Clean description for inventory (use CSV if available, fallback to seller description)
                csv_desc = csv_descriptions.get(part_row["Item Number"], "")
                if csv_desc:
                    # Use CSV description and clean it by removing seller note if present
                    clean_desc = csv_desc
                    if seller_desc and clean_desc.endswith(seller_desc):
                        clean_desc = clean_desc[: -len(seller_desc)].rstrip(" -")
                else:
                    # No CSV description available, use seller description as fallback
                    clean_desc = seller_desc
                # Update inventory with new values
                inventory[key].update({
                    'qty': new_qty,
                    'total_cost': new_total_cost,
                    'unit_cost': unit_cost,
                    'description': clean_desc,
                    'color_id': color_id if part_row["Item Type"] == "P" else None,
                    'color_name': part_row["Color"] if part_row["Item Type"] == "P" else None,
                })
                # Add item row to output rows for Google Sheets
                all_rows.append(part_row)

    # Return both inventory and all_rows if requested, otherwise just inventory
    return (inventory, all_rows) if return_rows else inventory
