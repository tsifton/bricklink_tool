import os
import csv
import xml.etree.ElementTree as ET
from collections import defaultdict
from colors import get_color_name
from config import ORDERS_DIR
from dataclasses import dataclass, field
from typing import List


@dataclass
class OrderItem:
    item_id: str
    item_type: str
    color_id: int
    qty: int
    price: float
    condition: str = ""
    description: str = ""
    unit_cost: float = 0.0  # added for list-based inventory calculations


@dataclass
class Order:
    order_id: str
    order_date: str
    seller: str
    order_total: float
    base_grand_total: float
    items: List[OrderItem] = field(default_factory=list)

    @classmethod
    def from_xml_element(cls, order_elem: ET.Element) -> "Order":
        order_id = (order_elem.findtext("ORDERID", "") or "").strip()
        order_date = (order_elem.findtext("ORDERDATE", "") or "").strip()
        seller = (order_elem.findtext("SELLER", "") or "").strip()
        order_total = float(order_elem.findtext("ORDERTOTAL", "0") or 0)
        base_total = float(order_elem.findtext("BASEGRANDTOTAL", "0") or 0)
        items: List[OrderItem] = []
        for it in order_elem.findall("ITEM"):
            items.append(OrderItem(
                item_id=(it.findtext("ITEMID", "") or "").strip(),
                item_type=(it.findtext("ITEMTYPE", "") or "").strip(),
                color_id=int(it.findtext("COLOR", "0") or 0),
                qty=int(it.findtext("QTY", "0") or 0),
                price=float(it.findtext("PRICE", "0") or 0),
                condition=(it.findtext("CONDITION", "") or "").strip(),
                description=(it.findtext("DESCRIPTION", "") or "").strip(),
            ))
        return cls(
            order_id=order_id,
            order_date=order_date,
            seller=seller,
            order_total=order_total,
            base_grand_total=base_total,
            items=items,
        )

    def to_xml_element(self) -> ET.Element:
        order_elem = ET.Element("ORDER")

        def _add(tag: str, text: str):
            el = ET.SubElement(order_elem, tag)
            el.text = text
            return el

        _add("ORDERID", self.order_id)
        _add("ORDERDATE", self.order_date)
        _add("SELLER", self.seller)
        _add("ORDERTOTAL", str(self.order_total))
        _add("BASEGRANDTOTAL", str(self.base_grand_total))

        for item in self.items:
            it = ET.SubElement(order_elem, "ITEM")
            ET.SubElement(it, "ITEMID").text = item.item_id
            ET.SubElement(it, "ITEMTYPE").text = item.item_type
            ET.SubElement(it, "COLOR").text = str(item.color_id)
            ET.SubElement(it, "QTY").text = str(item.qty)
            ET.SubElement(it, "PRICE").text = str(item.price)
            ET.SubElement(it, "CONDITION").text = item.condition
            ET.SubElement(it, "DESCRIPTION").text = item.description

        return order_elem


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

        # Parse Order objects, then iterate items
        for order_elem in root.findall("ORDER"):
            order = Order.from_xml_element(order_elem)
            order_id = order.order_id
            order_date = order.order_date
            seller = order.seller
            order_total = order.order_total
            base_total = order.base_grand_total
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

            # Build temporary list for fee distribution
            items_tmp = []
            for item in order.items:
                color_name = get_color_name(item.color_id)
                csv_desc = csv_descriptions.get(item.item_id, "")
                part_row = {
                    "Order ID": order_id,  # Include Order ID for proper key generation
                    "Item Number": item.item_id,
                    "Item Description": csv_desc,
                    "Color": color_name or item.item_type,
                    "Color ID": item.color_id,
                    "Item Type": item.item_type,
                    "Condition": item.condition,
                    "Qty": item.qty,
                    "Each": item.price,
                    "Total": item.qty * item.price
                }
                items_tmp.append((part_row, item.description, item.color_id))

            # Distribute fees across items and update inventory
            for part_row, seller_desc, color_id in items_tmp:
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

                csv_desc = csv_descriptions.get(part_row["Item Number"], "")
                if csv_desc:
                    clean_desc = csv_desc
                    if seller_desc and clean_desc.endswith(seller_desc):
                        clean_desc = clean_desc[: -len(seller_desc)].rstrip(" -")
                else:
                    clean_desc = seller_desc

                inventory[key].update({
                    'qty': new_qty,
                    'total_cost': new_total_cost,
                    'unit_cost': unit_cost,
                    'description': clean_desc,
                    'color_id': color_id if part_row["Item Type"] == "P" else None,
                    'color_name': part_row["Color"] if part_row["Item Type"] == "P" else None,
                })
                all_rows.append(part_row)

    # Return both inventory and all_rows if requested, otherwise just inventory
    return (inventory, all_rows) if return_rows else inventory
