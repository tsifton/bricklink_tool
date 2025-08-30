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


def load_orders():
    """
    Load BrickLink orders from XML and CSV.
    Returns:
      - inventory_list: list[OrderItem] with per-line unit_cost (fees allocated)
      - orders_list: list[Order] parsed from XML (item descriptions enriched from CSV)
    """
    inventory_list: List[OrderItem] = []
    orders_list: List[Order] = []
    csv_descriptions = {}  # Maps item_id -> BrickLink 'Item Description' from CSV

    # Check if merged files exist - if so, use only merged files to avoid duplication
    merged_csv_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.csv'))
    merged_xml_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.xml'))

    # --- Load CSV descriptions (full BrickLink descriptions) ---
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".csv"):
            continue
        if merged_csv_exists and fn != 'orders.csv':
            continue
        with open(os.path.join(ORDERS_DIR, fn), newline='', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                iid = (row.get('Item Number') or '').strip()
                desc = (row.get('Item Description') or '').strip()
                if iid and desc and iid not in csv_descriptions:
                    csv_descriptions[iid] = desc

    # --- Load XML orders, enrich descriptions, and build inventory list ---
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".xml"):
            continue
        if merged_xml_exists and fn != 'orders.xml':
            continue

        tree = ET.parse(os.path.join(ORDERS_DIR, fn))
        root = tree.getroot()

        for order_elem in root.findall("ORDER"):
            order = Order.from_xml_element(order_elem)

            # Enrich each item description from CSV (remove seller suffix when present)
            for it in order.items:
                seller_desc = it.description or ""
                csv_desc = csv_descriptions.get(it.item_id, "")
                if csv_desc:
                    clean_desc = csv_desc
                    if seller_desc and clean_desc.endswith(seller_desc):
                        clean_desc = clean_desc[: -len(seller_desc)].rstrip(" -")
                    it.description = clean_desc  # prefer CSV description
                # else keep seller description as-is

            # Allocate order-level fees proportionally and add per-line items to inventory_list
            order_total = order.order_total or 0.0
            base_total = order.base_grand_total or 0.0
            total_fees = base_total - order_total

            for it in order.items:
                line_total = float(it.qty) * float(it.price)
                share = (line_total / order_total) if order_total else 0.0
                fee_share = total_fees * share
                total_with_fees = line_total + fee_share
                unit_cost = (total_with_fees / it.qty) if it.qty else 0.0

                inventory_list.append(OrderItem(
                    item_id=it.item_id,
                    item_type=it.item_type,
                    color_id=it.color_id if it.item_type == "P" else 0,
                    qty=it.qty,
                    price=it.price,
                    condition=it.condition,
                    description=it.description,
                    unit_cost=unit_cost
                ))

            orders_list.append(order)

    return inventory_list, orders_list