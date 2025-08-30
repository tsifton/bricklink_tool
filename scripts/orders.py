import os
import csv
import xml.etree.ElementTree as ET
from collections import defaultdict
from colors import get_color_name
from config import ORDERS_DIR
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
import re


@dataclass
class OrderItem:
    item_id: str
    item_type: str
    color_id: int
    qty: int
    price: float
    condition: str = ""
    description: str = ""
    unit_cost: float = 0.0  # per-line unit cost incl. proportional fees
    lot_id: str = ""        # CSV Inv ID; used to match XML LOTID
    clean_description: str = ""  # CSV description with seller note suffix removed


@dataclass
class Order:
    order_id: str
    order_date: str
    seller: str
    order_total: float
    base_grand_total: float
    shipping: float = 0.0
    add_chrg_1: float = 0.0
    total_lots: int = 0
    total_items: int = 0
    tracking_no: str = ""
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
                # LOTID may exist in real BL XML; keep if present
                lot_id=(it.findtext("LOTID", "") or "").strip(),
            ))
        return cls(
            order_id=order_id,
            order_date=order_date,
            seller=seller,
            order_total=order_total,
            base_grand_total=base_total,
            # XML path doesn't provide the extra CSV-only fields; keep defaults
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
            # Preserve LOTID when present (no harm for merge output)
            if item.lot_id:
                ET.SubElement(it, "LOTID").text = item.lot_id

        return order_elem


# --- helpers ---

def _parse_money(val: Optional[str]) -> float:
    s = (val or "").strip()
    # Remove $ and commas
    s = s.replace("$", "").replace(",", "")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def _parse_int(val: Optional[str]) -> int:
    s = (val or "").strip()
    try:
        return int(float(s)) if s else 0
    except ValueError:
        return 0


def _map_item_type(label: str) -> str:
    lab = (label or "").strip().lower()
    if lab == "minifigure":
        return "M"
    if lab == "part":
        return "P"
    if lab == "set":
        return "S"
    # Fallback: first letter upper, default 'P' for unknowns to keep color logic consistent
    return lab[:1].upper() or "P"


def _normalize_spaces(text: str) -> str:
    """
    Collapse consecutive whitespace characters into single spaces and trim ends.
    """
    return re.sub(r"\s+", " ", (text or "").strip())


def _build_xml_indexes() -> Tuple[Dict[Tuple[str, str], int], Dict[Tuple[str, str], str]]:
    """
    Build both indexes from XML in a single pass:
      - color_index: (ORDERID, LOTID) -> COLOR
      - seller_note_index: (ORDERID, LOTID) -> DESCRIPTION (seller note)
    """
    color_index: Dict[Tuple[str, str], int] = {}
    seller_note_index: Dict[Tuple[str, str], str] = {}

    if not os.path.exists(ORDERS_DIR):
        return color_index, seller_note_index

    merged_xml_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.xml'))
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".xml"):
            continue
        if merged_xml_exists and fn != "orders.xml":
            continue
        try:
            tree = ET.parse(os.path.join(ORDERS_DIR, fn))
            root = tree.getroot()
            for order_elem in root.findall("ORDER"):
                oid = (order_elem.findtext("ORDERID") or "").strip()
                if not oid:
                    continue
                for it in order_elem.findall("ITEM"):
                    lot_id = (it.findtext("LOTID") or "").strip()
                    if not lot_id:
                        continue
                    # color id
                    try:
                        color_id = int(it.findtext("COLOR", "0") or 0)
                    except ValueError:
                        color_id = 0
                    color_index[(oid, lot_id)] = color_id
                    # seller note/description
                    note = (it.findtext("DESCRIPTION") or "").strip()
                    if note:
                        seller_note_index[(oid, lot_id)] = note
        except ET.ParseError:
            continue
        except Exception:
            continue

    return color_index, seller_note_index


def write_minimal_orders_xml(orders: List[Order], output_path: str):
    """
    Write a minimal XML capturing only ORDERID and ITEM LOTID for edited orders.
    This intentionally excludes all other fields.
    """
    root = ET.Element("ORDERS")
    for order in orders:
        o = ET.SubElement(root, "ORDER")
        ET.SubElement(o, "ORDERID").text = order.order_id
        for item in order.items:
            if not item.lot_id:
                continue
            it = ET.SubElement(o, "ITEM")
            ET.SubElement(it, "LOTID").text = item.lot_id
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)


def load_orders():
    """
    Load BrickLink orders with CSV-first strategy; use XML only to resolve color IDs.
    Returns:
      - inventory_list: list[OrderItem] with per-line unit_cost (fees allocated)
      - orders_list: list[Order] built from CSV with color_id filled from XML by (Order ID, Inv ID)
    """
    inventory_list: List[OrderItem] = []
    orders_list: List[Order] = []

    if not os.path.exists(ORDERS_DIR):
        return inventory_list, orders_list

    merged_csv_exists = os.path.exists(os.path.join(ORDERS_DIR, 'orders.csv'))

    # Build indices from XML in a single pass
    xml_color_index, xml_seller_note_index = _build_xml_indexes()

    current_order: Optional[Order] = None
    # Parse CSV, build orders and items
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".csv"):
            continue
        if merged_csv_exists and fn != "orders.csv":
            continue

        with open(os.path.join(ORDERS_DIR, fn), newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                order_id = (row.get("Order ID") or "").strip()
                item_number = (row.get("Item Number") or "").strip()

                if order_id:
                    # Close previous order if switching
                    if current_order and current_order.items:
                        orders_list.append(current_order)
                    # Start new order using CSV header fields
                    current_order = Order(
                        order_id=order_id,
                        order_date=(row.get("Order Date") or "").strip(),
                        seller=(row.get("Seller") or "").strip(),
                        shipping=_parse_money(row.get("Shipping")),
                        add_chrg_1=_parse_money(row.get("Add Chrg 1")),
                        order_total=_parse_money(row.get("Order Total")),
                        base_grand_total=_parse_money(row.get("Base Grand Total")),
                        total_lots=_parse_int(row.get("Total Lots")),
                        total_items=_parse_int(row.get("Total Items")),
                        tracking_no=(row.get("Tracking No") or "").strip(),
                        items=[],
                    )
                    continue

                # Item rows: must have Item Number
                if not current_order or not item_number:
                    continue

                # Extract item fields from CSV
                cond = (row.get("Condition") or "").strip()
                csv_desc_raw = (row.get("Item Description") or "")
                csv_desc = _normalize_spaces(csv_desc_raw)  # normalized full description for Orders sheet
                qty = _parse_int(row.get("Qty"))
                price = _parse_money(row.get("Each"))
                item_type_label = (row.get("Item Type") or "").strip()
                type_code = _map_item_type(item_type_label)
                lot_id = (row.get("Inv ID") or "").strip()

                color_id = 0
                if type_code == "P" and lot_id:
                    color_id = xml_color_index.get((current_order.order_id, lot_id), 0)

                # Clean description for inventory by removing seller note suffix from CSV desc
                seller_note = xml_seller_note_index.get((current_order.order_id, lot_id), "")
                if seller_note and csv_desc.endswith(seller_note):
                    clean_desc = csv_desc[: -len(seller_note)].rstrip(" -")
                else:
                    clean_desc = csv_desc

                # Compute unit_cost with proportional fees based on order header totals
                order_total = current_order.order_total or 0.0
                base_total = current_order.base_grand_total or 0.0
                total_fees = base_total - order_total
                line_total = float(qty) * float(price)
                share = (line_total / order_total) if order_total else 0.0
                fee_share = total_fees * share
                total_with_fees = line_total + fee_share
                unit_cost = (total_with_fees / qty) if qty else 0.0

                item = OrderItem(
                    item_id=item_number,
                    item_type=type_code,
                    color_id=color_id if type_code == "P" else 0,
                    qty=qty,
                    price=price,
                    condition=cond,
                    description=csv_desc,           # full CSV description for Orders sheet
                    clean_description=clean_desc,   # cleaned for inventory sheets
                    unit_cost=unit_cost,
                    lot_id=lot_id,
                )
                current_order.items.append(item)
                # Add to inventory list (flat list of per-line items)
                inventory_list.append(item)

    # Append the last order if present
    if current_order and current_order.items:
        orders_list.append(current_order)

    return inventory_list, orders_list