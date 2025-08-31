import os
import re
import csv
import xml.etree.ElementTree as ET
from config import ORDERS_DIR
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
from colors import get_color_name


@dataclass
class OrderItem:
    item_id: str
    item_type: str
    color_id: int
    qty: int
    price: float
    condition: str = ""
    description: str = ""
    unit_cost: float = 0.0
    lot_id: str = ""
    clean_description: str = ""
    color_name: str = ""


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
        """Create Order from XML element."""
        def get_text(elem: ET.Element, tag: str, default: str = "") -> str:
            return (elem.findtext(tag, default) or "").strip()
        
        def get_float(elem: ET.Element, tag: str, default: float = 0.0) -> float:
            try:
                return float(elem.findtext(tag, str(default)) or default)
            except ValueError:
                return default
        
        def get_int(elem: ET.Element, tag: str, default: int = 0) -> int:
            try:
                return int(elem.findtext(tag, str(default)) or default)
            except ValueError:
                return default

        items = []
        for item_elem in order_elem.findall("ITEM"):
            item_type = get_text(item_elem, "ITEMTYPE")
            color_id = get_int(item_elem, "COLOR")
            
            # Determine color name
            color_name = (get_color_name(color_id) if item_type == "P" and color_id 
                         else item_type or "")
            
            items.append(OrderItem(
                item_id=get_text(item_elem, "ITEMID"),
                item_type=item_type,
                color_id=color_id,
                qty=get_int(item_elem, "QTY"),
                price=get_float(item_elem, "PRICE"),
                condition=get_text(item_elem, "CONDITION"),
                description=get_text(item_elem, "DESCRIPTION"),
                lot_id=get_text(item_elem, "LOTID"),
                color_name=color_name,
            ))

        return cls(
            order_id=get_text(order_elem, "ORDERID"),
            order_date=get_text(order_elem, "ORDERDATE"),
            seller=get_text(order_elem, "SELLER"),
            order_total=get_float(order_elem, "ORDERTOTAL"),
            base_grand_total=get_float(order_elem, "BASEGRANDTOTAL"),
            items=items,
        )

    def to_xml_element(self) -> ET.Element:
        """Convert Order to XML element with minimal fields."""
        order_elem = ET.Element("ORDER")
        
        # Only include ORDERID for minimal XML
        ET.SubElement(order_elem, "ORDERID").text = self.order_id

        for item in self.items:
            # Only include items that have LOTID
            if item.lot_id:
                item_elem = ET.SubElement(order_elem, "ITEM")
                
                # Always include LOTID
                ET.SubElement(item_elem, "LOTID").text = item.lot_id
                
                # Include COLOR if present (non-zero)
                if item.color_id:
                    ET.SubElement(item_elem, "COLOR").text = str(item.color_id)
                
                # Include DESCRIPTION (seller notes) if present
                if item.description and item.description.strip():
                    ET.SubElement(item_elem, "DESCRIPTION").text = item.description

        return order_elem


# --- helpers ---

def _parse_money(val: Optional[str]) -> float:
    """Parse money string to float."""
    if not val:
        return 0.0
    cleaned = val.strip().replace("$", "").replace(",", "")
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0


def _parse_int(val: Optional[str]) -> int:
    """Parse string to int."""
    if not val:
        return 0
    try:
        return int(float(val.strip()))
    except ValueError:
        return 0


def _map_item_type(label: str) -> str:
    """Map item type label to code."""
    type_map = {"minifigure": "M", "part": "P", "set": "S"}
    return type_map.get((label or "").strip().lower(), 
                       (label or "").strip()[:1].upper() or "P")


def _normalize_spaces(text: str) -> str:
    """Normalize whitespace in text."""
    return re.sub(r"\s+", " ", (text or "").strip())


def _build_xml_indexes() -> Tuple[Dict[Tuple[str, str], int], Dict[Tuple[str, str], str]]:
    """Build color and seller note indexes from XML files."""
    color_index = {}
    seller_note_index = {}

    if not os.path.exists(ORDERS_DIR):
        return color_index, seller_note_index

    # Prefer merged file if it exists
    xml_files = ['orders.xml'] if os.path.exists(os.path.join(ORDERS_DIR, 'orders.xml')) else [
        f for f in os.listdir(ORDERS_DIR) if f.endswith('.xml')
    ]

    for filename in xml_files:
        filepath = os.path.join(ORDERS_DIR, filename)
        try:
            tree = ET.parse(filepath)
            for order_elem in tree.getroot().findall("ORDER"):
                order_id = (order_elem.findtext("ORDERID") or "").strip()
                if not order_id:
                    continue
                    
                for item_elem in order_elem.findall("ITEM"):
                    lot_id = (item_elem.findtext("LOTID") or "").strip()
                    if not lot_id:
                        continue
                        
                    key = (order_id, lot_id)
                    
                    # Store color ID
                    try:
                        color_index[key] = int(item_elem.findtext("COLOR", "0") or 0)
                    except ValueError:
                        color_index[key] = 0
                        
                    # Store seller note
                    note = (item_elem.findtext("DESCRIPTION") or "").strip()
                    if note:
                        seller_note_index[key] = note
                        
        except (ET.ParseError, Exception):
            continue

    return color_index, seller_note_index


def write_minimal_orders_xml(orders: List[Order], output_path: str) -> None:
    """Write minimal XML with only order IDs, lot IDs, colors, and descriptions."""
    root = ET.Element("ORDERS")
    for order in orders:
        order_elem = ET.SubElement(root, "ORDER")
        ET.SubElement(order_elem, "ORDERID").text = order.order_id
        for item in order.items:
            if item.lot_id:
                item_elem = ET.SubElement(order_elem, "ITEM")
                ET.SubElement(item_elem, "LOTID").text = item.lot_id
                
                # Include COLOR if present (non-zero)
                if item.color_id:
                    ET.SubElement(item_elem, "COLOR").text = str(item.color_id)
                
                # Include DESCRIPTION (seller notes) if present
                if item.description and item.description.strip():
                    ET.SubElement(item_elem, "DESCRIPTION").text = item.description
                
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)


def load_orders() -> Tuple[List[OrderItem], List[Order]]:
    """Load orders from CSV files with XML color resolution."""
    inventory_list = []
    orders_list = []

    if not os.path.exists(ORDERS_DIR):
        return inventory_list, orders_list

    # Build XML indexes
    xml_color_index, xml_seller_note_index = _build_xml_indexes()
    
    # Prefer merged file if it exists
    csv_files = ['orders.csv'] if os.path.exists(os.path.join(ORDERS_DIR, 'orders.csv')) else [
        f for f in os.listdir(ORDERS_DIR) if f.endswith('.csv')
    ]

    current_order = None
    
    for filename in csv_files:
        filepath = os.path.join(ORDERS_DIR, filename)
        with open(filepath, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                order_id = (row.get("Order ID") or "").strip()
                item_number = (row.get("Item Number") or "").strip()

                if order_id:
                    # Save previous order
                    if current_order and current_order.items:
                        orders_list.append(current_order)
                        
                    # Create new order
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
                    )
                    continue

                # Process item rows
                if not current_order or not item_number:
                    continue

                # Extract item data
                type_code = _map_item_type(row.get("Item Type", ""))
                lot_id = (row.get("Inv ID") or "").strip()
                csv_desc_raw = row.get("Item Description", "")
                csv_desc = _normalize_spaces(csv_desc_raw)
                
                # Get color ID from XML
                color_id = 0
                if type_code == "P" and lot_id:
                    color_id = xml_color_index.get((current_order.order_id, lot_id), 0)

                # Determine color name
                color_name = (get_color_name(color_id) if type_code == "P" and color_id 
                             else type_code)

                # Clean description by removing seller note
                seller_note = xml_seller_note_index.get((current_order.order_id, lot_id), "")
                clean_desc = csv_desc
                if seller_note and csv_desc_raw.rstrip().endswith(seller_note.strip()):
                    stripped = csv_desc_raw.rstrip()[:-len(seller_note.strip())].rstrip(" -")
                    clean_desc = _normalize_spaces(stripped)

                # Calculate unit cost with proportional fees
                qty = _parse_int(row.get("Qty"))
                price = _parse_money(row.get("Each"))
                line_total = qty * price
                
                if current_order.order_total:
                    fee_share = ((current_order.base_grand_total - current_order.order_total) * 
                               line_total / current_order.order_total)
                    unit_cost = (line_total + fee_share) / qty if qty else 0.0
                else:
                    unit_cost = price

                item = OrderItem(
                    item_id=item_number,
                    item_type=type_code,
                    color_id=color_id if type_code == "P" else 0,
                    qty=qty,
                    price=price,
                    condition=(row.get("Condition") or "").strip()[0],
                    description=csv_desc,
                    clean_description=clean_desc,
                    unit_cost=unit_cost,
                    lot_id=lot_id,
                    color_name=color_name,
                )
                
                current_order.items.append(item)
                inventory_list.append(item)

    # Add final order
    if current_order and current_order.items:
        orders_list.append(current_order)

    return inventory_list, orders_list