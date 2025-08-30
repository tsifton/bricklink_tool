import os
import xml.etree.ElementTree as ET
from config import WANTED_LISTS_DIR
from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class RequiredItem:
    """
    Represents a per-unit requirement in a wanted list (a 'sellable unit').
    qty is the number of this item needed to build one unit.
    color_id applies only to parts (Item Type 'P'); otherwise None.
    """
    item_id: str
    item_type: str  # 'S', 'M', 'P', etc.
    qty: int
    color_id: Optional[int] = None
    is_minifig_part: bool = False

@dataclass
class WantedList:
    title: str
    items: List[RequiredItem] = field(default_factory=list)

def parse_wanted_lists():
    """
    Parse all wanted list XML files and return a dictionary:
    {list_title: [wanted_items]}
    """
    wanted_lists = {}
    for fn in os.listdir(WANTED_LISTS_DIR):
        if not fn.endswith(".xml"):
            continue
        title = os.path.splitext(fn)[0]
        tree = ET.parse(os.path.join(WANTED_LISTS_DIR, fn))
        items = []
        for it in tree.getroot().findall("ITEM"):
            item_id = it.find("ITEMID").text.strip()
            item_type = it.find("ITEMTYPE").text.strip()
            # --- handle color ---
            color_id = int(it.find("COLOR").text.strip()) if item_type == "P" and it.find("COLOR") is not None else None
            # --- required quantity ---
            minqty_elem = it.find("MINQTY")
            minqty = float(minqty_elem.text.strip()) if minqty_elem is not None else 1
            # mark "minifig parts" (loose parts of a minifig)
            is_minifig_part = (minqty_elem is None and item_type == "P")
            items.append({
                'item_id': item_id,
                'item_type': item_type,
                'minqty': minqty,
                'color_id': color_id,
                'isMinifigPart': is_minifig_part
            })
        wanted_lists[title] = items
    return wanted_lists

def parse_wanted_lists_objects() -> List[WantedList]:
    """
    Parse all wanted list XMLs and return a list of WantedList objects.
    Each RequiredItem.qty represents the per-unit requirement for a sellable unit.
    """
    results: List[WantedList] = []
    for fn in os.listdir(WANTED_LISTS_DIR):
        if not fn.endswith(".xml"):
            continue
        title = os.path.splitext(fn)[0]
        tree = ET.parse(os.path.join(WANTED_LISTS_DIR, fn))
        root = tree.getroot()
        wl_items: List[RequiredItem] = []
        for it in root.findall("ITEM"):
            item_id = (it.findtext("ITEMID") or "").strip()
            item_type = (it.findtext("ITEMTYPE") or "").strip()
            # color only relevant for parts; None otherwise
            color_id: Optional[int] = None
            if item_type == "P":
                color_text = (it.findtext("COLOR") or "").strip() if it.find("COLOR") is not None else ""
                if color_text:
                    try:
                        color_id = int(color_text)
                    except ValueError:
                        color_id = None
            minqty_elem = it.find("MINQTY")
            # per-unit requirement; default 1 when missing or invalid
            try:
                qty = int(float(minqty_elem.text.strip())) if minqty_elem is not None else 1
            except Exception:
                qty = 1
            is_minifig_part = (minqty_elem is None and item_type == "P")
            wl_items.append(RequiredItem(
                item_id=item_id,
                item_type=item_type,
                qty=qty,
                color_id=color_id,
                is_minifig_part=is_minifig_part
            ))
        results.append(WantedList(title=title, items=wl_items))
    return results
