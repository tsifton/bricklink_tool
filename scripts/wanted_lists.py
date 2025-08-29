import os
import xml.etree.ElementTree as ET
from config import WANTED_LISTS_DIR

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
