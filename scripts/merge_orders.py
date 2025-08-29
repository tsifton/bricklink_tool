import os
import xml.etree.ElementTree as ET
import csv
from datetime import datetime

ORDERS_DIR = os.path.join(os.path.dirname(__file__), "..", "orders")
MAIN_XML = os.path.join(ORDERS_DIR, "orders.xml")
MAIN_CSV = os.path.join(ORDERS_DIR, "orders.csv")



def parse_orderdate_xml(order):
    date_str = order.findtext("ORDERDATE", "").strip()
    try:
        # Try ISO format first, fallback to other formats as needed
        return datetime.fromisoformat(date_str)
    except Exception:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except Exception:
            return datetime.min

def merge_xml():
    all_orders = {}
    def clone_element(elem):
        return ET.fromstring(ET.tostring(elem, encoding="utf-8"))
    # Gather all orders from main file
    if os.path.exists(MAIN_XML):
        tree = ET.parse(MAIN_XML)
        for order in tree.getroot().findall("ORDER"):
            oid = order.findtext("ORDERID", "").strip()
            if oid:
                all_orders[oid] = clone_element(order)
    # Gather all orders from other XML files
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".xml") or fn == "orders.xml":
            continue
        path = os.path.join(ORDERS_DIR, fn)
        tree = ET.parse(path)
        for order in tree.getroot().findall("ORDER"):
            oid = order.findtext("ORDERID", "").strip()
            if oid and oid not in all_orders:
                all_orders[oid] = clone_element(order)
    # Sort all orders by ORDERDATE descending (newest first)
    orders = list(all_orders.values())
    orders.sort(key=parse_orderdate_xml, reverse=True)
    # Write all orders to a new <ORDERS> root
    root = ET.Element("ORDERS")
    for order in orders:
        root.append(order)
    tree = ET.ElementTree(root)
    tree.write(MAIN_XML, encoding="utf-8", xml_declaration=True)

def parse_orderdate_csv(row):
    date_str = row.get("Order Date", "").strip()
    try:
        return datetime.fromisoformat(date_str)
    except Exception:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except Exception:
            return datetime.min

def merge_csv():
    # Collect all rows from main and all other CSV files
    all_rows = {}
    fieldnames = None
    # Load existing orders.csv if present
    if os.path.exists(MAIN_CSV):
        with open(MAIN_CSV, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            if reader.fieldnames:
                fieldnames = reader.fieldnames
            for row in reader:
                oid = row.get("Order ID", "").strip()
                if oid:
                    all_rows[oid] = row
    # Load from all other CSV files
    for fn in os.listdir(ORDERS_DIR):
        if not fn.endswith(".csv") or fn == "orders.csv":
            continue
        path = os.path.join(ORDERS_DIR, fn)
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            if fieldnames is None and reader.fieldnames:
                fieldnames = reader.fieldnames
            for row in reader:
                oid = row.get("Order ID", "").strip()
                if oid and oid not in all_rows:
                    all_rows[oid] = row
    # Sort all rows by Order Date descending (newest first)
    rows = list(all_rows.values())
    rows.sort(key=parse_orderdate_csv, reverse=True)
    if fieldnames is None and rows:
        fieldnames = rows[0].keys()
    if fieldnames:
        with open(MAIN_CSV, "w", newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)