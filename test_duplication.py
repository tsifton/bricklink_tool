#!/usr/bin/env python3
"""
Test script to demonstrate the order duplication issue.
"""

import os
import tempfile
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'scripts'))

import merge_orders

def create_test_xml(directory, filename, order_id, order_date, qty=10, price=2.50):
    """Create a test XML order file."""
    xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
<ORDER>
<ORDERID>{order_id}</ORDERID>
<ORDERDATE>{order_date}</ORDERDATE>
<SELLER>TestSeller</SELLER>
<ORDERTOTAL>{qty * price}</ORDERTOTAL>
<BASEGRANDTOTAL>{qty * price + 2.50}</BASEGRANDTOTAL>
<ITEM>
<ITEMID>3001</ITEMID>
<ITEMTYPE>P</ITEMTYPE>
<COLOR>4</COLOR>
<QTY>{qty}</QTY>
<PRICE>{price}</PRICE>
<CONDITION>N</CONDITION>
<DESCRIPTION>Test Brick</DESCRIPTION>
</ITEM>
</ORDER>
</ORDERS>'''
    with open(os.path.join(directory, filename), 'w', encoding='utf-8') as f:
        f.write(xml_content)

def simulate_load_orders(orders_dir):
    """Simulate the load_orders function to count total quantity."""
    total_qty = 0
    
    # This mimics the logic in orders.py line 39
    for fn in os.listdir(orders_dir):
        if not fn.endswith(".xml"):
            continue
        tree = ET.parse(os.path.join(orders_dir, fn))
        root = tree.getroot()
        for order in root.findall("ORDER"):
            for item in order.findall("ITEM"):
                qty = int(item.findtext("QTY", "0") or 0)
                total_qty += qty
                
    return total_qty

def test_duplication_bug():
    """Test that demonstrates the duplication bug."""
    # Create temporary directory
    test_dir = tempfile.mkdtemp()
    original_orders_dir = merge_orders.ORDERS_DIR
    
    try:
        # Set up test environment
        merge_orders.ORDERS_DIR = test_dir
        
        # Create a single test order file
        create_test_xml(test_dir, 'order1.xml', '12345', '2024-08-15T10:30:00.000Z', qty=5)
        
        print(f"Test directory: {test_dir}")
        print("Files before merge:", os.listdir(test_dir))
        
        # Check quantity before merge
        qty_before = simulate_load_orders(test_dir)
        print(f"Total quantity before merge: {qty_before}")
        
        # Run merge_orders (which creates orders.xml)
        merge_orders.merge_xml()
        
        print("Files after merge:", os.listdir(test_dir))
        
        # Check quantity after merge - this simulates what orders.load_orders() would see
        qty_after = simulate_load_orders(test_dir)
        print(f"Total quantity after merge: {qty_after}")
        
        if qty_after > qty_before:
            print("BUG CONFIRMED: Orders are being duplicated!")
            return True
        else:
            print("No duplication detected.")
            return False
            
    finally:
        # Clean up
        shutil.rmtree(test_dir)
        merge_orders.ORDERS_DIR = original_orders_dir

if __name__ == '__main__':
    test_duplication_bug()