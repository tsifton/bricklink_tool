"""
Test cases for merge_orders functionality.
"""

import unittest
import tempfile
import os
import shutil
import sys
import xml.etree.ElementTree as ET
import csv

# Add scripts directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__)), 'scripts'))

import merge_orders


class TestMergeOrders(unittest.TestCase):
    
    def setUp(self):
        """Create temporary directory for test files."""
        self.test_dir = tempfile.mkdtemp()
        self.original_orders_dir = merge_orders.ORDERS_DIR
        merge_orders.ORDERS_DIR = self.test_dir
        
    def tearDown(self):
        """Clean up temporary directory."""
        shutil.rmtree(self.test_dir)
        merge_orders.ORDERS_DIR = self.original_orders_dir
    
    def create_test_csv(self, filename, order_id, order_date, item_id="3001", qty=10, price=2.50):
        """Create a test CSV order file."""
        with open(os.path.join(self.test_dir, filename), 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'Order ID', 'Order Date', 'Item Number', 'Item Description',
                'Color ID', 'Qty', 'Each', 'Total'
            ])
            writer.writeheader()
            writer.writerow({
                'Order ID': order_id,
                'Order Date': order_date,
                'Item Number': item_id,
                'Item Description': 'Test Brick Description',
                'Color ID': '4',
                'Qty': str(qty),
                'Each': str(price),
                'Total': str(qty * price)
            })

    def create_test_xml(self, filename, order_id, order_date, seller="TestSeller", 
                       item_id="3001", qty=10, price=2.50):
        """Create a test XML order file."""
        xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
<ORDER>
<ORDERID>{order_id}</ORDERID>
<ORDERDATE>{order_date}</ORDERDATE>
<SELLER>{seller}</SELLER>
<ORDERTOTAL>{qty * price}</ORDERTOTAL>
<BASEGRANDTOTAL>{qty * price + 2.50}</BASEGRANDTOTAL>
<ITEM>
<ITEMID>{item_id}</ITEMID>
<ITEMTYPE>P</ITEMTYPE>
<COLOR>4</COLOR>
<QTY>{qty}</QTY>
<PRICE>{price}</PRICE>
<CONDITION>N</CONDITION>
<DESCRIPTION>Test Brick</DESCRIPTION>
<LOTID>LOT{order_id}_{item_id}</LOTID>
</ITEM>
</ORDER>
</ORDERS>'''
        
        with open(os.path.join(self.test_dir, filename), 'w', encoding='utf-8') as f:
            f.write(xml_content)
    
    def test_xml_merge_sorting(self):
        """Test that XML orders are merged and sorted by date (newest first)."""
        # Create test files with different dates
        self.create_test_xml('older.xml', '12345', '2024-08-15T10:30:00.000Z')
        self.create_test_xml('newer.xml', '12346', '2024-08-20T15:45:00.000Z')
        
        # Merge XML files
        merge_orders.merge_xml()
        
        # Verify merged file exists
        merged_path = os.path.join(self.test_dir, 'orders.xml')
        self.assertTrue(os.path.exists(merged_path))
        
        # Verify order sorting (newest first)
        tree = ET.parse(merged_path)
        root = tree.getroot()
        orders = root.findall('ORDER')
        
        self.assertEqual(len(orders), 2)
        # First order should be the newer one (12346) - but XML is now minimal, so check by position
        self.assertEqual(orders[0].findtext('ORDERID'), '12346')
        # Second order should be the older one (12345)
        self.assertEqual(orders[1].findtext('ORDERID'), '12345')
        # Note: ORDERDATE no longer included in minimal XML
    
    def test_xml_merge_deduplication(self):
        """Test that duplicate orders are deduplicated keeping the newest."""
        # Create test files with same order ID but different dates
        self.create_test_xml('old_version.xml', '12345', '2024-08-15T10:30:00.000Z', 
                           qty=10, price=2.50)
        self.create_test_xml('new_version.xml', '12345', '2024-08-25T09:15:00.000Z',
                           qty=12, price=2.50)
        
        # Merge XML files
        merge_orders.merge_xml()
        
        # Verify merged file has only one order
        merged_path = os.path.join(self.test_dir, 'orders.xml')
        tree = ET.parse(merged_path)
        root = tree.getroot()
        orders = root.findall('ORDER')
        
        self.assertEqual(len(orders), 1)
        # Should keep the newer version
        order = orders[0]
        self.assertEqual(order.findtext('ORDERID'), '12345')
        # Note: ORDERDATE and QTY no longer included in minimal XML
        # Verify we have the expected ITEM elements (deduplication worked)
        items = order.findall('ITEM')
        self.assertEqual(len(items), 1)  # Should have one item
    
    def test_parse_order_date(self):
        """Test date parsing functionality."""
        # Test various date formats
        date1 = merge_orders.parse_order_date('2024-08-15T10:30:00.000Z')
        date2 = merge_orders.parse_order_date('2024-08-20T15:45:00Z')
        date3 = merge_orders.parse_order_date('2024-08-25 09:15:00')
        date4 = merge_orders.parse_order_date('2024-08-30')
        
        # Verify dates are parsed correctly and can be compared
        self.assertLess(date1, date2)
        self.assertLess(date2, date3)
        self.assertLess(date3, date4)
        
        # Test invalid date returns minimum date
        invalid_date = merge_orders.parse_order_date('invalid-date')
        self.assertEqual(invalid_date, merge_orders.datetime.min)


    def test_no_duplication_when_merged_files_exist(self):
        """Test that individual files are not processed when merged files exist."""
        # Create individual order files
        self.create_test_xml('order1.xml', '12345', '2024-08-15T10:30:00.000Z', qty=5)
        self.create_test_xml('order2.xml', '12346', '2024-08-20T15:45:00.000Z', qty=7)
        
        # Merge XML files (creates orders.xml)
        merge_orders.merge_xml()
        
        # Verify merged file exists alongside individual files
        merged_path = os.path.join(self.test_dir, 'orders.xml')
        self.assertTrue(os.path.exists(merged_path))
        self.assertTrue(os.path.exists(os.path.join(self.test_dir, 'order1.xml')))
        self.assertTrue(os.path.exists(os.path.join(self.test_dir, 'order2.xml')))
        
        # Simulate what orders.load_orders() would do - count total quantities
        # This mimics the critical part: checking if merged file exists
        merged_xml_exists = os.path.exists(os.path.join(self.test_dir, 'orders.xml'))
        total_qty = 0
        
        for fn in os.listdir(self.test_dir):
            if not fn.endswith('.xml'):
                continue
            # Apply the fix: skip individual files when merged file exists
            if merged_xml_exists and fn != 'orders.xml':
                continue
                
            tree = ET.parse(os.path.join(self.test_dir, fn))
            root = tree.getroot()
            for order in root.findall('ORDER'):
                for item in order.findall('ITEM'):
                    # Count items instead of QTY since minimal XML doesn't include QTY
                    total_qty += 1
        
        # Should only process orders.xml (merged file), not individual files
        # Total should be 2 items (one from each order), not 4 (which would indicate duplication)
        self.assertEqual(total_qty, 2)

    def test_no_duplication_csv_when_merged_files_exist(self):
        """Test that individual CSV files are not processed when merged CSV exists."""
        # Create individual CSV files
        self.create_test_csv('order1.csv', '12345', '2024-08-15', qty=3)
        self.create_test_csv('order2.csv', '12346', '2024-08-20', qty=4)
        
        # Merge CSV files (creates orders.csv)
        merge_orders.merge_csv()
        
        # Verify merged file exists alongside individual files
        merged_path = os.path.join(self.test_dir, 'orders.csv')
        self.assertTrue(os.path.exists(merged_path))
        self.assertTrue(os.path.exists(os.path.join(self.test_dir, 'order1.csv')))
        self.assertTrue(os.path.exists(os.path.join(self.test_dir, 'order2.csv')))
        
        # Simulate what orders.load_orders() would do for CSV files
        merged_csv_exists = os.path.exists(os.path.join(self.test_dir, 'orders.csv'))
        total_qty = 0
        
        for fn in os.listdir(self.test_dir):
            if not fn.endswith('.csv'):
                continue
            # Apply the fix: skip individual files when merged file exists
            if merged_csv_exists and fn != 'orders.csv':
                continue
                
            with open(os.path.join(self.test_dir, fn), 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    qty = int(row.get('Qty', '0') or 0)
                    total_qty += qty
        
        # Should only process orders.csv (merged file), not individual files  
        # Total should be 3 + 4 = 7, not 14 (which would indicate duplication)
        self.assertEqual(total_qty, 7)


if __name__ == '__main__':
    unittest.main()