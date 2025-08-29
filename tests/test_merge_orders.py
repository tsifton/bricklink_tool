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
        # First order should be the newer one (12346)
        self.assertEqual(orders[0].findtext('ORDERID'), '12346')
        self.assertEqual(orders[0].findtext('ORDERDATE'), '2024-08-20T15:45:00.000Z')
        # Second order should be the older one (12345)
        self.assertEqual(orders[1].findtext('ORDERID'), '12345')
        self.assertEqual(orders[1].findtext('ORDERDATE'), '2024-08-15T10:30:00.000Z')
    
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
        self.assertEqual(order.findtext('ORDERDATE'), '2024-08-25T09:15:00.000Z')
        # Verify it kept the newer quantity
        item = order.find('ITEM')
        self.assertEqual(item.findtext('QTY'), '12')
    
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


if __name__ == '__main__':
    unittest.main()