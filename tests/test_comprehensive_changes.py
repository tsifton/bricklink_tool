"""
Test comprehensive change application functionality.
"""
import unittest
import tempfile
import shutil
import os
import sys
import xml.etree.ElementTree as ET
import csv

# Add scripts directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from sheets import apply_saved_changes_to_files


class TestComprehensiveChanges(unittest.TestCase):
    """Test applying all types of changes (edits, additions, deletions) in one step."""
    
    def setUp(self):
        """Set up test environment with temporary directories and files."""
        self.test_dir = tempfile.mkdtemp()
        self.orders_dir = os.path.join(self.test_dir, 'orders')
        os.makedirs(self.orders_dir, exist_ok=True)
        
        # Set up config to use test directory
        os.environ['ORDERS_DIR'] = self.orders_dir
        import config
        config.ORDERS_DIR = self.orders_dir
    
    def tearDown(self):
        """Clean up test environment."""
        shutil.rmtree(self.test_dir)
    
    def create_test_xml(self, filename="orders.xml"):
        """Create a test XML file with existing orders."""
        xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>OriginalSeller</SELLER>
    <ORDERTOTAL>25.00</ORDERTOTAL>
    <BASEGRANDTOTAL>27.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>10</QTY>
      <PRICE>2.50</PRICE>
      <DESCRIPTION>Original Brick Description</DESCRIPTION>
    </ITEM>
  </ORDER>
  <ORDER>
    <ORDERID>12346</ORDERID>
    <ORDERDATE>2024-01-02T11:30:00.000Z</ORDERDATE>
    <SELLER>DeleteMeSeller</SELLER>
    <ORDERTOTAL>15.00</ORDERTOTAL>
    <BASEGRANDTOTAL>17.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3002</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>5</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>5</QTY>
      <PRICE>3.00</PRICE>
      <DESCRIPTION>Delete Me Item</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        filepath = os.path.join(self.orders_dir, filename)
        with open(filepath, 'w') as f:
            f.write(xml_content)
    
    def create_test_csv(self, filename="orders.csv"):
        """Create a test CSV file with existing orders."""
        csv_data = [
            {
                "Order ID": "12345",
                "Seller": "OriginalSeller", 
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "25.00",
                "Base Grand Total": "27.50",
                "Item Number": ""
            },
            {
                "Order ID": "",
                "Seller": "",
                "Order Date": "",
                "Order Total": "",
                "Base Grand Total": "",
                "Item Number": "3001",
                "Item Description": "Original Brick Description",
                "Condition": "N",
                "Qty": "10",
                "Each": "2.50"
            },
            {
                "Order ID": "12346",
                "Seller": "DeleteMeSeller",
                "Order Date": "2024-01-02T11:30:00.000Z",
                "Order Total": "15.00",
                "Base Grand Total": "17.50",
                "Item Number": ""
            },
            {
                "Order ID": "",
                "Seller": "",
                "Order Date": "",
                "Order Total": "",
                "Base Grand Total": "",
                "Item Number": "3002",
                "Item Description": "Delete Me Item",
                "Condition": "N",
                "Qty": "5",
                "Each": "3.00"
            }
        ]
        
        fieldnames = ["Order ID", "Seller", "Order Date", "Order Total", "Base Grand Total", 
                     "Item Number", "Item Description", "Condition", "Qty", "Each"]
        
        filepath = os.path.join(self.orders_dir, filename)
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_data)
    
    def test_apply_comprehensive_changes_xml(self):
        """Test applying edits, additions, and deletions to XML files."""
        # Create test XML file
        self.create_test_xml()
        
        # Define comprehensive changes
        changes = {
            'edits': [
                {
                    'key': ('12345', ''),
                    'order_id': '12345',
                    'item_number': '',
                    'changes': {
                        'Seller': 'EditedSeller',
                        'Order Total': '30.00'
                    }
                },
                {
                    'key': ('12345', '3001'),
                    'order_id': '12345',
                    'item_number': '3001',
                    'changes': {
                        'Qty': '15',
                        'Item Description': 'Edited Brick Description'
                    }
                }
            ],
            'additions': [
                {
                    'key': ('12347', ''),
                    'order_id': '12347',
                    'item_number': '',
                    'data': {
                        'Order ID': '12347',
                        'Seller': 'NewSeller',
                        'Order Date': '2024-01-03T12:00:00.000Z',
                        'Order Total': '20.00',
                        'Base Grand Total': '22.50'
                    }
                },
                {
                    'key': ('12347', '3003'),
                    'order_id': '12347',
                    'item_number': '3003',
                    'data': {
                        'Item Number': '3003',
                        'Item Description': 'New Item',
                        'Condition': 'U',
                        'Qty': '8',
                        'Each': '2.75'
                    }
                }
            ],
            'deletions': [
                {
                    'key': ('12346', ''),
                    'order_id': '12346',
                    'item_number': ''
                }
            ]
        }
        
        # Apply comprehensive changes
        apply_saved_changes_to_files(changes, self.orders_dir)
        
        # Verify changes were applied correctly
        xml_file = os.path.join(self.orders_dir, 'orders.xml')
        tree = ET.parse(xml_file)
        root = tree.getroot()
        orders = root.findall('ORDER')
        
        # Should have 2 orders now (12345 edited, 12346 deleted, 12347 added)
        self.assertEqual(len(orders), 2)
        
        # Verify order 12345 was edited
        order_12345 = None
        for order in orders:
            if order.findtext('ORDERID') == '12345':
                order_12345 = order
                break
        
        self.assertIsNotNone(order_12345)
        self.assertEqual(order_12345.findtext('SELLER'), 'EditedSeller')
        self.assertEqual(order_12345.findtext('ORDERTOTAL'), '30.00')
        
        # Verify item in order 12345 was edited
        item = order_12345.find('ITEM')
        self.assertEqual(item.findtext('QTY'), '15')
        self.assertEqual(item.findtext('DESCRIPTION'), 'Edited Brick Description')
        
        # Verify order 12346 was deleted
        order_12346 = None
        for order in orders:
            if order.findtext('ORDERID') == '12346':
                order_12346 = order
                break
        self.assertIsNone(order_12346)
        
        # Verify order 12347 was added
        order_12347 = None
        for order in orders:
            if order.findtext('ORDERID') == '12347':
                order_12347 = order
                break
        
        self.assertIsNotNone(order_12347)
        self.assertEqual(order_12347.findtext('SELLER'), 'NewSeller')
        self.assertEqual(order_12347.findtext('ORDERTOTAL'), '20.00')
        
        # Verify item in order 12347 was added
        item = order_12347.find('ITEM')
        self.assertEqual(item.findtext('ITEMID'), '3003')
        self.assertEqual(item.findtext('DESCRIPTION'), 'New Item')
        self.assertEqual(item.findtext('CONDITION'), 'U')
    
    def test_apply_comprehensive_changes_csv(self):
        """Test applying edits, additions, and deletions to CSV files."""
        # Create test CSV file
        self.create_test_csv()
        
        # Define comprehensive changes
        changes = {
            'edits': [
                {
                    'key': ('12345', ''),
                    'order_id': '12345',
                    'item_number': '',
                    'changes': {
                        'Seller': 'EditedSeller',
                        'Order Total': '30.00'
                    }
                },
                {
                    'key': ('12345', '3001'),
                    'order_id': '12345',
                    'item_number': '3001',
                    'changes': {
                        'Qty': '15',
                        'Item Description': 'Edited Brick Description'
                    }
                }
            ],
            'additions': [
                {
                    'key': ('12347', ''),
                    'order_id': '12347',
                    'item_number': '',
                    'data': {
                        'Order ID': '12347',
                        'Seller': 'NewSeller',
                        'Order Date': '2024-01-03T12:00:00.000Z',
                        'Order Total': '20.00',
                        'Base Grand Total': '22.50',
                        'Item Number': ''
                    }
                },
                {
                    'key': ('12347', '3003'),
                    'order_id': '12347',
                    'item_number': '3003',
                    'data': {
                        'Order ID': '',
                        'Seller': '',
                        'Order Date': '',
                        'Order Total': '',
                        'Base Grand Total': '',
                        'Item Number': '3003',
                        'Item Description': 'New Item',
                        'Condition': 'U',
                        'Qty': '8',
                        'Each': '2.75'
                    }
                }
            ],
            'deletions': [
                {
                    'key': ('12346', ''),
                    'order_id': '12346',
                    'item_number': ''
                },
                {
                    'key': ('12346', '3002'),
                    'order_id': '12346',
                    'item_number': '3002'
                }
            ]
        }
        
        # Apply comprehensive changes
        apply_saved_changes_to_files(changes, self.orders_dir)
        
        # Verify changes were applied correctly
        csv_file = os.path.join(self.orders_dir, 'orders.csv')
        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        
        # Should have 4 rows now (2 original edited + 2 new added - 2 deleted)
        self.assertEqual(len(rows), 4)
        
        # Verify order 12345 header was edited
        header_12345 = None
        for row in rows:
            if row['Order ID'] == '12345' and not row['Item Number']:
                header_12345 = row
                break
        
        self.assertIsNotNone(header_12345)
        self.assertEqual(header_12345['Seller'], 'EditedSeller')
        self.assertEqual(header_12345['Order Total'], '30.00')
        
        # Verify item in order 12345 was edited
        item_12345 = None
        for row in rows:
            if not row['Order ID'] and row['Item Number'] == '3001':
                item_12345 = row
                break
        
        self.assertIsNotNone(item_12345)
        self.assertEqual(item_12345['Qty'], '15')
        self.assertEqual(item_12345['Item Description'], 'Edited Brick Description')
        
        # Verify order 12346 and its item were deleted
        for row in rows:
            self.assertNotEqual(row['Order ID'], '12346')
            self.assertNotEqual(row['Item Number'], '3002')
        
        # Verify order 12347 was added
        header_12347 = None
        for row in rows:
            if row['Order ID'] == '12347' and not row['Item Number']:
                header_12347 = row
                break
        
        self.assertIsNotNone(header_12347)
        self.assertEqual(header_12347['Seller'], 'NewSeller')
        self.assertEqual(header_12347['Order Total'], '20.00')
        
        # Verify item in order 12347 was added
        item_12347 = None
        for row in rows:
            if not row['Order ID'] and row['Item Number'] == '3003':
                item_12347 = row
                break
        
        self.assertIsNotNone(item_12347)
        self.assertEqual(item_12347['Item Description'], 'New Item')
        self.assertEqual(item_12347['Condition'], 'U')
        self.assertEqual(item_12347['Qty'], '8')


if __name__ == '__main__':
    unittest.main()