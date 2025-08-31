"""
Test change detection functionality before merging orders.
"""
import unittest
import tempfile
import shutil
import os
import sys
import xml.etree.ElementTree as ET

# Add scripts directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from sheets import detect_changes_before_merge


class TestChangeDetection(unittest.TestCase):
    """Test change detection between order files and sheet data."""
    
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
    
    def create_test_xml(self, filename="orders.xml", order_id="12345", 
                       order_date="2024-01-01T10:30:00.000Z", seller="TestSeller",
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
        
        with open(os.path.join(self.orders_dir, filename), 'w', encoding='utf-8') as f:
            f.write(xml_content)

    def create_test_csv(self, filename="orders.csv", order_id="12345", 
                       order_date="2024-01-01T10:30:00.000Z", seller="TestSeller",
                       item_id="3001", qty=10, price=2.50):
        """Create a test CSV order file."""
        import csv
        csv_path = os.path.join(self.orders_dir, filename)
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'Order ID', 'Order Date', 'Seller', 'Shipping', 'Add Chrg 1',
                'Order Total', 'Base Grand Total', 'Total Lots', 'Total Items',
                'Tracking No', 'Item Number', 'Item Description', 'Color ID',
                'Qty', 'Each', 'Total', 'Condition', 'Inv ID', 'Item Type'
            ])
            writer.writeheader()
            
            # Order header row
            writer.writerow({
                'Order ID': order_id,
                'Order Date': order_date,
                'Seller': seller,
                'Shipping': '0.0',
                'Add Chrg 1': '0.0',
                'Order Total': str(qty * price),
                'Base Grand Total': str(qty * price + 2.50),
                'Total Lots': '1',
                'Total Items': str(qty),
                'Tracking No': '',
                'Item Number': '',
                'Item Description': '',
                'Color ID': '',
                'Qty': '',
                'Each': '',
                'Total': '',
                'Condition': '',
                'Inv ID': '',
                'Item Type': ''
            })
            
            # Item row
            writer.writerow({
                'Order ID': '',
                'Order Date': '',
                'Seller': '',
                'Shipping': '',
                'Add Chrg 1': '',
                'Order Total': '',
                'Base Grand Total': '',
                'Total Lots': '',
                'Total Items': '',
                'Tracking No': '',
                'Item Number': item_id,
                'Item Description': 'Test Brick',
                'Color ID': '4',
                'Qty': str(qty),
                'Each': str(price),
                'Total': str(qty * price),
                'Condition': 'N',
                'Inv ID': '',
                'Item Type': 'part'
            })

    def create_test_files(self, **kwargs):
        """Create both XML and CSV test files for complete testing."""
        self.create_test_xml(**kwargs)
        self.create_test_csv(**kwargs)
    
    def test_no_changes_detected(self):
        """Test that no changes are detected when sheet and files match."""
        # Create test order files (both XML and CSV)
        self.create_test_files()
        
        # Create matching sheet edits (no actual changes)
        sheet_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "25.0",
                "Base Grand Total": "27.5",
                "Item Number": ""
            },
            ("12345", "3001"): {
                "Order ID": "12345", 
                "Item Number": "3001",
                "Item Description": "Test Brick",  # Match what load_orders produces
                "Color": "P",  # This should match what load_orders produces
                "Condition": "N",
                "Qty": "10",
                "Each": "2.5",
                "Total": "25.0"
            }
        }
        
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Should detect no changes
        self.assertEqual(len(changes['edits']), 0)
        self.assertEqual(len(changes['additions']), 0) 
        self.assertEqual(len(changes['deletions']), 0)
    
    def test_edits_detected(self):
        """Test that edits are properly detected."""
        # Create test order files (both XML and CSV)
        self.create_test_files()
        
        # Create sheet edits with changes
        sheet_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "EditedSeller",  # Changed from TestSeller
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "30.0",  # Changed from 25.0
                "Base Grand Total": "27.5",
                "Item Number": ""
            },
            ("12345", "3001"): {
                "Order ID": "12345",
                "Item Number": "3001", 
                "Item Description": "",
                "Color": "P",  # Changed to match load_orders output
                "Condition": "U",  # Changed from N
                "Qty": "12",  # Changed from 10
                "Each": "2.5",
                "Total": "25.0"
            }
        }
        
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Should detect edits
        self.assertEqual(len(changes['edits']), 2)  # Both order and item rows changed
        self.assertEqual(len(changes['additions']), 0)
        self.assertEqual(len(changes['deletions']), 0)
        
        # Check specific edits
        edit_keys = [edit['key'] for edit in changes['edits']]
        self.assertIn(("12345", ""), edit_keys)
        self.assertIn(("12345", "3001"), edit_keys)
    
    def test_additions_detected(self):
        """Test that additions are properly detected.""" 
        # Create test order files (both XML and CSV)
        self.create_test_files()
        
        # Create sheet edits with original data plus additional entry
        sheet_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "25.0",
                "Base Grand Total": "27.5",
                "Item Number": ""
            },
            ("12345", "3001"): {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "",
                "Color": "Red", 
                "Condition": "N",
                "Qty": "10",
                "Each": "2.5",
                "Total": "25.0"
            },
            ("67890", ""): {  # Additional order not in files
                "Order ID": "67890",
                "Seller": "NewSeller",
                "Order Date": "2024-01-02T10:30:00.000Z",
                "Order Total": "15.0",
                "Base Grand Total": "17.5",
                "Item Number": ""
            }
        }
        
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Should detect addition
        self.assertEqual(len(changes['edits']), 0)
        self.assertEqual(len(changes['additions']), 1) 
        self.assertEqual(len(changes['deletions']), 0)
        
        # Check the addition
        self.assertEqual(changes['additions'][0]['key'], ("67890", ""))
        self.assertEqual(changes['additions'][0]['order_id'], "67890")
    
    def test_deletions_detected(self):
        """Test that deletions are properly detected."""
        # Create test order file with multiple items
        xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>TestSeller</SELLER>
    <ORDERTOTAL>50.0</ORDERTOTAL>
    <BASEGRANDTOTAL>52.5</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <QTY>10</QTY>
      <PRICE>2.5</PRICE>
      <CONDITION>N</CONDITION>
      <DESCRIPTION>Test Brick 1</DESCRIPTION>
    </ITEM>
    <ITEM>
      <ITEMID>3002</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>2</COLOR>
      <QTY>10</QTY>
      <PRICE>2.5</PRICE>
      <CONDITION>N</CONDITION>
      <DESCRIPTION>Test Brick 2</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        with open(os.path.join(self.orders_dir, 'orders.xml'), 'w', encoding='utf-8') as f:
            f.write(xml_content)
        
        # Create sheet edits missing one of the items
        sheet_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "50.0",
                "Base Grand Total": "52.5",
                "Item Number": ""
            },
            ("12345", "3001"): {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "",
                "Color": "Red",
                "Condition": "N",
                "Qty": "10",
                "Each": "2.5",
                "Total": "25.0"
            }
            # Missing ("12345", "3002") - should be detected as deletion
        }
        
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Should detect deletion
        self.assertEqual(len(changes['edits']), 0)
        self.assertEqual(len(changes['additions']), 0)
        self.assertEqual(len(changes['deletions']), 1)
        
        # Check the deletion
        self.assertEqual(changes['deletions'][0]['key'], ("12345", "3002"))
        self.assertEqual(changes['deletions'][0]['order_id'], "12345")
        self.assertEqual(changes['deletions'][0]['item_number'], "3002")
    
    def test_empty_sheet_edits(self):
        """Test that empty sheet edits returns empty changes."""
        self.create_test_xml()
        
        changes = detect_changes_before_merge({}, self.orders_dir)
        
        self.assertEqual(len(changes['edits']), 0)
        self.assertEqual(len(changes['additions']), 0)
        self.assertEqual(len(changes['deletions']), 0)
    
    def test_no_sheet_edits(self):
        """Test that None sheet edits returns empty changes."""
        self.create_test_xml()
        
        changes = detect_changes_before_merge(None, self.orders_dir)
        
        self.assertEqual(len(changes['edits']), 0)
        self.assertEqual(len(changes['additions']), 0)
        self.assertEqual(len(changes['deletions']), 0)


if __name__ == '__main__':
    unittest.main()