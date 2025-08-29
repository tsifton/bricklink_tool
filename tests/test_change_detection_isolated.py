"""
Test change detection functionality in isolation without Google Sheets dependencies.
"""
import unittest
import tempfile
import shutil
import os
import sys
from unittest.mock import patch, MagicMock

# Add scripts directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))


class TestChangeDetectionIsolated(unittest.TestCase):
    """Test change detection logic without Google Sheets dependencies."""
    
    def setUp(self):
        """Set up test environment."""
        self.test_dir = tempfile.mkdtemp()
        self.orders_dir = os.path.join(self.test_dir, 'orders')
        os.makedirs(self.orders_dir, exist_ok=True)
    
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
    
    @patch('sheets.gspread', MagicMock())  # Mock gspread to avoid import error
    def test_detect_changes_no_changes(self):
        """Test that no changes are detected when data matches."""
        # Mock the load_orders function to return predictable data
        mock_order_rows = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller", 
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": 25.0,
                "Base Grand Total": 27.5,
                "Item Number": ""
            },
            {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "",
                "Color": "Red",
                "Condition": "N", 
                "Qty": 10,
                "Each": 2.5,
                "Total": 25.0
            }
        ]
        
        # Mock sheet edits that match the order data
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
            }
        }
        
        with patch('sheets.load_orders') as mock_load:
            mock_load.return_value = (None, mock_order_rows)
            
            from sheets import detect_changes_before_merge
            changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
            
            # Should detect no changes
            self.assertEqual(len(changes['edits']), 0)
            self.assertEqual(len(changes['additions']), 0)
            self.assertEqual(len(changes['deletions']), 0)
    
    @patch('sheets.gspread', MagicMock())  # Mock gspread to avoid import error
    def test_detect_changes_with_edits(self):
        """Test that edits are properly detected.""" 
        mock_order_rows = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": 25.0,
                "Base Grand Total": 27.5,
                "Item Number": ""
            },
            {
                "Order ID": "12345", 
                "Item Number": "3001",
                "Item Description": "",
                "Color": "Red",
                "Condition": "N",
                "Qty": 10,
                "Each": 2.5,
                "Total": 25.0
            }
        ]
        
        # Sheet edits with changes
        sheet_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "EditedSeller",  # Changed
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": "30.0",  # Changed
                "Base Grand Total": "27.5",
                "Item Number": ""
            },
            ("12345", "3001"): {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "",
                "Color": "Red",
                "Condition": "U",  # Changed 
                "Qty": "12",  # Changed
                "Each": "2.5",
                "Total": "25.0"
            }
        }
        
        with patch('sheets.load_orders') as mock_load:
            mock_load.return_value = (None, mock_order_rows)
            
            from sheets import detect_changes_before_merge
            changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
            
            # Should detect edits
            self.assertEqual(len(changes['edits']), 2)  # Both order and item rows changed
            self.assertEqual(len(changes['additions']), 0)
            self.assertEqual(len(changes['deletions']), 0)
            
            # Check specific edits
            edit_keys = [edit['key'] for edit in changes['edits']]
            self.assertIn(("12345", ""), edit_keys)
            self.assertIn(("12345", "3001"), edit_keys)
    
    @patch('sheets.gspread', MagicMock())  # Mock gspread to avoid import error  
    def test_detect_changes_with_deletions(self):
        """Test that deletions are properly detected."""
        mock_order_rows = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01T10:30:00.000Z",
                "Order Total": 50.0,
                "Base Grand Total": 52.5,
                "Item Number": ""
            },
            {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "", 
                "Color": "Red",
                "Condition": "N",
                "Qty": 10,
                "Each": 2.5,
                "Total": 25.0
            },
            {
                "Order ID": "12345", 
                "Item Number": "3002",
                "Item Description": "",
                "Color": "Blue",
                "Condition": "N",
                "Qty": 10,
                "Each": 2.5,
                "Total": 25.0
            }
        ]
        
        # Sheet edits missing one item (3002)
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
        
        with patch('sheets.load_orders') as mock_load:
            mock_load.return_value = (None, mock_order_rows)
            
            from sheets import detect_changes_before_merge
            changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
            
            # Should detect deletion
            self.assertEqual(len(changes['edits']), 0)
            self.assertEqual(len(changes['additions']), 0)
            self.assertEqual(len(changes['deletions']), 1)
            
            # Check the deletion
            self.assertEqual(changes['deletions'][0]['key'], ("12345", "3002"))
            self.assertEqual(changes['deletions'][0]['order_id'], "12345")
            self.assertEqual(changes['deletions'][0]['item_number'], "3002")
    
    @patch('sheets.gspread', MagicMock())  # Mock gspread to avoid import error
    def test_detect_changes_empty_sheet_edits(self):
        """Test that empty sheet edits returns empty changes."""
        with patch('sheets.load_orders') as mock_load:
            mock_load.return_value = (None, [])
            
            from sheets import detect_changes_before_merge
            changes = detect_changes_before_merge({}, self.orders_dir)
            
            self.assertEqual(len(changes['edits']), 0) 
            self.assertEqual(len(changes['additions']), 0)
            self.assertEqual(len(changes['deletions']), 0)


if __name__ == '__main__':
    unittest.main()