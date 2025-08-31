"""
Test the improved workflow that handles previously deleted orders being re-added during merge.
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

from sheets import detect_changes_before_merge, apply_saved_changes_to_files


class TestWorkflowImprovement(unittest.TestCase):
    """Test the improved workflow that preserves user deletions even after merging new files."""
    
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
    
    def test_workflow_preserves_deletions_after_merge(self):
        """Test that user deletions are preserved even when merge brings back the deleted orders."""
        
        # Step 1: Create initial merged file with orders 12345 and 12346
        xml_content_initial = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>Seller1</SELLER>
    <ORDERTOTAL>25.00</ORDERTOTAL>
    <BASEGRANDTOTAL>27.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>10</QTY>
      <PRICE>2.50</PRICE>
      <DESCRIPTION>Brick 1</DESCRIPTION>
    </ITEM>
  </ORDER>
  <ORDER>
    <ORDERID>12346</ORDERID>
    <ORDERDATE>2024-01-02T11:30:00.000Z</ORDERDATE>
    <SELLER>Seller2</SELLER>
    <ORDERTOTAL>15.00</ORDERTOTAL>
    <BASEGRANDTOTAL>17.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3002</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>5</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>5</QTY>
      <PRICE>3.00</PRICE>
      <DESCRIPTION>Brick 2</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        filepath = os.path.join(self.orders_dir, 'orders.xml')
        with open(filepath, 'w') as f:
            f.write(xml_content_initial)
        
        # Step 2: Simulate user deleting order 12346 from the sheet
        # The sheet_edits would only contain order 12345 now
        sheet_edits = {
            ('12345', ''): {
                'Order ID': '12345',
                'Seller': 'Seller1',
                'Order Date': '2024-01-01T10:30:00.000Z',
                'Order Total': '25.00',
                'Base Grand Total': '27.50',
                'Item Number': ''
            },
            ('12345', '3001'): {
                'Order ID': '12345',
                'Item Number': '3001',
                'Item Description': 'Brick 1',
                'Condition': 'N',
                'Qty': '10',
                'Each': '2.50',
                'Item Number': '3001'
            }
        }
        
        # Step 3: Detect changes before merge (this should detect that 12346 was deleted)
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Verify that the deletion was detected
        self.assertEqual(len(changes['deletions']), 2)  # Order header and item
        deleted_keys = [change['key'] for change in changes['deletions']]
        self.assertIn(('12346', ''), deleted_keys)
        self.assertIn(('12346', '3002'), deleted_keys)
        
        # Step 4: Simulate merge bringing back order 12346 (e.g., from a new downloaded file)
        # Let's say a new file had both orders, so after merge, we have both again
        xml_content_after_merge = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>Seller1</SELLER>
    <ORDERTOTAL>25.00</ORDERTOTAL>
    <BASEGRANDTOTAL>27.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>10</QTY>
      <PRICE>2.50</PRICE>
      <DESCRIPTION>Brick 1</DESCRIPTION>
    </ITEM>
  </ORDER>
  <ORDER>
    <ORDERID>12346</ORDERID>
    <ORDERDATE>2024-01-02T11:30:00.000Z</ORDERDATE>
    <SELLER>Seller2</SELLER>
    <ORDERTOTAL>15.00</ORDERTOTAL>
    <BASEGRANDTOTAL>17.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3002</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>5</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>5</QTY>
      <PRICE>3.00</PRICE>
      <DESCRIPTION>Brick 2</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        # Overwrite the file to simulate merge bringing back the deleted order
        with open(filepath, 'w') as f:
            f.write(xml_content_after_merge)
        
        # Step 5: Apply saved changes - this should re-apply the deletion
        apply_saved_changes_to_files(changes, self.orders_dir)
        
        # Step 6: Verify that order 12346 is deleted again, preserving the user's intent
        tree = ET.parse(filepath)
        root = tree.getroot()
        orders = root.findall('ORDER')
        
        # Should only have order 12345 left
        self.assertEqual(len(orders), 1)
        self.assertEqual(orders[0].findtext('ORDERID'), '12345')
        
        # Verify order 12346 is not present
        for order in orders:
            self.assertNotEqual(order.findtext('ORDERID'), '12346')
    
    def test_workflow_handles_additions_properly(self):
        """Test that the workflow properly adds new orders/items that were added in the sheet."""
        
        # Step 1: Create initial file with one order
        xml_content_initial = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>Seller1</SELLER>
    <ORDERTOTAL>25.00</ORDERTOTAL>
    <BASEGRANDTOTAL>27.50</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <CONDITION>N</CONDITION>
      <QTY>10</QTY>
      <PRICE>2.50</PRICE>
      <DESCRIPTION>Brick 1</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        filepath = os.path.join(self.orders_dir, 'orders.xml')
        with open(filepath, 'w') as f:
            f.write(xml_content_initial)
        
        # Step 2: Simulate user adding a new order in the sheet
        sheet_edits = {
            ('12345', ''): {
                'Order ID': '12345',
                'Seller': 'Seller1',
                'Order Date': '2024-01-01T10:30:00.000Z',
                'Order Total': '25.00',
                'Base Grand Total': '27.50',
                'Item Number': ''
            },
            ('12345', '3001'): {
                'Order ID': '12345',
                'Item Number': '3001',
                'Item Description': 'Brick 1',
                'Condition': 'N',
                'Qty': '10',
                'Each': '2.50'
            },
            ('12999', ''): {
                'Order ID': '12999',
                'Seller': 'NewSeller',
                'Order Date': '2024-01-05T15:00:00.000Z',
                'Order Total': '50.00',
                'Base Grand Total': '55.00',
                'Item Number': ''
            },
            ('12999', '3999'): {
                'Order ID': '12999',
                'Item Number': '3999',
                'Item Description': 'New Brick',
                'Condition': 'U',
                'Qty': '20',
                'Each': '2.50'
            }
        }
        
        # Step 3: Detect changes (should detect additions)
        changes = detect_changes_before_merge(sheet_edits, self.orders_dir)
        
        # Verify additions were detected
        self.assertEqual(len(changes['additions']), 2)  # Order and item
        addition_keys = [change['key'] for change in changes['additions']]
        self.assertIn(('12999', ''), addition_keys)
        self.assertIn(('12999', '3999'), addition_keys)
        
        # Step 4: Apply saved changes
        apply_saved_changes_to_files(changes, self.orders_dir)
        
        # Step 5: Verify that the new order was added to the file
        tree = ET.parse(filepath)
        root = tree.getroot()
        orders = root.findall('ORDER')
        
        # Should have 2 orders now
        self.assertEqual(len(orders), 2)
        
        # Verify the new order was added
        order_12999 = None
        for order in orders:
            if order.findtext('ORDERID') == '12999':
                order_12999 = order
                break
        
        self.assertIsNotNone(order_12999)
        self.assertEqual(order_12999.findtext('SELLER'), 'NewSeller')
        self.assertEqual(order_12999.findtext('ORDERTOTAL'), '50.00')
        
        # Verify the new item was added
        item = order_12999.find('ITEM')
        self.assertIsNotNone(item)
        self.assertEqual(item.findtext('ITEMID'), '3999')
        self.assertEqual(item.findtext('DESCRIPTION'), 'New Brick')
        self.assertEqual(item.findtext('CONDITION'), 'U')


if __name__ == '__main__':
    unittest.main()