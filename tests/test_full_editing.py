import os
import sys
import unittest
import tempfile
import shutil
import xml.etree.ElementTree as ET
import csv
from unittest.mock import Mock, patch

# Allow importing modules from the scripts directory
CURRENT_DIR = os.path.dirname(__file__)
SCRIPTS_DIR = os.path.abspath(os.path.join(CURRENT_DIR, '..', 'scripts'))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from sheets import (
    read_orders_sheet_edits,
    save_edits_to_files,
    detect_deleted_orders,
    remove_deleted_orders_from_files,
    update_orders_sheet
)


class TestFullSheetEditing(unittest.TestCase):
    
    def setUp(self):
        """Set up test environment with temporary directory and test files."""
        self.test_dir = tempfile.mkdtemp()
        self.orders_dir = self.test_dir
        
    def tearDown(self):
        """Clean up test environment."""
        shutil.rmtree(self.test_dir)
    
    def create_test_xml(self, filename, order_id="12345", order_date="2024-01-01T10:30:00.000Z",
                       seller="TestSeller", item_id="3001", qty=10, price=2.50):
        """Create a test XML order file."""
        xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>{order_id}</ORDERID>
    <ORDERDATE>{order_date}</ORDERDATE>
    <SELLER>{seller}</SELLER>
    <ORDERTOTAL>{qty * price}</ORDERTOTAL>
    <BASEGRANDTOTAL>{qty * price + 1.5}</BASEGRANDTOTAL>
    <ITEM>
      <ITEMID>{item_id}</ITEMID>
      <ITEMTYPE>P</ITEMTYPE>
      <COLOR>4</COLOR>
      <QTY>{qty}</QTY>
      <PRICE>{price}</PRICE>
      <CONDITION>N</CONDITION>
      <DESCRIPTION>Test Brick Description</DESCRIPTION>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        with open(os.path.join(self.test_dir, filename), 'w', encoding='utf-8') as f:
            f.write(xml_content)

    def create_test_csv(self, filename, order_id="12345", order_date="2024-01-01T10:30:00.000Z",
                       item_id="3001", qty=10, price=2.50):
        """Create a test CSV order file."""
        with open(os.path.join(self.test_dir, filename), 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'Order Number', 'Order Date', 'Item Number', 'Item Description',
                'Color ID', 'Qty', 'Each', 'Total', 'Condition'
            ])
            writer.writeheader()
            writer.writerow({
                'Order Number': order_id,
                'Order Date': order_date,
                'Item Number': item_id,
                'Item Description': 'Test Brick Description',
                'Color ID': '4',
                'Qty': str(qty),
                'Each': str(price),
                'Total': str(qty * price),
                'Condition': 'N'
            })

    def test_read_orders_sheet_edits_all_fields(self):
        """Test that read_orders_sheet_edits now captures all fields."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Mock sheet data with various edited fields
        mock_ws.get_all_records.return_value = [
            {
                "Order ID": "12345",
                "Seller": "EditedSeller",  # User edit
                "Order Date": "2024-02-01",  # User edit
                "Shipping": "5.99",  # User edit
                "Add Chrg 1": "2.00",
                "Order Total": "30.00",  # User edit
                "Base Grand Total": "32.50",
                "Total Lots": "3",
                "Total Items": "15", 
                "Tracking No": "1Z123456789",  # User edit
                "Condition": "",
                "Item Number": "",
                "Item Description": "",
                "Color": "",
                "Qty": "",
                "Each": "",
                "Total": ""
            },
            {
                "Order ID": "12345",
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Add Chrg 1": "",
                "Order Total": "",
                "Base Grand Total": "",
                "Total Lots": "",
                "Total Items": "",
                "Tracking No": "",
                "Condition": "U",  # User edit
                "Item Number": "3001",
                "Item Description": "Edited Brick Description",  # User edit
                "Color": "Blue",  # User edit
                "Qty": "12",  # User edit
                "Each": "3.00",  # User edit
                "Total": "36.00"  # User edit
            }
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws):
            edits = read_orders_sheet_edits(mock_sheet)
            
            # Check that ALL fields are captured, not just the previous limited set
            expected_edits = {
                ("12345", ""): {
                    "Order ID": "12345",
                    "Seller": "EditedSeller",
                    "Order Date": "2024-02-01",
                    "Shipping": "5.99",
                    "Add Chrg 1": "2.00",
                    "Order Total": "30.00",
                    "Base Grand Total": "32.50",
                    "Total Lots": "3",
                    "Total Items": "15",
                    "Tracking No": "1Z123456789",
                    "Condition": "",
                    "Item Number": "",
                    "Item Description": "",
                    "Color": "",
                    "Qty": "",
                    "Each": "",
                    "Total": ""
                },
                ("12345", "3001"): {
                    "Order ID": "12345",
                    "Seller": "",
                    "Order Date": "",
                    "Shipping": "",
                    "Add Chrg 1": "",
                    "Order Total": "",
                    "Base Grand Total": "",
                    "Total Lots": "",
                    "Total Items": "",
                    "Tracking No": "",
                    "Condition": "U",
                    "Item Number": "3001",
                    "Item Description": "Edited Brick Description",
                    "Color": "Blue",
                    "Qty": "12",
                    "Each": "3.00",
                    "Total": "36.00"
                }
            }
            self.assertEqual(edits, expected_edits)

    def test_save_edits_to_xml_files(self):
        """Test saving edited data back to XML files."""
        # Create test XML file
        self.create_test_xml('orders.xml')
        
        # Mock sheet edits
        sheet_edits = {
            ("12345", ""): {
                "Seller": "EditedSeller",
                "Order Date": "2024-02-01T12:00:00.000Z",
                "Order Total": "30.00",
                "Base Grand Total": "32.50"
            },
            ("12345", "3001"): {
                "Condition": "U",
                "Qty": "12",
                "Each": "3.00",
                "Item Description": "Edited Brick Description"
            }
        }
        
        # Save edits to files
        save_edits_to_files(sheet_edits, self.orders_dir)
        
        # Verify XML file was updated
        xml_file = os.path.join(self.test_dir, 'orders.xml')
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        order = root.find("ORDER")
        self.assertEqual(order.findtext("SELLER"), "EditedSeller")
        self.assertEqual(order.findtext("ORDERDATE"), "2024-02-01T12:00:00.000Z")
        self.assertEqual(order.findtext("ORDERTOTAL"), "30.00")
        self.assertEqual(order.findtext("BASEGRANDTOTAL"), "32.50")
        
        item = order.find("ITEM")
        self.assertEqual(item.findtext("CONDITION"), "U")
        self.assertEqual(item.findtext("QTY"), "12")
        self.assertEqual(item.findtext("PRICE"), "3.00")
        self.assertEqual(item.findtext("DESCRIPTION"), "Edited Brick Description")

    def test_save_edits_to_csv_files(self):
        """Test saving edited data back to CSV files."""
        # Create test CSV file
        self.create_test_csv('orders.csv')
        
        # Mock sheet edits
        sheet_edits = {
            ("12345", "3001"): {
                "Condition": "U",
                "Qty": "12",
                "Each": "3.00",
                "Total": "36.00",
                "Item Description": "Edited Brick Description"
            }
        }
        
        # Save edits to files
        save_edits_to_files(sheet_edits, self.orders_dir)
        
        # Verify CSV file was updated
        csv_file = os.path.join(self.test_dir, 'orders.csv')
        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            
            self.assertEqual(len(rows), 1)
            row = rows[0]
            self.assertEqual(row['Condition'], 'U')
            self.assertEqual(row['Qty'], '12')
            self.assertEqual(row['Each'], '3.00')
            self.assertEqual(row['Total'], '36.00')
            self.assertEqual(row['Item Description'], 'Edited Brick Description')

    def test_detect_deleted_orders(self):
        """Test detecting deleted orders/items."""
        # Original order rows
        original_rows = [
            {"Order ID": "12345", "Item Number": ""},  # Order header
            {"Order ID": "12345", "Item Number": "3001"},  # Item 1
            {"Order ID": "12345", "Item Number": "3002"},  # Item 2
            {"Order ID": "67890", "Item Number": ""},  # Another order header
            {"Order ID": "67890", "Item Number": "4001"},  # Another item
        ]
        
        # Sheet edits (missing some entries = they were deleted)
        sheet_edits = {
            ("12345", ""): {"Order ID": "12345"},  # Order header still there
            ("12345", "3001"): {"Order ID": "12345", "Item Number": "3001"},  # Item 1 still there
            # Missing: ("12345", "3002") - Item 2 was deleted
            # Missing: ("67890", "") - Order 67890 was deleted entirely
            # Missing: ("67890", "4001") - Item from deleted order
        }
        
        deleted_keys = detect_deleted_orders(original_rows, sheet_edits)
        
        expected_deleted = [("12345", "3002"), ("67890", ""), ("67890", "4001")]
        self.assertEqual(set(deleted_keys), set(expected_deleted))

    def test_remove_deleted_orders_from_xml_files(self):
        """Test removing deleted orders/items from XML files."""
        # Create XML with multiple orders and items
        xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
<ORDERS>
  <ORDER>
    <ORDERID>12345</ORDERID>
    <ORDERDATE>2024-01-01T10:30:00.000Z</ORDERDATE>
    <SELLER>TestSeller</SELLER>
    <ITEM>
      <ITEMID>3001</ITEMID>
      <QTY>10</QTY>
    </ITEM>
    <ITEM>
      <ITEMID>3002</ITEMID>
      <QTY>5</QTY>
    </ITEM>
  </ORDER>
  <ORDER>
    <ORDERID>67890</ORDERID>
    <ORDERDATE>2024-01-02T10:30:00.000Z</ORDERDATE>
    <SELLER>AnotherSeller</SELLER>
    <ITEM>
      <ITEMID>4001</ITEMID>
      <QTY>3</QTY>
    </ITEM>
  </ORDER>
</ORDERS>'''
        
        with open(os.path.join(self.test_dir, 'orders.xml'), 'w', encoding='utf-8') as f:
            f.write(xml_content)
        
        # Delete specific item and entire order
        deleted_keys = [("12345", "3002"), ("67890", "")]
        
        remove_deleted_orders_from_files(deleted_keys, self.orders_dir)
        
        # Verify XML file was updated
        xml_file = os.path.join(self.test_dir, 'orders.xml')
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # Should have only one order left (12345)
        orders = root.findall("ORDER")
        self.assertEqual(len(orders), 1)
        
        order = orders[0]
        self.assertEqual(order.findtext("ORDERID"), "12345")
        
        # Order 12345 should have only one item left (3001)
        items = order.findall("ITEM")
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].findtext("ITEMID"), "3001")

    def test_remove_deleted_orders_from_csv_files(self):
        """Test removing deleted orders/items from CSV files."""
        # Create CSV with multiple entries
        csv_file = os.path.join(self.test_dir, 'orders.csv')
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'Order Number', 'Order Date', 'Item Number', 'Qty'
            ])
            writer.writeheader()
            writer.writerow({
                'Order Number': '12345',
                'Order Date': '2024-01-01',
                'Item Number': '3001',
                'Qty': '10'
            })
            writer.writerow({
                'Order Number': '12345',
                'Order Date': '2024-01-01', 
                'Item Number': '3002',
                'Qty': '5'
            })
            writer.writerow({
                'Order Number': '67890',
                'Order Date': '2024-01-02',
                'Item Number': '4001', 
                'Qty': '3'
            })
        
        # Delete specific entries
        deleted_keys = [("12345", "3002"), ("67890", "4001")]
        
        remove_deleted_orders_from_files(deleted_keys, self.orders_dir)
        
        # Verify CSV file was updated
        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        
        # Should have only one row left
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['Order Number'], '12345')
        self.assertEqual(rows[0]['Item Number'], '3001')

    def test_update_orders_sheet_preserves_all_edits(self):
        """Test that update_orders_sheet preserves ALL user edits, not just limited fields."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Mock existing edits with various fields
        existing_edits = {
            ("12345", ""): {
                "Order ID": "12345",
                "Seller": "EditedSeller",  # User edit
                "Order Date": "2024-02-01",  # User edit
                "Shipping": "5.99",
                "Order Total": "30.00",  # User edit
                "Tracking No": "1Z123456789"
            },
            ("12345", "3001"): {
                "Order ID": "12345",
                "Item Number": "3001",
                "Item Description": "Edited Description",  # User edit
                "Condition": "U",  # User edit
                "Qty": "12",  # User edit
                "Each": "3.00"  # User edit
            }
        }
        
        # Mock order rows from the system
        order_rows = [
            {
                "Order ID": "12345",
                "Seller": "OriginalSeller",  # Should be replaced
                "Order Date": "2024-01-01",  # Should be replaced
                "Shipping": "",
                "Add Chrg 1": "",
                "Order Total": "25.00",  # Should be replaced
                "Base Grand Total": "26.50",
                "Total Lots": "",
                "Total Items": "",
                "Tracking No": "",  # Should be replaced
                "Condition": "",
                "Item Number": "",
                "Item Description": "",
                "Color": "",
                "Qty": "",
                "Each": "",
                "Total": ""
            },
            {
                "Order ID": "12345",
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Add Chrg 1": "",
                "Order Total": "",
                "Base Grand Total": "",
                "Total Lots": "",
                "Total Items": "",
                "Tracking No": "",
                "Condition": "N",  # Should be replaced
                "Item Number": "3001",
                "Item Description": "Original Description",  # Should be replaced
                "Color": "Red",
                "Qty": "10",  # Should be replaced
                "Each": "2.50",  # Should be replaced
                "Total": "25.00"
            }
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws), \
             patch('sheets.read_orders_sheet_edits', return_value=existing_edits):
            
            update_orders_sheet(mock_sheet, order_rows)
            
            # Verify ws.update was called
            mock_ws.update.assert_called_once()
            
            # Get the values that were written to the sheet
            call_args = mock_ws.update.call_args
            values = call_args[1]['values'] if 'values' in call_args[1] else call_args[0][0] if call_args[0] else None
            
            self.assertIsNotNone(values)
            self.assertTrue(len(values) > 1)  # Should have headers plus data
            
            headers = values[0]
            
            # Check that user edits were applied to the first row (order header)
            first_data_row = values[1]
            seller_index = headers.index("Seller")
            order_date_index = headers.index("Order Date") 
            order_total_index = headers.index("Order Total")
            tracking_index = headers.index("Tracking No")
            
            self.assertEqual(first_data_row[seller_index], "EditedSeller")
            self.assertEqual(first_data_row[order_date_index], "2024-02-01")
            self.assertEqual(first_data_row[order_total_index], "30.00")
            self.assertEqual(first_data_row[tracking_index], "1Z123456789")
            
            # Check that user edits were applied to the second row (item row)
            second_data_row = values[2]
            condition_index = headers.index("Condition")
            description_index = headers.index("Item Description")
            qty_index = headers.index("Qty")
            each_index = headers.index("Each")
            
            self.assertEqual(second_data_row[condition_index], "U")
            self.assertEqual(second_data_row[description_index], "Edited Description")
            self.assertEqual(second_data_row[qty_index], "12")
            self.assertEqual(second_data_row[each_index], "3.00")


if __name__ == '__main__':
    unittest.main()