import os
import sys
import unittest
from unittest.mock import Mock, patch

# Allow importing modules from the scripts directory
CURRENT_DIR = os.path.dirname(__file__)
SCRIPTS_DIR = os.path.abspath(os.path.join(CURRENT_DIR, '..', 'scripts'))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from sheets import read_orders_sheet_edits, update_orders_sheet  # noqa: E402
from orders import Order, OrderItem  # noqa: E402


class TestSheetsEditing(unittest.TestCase):
    
    def test_read_orders_sheet_edits_empty_sheet(self):
        """Test reading edits from an empty sheet returns empty dict."""
        mock_sheet = Mock()
        mock_ws = Mock()
        mock_ws.get_all_records.return_value = []
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws):
            edits = read_orders_sheet_edits(mock_sheet)
            self.assertEqual(edits, {})

    def test_read_orders_sheet_edits_with_user_edits(self):
        """Test reading edits from a sheet with user-edited fields."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Mock sheet data with some user edits
        mock_ws.get_all_records.return_value = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Order Date": "2024-01-01",
                "Shipping": "5.99",  # User edit
                "Add Chrg 1": "",
                "Order Total": "25.00",
                "Tracking No": "1Z123456789",  # User edit
                "Item Number": "",
                "Item Description": ""
            },
            {
                "Order ID": "12345",
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Add Chrg 1": "",
                "Order Total": "",
                "Tracking No": "",
                "Item Number": "3001",
                "Item Description": "Brick 2 x 4",
                "Total Lots": "5"  # User edit
            }
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws):
            edits = read_orders_sheet_edits(mock_sheet)
            
            # Check that ALL fields are captured (new behavior)
            expected_edits = {
                ("12345", ""): {
                    "Order ID": "12345",
                    "Seller": "TestSeller",
                    "Order Date": "2024-01-01",
                    "Shipping": "5.99",
                    "Add Chrg 1": "",
                    "Order Total": "25.00",
                    "Tracking No": "1Z123456789",
                    "Item Number": "",
                    "Item Description": ""
                },
                ("12345", "3001"): {
                    "Order ID": "12345",
                    "Seller": "",
                    "Order Date": "",
                    "Shipping": "",
                    "Add Chrg 1": "",
                    "Order Total": "",
                    "Tracking No": "",
                    "Item Number": "3001",
                    "Item Description": "Brick 2 x 4",
                    "Total Lots": "5"
                }
            }
            self.assertEqual(edits, expected_edits)

    def test_read_orders_sheet_edits_handles_exceptions(self):
        """Test that read_orders_sheet_edits handles exceptions gracefully."""
        mock_sheet = Mock()
        
        with patch('sheets.get_or_create_worksheet', side_effect=Exception("Test error")):
            edits = read_orders_sheet_edits(mock_sheet)
            self.assertEqual(edits, {})

    def test_update_orders_sheet_preserves_edits(self):
        """Test that update_orders_sheet preserves user edits."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Mock existing edits
        existing_edits = {
            ("12345", ""): {"Shipping": "5.99", "Tracking No": "1Z123456789"},
            ("12345", "3001"): {"Total Lots": "2"}
        }
        
        # Create Order objects instead of dictionaries
        orders = [
            Order(
                order_id="12345",
                order_date="2024-01-01T10:30:00.000Z",
                seller="TestSeller",
                order_total=25.0,
                base_grand_total=26.5,
                shipping=0.0,  # Should be replaced with user edit
                tracking_no="",  # Should be replaced with user edit
                items=[
                    OrderItem(
                        item_id="3001",
                        item_type="P",
                        color_id=4,
                        qty=10,
                        price=0.05,
                        condition="N",
                        description="Brick 2 x 4",
                        color_name="Orange"
                    )
                ]
            )
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws), \
             patch('sheets.read_orders_sheet_edits', return_value=existing_edits):
            
            update_orders_sheet(mock_sheet, orders)
            
            # Verify ws.update was called
            self.assertTrue(mock_ws.update.called)
            
            # Get the values that were written to the sheet
            call_args = mock_ws.update.call_args
            values = call_args[1]['values'] if 'values' in call_args[1] else call_args[0][0] if call_args[0] else None
            
            self.assertIsNotNone(values)
            self.assertTrue(len(values) > 1)  # Should have headers plus data
            
            # Check that headers are correct
            headers = values[0]
            expected_headers = [
                "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg 1",
                "Order Total", "Base Grand Total", "Total Lots", "Total Items",
                "Tracking No", "Condition", "Item Number", "Item Description",
                "Color", "Qty", "Each", "Total"
            ]
            self.assertEqual(headers, expected_headers)
            
            # Check that user edits were applied to the data row
            first_data_row = values[1]
            shipping_index = headers.index("Shipping")
            tracking_index = headers.index("Tracking No")
            total_lots_index = headers.index("Total Lots")
            self.assertEqual(first_data_row[shipping_index], "5.99")
            self.assertEqual(first_data_row[tracking_index], "1Z123456789")
            self.assertEqual(first_data_row[total_lots_index], "2")

    def test_update_orders_sheet_no_existing_edits(self):
        """Test update_orders_sheet works when there are no existing edits."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Create Order objects instead of dictionaries
        orders = [
            Order(
                order_id="12345",
                order_date="2024-01-01T10:30:00.000Z",
                seller="TestSeller",
                order_total=25.0,
                base_grand_total=27.5,
                items=[
                    OrderItem(
                        item_id="3001",
                        item_type="P",
                        color_id=4,
                        qty=10,
                        price=2.5,
                        condition="N",
                        description="Test Brick"
                    )
                ]
            )
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws), \
             patch('sheets.read_orders_sheet_edits', return_value={}):
            
            update_orders_sheet(mock_sheet, orders)
            
            # Verify the function completes without error
            self.assertTrue(mock_ws.update.called)

    def test_read_orders_sheet_edits_handles_order_structure(self):
        """Test reading edits from sheet with proper order structure (empty Order ID for item rows)."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        # Mock sheet data mimicking the real structure where item rows have empty Order IDs
        mock_ws.get_all_records.return_value = [
            {
                "Order ID": "12345",  # Order header has Order ID
                "Seller": "TestSeller",
                "Order Date": "2024-01-01",
                "Shipping": "5.99",
                "Tracking No": "",
                "Item Number": "",  # No item number for order header
                "Item Description": ""
            },
            {
                "Order ID": "",  # Item row has empty Order ID (as shown in sheet)
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Tracking No": "",
                "Item Number": "3001",  # Has item number
                "Item Description": "Brick 2 x 4",
                "Condition": "N",
                "Qty": "10",
                "Each": "1.25"
            },
            {
                "Order ID": "",  # Another item row with empty Order ID
                "Seller": "",
                "Order Date": "",
                "Shipping": "",
                "Tracking No": "",
                "Item Number": "3002",
                "Item Description": "Brick 2 x 2",
                "Condition": "U",
                "Qty": "5",
                "Each": "2.00"
            }
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws):
            edits = read_orders_sheet_edits(mock_sheet)
        
        # Should have correct keys with reconstructed Order IDs
        expected_keys = {
            ("12345", ""),      # Order header
            ("12345", "3001"),  # First item (Order ID reconstructed)
            ("12345", "3002")   # Second item (Order ID reconstructed)
        }
        actual_keys = set(edits.keys())
        
        self.assertEqual(actual_keys, expected_keys)
        
        # Verify that Order IDs were correctly reconstructed in the records
        self.assertEqual(edits[("12345", "")]["Order ID"], "12345")
        self.assertEqual(edits[("12345", "3001")]["Order ID"], "12345")
        self.assertEqual(edits[("12345", "3002")]["Order ID"], "12345")
        
        # Verify other fields are preserved
        self.assertEqual(edits[("12345", "")]["Shipping"], "5.99")
        self.assertEqual(edits[("12345", "3001")]["Condition"], "N")
        self.assertEqual(edits[("12345", "3002")]["Condition"], "U")


if __name__ == '__main__':
    unittest.main()