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
        
        # Mock order rows from the system
        order_rows = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller", 
                "Order Date": "2024-01-01",
                "Shipping": "",  # Should be replaced with user edit
                "Add Chrg 1": "",
                "Order Total": "25.00",
                "Base Grand Total": "26.50",
                "Total Lots": "",
                "Total Items": "",
                "Tracking No": "",  # Should be replaced with user edit
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
                "Total Lots": "",  # Should be replaced with user edit
                "Total Items": "",
                "Tracking No": "",
                "Condition": "N",
                "Item Number": "3001",
                "Item Description": "Brick 2 x 4",
                "Color": "Red",
                "Qty": "10",
                "Each": "0.05",
                "Total": "0.50"
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
            
            # Check that headers are correct
            headers = values[0]
            expected_headers = [
                "Order ID", "Seller", "Order Date", "Shipping", "Add Chrg 1",
                "Order Total", "Base Grand Total", "Total Lots", "Total Items",
                "Tracking No", "Condition", "Item Number", "Item Description",
                "Color", "Qty", "Each", "Total"
            ]
            self.assertEqual(headers, expected_headers)
            
            # Check that user edits were applied to the first row (order header)
            first_data_row = values[1]
            shipping_index = headers.index("Shipping")
            tracking_index = headers.index("Tracking No")
            self.assertEqual(first_data_row[shipping_index], "5.99")
            self.assertEqual(first_data_row[tracking_index], "1Z123456789")
            
            # Check that user edits were applied to the second row (item row)
            second_data_row = values[2]
            total_lots_index = headers.index("Total Lots")
            self.assertEqual(second_data_row[total_lots_index], "2")

    def test_update_orders_sheet_no_existing_edits(self):
        """Test update_orders_sheet works when there are no existing edits."""
        mock_sheet = Mock()
        mock_ws = Mock()
        
        order_rows = [
            {
                "Order ID": "12345",
                "Seller": "TestSeller",
                "Item Number": "",
                "Tracking No": ""
            }
        ]
        
        with patch('sheets.get_or_create_worksheet', return_value=mock_ws), \
             patch('sheets.read_orders_sheet_edits', return_value={}):
            
            update_orders_sheet(mock_sheet, order_rows)
            
            # Verify the function completes without error
            mock_ws.update.assert_called_once()


if __name__ == '__main__':
    unittest.main()