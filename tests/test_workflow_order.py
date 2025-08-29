"""
Test to verify the correct workflow order in main.py.
This test ensures that sheet edits are read BEFORE merging order files.
"""
import unittest
import re


class TestWorkflowOrder(unittest.TestCase):
    """Test that the main workflow follows the correct order of operations."""
    
    def test_main_function_order(self):
        """Test that main.py has the correct order of operations."""
        
        # Read the main.py file
        import os
        main_path = os.path.join(os.path.dirname(__file__), '..', 'scripts', 'main.py')
        with open(main_path, 'r') as f:
            content = f.read()
            
        # Find the main() function
        main_func_match = re.search(r'def main\(\):(.*?)^if __name__', content, re.DOTALL | re.MULTILINE)
        self.assertIsNotNone(main_func_match, "Could not find main() function")
        
        main_func = main_func_match.group(1)
        
        # Find the positions of key operations
        operations = {
            'load_google_sheet': main_func.find('load_google_sheet()'),
            'read_orders_sheet_edits': main_func.find('read_orders_sheet_edits('),
            'detect_changes_before_merge': main_func.find('detect_changes_before_merge('),
            'merge_xml': main_func.find('merge_orders.merge_xml()'),
            'merge_csv': main_func.find('merge_orders.merge_csv()')
        }
        
        # Verify all operations are present
        for op, pos in operations.items():
            self.assertNotEqual(pos, -1, f"Operation {op} not found in main() function")
        
        # Verify the correct order
        self.assertLess(operations['load_google_sheet'], operations['read_orders_sheet_edits'],
                       "load_google_sheet() should come before read_orders_sheet_edits()")
        
        self.assertLess(operations['read_orders_sheet_edits'], operations['detect_changes_before_merge'],
                       "read_orders_sheet_edits() should come before detect_changes_before_merge()")
        
        self.assertLess(operations['detect_changes_before_merge'], operations['merge_xml'],
                       "detect_changes_before_merge() should come before merge_orders.merge_xml()")
        
        self.assertLess(operations['merge_xml'], operations['merge_csv'],
                       "merge_orders.merge_xml() should come before merge_orders.merge_csv()")
        
        print("✓ Workflow order is correct:")
        print("  1. load_google_sheet()")
        print("  2. read_orders_sheet_edits()")
        print("  3. detect_changes_before_merge()")
        print("  4. merge_orders.merge_xml()")
        print("  5. merge_orders.merge_csv()")


if __name__ == '__main__':
    unittest.main()