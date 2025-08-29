import os
import sys
import unittest

# Allow importing modules from the scripts directory
CURRENT_DIR = os.path.dirname(__file__)
SCRIPTS_DIR = os.path.abspath(os.path.join(CURRENT_DIR, '..', 'scripts'))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

import colors  # noqa: E402
from build_logic import determine_buildable  # noqa: E402


class TestColors(unittest.TestCase):
    def test_get_color_name_known_int(self):
        self.assertEqual(colors.get_color_name(1), "White")

    def test_get_color_name_known_str(self):
        self.assertEqual(colors.get_color_name("1"), "White")

    def test_get_color_name_unknown(self):
        self.assertIsNone(colors.get_color_name(99999))

    def test_get_color_name_nonint_str(self):
        self.assertEqual(colors.get_color_name("abc"), "abc")


class TestBuildLogic(unittest.TestCase):
    def test_set_only_builds(self):
        inv = {('1234', None): {
            'qty': 5, 'total_cost': 0.0, 'unit_cost': 2.0,
            'description': '', 'color_id': None, 'color_name': None
        }}
        wanted = [{
            'item_id': '1234', 'item_type': 'S', 'minqty': 1,
            'color_id': None, 'isMinifigPart': False
        }]
        count, cost, updated = determine_buildable(wanted, inv)
        self.assertEqual(count, 5)
        self.assertEqual(cost, 10.0)
        self.assertEqual(updated[('1234', None)]['qty'], 0)

    def test_minifig_and_accessory_builds(self):
        inv = {
            ('m1', None): {
                'qty': 3, 'total_cost': 0.0, 'unit_cost': 1.0,
                'description': '', 'color_id': None, 'color_name': None
            },
            ('p1', 5): {
                'qty': 6, 'total_cost': 0.0, 'unit_cost': 0.5,
                'description': '', 'color_id': 5, 'color_name': 'Red'
            }
        }
        wanted = [
            {'item_id': 'm1', 'item_type': 'M', 'minqty': 1, 'color_id': None, 'isMinifigPart': False},
            {'item_id': 'p1', 'item_type': 'P', 'minqty': 2, 'color_id': 5, 'isMinifigPart': False}
        ]
        count, cost, updated = determine_buildable(wanted, inv)
        self.assertEqual(count, 3)
        self.assertEqual(cost, 6.0)
        self.assertEqual(updated[('m1', None)]['qty'], 0)
        self.assertEqual(updated[('p1', 5)]['qty'], 0)

    def test_parts_only_builds(self):
        inv = {
            ('p1', 1): {
                'qty': 4, 'total_cost': 0.0, 'unit_cost': 0.25,
                'description': '', 'color_id': 1, 'color_name': 'White'
            },
            ('p2', 2): {
                'qty': 6, 'total_cost': 0.0, 'unit_cost': 0.5,
                'description': '', 'color_id': 2, 'color_name': 'Tan'
            },
        }
        # Use only minifig-part items to exercise the parts-only path
        wanted = [
            {'item_id': 'p1', 'item_type': 'P', 'minqty': 2, 'color_id': 1, 'isMinifigPart': True},
            {'item_id': 'p2', 'item_type': 'P', 'minqty': 3, 'color_id': 2, 'isMinifigPart': True},
        ]
        count, cost, updated = determine_buildable(wanted, inv)
        # Limiting builds: p1 -> 4//2 = 2, p2 -> 6//3 = 2
        self.assertEqual(count, 2)
        # Cost: p1 -> 0.25*2*2 = 1.0, p2 -> 0.5*3*2 = 3.0
        self.assertEqual(cost, 4.0)
        self.assertEqual(updated[('p1', 1)]['qty'], 0)
        self.assertEqual(updated[('p2', 2)]['qty'], 0)


if __name__ == '__main__':
    unittest.main()


if __name__ == '__main__':
    unittest.main()
