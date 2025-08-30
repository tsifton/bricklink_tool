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
from orders import OrderItem  # noqa: E402
from wanted_lists import WantedList, RequiredItem  # noqa: E402


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
        inv = [
            OrderItem(item_id='1234', item_type='S', color_id=0, qty=5,
                      price=0.0, unit_cost=2.0, description='', condition='')
        ]
        wanted = WantedList(
            title="sets",
            items=[RequiredItem(item_id='1234', item_type='S', qty=1, color_id=None)]
        )
        count, cost, updated = determine_buildable(wanted, inv)
        self.assertEqual(count, 5)
        self.assertEqual(cost, 10.0)
        self.assertEqual([it for it in updated if it.item_id == '1234' and it.item_type == 'S'][0].qty, 0)

    def test_minifig_and_accessory_builds(self):
        inv = [
            OrderItem(item_id='m1', item_type='M', color_id=0, qty=3,
                      price=0.0, unit_cost=1.0, description='', condition=''),
            OrderItem(item_id='p1', item_type='P', color_id=5, qty=6,
                      price=0.0, unit_cost=0.5, description='', condition='')
        ]
        wanted = WantedList(
            title="minifig+parts",
            items=[
                RequiredItem(item_id='m1', item_type='M', qty=1, color_id=None),
                RequiredItem(item_id='p1', item_type='P', qty=2, color_id=5),
            ]
        )
        count, cost, updated = determine_buildable(wanted, inv)
        self.assertEqual(count, 3)
        self.assertEqual(cost, 6.0)
        self.assertEqual([it for it in updated if it.item_id == 'm1' and it.item_type == 'M'][0].qty, 0)
        self.assertEqual([it for it in updated if it.item_id == 'p1' and it.color_id == 5][0].qty, 0)

    def test_parts_only_builds(self):
        inv = [
            OrderItem(item_id='p1', item_type='P', color_id=1, qty=4,
                      price=0.0, unit_cost=0.25, description='', condition=''),
            OrderItem(item_id='p2', item_type='P', color_id=2, qty=6,
                      price=0.0, unit_cost=0.5, description='', condition=''),
        ]
        # Use only minifig-part items to exercise the parts-only path
        wanted = WantedList(
            title="parts-only",
            items=[
                RequiredItem(item_id='p1', item_type='P', qty=2, color_id=1, is_minifig_part=True),
                RequiredItem(item_id='p2', item_type='P', qty=3, color_id=2, is_minifig_part=True),
            ]
        )
        count, cost, updated = determine_buildable(wanted, inv)
        # Limiting builds: p1 -> 4//2 = 2, p2 -> 6//3 = 2
        self.assertEqual(count, 2)
        # Cost: p1 -> 0.25*2*2 = 1.0, p2 -> 0.5*3*2 = 3.0
        self.assertEqual(cost, 4.0)
        self.assertEqual([it for it in updated if it.item_id == 'p1' and it.color_id == 1][0].qty, 0)
        self.assertEqual([it for it in updated if it.item_id == 'p2' and it.color_id == 2][0].qty, 0)


if __name__ == '__main__':
    unittest.main()
