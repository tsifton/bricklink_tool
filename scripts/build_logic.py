from copy import deepcopy
from typing import List, Tuple, Dict, Any
from wanted_lists import WantedList, RequiredItem

def determine_buildable(wanted_list: WantedList, inventory) -> Tuple[int, float, List[Any]]:
    """
    Determine buildable count and cost from inventory for a wanted list.
    Returns: (build_count, total_cost, updated_inventory_list)
    """
    inv = deepcopy(inventory)
    total_builds = 0
    total_cost = 0.0
    req_items = [item for item in wanted_list.items if item.qty > 0]  # Skip zero-qty items

    def get_available_qty(item_id: str, item_type: str = None, color_id: int = None, 
                         ignore_color: bool = False) -> int:
        """Get total available quantity for an item."""
        return sum(
            getattr(item, 'qty', 0)
            for item in inv
            if (getattr(item, 'item_id', None) == item_id and
                (not item_type or getattr(item, 'item_type', None) == item_type) and
                (ignore_color or color_id is None or getattr(item, 'color_id', None) == color_id) and
                getattr(item, 'qty', 0) > 0)
        )

    def consume_items(item_id: str, amount: int, item_type: str = None, 
                     color_id: int = None, ignore_color: bool = False) -> float:
        """Consume items from inventory and return total cost."""
        remaining = amount
        cost = 0.0
        
        for item in inv:
            if remaining <= 0:
                break
                
            if (getattr(item, 'item_id', None) != item_id or
                getattr(item, 'qty', 0) <= 0 or
                (item_type and getattr(item, 'item_type', None) != item_type) or
                (not ignore_color and color_id is not None and 
                 getattr(item, 'color_id', None) != color_id)):
                continue
                
            take = min(item.qty, remaining)
            cost += float(getattr(item, 'unit_cost', 0.0)) * take
            item.qty -= take
            remaining -= take
            
        return cost

    # Process sets (only one set type per wanted list)
    set_items = [item for item in req_items if item.item_type == 'S']
    if set_items:
        set_item = set_items[0]  # Take first set item
        available = get_available_qty(set_item.item_id, 'S', ignore_color=True)
        builds = available // set_item.qty
        if builds > 0:
            cost = consume_items(set_item.item_id, set_item.qty * builds, 'S', ignore_color=True)
            total_builds += builds
            total_cost += cost

    # Process minifigs and accessories
    minifigs = {item.item_id: item.qty for item in req_items if item.item_type == 'M'}
    accessories = {
        (item.item_id, item.color_id): item.qty 
        for item in req_items 
        if item.item_type == 'P' and not item.is_minifig_part
    }

    if minifigs or accessories:
        build_limits = []
        
        # Calculate limits for minifigs
        for item_id, qty_needed in minifigs.items():
            available = get_available_qty(item_id, 'M', ignore_color=True)
            build_limits.append(available // qty_needed)
            
        # Calculate limits for accessories
        for (item_id, color_id), qty_needed in accessories.items():
            available = get_available_qty(item_id, 'P', color_id)
            build_limits.append(available // qty_needed)

        if build_limits:
            builds = min(build_limits)
            if builds > 0:
                # Consume minifigs
                for item_id, qty_needed in minifigs.items():
                    total_cost += consume_items(item_id, qty_needed * builds, 'M', ignore_color=True)
                
                # Consume accessories
                for (item_id, color_id), qty_needed in accessories.items():
                    total_cost += consume_items(item_id, qty_needed * builds, 'P', color_id)
                
                total_builds += builds

    # Process minifig parts (parts-only builds)
    minifig_parts = {
        (item.item_id, item.color_id): item.qty 
        for item in req_items 
        if item.item_type == 'P' and item.is_minifig_part
    }

    if minifig_parts:
        build_limits = [
            get_available_qty(item_id, 'P', color_id) // qty_needed
            for (item_id, color_id), qty_needed in minifig_parts.items()
        ]
        
        if build_limits:
            builds = min(build_limits)
            if builds > 0:
                for (item_id, color_id), qty_needed in minifig_parts.items():
                    total_cost += consume_items(item_id, qty_needed * builds, 'P', color_id)
                total_builds += builds

    return total_builds, total_cost, inv
