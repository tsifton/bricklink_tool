from copy import deepcopy

def determine_buildable(wanted_items, inventory):
    """
    Determine how many full builds can be made from inventory for a wanted list.
    Returns: (build_count, total_cost, updated_inventory)
    - wanted_items: list of dicts describing required items (sets, minifigs, parts)
    - inventory: dict mapping (item_id, color_id) or (item_id, None) to inventory info
    """
    # Make a deep copy so we don't mutate the original inventory
    inv = deepcopy(inventory)
    build_count = 0
    build_cost = 0.0

    # --- Sets (single set wanted per list) ---
    # Only one set per wanted list is supported. If present, check how many can be built.
    set_item = next((item for item in wanted_items if item['item_type'] == 'S'), None)
    if set_item:
        inv_item = inv.get((set_item['item_id'], None))
        if inv_item:
            builds = inv_item['qty'] // set_item['minqty']
            if builds:
                # Calculate total cost for all builds of this set
                build_cost += inv_item['unit_cost'] * set_item['minqty'] * builds
                # Deduct used quantity from inventory
                inv_item['qty'] -= set_item['minqty'] * builds
                build_count += builds

    # --- Minifigs (M) & Accessories (P that are not minifig parts) ---
    # Minifigs are keyed by item_id, accessories by (item_id, color_id)
    minifigs = {item['item_id']: item['minqty'] for item in wanted_items if item['item_type'] == 'M'}
    accessories = {
        (item['item_id'], item.get('color_id')): item['minqty']
        for item in wanted_items if item['item_type'] == 'P' and not item['isMinifigPart']
    }

    def total_qty(item_id):
        """
        Sum quantity for an item across all colors (tuple keys).
        Used for minifigs, which may have multiple color entries in inventory.
        """
        return sum(
            v['qty'] for k, v in inv.items()
            if (k == item_id if not isinstance(k, tuple) else k[0] == item_id)
        )

    def consume(item_id, amount):
        """
        Consume `amount` of a minifig across all its color entries in inventory.
        Returns the cost incurred for the consumed quantity.
        """
        left = amount
        cost = 0.0
        # Greedy: iterate in natural order; you could sort by unit_cost if desired
        for k, v in inv.items():
            match = (k == item_id) if not isinstance(k, tuple) else (k[0] == item_id)
            if not match or v['qty'] <= 0:
                continue
            take = min(v['qty'], left)
            if take:
                cost += v['unit_cost'] * take
                v['qty'] -= take
                left -= take
                if not left:
                    break
        return cost

    limits = []
    if minifigs or accessories:
        # Determine how many full builds can be made based on minifig and accessory requirements
        # For each required minifig, calculate how many full sets can be made from inventory
        limits += [total_qty(minifig_id) // req for minifig_id, req in minifigs.items()]
        # For each required accessory, calculate how many full sets can be made from inventory
        limits += [inv.get(part_id, {'qty': 0})['qty'] // req for part_id, req in accessories.items()]
        builds = min(limits) if limits else 0
        if builds:
            # Deduct minifigs (ignore color)
            build_cost += sum(consume(minifig_id, req * builds) for minifig_id, req in minifigs.items())
            # Deduct accessories (color-specific)
            for part_id, req in accessories.items():
                part_inv = inv[part_id]
                build_cost += part_inv['unit_cost'] * req * builds
                part_inv['qty'] -= req * builds
            build_count += builds

    # --- Parts-only builds (for lists with no sets/minifigs) ---
    # If the wanted list is just loose parts (e.g. minifig parts), handle as a parts-only build
    if any(item['isMinifigPart'] for item in wanted_items):
        # Build a dict of required parts and their quantities
        parts = {
            (item['item_id'], item.get('color_id')): item['minqty']
            for item in wanted_items if item['item_type'] == 'P'
        }
        if parts:
            # Find the limiting part (minimum number of builds possible)
            max_builds = min(inv.get(part_id, {'qty': 0})['qty'] // req for part_id, req in parts.items())
            if max_builds:
                for part_id, req in parts.items():
                    part_inv = inv[part_id]
                    build_cost += part_inv['unit_cost'] * req * max_builds
                    part_inv['qty'] -= req * max_builds
                build_count += max_builds

    # Return the number of builds, total cost, and the updated inventory after deduction
    return build_count, build_cost, inv
