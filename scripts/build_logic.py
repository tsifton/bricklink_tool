from copy import deepcopy
from wanted_lists import WantedList, RequiredItem

def determine_buildable(wanted_list: WantedList, inventory):
    """
    Determine how many full builds can be made from inventory for a wanted list.
    - wanted_list: WantedList with RequiredItem entries (per-unit requirements).
    - inventory: list of OrderItem-like objects with attrs item_id, item_type, color_id, qty, unit_cost.
    Returns: (build_count, total_cost, updated_inventory_list)
    """
    inv = deepcopy(inventory)
    build_count = 0
    build_cost = 0.0
    req_items = wanted_list.items

    def total_qty(item_id, *, item_type=None, color_id=None, ignore_color=False):
        total = 0
        for it in inv:
            if item_type and getattr(it, 'item_type', None) != item_type:
                continue
            if getattr(it, 'item_id', None) != item_id:
                continue
            if not ignore_color and color_id is not None and getattr(it, 'color_id', None) != color_id:
                continue
            total += getattr(it, 'qty', 0)
        return total

    def consume(item_id, amount, *, item_type=None, color_id=None, ignore_color=False):
        left = amount
        cost = 0.0
        for it in inv:
            if item_type and getattr(it, 'item_type', None) != item_type:
                continue
            if getattr(it, 'item_id', None) != item_id:
                continue
            if not ignore_color and color_id is not None and getattr(it, 'color_id', None) != color_id:
                continue
            if getattr(it, 'qty', 0) <= 0:
                continue
            take = min(it.qty, left)
            if take:
                cost += float(getattr(it, 'unit_cost', 0.0)) * take
                it.qty -= take
                left -= take
                if left == 0:
                    break
        return cost

    # Sets (only one set entry supported per wanted list)
    set_item = next((ri for ri in req_items if ri.item_type == 'S'), None)
    if set_item:
        builds = total_qty(set_item.item_id, item_type='S', ignore_color=True) // set_item.qty
        if builds:
            build_cost += consume(set_item.item_id, set_item.qty * builds, item_type='S', ignore_color=True)
            build_count += builds

    # Minifigs and accessories
    minifigs = {ri.item_id: ri.qty for ri in req_items if ri.item_type == 'M'}
    accessories = {
        (ri.item_id, ri.color_id): ri.qty
        for ri in req_items
        if ri.item_type == 'P' and not ri.is_minifig_part
    }

    limits = []
    if minifigs or accessories:
        limits += [
            total_qty(mid, item_type='M', ignore_color=True) // req
            for mid, req in minifigs.items()
        ]
        limits += [
            total_qty(pid, item_type='P', color_id=cid) // req
            for (pid, cid), req in accessories.items()
        ]
        builds = min(limits) if limits else 0
        if builds:
            # Consume minifigs (ignore color)
            for mid, req in minifigs.items():
                build_cost += consume(mid, req * builds, item_type='M', ignore_color=True)
            # Consume accessories (color-specific)
            for (pid, cid), req in accessories.items():
                build_cost += consume(pid, req * builds, item_type='P', color_id=cid)
            build_count += builds

    # Parts-only builds (e.g., minifig parts lists)
    if any(ri.is_minifig_part for ri in req_items):
        parts = {(ri.item_id, ri.color_id): ri.qty for ri in req_items if ri.item_type == 'P'}
        if parts:
            max_builds = min(
                total_qty(pid, item_type='P', color_id=cid) // req
                for (pid, cid), req in parts.items()
            )
            if max_builds:
                for (pid, cid), req in parts.items():
                    build_cost += consume(pid, req * max_builds, item_type='P', color_id=cid)
                build_count += max_builds

    return build_count, build_cost, inv
