"""Single-lane cargo planning. No network, AI, or routing calls."""
from contextvars import ContextVar
from functools import wraps

order_context = ContextVar('cargo_order', default=None)


def scoped_order(function):
    """Keep render/route simulations isolated across sessions and exceptions."""
    @wraps(function)
    def wrapped(*args, **kwargs):
        token = order_context.set(order_context.get())
        try:
            return function(*args, **kwargs)
        finally:
            order_context.reset(token)
    return wrapped


def validate(order, capacity=100):
    if not isinstance(order, (list, tuple)) or any(x not in ('bike', 'ebike') for x in order):
        raise ValueError('車輛排列格式不正確，請重新確認')
    if len(order) > capacity:
        raise ValueError('車輛排列超過最高載量')
    return list(order)


MAX_TEMPORARY_MOVES = 3


def plan_unload(order, bike, ebike, max_temporary=MAX_TEMPORARY_MOVES):
    """Estimate a station's unloads with at most three bikes temporarily aside.

    The list is ordered cab-to-tail. A skipped bike is held in a temporary
    LIFO stack; multiple requested bikes behind it can be unloaded together,
    and the skipped bikes are restored in their original relative order.
    Returns (truck_remaining, unloaded_by_type, handling_moves, max_held).
    """
    remaining = validate(order)
    needs = {'bike': int(bike), 'ebike': int(ebike)}
    if any(number < 0 for number in needs.values()):
        raise ValueError('下車數量不可為負數')
    max_temporary = max(0, int(max_temporary))
    unloaded = {'bike': 0, 'ebike': 0}
    temporary = []
    handling_moves = 0
    max_held = 0

    while remaining and any(needs.values()):
        tail = remaining[-1]
        if needs[tail] > 0:
            remaining.pop()
            needs[tail] -= 1
            unloaded[tail] += 1
            continue

        available = max_temporary - len(temporary)
        if available <= 0:
            break
        # Never move a blocker without a requested vehicle within reach.
        reachable = remaining[-(available + 1):]
        if not any(needs[kind] > 0 for kind in reachable):
            break
        temporary.append(remaining.pop())
        handling_moves += 1
        max_held = max(max_held, len(temporary))

    while temporary:
        remaining.append(temporary.pop())
        handling_moves += 1
    return remaining, unloaded, handling_moves, max_held


def tail_unload(order, bike, ebike):
    """Recommend unloads reachable by moving at most three tail bikes."""
    remaining, unloaded, _moves, _max_held = plan_unload(order, bike, ebike)
    return remaining, unloaded


def complete(order, unload_bike, unload_ebike, pickup_order, capacity):
    """Commit actual site work without rejecting an edited unload sequence.

    Station completion trusts the operator's actual counts, even if a vehicle
    was beyond the recommended three temporary moves or manually rearranged.
    Prefer the reachable unloading plan; for any remaining requested bikes,
    remove the closest matching kind from the tail without reordering others.
    """
    requested = {'bike': int(unload_bike), 'ebike': int(unload_ebike)}
    remaining, unloaded, _moves, _max_held = plan_unload(
        order, requested['bike'], requested['ebike']
    )
    missing = {kind: requested[kind] - unloaded[kind] for kind in requested}
    if any(missing.values()):
        for index in range(len(remaining) - 1, -1, -1):
            kind = remaining[index]
            if missing[kind] > 0:
                remaining.pop(index)
                missing[kind] -= 1
        if any(missing.values()):
            raise ValueError('實際下車數量超過車上現有車輛，請確認下車數量')
    return validate(remaining + validate(pickup_order), capacity)


def loading_order(bike, ebike, tail_kind='ebike'):
    groups = {'bike': ['bike'] * bike, 'ebike': ['ebike'] * ebike}
    other = 'bike' if tail_kind == 'ebike' else 'ebike'
    return groups[other] + groups[tail_kind]


def label(order):
    return '車頭 → ' + ('、'.join('一般' if x == 'bike' else '電' for x in order) or '空車') + ' → 車尾（先卸）'


def choose_loading(remaining, bike, ebike, rows, station_name):
    """Compare two grouped loading orders against the existing snapshot only.

    This is an unloading-access hint, not a road/priority-ranked next destination.
    Existing dispatch ranking remains authoritative after completion.
    """
    choices = [loading_order(bike, ebike, kind) for kind in ('ebike', 'bike')]
    best, best_score, hint = choices[0], (-1, float('-inf')), ''
    for choice in choices:
        score, name = (-1, float('-inf')), ''
        for row in rows:
            if row.get('場站名稱') == station_name:
                continue
            try:
                if int(row.get('服務狀態', 1)) != 1:
                    continue
                b = max(0, int(row['2.0 標準']) - int(row['2.0 現況']))
                e = max(0, int(row['2.0E 標準']) - int(row['2.0E 現況']))
            except (ValueError, TypeError, KeyError, OverflowError):
                continue
            _, down, moves, _ = plan_unload(remaining + choice, b, e)
            candidate = (sum(down.values()), -moves)
            if candidate > score:
                score, name = candidate, str(row.get('場站名稱') or '')
        if score > best_score:
            best, best_score, hint = choice, score, name
    return best, hint
