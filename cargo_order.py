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


def tail_unload(order, bike, ebike):
    remaining = validate(order)
    needs = {'bike': int(bike), 'ebike': int(ebike)}
    unloaded = {'bike': 0, 'ebike': 0}
    while remaining and needs[remaining[-1]] > 0:
        kind = remaining.pop()
        needs[kind] -= 1
        unloaded[kind] += 1
    return remaining, unloaded


def complete(order, unload_bike, unload_ebike, pickup_order, capacity):
    remaining, unloaded = tail_unload(order, unload_bike, unload_ebike)
    if unloaded != {'bike': unload_bike, 'ebike': unload_ebike}:
        raise ValueError('實際下車順序被車尾車輛擋住，請修正下車數量或車上排列')
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
    best, best_score, hint = choices[0], -1, ''
    for choice in choices:
        score, name = 0, ''
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
            _, down = tail_unload(remaining + choice, b, e)
            amount = sum(down.values())
            if amount > score:
                score, name = amount, str(row.get('場站名稱') or '')
        if score > best_score:
            best, best_score, hint = choice, score, name
    return best, hint
