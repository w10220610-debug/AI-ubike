"""Local, bounded 2–3 stop unloading forecast; never calls a road API.

A forecast uses configuration deficits, priority names, coordinates and the
confirmed truck sequence. Distances between future stops are *only* a proximity
heuristic, not road directions; the existing road router selects the actual next
destination after the operator completes each stop.
"""
from __future__ import annotations

from math import asin, cos, radians, sin, sqrt

from cargo_order import loading_order, plan_unload, validate


def _number(value):
    try:
        number = int(value)
        return max(0, number)
    except (ValueError, TypeError, OverflowError):
        return None


def _coordinate(value, limit):
    try:
        number = float(value)
        return number if -limit <= number <= limit else None
    except (ValueError, TypeError, OverflowError):
        return None


def _distance_km(first, second):
    """Geodesic proximity proxy; NEVER present this as a driving distance."""
    a, b = first["lat"], second["lat"]
    c, d = first["lon"], second["lon"]
    dl = radians(d - c)
    da = radians(b - a)
    h = sin(da / 2) ** 2 + cos(radians(a)) * cos(radians(b)) * sin(dl / 2) ** 2
    return 12742 * asin(min(1, sqrt(max(0, h))))


def _station_pool(rows, locations, current_name, priority_names):
    pool = []
    for row in rows:
        name = str(row.get("場站名稱") or "").strip()
        if not name or name == current_name:
            continue
        if _number(row.get("服務狀態", 1)) != 1:
            continue
        current_bike = _number(row.get("2.0 現況"))
        current_ebike = _number(row.get("2.0E 現況"))
        standard_bike = _number(row.get("2.0 標準"))
        standard_ebike = _number(row.get("2.0E 標準"))
        if None in (current_bike, current_ebike, standard_bike, standard_ebike):
            continue
        shortage_bike = max(0, standard_bike - current_bike)
        shortage_ebike = max(0, standard_ebike - current_ebike)
        if not shortage_bike and not shortage_ebike:
            continue
        loc = locations.get(name, {})
        lat = _coordinate(loc.get("latitude"), 90) if isinstance(loc, dict) else None
        lon = _coordinate(loc.get("longitude"), 180) if isinstance(loc, dict) else None
        if lat is None or lon is None:
            continue
        pool.append({
            "name": name, "lat": lat, "lon": lon,
            "zone": str(row.get("路線區域") or "").strip(),
            "need_bike": shortage_bike, "need_ebike": shortage_ebike,
            "priority": name in priority_names,
        })
    return pool


def _layouts(bike, ebike):
    """Up to 14 small, deterministic layouts; never permute the 14-bike truck."""
    all_choices = []
    seen = set()

    def add(value):
        result = tuple(value)
        if result not in seen:
            seen.add(result)
            all_choices.append(list(result))

    for tail in ("bike", "ebike"):
        add(loading_order(bike, ebike, tail))
        for start in (tail, "ebike" if tail == "bike" else "bike"):
            counts = {"bike": bike, "ebike": ebike}
            result = []
            next_kind = start
            while counts["bike"] or counts["ebike"]:
                if counts[next_kind]:
                    result.append(next_kind)
                    counts[next_kind] -= 1
                next_kind = "ebike" if next_kind == "bike" else "bike"
            add(result)
        # Place up to three examples of each type at the tail, leaving the
        # rest grouped at the front; this covers mixed unloading demands.
        for count in (1, 2, 3):
            counts = {"bike": bike, "ebike": ebike}
            last = min(count, counts[tail])
            counts[tail] -= last
            other = "ebike" if tail == "bike" else "bike"
            add([tail] * counts[tail] + [other] * counts[other] + [tail] * last)
    return all_choices


def _forecast(order, origin, pool, visited, depth):
    """Beam-style lookahead. One planned removal per stage, no imagined pickups."""
    if depth <= 0 or not order:
        return 0.0, []
    steps = []
    for station in pool:
        if station["name"] in visited:
            continue
        remaining, unloaded, moves, held = plan_unload(
            order, station["need_bike"], station["need_ebike"]
        )
        count = unloaded["bike"] + unloaded["ebike"]
        if not count:
            continue
        proximity = _distance_km(origin, station)
        # Crossing region boundaries is intentionally expensive because a
        # straight line through the mountain is not a verified driving road.
        cross_zone = bool(origin.get("zone") and station["zone"]
                          and origin["zone"] != station["zone"])
        score = (8.0 * count - 1.25 * moves - .30 * proximity
                 + (6.0 if station["priority"] else 0.0)
                 - (18.0 if cross_zone else 0.0))
        steps.append((score, station, remaining, unloaded, moves, held))
    if not steps:
        return 0.0, []
    steps.sort(key=lambda item: (-item[0], item[1]["name"]))
    best_score, best_path = 0.0, []
    # Prune to five immediately useful candidates; at most three stages.
    for local, station, remaining, unloaded, moves, held in steps[:5]:
        follow, path = _forecast(
            remaining, station, pool, visited | {station["name"]}, depth - 1
        )
        candidate_score = local + .55 * follow
        if candidate_score > best_score:
            best_score = candidate_score
            best_path = [{
                "station_name": station["name"],
                "unload_bike": unloaded["bike"],
                "unload_ebike": unloaded["ebike"],
                "temporary_bikes": held,
                "handling_moves": moves,
                "priority": station["priority"],
                "proximity_km": round(_distance_km(origin, station), 1),
            }, *path]
    return best_score, best_path


def plan_loading_lookahead(remaining, pickup_bike, pickup_ebike, current_plan,
                           rows, station_locations, *, priority_names=(),
                           max_stops=3):
    """Return (pickup_order, projected_stops) with bounded local computations.

    It preserves the actual truck prefix; only newly loaded bikes are reordered.
    If coordinates or deficits are missing, return a safe grouped arrangement.
    No state is committed until the operator accepts the stop and completes it.
    """
    prefix = validate(remaining)
    bike, ebike = int(pickup_bike), int(pickup_ebike)
    if bike < 0 or ebike < 0:
        raise ValueError("上車數量不可為負數")
    if len(prefix) + bike + ebike > 100:
        raise ValueError("上車後排列超出安全限制")
    default = loading_order(bike, ebike)
    lat = _coordinate(current_plan.get("latitude"), 90)
    lon = _coordinate(current_plan.get("longitude"), 180)
    if lat is None or lon is None:
        return default, []
    current_name = str(current_plan.get("station_name") or "").strip()
    origin = {
        "name": current_name, "lat": lat, "lon": lon,
        "zone": str(current_plan.get("route_zone") or "").strip(),
    }
    pool = _station_pool(rows, station_locations, current_name, set(priority_names))
    if not pool:
        return default, []

    # Minimise total work by evaluating plausible suffixes, not brute-force
    # permutations of all carried bikes; at most 14 layouts/5 branches/3 steps.
    best_order, best_stops, best_score = default, [], float("-inf")
    for candidate in _layouts(bike, ebike):
        score, steps = _forecast(
            prefix + candidate, origin, pool, {current_name}, min(3, max(1, int(max_stops)))
        )
        # Deterministic tie-breaking avoids unexpected UI reorder on refresh.
        if score > best_score:
            best_order, best_stops, best_score = candidate, steps, score
    return best_order, best_stops
