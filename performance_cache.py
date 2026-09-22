"""Bounded caches for immutable code and static station identities, not live data."""
from __future__ import annotations

from collections import OrderedDict
from copy import deepcopy
from functools import lru_cache
from threading import RLock


@lru_cache(maxsize=2)
def compile_legacy_source(source: str, filename: str):
    """Exact source is the key; execute separately in each session's namespace."""
    return compile(source, filename, "exec", dont_inherit=True)


class CatalogMapCache:
    """Keep one catalog generation alive; never key solely on a recycled id()."""

    def __init__(self, max_entries: int = 16):
        self.max_entries = max_entries
        self._catalog = None
        self._entries = OrderedDict()
        self._lock = RLock()

    def resolve(self, clean_map, catalog, build):
        key = tuple(
            (zone, tuple((item["name"], item["district"]) for item in items))
            for zone, items in clean_map.items()
        )
        with self._lock:
            if self._catalog is not catalog:
                self._catalog = catalog
                self._entries.clear()
            if key not in self._entries:
                # build must only do local matching, never network requests.
                result = build(clean_map, catalog)
                self._entries[key] = deepcopy(result)
                while len(self._entries) > self.max_entries:
                    self._entries.popitem(last=False)
            self._entries.move_to_end(key)
            # Callers may annotate their own result without contaminating others.
            return deepcopy(self._entries[key])
