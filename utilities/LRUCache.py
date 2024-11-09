#
# Copyright ©2024 Dana Basken
#

from collections import OrderedDict
import time

class LRUCache:

    def __init__(self, size: int, ttl: int, on_event=None):
        self._size = size
        self._ttl = ttl
        self._on_event = on_event
        self._cache = OrderedDict()

    def get(self, key: str):
        self._age()
        if key not in self._cache:
            if self._on_event: self._on_event({"type": "miss", "key": key})
            return None
        value, timestamp = self._cache.pop(key)
        self._cache[key] = (value, time.time())
        if self._on_event: self._on_event({"type": "hit", "key": key})
        return value

    def put(self, key, value):
        self._age()
        if key in self._cache: self._cache.pop(key)
        elif len(self._cache) >= self._size: self._cache.popitem(last=False)
        if self._on_event: self._on_event({"type": "add", "key": key})
        self._cache[key] = (value, time.time())

    def _age(self):
        current_time = time.time()
        keys_to_delete = [key for key, (value, timestamp) in self._cache.items() if current_time - timestamp >= self._ttl]
        for key in keys_to_delete:
            del self._cache[key]
            if self._on_event: self._on_event({"type": "evict", "key": key})
