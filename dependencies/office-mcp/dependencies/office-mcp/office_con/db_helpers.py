"""Shared MongoDB and Redis client helpers (no external dependencies)."""
from __future__ import annotations

import os
import threading

_mongo_async_cache: dict = {}
_mongo_async_lock = threading.Lock()

_mongo_sync_cache: dict = {}
_mongo_sync_lock = threading.Lock()

_redis_sync_cache: dict = {}
_redis_sync_lock = threading.Lock()


def get_async_mongo_client(url: str | None = None) -> "AsyncMongoClient[dict[str, object]]":
    """Return a shared AsyncMongoClient for the given URL."""
    from pymongo import AsyncMongoClient

    url = url or os.getenv("MONGODB_CONNECTION") or os.getenv("O365_MONGODB_URL")
    if not url:
        raise ValueError("MongoDB URL must be provided or set in MONGODB_CONNECTION env var")
    with _mongo_async_lock:
        if url not in _mongo_async_cache:
            _mongo_async_cache[url] = AsyncMongoClient(
                url,
                serverSelectionTimeoutMS=15000,
                connectTimeoutMS=10000,
                maxPoolSize=10,
            )
        return _mongo_async_cache[url]


def get_mongo_client(url: str | None = None) -> "MongoClient[dict[str, object]]":
    """Return a shared sync MongoClient for the given URL."""
    from pymongo import MongoClient

    url = url or os.getenv("MONGODB_CONNECTION") or os.getenv("O365_MONGODB_URL")
    if not url:
        raise ValueError("MongoDB URL must be provided or set in MONGODB_CONNECTION env var")
    with _mongo_sync_lock:
        if url not in _mongo_sync_cache:
            _mongo_sync_cache[url] = MongoClient(
                url,
                serverSelectionTimeoutMS=5000,
                connectTimeoutMS=5000,
                maxPoolSize=10,
            )
        return _mongo_sync_cache[url]


async def get_async_redis_client(url: str) -> object:
    """Return a shared async Redis client for the given URL.

    Supports both standalone Redis and Redis Cluster (detected via comma-separated URLs).
    """
    import asyncio
    from urllib.parse import urlparse, urlunparse

    url_stripped = url.rstrip(",")
    current_loop = asyncio.get_running_loop()
    key = (url_stripped, id(current_loop))
    if key not in _redis_sync_cache:
        parts = [p.strip() for p in url.split(",") if p.strip()]
        is_cluster = "," in url
        if is_cluster:
            from redis.asyncio.cluster import RedisCluster as AsyncRedisCluster, ClusterNode
            first = urlparse(parts[0] if "://" in parts[0] else f"redis://{parts[0]}")
            ssl = first.scheme == "rediss"
            password = first.password or None
            nodes = []
            for part in parts:
                node_url = part if "://" in part else urlunparse(
                    (first.scheme, part, "", "", "", "")
                )
                parsed = urlparse(node_url)
                nodes.append(ClusterNode(parsed.hostname or "localhost", parsed.port or (6380 if ssl else 6379)))
            client = AsyncRedisCluster(
                startup_nodes=nodes, password=password, ssl=ssl,
                ssl_check_hostname=False, decode_responses=False
            )
        else:
            import redis.asyncio as aioredis
            client = aioredis.from_url(url_stripped, decode_responses=False)  # type: ignore[assignment]
        _redis_sync_cache[key] = client
    return _redis_sync_cache[key]


def get_redis_client(url: str) -> object:
    """Return a shared sync Redis client for the given URL.

    Supports both standalone Redis and Redis Cluster (detected via comma-separated URLs).
    """
    from urllib.parse import urlparse, urlunparse

    url_stripped = url.rstrip(",")
    with _redis_sync_lock:
        if url_stripped not in _redis_sync_cache:
            parts = [p.strip() for p in url.split(",") if p.strip()]
            is_cluster = "," in url
            if is_cluster:
                from redis.cluster import RedisCluster, ClusterNode
                first = urlparse(parts[0] if "://" in parts[0] else f"redis://{parts[0]}")
                ssl = first.scheme == "rediss"
                password = first.password or None
                nodes = []
                for part in parts:
                    node_url = part if "://" in part else urlunparse(
                        (first.scheme, part, "", "", "", "")
                    )
                    parsed = urlparse(node_url)
                    nodes.append(ClusterNode(parsed.hostname or "localhost", parsed.port or (6380 if ssl else 6379)))
                _redis_sync_cache[url_stripped] = RedisCluster(  # type: ignore[assignment]
                    startup_nodes=nodes, password=password, ssl=ssl,
                    ssl_check_hostname=False, decode_responses=False
                )
            else:
                import redis as _redis
                _redis_sync_cache[url_stripped] = _redis.from_url(url_stripped, decode_responses=False)
        return _redis_sync_cache[url_stripped]
