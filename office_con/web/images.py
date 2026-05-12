"""Authenticated image serving and browser-side image cache helpers."""

from __future__ import annotations

import asyncio
import inspect
import time
from dataclasses import dataclass
from typing import Awaitable, Callable, Protocol

from fastapi import APIRouter, Query, Request
from fastapi.responses import JSONResponse, Response

from office_con.db.company_dir import CompanyDir, generate_initials_avatar


class OfficeImageProvider(Protocol):
    """Provides the current company directory for a host application."""

    def __call__(self) -> CompanyDir | None:
        """Return the current directory, or None if it is not ready."""


OfficeImageAuthChecker = Callable[[Request], bool | Awaitable[bool]]

_VALID_IMAGE_SIZES = (32, 48, 64, 96, 128, 256, 512)


@dataclass(frozen=True)
class OfficeImageCacheConfig:
    """Cache and route settings for Office Connect image endpoints."""

    prefix: str = "/api/photo"
    ttl_seconds: int = 86_400
    default_size: int = 128


async def _check(callback: OfficeImageAuthChecker, request: Request) -> bool:
    result = callback(request)
    if inspect.isawaitable(result):
        result = await result
    return bool(result)


def _truthy(value: bool | None) -> bool:
    if isinstance(value, bool):
        return value
    return False


def _image_size(size: int | None, width: int | None, height: int | None, default: int) -> int:
    requested = size or width or height or default
    requested = max(_VALID_IMAGE_SIZES[0], min(_VALID_IMAGE_SIZES[-1], requested))
    for candidate in _VALID_IMAGE_SIZES:
        if requested <= candidate:
            return candidate
    return _VALID_IMAGE_SIZES[-1]


def _cache_headers(ttl_seconds: int, source: str, version: str) -> dict[str, str]:
    return {
        "Cache-Control": f"private, max-age={ttl_seconds}",
        "X-Office-Image-Source": source,
        "X-Office-Image-Cache-Ttl": str(ttl_seconds),
        "X-Office-Image-Cache-Version": version,
    }


def create_company_image_router(
    company_dir_provider: OfficeImageProvider,
    auth_checker: OfficeImageAuthChecker,
    *,
    admin_checker: OfficeImageAuthChecker | None = None,
    config: OfficeImageCacheConfig | None = None,
) -> APIRouter:
    """Create authenticated CompanyDir image routes.

    The host app owns authentication and admin semantics. This router only
    calls the supplied checkers and never exposes images to anonymous users.
    """

    cfg = config or OfficeImageCacheConfig()
    router = APIRouter(prefix=cfg.prefix.rstrip("/"))
    cache_version = str(int(time.time() * 1000))

    async def require_auth(request: Request) -> Response | None:
        if await _check(auth_checker, request):
            return None
        return Response(status_code=401, content="Not authenticated")

    async def require_admin(request: Request) -> Response | None:
        if admin_checker is None:
            return Response(status_code=403, content="Admin required")
        if await _check(admin_checker, request):
            return None
        return Response(status_code=403, content="Admin required")

    async def serve_user(
        user_id: str,
        display_name: str,
        *,
        size: int,
        real_only: bool,
    ) -> Response:
        directory = company_dir_provider()
        if directory is None:
            return Response(status_code=503, content="Directory not ready")

        blob = await directory.get_user_image_bytes_async(user_id)
        if blob:
            resized = await asyncio.to_thread(
                directory.get_user_image, user_id, width=size, height=size,
            )
            return Response(
                content=resized,
                media_type="image/jpeg",
                headers=_cache_headers(cfg.ttl_seconds, "real", cache_version),
            )

        if real_only:
            return Response(status_code=404, content="No real photo")

        avatar = await asyncio.to_thread(
            generate_initials_avatar, display_name, user_id, size, size,
        )
        return Response(
            content=avatar,
            media_type="image/jpeg",
            headers=_cache_headers(cfg.ttl_seconds, "fallback", cache_version),
        )

    @router.get("/cache-info")
    async def cache_info(request: Request):
        denied = await require_auth(request)
        if denied:
            return denied
        return JSONResponse(
            {
                "cache_name": "office-connect-images-v1",
                "ttl_seconds": cfg.ttl_seconds,
                "version": cache_version,
                "sizes": list(_VALID_IMAGE_SIZES),
            }
        )

    @router.post("/admin/cache-bust")
    async def cache_bust(request: Request):
        nonlocal cache_version
        denied = await require_auth(request)
        if denied:
            return denied
        forbidden = await require_admin(request)
        if forbidden:
            return forbidden
        cache_version = str(int(time.time() * 1000))
        return JSONResponse({"ok": True, "version": cache_version})

    @router.get("/by-email/{email}")
    async def image_by_email(
        email: str,
        request: Request,
        size: int | None = Query(default=None),
        w: int | None = Query(default=None),
        h: int | None = Query(default=None),
        real: bool | None = Query(default=None),
    ):
        denied = await require_auth(request)
        if denied:
            return denied
        directory = company_dir_provider()
        if directory is None:
            return Response(status_code=503, content="Directory not ready")
        user = directory.get_user_by_email(email)
        if user is None:
            return Response(status_code=404, content="Unknown email")
        display_name = user.display_name or email
        return await serve_user(
            user.id,
            display_name,
            size=_image_size(size, w, h, cfg.default_size),
            real_only=_truthy(real),
        )

    @router.get("/{user_id}")
    async def image_by_user_id(
        user_id: str,
        request: Request,
        size: int | None = Query(default=None),
        w: int | None = Query(default=None),
        h: int | None = Query(default=None),
        real: bool | None = Query(default=None),
    ):
        denied = await require_auth(request)
        if denied:
            return denied
        directory = company_dir_provider()
        if directory is None:
            return Response(status_code=503, content="Directory not ready")
        user = directory.get_user_by_id(user_id)
        display_name = user.display_name if user else user_id
        return await serve_user(
            user_id,
            display_name,
            size=_image_size(size, w, h, cfg.default_size),
            real_only=_truthy(real),
        )

    return router


def office_image_cache_client_script() -> str:
    """Return a browser helper for cached, authenticated Office images."""

    return r"""
(function() {
  if (window.OfficeImageCache) return;

  var CACHE_NAME = 'office-connect-images-v1';
  var META_KEY = 'office-connect-images-meta-v1';
  var DEFAULT_TTL_MS = 24 * 60 * 60 * 1000;
  var VALID_SIZES = [32, 48, 64, 96, 128, 256, 512];
  var objectUrls = new Set();

  function readMeta() {
    try { return JSON.parse(localStorage.getItem(META_KEY) || '{}'); }
    catch (e) { return {}; }
  }

  function writeMeta(meta) {
    try { localStorage.setItem(META_KEY, JSON.stringify(meta)); }
    catch (e) {}
  }

  function normalizeSize(value) {
    var requested = Number(value || 128);
    if (!Number.isFinite(requested)) requested = 128;
    requested = Math.max(VALID_SIZES[0], Math.min(VALID_SIZES[VALID_SIZES.length - 1], requested));
    for (var i = 0; i < VALID_SIZES.length; i++) {
      if (requested <= VALID_SIZES[i]) return VALID_SIZES[i];
    }
    return VALID_SIZES[VALID_SIZES.length - 1];
  }

  function withParams(url, opts) {
    var u = new URL(url, window.location.origin);
    u.searchParams.set('size', String(normalizeSize(opts.size)));
    if (opts.realOnly) u.searchParams.set('real', '1');
    return u.pathname + u.search;
  }

  async function responseToObjectUrl(response) {
    var blob = await response.blob();
    var url = URL.createObjectURL(blob);
    objectUrls.add(url);
    return url;
  }

  async function resolve(url, options) {
    var opts = options || {};
    var ttlMs = Number(opts.ttlMs || DEFAULT_TTL_MS);
    var key = withParams(url, opts);
    if (!('caches' in window)) {
      var direct = await fetch(key, { credentials: 'same-origin' });
      if (!direct.ok) return '';
      return responseToObjectUrl(direct);
    }
    var meta = readMeta();
    var now = Date.now();
    var cache = await caches.open(CACHE_NAME);
    var cached = await cache.match(key);
    if (cached && meta[key] && now - meta[key] < ttlMs) {
      return responseToObjectUrl(cached);
    }
    var fetched = await fetch(key, { credentials: 'same-origin', cache: 'no-store' });
    if (!fetched.ok) {
      await cache.delete(key);
      delete meta[key];
      writeMeta(meta);
      return '';
    }
    await cache.put(key, fetched.clone());
    meta[key] = now;
    writeMeta(meta);
    return responseToObjectUrl(fetched);
  }

  function byEmail(email, options) {
    var opts = options || {};
    var prefix = (opts.prefix || '/api/photo').replace(/\/+$/, '');
    return resolve(prefix + '/by-email/' + encodeURIComponent(email), opts);
  }

  function byUserId(userId, options) {
    var opts = options || {};
    var prefix = (opts.prefix || '/api/photo').replace(/\/+$/, '');
    return resolve(prefix + '/' + encodeURIComponent(userId), opts);
  }

  async function clear() {
    objectUrls.forEach(function(url) { URL.revokeObjectURL(url); });
    objectUrls.clear();
    if ('caches' in window) await caches.delete(CACHE_NAME);
    localStorage.removeItem(META_KEY);
  }

  async function adminWipe(options) {
    var opts = options || {};
    var prefix = (opts.prefix || '/api/photo').replace(/\/+$/, '');
    await fetch(prefix + '/admin/cache-bust', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: '{}'
    }).catch(function() {});
    await clear();
  }

  window.OfficeImageCache = {
    byEmail: byEmail,
    byUserId: byUserId,
    clear: clear,
    adminWipe: adminWipe,
    normalizeSize: normalizeSize
  };
})();
"""
