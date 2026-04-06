import asyncio
import os
import signal
import threading
import time
import logging
from datetime import datetime, timezone
from typing import Optional

import requests
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

# Global variables to store database configurations and health status
db_configs = {}
db_health_status: dict[str, dict] = {}
health_check_interval = 15  # seconds


class DatabaseConfig(BaseModel):
    """Configuration for a database connection."""
    name: str = Field(..., description="Unique name for this database connection")
    url: str = Field(..., description="Connection URL for the database")
    type: str = Field(..., description="Type of database (redis or mongodb)")
    timeout: float = Field(5.0, description="Connection timeout in seconds")


class HealthStatus(BaseModel):
    """Health status for a database connection."""
    name: str = Field(..., description="Name of the database connection")
    type: str = Field(..., description="Type of database (redis or mongodb)")
    status: str = Field(..., description="Status of the connection (healthy or unhealthy)")
    last_checked: str = Field(..., description="ISO timestamp of the last health check")
    error: Optional[str] = Field(None, description="Error message if the connection is unhealthy")


def create_health_response(status_str: str = "healthy", error: Optional[str] = None) -> dict:
    response = {
        "status": status_str,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "pid": os.getpid(),
        "port": os.environ.get("PORT", "8080"),
    }
    if db_health_status:
        response["databases"] = list(db_health_status.values())
    if error:
        response["error"] = error
    return response


def check_redis_health(config: DatabaseConfig) -> HealthStatus:
    """Check the health of the SHARED Redis connection (the one the app actually uses)."""
    try:
        from office_con.db_helpers import get_redis_client
        # Use the shared singleton — no kwargs = shared instance
        redis_client = get_redis_client(config.url)
        redis_client.ping()

        return HealthStatus(
            name=config.name,
            type="redis",
            status="healthy",
            last_checked=datetime.now(timezone.utc).isoformat(),
            error=None
        )
    except Exception as e:
        logger.error(f"Redis health check failed for {config.name} (shared connection): {e}")
        return HealthStatus(
            name=config.name,
            type="redis",
            status="unhealthy",
            last_checked=datetime.now(timezone.utc).isoformat(),
            error=str(e)
        )


def check_mongodb_health(config: DatabaseConfig) -> HealthStatus:
    """Check the health of the SHARED MongoDB connection (the one the app actually uses)."""
    try:
        from office_con.db_helpers import get_mongo_client
        # Use the shared singleton — same connection pool the app uses
        client = get_mongo_client(config.url)
        client.admin.command('ping')

        return HealthStatus(
            name=config.name,
            type="mongodb",
            status="healthy",
            last_checked=datetime.now(timezone.utc).isoformat(),
            error=None
        )
    except Exception as e:
        logger.error(f"MongoDB health check failed for {config.name} (shared connection): {e}")
        return HealthStatus(
            name=config.name,
            type="mongodb",
            status="unhealthy",
            last_checked=datetime.now(timezone.utc).isoformat(),
            error=str(e)
        )


def check_database_health(config: DatabaseConfig) -> HealthStatus:
    if config.type.lower() == "redis":
        return check_redis_health(config)
    elif config.type.lower() == "mongodb":
        return check_mongodb_health(config)
    else:
        return HealthStatus(
            name=config.name,
            type=config.type,
            status="unhealthy",
            last_checked=datetime.now(timezone.utc).isoformat(),
            error=f"Unsupported database type: {config.type}"
        )


def add_database(config: DatabaseConfig):
    db_configs[config.name] = config
    status = check_database_health(config)
    db_health_status[config.name] = status.model_dump()
    logger.info(f"Added database {config.name} to health check system")


def remove_database(name: str):
    if name in db_configs:
        del db_configs[name]
    if name in db_health_status:
        del db_health_status[name]
    logger.info(f"Removed database {name} from health check system")


def configure_database_health_checks() -> None:
    """Configure database health checks based on environment variables."""
    redis_url = os.environ.get("O365_REDIS_URL")
    if redis_url:
        add_database(DatabaseConfig(
            name="main-redis",
            url=redis_url,
            type="redis",
            timeout=5.0
        ))

    mongodb_url = os.environ.get("MONGODB_CONNECTION")
    if mongodb_url:
        add_database(DatabaseConfig(
            name="main-mongodb",
            url=mongodb_url,
            type="mongodb",
            timeout=5.0
        ))

    for key, value in os.environ.items():
        if key.startswith("MONGODB_CONNECTION_") and key != "MONGODB_CONNECTION":
            name = key[len("MONGODB_CONNECTION_"):].lower()
            add_database(DatabaseConfig(
                name=f"mongodb-{name}",
                url=value,
                type="mongodb",
                timeout=5.0
            ))


def database_health_check_worker() -> None:
    """Background thread that probes shared DB connections."""
    logger.info("Starting database health check worker")
    while True:
        try:
            for name, config in db_configs.items():
                status = check_database_health(config)
                db_health_status[name] = status.model_dump()
            time.sleep(health_check_interval)
        except Exception as e:
            logger.error(f"Database health check worker failed: {e}")
            time.sleep(health_check_interval)


def _kill_worker(reason: str) -> None:
    """Kill this worker process. Tries SIGTERM first, then SIGKILL after 5s."""
    pid = os.getpid()
    logger.critical(f"[HEALTH] {reason} — sending SIGTERM to pid={pid}")
    os.kill(pid, signal.SIGTERM)
    # Give uvicorn 5 seconds to shut down gracefully
    time.sleep(5)
    # If we're still alive, the event loop is blocked and can't process SIGTERM
    logger.critical(f"[HEALTH] Worker pid={pid} did not exit after SIGTERM — sending SIGKILL")
    os.kill(pid, signal.SIGKILL)


def health_check_worker() -> None:
    """Background thread that calls THIS worker's own /health endpoint.

    CRITICAL: We call http://localhost:{PORT}/health directly — NOT through nginx.
    This ensures we test OUR OWN event loop responsiveness. If we went through
    nginx (port 80), the request could be routed to a different worker, and a
    blocked worker would never detect its own problem.
    """
    port = os.environ.get("PORT", "8080")
    own_url = f"http://localhost:{port}/health"
    logger.info(f"Health check worker targeting own endpoint: {own_url}")

    # Give the server time to start up before first check
    time.sleep(15)

    consecutive_failures = 0
    max_failures = 2  # Allow 1 transient failure before killing

    while True:
        try:
            response = requests.get(own_url, timeout=10.0)

            if response.status_code != 200:
                consecutive_failures += 1
                logger.error(f"Health check failed with status {response.status_code} "
                             f"(attempt {consecutive_failures}/{max_failures})")
                if consecutive_failures >= max_failures:
                    _kill_worker("Health check failed")
                    return
            else:
                consecutive_failures = 0

            time.sleep(health_check_interval)

        except requests.Timeout:
            consecutive_failures += 1
            logger.error(f"Health check timed out — event loop likely blocked "
                         f"(attempt {consecutive_failures}/{max_failures})")
            if consecutive_failures >= max_failures:
                _kill_worker("Event loop blocked (health check timeout)")
                return
            time.sleep(health_check_interval)

        except Exception as e:
            consecutive_failures += 1
            logger.error(f"Health check worker error: {e} "
                         f"(attempt {consecutive_failures}/{max_failures})")
            if consecutive_failures >= max_failures:
                _kill_worker(f"Health check error: {e}")
                return
            time.sleep(health_check_interval)


def start_health_check_thread() -> None:
    """Start the background health check threads."""
    # Ensure DBs are configured (idempotent — safe to call multiple times)
    configure_database_health_checks()

    db_thread = threading.Thread(target=database_health_check_worker, daemon=True)
    db_thread.start()
    logger.info("Started database health check thread (shared connections)")

    thread = threading.Thread(target=health_check_worker, daemon=True)
    thread.start()
    port = os.environ.get("PORT", "8080")
    logger.info(f"Started event loop health check thread (self-check on port {port})")


async def health_check() -> dict:
    """Async health check — reads the latest DB probe results."""
    try:
        unhealthy_dbs = [
            status for status in db_health_status.values()
            if status["status"] == "unhealthy"
        ]
        if unhealthy_dbs:
            error_msg = ", ".join([
                f"{db['name']} ({db['error']})" for db in unhealthy_dbs
            ])
            return create_health_response("unhealthy", f"Unhealthy databases: {error_msg}")
        return create_health_response()
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return create_health_response("unhealthy", str(e))


async def health_endpoint() -> JSONResponse:
    """FastAPI /health endpoint. Runs on the event loop — if the event loop is blocked,
    this endpoint won't respond, and the self-check thread will detect the timeout."""
    try:
        # Run a fresh DB check inline (tests shared connections from event loop context)
        # Offload blocking sync Redis/MongoDB pings to a thread so we don't block the event loop
        for name, config in db_configs.items():
            status = await asyncio.to_thread(check_database_health, config)
            db_health_status[name] = status.model_dump()

        result = await health_check()

        if result["status"] == "healthy":
            return JSONResponse(content=result)
        else:
            return JSONResponse(status_code=503, content=result)
    except Exception as e:
        return JSONResponse(
            status_code=503,
            content=create_health_response("unhealthy", str(e))
        )
