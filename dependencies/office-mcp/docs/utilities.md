# Utilities

The `office_mcp.utils` package provides infrastructure-level modules that
support the office-mcp application at runtime. These utilities are not tied to
Microsoft Graph functionality; instead they handle operational concerns such as
service health monitoring and file format parsing.

## Table of Contents

- [Package Layout](#package-layout)
- [Health Check — `office_mcp.utils.health_check`](#health-check--office_mcputilshealth_check)
  - [Overview](#overview)
  - [Data Models](#data-models)
  - [Module-Level State](#module-level-state)
  - [Functions](#functions)
  - [Async Functions](#async-functions)
  - [Configuration (Environment Variables)](#configuration-environment-variables)
- [Excel Parser — `office_mcp.utils.file_formats.excel_parser`](#excel-parser--office_mcputilsfile_formatsexcel_parser)
  - [Dependencies](#dependencies)
  - [`ExcelParser` Class](#excelparser-class)
  - [Usage Examples](#usage-examples)
- [Full Integration Example](#full-integration-example)

## Package Layout

```text
office_mcp/utils/
├── __init__.py
├── health_check.py              # Service & database health monitoring
└── file_formats/
    ├── __init__.py
    └── excel_parser.py          # .xls / .xlsx spreadsheet parser
```

## Health Check — `office_mcp.utils.health_check`

The health check module provides a multi-layered monitoring system designed for
production FastAPI deployments behind nginx or a load balancer. It covers three
concerns:

1. **Database connectivity** — periodic background probes against shared
   Redis and MongoDB connections.
2. **Event loop liveness** — a self-check thread that calls the worker's own
   `/health` HTTP endpoint to detect a blocked asyncio event loop.
3. **Automatic recovery** — if the self-check detects consecutive failures,
   the worker process is terminated (`SIGTERM`, then `SIGKILL`) so that the
   process supervisor can restart it.

The module is imported and started in the application's lifespan hook; it
requires no additional configuration beyond the standard database environment
variables.

### Overview

```text
┌─────────────────────────────────────────────────────────┐
│                    FastAPI Worker                       │
│                                                        │
│  ┌──────────────────┐    ┌──────────────────────────┐  │
│  │  DB health thread │    │  Self-check thread       │  │
│  │  (daemon)         │    │  (daemon)                │  │
│  │                   │    │                          │  │
│  │  Every 15s:       │    │  Every 15s:              │  │
│  │  - ping Redis     │    │  - GET /health on own    │  │
│  │  - ping MongoDB   │    │    port (not nginx)      │  │
│  │  - update status  │    │  - on 2 consecutive      │  │
│  │    dict            │    │    failures: kill worker │  │
│  └──────────────────┘    └──────────────────────────┘  │
│                                                        │
│  ┌──────────────────────────────────────────────────┐  │
│  │  GET /health  (async endpoint on event loop)     │  │
│  │  - runs inline DB probes                         │  │
│  │  - returns 200 if all healthy, 503 otherwise     │  │
│  └──────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘
```

### Data Models

#### `DatabaseConfig`

Pydantic model describing a database connection to monitor.

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `name` | `str` | | Unique identifier for this database connection (e.g. `"main-redis"`). |
| `url` | `str` | | Connection URL for the database. |
| `type` | `str` | | Database engine type. Supported values: `"redis"`, `"mongodb"`. |
| `timeout` | `float` | `5.0` | Connection timeout in seconds. |

#### `HealthStatus`

Pydantic model representing the result of a single health probe.

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `name` | `str` | | Name of the database connection. |
| `type` | `str` | | Database type (`"redis"` or `"mongodb"`). |
| `status` | `str` | | Result of the check: `"healthy"` or `"unhealthy"`. |
| `last_checked` | `str` | | ISO 8601 timestamp of when the check was performed (UTC). |
| `error` | `Optional[str]` | `None` | Error message when `status` is `"unhealthy"`. |

### Module-Level State

The module maintains the following global variables:

| Variable | Purpose |
|----------|---------|
| `db_configs` | Dictionary mapping database names to their `DatabaseConfig` objects. |
| `db_health_status` | Dictionary mapping database names to their latest health probe results (serialized `HealthStatus` dicts). |
| `health_check_interval` | Interval in seconds between background probe cycles. Default: **15**. |

### Functions

#### `create_health_response(status_str, error)`

Build the JSON-serializable response dict returned by the `/health` endpoint.

| Parameter | Type | Description |
|-----------|------|-------------|
| `status_str` | `str` | Overall status, typically `"healthy"` or `"unhealthy"`. Default: `"healthy"`. |
| `error` | `Optional[str]` | Optional top-level error message. Default: `None`. |

**Returns:** A dict with keys `status`, `timestamp` (ISO 8601 UTC), `pid`,
`port`, and optionally `databases` (list of per-DB statuses) and `error`.

```python
>>> create_health_response()
{
    "status": "healthy",
    "timestamp": "2026-03-07T10:00:00+00:00",
    "pid": 12345,
    "port": "8080",
}
```

#### `check_redis_health(config)`

Probe the **shared** Redis connection obtained from
`office_mcp._db_helpers.get_redis_client`. This ensures the health check
tests the same connection pool that the application uses, rather than opening a
throwaway connection.

**Parameters:** `config` — a `DatabaseConfig` with `type="redis"`.

**Returns:** `HealthStatus`.

#### `check_mongodb_health(config)`

Probe the **shared** MongoDB connection obtained from
`office_mcp._db_helpers.get_mongo_client` by issuing an `admin.command('ping')`.

**Parameters:** `config` — a `DatabaseConfig` with `type="mongodb"`.

**Returns:** `HealthStatus`.

#### `check_database_health(config)`

Dispatch to the appropriate engine-specific health check based on
`config.type`. Returns an `"unhealthy"` status with an error message for
unsupported database types.

**Parameters:** `config` — a `DatabaseConfig`.

**Returns:** `HealthStatus`.

#### `add_database(config)`

Register a database for periodic health monitoring. The database is immediately
probed and its status is stored in `db_health_status`.

**Parameters:** `config` — a `DatabaseConfig`.

#### `remove_database(name)`

Remove a database from the health monitoring system.

**Parameters:** `name` — the string name of the database to remove.

#### `configure_database_health_checks()`

Auto-discover databases from environment variables and register them. This
function is called automatically by `start_health_check_thread()` and is
idempotent (safe to call multiple times).

The discovery logic:

1. If `O365_REDIS_URL` is set, register it as `"main-redis"` (type `redis`).
2. If `MONGODB_CONNECTION` is set, register it as `"main-mongodb"` (type `mongodb`).
3. Any additional `MONGODB_CONNECTION_<NAME>` variables are registered as
   `"mongodb-<name>"` (the suffix is lowercased).

#### `database_health_check_worker()`

Background thread target that continuously probes all registered databases at
`health_check_interval` intervals. Results are written into `db_health_status`.
This function runs in an infinite loop and is intended to be started as a daemon
thread.

#### `health_check_worker()`

Background thread target that performs **self-check** HTTP requests against the
worker's own `/health` endpoint.

Key design decisions:

- The request targets `http://localhost:{PORT}/health` **directly** — not
  through nginx or a load balancer. This ensures each worker tests its own
  event loop responsiveness.
- An initial 15-second delay allows the FastAPI server to finish startup before
  the first check.
- A 10-second HTTP timeout is used per request.
- After **2 consecutive failures** (non-200 status, timeout, or exception), the
  worker process is killed via `_kill_worker()`.

#### `_kill_worker(reason)`

Terminate the current worker process. Sends `SIGTERM` first to allow graceful
shutdown of uvicorn, then waits 5 seconds and sends `SIGKILL` if the process
is still alive (e.g., because the event loop is blocked and cannot process the
signal).

**Parameters:** `reason` — a string describing why the worker is being killed
(logged at `CRITICAL` level).

#### `start_health_check_thread()`

Entry point to activate the health monitoring system. Call this once during
application startup (typically in a FastAPI lifespan hook).

This function:

1. Calls `configure_database_health_checks()` to discover databases.
2. Starts the database probe daemon thread (`database_health_check_worker`).
3. Starts the self-check daemon thread (`health_check_worker`).

```python
from office_mcp.utils.health_check import start_health_check_thread

@asynccontextmanager
async def lifespan(app: FastAPI):
    start_health_check_thread()
    yield

app = FastAPI(lifespan=lifespan)
```

### Async Functions

#### `health_check()`

Coroutine that reads the latest background probe results from
`db_health_status` and builds a response dict. If any database is unhealthy,
the overall status is set to `"unhealthy"` with a summary of the failing
databases.

**Returns:** `dict` suitable for JSON serialization.

#### `health_endpoint()`

FastAPI route handler for `GET /health`. This coroutine:

1. Runs an **inline** (synchronous) database probe for all registered databases
   to test connectivity from the event loop context.
2. Calls `health_check()` to build the response.
3. Returns `200` with the health payload if all databases are healthy, or
   `503` if any are unhealthy or an exception occurs.

Because this handler runs on the event loop, a blocked event loop will prevent
it from responding — which is exactly what the self-check thread detects.

```python
from office_mcp.utils.health_check import health_endpoint

app.add_api_route("/health", health_endpoint, methods=["GET"])
```

### Configuration (Environment Variables)

| Variable | Description | Required |
|----------|-------------|----------|
| `PORT` | Port the FastAPI worker listens on. Used by the self-check thread to build `http://localhost:{PORT}/health`. Default: `8080`. | No |
| `O365_REDIS_URL` | Redis connection URL. If set, a Redis health check is registered automatically as `"main-redis"`. | No |
| `MONGODB_CONNECTION` | MongoDB connection string. If set, a MongoDB health check is registered automatically as `"main-mongodb"`. | No |
| `MONGODB_CONNECTION_<NAME>` | Additional MongoDB connections. Each variable with this prefix (excluding `MONGODB_CONNECTION` itself) is registered as `"mongodb-<name>"`. | No |

## Excel Parser — `office_mcp.utils.file_formats.excel_parser`

The Excel parser module provides a unified interface for extracting tabular data
from both legacy `.xls` (BIFF/Excel 97-2003) and modern `.xlsx`
(Office Open XML) spreadsheet files. It reads the **first worksheet** of the
workbook and normalizes all cell values to stripped strings.

### Dependencies

| Library | Purpose | File Formats |
|---------|---------|--------------|
| `openpyxl` | Parse `.xlsx` files (Office Open XML). | `.xlsx` |
| `xlrd` | Parse `.xls` files (BIFF / Excel 97-2003). | `.xls` |

### `ExcelParser` Class

```python
from office_mcp.utils.file_formats.excel_parser import ExcelParser

parser = ExcelParser(data=file_bytes, filename="report.xlsx", remove_blank=True)
for row in parser.values:
    print(row)
```

#### Constructor

```python
ExcelParser(data: bytes, filename: str, remove_blank: bool = False)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `data` | `bytes` | Raw file content as a byte string. This is typically obtained from an HTTP upload, a OneDrive/SharePoint download, or `open(..., "rb").read()`. |
| `filename` | `str` | Original filename. The file extension (`.xls` or `.xlsx`) determines which parsing backend is used. The comparison is case-insensitive. |
| `remove_blank` | `bool` | If `True`, rows and columns that are entirely empty are removed from the output after parsing. Default: `False`. |

The constructor performs all parsing immediately. After construction, the parsed
data is available in the `values` attribute.

**Raises:** `ValueError` if the filename does not end with `.xls` or `.xlsx`.

#### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `values` | `list[list[str]]` | Two-dimensional list of cell values. Every cell is a stripped string. `None` values are converted to empty strings (`""`). |

#### Methods

##### `parse_xlsx(data, filename)`

Parse an `.xlsx` file using `openpyxl`. Loads the workbook from an
in-memory `BytesIO` wrapper, reads the first sheet, and skips hidden rows
(rows with `row_dimensions[i].hidden == True`). Each cell value is converted
to a string; `None` values become empty strings.

This method is called automatically by the constructor when the filename ends
with `.xlsx`.

##### `parse_xls(data, filename)`

Parse a legacy `.xls` file using `xlrd`. Opens the workbook from raw bytes,
reads the first sheet, and determines the bounding rectangle of non-`None`
cells (last valid row and column). A uniformly sized 2D array is pre-allocated
and filled with cell values.

This method is called automatically by the constructor when the filename ends
with `.xls`.

##### `remove_empty_rows()`

Remove all rows from `values` where every cell is an empty string. Iterates
in reverse order to maintain correct indices during deletion.

This method is called by the constructor when `remove_blank=True`, after
parsing and before string normalization.

##### `remove_empty_columns()`

Remove all columns from `values` where every cell is an empty string. Iterates
column indices in reverse order and deletes the corresponding element from each
row.

This method is called by the constructor when `remove_blank=True`, after
`remove_empty_rows()` and before string normalization.

### Usage Examples

#### Parsing an uploaded Excel file

```python
from office_mcp.utils.file_formats.excel_parser import ExcelParser

# Read from disk
with open("/tmp/report.xlsx", "rb") as f:
    data = f.read()

parser = ExcelParser(data=data, filename="report.xlsx", remove_blank=True)

# The first row is typically a header
headers = parser.values[0]
print("Columns:", headers)

# Iterate data rows
for row in parser.values[1:]:
    record = dict(zip(headers, row))
    print(record)
```

#### Handling files from an HTTP upload (FastAPI)

```python
from fastapi import UploadFile
from office_mcp.utils.file_formats.excel_parser import ExcelParser

async def process_upload(file: UploadFile):
    content = await file.read()
    parser = ExcelParser(
        data=content,
        filename=file.filename,
        remove_blank=True,
    )
    return {"rows": len(parser.values), "columns": len(parser.values[0])}
```

#### Converting to a dictionary list

```python
parser = ExcelParser(data=raw_bytes, filename="data.xls")

headers = parser.values[0]
records = [dict(zip(headers, row)) for row in parser.values[1:]]

# records is now a list of dicts, one per data row
for rec in records:
    print(rec)
```

## Full Integration Example

The following example demonstrates starting the health check system and
registering the `/health` endpoint in a FastAPI application, alongside using
the Excel parser to process uploaded files.

```python
import os
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile

from office_mcp.utils.health_check import (
    start_health_check_thread,
    health_endpoint,
    add_database,
    DatabaseConfig,
)
from office_mcp.utils.file_formats.excel_parser import ExcelParser

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Start background health monitoring (auto-discovers DBs from env vars)
    start_health_check_thread()

    # Optionally register additional databases at runtime
    custom_redis = os.environ.get("CACHE_REDIS_URL")
    if custom_redis:
        add_database(DatabaseConfig(
            name="cache-redis",
            url=custom_redis,
            type="redis",
            timeout=3.0,
        ))

    yield

app = FastAPI(lifespan=lifespan)

# Wire up the health endpoint
app.add_api_route("/health", health_endpoint, methods=["GET"])

@app.post("/upload/excel")
async def upload_excel(file: UploadFile):
    content = await file.read()
    parser = ExcelParser(
        data=content,
        filename=file.filename,
        remove_blank=True,
    )
    return {
        "filename": file.filename,
        "rows": len(parser.values),
        "columns": len(parser.values[0]) if parser.values else 0,
        "headers": parser.values[0] if parser.values else [],
    }
```
