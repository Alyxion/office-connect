# Company Directory

The company directory module provides a high-level abstraction over the
Microsoft Graph directory API. It fetches, enriches, caches, and serves
user records and profile photos for an entire Azure AD tenant.

The module lives in `office_mcp.db.company` and exposes three public
classes plus one standalone helper:

| Symbol | Purpose |
| --- | --- |
| `CompanyDir` | Main facade -- user lookups, image serving, search, DataFrame export. |
| `LiveCompanyDirData` | Live backend that populates users from MS Graph in a background task and fetches photos on demand. |
| `CompanyDirBuilder` | Offline / batch builder that downloads users and photos to disk. |
| `apply_user_replacements()` | Standalone function that applies replacement rules to a user list in-place. |

## Overview

There are two distinct workflows for obtaining a company directory:

1. **Live mode** -- `LiveCompanyDirData` fetches all users asynchronously
   on first authentication, keeps them in memory, and fetches profile photos
   on demand from MS Graph. The `CompanyDir` facade delegates to this
   backend at runtime. This is the recommended mode for web applications.

2. **Batch / offline mode** -- `CompanyDirBuilder` downloads every user
   and every profile photo to a local directory (`users.json` plus a
   `photos/` folder). The resulting `CompanyDirData` can then be loaded
   into `CompanyDir` for purely local operation without further network
   access.

Both modes share the same data model (`CompanyUser`, `CompanyUserList`)
and the same replacement / enrichment pipeline (`UserReplacements`).

## Data Model

### DirectoryUser (base)

`DirectoryUser` is defined in `office_mcp.msgraph.directory_handler` and
represents the raw record returned by MS Graph.

| Field | Type | Description |
| --- | --- | --- |
| `id` | `str` | Azure AD object UUID (required). |
| `display_name` | `str \| None` | Full display name, e.g. "Jane Doe". |
| `email` | `str \| None` | Primary email (`mail` or `userPrincipalName`). |
| `job_title` | `str \| None` | Job title, e.g. "Sales Manager". |
| `department` | `str \| None` | Department, e.g. "Sales". |
| `manager_id` | `str \| None` | UUID of the user's manager. |
| `account_enabled` | `bool \| None` | Whether the Azure AD account is enabled. |
| `surname` | `str \| None` | Family name. |
| `given_name` | `str \| None` | Given (first) name. |
| `office_location` | `str \| None` | Office location string. |
| `mobile_phone` | `str \| None` | Mobile phone number. |

All fields except `id` are optional and default to `None`.

### CompanyUser

`CompanyUser` extends `DirectoryUser` with company-specific metadata.
It is a Pydantic `BaseModel`, so all standard serialization
(`model_dump()`, `model_validate()`, JSON schema) works out of the box.

#### Additional fields

| Field | Type | Description |
| --- | --- | --- |
| `company` | `str \| None` | Company / legal entity name. |
| `external` | `bool` | `True` for external (guest) accounts. Default `False`. |
| `gender` | `Literal["undefined", "male", "female", "other"]` | Gender, optionally guessed from the first name. |
| `object_type` | `str` | Discriminator for the kind of directory object (see constants below). Default `"user"`. |
| `building` | `str \| None` | Building name or identifier. |
| `street` | `str \| None` | Street address. |
| `manager_email` | `str \| None` | Email address of the user's manager (enriched field). |
| `zip` | `str \| None` | Postal / ZIP code. |
| `city` | `str \| None` | City. |
| `country` | `str \| None` | Country. |
| `room_name` | `str \| None` | Room or desk identifier. |
| `guessed_fields` | `dict[str, bool]` | Tracks which fields were populated by heuristic guessing (as opposed to authoritative data). Keys are field names. |
| `has_image` | `bool` | `True` if a profile photo is known to exist. |
| `join_date` | `str \| None` | Date the user joined the organization. |
| `birth_date` | `str \| None` | Date of birth. |
| `termination_date` | `str \| None` | Date the user left / will leave the organization. |

#### Object type constants

`CompanyUser` defines class-level constants for the `object_type` field.
The value may also contain a sub-type separated by a dot
(e.g. `"serviceAccount.automation"`).

```python
CompanyUser.OT_USER              # "user"
CompanyUser.OT_SERVICE_ACCOUNT   # "serviceAccount"
CompanyUser.OT_ROOM              # "room"
CompanyUser.OT_DEVICE            # "device"
CompanyUser.OT_GROUP             # "group"
CompanyUser.OT_APPLICATION       # "application"
CompanyUser.OT_CALENDAR          # "calendar"
CompanyUser.OT_SERVICE_PRINCIPAL # "servicePrincipal"
```

Use the `main_type` and `sub_type` properties to decompose compound
types:

```python
user.object_type = "serviceAccount.automation"
user.main_type   # "serviceAccount"
user.sub_type    # "automation"

user.object_type = "user"
user.main_type   # "user"
user.sub_type    # None
```

#### Rule matching

`CompanyUser.matches_rule(rule, use_guessed=True)` tests whether a user
matches a dictionary-based filter rule. Rules are used extensively by the
replacement pipeline (see [UserReplacements](#userreplacements)).

A rule is a `dict` that must contain exactly one key starting with `_`
(the **mask key**). The mask key selects which fields to match against and
the corresponding value is a glob pattern (case-insensitive, using
`fnmatch`).

| Mask key | Behaviour |
| --- | --- |
| `_` | Match the pattern against **all** matchable string fields (catch-all). |
| `_<field>` | Match only the named field, e.g. `_department`, `_email`. |

Multiple alternative patterns can be separated with `|` (logical OR):

```python
# Matches users whose department is "Sales" OR "Marketing"
{"_department": "sales|marketing", "external": True}
```

When `use_guessed=False`, fields listed in `guessed_fields` are
excluded from matching.

### CompanyUserList

A thin wrapper used for serialization:

```python
class CompanyUserList(BaseModel):
    CURRENT_VERSION: ClassVar[str] = "1.0"
    version: str = "1.0"
    users: List[CompanyUser] = []
```

The `version` field enables future schema migrations.

## CompanyDir (facade)

`CompanyDir` is the primary entry point for consuming directory data. It
wraps either a `CompanyDirData` (static / in-memory) or a
`LiveCompanyDirData` (live MS Graph) backend and provides a unified API
for lookups, search, image serving, and export.

### Construction

```python
from office_mcp.db.company import CompanyDir, LiveCompanyDirData

# --- Live mode (recommended for web apps) ---
live_data = LiveCompanyDirData(user_replacements=replacements, add_genders=True)
directory = CompanyDir(source=live_data)

# --- Static mode (from a CompanyDirData or path) ---
directory = CompanyDir(source=data_object)

# --- Empty directory ---
directory = CompanyDir()
```

Parameters:

| Parameter | Description |
| --- | --- |
| `source` | A `LiveCompanyDirData`, `CompanyDirData`, path string, or `None`. |
| `user_image_url_callback` | Optional callable `(user_id, ext, width, height) -> url` for custom photo URL generation. |
| `in_memory` | Passed to `CompanyDirData` when `source` is a path. |

### User lookups

```python
# By Azure AD object ID (O(1) dict lookup)
user = directory.get_user_by_id("aabbccdd-1234-...")

# By email, case-insensitive (O(1) dict lookup)
user = directory.get_user_by_email("jane.doe@example.com")

# Full list
all_users = directory.users  # List[CompanyUser]
```

### Searching users

`find_users()` performs a multi-criteria search. Name fields use
case-insensitive substring matching by default; email always uses exact
matching.

```python
# Substring match on display name
results = directory.find_users(display_name="Doe")

# Exact match on first + last name
results = directory.find_users(first_name="Jane", last_name="Doe", exact=True)

# By email (always exact)
results = directory.find_users(email="jane.doe@example.com")

# First match shortcut
user = directory.get_user(display_name="Doe")
```

Parameters for `find_users()`:

| Parameter | Type | Description |
| --- | --- | --- |
| `display_name` | `str` | Substring (or exact) match on `display_name`. |
| `email` | `str` | Case-insensitive exact match on `email`. |
| `first_name` | `str` | Substring (or exact) match on `given_name`. |
| `last_name` | `str` | Substring (or exact) match on `surname`. |
| `exact` | `bool` | If `True`, name fields require an exact match instead of substring. |

### Profile images

`CompanyDir` provides two image-related methods:

**`get_user_image(user_id, *, width=256, height=256)`**

Returns JPEG bytes of the user's profile photo, cropped to a square and
resized to the requested dimensions. If no photo exists, a
deterministic pastel-colored **initials avatar** is generated instead.

Valid resolutions: `32`, `48`, `64`, `96`, `128`, `256`,
`512`. Passing any other value raises `ValueError`.

Resized images are cached in memory (keyed by
`"{user_id}_{width}_{height}"`). Initials avatars are intentionally
**not** cached so that a real photo can be picked up once it becomes
available.

**`get_user_image_url(user_id, ext='.jpg', *, width=256, height=256)`**

Returns the URL where the photo can be retrieved. Delegates to the
`user_image_url_callback` if one was provided; otherwise builds a
URL from the `WEBSITE_URL` environment variable:

```
{WEBSITE_URL}/profiles/{user_id}.jpg
```

#### Initials avatar generation

When no photo is available, `generate_initials_avatar()` creates a JPEG
image with the user's initials on a pastel background:

- Initials are extracted from the first and last parts of `display_name`.
- The background hue is derived deterministically from `user_id`
  (MD5-based), so the same user always gets the same color.
- Font resolution cascades through Helvetica (macOS), DejaVu Sans Bold
  (Debian/Ubuntu), and DejaVu Sans Bold (RHEL/Fedora), falling back to
  Pillow's built-in default font.

### DataFrame export

`users_to_dataframe(rules, use_guessed=True)` filters users by one or
more match rules and returns a `pandas.DataFrame` with all fields
flattened to strings.

```python
import pandas as pd

# All users in the "Sales" department
df = directory.users_to_dataframe({"_department": "sales"})

# Multiple rules (OR logic)
df = directory.users_to_dataframe([
    {"_department": "sales"},
    {"_department": "marketing"},
])
```

In the output, `manager_id` (a UUID) is replaced by `manager_email`
for readability. Nested dicts are JSON-serialized; lists are
comma-joined.

### Refreshing

`CompanyDir.refresh()` rebuilds the internal lookup indices
(`_user_by_id`, `_user_by_email`) from the underlying data source and
clears the resized-image cache. In live mode, `LiveCompanyDirData`
registers this as the `on_done` callback so the facade is kept in sync
automatically after background population completes.

## LiveCompanyDirData

`LiveCompanyDirData` is the recommended backend for production web
applications. It populates users asynchronously, stores them in memory,
and fetches profile photos on demand from MS Graph.

All public methods are **thread-safe** (guarded by `threading.RLock`).

### Construction

```python
from office_mcp.db.company import LiveCompanyDirData
from office_mcp.db.company.company_dir_builder import UserReplacements

replacements = UserReplacements(
    guesses=[...],
    data=[...],
)
live = LiveCompanyDirData(user_replacements=replacements, add_genders=True)
```

Parameters:

| Parameter | Description |
| --- | --- |
| `user_replacements` | Optional `UserReplacements` applied after fetching users from Graph. |
| `add_genders` | If `True`, genders are guessed from first names using the `names_dataset` library. |

### Population lifecycle

1. **Start population** -- call `start_population(graph_instance,
   on_done=callback)` once after the first OAuth authentication succeeds.
   This creates a `DirectoryHandler`, then spawns an async task on the
   current event loop that:

   a. Calls `DirectoryHandler.get_all_users_async()` to page through all
      Azure AD users.
   b. Converts each `DirectoryUser` to a `CompanyUser`.
   c. Applies `UserReplacements` (guesses, then definitive data).
   d. Scans the in-memory photo cache to restore `has_image` flags.
   e. Atomically swaps the user list and index dicts under the lock.
   f. Invokes the `on_done` callback (typically `CompanyDir.refresh()`).

2. **Subsequent logins** -- call `refresh_graph(graph_instance)` to keep
   the Graph access token fresh without re-triggering population.

3. **Status check** -- the `is_populated` property returns `True` once
   the background task has completed successfully.

```python
# Typical integration in a login handler
if not live.is_populated:
    live.start_population(graph, on_done=directory.refresh)
else:
    live.refresh_graph(graph)
```

### Photo handling

Photos are never bulk-downloaded in live mode. Instead:

**`prefetch_photo_async(user_id)`**

Fetches a single user's photo from MS Graph and stores the result
(present or absent) in `_photo_known`. This should be called during
login so that `has_image` is accurate by the time the page renders.
Skips the fetch if the photo status is already known.

**`get_image_bytes_async(user_id)`**

On-demand photo fetch. Returns the raw bytes or `None`. Updates
`_photo_known` and the user's `has_image` flag.

**`check_photo_cache(user_id)`**

Synchronous check against the in-memory cache. Returns `True` if the
photo status is already known (no network call needed), `False`
otherwise. Sets `has_image` on the user object if known.

**`get_image_bytes(user_id)`**

Always returns `None` -- exists for interface compatibility with
`CompanyDirData`. Use the async variants for actual photo retrieval.

The `_photo_known` dict survives population swaps, so a photo that was
fetched for a logged-in user does not need to be re-fetched after a
background refresh.

### User lookups

`LiveCompanyDirData` provides O(1) lookups under the lock:

```python
user = live.get_user_by_id("aabbccdd-1234-...")
user = live.get_user_by_email("jane.doe@example.com")
```

### Thread safety

All mutable state is protected by a single `threading.RLock`. The
background population task performs an **atomic swap** of three data
structures (`_users`, `_user_by_id`, `_user_by_email`) under the
lock, so readers never see a partially populated directory.

## CompanyDirBuilder

`CompanyDirBuilder` is designed for batch / offline scenarios where the
entire directory (users and photos) is downloaded to the local filesystem.

### Construction

```python
from pathlib import Path
from office_mcp.db.company.company_dir_builder import CompanyDirBuilder, UserReplacements

builder = CompanyDirBuilder(
    target_dir=Path("/data/company_dir"),
    msgraph=graph_instance,
    logger=logging.getLogger("builder"),
    clear_images=True,
    user_replacements=UserReplacements(guesses=[...], data=[...]),
    add_genders=True,
)
```

Parameters:

| Parameter | Description |
| --- | --- |
| `target_dir` | `Path` to the output directory. |
| `msgraph` | Authenticated `MsGraphInstance`. |
| `logger` | Logger for progress and error messages. |
| `clear_images` | If `True` (default), removes existing `users.json` and `photos/` before starting. |
| `user_replacements` | `UserReplacements` rules to apply after fetching. |
| `add_genders` | Guess gender from first names. |

### Build process

Call `await builder.build()` to execute the full pipeline:

1. Fetch all users from MS Graph via paginated `get_all_users_async()`.
2. Convert each `DirectoryUser` to a `CompanyUser`.
3. Iterate over every user and download their profile photo:

   - If `<user_id>.jpg` or `__<user_id>.ph` (placeholder) already
     exists on disk, skip the download.
   - If the Graph API returns a photo, save it as
     `<target_dir>/photos/<user_id>.jpg` and set `has_image = True`.
   - If no photo exists, write an empty placeholder file
     `__<user_id>.ph` to avoid re-fetching on subsequent runs.

4. Apply `UserReplacements` (guesses first, then definitive data;
   optionally guess genders).
5. Write the final `users.json` to `<target_dir>/users.json`.

```python
await builder.build()
# Output:
#   /data/company_dir/users.json
#   /data/company_dir/photos/<uuid>.jpg
#   /data/company_dir/photos/__<uuid>.ph   (no-photo placeholders)
```

### Output directory structure

```text
<target_dir>/
    users.json                    # CompanyUserList serialized as JSON
    photos/
        aabbccdd-1234-....jpg     # Profile photos (JPEG)
        __eeffgghh-5678-....ph    # Empty placeholder (no photo available)
```

### Clearing data

`builder.clear()` removes `users.json` and recursively deletes the
`photos/` directory. It is called automatically when `clear_images=True`
(the default) during construction.

## UserReplacements

`UserReplacements` is a Pydantic model that defines two ordered lists of
transformation rules:

```python
class UserReplacements(BaseModel):
    guesses: List[Dict[str, Any]] = []
    data: List[Dict[str, Any]] = []
```

**`guesses`**

Applied first. When a rule matches a user, its non-mask fields are set
on the user **and** the field name is added to `guessed_fields`. This
marks the value as heuristic and allows downstream logic to distinguish
authoritative data from guesses.

**`data`**

Applied second (after guesses and optional gender guessing). These are
**definitive** overrides. If a field was previously guessed, it is
removed from `guessed_fields` when overwritten by a data rule.

Each rule dict must contain exactly one mask key (prefixed with `_`)
for matching, plus one or more plain keys for the values to set:

```python
replacements = UserReplacements(
    guesses=[
        # Guess: anyone with "intern" in the job title is external
        {"_job_title": "*intern*", "external": True},
        # Guess: users in Office K are in building "HQ"
        {"_office_location": "office k", "building": "HQ"},
    ],
    data=[
        # Definitive: set company for all users matching a department
        {"_department": "sales", "company": "Acme Corp"},
        # Definitive: override a specific user by email
        {"_email": "bot@example.com", "object_type": "serviceAccount"},
    ],
)
```

### apply_user_replacements()

```python
from office_mcp.db.company.company_dir_builder import apply_user_replacements

apply_user_replacements(user_list, replacements, add_genders=False)
```

This standalone function applies the full replacement pipeline to a
`CompanyUserList` **in-place**:

1. Apply `guesses` rules (fields marked in `guessed_fields`).
2. Remove users with empty `id` strings.
3. If `add_genders=True`, run the gender-guessing heuristic.
4. Apply `data` rules (fields removed from `guessed_fields`).

## Gender guessing

When `add_genders=True`, the pipeline uses the `names_dataset` library
to infer gender from first names. The logic handles several edge cases:

- Academic titles (`Prof.`, `Dr.`) are stripped.
- Hyphenated first names use only the first component.
- For names with a European name in parentheses (common in Chinese naming
  conventions, e.g. "Zhang Wei (David)"), the European name is preferred
  for lookup.
- If the initial lookup returns `"other"`, a second attempt is made with
  the raw `given_name`.

All gender values set by this heuristic are marked in `guessed_fields`
so they can be distinguished from authoritative data.

## MongoDB Collection Structure

When `CompanyDirBuilder` writes to disk, the canonical on-disk format is
a single JSON file (`users.json`) containing a serialized
`CompanyUserList`.

The JSON structure is:

```json
{
  "version": "1.0",
  "users": [
    {
      "id": "aabbccdd-1234-5678-9abc-def012345678",
      "display_name": "Jane Doe",
      "email": "jane.doe@example.com",
      "job_title": "Sales Manager",
      "department": "Sales",
      "manager_id": "11223344-...",
      "account_enabled": true,
      "surname": "Doe",
      "given_name": "Jane",
      "office_location": "Office K",
      "mobile_phone": "+49 123 456 789",
      "company": "Acme Corp",
      "external": false,
      "gender": "female",
      "object_type": "user",
      "building": "HQ",
      "street": "Main Street 1",
      "manager_email": "boss@example.com",
      "zip": "12345",
      "city": "Springfield",
      "country": "DE",
      "room_name": "3.14",
      "guessed_fields": {"gender": true},
      "has_image": true,
      "join_date": "2020-01-15",
      "birth_date": null,
      "termination_date": null
    }
  ]
}
```

When the directory is consumed by a MongoDB-backed application, each user
document maps directly to the `CompanyUser` schema above. A typical
MongoDB collection layout would store each user as an individual document
with the same field names, enabling indexed queries on `email`,
`department`, `object_type`, etc.

## Complete Usage Example

The following example demonstrates the full live-mode lifecycle: creating
the backend, wiring it into `CompanyDir`, starting population, and
querying the directory.

```python
import asyncio
import logging
from office_mcp.db.company import CompanyDir, LiveCompanyDirData
from office_mcp.db.company.company_dir_builder import UserReplacements

log = logging.getLogger(__name__)

# 1. Define replacement rules
replacements = UserReplacements(
    guesses=[
        {"_office_location": "office k", "building": "HQ"},
    ],
    data=[
        {"_email": "bot@example.com", "object_type": "serviceAccount"},
    ],
)

# 2. Create the live backend and the facade
live = LiveCompanyDirData(user_replacements=replacements, add_genders=True)
directory = CompanyDir(source=live)

# 3. On first login, start background population
#    (graph_instance is an authenticated MsGraphInstance)
live.start_population(graph_instance, on_done=directory.refresh)

# 4. On subsequent logins, just refresh the token
live.refresh_graph(graph_instance)

# 5. Prefetch the logged-in user's photo during login
await live.prefetch_photo_async(current_user_id)

# 6. Query the directory once populated
if live.is_populated:
    user = directory.get_user_by_email("jane.doe@example.com")
    if user:
        log.info("Found: %s (%s)", user.display_name, user.department)

    # Get a 128x128 JPEG (real photo or initials fallback)
    image_bytes = directory.get_user_image(user.id, width=128, height=128)

    # Search by partial name
    matches = directory.find_users(last_name="Doe")

    # Export filtered users to a DataFrame
    df = directory.users_to_dataframe({"_department": "sales"})
    print(df[["display_name", "email", "job_title"]].to_string())
```

## Class Diagram

```text
DirectoryUser (Pydantic BaseModel)
     |
     v
CompanyUser (extends DirectoryUser)
     |
     +--- CompanyUserList (wrapper with version)
     |
     +---> CompanyDirData (static in-memory container)
     |          |
     |          v
     +---> LiveCompanyDirData (live MS Graph backend)
     |          |
     |          +--- DirectoryHandler (Graph API calls)
     |          +--- UserReplacements (enrichment rules)
     |
     +---> CompanyDirBuilder (batch download to disk)
     |          |
     |          +--- MsGraphInstance (Graph auth)
     |          +--- UserReplacements
     |
     v
CompanyDir (facade)
     |
     +--- get_user_by_id / get_user_by_email
     +--- find_users / get_user
     +--- get_user_image / get_user_image_url
     +--- users_to_dataframe
     +--- refresh
```

## Source Files

| File | Contents |
| --- | --- |
| `office_mcp/db/company/company_dir.py` | `CompanyUser`, `CompanyUserList`, `CompanyDirData`, `LiveCompanyDirData`, `CompanyDir`, `generate_initials_avatar()`. |
| `office_mcp/db/company/company_dir_builder.py` | `UserReplacements`, `CompanyDirBuilder`, `apply_user_replacements()`, `_guess_genders_sync()`. |
| `office_mcp/msgraph/directory_handler.py` | `DirectoryUser`, `DirectoryUserList`, `DirectoryHandler`. |
| `office_mcp/db/company/__init__.py` | Public re-exports: `CompanyDir`, `LiveCompanyDirData`, `CompanyDirBuilder`. |
| `office_mcp/db/__init__.py` | Top-level re-exports including `apply_user_replacements` and `UserReplacements`. |
