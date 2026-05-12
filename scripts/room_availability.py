#!/usr/bin/env python3
"""Generate an HTML room availability overview from MS Graph.

Reads config from ``tests/test_config.json`` (or ``OFFICE_CONNECT_TEST_CONFIG``).

Config ``room_availability`` section::

    {
        "token_file": "~/Downloads/token_export.json",
        "room_availability": {
            "title": "Meeting Rooms",
            "domain": "@example.com",
            "exclude": ["silentroom", "training"],
            "future_rooms": ["building-x"],
            "hours_start": 7,
            "hours_end": 20,
            "timezone": "W. Europe Standard Time",
            "output": "~/Downloads/room_availability.html"
        }
    }

Usage::

    poetry run python scripts/room_availability.py
"""

from __future__ import annotations

import asyncio
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

# Ensure the package is importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from office_con.testing.mock_config import get_test_config, get_token_data

# Map Windows timezone names to IANA (common ones)
_WIN_TO_IANA = {
    "W. Europe Standard Time": "Europe/Berlin",
    "Central European Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris",
    "GMT Standard Time": "Europe/London",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Pacific Standard Time": "America/Los_Angeles",
    "China Standard Time": "Asia/Shanghai",
    "India Standard Time": "Asia/Kolkata",
    "Tokyo Standard Time": "Asia/Tokyo",
    "Singapore Standard Time": "Asia/Singapore",
}


def _to_iana(win_tz: str) -> str:
    """Convert a Windows timezone name to IANA, or return as-is if already IANA."""
    return _WIN_TO_IANA.get(win_tz, win_tz)


async def main():
    from office_con.msgraph.ms_graph_handler import MsGraphInstance
    from office_con.msgraph.places_handler import PlacesHandler
    from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler

    cfg = get_test_config()
    token_data = get_token_data()
    if not token_data:
        print("No token file configured. Create tests/test_config.json with 'token_file'.")
        sys.exit(1)

    ra = cfg.get("room_availability", {})
    title = ra.get("title", "Meeting Rooms")
    domain = ra.get("domain", "")
    exclude = [x.lower() for x in ra.get("exclude", [])]
    future_rooms = [x.lower() for x in ra.get("future_rooms", [])]
    hours_start = ra.get("hours_start", 7)
    hours_end = ra.get("hours_end", 20)
    win_tz = ra.get("timezone", "W. Europe Standard Time")
    output = Path(ra.get("output", "~/Downloads/room_availability.html")).expanduser()

    graph = MsGraphInstance(scopes=None, endpoint="https://graph.microsoft.com/v1.0/")
    graph.cache_dict = token_data
    graph.email = token_data.get("email", "")
    graph.client_id = token_data.get("client_id", "")
    graph.client_secret = token_data.get("client_secret", "")
    graph.tenant_id = token_data.get("tenant_id", "")

    # Refresh token if expired
    token = token_data.get("access_token")
    ep = graph.msg_endpoint
    test = await graph.run_async(url=ep + "me?$select=displayName", token=token)
    if test.status_code == 401:
        print("Token expired, refreshing...")
        token = await graph.refresh_token_async()
        if not token:
            print("Token refresh failed. Export a new token.")
            sys.exit(1)
        # Persist refreshed token back to file
        token_path = Path(cfg.get("token_file", "")).expanduser()
        if token_path.is_file():
            token_data["access_token"] = token
            new_refresh = graph.cache_dict.get("refresh_token")
            if new_refresh:
                token_data["refresh_token"] = new_refresh
            token_path.write_text(json.dumps(token_data, indent=2))
            print("Token refreshed and saved.")

    # Get user's timezone from mailbox settings
    mbs = MailboxSettingsHandler(graph)
    settings = await mbs.get_mailbox_settings_async()
    user_tz_win = settings.get("timeZone", win_tz)
    iana_tz = _to_iana(user_tz_win)
    local_tz = ZoneInfo(iana_tz)
    now_local = datetime.now(local_tz)
    print(f"User timezone: {user_tz_win} ({iana_tz}), local time: {now_local.strftime('%H:%M')}")

    # Get rooms, filter by domain
    ph = PlacesHandler(graph)
    all_rooms = await ph.get_rooms_async()

    if domain:
        rooms = [r for r in all_rooms if r.get("emailAddress", "").lower().endswith(domain.lower())]
    else:
        rooms = list(all_rooms)

    # Exclude patterns
    rooms = [r for r in rooms
             if not any(ex in r.get("displayName", "").lower() for ex in exclude)]

    # Separate current vs future
    current_rooms = [r for r in rooms
                     if not any(f in r.get("displayName", "").lower() for f in future_rooms)]
    future_room_list = [r for r in rooms
                        if any(f in r.get("displayName", "").lower() for f in future_rooms)]

    print(f"Current rooms: {len(current_rooms)}, future: {len(future_room_list)}")

    # Get schedule for today (in user's timezone)
    start_local = now_local.replace(hour=hours_start, minute=0, second=0, microsecond=0)
    end_local = now_local.replace(hour=hours_end, minute=0, second=0, microsecond=0)
    start_str = start_local.strftime("%Y-%m-%dT%H:%M:%S")
    end_str = end_local.strftime("%Y-%m-%dT%H:%M:%S")

    emails = [r["emailAddress"] for r in current_rooms]

    body = {
        "schedules": emails,
        "startTime": {"dateTime": start_str, "timeZone": user_tz_win},
        "endTime": {"dateTime": end_str, "timeZone": user_tz_win},
        "availabilityViewInterval": 30,
    }

    result = await graph.run_async(url=ep + "me/calendar/getSchedule", method="POST", json=body, token=token)
    schedules = result.json().get("value", [])

    # Build time slots
    slots = []
    for h in range(hours_start, hours_end):
        slots.append(f"{h:02d}:00")
        slots.append(f"{h:02d}:30")

    room_map = {r["emailAddress"].lower(): r for r in current_rooms}
    current_slot = f"{now_local.hour:02d}:{(now_local.minute // 30) * 30:02d}"

    # Parse schedules into slot maps
    room_slots: dict[str, dict[str, tuple[str, str]]] = {}
    for sched in schedules:
        email = sched.get("scheduleId", "").lower()
        items = sched.get("scheduleItems", [])
        slot_map: dict[str, tuple[str, str]] = {}
        for item in items:
            s_start = item.get("start", {}).get("dateTime", "")
            s_end = item.get("end", {}).get("dateTime", "")
            status = item.get("status", "busy")
            subject = item.get("subject", "")
            if s_start and s_end:
                try:
                    # API returns UTC — convert to user's local timezone
                    st = datetime.fromisoformat(s_start.rstrip("Z")).replace(tzinfo=timezone.utc).astimezone(local_tz)
                    en = datetime.fromisoformat(s_end.rstrip("Z")).replace(tzinfo=timezone.utc).astimezone(local_tz)
                    cur = st
                    while cur < en:
                        key = f"{cur.hour:02d}:{(cur.minute // 30) * 30:02d}"
                        slot_map[key] = (status, subject)
                        cur += timedelta(minutes=30)
                except Exception:
                    pass
        room_slots[email] = slot_map

    # Generate HTML
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{title}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: -apple-system, system-ui, sans-serif; background:#1a1a2e; color:#e0e0e0; padding:24px; }}
h1 {{ margin-bottom:4px; font-size:1.5em; color:#7fdbca; }}
.sub {{ color:#888; margin-bottom:20px; font-size:0.9em; }}
table {{ border-collapse:collapse; width:100%; margin-bottom:24px; background:#16213e; border-radius:8px; overflow:hidden; }}
th, td {{ padding:4px 2px; text-align:center; font-size:10px; border:1px solid #2a2a4a; white-space:nowrap; }}
th {{ background:#0f3460; color:#7fdbca; font-weight:600; position:sticky; top:0; z-index:1; }}
th:first-child {{ text-align:left; min-width:180px; }}
td:first-child {{ text-align:left; font-weight:500; padding:4px 8px; background:#0f3460; font-size:11px; }}
.busy {{ background:#c62828; color:#fff; font-size:9px; overflow:hidden; max-width:60px; text-overflow:ellipsis; }}
.free {{ background:#1b4332; }}
.tentative {{ background:#e65100; color:#fff; font-size:9px; }}
.oof {{ background:#1565c0; color:#fff; font-size:9px; }}
.now {{ border-left:2px solid #7fdbca; }}
.cap {{ color:#888; font-size:10px; }}
.future {{ opacity:0.4; }}
.section {{ color:#7fdbca; font-size:13px; font-weight:600; margin:20px 0 8px; }}
.legend {{ margin-top:16px; font-size:12px; color:#888; display:flex; gap:16px; }}
.legend span {{ display:inline-flex; align-items:center; gap:4px; }}
.legend i {{ display:inline-block; width:12px; height:12px; border-radius:2px; }}
</style>
</head><body>
<h1>{title}</h1>
<p class="sub">{now_local.strftime("%A, %B %d, %Y")} — {len(current_rooms)} rooms, {now_local.strftime("%H:%M")} {iana_tz}</p>
<table>
<tr><th>Room</th>"""

    for s in slots:
        cls = ' class="now"' if s == current_slot else ""
        html += f"<th{cls}>{s}</th>"
    html += "</tr>\n"

    for sched in schedules:
        email = sched.get("scheduleId", "").lower()
        room = room_map.get(email)
        if not room:
            continue
        name = room["displayName"]
        cap = room.get("capacity")
        label = name
        if cap:
            label += f' <span class="cap">({cap})</span>'

        sm = room_slots.get(email, {})
        html += f"<tr><td>{label}</td>"
        for s in slots:
            cls_now = " now" if s == current_slot else ""
            if s in sm:
                status, subj = sm[s]
                cls = {"busy": "busy", "tentative": "tentative", "oof": "oof"}.get(status, "busy")
                short = (subj[:12] + "\u2026") if len(subj) > 12 else subj
                esc_subj = subj.replace('"', '&quot;').replace('<', '&lt;')
                html += f'<td class="{cls}{cls_now}" title="{esc_subj}">{short}</td>'
            else:
                html += f'<td class="free{cls_now}"></td>'
        html += "</tr>\n"

    html += "</table>\n"

    # Future rooms section
    if future_room_list:
        html += f'<p class="section">Planned / Future Rooms ({len(future_room_list)})</p>\n'
        html += '<table class="future"><tr><th>Room</th><th>Capacity</th><th>Status</th></tr>\n'
        for r in future_room_list:
            cap = r.get("capacity") or "—"
            html += f'<tr><td>{r["displayName"]}</td><td>{cap}</td><td>Not yet available</td></tr>\n'
        html += "</table>\n"

    html += """<div class="legend">
  <span><i style="background:#c62828"></i> Busy</span>
  <span><i style="background:#e65100"></i> Tentative</span>
  <span><i style="background:#1565c0"></i> Out of Office</span>
  <span><i style="background:#1b4332"></i> Free</span>
</div>
</body></html>"""

    output.write_text(html)
    print(f"Saved to {output}")

    booked = sum(1 for s in schedules if s.get("scheduleItems"))
    print(f"{booked}/{len(current_rooms)} rooms have bookings today")


if __name__ == "__main__":
    asyncio.run(main())
