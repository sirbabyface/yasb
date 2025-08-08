import os
import re
import urllib.error
import urllib.request
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, List, Optional, Tuple
from zoneinfo import ZoneInfo

from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices, QGuiApplication
from PyQt6.QtWidgets import QHBoxLayout, QLabel, QWidget, QVBoxLayout

from core.utils.tooltip import set_tooltip
from core.utils.utilities import PopupWidget, add_shadow, build_widget_label, get_app_identifier
from core.utils.widgets.animation_manager import AnimationManager
from core.validation.widgets.yasb.google_meet import VALIDATION_SCHEMA
from core.widgets.base import BaseWidget


logger = logging.getLogger(__name__)


class GoogleMeetWidget(BaseWidget):
    """
    Boilerplate widget template for YASB.

    How to use:
    - Duplicate/rename this file and class for your widget.
    - Extend _update_label and/or add new callbacks to implement logic.
    - Add new options to the validation schema at
      src/core/validation/widgets/yasb/google_meet.py and import here.
    - Reference in your config like:

        google_meet:
          type: "yasb.google_meet.GoogleMeetWidget"
          options:
            label: "\u003cspan\u003e\ue71a\u003c/span\u003e {data}"
            label_alt: "{data} items"
            tooltip: true
            update_interval: 1000
            animation:
              enabled: true
              type: "fadeInOut"
              duration: 200
            container_padding: { top: 0, left: 0, bottom: 0, right: 0 }
            callbacks: { on_left: "toggle_label", on_middle: "do_nothing", on_right: "do_nothing" }
    """

    validation_schema = VALIDATION_SCHEMA

    def __init__(
        self,
        label: str,
        label_alt: str,
        tooltip: bool,
        update_interval: int,
        animation: dict[str, Any],
        container_padding: dict[str, int],
        callbacks: dict[str, str],
        env_var: str = "YASB_ICAL_URL",
        meeting_filter: str = "all",
        next_time_format: str = "%Y-%m-%d %H:%M",
        horizon_hours: int | None = None,
        timezone_name: str | None = None,
        menu: dict | None = None,
        icons: dict | None = None,
        notifications: dict | None = None,
        label_shadow: dict | None = None,
        container_shadow: dict | None = None,
    ):
        # Pass update_interval (ms) directly to BaseWidget timer
        super().__init__(update_interval, class_name="google-meet-widget")

        # Store options
        self._label_content = label
        self._label_alt_content = label_alt
        self._tooltip = tooltip
        self._animation = animation or {"enabled": False, "type": "fadeInOut", "duration": 200}
        self._padding = container_padding or {"top": 0, "left": 0, "bottom": 0, "right": 0}
        self._label_shadow = label_shadow or {"enabled": False, "color": "black", "offset": [1, 1], "radius": 3}
        self._container_shadow = container_shadow or {"enabled": False, "color": "black", "offset": [1, 1], "radius": 3}
        self._ical_env_var = env_var
        self._meeting_filter = meeting_filter  # all | online | in_person
        self._next_time_format = next_time_format
        self._horizon_hours = horizon_hours  # None -> end of local day; int -> next N hours
        self._tzname = timezone_name
        self._tz = None
        self._menu_conf = menu or {"blur": True, "round_corners": True, "round_corners_type": "normal", "border_color": "System", "alignment": "right", "direction": "down", "offset_top": 6, "offset_left": 0}
        self._icons = icons or {"online": "\uea64", "time": "\uf017", "copy": "\uebcc"}  # defaults
        self._notifications = notifications or {
            "enabled": False,
            "offsets_minutes": [10, 0],
            "title": "Upcoming: {summary}",
            "message": "Starts at {start_time} ({countdown})",
        }
        self._notified: set[tuple[str, int]] = set()
        try:
            if self._tzname:
                self._tz = ZoneInfo(self._tzname)
        except Exception:
            self._tz = None

        # Internal state example (replace/extend as needed)
        self._show_alt_label = False
        self._data = None  # Set and update this from your logic to feed into {data}
        self._next_event: Optional[dict] = None

        # Construct container
        self._widget_container_layout: QHBoxLayout = QHBoxLayout()
        self._widget_container_layout.setSpacing(0)
        self._widget_container_layout.setContentsMargins(
            self._padding["left"], self._padding["top"], self._padding["right"], self._padding["bottom"]
        )

        self._widget_container: QWidget = QWidget()
        self._widget_container.setLayout(self._widget_container_layout)
        self._widget_container.setProperty("class", "widget-container")
        add_shadow(self._widget_container, self._container_shadow)
        self.widget_layout.addWidget(self._widget_container)

        # Build initial label(s)
        build_widget_label(self, self._label_content, self._label_alt_content, self._label_shadow)

        # Optional tooltip on container
        if self._tooltip:
            set_tooltip(self._widget_container, "Google Meet")

        # Register callbacks and mouse actions
        self.register_callback("toggle_label", self._toggle_label)
        self.register_callback("update_label", self._update_label)
        self.register_callback("refresh_calendar", self._refresh_calendar)
        self.register_callback("toggle_menu", self._toggle_menu)

        # Map primary callbacks
        self.callback_left = callbacks.get("on_left", "toggle_menu")
        self.callback_middle = callbacks.get("on_middle", "do_nothing")
        self.callback_right = callbacks.get("on_right", "do_nothing")

        # Timer callback
        # Fetch calendar on timer tick (defaults to 10 minutes via schema)
        self.callback_timer = "refresh_calendar"
        self.start_timer()
        # Initial immediate refresh to populate label quickly
        self._refresh_calendar()

    def _animate(self):
        if self._animation and self._animation.get("enabled"):
            AnimationManager.animate(self, self._animation.get("type"), self._animation.get("duration"))

    def _toggle_label(self):
        self._animate()
        self._show_alt_label = not self._show_alt_label
        # Toggle visibility of built labels
        for widget in getattr(self, "_widgets", []):
            widget.setVisible(not self._show_alt_label)
        for widget in getattr(self, "_widgets_alt", []):
            widget.setVisible(self._show_alt_label)
        self._update_label()

    def _format_countdown(self, dt: Optional[datetime]) -> str:
        if not isinstance(dt, datetime):
            return ""
        now = datetime.now(timezone.utc)
        # Ensure both in same tz
        start = dt.astimezone(timezone.utc)
        delta = start - now
        if delta.total_seconds() <= 0:
            return "now"
        minutes = int(delta.total_seconds() // 60)
        hours = minutes // 60
        minutes = minutes % 60
        if hours > 0:
            return f"{hours}h {minutes:02d}m"
        return f"{minutes}m"

    def _update_label(self):
        """
        Update both icon and text portions based on current state.
        Placeholders supported in labels:
        - {data} or {count}: number of upcoming events in the next 12 hours after filtering
        - {next_time}: formatted next event start time (next_time_format)
        - {countdown}: time until next event (e.g., 1h 05m)
        - {next_summary}: next event summary/title
        """
        active_widgets = getattr(self, "_widgets_alt", []) if self._show_alt_label else getattr(self, "_widgets", [])
        active_label_content = self._label_alt_content if self._show_alt_label else self._label_content

        # Split into span and non-span parts to update in place
        label_parts = [part.strip() for part in re.split(r"(\u003cspan.*?\u003e.*?\u003c/span\u003e)", active_label_content) if part]
        widget_index = 0
        for part in label_parts:
            if widget_index >= len(active_widgets) or not isinstance(active_widgets[widget_index], QLabel):
                continue
            formatted_text = part
            # Placeholder mapping
            count_str = str(self._data) if self._data is not None else "0"
            next_dt = self._next_event.get("start") if self._next_event else None
            # Format next time in configured timezone (if provided) or local
            if isinstance(next_dt, datetime):
                if self._tz:
                    next_time_str = next_dt.astimezone(self._tz).strftime(self._next_time_format)
                else:
                    next_time_str = next_dt.astimezone().strftime(self._next_time_format)
            else:
                next_time_str = ""
            countdown_str = self._format_countdown(next_dt)
            next_summary = self._next_event.get("summary", "") if self._next_event else ""

            replacements = {
                "{data}": count_str,
                "{count}": count_str,
                "{next_time}": next_time_str,
                "{countdown}": countdown_str,
                "{next_summary}": next_summary,
            }
            for k, v in replacements.items():
                formatted_text = formatted_text.replace(k, v)

            if "\u003cspan" in part and "\u003c/span\u003e" in part:
                # Strip span tags; class is set at creation time via build_widget_label
                icon_text = re.sub(r"\u003cspan.*?\u003e|\u003c/span\u003e", "", formatted_text).strip()
                active_widgets[widget_index].setText(icon_text)
            else:
                active_widgets[widget_index].setText(formatted_text)
            active_widgets[widget_index].style().unpolish(active_widgets[widget_index])
            widget_index += 1

    # --- iCal fetching and parsing ---
    def _get_ical_url(self) -> Optional[str]:
        url = os.getenv(self._ical_env_var)
        return url

    def _fetch_ical(self, url: str) -> Optional[str]:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "YASB/GoogleMeetWidget"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                charset = resp.headers.get_content_charset() or "utf-8"
                return resp.read().decode(charset, errors="replace")
        except (urllib.error.URLError, urllib.error.HTTPError, ValueError):
            return None

    @staticmethod
    def _unfold_ical_lines(text: str) -> List[str]:
        # RFC 5545 line unfolding: concatenate lines beginning with space
        lines = text.splitlines()
        unfolded = []
        for line in lines:
            if line.startswith(" ") and unfolded:
                unfolded[-1] += line[1:]
            else:
                unfolded.append(line)
        return unfolded

    @staticmethod
    def _parse_dt(value: str) -> Optional[datetime]:
        # Handles formats like 20250808T112233Z or local time without Z
        try:
            if value.endswith("Z"):
                return datetime.strptime(value, "%Y%m%dT%H%M%SZ").replace(tzinfo=timezone.utc)
            # Some entries may include timezone params like DTSTART;TZID=...:YYYYMMDDT...
            # Here we treat as local time if no trailing Z
            return datetime.strptime(value, "%Y%m%dT%H%M%S").astimezone()
        except Exception:
            try:
                # All-day date form YYYYMMDD
                return datetime.strptime(value, "%Y%m%d").astimezone()
            except Exception:
                return None

    def _is_online(self, event: dict) -> bool:
        # Heuristic: if any URL present in known fields or meet/zoom/teams in text
        text = " ".join(
            str(event.get(k, "")) for k in ("location", "description", "url", "summary")
        ).lower()
        return ("http://" in text) or ("https://" in text) or any(
            d in text for d in ["meet.google.com", "zoom.us", "teams.microsoft.com", "webex.com"]
        )

    def _extract_first_url(self, event: dict) -> Optional[str]:
        # Try URL field first, then search in description/location/summary
        url = event.get("url")
        if isinstance(url, str) and url.startswith("http"):
            return url
        import re as _re
        text = " \n".join(str(event.get(k, "")) for k in ("url", "description", "location", "summary"))
        m = _re.search(r"https?://\S+", text)
        return m.group(0) if m else None
    def _parse_events(self, ical_text: str, window_start: datetime | None = None, window_end: datetime | None = None):
        if not ical_text:
            return []
        lines = self._unfold_ical_lines(ical_text)
        raw_events: list[dict] = []
        in_event = False
        current: dict[str, Any] = {}
        for line in lines:
            if line.startswith("BEGIN:VEVENT"):
                in_event = True
                current = {}
                continue
            if line.startswith("END:VEVENT"):
                if current:
                    raw_events.append(current)
                in_event = False
                current = {}
                continue
            if not in_event:
                continue
            # Extract fields
            if line.startswith("DTSTART"):
                # Split at ':' to get value (supports params before colon)
                val = line.split(":", 1)[-1].strip()
                current["start"] = self._parse_dt(val)
            elif line.startswith("DTEND"):
                val = line.split(":", 1)[-1].strip()
                current["end"] = self._parse_dt(val)
            elif line.startswith("UID:"):
                current["uid"] = line[len("UID:") :].strip()
            elif line.startswith("SUMMARY:"):
                current["summary"] = line[len("SUMMARY:") :].strip()
            elif line.startswith("DESCRIPTION:"):
                current["description"] = line[len("DESCRIPTION:") :].strip()
            elif line.startswith("LOCATION:"):
                current["location"] = line[len("LOCATION:") :].strip()
            elif line.startswith("URL:"):
                current["url"] = line[len("URL:") :].strip()
            elif line.startswith("RRULE:"):
                current["rrule"] = line[len("RRULE:") :].strip()
            elif line.startswith("RDATE"):
                # RDATE may be comma-separated values
                val = line.split(":", 1)[-1].strip()
                dates = [self._parse_dt(v.strip()) for v in val.split(",") if v.strip()]
                current.setdefault("rdate", []).extend([d for d in dates if isinstance(d, datetime)])
            elif line.startswith("EXDATE"):
                val = line.split(":", 1)[-1].strip()
                dates = [self._parse_dt(v.strip()) for v in val.split(",") if v.strip()]
                current.setdefault("exdate", []).extend([d for d in dates if isinstance(d, datetime)])
            elif line.startswith("RECURRENCE-ID"):
                val = line.split(":", 1)[-1].strip()
                current["recurrence_id"] = self._parse_dt(val)
        # Filter invalid
        raw_events = [e for e in raw_events if isinstance(e.get("start"), datetime)]

        # Group into masters and overrides (RECURRENCE-ID)
        def _uid_of(ev: dict) -> str:
            u = ev.get("uid")
            if u:
                return u
            s = ev.get("start")
            return f"NOUID|{s.isoformat() if isinstance(s, datetime) else ''}"

        masters = [e for e in raw_events if not e.get("recurrence_id")]
        overrides = [e for e in raw_events if e.get("recurrence_id")]
        overrides_by_uid: dict[str, list[dict]] = {}
        for ov in overrides:
            overrides_by_uid.setdefault(_uid_of(ov), []).append(ov)

        # Expand recurrences within window
        def _parse_rrule(rr: str) -> dict[str, Any]:
            parts = {}
            for token in rr.split(";"):
                if not token:
                    continue
                if "=" in token:
                    k, v = token.split("=", 1)
                    parts[k.upper()] = v
            return parts

        def _weekday_to_py(day: str) -> int | None:
            # MO TU WE TH FR SA SU -> 0..6 (Python Monday=0)
            mapping = {"MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6}
            return mapping.get(day.upper())

        def _expand_event(ev: dict, ws: datetime | None, we: datetime | None) -> list[dict]:
            start: datetime = ev.get("start")
            end: datetime | None = ev.get("end")
            duration = None
            if isinstance(start, datetime) and isinstance(end, datetime):
                try:
                    duration = end - start
                except Exception:
                    duration = None
            rrule_text = ev.get("rrule")
            rdates: list[datetime] = ev.get("rdate", [])
            exdates: list[datetime] = ev.get("exdate", [])

            # Adjust RDATE/EXDATE that are date-only to carry base time-of-day
            def _adjust_to_base_time(dt: datetime) -> datetime:
                if not isinstance(dt, datetime):
                    return dt
                # Heuristic: if midnight and base has time
                if dt.hour == 0 and dt.minute == 0 and dt.second == 0 and isinstance(start, datetime):
                    if start.hour or start.minute or start.second or start.microsecond:
                        return dt.replace(
                            hour=start.hour,
                            minute=start.minute,
                            second=start.second,
                            microsecond=start.microsecond,
                        )
                return dt

            rdates = [_adjust_to_base_time(d) for d in rdates]
            exdates = [_adjust_to_base_time(d) for d in exdates]

            # If no RRULE/RDATE, return as-is
            if not rrule_text and not rdates:
                return [ev]

            # Normalize window
            ws_utc = ws.astimezone(timezone.utc) if isinstance(ws, datetime) else None
            we_utc = we.astimezone(timezone.utc) if isinstance(we, datetime) else None

            instances: list[tuple[datetime, datetime | None]] = []

            # Try robust expansion using dateutil.rrule
            used_dateutil = False
            if rrule_text:
                try:
                    from dateutil.rrule import rrulestr, rruleset
                    rs = rruleset()
                    # Build rule with dtstart
                    rr = rrulestr(rrule_text, dtstart=start)
                    rs.rrule(rr)
                    # Add RDATEs and EXDATEs
                    for rd in rdates:
                        if isinstance(rd, datetime):
                            rs.rdate(rd)
                    for ex in exdates:
                        if isinstance(ex, datetime):
                            rs.exdate(ex)
                    # Get occurrences in window
                    if ws_utc and we_utc:
                        occ = rs.between(ws_utc, we_utc, inc=True)
                    elif ws_utc and not we_utc:
                        occ = list(rs.xafter(ws_utc, count=365))
                    else:
                        occ = []
                    for s in occ:
                        e = (s + duration) if duration else None
                        instances.append((s, e))
                    used_dateutil = True
                except Exception:
                    used_dateutil = False

            if not used_dateutil:
                # Fallback: simple expansion for basic DAILY/WEEKLY/MONTHLY/YEARLY with INTERVAL/COUNT/UNTIL and BYMONTHDAY
                params = _parse_rrule(rrule_text) if rrule_text else {}
                freq = (params.get("FREQ", "") or "").upper()
                interval = int(params.get("INTERVAL", "1") or 1)
                until = params.get("UNTIL")
                until_dt = self._parse_dt(until) if until else None
                count = int(params.get("COUNT", "0") or 0)
                bymonth = params.get("BYMONTH")
                bymonthdays = params.get("BYMONTHDAY")
                bydays = params.get("BYDAY")
                bymonth_vals = [int(m) for m in bymonth.split(",")] if bymonth else None
                bymonthday_vals = [int(d) for d in bymonthdays.split(",")] if bymonthdays else None

                produced = 0
                cur = start

                def cont(cdt: datetime) -> bool:
                    if until_dt and cdt > until_dt:
                        return False
                    if count and produced >= count:
                        return False
                    if we_utc and cdt.astimezone(timezone.utc) > we_utc:
                        return False
                    return True

                if freq == "DAILY":
                    step = timedelta(days=interval)
                    while cont(cur):
                        if not ws_utc or cur.astimezone(timezone.utc) >= ws_utc:
                            instances.append((cur, (cur + duration) if duration else None))
                            produced += 1
                        cur = cur + step
                elif freq == "WEEKLY":
                    # If BYDAY provided, iterate day-by-day; else jump by weeks
                    if bydays:
                        target_dows = [_weekday_to_py(d) for d in [x.strip().upper() for x in bydays.split(",")]]
                        cur_day = cur
                        while cont(cur_day):
                            dow = cur_day.weekday()
                            if dow in target_dows and (not ws_utc or cur_day.astimezone(timezone.utc) >= ws_utc):
                                instances.append((cur_day, (cur_day + duration) if duration else None))
                                produced += 1
                            cur_day = cur_day + timedelta(days=1)
                    else:
                        step = timedelta(weeks=interval)
                        while cont(cur):
                            if not ws_utc or cur.astimezone(timezone.utc) >= ws_utc:
                                instances.append((cur, (cur + duration) if duration else None))
                                produced += 1
                            cur = cur + step
                elif freq == "MONTHLY":
                    from calendar import monthrange
                    def add_months(dt: datetime, n: int) -> datetime:
                        y = dt.year + (dt.month - 1 + n) // 12
                        m = (dt.month - 1 + n) % 12 + 1
                        d = min(dt.day, monthrange(y, m)[1])
                        return dt.replace(year=y, month=m, day=d)
                    cur = start
                    while cont(cur):
                        if bymonthday_vals:
                            for md in bymonthday_vals:
                                try:
                                    y = cur.year; m = cur.month
                                    from calendar import monthrange as _mr
                                    last = _mr(y, m)[1]
                                    day = md if md > 0 else last + md + 1
                                    occ = cur.replace(day=day)
                                except Exception:
                                    continue
                                if not ws_utc or occ.astimezone(timezone.utc) >= ws_utc:
                                    instances.append((occ, (occ + duration) if duration else None))
                                    produced += 1
                        else:
                            if not ws_utc or cur.astimezone(timezone.utc) >= ws_utc:
                                instances.append((cur, (cur + duration) if duration else None))
                                produced += 1
                        cur = add_months(cur, interval)
                elif freq == "YEARLY":
                    from calendar import monthrange
                    def add_years(dt: datetime, n: int) -> datetime:
                        y = dt.year + n
                        m = dt.month
                        d = min(dt.day, monthrange(y, m)[1])
                        return dt.replace(year=y, month=m, day=d)
                    cur = start
                    while cont(cur):
                        months = bymonth_vals if bymonth_vals else [cur.month]
                        days = bymonthday_vals if bymonthday_vals else [cur.day]
                        for m in months:
                            for d in days:
                                try:
                                    last = monthrange(cur.year, m)[1]
                                    day = d if d > 0 else last + d + 1
                                    occ = cur.replace(month=m, day=day)
                                except Exception:
                                    continue
                                if not ws_utc or occ.astimezone(timezone.utc) >= ws_utc:
                                    instances.append((occ, (occ + duration) if duration else None))
                                    produced += 1
                        cur = add_years(cur, interval)
                else:
                    instances.append((start, (start + duration) if duration else None))

            # Add RDATE explicit instances (already time-adjusted)
            for rd in rdates:
                instances.append((rd, (rd + duration) if duration else None))

            # De-duplicate and filter to window
            def _k(dt: datetime) -> datetime:
                return dt.astimezone(timezone.utc).replace(tzinfo=timezone.utc)

            out: list[dict] = []
            seen: set[datetime] = set()
            for s, e in instances:
                if not isinstance(s, datetime):
                    continue
                ks = _k(s)
                if ws_utc and ks < ws_utc:
                    continue
                if we_utc and ks > we_utc:
                    continue
                if ks in seen:
                    continue
                seen.add(ks)
                ev_copy = dict(ev)
                ev_copy["start"] = s
                if isinstance(e, datetime):
                    ev_copy["end"] = e
                out.append(ev_copy)
            return out

        # If no window given, return raw (non-expanded) to avoid unbounded expansion
        if window_start is None or window_end is None:
            logger.debug("GoogleMeet: parse_events called without window; returning non-expanded %d events", len(raw_events))
            return raw_events

        # Expand masters, then apply overrides (use override DTSTART for start time)
        expanded: list[dict] = []
        override_excl_keys: set[datetime] = set()
        def _key(dt: datetime) -> datetime:
            return dt.astimezone(timezone.utc).replace(tzinfo=timezone.utc)

        for ev in masters:
            expanded.extend(_expand_event(ev, window_start, window_end))
            uid = _uid_of(ev)
            for ov in overrides_by_uid.get(uid, []):
                recid = ov.get("recurrence_id")
                if isinstance(recid, datetime):
                    override_excl_keys.add(_key(recid))

        # Remove master instances at overridden times
        expanded = [e for e in expanded if not isinstance(e.get("start"), datetime) or _key(e["start"]) not in override_excl_keys]

        # Add overrides as explicit instances using their DTSTART/DTEND
        for uid, ov_list in overrides_by_uid.items():
            for ov in ov_list:
                s = ov.get("start")
                if not isinstance(s, datetime):
                    continue
                if s < window_start or s > window_end:
                    continue
                e = ov.get("end")
                inst = dict(ov)
                inst.pop("recurrence_id", None)
                expanded.append(inst)

        return expanded

    def _parse_events_ical(self, ical_text: str, window_start: datetime | None = None, window_end: datetime | None = None):
        try:
            import icalendar as _ical
        except Exception as e:
            logger.error("GoogleMeet: icalendar library not available: %s", e)
            return self._parse_events(ical_text, window_start, window_end)
        if not ical_text:
            return []
        try:
            cal = _ical.Calendar.from_ical(ical_text)
        except Exception as e:
            logger.error("GoogleMeet: failed to parse ICS via icalendar: %s", e)
            return []

        # Helpers
        def _to_datetime(val, base_time: datetime | None):
            if isinstance(val, datetime):
                dt = val
            elif hasattr(val, 'dt'):
                v = val.dt
                if isinstance(v, datetime):
                    dt = v
                else:
                    # date only
                    dt = datetime(v.year, v.month, v.day)
            elif hasattr(val, 'to_ical'):
                try:
                    v = val.dt
                    if isinstance(v, datetime):
                        dt = v
                    else:
                        dt = datetime(v.year, v.month, v.day)
                except Exception:
                    return None
            else:
                return None
            # Ensure tz-aware
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
            # If date-only and base has time, copy time-of-day
            if base_time and (dt.hour == 0 and dt.minute == 0 and dt.second == 0) and (
                base_time.hour or base_time.minute or base_time.second or base_time.microsecond
            ):
                dt = dt.replace(hour=base_time.hour, minute=base_time.minute, second=base_time.second, microsecond=base_time.microsecond)
            return dt

        def _component_to_dict(comp):
            def _get_text(name: str, default: str = ""):
                v = comp.get(name)
                try:
                    if v is None:
                        return default
                    vv = v.to_ical().decode() if hasattr(v, 'to_ical') else str(v)
                    return vv
                except Exception:
                    return default
            uid = _get_text('UID', '')
            summary = _get_text('SUMMARY', '')
            description = _get_text('DESCRIPTION', '')
            location = _get_text('LOCATION', '')
            url = _get_text('URL', '')
            # dtstart/dtend
            dtstart_raw = comp.get('DTSTART')
            base_start = _to_datetime(dtstart_raw.dt if hasattr(dtstart_raw,'dt') else dtstart_raw, None)
            dtend_raw = comp.get('DTEND')
            base_end = _to_datetime(dtend_raw.dt if hasattr(dtend_raw,'dt') else dtend_raw, base_start)
            # RRULE
            rrule_val = comp.get('RRULE')
            rrule_str = None
            if rrule_val:
                try:
                    rrule_str = _ical.vRecur(rrule_val).to_ical().decode()
                except Exception:
                    # Some icalendar versions store RRULE as dict-like
                    try:
                        # Build simple key=value; comma joined for lists
                        parts = []
                        for k, v in (rrule_val.items() if hasattr(rrule_val, 'items') else []):
                            if isinstance(v, (list, tuple)):
                                v = ",".join([str(x) for x in v])
                            parts.append(f"{str(k).upper()}={v}")
                        rrule_str = ";".join(parts) if parts else None
                    except Exception:
                        rrule_str = None
            # RDATE
            rdate_vals = comp.get('RDATE')
            rdates: list[datetime] = []
            if rdate_vals is not None:
                try:
                    # Could be vDDDList
                    seq = rdate_vals.dts if hasattr(rdate_vals, 'dts') else rdate_vals
                    for d in seq:
                        dv = d.dt if hasattr(d, 'dt') else d
                        dt = _to_datetime(dv, base_start)
                        if isinstance(dt, datetime):
                            rdates.append(dt)
                except Exception:
                    pass
            # EXDATE
            exdate_vals = comp.get('EXDATE')
            exdates: list[datetime] = []
            if exdate_vals is not None:
                try:
                    seq = exdate_vals.dts if hasattr(exdate_vals, 'dts') else exdate_vals
                    for d in seq:
                        dv = d.dt if hasattr(d, 'dt') else d
                        dt = _to_datetime(dv, base_start)
                        if isinstance(dt, datetime):
                            exdates.append(dt)
                except Exception:
                    pass
            # RECURRENCE-ID (override instance)
            recurid = comp.get('RECURRENCE-ID')
            recurid_dt = _to_datetime(recurid.dt if hasattr(recurid,'dt') else recurid, base_start) if recurid else None

            return {
                'uid': uid,
                'summary': summary,
                'description': description,
                'location': location,
                'url': url,
                'start': base_start,
                'end': base_end,
                'rrule': rrule_str,
                'rdates': rdates,
                'exdates': exdates,
                'recurid': recurid_dt,
            }

        # Build masters and overrides
        masters: dict[str, dict] = {}
        overrides: dict[str, list[dict]] = {}
        for comp in cal.walk('VEVENT'):
            ev = _component_to_dict(comp)
            if not isinstance(ev.get('start'), datetime):
                continue
            uid = ev.get('uid') or f"NOUID|{ev['start'].isoformat()}"
            if ev.get('recurid'):
                overrides.setdefault(uid, []).append(ev)
            else:
                masters[uid] = ev

        # Expand masters within window
        results: list[dict] = []
        if window_start is None or window_end is None:
            window_start = datetime.min.replace(tzinfo=timezone.utc)
            window_end = datetime.max.replace(tzinfo=timezone.utc)
        ws_utc = window_start.astimezone(timezone.utc)
        we_utc = window_end.astimezone(timezone.utc)

        for uid, ev in masters.items():
            start = ev['start']
            end = ev.get('end')
            duration = (end - start) if isinstance(end, datetime) else None
            rrule_text = ev.get('rrule')
            rdates = ev.get('rdates') or []
            exdates = ev.get('exdates') or []
            instances: list[datetime] = []
            if rrule_text:
                try:
                    from dateutil.rrule import rrulestr, rruleset
                    rs = rruleset()
                    rr = rrulestr(rrule_text, dtstart=start)
                    rs.rrule(rr)
                    for rd in rdates:
                        rs.rdate(rd)
                    for ex in exdates:
                        rs.exdate(ex)
                    occ = rs.between(ws_utc, we_utc, inc=True)
                    instances.extend(occ)
                except Exception as e:
                    logger.debug("GoogleMeet: rrule expansion failed for uid=%s: %s", uid, e)
                    # fallback: include the base if in window
                    if ws_utc <= start.astimezone(timezone.utc) <= we_utc:
                        instances.append(start)
            else:
                # No rrule: consider base start and any RDATEs
                if ws_utc <= start.astimezone(timezone.utc) <= we_utc:
                    instances.append(start)
                for rd in rdates:
                    if ws_utc <= rd.astimezone(timezone.utc) <= we_utc:
                        instances.append(rd)
            # Apply overrides: remove master at recurid, add override DTSTART
            excl_keys = set()
            for ov in overrides.get(uid, []):
                rid = ov.get('recurid')
                if isinstance(rid, datetime):
                    excl_keys.add(rid.astimezone(timezone.utc))
            final_starts: list[datetime] = []
            seen = set()
            for s in instances:
                su = s.astimezone(timezone.utc)
                if su in excl_keys or su in seen:
                    continue
                seen.add(su)
                final_starts.append(s)
            # Add instances
            for s in final_starts:
                e = (s + duration) if duration else None
                results.append({
                    'uid': uid,
                    'summary': ev['summary'],
                    'description': ev['description'],
                    'location': ev['location'],
                    'url': ev['url'],
                    'start': s,
                    'end': e,
                })
            # Add overrides as explicit events if within window
            for ov in overrides.get(uid, []):
                s = ov.get('start')
                if not isinstance(s, datetime):
                    continue
                su = s.astimezone(timezone.utc)
                if ws_utc <= su <= we_utc:
                    results.append({
                        'uid': uid,
                        'summary': ov['summary'] or ev['summary'],
                        'description': ov['description'] or ev['description'],
                        'location': ov['location'] or ev['location'],
                        'url': ov['url'] or ev['url'],
                        'start': s,
                        'end': ov.get('end') or ((s + duration) if duration else None),
                    })
        return results

    def _refresh_calendar(self):
        logger.debug(
            "GoogleMeet: refreshing calendar (env_var=%s, filter=%s, tz=%s, horizon_hours=%s)",
            self._ical_env_var,
            self._meeting_filter,
            self._tzname if getattr(self, "_tzname", None) else (getattr(self, "_tz", None).key if getattr(self, "_tz", None) else None),
            self._horizon_hours,
        )
        url = self._get_ical_url()
        if not url:
            logger.debug("GoogleMeet: no iCal URL found in env var %s", self._ical_env_var)
            # No URL configured; set empty state
            self._data = 0
            self._next_event = None
            self._update_label()
            return
        # Build horizon window
        base_tz = self._tz if self._tz else datetime.now().astimezone().tzinfo
        now_local = datetime.now(base_tz).astimezone(base_tz)
        if self._horizon_hours is None:
            end_of_day_local = now_local.replace(hour=23, minute=59, second=59, microsecond=0)
            horizon_dt_utc = end_of_day_local.astimezone(timezone.utc)
        else:
            horizon_dt_utc = (now_local + timedelta(hours=max(0, int(self._horizon_hours)))).astimezone(timezone.utc)
        now_utc = now_local.astimezone(timezone.utc)
        ical_text = self._fetch_ical(url)
        logger.debug("GoogleMeet: fetched iCal text (%d chars)", len(ical_text) if isinstance(ical_text, str) else 0)
        events = self._parse_events_ical(ical_text or "", window_start=now_utc, window_end=horizon_dt_utc)
        logger.debug("GoogleMeet: parsed+expanded %d events from iCal within window (icalendar)", len(events))
        logger.debug(
            "GoogleMeet: time window UTC [%s -> %s] (now_local=%s)",
            now_utc.isoformat(),
            horizon_dt_utc.isoformat(),
            now_local.isoformat(),
        )
        # Events are already windowed by parser; keep variable naming for clarity
        windowed = events
        # Apply meeting_filter
        if self._meeting_filter == "online":
            filtered = [e for e in windowed if self._is_online(e)]
        elif self._meeting_filter == "in_person":
            filtered = [e for e in windowed if not self._is_online(e)]
        else:
            filtered = windowed
        logger.debug("GoogleMeet: %d events after applying meeting_filter='%s'", len(filtered), self._meeting_filter)
        upcoming = filtered
        upcoming.sort(key=lambda e: e.get("start"))
        self._data = len(upcoming)
        self._next_event = upcoming[0] if upcoming else None
        if self._next_event:
            try:
                _ns = self._next_event.get("summary", "")
                _st = self._next_event.get("start")
                _st_disp = _st.astimezone().isoformat() if isinstance(_st, datetime) else str(_st)
                logger.debug("GoogleMeet: next event '%s' at %s; total upcoming=%d", _ns, _st_disp, self._data)
            except Exception:
                logger.debug("GoogleMeet: next event set; total upcoming=%d", self._data)
        else:
            logger.debug("GoogleMeet: no upcoming events in window")
        self._update_label()

        # Notifications
        self._maybe_send_notifications(upcoming, now_local)

        # Cache upcoming list for menu
        self._upcoming = upcoming

        # Update tooltip with next event details
        if self._tooltip:
            if self._next_event:
                start = self._next_event.get("start")
                summary = self._next_event.get("summary", "Event")
                when = start.astimezone().strftime("%Y-%m-%d %H:%M") if isinstance(start, datetime) else str(start)
                set_tooltip(self._widget_container, f"Next: {summary}\n{when}")
            else:
                set_tooltip(self._widget_container, "No upcoming events")

    # --- Popup menu for upcoming events ---
    def _toggle_menu(self):
        if self._animation and self._animation.get("enabled"):
            AnimationManager.animate(self, self._animation.get("type"), self._animation.get("duration"))
        self._show_menu()

    def _show_menu(self):
        # Build popup using PopupWidget
        dialog = PopupWidget(
            self,
            blur=self._menu_conf.get("blur", True),
            round_corners=self._menu_conf.get("round_corners", True),
            round_corners_type=self._menu_conf.get("round_corners_type", "normal"),
            border_color=self._menu_conf.get("border_color", "System"),
        )
        dialog.setProperty("class", "google-meet-menu")

        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        items = getattr(self, "_upcoming", []) or []
        row_widgets = []
        row_urls = []
        if not items:
            empty = QLabel("No upcoming events")
            empty.setProperty("class", "empty")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(empty)
        else:
            for ev in items[:20]:  # cap list for readability
                row = QWidget()
                is_online = self._is_online(ev)
                row.setProperty("class", "item online" if is_online else "item")
                rlay = QHBoxLayout(row)
                rlay.setContentsMargins(6, 4, 6, 4)
                rlay.setSpacing(6)

                # Online icon
                icon_lbl = QLabel(self._icons.get("online", "")) if is_online else QLabel("")
                icon_lbl.setProperty("class", "icon")
                rlay.addWidget(icon_lbl)

                # Time label (start – end)
                t_start = ev.get("start")
                t_end = ev.get("end")
                show_tz = self._tz if self._tz else None
                if isinstance(t_start, datetime):
                    start_disp = (t_start.astimezone(show_tz) if show_tz else t_start.astimezone()).strftime(self._next_time_format)
                else:
                    start_disp = ""
                if isinstance(t_end, datetime):
                    end_disp = (t_end.astimezone(show_tz) if show_tz else t_end.astimezone()).strftime(self._next_time_format)
                else:
                    end_disp = ""
                time_txt = f"{start_disp} - {end_disp}" if end_disp else start_disp
                time_lbl = QLabel(f"{self._icons.get('time', '')} {time_txt}" if self._icons.get('time') else time_txt)
                time_lbl.setProperty("class", "time")
                rlay.addWidget(time_lbl)

                # Summary
                summary_lbl = QLabel(ev.get("summary", "Event"))
                summary_lbl.setProperty("class", "title")
                rlay.addWidget(summary_lbl, stretch=1)

                # Click behavior on row (open meeting)
                url = self._extract_first_url(ev) if is_online else None
                if url:
                    def _on_click(u=url):
                        def handler(event):
                            dialog.hide()
                            QDesktopServices.openUrl(QUrl(u))
                        return handler
                    row.mousePressEvent = _on_click()
                    row.setCursor(Qt.CursorShape.PointingHandCursor)

                # Copy URL icon button (does not show URL)
                if url:
                    copy_lbl = QLabel(self._icons.get("copy", ""))
                    copy_lbl.setProperty("class", "copy")
                    copy_lbl.setCursor(Qt.CursorShape.PointingHandCursor)
                    def _on_copy(u=url):
                        def handler(event):
                            cb = QGuiApplication.clipboard()
                            cb.setText(u)
                            # Prevent row click from firing
                            event.accept()
                        return handler
                    copy_lbl.mousePressEvent = _on_copy()
                    rlay.addWidget(copy_lbl)

                layout.addWidget(row)
                row_widgets.append(row)
                row_urls.append(url)

        # Keyboard navigation support
        dialog.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        dialog.setFocus()
        dialog._row_widgets = row_widgets
        dialog._row_urls = row_urls
        dialog._selected_index = 0 if row_widgets else -1

        def _set_selected(idx):
            # clear
            for i, w in enumerate(dialog._row_widgets):
                w.setProperty("selected", i == idx)
                try:
                    w.style().unpolish(w); w.style().polish(w)
                except Exception:
                    pass
        _set_selected(dialog._selected_index)

        def _activate(idx):
            if 0 <= idx < len(dialog._row_widgets):
                u = dialog._row_urls[idx]
                if u:
                    dialog.hide()
                    QDesktopServices.openUrl(QUrl(u))

        def _copy(idx):
            if 0 <= idx < len(dialog._row_widgets):
                u = dialog._row_urls[idx]
                if u:
                    QGuiApplication.clipboard().setText(u)

        def _keyPressEvent(event):
            key = event.key()
            if key in (Qt.Key.Key_Escape,):
                dialog.hide(); event.accept(); return
            if not dialog._row_widgets:
                return
            if key in (Qt.Key.Key_Up,):
                dialog._selected_index = max(0, dialog._selected_index - 1)
                _set_selected(dialog._selected_index)
                event.accept(); return
            if key in (Qt.Key.Key_Down,):
                dialog._selected_index = min(len(dialog._row_widgets) - 1, dialog._selected_index + 1)
                _set_selected(dialog._selected_index)
                event.accept(); return
            if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                _activate(dialog._selected_index); event.accept(); return
            if key in (Qt.Key.Key_C,):
                _copy(dialog._selected_index); event.accept(); return
            # fallthrough to default
            try:
                return QWidget.keyPressEvent(dialog, event)
            except Exception:
                pass
        dialog.keyPressEvent = _keyPressEvent

        dialog.setLayout(layout)
        dialog.adjustSize()
        dialog.setPosition(
            alignment=self._menu_conf.get("alignment", "right"),
            direction=self._menu_conf.get("direction", "down"),
            offset_left=self._menu_conf.get("offset_left", 0),
            offset_top=self._menu_conf.get("offset_top", 6),
        )
        dialog.show()
        self.dialog = dialog

    # --- Windows toast notifications ---
    def _format_notification(self, ev: dict, now_local: datetime, is_now: bool = False) -> tuple[str, str]:
        if is_now:
            title_tpl = self._notifications.get("title_now") or self._notifications.get("title", "Upcoming: {summary}")
            msg_tpl = self._notifications.get("message_now") or self._notifications.get("message", "Starts at {start_time} ({countdown})")
        else:
            title_tpl = self._notifications.get("title", "Upcoming: {summary}")
            msg_tpl = self._notifications.get("message", "Starts at {start_time} ({countdown})")
        summary = ev.get("summary", "Event")
        start: datetime | None = ev.get("start")
        show_tz = self._tz if self._tz else None
        if isinstance(start, datetime):
            start_local = start.astimezone(show_tz) if show_tz else start.astimezone()
            start_time = start_local.strftime(self._next_time_format)
        else:
            start_time = ""
        countdown = self._format_countdown(start)
        return (
            title_tpl.replace("{summary}", summary).replace("{start_time}", start_time).replace("{countdown}", countdown),
            msg_tpl.replace("{summary}", summary).replace("{start_time}", start_time).replace("{countdown}", countdown),
        )

    def _show_toast(self, title: str, message: str, url: str | None = None, minutes_remaining: int | None = None):
        try:
            app_id = get_app_identifier()
            from winrt.windows.ui.notifications import ToastNotificationManager, ToastNotification
            from winrt.windows.data.xml.dom import XmlDocument

            # Actions: Join (protocol), Snooze (system), Dismiss (system)
            join_action_xml = ""
            if url:
                join_text = self._notifications.get("join_button_text", "Join")
                join_action_xml = f"<action content=\"{join_text}\" activationType=\"protocol\" arguments=\"{url}\"/>"

            # Build snooze options filtered by time remaining (if provided)
            default_snooze = int(self._notifications.get("snooze_minutes", 5) or 5)
            options = self._notifications.get("snooze_options_minutes") or [default_snooze]
            try:
                options = [int(x) for x in options if int(x) > 0]
            except Exception:
                options = [default_snooze]
            if minutes_remaining is not None and minutes_remaining >= 1:
                options = [m for m in options if m <= minutes_remaining]
            if not options:
                options = [1 if (minutes_remaining is None or minutes_remaining >= 1) else 0]
            options = [m for m in options if m > 0]
            default_input = str(options[0])
            selection_items = "".join([f"\u003cselection id=\\\"{m}\\\" content=\\\"{m} minutes\\\"/\u003e" for m in options])
            snooze_text = self._notifications.get("snooze_button_text", "Snooze")
            actions_xml = f"""
            \u003cactions\u003e
              {join_action_xml}
              \u003cinput id=\"snoozeTime\" type=\"selection\" defaultInput=\"{default_input}\"\u003e 
                {selection_items}
              \u003c/input\u003e
              \u003caction content=\"{snooze_text}\" activationType=\"system\" arguments=\"snooze\" hint-inputId=\"snoozeTime\"/\u003e
              \u003caction content=\"Dismiss\" activationType=\"system\" arguments=\"dismiss\"/\u003e
            \u003c/actions\u003e
            """

            xml = XmlDocument()
            xml.load_xml(f"""
            <toast scenario=\"reminder\"> 
              <visual>
                <binding template=\"ToastGeneric\">
                  <text>{title}</text>
                  <text>{message}</text>
                </binding>
              </visual>
              {actions_xml}
            </toast>
            """)
            notifier = ToastNotificationManager.create_toast_notifier(app_id)
            notifier.show(ToastNotification(xml))
        except Exception:
            # Silently ignore notification errors
            pass

    def _event_key(self, ev: dict) -> str:
        uid = ev.get("uid")
        if uid:
            return uid
        start = ev.get("start")
        summary = ev.get("summary", "")
        return f"{summary}|{start.isoformat() if isinstance(start, datetime) else ''}"

    def _maybe_send_notifications(self, upcoming: list[dict], now_local: datetime):
        if not self._notifications.get("enabled", False):
            return
        offsets = self._notifications.get("offsets_minutes", [10, 0])
        try:
            offsets = list({int(x) for x in offsets if int(x) >= 0})
        except Exception:
            offsets = [10, 0]
        nfilter = self._notifications.get("filter", "all")
        for ev in upcoming:
            start: datetime | None = ev.get("start")
            if not isinstance(start, datetime):
                continue
            # Apply notifications.filter
            is_online = self._is_online(ev)
            if nfilter == "online" and not is_online:
                continue
            if nfilter == "in_person" and is_online:
                continue
            show_tz = self._tz if self._tz else now_local.tzinfo
            start_local = start.astimezone(show_tz)
            delta_min = int((start_local - now_local).total_seconds() // 60)
            if delta_min in offsets:
                key = (self._event_key(ev), delta_min)
                if key not in self._notified:
                    title, message = self._format_notification(ev, now_local, is_now=(delta_min == 0))
                    url = self._extract_first_url(ev) if self._is_online(ev) else None
                    self._show_toast(title, message, url=url, minutes_remaining=delta_min)
                    self._notified.add(key)
