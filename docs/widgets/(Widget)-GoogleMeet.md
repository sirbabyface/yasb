# Google Meet Widget

This widget displays upcoming meetings from an iCal URL and provides a popup list with click-to-join and copy-link actions.

- Source of events: iCal URL taken from an environment variable (default: YASB_ICAL_URL)
- Label placeholders: {count}/{data}, {next_time}, {countdown}, {next_summary}
- Popup: shows start/end times, online indicator, and a copy URL button
- Filtering: all | online | in_person
- Horizon window: Until end-of-day (default) or next N hours
- Timezone: Optional explicit timezone for horizon and formatting

## Example Configuration

```yaml
google_meet:
  type: "yasb.google_meet.GoogleMeetWidget"
  options:
    env_var: "YASB_ICAL_URL"         # iCal URL environment variable name
    meeting_filter: "online"         # all | online | in_person
    timezone_name: "Europe/Lisbon"   # optional; IANA timezone
    horizon_hours: null               # null=end-of-day, or number of hours
    next_time_format: "%H:%M"

    # Label placeholders
    label: "<span></span> {count}"
    label_alt: "{next_summary} at {next_time} ({countdown})"

    update_interval: 600000           # 10 minutes (ms)

    # Popup menu options
    menu:
      blur: true
      round_corners: true
      round_corners_type: "normal"
      border_color: "System"
      alignment: "right"
      direction: "down"
      offset_top: 6
      offset_left: 0

    icons:
      online: "🔗"  # displayed when event is online
      time: "🕒"    # prefix for time in popup
      copy: "📋"    # click to copy meeting URL

    callbacks:
      on_left: "toggle_menu"
      on_middle: "do_nothing"
      on_right: "toggle_label"

    container_padding:
      top: 0
      left: 0
      bottom: 0
      right: 0

    animation:
      enabled: true
      type: "fadeInOut"
      duration: 200
```

## Placeholders
- {count} or {data}: number of upcoming events in the current horizon
- {next_time}: next event start time formatted with `next_time_format` (and `timezone_name` if set)
- {countdown}: time until next event (e.g., `1h 05m` or `now` if started)
- {next_summary}: next event title

## Environment Variable
Set the environment variable to your iCal URL:
- PowerShell (temporary in current shell):
  ```pwsh
  $env:YASB_ICAL_URL = "https://example.com/calendar.ics"
  ```
- Persist across sessions:
  ```pwsh
  setx YASB_ICAL_URL "https://example.com/calendar.ics"
  ```
  Restart YASB (and shells) to pick up changes.

## Styling Hooks
Widget root: `.google-meet-widget`
Popup root: `.google-meet-menu`
Row classes: `.google-meet-menu .item`, `.google-meet-menu .item.online`
Selected row (keyboard navigation): `.google-meet-menu .item[selected="true"]`
Row parts: `.icon`, `.time`, `.title`, `.copy`

### Minimal CSS Skeleton
```css
/* Widget container */
.google-meet-widget {}
.google-meet-widget .widget-container {}
.google-meet-widget .label {}

/* Popup base */
.google-meet-menu {
  min-width: 360px;
  max-width: 480px;
  max-height: 600px;
}
.google-meet-menu .empty {
  padding: 12px;
  color: #a6adc8;
}

/* Event row (Qt stylesheets do not support flex/align-items/gap) */
.google-meet-menu .item {
  padding: 6px 8px; /* spacing handled by layout and margins below */
}
.google-meet-menu .item:hover {
  background-color: rgba(255,255,255,0.06);
}
/* Selected via keyboard navigation */
.google-meet-menu .item[selected="true"] {
  background-color: rgba(0, 153, 255, 0.20);
}
.google-meet-menu .item.online .icon {
  color: #3fb950; /* green-ish for online */
}

/* Parts */
.google-meet-menu .icon {
  font-size: 14px;
  min-width: 18px;
  margin-right: 8px; /* simulate gap */
}
.google-meet-menu .time {
  color: #a6adc8;
  font-size: 12px;
  min-width: 150px; /* align columns */
  margin-right: 8px; /* simulate gap */
}
.google-meet-menu .title {
  color: #e6e6e6;
  font-size: 13px;
}
.google-meet-menu .copy {
  margin-left: 8px;
  color: #c8d3f5;
}
.google-meet-menu .copy:hover {
  color: #ffffff;
}
```

### Alternative Compact Style
```css
.google-meet-menu {
  min-width: 320px;
}
.google-meet-menu .item {
  padding: 4px 6px;
}
.google-meet-menu .time {
  font-size: 11px;
  min-width: 120px;
}
.google-meet-menu .title {
  font-size: 12px;
}
```

## Tips
- If you prefer only online meeting with links, set `meeting_filter: online`..
- Use a more detailed `next_time_format` when also setting a fixed `timezone_name` (e.g., `%Y-%m-%d %H:%M`).

