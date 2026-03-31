#!/usr/bin/env python3
"""
update_from_sheets.py — WA Topline Intelligence Auto-Updater
=============================================================
Location: WAtopline/scripts/update_from_sheets.py

Reads the Google Sheet, filters to the rolling last 10 days
where Cold-Hot >= 2, and injects into index.html:
  1. State chips colored by cold-hot density
  2. Top issues summary for the period
  3. An inline SVG national heat map
  4. Full entry cards grouped by state

Usage:
    python scripts/update_from_sheets.py

The script is IDEMPOTENT — safe to run multiple times.
"""

import datetime
import re
import sys
from collections import Counter
from html import escape
from pathlib import Path

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ── PATH RESOLUTION ─────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT  = SCRIPT_DIR.parent
CREDS_FILE = SCRIPT_DIR / "hmh-index-updates-d657b7e7e128.json"
INDEX_HTML = REPO_ROOT / "index.html"

# ── GOOGLE SHEETS CONFIG ─────────────────────────────────────────────────────────
SHEET_URL = "https://docs.google.com/spreadsheets/d/1CHepyOinqrY5nSQqx66C_HAQ0NkKVXBWj5QW13OHfw8"
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ── FILTER CONFIG ────────────────────────────────────────────────────────────────
DAYS_WINDOW          = 7
DATE_COL             = "Date"
ACTION_COL           = "Cold-Hot"
ACTION_MIN_THRESHOLD = 3            # Only Score 3 items
HOT_THRESHOLD        = 3
BAD_GOOD_COL         = "Bad-Good"
ISSUE_COL            = "Issue"
SECTOR_COL           = "Sector"
DISTRICT_COL         = "District Validatation"
CONTENT_COL          = "Content Text"
LINK_COL             = "Link"
WHY_COL              = "WHY THIS MATTERS1"
SOUNDBYTE_COL        = "SoundByte"
STATE_COL            = "State"
SHOW_BAD_GOOD_FOR    = {"Budget", "Finance"}

# ── HTML INJECTION MARKERS ───────────────────────────────────────────────────────
MARKER_BEGIN = "<!-- UPDATES:BEGIN -->"
MARKER_END   = "<!-- UPDATES:END -->"

PANEL_MARKERS = {
    "LAST_GENERATED":    lambda v: f"📅 <strong>Latest Update:</strong> {v}",
    "DATA_SOURCE":       lambda v: f"📂 <strong>Source:</strong> {v}",
    "DATE_WINDOW":       lambda v: f"📆 <strong>Window:</strong> {v}",
    "ROW_COUNT":         lambda v: f"📊 <strong>Rows Processed:</strong> {v}",
    "GENERATION_METHOD": lambda v: (
        '<span class="auto-status active">'
        '<span class="auto-dot"></span> Daily Updates Active</span>'
    ),
}

# ── US STATE SVG PATHS (simplified borders) ──────────────────────────────────────
# Keyed by state name → {"abbrev", "path", "tx", "ty"} for label placement
# These are simplified path outlines for a 960×600 US map projection
STATE_GEO = {
    "Alabama":        {"abbrev": "AL", "path": "M628,396 L628,440 L620,453 L613,452 L611,396 Z", "tx": 620, "ty": 425},
    "Alaska":         {"abbrev": "AK", "path": "M161,485 L195,485 L195,510 L161,510 Z", "tx": 178, "ty": 500},
    "Arizona":        {"abbrev": "AZ", "path": "M222,378 L278,378 L278,438 L222,438 Z", "tx": 250, "ty": 410},
    "Arkansas":       {"abbrev": "AR", "path": "M557,393 L601,393 L601,432 L557,432 Z", "tx": 579, "ty": 413},
    "California":     {"abbrev": "CA", "path": "M138,280 L178,280 L190,380 L155,430 L138,380 Z", "tx": 155, "ty": 350},
    "Colorado":       {"abbrev": "CO", "path": "M300,300 L378,300 L378,350 L300,350 Z", "tx": 339, "ty": 327},
    "Connecticut":    {"abbrev": "CT", "path": "M808,225 L830,225 L830,242 L808,242 Z", "tx": 819, "ty": 234},
    "DC":             {"abbrev": "DC", "path": "M755,300 L765,300 L765,310 L755,310 Z", "tx": 760, "ty": 306},
    "Delaware":       {"abbrev": "DE", "path": "M775,290 L790,290 L790,315 L775,315 Z", "tx": 782, "ty": 303},
    "Florida":        {"abbrev": "FL", "path": "M642,455 L700,440 L715,490 L680,520 L645,490 Z", "tx": 678, "ty": 475},
    "Georgia":        {"abbrev": "GA", "path": "M645,396 L690,396 L690,453 L645,453 Z", "tx": 667, "ty": 425},
    "Hawaii":         {"abbrev": "HI", "path": "M260,490 L290,490 L290,515 L260,515 Z", "tx": 275, "ty": 503},
    "Idaho":          {"abbrev": "ID", "path": "M225,170 L265,170 L265,270 L225,270 Z", "tx": 245, "ty": 225},
    "Illinois":       {"abbrev": "IL", "path": "M578,268 L608,268 L608,348 L578,348 Z", "tx": 593, "ty": 310},
    "Indiana":        {"abbrev": "IN", "path": "M618,275 L648,275 L648,345 L618,345 Z", "tx": 633, "ty": 312},
    "Iowa":           {"abbrev": "IA", "path": "M508,250 L575,250 L575,296 L508,296 Z", "tx": 541, "ty": 275},
    "Kansas":         {"abbrev": "KS", "path": "M420,330 L510,330 L510,370 L420,370 Z", "tx": 465, "ty": 352},
    "Kentucky":       {"abbrev": "KY", "path": "M630,340 L720,330 L720,365 L630,370 Z", "tx": 675, "ty": 350},
    "Louisiana":      {"abbrev": "LA", "path": "M557,440 L608,440 L608,490 L557,490 Z", "tx": 582, "ty": 465},
    "Maine":          {"abbrev": "ME", "path": "M830,135 L860,135 L860,200 L830,200 Z", "tx": 845, "ty": 168},
    "Maryland":       {"abbrev": "MD", "path": "M730,290 L775,280 L775,305 L730,310 Z", "tx": 752, "ty": 296},
    "Massachusetts":  {"abbrev": "MA", "path": "M808,210 L850,210 L850,225 L808,225 Z", "tx": 829, "ty": 218},
    "Michigan":       {"abbrev": "MI", "path": "M598,185 L660,185 L660,270 L598,270 Z", "tx": 635, "ty": 235},
    "Minnesota":      {"abbrev": "MN", "path": "M488,150 L548,150 L548,240 L488,240 Z", "tx": 518, "ty": 200},
    "Mississippi":    {"abbrev": "MS", "path": "M592,400 L618,400 L618,465 L592,465 Z", "tx": 605, "ty": 433},
    "Missouri":       {"abbrev": "MO", "path": "M520,310 L580,310 L580,385 L520,385 Z", "tx": 550, "ty": 348},
    "Montana":        {"abbrev": "MT", "path": "M270,130 L385,130 L385,195 L270,195 Z", "tx": 327, "ty": 165},
    "Nebraska":       {"abbrev": "NE", "path": "M390,275 L498,275 L498,315 L390,315 Z", "tx": 444, "ty": 297},
    "Nevada":         {"abbrev": "NV", "path": "M185,250 L230,250 L222,378 L185,378 Z", "tx": 205, "ty": 318},
    "New Hampshire":  {"abbrev": "NH", "path": "M825,165 L840,165 L840,210 L825,210 Z", "tx": 832, "ty": 190},
    "New Jersey":     {"abbrev": "NJ", "path": "M785,250 L802,250 L802,300 L785,300 Z", "tx": 793, "ty": 276},
    "New Mexico":     {"abbrev": "NM", "path": "M278,370 L345,370 L345,445 L278,445 Z", "tx": 311, "ty": 410},
    "New York":       {"abbrev": "NY", "path": "M740,185 L820,185 L820,250 L740,250 Z", "tx": 780, "ty": 220},
    "North Carolina": {"abbrev": "NC", "path": "M665,355 L775,345 L775,375 L665,385 Z", "tx": 720, "ty": 365},
    "North Dakota":   {"abbrev": "ND", "path": "M400,145 L488,145 L488,195 L400,195 Z", "tx": 444, "ty": 172},
    "Ohio":           {"abbrev": "OH", "path": "M660,265 L720,265 L720,330 L660,330 Z", "tx": 690, "ty": 300},
    "Oklahoma":       {"abbrev": "OK", "path": "M400,375 L520,370 L520,408 L400,412 Z", "tx": 465, "ty": 392},
    "Oregon":         {"abbrev": "OR", "path": "M140,165 L225,165 L225,230 L140,230 Z", "tx": 182, "ty": 200},
    "Pennsylvania":   {"abbrev": "PA", "path": "M720,245 L790,245 L790,280 L720,280 Z", "tx": 755, "ty": 264},
    "Rhode Island":   {"abbrev": "RI", "path": "M830,228 L845,228 L845,242 L830,242 Z", "tx": 837, "ty": 236},
    "South Carolina": {"abbrev": "SC", "path": "M680,380 L735,370 L735,410 L680,415 Z", "tx": 707, "ty": 392},
    "South Dakota":   {"abbrev": "SD", "path": "M400,200 L488,200 L488,255 L400,255 Z", "tx": 444, "ty": 230},
    "Tennessee":      {"abbrev": "TN", "path": "M590,365 L710,355 L710,385 L590,392 Z", "tx": 650, "ty": 374},
    "Texas":          {"abbrev": "TX", "path": "M370,410 L520,408 L530,510 L430,520 L370,470 Z", "tx": 460, "ty": 460},
    "US":             {"abbrev": "US", "path": "", "tx": 0, "ty": 0},
    "Utah":           {"abbrev": "UT", "path": "M248,268 L300,268 L300,358 L248,358 Z", "tx": 274, "ty": 315},
    "Vermont":        {"abbrev": "VT", "path": "M810,165 L825,165 L825,210 L810,210 Z", "tx": 817, "ty": 190},
    "Virginia":       {"abbrev": "VA", "path": "M680,320 L775,305 L775,350 L680,360 Z", "tx": 730, "ty": 335},
    "Washington":     {"abbrev": "WA", "path": "M150,110 L230,110 L230,170 L150,170 Z", "tx": 190, "ty": 143},
    "West Virginia":  {"abbrev": "WV", "path": "M700,300 L740,290 L740,340 L700,340 Z", "tx": 720, "ty": 320},
    "Wisconsin":      {"abbrev": "WI", "path": "M548,180 L600,180 L600,260 L548,260 Z", "tx": 574, "ty": 222},
    "Wyoming":        {"abbrev": "WY", "path": "M290,200 L370,200 L370,260 L290,260 Z", "tx": 330, "ty": 232},
}


# ── STEP 1: AUTHENTICATE & FETCH ─────────────────────────────────────────────────
def fetch_records():
    if not CREDS_FILE.exists():
        print(f"ERROR: Credentials file not found: {CREDS_FILE}")
        sys.exit(1)

    try:
        creds  = ServiceAccountCredentials.from_json_keyfile_name(str(CREDS_FILE), SCOPE)
        client = gspread.authorize(creds)
        print("✓ Authenticated with Google API")
    except Exception as e:
        print(f"ERROR authenticating with Google API: {e}")
        sys.exit(1)

    try:
        spreadsheet = client.open_by_url(SHEET_URL)
        try:
            sheet = spreadsheet.worksheet("K12 Track 2026")
            print("✓ Opened tab: K12 Track 2026")
        except gspread.exceptions.WorksheetNotFound:
            print("⚠ Tab 'K12 Track 2026' not found — falling back to first sheet")
            sheet = spreadsheet.sheet1

        all_values = sheet.get_all_values()
        if not all_values:
            print("⚠ Sheet appears empty")
            return []

        headers = all_values[0]
        records = []
        for row in all_values[1:]:
            row = list(row) + [""] * (len(headers) - len(row))
            record = {
                headers[i]: row[i]
                for i in range(len(headers))
                if headers[i].strip()
            }
            records.append(record)
        print(f"✓ Retrieved {len(records)} records from Google Sheet")
        return records

    except gspread.exceptions.SpreadsheetNotFound:
        print("ERROR: Spreadsheet not found. Check the URL and sharing permissions.")
        sys.exit(1)
    except gspread.exceptions.APIError as e:
        print(f"ERROR: Google Sheets API error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR accessing Google Sheet: {e}")
        sys.exit(1)


# ── STEP 2: FILTER ───────────────────────────────────────────────────────────────
def safe(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s


def filter_records(records):
    cutoff = datetime.datetime.utcnow() - datetime.timedelta(days=DAYS_WINDOW)
    updates = []

    for record in records:
        date_str = str(record.get(DATE_COL, "") or record.get("date", "")).strip()
        if not date_str:
            continue

        date_obj = None
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%Y/%m/%d"):
            try:
                date_obj = datetime.datetime.strptime(date_str, fmt)
                break
            except ValueError:
                continue

        if date_obj is None:
            continue

        try:
            cold_hot = int(record.get(ACTION_COL, 0) or 0)
        except (ValueError, TypeError):
            cold_hot = 0

        if date_obj >= cutoff and cold_hot >= ACTION_MIN_THRESHOLD:
            updates.append({**record, "_date_obj": date_obj, "_date_str": date_str})

    updates.sort(key=lambda r: (
        str(r.get(STATE_COL, "")),
        -int(r.get(ACTION_COL, 0) or 0),
        str(r.get(ISSUE_COL, ""))
    ))
    print(f"✓ {len(updates)} qualifying records (last {DAYS_WINDOW} days, Cold-Hot >= {ACTION_MIN_THRESHOLD})")
    return updates


# ── STEP 3: BUILD HTML ───────────────────────────────────────────────────────────

def score_class(cold_hot):
    try:
        return {1: "score1", 2: "score2", 3: "score3"}.get(int(cold_hot), "score2")
    except (TypeError, ValueError):
        return "score2"


def build_entry_card(record):
    state     = safe(record.get(STATE_COL))
    sector    = safe(record.get(SECTOR_COL))
    issue     = safe(record.get(ISSUE_COL))
    district  = safe(record.get(DISTRICT_COL))
    content   = safe(record.get(CONTENT_COL))
    link      = safe(record.get(LINK_COL))
    why       = safe(record.get(WHY_COL))
    soundbyte = safe(record.get(SOUNDBYTE_COL))

    try:
        cold_hot = int(record.get(ACTION_COL, 0) or 0)
    except (TypeError, ValueError):
        cold_hot = 0

    try:
        bad_good_int = int(record.get(BAD_GOOD_COL, 0) or 0)
    except (ValueError, TypeError):
        bad_good_int = 0

    try:
        display_date = record["_date_obj"].strftime("%B %-d, %Y")
    except Exception:
        display_date = record.get("_date_str", "")

    card_cls = score_class(cold_hot)

    if cold_hot == 3:
        score_tag = '<span class="score3-tag">🔥 Action Score: 3</span>'
    elif cold_hot == 2:
        score_tag = '<span class="score2-tag">Action Score: 2</span>'
    else:
        score_tag = f'<span class="score1-tag">Action Score: {cold_hot}</span>'

    bg_chip = ""
    if issue in SHOW_BAD_GOOD_FOR and bad_good_int:
        if bad_good_int == 1:
            bg_chip = '<span class="badgood-tag badgood-1">📉 Declining</span>'
        elif bad_good_int == 2:
            bg_chip = '<span class="badgood-tag badgood-2">➡ Flat</span>'
        elif bad_good_int == 3:
            bg_chip = '<span class="badgood-tag badgood-3">📈 Improving</span>'

    why_html = f'<span class="why-matters">{escape(why)}</span> ' if why else ""

    sb_banner = ""
    sb_extra_class = ""
    if soundbyte and soundbyte.startswith("http"):
        sb_extra_class = " has-soundbyte"
        sb_banner = (
            '  <div class="wa-video-callout">\n'
            '    <div class="wa-video-callout-icon">🎬</div>\n'
            '    <div class="wa-video-callout-text">\n'
            '      <span class="wa-video-callout-label">W/A Video Commentary</span>\n'
            f'      <a class="wa-video-link" href="{escape(soundbyte)}" target="_blank">Watch Analysis →</a>\n'
            '    </div>\n'
            '  </div>\n'
        )

    link_html = ""
    if link and link.startswith("http"):
        link_html = (
            f'\n        <br/><a class="source-link" href="{escape(link)}" target="_blank">'
            f'View Source →</a>'
        )

    return f"""\
  <div class="entry-card {card_cls}{sb_extra_class}">
{sb_banner}    <div class="entry-meta">
      <span class="meta-chip level">Level: {escape(sector)}</span>
      <span class="meta-chip issue">{escape(issue)}</span>
      <span class="meta-chip date">{escape(display_date)}</span>
      {bg_chip}
      <div class="actionable-score">{score_tag}</div>
    </div>
    <div class="district-badge">{escape(district)}</div>
    <div class="entry-content">
      {why_html}{escape(content)}{link_html}
    </div>
  </div>"""


def group_by_state(updates):
    groups = {}
    for r in updates:
        state = safe(r.get(STATE_COL, "Unknown"))
        groups.setdefault(state, []).append(r)
    return sorted(groups.items())


def build_state_section(state, records):
    abbrev = safe(records[0].get("State Abbrev", state[:2].upper()))
    count  = len(records)
    items_label = "item" if count == 1 else "items"
    cards  = "\n\n".join(build_entry_card(r) for r in records)
    state_id = "state-" + re.sub(r"[^a-z0-9]+", "-", state.lower()).strip("-")
    return f"""\
  <div class="state-section" id="{state_id}">
    <div class="state-header">
      <span class="state-name">{escape(state)}</span>
      <span class="state-abbrev">{escape(abbrev)}</span>
      <span class="state-count">{count} {items_label}</span>
    </div>

{cards}
    <div class="section-back-top"><a href="#page-top" class="back-to-top-link">↑ Back to top</a></div>
  </div>"""


def build_updates_html(updates):
    if not updates:
        return f"""\
  <div class="entry-card score2">
    <div class="entry-content" style="text-align:center;padding:32px;color:var(--text-muted);">
      No qualifying items found in the last {DAYS_WINDOW} days.
    </div>
  </div>"""

    sections = []
    for state, records in group_by_state(updates):
        sections.append(build_state_section(state, records))
    return "\n\n".join(sections)


# ── STATE HEAT DATA ──────────────────────────────────────────────────────────────

def compute_state_heat(updates):
    """Return dict: state_name → {total, hot3, abbrev, id}."""
    state_info = {}
    for r in updates:
        state = safe(r.get(STATE_COL, "Unknown"))
        if state not in state_info:
            abbrev   = safe(r.get("State Abbrev", state[:2].upper()))
            state_id = "state-" + re.sub(r"[^a-z0-9]+", "-", state.lower()).strip("-")
            state_info[state] = {"abbrev": abbrev, "hot3": 0, "total": 0, "id": state_id}
        state_info[state]["total"] += 1
        try:
            if int(r.get(ACTION_COL, 0) or 0) >= HOT_THRESHOLD:
                state_info[state]["hot3"] += 1
        except (ValueError, TypeError):
            pass
    return state_info


def heat_class(n):
    """Map count of Score-3 items to a heat tier."""
    if n == 0: return "hot-0"
    if n == 1: return "hot-1"
    if n == 2: return "hot-2"
    return "hot-3"


def heat_fill(n):
    """Map count of Score-3 items to an SVG fill color."""
    if n == 0: return "#DDE5F0"
    if n == 1: return "#E8ECF2"
    if n == 2: return "#FFF8CC"
    if n == 3: return "#FFD0D0"
    if n == 4: return "#FFB0B0"
    return "#FF8888"


# ── BUILD TOC CHIPS ──────────────────────────────────────────────────────────────

def build_toc_html(state_info):
    chips = []
    for state in sorted(state_info):
        info = state_info[state]
        cls  = heat_class(info["hot3"])
        chips.append(
            f'      <a href="#{info["id"]}" class="toc-chip {cls}" '
            f'title="{escape(state)} — {info["total"]} items, {info["hot3"]} hot">'
            f'{escape(info["abbrev"])}</a>'
        )
    return '    <div class="toc-chips">\n' + "\n".join(chips) + '\n    </div>'


# ── BUILD TOP ISSUES ─────────────────────────────────────────────────────────────

def build_top_issues_html(updates):
    issue_counts = Counter(safe(r.get(ISSUE_COL)) for r in updates if safe(r.get(ISSUE_COL)))
    ranked = issue_counts.most_common(5)
    if not ranked:
        return '<div class="top-issues-list">\n      </div>'
    rows = []
    for rank, (issue, count) in enumerate(ranked, 1):
        rows.append(
            f'        <div class="top-issue">'
            f'<span class="top-issue-rank">{rank}</span>'
            f'<span class="top-issue-name">{escape(issue)}</span>'
            f'<span class="top-issue-count">{count}</span>'
            f'</div>'
        )
    return '<div class="top-issues-list">\n' + "\n".join(rows) + '\n      </div>'


# ── BUILD SVG HEAT MAP ───────────────────────────────────────────────────────────

def build_svg_map(state_info):
    """Generate an inline SVG national map colored by cold-hot density."""
    shapes = []
    labels = []

    for state_name, geo in STATE_GEO.items():
        if not geo["path"]:
            continue
        info = state_info.get(state_name)
        if info:
            fill   = heat_fill(info["hot3"])
            stroke = "#5A6880"
            total  = info["total"]
            hot3   = info["hot3"]
            title  = f"{state_name}: {total} items, {hot3} high-heat"
        else:
            fill   = "#EDF1F7"
            stroke = "#B8C4D4"
            title  = f"{state_name}: no items"

        shapes.append(
            f'    <path d="{geo["path"]}" fill="{fill}" stroke="{stroke}" '
            f'stroke-width="1" opacity="0.9">'
            f'<title>{escape(title)}</title></path>'
        )
        if info and geo["tx"] > 0:
            labels.append(
                f'    <text x="{geo["tx"]}" y="{geo["ty"]}" '
                f'font-size="9" font-weight="700" fill="#0D1B2E" '
                f'text-anchor="middle" dominant-baseline="central" '
                f'style="pointer-events:none;">{escape(geo["abbrev"])}</text>'
            )

    # Legend
    legend_items = [
        ("#EDF1F7", "No data"),
        ("#DDE5F0", "0 hot"),
        ("#FFF8CC", "1–2 hot"),
        ("#FFD0D0", "3 hot"),
        ("#FF8888", "5+ hot"),
    ]
    legend = []
    for i, (color, label) in enumerate(legend_items):
        x = 70 + i * 100
        legend.append(f'    <rect x="{x}" y="530" width="14" height="14" fill="{color}" stroke="#9AAFC8" stroke-width="0.5" rx="2"/>')
        legend.append(f'    <text x="{x+20}" y="541" font-size="10" fill="#5A6880">{label}</text>')

    svg_lines = [
        '<svg viewBox="60 100 820 470" xmlns="http://www.w3.org/2000/svg" '
        'style="width:100%;max-width:900px;height:auto;display:block;margin:0 auto 32px;">',
        '  <g>',
    ] + shapes + labels + [
        '  </g>',
    ] + legend + [
        '</svg>',
    ]
    return "\n".join(svg_lines)


# ── INJECTION FUNCTIONS ──────────────────────────────────────────────────────────

def inject_between(html, begin_marker, end_marker, content):
    pattern = re.compile(re.escape(begin_marker) + r".*?" + re.escape(end_marker), re.DOTALL)
    if not pattern.search(html):
        print(f"  ⚠ Markers {begin_marker} not found — skipping")
        return html
    return pattern.sub(f"{begin_marker}\n{content}\n  {end_marker}", html)


def update_panel_marker(html, marker_key, new_inner_html):
    begin = f"<!-- PYTHON:{marker_key} -->"
    end   = f"<!-- PYTHON:END_{marker_key} -->"
    pattern = re.compile(re.escape(begin) + r".*?" + re.escape(end), re.DOTALL)
    if not pattern.search(html):
        return html
    return pattern.sub(f"{begin}\n      {new_inner_html}\n      {end}", html)


def update_script_config(html, latest_item_date, window_str, total, qualifying, source):
    config_block = f"""/* PYTHON:BEGIN_SCRIPT_CONFIG */
const MEMO_CONFIG = {{
  latestItemDate:   "{latest_item_date}",
  reportingWindow:  "{window_str}",
  totalRowsRead:    {total},
  qualifyingRows:   {qualifying},
  dataSource:       "{escape(source)}",
  generationMethod: "auto",
}};
/* PYTHON:END_SCRIPT_CONFIG */"""
    pattern = re.compile(
        r"/\* PYTHON:BEGIN_SCRIPT_CONFIG \*/.*?/\* PYTHON:END_SCRIPT_CONFIG \*/",
        re.DOTALL
    )
    if pattern.search(html):
        html = pattern.sub(config_block, html)
    return html


def update_static_dates(html, latest_display, cutoff_display, qualifying, hot_count, state_count):
    html = re.sub(r'(<div class="stat-number total">)\d+(</div>)', rf'\g<1>{qualifying}\2', html)
    html = re.sub(r'(<div class="stat-number hot">)\d+(</div>)', rf'\g<1>{hot_count}\2', html)
    html = re.sub(r'(<div class="stat-number states">)\d+(</div>)', rf'\g<1>{state_count}\2', html)
    html = re.sub(
        r'(<div class="stat-number"[^>]*>)\d+(</div>\s*<div class="stat-label">Days Covered)',
        rf'\g<1>{DAYS_WINDOW}\2', html
    )
    html = re.sub(
        r'(<span class="label">Date</span>\s*<span class="val">)[^<]*(</span>)',
        rf'\g<1>{latest_display}\2', html
    )
    html = re.sub(
        r'(<span class="label">Re</span>\s*<span class="val">)[^<]*(</span>)',
        rf'\g<1>Actionable Education Policy Developments — {cutoff_display} – {latest_display}\2', html
    )
    html = re.sub(
        r'(<span class="label">Source</span>\s*<span class="val">)[^<]*(</span>)',
        rf'\g<1>Whiteboard Advisors Review and Analysis — Action Score: 3 (Actionable) | Rolling {DAYS_WINDOW}-Day Window\2', html
    )
    html = re.sub(r'(<div class="memo-date-badge">)[^<]*(</div>)', rf'\g<1>{latest_display}\2', html)
    html = re.sub(
        r'(<div style="margin-top:4px;">)[^<]*(· Confidential</div>)',
        rf'\g<1>{latest_display} · Confidential\2', html
    )
    return html


# ── MAIN ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("  WA Topline Intelligence — update_from_sheets.py")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")

    print("[1] Fetching from Google Sheets...")
    records = fetch_records()
    total   = len(records)

    print(f"\n[2] Filtering (last {DAYS_WINDOW} days, Cold-Hot >= {ACTION_MIN_THRESHOLD})...")
    updates = filter_records(records)

    if updates:
        latest_dt       = max(r["_date_obj"] for r in updates)
        cutoff_dt       = min(r["_date_obj"] for r in updates)
        latest_item_str = latest_dt.strftime("%Y-%m-%d")
        latest_display  = latest_dt.strftime("%B %-d, %Y")
        cutoff_display  = cutoff_dt.strftime("%B %-d, %Y")
    else:
        now             = datetime.datetime.utcnow()
        latest_item_str = now.strftime("%Y-%m-%d")
        latest_display  = now.strftime("%B %-d, %Y")
        cutoff_display  = (now - datetime.timedelta(days=DAYS_WINDOW)).strftime("%B %-d, %Y")

    state_info  = compute_state_heat(updates)
    hot_count   = sum(1 for r in updates if int(r.get(ACTION_COL, 0) or 0) >= HOT_THRESHOLD)
    state_count = len(state_info)

    window_str = f"Rolling {DAYS_WINDOW} days · latest item: {latest_display}"
    rows_str   = f"{total} total → {len(updates)} qualifying"
    source_str = "Whiteboard Advisors Review and Analysis"

    print("\n[3] Building components...")
    updates_html    = build_updates_html(updates)
    toc_html        = build_toc_html(state_info)
    top_issues_html = build_top_issues_html(updates)
    svg_map_html    = build_svg_map(state_info)
    print(f"  ✓ {len(updates)} cards | {hot_count} high-heat | {state_count} states")

    print(f"\n[4] Reading {INDEX_HTML}...")
    if not INDEX_HTML.exists():
        print(f"ERROR: index.html not found at {INDEX_HTML}")
        sys.exit(1)
    html = INDEX_HTML.read_text(encoding="utf-8")

    print("\n[5] Injecting all sections...")
    html = inject_between(html, MARKER_BEGIN, MARKER_END, updates_html)
    html = inject_between(html, "<!-- TOC:BEGIN -->", "<!-- TOC:END -->", toc_html)
    html = inject_between(html, "<!-- TOP_ISSUES:BEGIN -->", "<!-- TOP_ISSUES:END -->", "      " + top_issues_html)
    html = inject_between(html, "<!-- HEATMAP:BEGIN -->", "<!-- HEATMAP:END -->", svg_map_html)

    print("\n[6] Updating dates and stats...")
    html = update_static_dates(html, latest_display, cutoff_display, len(updates), hot_count, state_count)

    print("\n[7] Updating automation panel...")
    for key in PANEL_MARKERS:
        if key == "LAST_GENERATED":
            val = PANEL_MARKERS[key](latest_display)
        elif key == "DATA_SOURCE":
            val = PANEL_MARKERS[key](source_str)
        elif key == "DATE_WINDOW":
            val = PANEL_MARKERS[key](window_str)
        elif key == "ROW_COUNT":
            val = PANEL_MARKERS[key](rows_str)
        else:
            val = PANEL_MARKERS[key](None)
        html = update_panel_marker(html, key, val)

    html = update_script_config(html, latest_item_str, window_str, total, len(updates), source_str)

    print(f"\n[8] Writing {INDEX_HTML}...")
    INDEX_HTML.write_text(html, encoding="utf-8")
    print("  ✓ Successfully updated index.html")

    print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"  Done — {len(updates)} entries | latest: {latest_display}")
    print(f"  Score 3: {hot_count} | States: {state_count}")
    print(f"  Window: {cutoff_display} → {latest_display}")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")


if __name__ == "__main__":
    main()
