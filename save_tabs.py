#!/usr/bin/env python3
"""Capture all open Microsoft Edge browser tabs and save to markdown.

Reads tab group data from Edge's Preferences file (always current) and
ungrouped tabs from the most recent readable session file.
Works with a running Edge browser — no restart or special flags needed.
"""

import argparse
import json
import os
import struct
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

DEFAULT_OUTPUT_DIR = "snapshots"
SCRIPT_DIR = Path(__file__).resolve().parent
INTERNAL_SCHEMES = ("edge://", "chrome://", "about:", "chrome-extension://", "devtools://")


def get_edge_profile_dir(profile: str = "Default") -> Path:
    """Locate the Edge profile directory."""
    local_appdata = os.environ.get("LOCALAPPDATA")
    if not local_appdata:
        local_appdata = str(Path.home() / "AppData" / "Local")
    return Path(local_appdata) / "Microsoft" / "Edge" / "User Data" / profile


def load_tab_groups(profile_dir: Path) -> list[dict]:
    """Load tab groups from Edge Preferences file.

    Returns a list of groups, each with 'title' and 'tabs' (list of {title, url}).
    """
    prefs_path = profile_dir / "Preferences"
    try:
        data = json.loads(prefs_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        print(f"Error: Preferences file not found at {prefs_path}", file=sys.stderr)
        print("Make sure Microsoft Edge is installed.", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"Error: Could not parse {prefs_path}", file=sys.stderr)
        sys.exit(1)

    groups = []
    for g in data.get("tab_groups", []):
        title = g.get("tabGroupTitle", "Untitled Group")
        tabs = []
        for t in g.get("tabsInGroup", []):
            tab_url = t.get("url", "")
            tab_title = t.get("title", "").strip() or tab_url
            tabs.append({"title": tab_title, "url": tab_url})
        if tabs:
            groups.append({"title": title, "tabs": tabs})
    return groups


def find_readable_tabs_file(profile_dir: Path) -> Path | None:
    """Find the most recent readable SNSS Tabs file in the Sessions directory."""
    sessions_dir = profile_dir / "Sessions"
    if not sessions_dir.exists():
        return None

    candidates = []
    for f in sessions_dir.iterdir():
        if not f.name.startswith("Tabs_"):
            continue
        try:
            with open(f, "rb") as fh:
                magic = fh.read(4)
                if magic == b"SNSS":
                    candidates.append(f)
        except (PermissionError, OSError):
            continue

    if not candidates:
        return None
    # Return the one with the most recent modification time
    return max(candidates, key=lambda p: p.stat().st_mtime)


def parse_snss_tabs(tabs_file: Path) -> list[dict]:
    """Parse an SNSS Tabs file to extract tab URLs and titles.

    Returns list of {title, url} for each tab (latest navigation entry per tab).
    """
    with open(tabs_file, "rb") as f:
        data = f.read()

    if data[:4] != b"SNSS":
        return []

    # Parse command 1 entries (UpdateTabNavigation) which contain URL + title
    tab_navs = {}  # tab_id -> (url, title, nav_index)
    pos = 8  # skip header (magic + version)
    while pos < len(data) - 2:
        size = struct.unpack_from("<H", data, pos)[0]
        if size == 0 or pos + 2 + size > len(data):
            break
        cmd_id = data[pos + 2]
        if cmd_id == 1:
            payload = data[pos + 3 : pos + 2 + size]
            try:
                p = 4  # skip internal pickle size
                tab_id = struct.unpack_from("<I", payload, p)[0]
                p += 4
                nav_idx = struct.unpack_from("<I", payload, p)[0]
                p += 4
                url_len = struct.unpack_from("<I", payload, p)[0]
                p += 4
                url = payload[p : p + url_len].decode("utf-8", errors="ignore").rstrip("\x00")
                p += url_len
                # Skip null padding before title
                while p < len(payload) - 4 and payload[p] == 0:
                    p += 1
                title = ""
                if p + 4 <= len(payload):
                    title_len = struct.unpack_from("<I", payload, p)[0]
                    p += 4
                    if title_len < 2000 and p + title_len * 2 <= len(payload):
                        title = payload[p : p + title_len * 2].decode(
                            "utf-16-le", errors="ignore"
                        )
                # Keep the highest navigation index per tab (most recent page)
                if tab_id not in tab_navs or nav_idx > tab_navs[tab_id][2]:
                    tab_navs[tab_id] = (url, title, nav_idx)
            except (struct.error, IndexError, UnicodeDecodeError):
                pass
        pos += 2 + size

    tabs = []
    for url, title, _ in tab_navs.values():
        if any(url.startswith(s) for s in INTERNAL_SCHEMES):
            continue
        tabs.append({"title": title.strip() or url, "url": url})
    return tabs


def escape_markdown_link(title: str, url: str) -> str:
    """Format a single tab as a markdown link, handling special chars."""
    safe_title = title.replace("[", "\\[").replace("]", "\\]")
    safe_url = url.replace(")", "%29")
    return f"- [{safe_title}]({safe_url})"


def format_markdown(
    groups: list[dict], ungrouped: list[dict], timestamp: datetime
) -> str:
    """Format all tabs as a markdown document."""
    grouped_count = sum(len(g["tabs"]) for g in groups)
    total = grouped_count + len(ungrouped)
    lines = [
        f"# Browser Tabs - {timestamp.strftime('%Y-%m-%d %H:%M')}",
        "",
        f"Captured {total} tabs ({grouped_count} in {len(groups)} groups, "
        f"{len(ungrouped)} ungrouped) at {timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]
    for group in groups:
        lines.append(f"## {group['title']}")
        lines.append("")
        for tab in group["tabs"]:
            lines.append(escape_markdown_link(tab["title"], tab["url"]))
        lines.append("")
    if ungrouped:
        lines.append("## Ungrouped")
        lines.append("")
        for tab in ungrouped:
            lines.append(escape_markdown_link(tab["title"], tab["url"]))
        lines.append("")
    return "\n".join(lines)


def format_json_output(
    groups: list[dict], ungrouped: list[dict], timestamp: datetime
) -> str:
    """Format all tabs as JSON."""
    grouped_count = sum(len(g["tabs"]) for g in groups)
    output = {
        "timestamp": timestamp.isoformat(),
        "total_tabs": grouped_count + len(ungrouped),
        "total_groups": len(groups),
        "groups": groups,
        "ungrouped": ungrouped,
    }
    return json.dumps(output, indent=2, ensure_ascii=False)


def resolve_output_path(output_dir: Path, date: datetime) -> Path:
    """Find next available filename for today's date."""
    output_dir.mkdir(parents=True, exist_ok=True)
    base = date.strftime("%Y-%m-%d")
    path = output_dir / f"{base}.md"
    if not path.exists():
        return path
    counter = 2
    while True:
        path = output_dir / f"{base}_{counter}.md"
        if not path.exists():
            return path
        counter += 1


def main():
    parser = argparse.ArgumentParser(
        description="Capture open Microsoft Edge tabs to markdown."
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory relative to script location (default: {DEFAULT_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--profile",
        default="Default",
        help="Edge profile name (default: Default)",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print to stdout instead of saving to file",
    )
    parser.add_argument(
        "--format",
        choices=["markdown", "json"],
        default="markdown",
        dest="output_format",
        help="Output format (default: markdown)",
    )
    args = parser.parse_args()

    now = datetime.now()
    profile_dir = get_edge_profile_dir(args.profile)
    groups = load_tab_groups(profile_dir)

    # Find ungrouped tabs from SNSS session file
    grouped_urls = set()
    for g in groups:
        for t in g["tabs"]:
            grouped_urls.add(t["url"])

    ungrouped = []
    tabs_file = find_readable_tabs_file(profile_dir)
    if tabs_file:
        all_snss_tabs = parse_snss_tabs(tabs_file)
        # Deduplicate and exclude grouped URLs
        seen_urls = set()
        for t in all_snss_tabs:
            if t["url"] not in grouped_urls and t["url"] not in seen_urls:
                seen_urls.add(t["url"])
                ungrouped.append(t)
        # Sort by title for readability
        ungrouped.sort(key=lambda t: t["title"].lower())

    if not groups and not ungrouped:
        print("No tabs found.")
        sys.exit(0)

    if args.output_format == "json":
        content = format_json_output(groups, ungrouped, now)
    else:
        content = format_markdown(groups, ungrouped, now)

    if args.stdout:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        print(content)
    else:
        out_dir = SCRIPT_DIR / args.output_dir
        out_path = resolve_output_path(out_dir, now)
        out_path.write_text(content, encoding="utf-8")
        total = sum(len(g["tabs"]) for g in groups) + len(ungrouped)
        print(f"Saved {total} tabs ({len(groups)} groups + {len(ungrouped)} ungrouped) to {out_path}")


if __name__ == "__main__":
    main()
