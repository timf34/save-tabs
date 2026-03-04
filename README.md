# save-tabs

Snapshot all open Microsoft Edge tabs to a markdown file. Zero dependencies — just Python 3.7+.

Works with your currently running browser. No restart, no flags, no setup.

## Usage

```bash
python save_tabs.py                       # Save to snapshots/YYYY-MM-DD.md
python save_tabs.py --stdout              # Print to terminal
python save_tabs.py --format json --stdout # Output as JSON
python save_tabs.py -o my-tabs            # Custom output directory
```

Multiple runs on the same day auto-increment: `2026-03-04.md`, `2026-03-04_2.md`, etc.

## How it works

- **Tab groups** — read from Edge's `Preferences` JSON file, which is updated live while Edge runs. Always current.
- **Ungrouped tabs** — parsed from the most recent unlocked SNSS session file in `Sessions/`. These reflect the state from the previous session, so tabs opened since the last restart won't appear.

## Limitations

- **Ungrouped tabs may be slightly stale** — from the previous session file, not the live one.
- **Windows only** — reads from `%LOCALAPPDATA%\Microsoft\Edge\User Data\`.
