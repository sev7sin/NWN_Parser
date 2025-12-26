# AetherLens (NWN Log Watcher)

AetherLens is a Windows desktop companion for Neverwinter Nights logs. It monitors a selected log file and surfaces structured insights in real time:

- **Combat**: parses attacks, damage, and rolls; tracks per-combat and overall averages and prevents double-counting through file-ingestion dedupe.
- **Entities**: aggregates stats per creature (AC bands, seen times, spells, abilities, and average d20 rolls) with reset/delete controls.
- **Social**: remembers speakers, aliases, and notes across sessions with per-player profiles.
- **Reminders**: keyword alerts with case/whole-word matching, match context, and subtle UI cues.

## Requirements
- .NET SDK 8.0+ (uses `net10.0-windows` target framework)
- Windows with access to NWN log files

## Build
```bash
dotnet build WatcherNWN/WatcherNWN.sln
```

## Run
```bash
dotnet run --project WatcherNWN/WatcherNWNApp
```
Then choose your NWN log file with **Select File** to start monitoring. The app persists dedupe state, dice stats, and social/reminder data under the app directory.

## Repository layout
- `WatcherNWN/WatcherNWNApp/` – WPF application source
- `.vscode/` – editor/debug configuration
- `Settings/`, `Stats.db`, `Social*.json`, `Reminders.json` – runtime data (created at first run)

## Credits
- Sev7Sin

## License
See [LICENSE](LICENSE) for details.
