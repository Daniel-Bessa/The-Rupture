## Context for Claude Code — WarcraftLogs Guild Audit Tool

I'm Daniel, an officer in the WoW guild "The Rupture" (EU servers). I've been building a Python tool with Claude.ai that pulls data from the WarcraftLogs API and generates raid audit reports. The code is in `wcl_craft_audit.py` with credentials in `wcl_config.txt`.

### What the tool does now:
1. **Authenticates** with WarcraftLogs v2 GraphQL API (client credentials flow)
2. **Fetches all boss kill fights** from a report (both splits — we run split raids)
3. **Pulls gear data** via `CombatantInfo` events endpoint (not `playerDetails` — that returned empty `combatantInfo`)
4. **Detects crafted items** by checking for bonus ID `8960` (Midnight crafting indicator)
5. **Maps character names to players** using a hardcoded roster (with Main/Alt designation and Tank/Healer/DPS role)
6. **Auto-detects Split 1 vs Split 2** — if a boss has 2 kills, first kill = Split 1, second = Split 2
7. **Fetches cast events** per fight to track health pots, combat pots, and defensive cooldowns with timestamps
8. **Outputs HTML** (dark theme, class colors, Wowhead tooltips, horizontal item layout)
9. **Outputs XLSX** with 4 tabs:
   - **Mains** — crafted gear per main character (Player | Char | Slot | Ilvl | Spark?)
   - **Alts** — same for alt characters
   - **Split 1** — spell usage per player per boss (Health Pots | Combat Pots | Defensives with timestamps)
   - **Split 2** — same for second split

### Key technical details:
- **masterData.actors** gives us player names/classes (field `subType` has the class, not `type`)
- **Gear comes from events endpoint** with `dataType: CombatantInfo`, NOT from `playerDetails` (which returns empty combatantInfo arrays)
- **Crafted items detected by bonus ID `8960`** in the item's `bonusIDs` array
- **Fights query uses no killType filter** to catch all kills (both splits of same boss)
- **Class colors for XLSX** use semi-transparent hex (`77C41E3A` format) as row backgrounds with black text
- **Role ordering**: Tanks → Healers → DPS with separator rows

### What still needs work / known issues:
- **Craft Rank detection** shows "Unknown" — the bonus IDs for quality ranks in Midnight haven't been identified yet
- **Spark detection** is based on ilvl thresholds (252+ = spark used) which is approximate
- **The roster is hardcoded** — needs updating when players join/leave or swap mains
- **Only 26 players showing from first split** — second split's players should also appear (the fights query was recently fixed to include all kills)
- **Cast tracking** was just added and hasn't been tested with real data yet — spell name list may need expanding
- **Design polish** — the XLSX uses class-colored row backgrounds like the guild roster screenshots Daniel shared

### Roster structure:
The `_roster_raw` list maps: `(PlayerName, Role, [(CharName, "Main"/"Alt"), ...])`
- Role is per-player: "Tank", "Healer", or "DPS"
- Each player can have multiple characters across both splits
- Characters are matched case-insensitively

### How to run:
```
python wcl_craft_audit.py
```
It reads `wcl_config.txt` for CLIENT_ID, CLIENT_SECRET, and REPORT_URL. Type `all` when asked for fight IDs.

### What Daniel wants to build next:
- Keep improving the split tabs with spell tracking
- Eventually add more raid checks (avoidable damage, interrupt tracking)
- The goal is a weekly officer report tool that auto-generates after each raid
- Design should look clean and match WoW aesthetic (dark theme, class colors)

Read the code, run it, and let's keep iterating!