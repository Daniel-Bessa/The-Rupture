#!/usr/bin/env python3
"""
WarcraftLogs Crafted Gear Audit Tool — Midnight Season 1
=========================================================
Pulls player gear from a WarcraftLogs report and identifies crafted items,
spark usage, crest tier, and embellishments. Outputs to .xlsx.

Usage:
    python wcl_craft_audit.py

You will be prompted for:
    - WarcraftLogs Client ID & Secret
    - Report code (the alphanumeric part of a WCL report URL)

Requirements:
    pip install requests openpyxl
"""

import sys
import json
import requests
from datetime import datetime
from html import escape
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Midnight Season 1 Constants ─────────────────────────────────────────────

# Crafting quality bonus IDs (WoW internal — ranks 1 through 5)
# These are the bonus IDs appended to items when they are player-crafted.
CRAFT_QUALITY_BONUS_IDS = {
    10249: 1, 10250: 2, 10251: 3, 10252: 4, 10253: 5,  # TWW-era IDs
    10255: 1, 10256: 2, 10257: 3, 10258: 4, 10259: 5,  # Midnight-era IDs (if changed)
    11109: 1, 11110: 2, 11111: 3, 11112: 4, 11113: 5,  # Alternate range seen in beta
}

# Bonus IDs that indicate an item was crafted via the crafting order / profession system
CRAFTED_INDICATOR_BONUS_IDS = {
    8960,   # Midnight crafted tag
    8791,   # Midnight crafted (often paired with 8960)
    9497,   # TWW-era "Crafted" tag
    9498,   # TWW-era "Crafted" tag (alternate)
    10222,  # Crafting work order
    10343,  # Recrafted
}

# Known embellishment bonus IDs for Midnight Season 1
# Add more as they're discovered — these help confirm an item is crafted
EMBELLISHMENT_BONUS_IDS = {
    # Placeholder — update with actual Midnight embellishment bonus IDs
    # Format: bonus_id: "Embellishment Name"
}

# Item level thresholds for Midnight Season 1 crafted gear
ILVL_TIERS = {
    "Myth (Spark + Myth Crests)":  (272, 285),
    "Hero (Spark + Hero Crests)":  (259, 272),
    "Epic Base (Spark, no crests)": (252, 259),
    "Rare (Veteran Crests)":       (233, 246),
    "Rare (Adventurer Crests)":    (214, 233),
    "Rare Base (no spark)":        (200, 214),
}

# Gear slot mapping from WCL slot IDs
SLOT_NAMES = {
    0: "Head", 1: "Neck", 2: "Shoulder", 3: "Shirt", 4: "Chest",
    5: "Waist", 6: "Legs", 7: "Feet", 8: "Wrist", 9: "Hands",
    10: "Ring 1", 11: "Ring 2", 12: "Trinket 1", 13: "Trinket 2",
    14: "Back", 15: "Main Hand", 16: "Off Hand",
}


# ─── WarcraftLogs API ────────────────────────────────────────────────────────

def get_access_token(client_id: str, client_secret: str) -> str:
    """Get OAuth2 access token using client credentials flow."""
    resp = requests.post(
        "https://www.warcraftlogs.com/oauth/token",
        data={"grant_type": "client_credentials"},
        auth=(client_id, client_secret),
    )
    if resp.status_code != 200:
        print(f"[ERROR] Auth failed ({resp.status_code}): {resp.text}")
        sys.exit(1)
    return resp.json()["access_token"]


def query_wcl(token: str, query: str, variables: dict = None) -> dict:
    """Execute a GraphQL query against WarcraftLogs v2 API."""
    resp = requests.post(
        "https://www.warcraftlogs.com/api/v2/client",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"query": query, "variables": variables or {}},
    )
    if resp.status_code != 200:
        print(f"[ERROR] API request failed ({resp.status_code}): {resp.text}")
        sys.exit(1)
    data = resp.json()
    if "errors" in data:
        print(f"[ERROR] GraphQL errors: {json.dumps(data['errors'], indent=2)}")
        sys.exit(1)
    return data["data"]


def fetch_report_info(token: str, report_code: str) -> dict:
    """Fetch basic report metadata (title, fights, actors, etc.)."""
    query = """
    query ($code: String!) {
        reportData {
            report(code: $code) {
                title
                startTime
                endTime
                guild { name server { name region { slug } } }
                fights {
                    id
                    name
                    kill
                    difficulty
                    encounterID
                    startTime
                    endTime
                }
                masterData(translate: true) {
                    actors(type: "Player") {
                        id
                        name
                        type
                        subType
                        server
                    }
                }
            }
        }
    }
    """
    return query_wcl(token, query, {"code": report_code})["reportData"]["report"]


def fetch_player_details(token: str, report_code: str, start_time: float = None, end_time: float = None, fight_ids: list = None) -> dict:
    """Fetch player info including gear for the report."""
    # Get player list
    if fight_ids:
        query = """
        query ($code: String!, $fightIDs: [Int]!) {
            reportData {
                report(code: $code) {
                    playerDetails(fightIDs: $fightIDs)
                }
            }
        }
        """
        variables = {"code": report_code, "fightIDs": fight_ids}
    else:
        query = """
        query ($code: String!, $startTime: Float!, $endTime: Float!) {
            reportData {
                report(code: $code) {
                    playerDetails(startTime: $startTime, endTime: $endTime)
                }
            }
        }
        """
        variables = {"code": report_code, "startTime": start_time, "endTime": end_time}
    details = query_wcl(token, query, variables)["reportData"]["report"]["playerDetails"]
    return details


def fetch_combatant_info_events(token: str, report_code: str, fight_id: int) -> list:
    """Fetch gear data via the events endpoint for a specific fight."""
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                events(dataType: CombatantInfo, fightIDs: [$fightID], limit: 500) {
                    data
                    nextPageTimestamp
                }
            }
        }
    }
    """
    variables = {"code": report_code, "fightID": fight_id}
    result = query_wcl(token, query, variables)
    events = result["reportData"]["report"]["events"]
    return events.get("data", [])


def fetch_cast_events(token: str, report_code: str, fight_id: int, start_time: float = None, end_time: float = None) -> list:
    """Fetch all cast events for a fight. Paginates if needed."""
    all_data = []
    next_ts = start_time
    
    for _ in range(10):  # Max 10 pages to avoid infinite loop
        query = """
        query ($code: String!, $fightID: Int!, $startTime: Float, $endTime: Float) {
            reportData {
                report(code: $code) {
                    events(dataType: Casts, fightIDs: [$fightID], startTime: $startTime, endTime: $endTime, hostilityType: Friendlies, limit: 10000) {
                        data
                        nextPageTimestamp
                    }
                }
            }
        }
        """
        variables = {"code": report_code, "fightID": fight_id}
        if next_ts is not None:
            variables["startTime"] = next_ts
        if end_time is not None:
            variables["endTime"] = end_time
        
        result = query_wcl(token, query, variables)
        events = result["reportData"]["report"]["events"]
        all_data.extend(events.get("data", []))
        next_ts = events.get("nextPageTimestamp")
        if next_ts is None:
            break
    
    return all_data


# ─── Tracked Spells (matched by abilityGameID) ───────────────────────────────
# WCL cast events return abilityGameID (numeric), not ability names.

# Health Potions & Healthstone
HEALTH_POT_IDS = {
    5512,    # Healthstone
    432112,  # Algari Healing Potion (TWW)
    431924,  # Algari Healing Potion (alternate)
    # Add Midnight health potion IDs here when known
}

# Combat Potions
COMBAT_POT_IDS = {
    431932,  # Tempered Potion (TWW)
    432098,  # Potion of Unwavering Focus (TWW)
    431945,  # Light's Potential (TWW)
    432106,  # Void-Shrouded Tincture (TWW)
    # Add Midnight combat potion IDs here when known
}

# Class Defensives
CLASS_DEFENSIVE_IDS = {
    # Death Knight
    48792,   # Icebound Fortitude
    48707,   # Anti-Magic Shell
    55233,   # Vampiric Blood
    49028,   # Dancing Rune Weapon
    # Demon Hunter
    198589,  # Blur
    196718,  # Darkness
    196555,  # Netherwalk
    187827,  # Metamorphosis (Vengeance)
    # Druid
    22812,   # Barkskin
    61336,   # Survival Instincts
    102342,  # Ironbark
    # Evoker
    363916,  # Obsidian Scales
    374348,  # Renewing Blaze
    # Hunter
    186265,  # Aspect of the Turtle
    109304,  # Exhilaration
    264735,  # Survival of the Fittest
    # Mage
    45438,   # Ice Block
    55342,   # Mirror Image
    108978,  # Alter Time
    110959,  # Greater Invisibility
    # Monk
    115203,  # Fortifying Brew
    122783,  # Diffuse Magic
    131645,  # Zen Meditation
    122278,  # Dampen Harm
    # Paladin
    642,     # Divine Shield
    31850,   # Ardent Defender
    86659,   # Guardian of Ancient Kings
    184662,  # Shield of Vengeance
    # Priest
    19236,   # Desperate Prayer
    47585,   # Dispersion
    586,     # Fade
    15286,   # Vampiric Embrace
    # Rogue
    31224,   # Cloak of Shadows
    5277,    # Evasion
    1966,    # Feint
    1856,    # Vanish
    # Shaman
    108271,  # Astral Shift
    98008,   # Spirit Link Totem
    # Warlock
    104773,  # Unending Resolve
    108416,  # Dark Pact
    6789,    # Mortal Coil
    # Warrior
    871,     # Shield Wall
    118038,  # Die by the Sword
    184364,  # Enraged Regeneration
    97462,   # Rallying Cry
    23920,   # Spell Reflection
}

ALL_TRACKED_IDS = HEALTH_POT_IDS | COMBAT_POT_IDS | CLASS_DEFENSIVE_IDS

# ID → display name
SPELL_NAMES = {
    5512: "Healthstone", 432112: "Algari Healing Potion", 431924: "Algari Healing Potion",
    431932: "Tempered Potion", 432098: "Potion of Unwavering Focus",
    431945: "Light's Potential", 432106: "Void-Shrouded Tincture",
    48792: "Icebound Fortitude", 48707: "Anti-Magic Shell", 55233: "Vampiric Blood", 49028: "Dancing Rune Weapon",
    198589: "Blur", 196718: "Darkness", 196555: "Netherwalk", 187827: "Metamorphosis",
    22812: "Barkskin", 61336: "Survival Instincts", 102342: "Ironbark",
    363916: "Obsidian Scales", 374348: "Renewing Blaze",
    186265: "Aspect of the Turtle", 109304: "Exhilaration", 264735: "Survival of the Fittest",
    45438: "Ice Block", 55342: "Mirror Image", 108978: "Alter Time", 110959: "Greater Invisibility",
    115203: "Fortifying Brew", 122783: "Diffuse Magic", 131645: "Zen Meditation", 122278: "Dampen Harm",
    642: "Divine Shield", 31850: "Ardent Defender", 86659: "Guardian of Ancient Kings", 184662: "Shield of Vengeance",
    19236: "Desperate Prayer", 47585: "Dispersion", 586: "Fade", 15286: "Vampiric Embrace",
    31224: "Cloak of Shadows", 5277: "Evasion", 1966: "Feint", 1856: "Vanish",
    108271: "Astral Shift", 98008: "Spirit Link Totem",
    104773: "Unending Resolve", 108416: "Dark Pact", 6789: "Mortal Coil",
    871: "Shield Wall", 118038: "Die by the Sword", 184364: "Enraged Regeneration",
    97462: "Rallying Cry", 23920: "Spell Reflection",
}


def classify_spell(ability_game_id: int) -> str | None:
    """Classify a spell by its game ID into a category, or None if not tracked."""
    if ability_game_id in HEALTH_POT_IDS:
        return "Health"
    if ability_game_id in COMBAT_POT_IDS:
        return "Combat Pot"
    if ability_game_id in CLASS_DEFENSIVE_IDS:
        return "Defensive"
    return None


def analyze_fight_casts(cast_events: list, fight_start_time: float, actors: dict) -> dict:
    """Analyze cast events for tracked spells. Returns {sourceID: [{spell, category, timestamp}]}."""
    results = {}
    for event in cast_events:
        if event.get("type") != "cast":
            continue
        ability_id = event.get("abilityGameID")
        if ability_id is None:
            continue
        category = classify_spell(ability_id)
        if category is None:
            continue

        source_id = event.get("sourceID")
        if source_id is None:
            continue
        
        # Timestamp relative to fight start (in seconds)
        ts_ms = event.get("timestamp", 0)
        relative_sec = (ts_ms - fight_start_time) / 1000.0
        minutes = int(relative_sec // 60)
        seconds = int(relative_sec % 60)
        time_str = f"{minutes}:{seconds:02d}"
        
        if source_id not in results:
            results[source_id] = []
        results[source_id].append({
            "spell": SPELL_NAMES.get(ability_id, f"ID:{ability_id}"),
            "category": category,
            "time": time_str,
            "timestamp_ms": ts_ms,
        })
    
    return results


# ─── Gear Analysis ───────────────────────────────────────────────────────────

def detect_craft_quality(bonus_ids: list) -> int | None:
    """Return crafting quality rank (1-5) if item has a craft quality bonus ID."""
    for bid in bonus_ids:
        if bid in CRAFT_QUALITY_BONUS_IDS:
            return CRAFT_QUALITY_BONUS_IDS[bid]
    return None


def is_crafted(bonus_ids: list) -> bool:
    """Check if an item is crafted based on bonus IDs."""
    if any(bid in CRAFTED_INDICATOR_BONUS_IDS for bid in bonus_ids):
        return True
    if detect_craft_quality(bonus_ids) is not None:
        return True
    return False


def classify_ilvl_tier(ilvl: int, quality: int) -> str:
    """Classify an item's crafting tier based on ilvl."""
    if quality >= 4:  # Epic
        if ilvl >= 272:
            return "Myth (Spark + Myth Crests)"
        elif ilvl >= 259:
            return "Hero (Spark + Hero Crests)"
        elif ilvl >= 252:
            return "Epic Base (Spark, no crests)"
        else:
            return f"Epic (ilvl {ilvl})"
    else:  # Rare
        if ilvl >= 233:
            return "Rare (Veteran Crests)"
        elif ilvl >= 214:
            return "Rare (Adventurer Crests)"
        else:
            return "Rare Base (no spark)"


def estimate_spark_usage(ilvl: int, quality: int, slot: int) -> str:
    """Estimate spark usage based on ilvl and quality."""
    if quality < 4:  # Rare items don't use sparks
        return "No"
    if ilvl >= 252:
        if slot in (15, 16):  # Weapons
            return "Yes (2H = 4 sparks)" if ilvl >= 252 else "Yes (2 sparks)"
        return "Yes (2 sparks)"
    return "No"


def analyze_players(actors: list, combatant_events: list) -> list:
    """Analyze all players' gear using masterData actors + combatant events."""
    records = []
    
    # Build actor lookup from masterData: id -> actor info
    actor_lookup = {}
    for actor in actors:
        actor_lookup[actor["id"]] = {
            "name": actor.get("name", "Unknown"),
            "class": actor.get("subType", actor.get("type", "Unknown")),
            "server": actor.get("server", "Unknown"),
        }
    
    # Build gear lookup from combatant events: sourceID -> gear list
    gear_lookup = {}
    for event in combatant_events:
        source_id = event.get("sourceID")
        gear = event.get("gear", [])
        if source_id is not None and gear:
            # Keep the most complete gear entry per player
            if source_id not in gear_lookup or len(gear) > len(gear_lookup.get(source_id, [])):
                gear_lookup[source_id] = gear
    
    print(f"[DEBUG] Actors from masterData: {len(actor_lookup)}")
    print(f"[DEBUG] Players with gear from events: {len(gear_lookup)}")
    
    # For each player that has gear events, analyze their gear
    for source_id, gear in gear_lookup.items():
        actor = actor_lookup.get(source_id, {"name": f"Unknown-{source_id}", "class": "Unknown", "server": "Unknown"})
        player_has_crafted = False
        
        for idx, item in enumerate(gear):
            if not item or not isinstance(item, dict):
                continue
            
            item_id = item.get("id", 0)
            if item_id == 0:
                continue
            
            ilvl = item.get("itemLevel", 0)
            quality = item.get("quality", 0)
            bonus_ids = item.get("bonusIDs", []) or []
            slot = idx
            
            if is_crafted(bonus_ids):
                player_has_crafted = True
                craft_rank = detect_craft_quality(bonus_ids)
                tier = classify_ilvl_tier(ilvl, quality)
                spark = estimate_spark_usage(ilvl, quality, slot)
                
                records.append({
                    "player": actor["name"],
                    "class": actor["class"],
                    "server": actor["server"],
                    "slot": SLOT_NAMES.get(slot, f"Slot {slot}"),
                    "item_id": item_id,
                    "item_level": ilvl,
                    "quality": "Epic" if quality >= 4 else "Rare" if quality >= 3 else "Uncommon",
                    "craft_rank": f"Rank {craft_rank}" if craft_rank else "Unknown",
                    "tier": tier,
                    "spark_used": spark,
                    "bonus_ids": ", ".join(str(b) for b in bonus_ids),
                })
        
        if not player_has_crafted:
            records.append({
                "player": actor["name"],
                "class": actor["class"],
                "server": actor["server"],
                "slot": "—",
                "item_id": 0,
                "item_level": 0,
                "quality": "—",
                "craft_rank": "—",
                "tier": "NO CRAFTED GEAR FOUND",
                "spark_used": "—",
                "bonus_ids": "",
            })
    
    return records
    
    return records
    
    return records


# ─── HTML Output ─────────────────────────────────────────────────────────────

CLASS_COLORS = {
    "DeathKnight": "#C41E3A", "DemonHunter": "#A330C9", "Druid": "#FF7C0A",
    "Evoker": "#33937F", "Hunter": "#AAD372", "Mage": "#3FC7EB",
    "Monk": "#00FF98", "Paladin": "#F48CBA", "Priest": "#FFFFFF",
    "Rogue": "#FFF468", "Shaman": "#0070DD", "Warlock": "#8788EE",
    "Warrior": "#C69B6D",
}


def _build_split_html(split_data: dict, actors: list) -> dict:
    """Build HTML for each split. Returns {split_name: html_string}.
    Layout: rows = players, columns = bosses (3 sub-cols: Health | Combat | Defensive).
    """
    actor_lookup = {a["id"]: a for a in actors}
    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}

    def player_sort_key(pid):
        name = actor_lookup.get(pid, {}).get("name", "")
        pname, _ = lookup_roster(name)
        role = PLAYER_ROLES.get(pname, "DPS")
        return (ROLE_ORDER.get(role, 2), pname.lower())

    results = {}
    for split_name, fights in split_data.items():
        if not fights:
            continue

        all_pids = sorted(
            set(pid for fd in fights for pid in fd.get("player_casts", {})),
            key=player_sort_key
        )

        html = '<div class="table-wrap"><table><thead>'
        # Row 1: boss names (3 cols each)
        html += '<tr><th class="player-header" rowspan="2">Player</th>'
        for fd in fights:
            html += f'<th colspan="3" class="boss-name divider">{escape(fd.get("fight_name", "Boss"))}</th>'
        html += '</tr>'
        # Row 2: sub-column headers
        html += '<tr>'
        for _ in fights:
            html += '<th class="cast-h health-h divider">⚗ Health</th>'
            html += '<th class="cast-h combat-h">⚔ Combat</th>'
            html += '<th class="cast-h def-h">🛡 Defensive</th>'
        html += '</tr></thead><tbody>'

        current_role = None
        for pid in all_pids:
            actor = actor_lookup.get(pid, {})
            char_name = actor.get("name", f"ID-{pid}")
            cls = actor.get("subType", "Unknown")
            cls_color = CLASS_COLORS.get(cls, "#ccc")
            pname, _ = lookup_roster(char_name)
            role = PLAYER_ROLES.get(pname, "DPS")

            if role != current_role:
                current_role = role
                colspan = 1 + len(fights) * 3
                html += f'<tr class="role-sep"><td colspan="{colspan}">── {current_role}s ──</td></tr>'

            html += f'<tr>'
            html += (f'<td class="player-cell" style="color:{cls_color}">'
                     f'<span class="pname">{escape(pname)}</span><br>'
                     f'<span class="cname">{escape(char_name)}</span></td>')

            for fd in fights:
                casts = fd.get("player_casts", {}).get(pid, [])
                health    = [c for c in casts if c["category"] == "Health"]
                combat    = [c for c in casts if c["category"] == "Combat Pot"]
                defensive = [c for c in casts if c["category"] == "Defensive"]

                for items, cat_cls, is_first in [
                    (health,    "health-cell", True),
                    (combat,    "combat-cell", False),
                    (defensive, "def-cell",    False),
                ]:
                    div = " divider" if is_first else ""
                    if items:
                        lines = "".join(
                            f'<div class="cast-line">{escape(it["spell"])}'
                            f' <span class="cast-time">@ {it["time"]}</span></div>'
                            for it in items
                        )
                        html += f'<td class="{cat_cls}{div}">{lines}</td>'
                    else:
                        html += f'<td class="empty-cast{div}">—</td>'

            html += '</tr>'

        html += '</tbody></table></div>'
        results[split_name] = html

    return results


def write_html(records: list, report_info: dict, report_code: str, output_path: str,
               split_data: dict = None, actors: list = None):
    """Write the full audit report as a tabbed HTML file."""
    guild_name = "Unknown Guild"
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "Unknown Guild")
    report_title = report_info.get("title", report_code)
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.utcfromtimestamp(report_info["startTime"] / 1000).strftime("%Y-%m-%d")

    # ── Gear tab data ──
    players = {}
    for rec in records:
        p = rec["player"]
        if p not in players:
            players[p] = {"class": rec["class"], "items": []}
        if rec["item_id"] != 0:
            players[p]["items"].append(rec)

    max_items = max((len(p["items"]) for p in players.values()), default=2)
    max_items = max(max_items, 2)
    total_players = len(players)
    total_crafted = sum(len(p["items"]) for p in players.values())
    no_craft = sum(1 for p in players.values() if len(p["items"]) == 0)

    # Gear table rows
    gear_rows = ""
    for pname, pdata in sorted(players.items(), key=lambda x: (-len(x[1]["items"]), x[0].lower())):
        cls_color = CLASS_COLORS.get(pdata["class"], "#ccc")
        row_class = "no-craft" if not pdata["items"] else ""
        row = f'<tr class="{row_class}"><td class="player-cell" style="color:{cls_color}">{escape(pname)}</td>'
        for i in range(max_items):
            div = ' divider' if i > 0 else ''
            if i < len(pdata["items"]):
                item = pdata["items"][i]
                iid = item["item_id"]
                spark_cls = "spark-yes" if "Yes" in str(item["spark_used"]) else ""
                spark_txt = "Yes" if "Yes" in str(item["spark_used"]) else "No"
                item_link = f'<a href="https://www.wowhead.com/item={iid}" data-wowhead="item={iid}" target="_blank">#{iid}</a>'
                row += f'<td class="item-cell{div}">{item_link}</td>'
                row += f'<td>{escape(item["slot"])}</td>'
                row += f'<td class="center">{item["item_level"]}</td>'
                row += f'<td>{escape(item["craft_rank"])}</td>'
                row += f'<td class="center {spark_cls}">{spark_txt}</td>'
            else:
                row += f'<td class="empty{div}">—</td>' + '<td class="empty">—</td>' * 4
        row += '</tr>'
        gear_rows += row

    gear_header = '<th class="player-header">Player</th>'
    for i in range(max_items):
        div = ' divider' if i > 0 else ''
        gear_header += f'<th class="{div}">Item {i+1}</th><th>Slot</th><th>Ilvl</th><th>Rank</th><th>Spark?</th>'

    # ── Split tabs data (one tab per split) ──
    split_htmls = _build_split_html(split_data or {}, actors or [])

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Raid Audit — {escape(guild_name)}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #1a1a2e; color: #e0e0e0; font-family: -apple-system, 'Segoe UI', sans-serif; padding: 20px; }}
h1 {{ color: #7289DA; font-size: 22px; margin-bottom: 4px; }}
.subtitle {{ color: #888; font-size: 13px; margin-bottom: 16px; }}
.subtitle a {{ color: #7289DA; text-decoration: none; }}

/* ── Tabs ── */
.tab-bar {{ display: flex; gap: 4px; margin-bottom: 20px; border-bottom: 2px solid #2a2a4a; padding-bottom: 0; }}
.tab-btn {{ background: #16213e; color: #888; border: none; padding: 10px 24px; border-radius: 8px 8px 0 0; font-size: 14px; font-weight: 600; cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -2px; transition: all 0.15s; }}
.tab-btn:hover {{ color: #bbb; background: #1e2d50; }}
.tab-btn.active {{ color: #7289DA; background: #1a1a2e; border-bottom: 2px solid #7289DA; }}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}

/* ── Stats ── */
.stats {{ display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }}
.stat-box {{ background: #16213e; border-radius: 8px; padding: 10px 18px; }}
.stat-box .num {{ font-size: 24px; font-weight: bold; color: #7289DA; }}
.stat-box .label {{ font-size: 11px; color: #888; text-transform: uppercase; }}
.stat-box.warn .num {{ color: #ffc107; }}

/* ── Search ── */
.search-box {{ margin: 0 0 12px; }}
.search-box input {{ background: #16213e; border: 1px solid #2a2a4a; color: #e0e0e0; padding: 7px 14px; border-radius: 6px; width: 280px; font-size: 13px; }}
.search-box input::placeholder {{ color: #555; }}

/* ── Tables (shared) ── */
.table-wrap {{ overflow-x: auto; border-radius: 8px; margin-bottom: 28px; }}
table {{ border-collapse: collapse; min-width: 100%; }}
th {{ background: #16213e; color: #7289DA; padding: 8px 10px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; position: sticky; top: 0; z-index: 1; }}
th.player-header {{ position: sticky; left: 0; z-index: 2; background: #16213e; min-width: 140px; }}
td {{ padding: 6px 10px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; white-space: nowrap; }}
.player-cell {{ font-weight: bold; position: sticky; left: 0; background: #0d1f3c; z-index: 1; min-width: 140px; }}
tr:hover td {{ background: rgba(114,137,218,0.07); }}
tr:hover .player-cell {{ background: #152848; }}
td.divider, th.divider {{ border-left: 2px solid rgba(114,137,218,0.25); }}
.center {{ text-align: center; }}
.empty {{ color: #2a2a4a; text-align: center; }}

/* ── Gear tab ── */
.no-craft {{ background: rgba(255,160,0,0.05); }}
.no-craft .player-cell {{ background: rgba(80,50,0,0.4); }}
.spark-yes {{ color: #4caf50; font-weight: bold; }}
.item-cell a {{ color: #a48cff; text-decoration: none; }}
.item-cell a:hover {{ text-decoration: underline; }}

/* ── Split tab ── */
.split-section {{ margin-bottom: 40px; }}
.split-title {{ font-size: 17px; color: #7289DA; margin-bottom: 14px; padding-bottom: 6px; border-bottom: 1px solid #2a2a4a; }}
.boss-name {{ text-align: center; color: #c4b0ff; font-size: 12px; background: #111827; padding: 6px 10px; }}
.cast-h {{ font-size: 10px; padding: 5px 8px; }}
.health-h {{ color: #e57373; }}
.combat-h {{ color: #ce93d8; }}
.def-h {{ color: #64b5f6; }}
.cast-line {{ margin: 2px 0; white-space: nowrap; }}
.cast-time {{ color: #888; font-size: 11px; }}
.health-cell {{ background: rgba(229,115,115,0.06); }}
.combat-cell {{ background: rgba(206,147,216,0.06); }}
.def-cell {{ background: rgba(100,181,246,0.06); white-space: normal; min-width: 160px; }}
.empty-cast {{ color: #2a2a4a; text-align: center; }}
.role-sep td {{ background: #111827; color: #7289DA; font-size: 11px; font-weight: bold; padding: 4px 10px; }}
.pname {{ font-weight: bold; }}
.cname {{ color: #888; font-size: 11px; }}
</style>
<script>const whTooltips = {{colorLinks: true, iconizeLinks: true, iconSize: 'small'}};</script>
<script src="https://wow.zamimg.com/js/tooltips.js"></script>
</head>
<body>

<h1>Raid Audit — {escape(guild_name)}</h1>
<div class="subtitle">{escape(report_title)} · {report_date} · <a href="https://www.warcraftlogs.com/reports/{report_code}" target="_blank">View on WarcraftLogs ↗</a></div>

<div class="tab-bar">
  <button class="tab-btn active" onclick="switchTab('gear', this)">⚙ Gear Audit</button>
  {''.join(f'<button class="tab-btn" onclick="switchTab(\'split-{i}\', this)">📋 {escape(name)}</button>' for i, name in enumerate(split_htmls))}
</div>

<!-- ── GEAR TAB ── -->
<div id="tab-gear" class="tab-content active">
  <div class="stats">
    <div class="stat-box"><div class="num">{total_players}</div><div class="label">Players</div></div>
    <div class="stat-box"><div class="num">{total_crafted}</div><div class="label">Crafted Items</div></div>
    <div class="stat-box warn"><div class="num">{no_craft}</div><div class="label">No Crafted Gear</div></div>
  </div>
  <div class="search-box"><input type="text" placeholder="Search player..." onkeyup="filterGear(this)"></div>
  <div class="table-wrap">
    <table id="gear-table">
      <thead><tr>{gear_header}</tr></thead>
      <tbody>{gear_rows}</tbody>
    </table>
  </div>
</div>

<!-- ── SPLIT TABS ── -->
{''.join(f'<div id="tab-split-{i}" class="tab-content">{content}</div>' for i, content in enumerate(split_htmls.values()))}

<script>
function switchTab(name, btn) {{
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  btn.classList.add('active');
}}
function filterGear(input) {{
  const filter = input.value.toLowerCase();
  for (let row of document.getElementById('gear-table').tBodies[0].rows) {{
    row.style.display = row.cells[0].textContent.toLowerCase().includes(filter) ? '' : 'none';
  }}
}}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n[OK] HTML report saved to: {output_path}")


# ─── Roster Mapping (The Rupture) ────────────────────────────────────────────

# Maps character name (lowercase) -> (Player name, "Main" or "Alt")
ROSTER = {}
# (Player, [(char_name, Main/Alt)])
# Role is determined per-player, not per-char
_roster_raw = [
    ("Nope",        "Tank",    [("NopeDK","Alt"), ("Nopebrew","Main")]),
    ("Phyxius",     "Tank",    [("Phyxius","Main"), ("Phyxy","Alt")]),
    ("Toshiko",     "Healer",  [("Toshiko","Alt"), ("Evokenooblal","Alt")]),
    ("Minxy",       "Healer",  [("Minxymender","Alt"), ("Minxycat","Main")]),
    ("Hipe",        "Healer",  [("Cype","Alt"), ("Wype","Main")]),
    ("Zush",        "Healer",  [("Zush","Main"), ("Züsh","Main"), ("Ryukho","Alt")]),
    ("Jimy",        "DPS",     [("jaime","Main")]),
    ("Ice",         "DPS",     [("Icecoldleap","Main"), ("Chocoice","Alt")]),
    ("Zodiacos",    "DPS",     [("Zodiacos","Alt")]),
    ("Kutcher",     "DPS",     [("Kutcherdhtwo","Alt"), ("Kutchersplit","Alt")]),
    ("Hypno",       "DPS",     [("Hypno","Main"), ("Hypnodh","Alt")]),
    ("Brunaine",    "DPS",     [("Brunainevoke","Alt"), ("Brunainehunt","Alt")]),
    ("Beldryk",     "DPS",     [("beldrýk","Main"), ("Beldrýk","Main"), ("beldryc","Alt")]),
    ("Potrenu",     "DPS",     [("Potrenu","Main"), ("Potrenuu","Alt")]),
    ("Shamishan",   "DPS",     [("Samdracson","Alt"), ("Shamishan","Main")]),
    ("Madonis",     "DPS",     [("Madonis","Alt"), ("Madonisvoker","Alt")]),
    ("Kaze",        "DPS",     [("Kazeofscales","Main"), ("Käz","Main"), ("Kazeoflight","Alt")]),
    ("Upyeah",      "DPS",     [("upyeah","Alt"), ("Upyeah","Alt"), ("upyeäh","Alt")]),
    ("Mindhacker",  "DPS",     [("Mindrage","Alt"), ("Mindhacker","Main")]),
    ("Uncleyoinky", "DPS",     [("allblues","Alt"), ("Allblues","Alt"), ("uncleyoinky","Main")]),
    ("Nizze",       "DPS",     [("Nizzedk","Alt"), ("Nizze","Main")]),
    ("Mostbanned",  "DPS",     [("Mosta","Alt"), ("Mostbanned","Alt")]),
    ("Malheiro",    "DPS",     [("Rödinhas","Alt"), ("Malheiro","Main")]),
    ("Bolters",     "DPS",     [("Schmosba","Main"), ("Ipala","Alt")]),
    ("Tinet",       "DPS",     [("Pingveryhigh","Alt"), ("Guldanrämsay","Alt")]),
    ("Doomkry",     "DPS",     [("Doomkry","Main"), ("Lockry","Alt")]),
]
PLAYER_ROLES = {}  # player_name -> "Tank"/"Healer"/"DPS"
for player_name, role, chars in _roster_raw:
    PLAYER_ROLES[player_name] = role
    for char_name, main_alt in chars:
        ROSTER[char_name.lower()] = (player_name, main_alt)


def lookup_roster(char_name: str):
    """Look up a character in the roster. Returns (player_name, 'Main'/'Alt') or (char_name, 'Unknown')."""
    return ROSTER.get(char_name.lower(), (char_name, "Unknown"))


# ─── XLSX Output ─────────────────────────────────────────────────────────────

# Class background colors (lighter/muted versions for row backgrounds)
XLSX_CLASS_BG = {
    "DeathKnight": "77C41E3A", "DemonHunter": "77A330C9", "Druid": "77FF7C0A",
    "Evoker": "7733937F", "Hunter": "77AAD372", "Mage": "773FC7EB",
    "Monk": "7700FF98", "Paladin": "77F48CBA", "Priest": "77AAAAAA",
    "Rogue": "77FFF468", "Shaman": "770070DD", "Warlock": "778788EE",
    "Warrior": "77C69B6D",
}


def write_xlsx(records: list, report_info: dict, report_code: str, output_path: str, split_data: dict = None):
    """Write crafted gear data to XLSX with Mains, Alts, and Split tabs."""
    wb = Workbook()
    
    guild_name = "Unknown Guild"
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "Unknown Guild")
    report_title = report_info.get("title", report_code)
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.utcfromtimestamp(report_info["startTime"] / 1000).strftime("%Y-%m-%d")

    # Group records by player and enrich with roster data
    players = {}
    for rec in records:
        char_name = rec["player"]
        player_name, role = lookup_roster(char_name)
        key = (player_name, char_name)
        if key not in players:
            players[key] = {"player": player_name, "char": char_name, "class": rec["class"], "role": role, "items": []}
        if rec["item_id"] != 0:
            players[key]["items"].append(rec)

    # Split into mains and alts
    mains = {k: v for k, v in players.items() if v["role"] == "Main"}
    alts = {k: v for k, v in players.items() if v["role"] != "Main"}

    # Find max items for column count
    all_items_counts = [len(v["items"]) for v in players.values()]
    max_items = max(all_items_counts) if all_items_counts else 2
    max_items = max(max_items, 2)

    # Role sort order
    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}

    def sort_key(k):
        pdata = players.get(k) or mains.get(k) or alts.get(k)
        player_name = pdata["player"]
        role_rank = ROLE_ORDER.get(PLAYER_ROLES.get(player_name, "DPS"), 2)
        return (role_rank, -len(pdata["items"]), player_name.lower())

    # Styling
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1a1a2e")
    black_font = Font(name="Arial", size=10, color="000000")
    black_bold = Font(name="Arial", size=10, color="000000", bold=True)
    border = Border(bottom=Side(style="thin", color="333333"))
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    spark_font = Font(name="Arial", size=10, color="006600", bold=True)
    no_spark_font = Font(name="Arial", size=10, color="000000")
    no_craft_fill = PatternFill("solid", fgColor="3D3D3D")
    no_craft_font = Font(name="Arial", size=10, color="999999")

    def get_class_fill(cls_name):
        color = XLSX_CLASS_BG.get(cls_name, "77666666")
        return PatternFill("solid", fgColor=color)

    def build_sheet(ws, title, player_data):
        # Title
        total_cols = 2 + max_items * 3
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        ws["A1"].value = f"{title} — {guild_name} — {report_title} ({report_date})"
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color="7289DA")
        ws.row_dimensions[1].height = 26

        ws["A2"].value = f"Report: https://www.warcraftlogs.com/reports/{report_code}"
        ws["A2"].font = Font(name="Arial", size=9, color="888888", italic=True)

        # Headers: Player | Char | Slot 1 | Ilvl 1 | Spark? 1 | Slot 2 | Ilvl 2 | Spark? 2 ...
        hr = 4
        headers = ["Player", "Char"]
        for i in range(max_items):
            headers += [f"Slot {i+1}", f"Ilvl", f"Spark?"]
        
        col_widths = [16, 18] + [12, 8, 8] * max_items
        for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=hr, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[get_column_letter(ci)].width = w
            # Divider before each item group
            if ci > 2 and (ci - 3) % 3 == 0 and (ci - 3) // 3 > 0:
                cell.border = Border(bottom=Side(style="thin", color="333333"), left=Side(style="medium", color="7289DA"))
        ws.row_dimensions[hr].height = 20
        ws.freeze_panes = "A5"

        # Data rows — sorted by role then items
        row = hr + 1
        current_role = None
        for key in sorted(player_data.keys(), key=sort_key):
            pdata = player_data[key]
            cls = pdata["class"]
            cls_fill = get_class_fill(cls)
            has_items = len(pdata["items"]) > 0
            player_name = pdata["player"]
            player_role = PLAYER_ROLES.get(player_name, "DPS")

            # Role separator
            if player_role != current_role:
                current_role = player_role
                sep_cell = ws.cell(row=row, column=1, value=f"── {current_role}s ──")
                sep_cell.font = Font(name="Arial", size=9, bold=True, color="7289DA")
                for ci in range(1, total_cols + 1):
                    ws.cell(row=row, column=ci).fill = PatternFill("solid", fgColor="111122")
                    ws.cell(row=row, column=ci).border = border
                row += 1

            if not has_items:
                fill = no_craft_fill
                name_font = no_craft_font
                txt_font = no_craft_font
            else:
                fill = cls_fill
                name_font = black_bold
                txt_font = black_font

            # Player name
            c = ws.cell(row=row, column=1, value=player_name)
            c.font = name_font
            c.fill = fill
            c.border = border
            c.alignment = left

            # Char name
            c = ws.cell(row=row, column=2, value=pdata["char"])
            c.font = txt_font
            c.fill = fill
            c.border = border
            c.alignment = left

            # Items: Slot | Ilvl | Spark?
            for i in range(max_items):
                base_col = 3 + i * 3
                is_divider = i > 0

                if i < len(pdata["items"]):
                    item = pdata["items"][i]
                    spark_yes = "Yes" in str(item["spark_used"])
                    
                    vals = [item["slot"], item["item_level"], "Yes" if spark_yes else "No"]
                    for j, val in enumerate(vals):
                        cell = ws.cell(row=row, column=base_col + j, value=val)
                        cell.fill = fill
                        cell.border = border
                        if j == 1:
                            cell.alignment = center
                            cell.font = txt_font
                        elif j == 2:
                            cell.alignment = center
                            cell.font = spark_font if spark_yes else txt_font
                        else:
                            cell.font = txt_font
                        # Divider
                        if j == 0 and is_divider:
                            cell.border = Border(bottom=Side(style="thin", color="333333"), left=Side(style="medium", color="7289DA"))
                else:
                    for j in range(3):
                        cell = ws.cell(row=row, column=base_col + j, value="—")
                        cell.font = no_craft_font if not has_items else Font(name="Arial", size=10, color="444444")
                        cell.fill = fill
                        cell.border = border
                        cell.alignment = center
                        if j == 0 and is_divider:
                            cell.border = Border(bottom=Side(style="thin", color="333333"), left=Side(style="medium", color="7289DA"))

            row += 1

        # Auto-filter
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(total_cols)}{row - 1}"

    # Build crafted gear sheets
    ws_mains = wb.active
    ws_mains.title = "Mains"
    build_sheet(ws_mains, "Mains — Crafted Gear", mains)

    ws_alts = wb.create_sheet("Alts")
    build_sheet(ws_alts, "Alts — Crafted Gear", alts)

    # Build split consumable/defensive tabs
    if split_data:
        for split_name, split_fights in split_data.items():
            ws_split = wb.create_sheet(split_name)
            _build_split_sheet(ws_split, split_name, split_fights, report_info, actors_list=None)

    wb.save(output_path)
    print(f"[OK] XLSX report saved to: {output_path}")


def _build_split_sheet(ws, title, fights_data, report_info, actors_list):
    """Build a split tab showing consumable/defensive usage per player per fight.
    
    fights_data: list of { "fight_name": str, "difficulty": str, "player_casts": { sourceID: [cast_info] } }
    """
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1a1a2e")
    data_font = Font(name="Arial", size=10, color="000000")
    border = Border(bottom=Side(style="thin", color="333333"))
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    cat_fills = {
        "Health": PatternFill("solid", fgColor="77CC4444"),
        "Combat Pot": PatternFill("solid", fgColor="77AA44CC"),
        "Defensive": PatternFill("solid", fgColor="774488CC"),
    }

    ws["A1"].value = title
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="7289DA")
    ws.row_dimensions[1].height = 26

    # Collect all unique players across all fights in this split
    all_player_ids = set()
    for fd in fights_data:
        all_player_ids.update(fd.get("player_casts", {}).keys())

    # Get actor lookup
    actors = {}
    if report_info.get("masterData") and report_info["masterData"].get("actors"):
        for a in report_info["masterData"]["actors"]:
            actors[a["id"]] = a

    # Sort players by roster role
    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}
    def player_sort(pid):
        actor = actors.get(pid, {})
        name = actor.get("name", "")
        player_name, _ = lookup_roster(name)
        role = PLAYER_ROLES.get(player_name, "DPS")
        return (ROLE_ORDER.get(role, 2), player_name.lower())

    sorted_players = sorted(all_player_ids, key=player_sort)

    # Two-row header: row 3 = boss names (merged 3 cols), row 4 = sub-columns
    hr_boss = 3
    hr_sub  = 4

    # "Player" cell spans both header rows
    ws.merge_cells(start_row=hr_boss, start_column=1, end_row=hr_sub, end_column=1)
    c = ws.cell(row=hr_boss, column=1, value="Player")
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border
    ws.column_dimensions["A"].width = 18

    sub_labels = ["Health Pots", "Combat Pots", "Defensives"]
    col_widths  = [20, 20, 28]

    for fi, fd in enumerate(fights_data):
        fname   = fd.get("fight_name", "Boss")
        base_ci = 2 + fi * 3  # first column of this boss group

        # Merge boss name across 3 columns
        ws.merge_cells(start_row=hr_boss, start_column=base_ci, end_row=hr_boss, end_column=base_ci + 2)
        bc = ws.cell(row=hr_boss, column=base_ci, value=fname)
        bc.font = header_font
        bc.fill = PatternFill("solid", fgColor="111827")
        bc.alignment = Alignment(horizontal="center", vertical="center")
        bc.border = Border(bottom=Side(style="thin", color="333333"),
                           left=Side(style="medium", color="7289DA"))

        # Sub-column headers
        for j, (label, w) in enumerate(zip(sub_labels, col_widths)):
            ci = base_ci + j
            sc = ws.cell(row=hr_sub, column=ci, value=label)
            sc.font = header_font
            sc.fill = header_fill
            sc.alignment = Alignment(horizontal="center", vertical="center")
            sc.border = border
            if j == 0:
                sc.border = Border(bottom=Side(style="thin", color="333333"),
                                   left=Side(style="medium", color="7289DA"))
            ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[hr_boss].height = 20
    ws.row_dimensions[hr_sub].height  = 18
    ws.freeze_panes = "A5"

    # Data rows
    row = hr_sub + 1
    for pid in sorted_players:
        actor = actors.get(pid, {})
        char_name = actor.get("name", f"ID-{pid}")
        cls = actor.get("subType", "Unknown")
        cls_bg = XLSX_CLASS_BG.get(cls, "77666666")
        cls_fill = PatternFill("solid", fgColor=cls_bg)
        player_name, _ = lookup_roster(char_name)

        c = ws.cell(row=row, column=1, value=f"{player_name}\n({char_name})")
        c.font = Font(name="Arial", size=10, color="000000", bold=True)
        c.fill = cls_fill
        c.border = border
        c.alignment = Alignment(vertical="center", wrap_text=True)

        for fi, fd in enumerate(fights_data):
            base_col = 2 + fi * 3
            casts = fd.get("player_casts", {}).get(pid, [])
            
            # Group by category
            health = [c for c in casts if c["category"] == "Health"]
            combat = [c for c in casts if c["category"] == "Combat Pot"]
            defensives = [c for c in casts if c["category"] == "Defensive"]

            for j, (items, cat) in enumerate([(health, "Health"), (combat, "Combat Pot"), (defensives, "Defensive")]):
                cell = ws.cell(row=row, column=base_col + j)
                cell.fill = cls_fill
                cell.border = border
                
                if items:
                    lines = [f"{it['spell']} @ {it['time']}" for it in items]
                    cell.value = "\n".join(lines)
                    cell.font = data_font
                    cell.alignment = wrap
                else:
                    cell.value = "—"
                    cell.font = Font(name="Arial", size=10, color="666666")
                    cell.alignment = center

                # Divider
                if j == 0:
                    cell.border = Border(bottom=Side(style="thin", color="333333"), left=Side(style="medium", color="7289DA"))

        # Height = tallest cell: max casts in any single category across all fights
        max_lines = 1
        for fd in fights_data:
            player_casts = fd.get("player_casts", {}).get(pid, [])
            for cat in ("Health", "Combat Pot", "Defensive"):
                count = sum(1 for c in player_casts if c["category"] == cat)
                max_lines = max(max_lines, count)
        ws.row_dimensions[row].height = max(18, 15 * max_lines)
        row += 1

def load_config(path="wcl_config.txt"):
    """Load credentials from a config file if it exists."""
    config = {}
    try:
        with open(path, "r") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    key, val = line.split("=", 1)
                    config[key.strip()] = val.strip()
    except FileNotFoundError:
        pass
    return config


def main():
    print("=" * 60)
    print("  WarcraftLogs Crafted Gear Audit — Midnight Season 1")
    print("=" * 60)
    
    # Try loading from config file
    config = load_config()
    
    if config.get("CLIENT_ID") and config.get("CLIENT_SECRET"):
        client_id = config["CLIENT_ID"]
        client_secret = config["CLIENT_SECRET"]
        print(f"\n[OK] Loaded credentials from wcl_config.txt (Client ID: {client_id[:8]}...)")
    else:
        client_id = input("\nWarcraftLogs Client ID: ").strip()
        client_secret = input("WarcraftLogs Client Secret: ").strip()
    
    if not client_id or not client_secret:
        print("[ERROR] Both Client ID and Client Secret are required.")
        sys.exit(1)
    
    # Authenticate
    print("\n[...] Authenticating with WarcraftLogs...")
    token = get_access_token(client_id, client_secret)
    print("[OK] Authenticated successfully.")
    
    # Report code — check config first
    if config.get("REPORT_URL"):
        report_input = config["REPORT_URL"]
        print(f"[OK] Loaded report URL from config: {report_input}")
    else:
        report_input = input("\nReport code or URL: ").strip()
    # Extract code from URL if full URL was pasted
    if "warcraftlogs.com" in report_input:
        # URL format: https://www.warcraftlogs.com/reports/XXXXXXXXXXXX
        parts = report_input.rstrip("/").split("/")
        report_code = parts[-1].split("#")[0].split("?")[0]
    else:
        report_code = report_input
    
    print(f"\n[...] Fetching report info for: {report_code}")
    report_info = fetch_report_info(token, report_code)
    
    guild_name = ""
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "")
    print(f"[OK] Report: {report_info.get('title', 'Untitled')} ({guild_name})")
    
    print("\n[...] Fetching players from ALL fights...")
    
    # Get all fights, filter to boss encounters (encounterID > 0) with kills
    all_raw_fights = report_info.get("fights", [])
    fights = [f for f in all_raw_fights if f.get("encounterID", 0) > 0 and f.get("kill")]
    all_fights = fights if fights else []
    selected_fights = all_fights  # Default to all
    
    if all_fights:
        print(f"\nBoss kills found ({len(all_fights)}):")
        for f in all_fights:
            diff = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(f.get("difficulty", 0), f"D{f.get('difficulty', '?')}")
            print(f"  [{f['id']:>3}] {f['name']} — {diff} (Kill)")
        
        fight_input = input("\nFight IDs to analyze (comma-separated, or 'all'): ").strip()
        if fight_input.lower() != "all" and fight_input != "":
            selected_ids = [int(x.strip()) for x in fight_input.split(",") if x.strip().isdigit()]
            selected_fights = [f for f in all_fights if f["id"] in selected_ids]
    
    # Get actors from masterData (player names, classes, servers)
    actors = []
    if report_info.get("masterData") and report_info["masterData"].get("actors"):
        actors = report_info["masterData"]["actors"]
    print(f"[OK] Found {len(actors)} player actors in report.")
    
    # Fetch gear via CombatantInfo events for each fight
    all_combatant_events = []
    for fight in selected_fights:
        fid = fight["id"]
        fname = fight["name"]
        print(f"  [...] Fetching gear from fight {fid}: {fname}...")
        try:
            events = fetch_combatant_info_events(token, report_code, fid)
            all_combatant_events.extend(events)
            print(f"         Got {len(events)} players' gear.")
        except Exception as e:
            print(f"    [WARN] Could not fetch gear events for fight {fid}: {e}")
    
    # Debug: dump raw API responses
    debug_path = "wcl_debug_response.json"
    with open(debug_path, "w", encoding="utf-8") as f:
        json.dump({
            "actors": actors,
            "combatantEvents": all_combatant_events[:3],
            "totalCombatantEvents": len(all_combatant_events),
        }, f, indent=2, ensure_ascii=False)
    print(f"[DEBUG] Raw API responses saved to: {debug_path}")
    
    print("\n[...] Analyzing crafted gear...")
    records = analyze_players(actors, all_combatant_events)
    
    crafted_count = sum(1 for r in records if r["item_id"] != 0)
    player_count = len(set(r["player"] for r in records))
    no_craft_count = sum(1 for r in records if r["item_id"] == 0)
    
    print(f"[OK] Found {crafted_count} crafted items across {player_count} players.")
    if no_craft_count > 0:
        print(f"[!!] {no_craft_count} player(s) have NO detected crafted gear.")
    
    # ── Split detection & cast tracking ──
    # Group kills by encounter — first kill = Split 1, second = Split 2
    encounter_kills = {}
    for fight in selected_fights:
        eid = fight.get("encounterID", 0)
        if eid not in encounter_kills:
            encounter_kills[eid] = []
        encounter_kills[eid].append(fight)
    
    split1_fights = []
    split2_fights = []
    for eid, kills in encounter_kills.items():
        kills.sort(key=lambda f: f["id"])
        if len(kills) >= 1:
            split1_fights.append(kills[0])
        if len(kills) >= 2:
            split2_fights.append(kills[1])
    
    print(f"\n[...] Detected {len(split1_fights)} Split 1 fights, {len(split2_fights)} Split 2 fights.")
    
    # Fetch cast events for each split
    split_data = {}
    actor_lookup = {a["id"]: a for a in actors}
    
    for split_name, split_fights_list in [("Split 1", split1_fights), ("Split 2", split2_fights)]:
        if not split_fights_list:
            continue
        
        print(f"\n[...] Fetching spell usage for {split_name}...")
        split_fight_data = []
        
        for fight in split_fights_list:
            fid = fight["id"]
            fname = fight["name"]
            diff = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(fight.get("difficulty", 0), "")
            print(f"  [...] Fetching casts for {fname} ({diff})...")
            
            try:
                cast_events = fetch_cast_events(token, report_code, fid)
                # Event timestamps are relative to report start, same as fight.startTime
                fight_start = fight.get("startTime", 0)

                player_casts = analyze_fight_casts(cast_events, fight_start, actor_lookup)

                total_tracked = sum(len(v) for v in player_casts.values())
                print(f"         Found {total_tracked} tracked spell uses across {len(player_casts)} players.")
                
                split_fight_data.append({
                    "fight_name": f"{fname} ({diff})",
                    "fight_id": fid,
                    "player_casts": player_casts,
                })
            except Exception as e:
                print(f"    [WARN] Could not fetch casts for fight {fid}: {e}")
        
        split_data[split_name] = split_fight_data
    
    # Output
    html_path = "craft_audit.html"
    write_html(records, report_info, report_code, html_path, split_data=split_data, actors=actors)

    xlsx_path = "craft_audit.xlsx"
    write_xlsx(records, report_info, report_code, xlsx_path, split_data=split_data)
    
    print(f"\nDone! Refresh craft_audit.html in your browser, or open craft_audit.xlsx.\n")


if __name__ == "__main__":
    main()