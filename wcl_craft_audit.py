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
from datetime import datetime, timezone
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
                    abilities {
                        gameID
                        name
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


def fetch_death_events(token: str, report_code: str, fight_id: int) -> list:
    """Fetch death events for a fight (friendly players only)."""
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                events(dataType: Deaths, fightIDs: [$fightID], hostilityType: Friendlies, limit: 100) {
                    data
                }
            }
        }
    }
    """
    result = query_wcl(token, query, {"code": report_code, "fightID": fight_id})
    return result["reportData"]["report"]["events"].get("data", [])


def fetch_interrupt_events(token: str, report_code: str, fight_id: int) -> list:
    """Fetch interrupt events for a fight (friendly players), with pagination."""
    query = """
    query ($code: String!, $fightID: Int!, $startTime: Float) {
        reportData {
            report(code: $code) {
                events(dataType: Interrupts, fightIDs: [$fightID], hostilityType: Friendlies,
                       limit: 10000, startTime: $startTime) {
                    data
                    nextPageTimestamp
                }
            }
        }
    }
    """
    all_events = []
    variables = {"code": report_code, "fightID": fight_id, "startTime": None}
    while True:
        result = query_wcl(token, query, variables)
        page = result["reportData"]["report"]["events"]
        all_events.extend(page.get("data", []))
        next_ts = page.get("nextPageTimestamp")
        if not next_ts:
            break
        variables = {"code": report_code, "fightID": fight_id, "startTime": next_ts}
    return all_events


def analyze_interrupts(interrupt_events: list, actor_lookup: dict) -> dict:
    """Count interrupts performed by each player.
    Returns {pid: int}
    """
    totals = {}
    for event in interrupt_events:
        pid = event.get("sourceID")
        if pid is None or pid not in actor_lookup:
            continue
        totals[pid] = totals.get(pid, 0) + 1
    return totals


def analyze_boss_mechanics(damage_events: list, actor_lookup: dict, mechanics: list) -> dict:
    """Count per-player hits for each boss mechanic based on spell IDs.
    Returns {pid: {mechanic_label: count}}
    """
    if not mechanics:
        return {}
    spell_to_label = {}
    for mech in mechanics:
        for sid in mech["spell_ids"]:
            spell_to_label[sid] = mech["label"]
    result = {}
    for event in damage_events:
        if event.get("type") != "damage":
            continue
        pid = event.get("targetID")
        if pid not in actor_lookup:
            continue
        aid = event.get("abilityGameID", 0)
        label = spell_to_label.get(aid)
        if label is None:
            continue
        if pid not in result:
            result[pid] = {}
        result[pid][label] = result[pid].get(label, 0) + 1
    return result


def fetch_damage_taken_events(token: str, report_code: str, fight_id: int) -> list:
    """Fetch avoidable damage-taken events for a fight (friendly players only), with pagination."""
    query = """
    query ($code: String!, $fightID: Int!, $startTime: Float) {
        reportData {
            report(code: $code) {
                events(dataType: DamageTaken, fightIDs: [$fightID], hostilityType: Friendlies,
                       limit: 10000, startTime: $startTime) {
                    data
                    nextPageTimestamp
                }
            }
        }
    }
    """
    all_events = []
    variables = {"code": report_code, "fightID": fight_id, "startTime": None}
    while True:
        result = query_wcl(token, query, variables)
        page = result["reportData"]["report"]["events"]
        all_events.extend(page.get("data", []))
        next_ts = page.get("nextPageTimestamp")
        if not next_ts:
            break
        variables = {"code": report_code, "fightID": fight_id, "startTime": next_ts}
    return all_events


def fetch_uptime_table(token: str, report_code: str, fight_id: int) -> dict:
    """Fetch DPS active time + damage done per player from WCL table endpoint.
    Returns {sourceID: {"activeTime": ms, "total": damage_int}}.
    """
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                table(dataType: DamageDone, fightIDs: [$fightID])
            }
        }
    }
    """
    result = query_wcl(token, query, {"code": report_code, "fightID": fight_id})
    table = result["reportData"]["report"]["table"]
    entries = table.get("entries", []) if isinstance(table, dict) else []
    return {e["id"]: {"activeTime": e.get("activeTime", 0), "total": e.get("total", 0)}
            for e in entries if "id" in e}


def fetch_healing_table(token: str, report_code: str, fight_id: int) -> dict:
    """Fetch healing done per player from WCL table endpoint.
    Returns {sourceID: total_healing_int}.
    """
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                table(dataType: Healing, fightIDs: [$fightID])
            }
        }
    }
    """
    result = query_wcl(token, query, {"code": report_code, "fightID": fight_id})
    table = result["reportData"]["report"]["table"]
    entries = table.get("entries", []) if isinstance(table, dict) else []
    return {e["id"]: e.get("total", 0) for e in entries if "id" in e}


def fetch_rankings(token: str, report_code: str, fight_id: int) -> dict:
    """Fetch player parse percentiles for a fight.
    Returns {player_name_lower: rankPercent}.
    """
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                rankings(fightIDs: [$fightID])
            }
        }
    }
    """
    result = query_wcl(token, query, {"code": report_code, "fightID": fight_id})
    rankings = result["reportData"]["report"]["rankings"]
    if not isinstance(rankings, dict):
        return {}
    out = {}
    for entry in rankings.get("data", []):
        name = entry.get("name", "").lower()
        pct = entry.get("rankPercent", 0)
        if name:
            out[name] = round(pct)
    return out


def aggregate_damage_taken(damage_events: list, actor_lookup: dict) -> dict:
    """Sum total actual damage taken per player (post-mitigation, all sources).
    Returns {pid: total_damage_int}
    """
    totals = {}
    for event in damage_events:
        if event.get("type") != "damage":
            continue
        pid = event.get("targetID")
        if pid is None or pid not in actor_lookup:
            continue
        totals[pid] = totals.get(pid, 0) + event.get("amount", 0)
    return totals


def analyze_avoidable_damage(damage_events: list, actor_lookup: dict,
                             fight_start_ms: int = 0, ability_names: dict = None,
                             player_max_hp: dict = None) -> dict:
    """Count enemy-source damage-taken hits per player (proxy for avoidable damage).
    Returns {pid: {"hits": int, "big_hits": int, "details": [{"ability": str, "amount_k": str, "time": str}]}}
    where big_hits = hits > 10% of player's estimated max HP (stamina * 20).
    """
    ability_names = ability_names or {}
    player_max_hp = player_max_hp or {}
    friendly_ids = set(actor_lookup.keys())
    results = {}
    for event in sorted(damage_events, key=lambda e: e.get("timestamp", 0)):
        if event.get("type") != "damage":
            continue
        pid = event.get("targetID")
        if pid is None or pid not in actor_lookup:
            continue
        # Only count hits from non-friendly sources (enemy/boss abilities)
        if event.get("sourceID") in friendly_ids:
            continue
        total_hit = event.get("unmitigatedAmount", 0) or event.get("amount", 0)
        if total_hit == 0:
            continue
        ability_id = event.get("abilityGameID", 0)
        ability_name = ability_names.get(ability_id, f"#{ability_id}")
        ts = event.get("timestamp", 0)
        elapsed_ms = ts - fight_start_ms
        minutes = int(elapsed_ms // 60000)
        seconds = int((elapsed_ms % 60000) // 1000)
        time_str = f"{minutes}:{seconds:02d}"
        amount_k = f"{total_hit / 1000:.1f}k"
        entry = results.setdefault(pid, {"hits": 0, "big_hits": 0, "details": []})
        entry["hits"] += 1
        max_hp = player_max_hp.get(pid, 0)
        if max_hp > 0 and total_hit >= max_hp * 0.10:
            entry["big_hits"] += 1
            entry["details"].append({"ability": ability_name, "amount_k": amount_k, "time": time_str})
    return results


# How many deaths before we stop counting (wipe cascade filter)
WIPE_DEATH_THRESHOLD = 4


def analyze_deaths(death_events: list, fight_start_ms: int, ability_names: dict = None,
                   fight_end_ms: int = 0, damage_events: list = None) -> dict:
    """Analyze death events, ignoring deaths after the Nth death (wipe cascade).
    Returns {targetID: [{"time", "ability", "fight_pct", "timeline"}, ...]}
    timeline = last 5s of damage hits before death: [{"ability", "amount_k", "sec_before"}, ...]
    """
    results = {}
    total_deaths = 0
    ability_names = ability_names or {}

    # Build per-player damage index for fast 5s window lookup
    dmg_by_player = {}
    for ev in (damage_events or []):
        if ev.get("type") != "damage":
            continue
        pid = ev.get("targetID")
        if pid is None:
            continue
        dmg_by_player.setdefault(pid, []).append(ev)

    for event in sorted(death_events, key=lambda e: e.get("timestamp", 0)):
        if total_deaths >= WIPE_DEATH_THRESHOLD:
            break
        target_id = event.get("targetID")
        if target_id is None:
            continue
        total_deaths += 1
        ts_ms = event.get("timestamp", 0)
        relative_sec = (ts_ms - fight_start_ms) / 1000.0
        minutes = int(relative_sec // 60)
        seconds = int(relative_sec % 60)
        time_str = f"{minutes}:{seconds:02d}"
        killing_id = event.get("killingAbilityGameID", 0)
        killing_name = ability_names.get(killing_id, f"ID:{killing_id}" if killing_id else "Unknown")
        fight_elapsed = ts_ms - fight_start_ms
        pct = round(fight_elapsed / fight_end_ms * 100) if fight_end_ms > 0 else 0

        # Last 5 seconds of damage events before this death
        timeline = []
        window_start = ts_ms - 5000
        for dev in dmg_by_player.get(target_id, []):
            dts = dev.get("timestamp", 0)
            if window_start <= dts <= ts_ms:
                daid = dev.get("abilityGameID", 0)
                dname = ability_names.get(daid, f"#{daid}")
                amt = dev.get("amount", 0)
                sec_before = (ts_ms - dts) / 1000.0
                timeline.append({"ability": dname, "amount_k": f"{amt/1000:.0f}k" if amt >= 1000 else str(amt), "sec_before": sec_before})
        timeline.sort(key=lambda x: x["sec_before"], reverse=True)  # oldest first

        results.setdefault(target_id, []).append({
            "time": time_str, "ability": killing_name, "fight_pct": pct, "timeline": timeline,
        })

    return results


# ─── Tracked Spells (matched by abilityGameID) ───────────────────────────────
# WCL cast events return abilityGameID (numeric), not ability names.

# Healthstone (tracked separately from health pots)
HEALTHSTONE_IDS = {
    5512,    # Healthstone
    6262,    # Healthstone (Midnight)
}

# Health Potions
HEALTH_POT_IDS = {
    432112,  # Algari Healing Potion (TWW)
    431924,  # Algari Healing Potion (TWW alternate)
    241304,  # Silvermoon Health Potion (Midnight)
    241305,  # Silvermoon Health Potion (Midnight alternate)
    258138,  # Silvermoon Health Potion (Midnight alternate)
    1234768, # Silvermoon Health Potion (Midnight)
}

# Combat Potions
COMBAT_POT_IDS = {
    431932,  # Tempered Potion (TWW)
    432098,  # Potion of Unwavering Focus (TWW)
    431945,  # Light's Potential (TWW)
    432106,  # Void-Shrouded Tincture (TWW)
    1236616, # Light's Potential (Midnight)
    245898,  # Light's Potential (Midnight alternate)
    245897,  # Light's Potential (Midnight alternate)
    241308,  # Light's Potential (Midnight alternate)
    241309,  # Light's Potential (Midnight alternate)
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

ALL_TRACKED_IDS = HEALTHSTONE_IDS | HEALTH_POT_IDS | COMBAT_POT_IDS | CLASS_DEFENSIVE_IDS

# ID → display name
SPELL_NAMES = {
    5512: "Healthstone", 6262: "Healthstone", 432112: "Health Potion", 431924: "Health Potion",
    431932: "Tempered Potion", 432098: "Potion of Unwavering Focus",
    431945: "Light's Potential", 432106: "Void-Shrouded Tincture",
    1236616: "Light's Potential", 245898: "Light's Potential",
    245897: "Light's Potential", 241308: "Light's Potential", 241309: "Light's Potential",
    241304: "Health Potion", 241305: "Health Potion", 258138: "Health Potion", 1234768: "Health Potion",
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
    if ability_game_id in HEALTHSTONE_IDS:
        return "Healthstone"
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
    Layout: rows = players, columns = bosses (4 sub-cols: Health | Healthstone | Combat | Defensive).
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
        # Row 1: boss names (4 cols each)
        html += '<tr><th class="player-header" rowspan="2">Player</th>'
        for fd in fights:
            html += f'<th colspan="4" class="boss-name divider">{escape(fd.get("fight_name", "Boss"))}</th>'
        html += '</tr>'
        # Row 2: sub-column headers
        html += '<tr>'
        for _ in fights:
            html += '<th class="cast-h health-h divider">⚗ Health</th>'
            html += '<th class="cast-h health-h">💚 Stone</th>'
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
                colspan = 1 + len(fights) * 4
                html += f'<tr class="role-sep"><td colspan="{colspan}">── {current_role}s ──</td></tr>'

            html += f'<tr>'
            html += (f'<td class="player-cell" style="color:{cls_color}">'
                     f'<span class="pname">{escape(pname)}</span><br>'
                     f'<span class="cname">{escape(char_name)}</span></td>')

            for fd in fights:
                casts = fd.get("player_casts", {}).get(pid, [])
                health      = [c for c in casts if c["category"] == "Health"]
                healthstone = [c for c in casts if c["category"] == "Healthstone"]
                combat      = [c for c in casts if c["category"] == "Combat Pot"]
                defensive   = [c for c in casts if c["category"] == "Defensive"]

                for items, cat_cls, is_first in [
                    (health,      "health-cell", True),
                    (healthstone, "health-cell", False),
                    (combat,      "combat-cell", False),
                    (defensive,   "def-cell",    False),
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


def _build_boss_html(boss_data: dict, actor_lookup: dict) -> dict:
    """Build HTML for each boss tab. Returns {boss_name: html_string}."""
    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}
    results = {}

    for boss_idx, (boss_name, fights) in enumerate(boss_data.items()):
        table_id = f"boss-tbl-{boss_idx}"
        boss_name_base = boss_name.rsplit(" (", 1)[0]
        mech_defs      = BOSS_MECHANICS.get(boss_name_base, [])
        has_interrupts = boss_name_base in BOSS_HAS_INTERRUPTS
        total_cols     = 16 + len(mech_defs)  # base 16 + mechanic columns

        # Collect all pids across all fights for this boss
        all_pids_set = set()
        for fight in fights:
            all_pids_set.update(fight.get("deaths", {}).keys())
            all_pids_set.update(fight.get("all_player_ids", set()))

        def pid_sort(pid):
            actor = actor_lookup.get(pid, {})
            char_name = actor.get("name", "")
            pname, _ = lookup_roster(char_name)
            role = PLAYER_ROLES.get(pname, "DPS")
            return (ROLE_ORDER.get(role, 2), pname.lower())

        sorted_pids = sorted(all_pids_set, key=pid_sort)

        # ── Fight overview (aggregated across all splits) ──
        ov_dmg_done = ov_healing = ov_dmg_taken = ov_avoid = 0
        for fight in fights:
            for pid, v in fight.get("uptime_map", {}).items():
                ov_dmg_done += v.get("total", 0) if isinstance(v, dict) else 0
            for pid, v in fight.get("healing_map", {}).items():
                ov_healing += v if isinstance(v, int) else 0
            for pid, v in fight.get("dmg_taken", {}).items():
                ov_dmg_taken += v if isinstance(v, int) else 0
            for pid, av in fight.get("avoidable_damage", {}).items():
                ov_avoid += av.get("hits", 0) if isinstance(av, dict) else 0
        def _hfmt(n):
            return f"{n/1_000_000_000:.1f}B" if n >= 1_000_000_000 else (f"{n/1_000_000:.1f}M" if n >= 1_000_000 else f"{n//1000}k")
        overview_html  = '<div class="boss-overview">'
        overview_html += f'<span class="ov-item"><span class="ov-label">⚔ Dmg Done</span><span class="ov-val">{_hfmt(ov_dmg_done) if ov_dmg_done else "—"}</span></span>'
        overview_html += f'<span class="ov-item"><span class="ov-label">💚 Healing</span><span class="ov-val">{_hfmt(ov_healing) if ov_healing else "—"}</span></span>'
        overview_html += f'<span class="ov-item"><span class="ov-label">🛡 Dmg Taken</span><span class="ov-val">{_hfmt(ov_dmg_taken) if ov_dmg_taken else "—"}</span></span>'
        overview_html += f'<span class="ov-item"><span class="ov-label">☠ Avoid Hits</span><span class="ov-val">{ov_avoid if ov_avoid else "—"}</span></span>'
        overview_html += '</div>'

        html = overview_html
        html += f'<div class="table-wrap"><table id="{table_id}" class="detail-col-hidden"><thead>'
        html += '<tr>'
        html += '<th style="width:10px"></th>'
        html += '<th class="player-header">Player</th>'
        html += '<th>Char</th>'
        html += '<th>Split</th>'
        html += '<th class="parse-h">Parse %</th>'
        html += '<th class="dmg-h">Damage</th>'
        html += '<th class="heal-h">Healing</th>'
        html += '<th class="death-h">Deaths</th>'
        html += '<th class="death-h">Killed by (time)</th>'
        html += '<th class="dmg-h">Dmg Taken</th>'
        html += '<th class="uptime-h">Uptime %</th>'
        html += '<th class="interrupt-h">Interrupts</th>'
        html += '<th class="avoid-h">Avoid Hits</th>'
        html += '<th class="avoid-h">&gt;10% HP Hits</th>'
        html += f'<th class="detail-h" onclick="toggleDetails(\'{table_id}\', this)" title="Click to expand/collapse">▶ Details</th>'
        for m in mech_defs:
            css = "mech-soak-h" if m["type"] == "soak" else "mech-bad-h"
            tip = escape(m.get("name", m["label"])).replace("'", "&#39;")
            html += f'<th class="{css}" style="cursor:help" data-htip=\'{tip}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()">{escape(m["label"])}</th>'
        html += '<th>Notes</th>'
        html += '</tr></thead><tbody>'

        PARSE_COLORS = {99:"#E5CC80", 95:"#FF8000", 75:"#A335EE", 50:"#0070DD", 25:"#1EFF00"}

        # Group by split, then role, then player
        for fi, fight in enumerate(fights, 1):
            html += f'<tr class="role-sep"><td colspan="{total_cols}" style="color:#a0b4ff;font-size:13px;padding:8px 10px;">── Split {fi} ──</td></tr>'
            current_role = None
            for pid in sorted_pids:
                deaths_map = fight.get("deaths", {})
                if pid not in fight.get("all_player_ids", set()) and pid not in deaths_map:
                    continue
                actor = actor_lookup.get(pid, {})
                char_name = actor.get("name", f"ID-{pid}")
                cls = actor.get("subType", "Unknown")
                cls_color = CLASS_COLORS.get(cls, "#ccc")
                pname, _ = lookup_roster(char_name)
                role = PLAYER_ROLES.get(pname, "DPS")

                if role != current_role:
                    current_role = role
                    html += f'<tr class="role-sep"><td colspan="{total_cols}">── {current_role}s ──</td></tr>'

                avoidable    = fight.get("avoidable_damage", {})
                dmg_taken    = fight.get("dmg_taken", {})
                uptime_map   = fight.get("uptime_map", {})
                interrupts   = fight.get("interrupts", {})
                healing_map  = fight.get("healing_map", {})
                rankings_map = fight.get("rankings_map", {})
                fight_dur    = fight.get("fight_dur_ms", 0)
                fight_mechs  = fight.get("mechanics_data", {})

                death_list = deaths_map.get(pid, [])
                av = avoidable.get(pid, {})
                d_raw = dmg_taken.get(pid, 0)
                dmg_str = f"{d_raw/1_000_000:.1f}M" if d_raw >= 1_000_000 else (f"{d_raw/1000:.0f}k" if d_raw > 0 else "—")
                active = uptime_map.get(pid, {}).get("activeTime", 0) if isinstance(uptime_map.get(pid), dict) else 0
                uptime_str = f"{min(active / fight_dur * 100, 100):.0f}%" if fight_dur > 0 and active > 0 else "—"
                interrupt_count = interrupts.get(pid, 0)
                interrupt_str = str(interrupt_count) if interrupt_count > 0 else "—"
                interrupt_style = (
                    ' style="background:#1A3A1A"' if interrupt_count > 0
                    else (' style="background:#5D1A1A"' if has_interrupts else "")
                )
                hits = av.get("hits", 0) or 0
                big_hits = av.get("big_hits", 0) or 0
                details_list = av.get("details", [])
                details_str = "<br>".join(f'{escape(d["ability"])}: {escape(d["amount_k"])} @ {escape(d["time"])}' for d in details_list) or "—"

                # Parse %
                parse_pct = rankings_map.get(char_name.lower())
                if parse_pct is not None:
                    pc = int(parse_pct)
                    pc_color = next((v for k, v in sorted(PARSE_COLORS.items(), reverse=True) if pc >= k), "#666")
                    parse_cell = f'<span style="color:{pc_color};font-weight:600">{pc}</span>'
                else:
                    parse_cell = "—"

                dmg_done_raw = uptime_map.get(pid, {}).get("total", 0) if isinstance(uptime_map.get(pid), dict) else 0
                dmg_done_str = _hfmt(dmg_done_raw) if dmg_done_raw > 0 else "—"
                heal_raw = healing_map.get(pid, 0)
                heal_str = _hfmt(heal_raw) if heal_raw > 0 else "—"

                death_count = len(death_list)
                killed_str = "<br>".join(f'{escape(d["ability"])} @ {escape(d["time"])} ({d.get("fight_pct", 0)}%)' for d in death_list) or "—"

                # Build pre-death timeline tooltip HTML
                death_tip_html = ""
                if death_list:
                    parts = []
                    for d in death_list:
                        tl = d.get("timeline", [])
                        block = f"<b style='color:#e57373'>☠ {escape(d['ability'])}</b> @ {escape(d['time'])} ({d.get('fight_pct',0)}%)"
                        if tl:
                            rows = []
                            for t in tl:
                                is_killing = abs(t["sec_before"]) < 0.05
                                color = "#ff6b6b" if is_killing else "#aaa"
                                label = " ← killing blow" if is_killing else f"-{t['sec_before']:.1f}s"
                                rows.append(f"<span style='color:{color}'>{escape(t['ability'])}: {escape(t['amount_k'])} &nbsp;<span style='color:#666;font-size:11px'>{label}</span></span>")
                            block += "<br>" + "<br>".join(rows)
                        parts.append(block)
                    death_tip_html = ("<hr style='border-color:#333;margin:6px 0'>").join(parts)

                row_cls = "boss-death-row" if death_count > 0 else ""
                bighit_cls = " bighit-row" if big_hits > 0 else ""

                html += f'<tr class="{row_cls}">'
                html += f'<td style="background:{cls_color};width:4px;padding:0;min-width:4px"></td>'
                html += f'<td class="player-cell"><span class="pname">{escape(pname)}</span></td>'
                html += f'<td><span class="cname" style="color:{cls_color}">{escape(char_name)}</span></td>'
                html += f'<td class="center">Split {fi}</td>'
                html += f'<td class="center parse-h">{parse_cell}</td>'
                html += f'<td class="center dmg-h">{dmg_done_str}</td>'
                html += f'<td class="center heal-h">{heal_str}</td>'
                if death_count > 0 and death_tip_html:
                    tip_attr = death_tip_html.replace("'", "&#39;")
                    html += f'<td class="center death-h" style="cursor:help" data-htip=\'{tip_attr}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()"><span class="death-num">{death_count}</span></td>'
                    html += f'<td class="death-h" data-htip=\'{tip_attr}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()" style="cursor:help">{killed_str}</td>'
                else:
                    html += f'<td class="center death-h">{"<span class=death-num>" + str(death_count) + "</span>" if death_count > 0 else "—"}</td>'
                    html += f'<td class="death-h">{killed_str}</td>'
                html += f'<td class="center dmg-h">{dmg_str}</td>'
                html += f'<td class="center uptime-h">{uptime_str}</td>'
                html += f'<td class="center interrupt-h"{interrupt_style}>{interrupt_str}</td>'
                html += f'<td class="center avoid-h">{hits if hits else "—"}</td>'
                html += f'<td class="center avoid-h{bighit_cls}">{big_hits if big_hits else "—"}</td>'
                html += f'<td class="detail-cell">{details_str}</td>'
                pid_mechs = fight_mechs.get(pid, {})
                for m in mech_defs:
                    count = pid_mechs.get(m["label"], 0)
                    if count:
                        bg = "#1A3D1A" if m["type"] == "soak" else "#5D1A1A"
                        html += f'<td class="center" style="background:{bg}">{count}</td>'
                    else:
                        html += '<td class="center">—</td>'
                html += '<td></td>'
                html += '</tr>'

        html += '</tbody></table></div>'
        results[boss_name] = html

    return results


def write_html(records: list, report_info: dict, report_code: str, output_path: str,
               split_data: dict = None, actors: list = None, boss_data: dict = None):
    """Write the full audit report as a tabbed HTML file."""
    guild_name = "Unknown Guild"
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "Unknown Guild")
    report_title = report_info.get("title", report_code)
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.fromtimestamp(report_info["startTime"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")

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

    # ── Boss tabs data ──
    _boss_actor_lookup = {a["id"]: a for a in (actors or [])}
    boss_htmls = _build_boss_html(boss_data or {}, _boss_actor_lookup)

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

/* ── Column borders ── */
td, th {{ border-left: 1px solid rgba(255,255,255,0.07); }}
td:first-child, th:first-child {{ border-left: none; }}

/* ── Boss overview bar ── */
.boss-overview {{
  display: flex; flex-wrap: wrap; gap: 12px;
  margin-bottom: 14px; padding: 10px 14px;
  background: rgba(255,255,255,0.04); border-radius: 8px;
  border: 1px solid rgba(255,255,255,0.08);
}}
.ov-item {{ display: flex; flex-direction: column; align-items: center; min-width: 90px; }}
.ov-label {{ font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: .05em; }}
.ov-val {{ font-size: 16px; font-weight: 700; color: #e0e0e0; margin-top: 2px; }}

/* ── Boss tabs ── */
.death-h {{ color: #e57373; }}
.dmg-h {{ color: #ffb74d; }}
.heal-h {{ color: #81c784; }}
.parse-h {{ color: #E5CC80; }}
.uptime-h {{ color: #81c784; }}
.interrupt-h {{ color: #64b5f6; }}
.avoid-h {{ color: #ce93d8; }}
.detail-h {{ color: #7289DA; white-space: normal; min-width: 30px; cursor: pointer; user-select: none; }}
.mech-bad-h {{ color: #ff7070; }}
.mech-soak-h {{ color: #66bb6a; }}
.has-tip {{ position: relative; cursor: help; }}
.has-tip::after {{ content: attr(data-tip); position: absolute; bottom: 130%; left: 50%; transform: translateX(-50%);
    background: #1e2a3a; color: #e0e0e0; border: 1px solid #4a5568; border-radius: 6px;
    padding: 6px 10px; font-size: 12px; font-weight: normal; white-space: nowrap;
    pointer-events: none; opacity: 0; transition: opacity 0.15s; z-index: 100; max-width: 340px; white-space: normal; text-align: left; }}
.has-tip:hover::after {{ opacity: 1; }}
.detail-h:hover {{ color: #a0b4ff; }}
.detail-cell {{ white-space: normal; min-width: 180px; font-size: 12px; color: #ccc; }}
.detail-col-hidden .detail-h span.detail-content,
.detail-col-hidden .detail-cell {{ display: none; }}
.detail-col-hidden .detail-h {{ min-width: 0; }}
.boss-death-row td {{ background: rgba(61,16,16,0.4); }}
.boss-death-row .player-cell {{ background: rgba(80,10,10,0.6); }}
.death-num {{ color: #e57373; font-weight: bold; }}
.interrupt-yes {{ color: #00FF88; font-weight: bold; }}
.bighit-row {{ color: #ff8a65; font-weight: bold; }}
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
  {''.join(f'<button class="tab-btn" onclick="switchTab(\'boss-{i}\', this)">⚔ {escape(name)}</button>' for i, name in enumerate(boss_htmls))}
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

<!-- ── BOSS TABS ── -->
{''.join(f'<div id="tab-boss-{i}" class="tab-content">{content}</div>' for i, content in enumerate(boss_htmls.values()))}

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
function toggleDetails(tableId, btn) {{
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  const hidden = tbl.classList.toggle('detail-col-hidden');
  btn.textContent = hidden ? '▶ Details' : '▼ Details';
}}
const _htip = document.createElement('div');
_htip.id = 'htip';
Object.assign(_htip.style, {{
  position:'fixed', display:'none', pointerEvents:'none', zIndex:'9999',
  background:'#1a2233', border:'1px solid #4a5568', borderRadius:'8px',
  padding:'10px 14px', fontSize:'12px', color:'#e0e0e0',
  maxWidth:'340px', lineHeight:'1.6', boxShadow:'0 4px 16px #0008'
}});
document.body.appendChild(_htip);
document.addEventListener('mousemove', e => {{
  if (_htip.style.display !== 'none') {{
    const x = e.clientX + 16, y = e.clientY - 10;
    _htip.style.left = (x + _htip.offsetWidth > window.innerWidth ? e.clientX - _htip.offsetWidth - 8 : x) + 'px';
    _htip.style.top  = (y + _htip.offsetHeight > window.innerHeight ? e.clientY - _htip.offsetHeight - 8 : y) + 'px';
  }}
}});
function showHTip(el) {{
  const d = el.getAttribute('data-htip');
  if (!d) return;
  _htip.innerHTML = d;
  _htip.style.display = 'block';
}}
function hideHTip() {{ _htip.style.display = 'none'; }}
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
    ("Nope",        "Tank",    [("NopeDK","Alt"), ("Nopebrew","Main"), ("Nøpæ","Main")]),
    ("Phyxius",     "Tank",    [("Phyxius","Main"), ("Phyxy","Alt")]),
    ("Toshiko",     "Healer",  [("Toshiko","Main"), ("Evokenooblal","Alt")]),
    ("Minxy",       "Healer",  [("Minxymender","Alt"), ("Minxycat","Main")]),
    ("Hipe",        "Healer",  [("Cype","Alt"), ("Wype","Main")]),
    ("Zush",        "Healer",  [("Zush","Main"), ("Züsh","Main"), ("Ryukho","Alt"), ("Ryuhko","Alt")]),
    ("Jimy",        "DPS",     [("jaime","Main"), ("Käz","Alt"), ("Madhmag","Alt")]),
    ("Ice",         "DPS",     [("Icecoldleap","Main"), ("Chocoice","Alt")]),
    ("Zodiacos",    "DPS",     [("Zodiacos","Alt")]),
    ("Kutcher",     "DPS",     [("Kutcherdhtwo","Alt"), ("Kutchersplit","Alt")]),
    ("Hypno",       "DPS",     [("Hypno","Main"), ("Hypnodh","Alt")]),
    ("Brunaine",    "DPS",     [("Brunainevoke","Alt"), ("Brunainehunt","Alt")]),
    ("Beldryk",     "DPS",     [("beldrýk","Main"), ("Beldrýk","Main"), ("beldryc","Alt")]),
    ("Potrenu",     "DPS",     [("Potrenu","Main"), ("Potrenuu","Alt")]),
    ("Shamishan",   "DPS",     [("Samdracson","Alt"), ("Shamishan","Main")]),
    ("Madonis",     "DPS",     [("Madonis","Alt"), ("Madonisvoker","Alt")]),
    ("Kaze",        "DPS",     [("Kazeofscales","Main"), ("Kazeoflight","Alt")]),
    ("Upyeah",      "DPS",     [("upyeah","Alt"), ("Upyeah","Alt"), ("upyeäh","Alt")]),
    ("Mindhacker",  "DPS",     [("Mindrage","Alt"), ("Mindhacker","Main")]),
    ("Uncleyoinky", "DPS",     [("allblues","Alt"), ("Allblues","Alt"), ("uncleyoinky","Main")]),
    ("Nizze",       "DPS",     [("Nizzedk","Alt"), ("Nizze","Main")]),
    ("Mostbanned",  "DPS",     [("Mosta","Alt"), ("Mostbanned","Alt")]),
    ("Malheiro",    "DPS",     [("Rödinhas","Alt"), ("Malheiro","Main")]),
    ("Bolters",     "DPS",     [("Schmosba","Main"), ("Ipala","Alt"), ("Devert","Alt")]),
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


# ─── Boss Mechanic Definitions ────────────────────────────────────────────────
# Per-boss spell tracking. type: "dmg_hits" = bad (red), "soak" = positive (green)
BOSS_MECHANICS = {
    "Imperator Averzian": [
        {"label": "Shad.Advance",  "name": "Shadow's Advance — stood within 10y when cast (knockback)",         "spell_ids": {1253691},                             "type": "dmg_hits"},
        {"label": "Void Fall",     "name": "Void Fall — stood in impact zone (avoidable)",                      "spell_ids": {1258883, 1269160},                    "type": "dmg_hits"},
        {"label": "Obliv.Wrath",   "name": "Oblivion's Wrath — stood in path (knockback)",                      "spell_ids": {1260718},                             "type": "dmg_hits"},
        {"label": "Umbral Col.",   "name": "Umbral Collapse — soak mechanic (higher = better)",                  "spell_ids": {1249262},                             "type": "soak"},
        {"label": "Gnash.Void",    "name": "Gnashing Void — hit by stacking debuff explosion",                  "spell_ids": {1255683},                             "type": "dmg_hits"},
        {"label": "Shad.Phalanx",  "name": "Shadow Phalanx — death source / avoidable burst",                  "spell_ids": {1284786},                             "type": "dmg_hits"},
    ],
    "Vorasius": [
        {"label": "Blisterburst",  "name": "Blisterburst — stood within 8y of explosion (avoidable debuff)",    "spell_ids": {1259186},                             "type": "dmg_hits"},
        {"label": "Claw Slam",     "name": "Shadowclaw Slam — non-tank avoidable hit",                          "spell_ids": {1241808, 1281954, 1281906, 1272328},  "type": "dmg_hits"},
        {"label": "Parasite Exp.", "name": "Parasite Expulsion — stood in swirlies (avoidable)",                "spell_ids": {1275558, 1275556},                    "type": "dmg_hits"},
        {"label": "Void Breath",   "name": "Void Breath — frontal breath (avoidable)",                          "spell_ids": {1257607},                             "type": "dmg_hits"},
    ],
    "Fallen-King Salhadaar": [
        {"label": "Tort.Extract",  "name": "Torturous Extract — stood in puddles (avoidable)",                  "spell_ids": {1245592},                             "type": "dmg_hits"},
        {"label": "Umbral Beams",  "name": "Umbral Beams — stood in rotating beams (avoidable)",                "spell_ids": {1260030},                             "type": "dmg_hits"},
        {"label": "Void Exposure", "name": "Void Exposure — touched a Void Orb (avoidable)",                   "spell_ids": {1250828},                             "type": "dmg_hits"},
        {"label": "Twilight Spk.", "name": "Twilight Spikes — stood in spike zone (avoidable)",                 "spell_ids": {1251213},                             "type": "dmg_hits"},
    ],
    "Vaelgor & Ezzorak": [
        {"label": "Impale",        "name": "Impale (Rakfang) — avoidable hit, causes stun",                     "spell_ids": {1265152},                             "type": "dmg_hits"},
        {"label": "Dread Breath",  "name": "Dread Breath — non-tank hit (avoidable), causes fear",              "spell_ids": {1244225, 1255979},                    "type": "dmg_hits"},
        {"label": "Gloomfield",    "name": "Gloomfield — stood in void zone ticks (avoidable)",                 "spell_ids": {1245421},                             "type": "dmg_hits"},
        {"label": "Tail Lash",     "name": "Tail Lash (Vaelwing) — avoidable knockback hit",                    "spell_ids": {1264467},                             "type": "dmg_hits"},
        {"label": "Nullbeam",      "name": "Nullbeam — soak stacks to reduce Nullzone size (higher = better)",  "spell_ids": {1283856, 1262688},                    "type": "soak"},
    ],
    "Lightblinded Vanguard": [
        {"label": "Final Verdict", "name": "Final Verdict — avoidable hit after gaining Judgment",              "spell_ids": {1251812},                             "type": "dmg_hits"},
        {"label": "Divine Toll",   "name": "Divine Toll — stood in path (avoidable, causes silence)",           "spell_ids": {1248652},                             "type": "dmg_hits"},
        {"label": "Exec.Sentence", "name": "Execution Sentence — soak mechanic (higher = better)",             "spell_ids": {1249024},                             "type": "soak"},
        {"label": "Trampled",      "name": "Trampled — knocked away / run over (avoidable)",                    "spell_ids": {1249135},                             "type": "dmg_hits"},
        {"label": "Div.Hammer",    "name": "Divine Hammer — spiraling hammer from Execution Sentence (avoid)",  "spell_ids": {1249047},                             "type": "dmg_hits"},
    ],
    "Crown of the Cosmos": [
        {"label": "Silverstrike",  "name": "Silverstrike Arrow/Ricochet — hit without void effect (avoidable)", "spell_ids": {1233649, 1237729},                    "type": "dmg_hits"},
        {"label": "Brstng Empty.", "name": "Bursting Emptiness — stood in path of burst (avoidable)",           "spell_ids": {1255378},                             "type": "dmg_hits"},
        {"label": "Void Remnants", "name": "Void Remnants/Expulsion — stood in void pool (avoidable)",          "spell_ids": {1233826, 1242553},                    "type": "dmg_hits"},
        {"label": "Singularity",   "name": "Singularity Eruption — stood in impact zone (avoidable)",           "spell_ids": {1235631},                             "type": "dmg_hits"},
        {"label": "Dev.Cosmos",    "name": "Devouring Cosmos — stood in zone, 99% reduced healing (avoidable)", "spell_ids": {1238882},                             "type": "dmg_hits"},
        {"label": "Grav.Collapse", "name": "Gravity Collapse — failed to space Aspect of the End removals",     "spell_ids": {1239095},                             "type": "dmg_hits"},
    ],
    "Chimaerus, the Undreamt God": [
        {"label": "Alndust Ess.",  "name": "Alndust Essence — stood in pool ticks (avoidable)",                 "spell_ids": {1245919},                             "type": "dmg_hits"},
        {"label": "Alndust Uph.",  "name": "Alndust Upheaval — soak mechanic (higher = better)",                "spell_ids": {1262305, 1246827},                    "type": "soak"},
        {"label": "Disc.Roar",     "name": "Discordant Roar — avoidable damage from Colossal Horrors",          "spell_ids": {1249207},                             "type": "dmg_hits"},
        {"label": "Rift Emerg.",   "name": "Rift Emergence — death source / avoidable burst",                   "spell_ids": {1258610},                             "type": "dmg_hits"},
    ],
}

# Bosses where 0 interrupts = red (boss has meaningful interruptible abilities)
BOSS_HAS_INTERRUPTS = {
    "Imperator Averzian", "Fallen-King Salhadaar", "Vaelgor & Ezzorak",
    "Lightblinded Vanguard", "Crown of the Cosmos", "Chimaerus, the Undreamt God",
}


# ─── XLSX Output ─────────────────────────────────────────────────────────────

# Class background colors (lighter/muted versions for row backgrounds)
XLSX_CLASS_BG = {
    "DeathKnight": "77C41E3A", "DemonHunter": "77A330C9", "Druid": "77FF7C0A",
    "Evoker": "7733937F", "Hunter": "77AAD372", "Mage": "773FC7EB",
    "Monk": "7700FF98", "Paladin": "77F48CBA", "Priest": "77FFFFFF",
    "Rogue": "77FFF468", "Shaman": "770070DD", "Warlock": "778788EE",
    "Warrior": "77C69B6D",
}


def _build_boss_sheet(ws, boss_name: str, fights: list, report_info: dict, actor_lookup: dict):
    """One sheet per boss: rows = chars from both splits, cols = Split | Deaths | Death Times | Dmg Hits | Big Hits | Notes."""
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill  = PatternFill("solid", fgColor="1a1a2e")
    data_font    = Font(name="Arial", size=10, color="FFFFFF")
    bold_font    = Font(name="Arial", size=10, color="FFFFFF", bold=True)
    death_font   = Font(name="Arial", size=10, color="CC0000", bold=True)
    border       = Border(bottom=Side(style="thin", color="333333"))
    center       = Alignment(horizontal="center", vertical="center")
    left         = Alignment(horizontal="left", vertical="center")
    wrap         = Alignment(horizontal="left", vertical="top", wrap_text=True)
    notes_fill   = PatternFill("solid", fgColor="1E2736")
    no_death_fill = PatternFill("solid", fgColor="111827")  # dark = clean
    death_fill   = PatternFill("solid", fgColor="3D1010")   # dark red = died

    def _parse_color(pct):
        """Return ARGB fill color for parse percentile (WarcraftLogs style)."""
        if not isinstance(pct, int): return "111827"
        if pct >= 99: return "E5CC80"   # gold/legendary
        if pct >= 95: return "FF8000"   # orange/epic
        if pct >= 75: return "A335EE"   # purple/rare
        if pct >= 50: return "0070DD"   # blue/uncommon
        if pct >= 25: return "1EFF00"   # green/common
        return "666666"                  # grey/low

    guild_name  = report_info.get("guild", {}).get("name", "") if report_info.get("guild") else ""
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.fromtimestamp(report_info["startTime"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")

    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}

    # Pre-fill entire sheet with dark background
    dk = PatternFill("solid", fgColor="0D1117")
    for r in range(1, 300):
        for c in range(1, 55):
            ws.cell(row=r, column=c).fill = dk

    # Title
    ws.merge_cells("A1:M1")
    ws["A1"].value = f"{boss_name} — {guild_name} ({report_date})"
    ws["A1"].font  = Font(name="Arial", bold=True, size=13, color="7289DA")
    ws.row_dimensions[1].height = 22

    ws["A2"].value = f"Deaths filtered: only first {WIPE_DEATH_THRESHOLD} deaths shown (wipe cascade excluded)"
    ws["A2"].font  = Font(name="Arial", size=9, color="888888", italic=True)

    # Row 3: fight overview (aggregated across all splits)
    def _k(v):
        if v >= 1_000_000_000: return f"{v/1_000_000_000:.1f}B"
        if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
        if v > 0: return f"{v/1000:.0f}k"
        return "—"

    total_dur_ms   = sum(f.get("fight_dur_ms", 0) for f in fights)
    total_dmg_done = sum(e.get("total", 0) for f in fights for e in f.get("uptime_map", {}).values())
    total_healing  = sum(v for f in fights for v in f.get("healing_map", {}).values())
    total_dmg_taken = sum(v for f in fights for v in f.get("dmg_taken", {}).values())
    total_avoidable = sum(e.get("hits", 0) for f in fights for e in f.get("avoidable_damage", {}).values())
    dur_m, dur_s   = divmod(total_dur_ms // 1000, 60)

    overview_items = [
        ("⏱ Duration", f"{dur_m}:{dur_s:02d}"),
        ("⚔ Dmg Done", _k(total_dmg_done)),
        ("💚 Healing", _k(total_healing)),
        ("🛡 Dmg Taken", _k(total_dmg_taken)),
        ("☠ Avoid Hits", str(total_avoidable) if total_avoidable else "—"),
    ]
    ov_fill  = PatternFill("solid", fgColor="16213e")
    ov_label = Font(name="Arial", size=9, color="888888")
    ov_value = Font(name="Arial", size=11, bold=True, color="7289DA")
    for i, (label, value) in enumerate(overview_items):
        col = 2 + i * 2
        ws.cell(row=3, column=col,     value=label).font  = ov_label
        ws.cell(row=3, column=col,     value=label).fill  = ov_fill
        ws.cell(row=3, column=col + 1, value=value).font  = ov_value
        ws.cell(row=3, column=col + 1, value=value).fill  = ov_fill
        ws.cell(row=3, column=col + 1).alignment = center
    ws.row_dimensions[3].height = 22

    # Dynamic columns: base + boss-specific mechanics + Notes
    boss_name_base = boss_name.rsplit(" (", 1)[0]
    mech_defs = BOSS_MECHANICS.get(boss_name_base, [])
    has_interrupts = boss_name_base in BOSS_HAS_INTERRUPTS

    hr = 5
    base_cols   = ["", "Player", "Char", "Split", "Parse %", "Damage", "Healing",
                   "Deaths", "Killed by (time)", "Dmg Taken", "Uptime %",
                   "Interrupts", "Avoid Hits", ">10% HP Hits", "Big Hit Details"]
    base_widths = [3, 16, 18, 8, 9, 12, 12, 10, 28, 12, 10, 11, 11, 13, 40]
    mech_cols   = [m["label"] for m in mech_defs]
    cols   = base_cols + mech_cols + ["Notes"]
    widths = base_widths + [11] * len(mech_cols) + [40]

    for ci, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[hr].height = 30
    ws.freeze_panes = "A6"
    # Hide "Big Hit Details" column (col 15) by default — unhide to expand
    ws.column_dimensions[get_column_letter(15)].hidden = True

    # Collect all chars across both splits, tag with split number
    rows = []
    for fi, fight in enumerate(fights, 1):
        split_num    = fi
        avoidable    = fight.get("avoidable_damage", {})
        dmg_taken    = fight.get("dmg_taken", {})
        uptime_map   = fight.get("uptime_map", {})
        healing_map  = fight.get("healing_map", {})
        rankings_map = fight.get("rankings_map", {})
        interrupts_map  = fight.get("interrupts", {})
        mechanics_data  = fight.get("mechanics_data", {})
        fight_dur    = fight.get("fight_dur_ms", 0)
        seen_pids    = set()

        def _fmt_k(val):
            if val >= 1_000_000: return f"{val/1_000_000:.1f}M"
            if val > 0:          return f"{val/1000:.0f}k"
            return "—"

        def _fmt_dmg_taken(pid):
            return _fmt_k(dmg_taken.get(pid, 0))

        def _fmt_dps_dmg(pid):
            return _fmt_k(uptime_map.get(pid, {}).get("total", 0))

        def _fmt_healing(pid):
            return _fmt_k(healing_map.get(pid, 0))

        def _fmt_uptime(pid, role):
            if role not in ("DPS", "Healer"):
                return "—"
            active = uptime_map.get(pid, {}).get("activeTime", 0)
            if fight_dur > 0 and active > 0:
                return f"{min(active / fight_dur * 100, 100):.0f}%"
            return "—"

        def _fmt_parse(pid, char_name):
            pct = rankings_map.get(char_name.lower(), None)
            return pct if pct is not None else "—"

        def _fmt_interrupts(pid):
            count = interrupts_map.get(pid, 0)
            return count if count > 0 else "—"

        for pid, death_times in fight.get("deaths", {}).items():
            actor = actor_lookup.get(pid, {})
            char_name = actor.get("name", f"ID-{pid}")
            cls = actor.get("subType", "Unknown")
            player_name, _ = lookup_roster(char_name)
            role = PLAYER_ROLES.get(player_name, "DPS")
            av = avoidable.get(pid, {})
            details_str = "\n".join(f"{d['ability']}: {d['amount_k']} @ {d['time']}" for d in av.get("details", []))
            rows.append({
                "player": player_name, "char": char_name, "cls": cls,
                "split": split_num, "role": role,
                "deaths": len(death_times),
                "death_times": "\n".join(f"{d['ability']} @ {d['time']} ({d.get('fight_pct', 0)}%)" for d in death_times),
                "parse": _fmt_parse(pid, char_name),
                "dps_dmg": _fmt_dps_dmg(pid),
                "healing": _fmt_healing(pid),
                "dmg_taken": _fmt_dmg_taken(pid),
                "uptime": _fmt_uptime(pid, role),
                "interrupts": _fmt_interrupts(pid),
                "dmg_hits": av.get("hits", "—"),
                "big_hits": av.get("big_hits", "—"),
                "details": details_str,
                "mechanics": mechanics_data.get(pid, {}),
            })
            seen_pids.add(pid)
        # Also include players with 0 deaths from cast events (they were in the fight)
        for pid in fight.get("all_player_ids", set()):
            if pid not in seen_pids:
                actor = actor_lookup.get(pid, {})
                char_name = actor.get("name", f"ID-{pid}")
                cls = actor.get("subType", "Unknown")
                player_name, _ = lookup_roster(char_name)
                role = PLAYER_ROLES.get(player_name, "DPS")
                av = avoidable.get(pid, {})
                details_str = "\n".join(f"{d['ability']}: {d['amount_k']} @ {d['time']}" for d in av.get("details", []))
                rows.append({
                    "player": player_name, "char": char_name, "cls": cls,
                    "split": split_num, "role": role,
                    "deaths": 0, "death_times": "",
                    "parse": _fmt_parse(pid, char_name),
                    "dps_dmg": _fmt_dps_dmg(pid),
                    "healing": _fmt_healing(pid),
                    "dmg_taken": _fmt_dmg_taken(pid),
                    "uptime": _fmt_uptime(pid, role),
                    "interrupts": _fmt_interrupts(pid),
                    "dmg_hits": av.get("hits", "—"),
                    "big_hits": av.get("big_hits", "—"),
                    "details": details_str,
                    "mechanics": mechanics_data.get(pid, {}),
                })

    # Sort: role → player name → split
    rows.sort(key=lambda r: (ROLE_ORDER.get(r["role"], 2), r["player"].lower(), r["split"]))

    row = hr + 1
    current_role = None
    for r in rows:
        if r["role"] != current_role:
            current_role = r["role"]
            for ci in range(1, len(cols) + 1):
                sc = ws.cell(row=row, column=ci)
                sc.fill = PatternFill("solid", fgColor="111827")
                sc.border = border
            ws.cell(row=row, column=2).value = f"── {current_role}s ──"
            ws.cell(row=row, column=2).font = Font(name="Arial", size=9, bold=True, color="7289DA")
            ws.row_dimensions[row].height = 30
            row += 1

        cls_hex = CLASS_COLORS.get(r["cls"], "#CCCCCC").lstrip("#")
        cls_bar_fill = PatternFill("solid", fgColor=cls_hex)
        died = r["deaths"] > 0
        row_fill = death_fill if died else no_death_fill
        dmg_hits      = r.get("dmg_hits", "—")
        big_hits      = r.get("big_hits", "—")
        details       = r.get("details", "")
        interrupt_val = r.get("interrupts", "—")
        big_fill      = PatternFill("solid", fgColor="5D2D1A") if isinstance(big_hits, int) and big_hits > 0 else row_fill
        interrupt_fill = (
            PatternFill("solid", fgColor="1A3A1A") if isinstance(interrupt_val, int) and interrupt_val > 0
            else PatternFill("solid", fgColor="5D1A1A") if has_interrupts
            else row_fill
        )
        char_font = Font(name="Arial", size=10, color=cls_hex, bold=False)

        parse_val  = r.get("parse", "—")
        parse_fill = PatternFill("solid", fgColor=_parse_color(parse_val)) if isinstance(parse_val, int) else row_fill
        parse_font = Font(name="Arial", size=10, color="000000" if isinstance(parse_val, int) else "FFFFFF", bold=isinstance(parse_val, int))

        def _mech_fill(mech, count):
            if not isinstance(count, int) or count == 0:
                return row_fill
            return PatternFill("solid", fgColor="1A3D1A") if mech["type"] == "soak" else PatternFill("solid", fgColor="5D1A1A")

        mech_data  = r.get("mechanics", {})
        mech_vals  = [mech_data.get(m["label"], 0) or "—" for m in mech_defs]
        mech_fills = [_mech_fill(m, mech_data.get(m["label"], 0)) for m in mech_defs]

        base_vals  = ["", r["player"], r["char"], f"Split {r['split']}",
                      parse_val, r.get("dps_dmg", "—"), r.get("healing", "—"),
                      r["deaths"] or "—", r["death_times"] or "—",
                      r.get("dmg_taken", "—"), r.get("uptime", "—"), interrupt_val,
                      dmg_hits, big_hits, details or "—"]
        base_fills = [cls_bar_fill, row_fill, row_fill, row_fill,
                      parse_fill, row_fill, row_fill,
                      row_fill, row_fill, row_fill, row_fill, interrupt_fill, row_fill, big_fill, big_fill]
        base_aligns = [left, left, left, center, center, center, center, center, wrap, center, center, center, center, center, wrap]
        base_fonts  = [data_font, bold_font, char_font, data_font,
                       parse_font, data_font, data_font,
                       death_font if died else data_font, data_font,
                       data_font, data_font, data_font, data_font, data_font, data_font]

        vals   = base_vals  + mech_vals  + [""]
        fills  = base_fills + mech_fills + [notes_fill]
        aligns = base_aligns + [center] * len(mech_defs) + [wrap]
        fonts  = base_fonts  + [data_font] * len(mech_defs) + [data_font]

        for ci, (val, fill, align, font) in enumerate(zip(vals, fills, aligns, fonts), 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = fill
            c.font = font
            c.alignment = align
            c.border = border

        ws.row_dimensions[row].height = 30
        row += 1


def write_xlsx(records: list, report_info: dict, report_code: str, output_path: str, split_data: dict = None, boss_data: dict = None, actors: list = None):
    """Write crafted gear data to XLSX with Mains, Alts, and Split tabs."""
    wb = Workbook()
    
    guild_name = "Unknown Guild"
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "Unknown Guild")
    report_title = report_info.get("title", report_code)
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.fromtimestamp(report_info["startTime"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")

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
    black_font = Font(name="Arial", size=10, color="FFFFFF")
    black_bold = Font(name="Arial", size=10, color="FFFFFF", bold=True)
    border = Border(bottom=Side(style="thin", color="333333"))
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    spark_font = Font(name="Arial", size=10, color="00FF88", bold=True)
    no_spark_font = Font(name="Arial", size=10, color="FFFFFF")
    no_craft_fill = PatternFill("solid", fgColor="3D3D3D")
    no_craft_font = Font(name="Arial", size=10, color="999999")

    def get_class_fill(cls_name):
        color = XLSX_CLASS_BG.get(cls_name, "77666666")
        return PatternFill("solid", fgColor=color)

    def prefill_dark(ws, rows=300, cols=60):
        """Pre-fill entire sheet area with dark background so no white cells show."""
        dk = PatternFill("solid", fgColor="0D1117")
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c).fill = dk

    def build_sheet(ws, title, player_data):
        # Title
        total_cols = 2 + max_items * 3
        prefill_dark(ws, rows=200, cols=total_cols + 5)
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
                        cell.font = no_craft_font if not has_items else Font(name="Arial", size=10, color="AAAAAA")
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

    # Build boss analysis tabs (deaths + notes)
    if boss_data:
        actor_lookup = {a["id"]: a for a in (actors or [])}
        for boss_name, fights in boss_data.items():
            safe_name = boss_name[:28]  # Excel sheet name limit
            ws_boss = wb.create_sheet(safe_name)
            _build_boss_sheet(ws_boss, boss_name, fights, report_info, actor_lookup)

    wb.save(output_path)
    print(f"[OK] XLSX report saved to: {output_path}")


def _build_split_sheet(ws, title, fights_data, report_info, actors_list):
    """Build a split tab showing consumable/defensive usage per player per fight.
    
    fights_data: list of { "fight_name": str, "difficulty": str, "player_casts": { sourceID: [cast_info] } }
    """
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1a1a2e")
    data_font = Font(name="Arial", size=10, color="FFFFFF")
    border = Border(bottom=Side(style="thin", color="333333"))
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    cat_fills = {
        "Health": PatternFill("solid", fgColor="77CC4444"),
        "Combat Pot": PatternFill("solid", fgColor="77AA44CC"),
        "Defensive": PatternFill("solid", fgColor="774488CC"),
    }

    guild_name = ""
    if report_info.get("guild"):
        guild_name = report_info["guild"].get("name", "")
    report_title = report_info.get("title", "")
    report_date = ""
    if report_info.get("startTime"):
        report_date = datetime.fromtimestamp(report_info["startTime"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")

    # Pre-fill entire sheet with dark background
    dk = PatternFill("solid", fgColor="0D1117")
    num_cols = 1 + len(fights_data) * 4 + 10
    for r in range(1, 200):
        for c in range(1, num_cols + 1):
            ws.cell(row=r, column=c).fill = dk

    full_title = f"{title} — {guild_name} — {report_title} ({report_date})"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(fights_data) * 3)
    ws["A1"].value = full_title
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="7289DA")
    ws.row_dimensions[1].height = 26

    ws["A2"].value = f"Report: {report_title} ({report_date})"
    ws["A2"].font = Font(name="Arial", size=9, color="888888", italic=True)

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

    sub_labels = ["Health Pots", "Healthstone", "Combat Pots", "Defensives"]
    col_widths  = [18, 14, 18, 28]
    num_cols    = len(sub_labels)

    for fi, fd in enumerate(fights_data):
        fname   = fd.get("fight_name", "Boss")
        base_ci = 2 + fi * num_cols  # first column of this boss group

        # Merge boss name across num_cols columns
        ws.merge_cells(start_row=hr_boss, start_column=base_ci, end_row=hr_boss, end_column=base_ci + num_cols - 1)
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
        c.font = Font(name="Arial", size=10, color="FFFFFF", bold=True)
        c.fill = cls_fill
        c.border = border
        c.alignment = Alignment(vertical="center", wrap_text=True)

        for fi, fd in enumerate(fights_data):
            base_col = 2 + fi * num_cols
            casts = fd.get("player_casts", {}).get(pid, [])

            # Group by category
            health      = [c for c in casts if c["category"] == "Health"]
            healthstone = [c for c in casts if c["category"] == "Healthstone"]
            combat      = [c for c in casts if c["category"] == "Combat Pot"]
            defensives  = [c for c in casts if c["category"] == "Defensive"]

            for j, (items, cat) in enumerate([(health, "Health"), (healthstone, "Healthstone"), (combat, "Combat Pot"), (defensives, "Defensive")]):
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
                    cell.font = Font(name="Arial", size=10, color="AAAAAA")
                    cell.alignment = center

                # Divider
                if j == 0:
                    cell.border = Border(bottom=Side(style="thin", color="333333"), left=Side(style="medium", color="7289DA"))

        # Height = tallest cell: max casts in any single category across all fights
        max_lines = 1
        for fd in fights_data:
            player_casts = fd.get("player_casts", {}).get(pid, [])
            for cat in ("Health", "Healthstone", "Combat Pot", "Defensive"):
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

    # Build ability name lookup from masterData
    ability_names = {}
    if report_info.get("masterData") and report_info["masterData"].get("abilities"):
        for ab in report_info["masterData"]["abilities"]:
            ability_names[ab["gameID"]] = ab["name"]
    
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
    
    # Build player max HP lookup from combatant info (stamina * 20 approximation)
    player_max_hp = {}
    for ce in all_combatant_events:
        pid = ce.get("sourceID")
        stamina = ce.get("stamina", 0)
        if pid and stamina:
            hp = stamina * 20
            player_max_hp[pid] = max(player_max_hp.get(pid, 0), hp)

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

    # ── Boss analysis: deaths + avoidable damage per boss (both splits combined) ──
    # boss_data: { boss_name: [ {fight_id, deaths, all_player_ids, avoidable_damage}, ... ] }
    boss_data = {}
    print(f"\n[...] Fetching death + damage events per boss...")
    for fight in selected_fights:
        fid   = fight["id"]
        fname = fight["name"]
        diff  = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(fight.get("difficulty", 0), "")
        boss_key = f"{fname} ({diff})"
        try:
            death_events     = fetch_death_events(token, report_code, fid)
            damage_events    = fetch_damage_taken_events(token, report_code, fid)
            interrupt_events = fetch_interrupt_events(token, report_code, fid)
            fight_start      = fight.get("startTime", 0)
            fight_end        = fight.get("endTime", 0)
            fight_dur_ms     = fight_end - fight_start
            deaths           = analyze_deaths(death_events, fight_start, ability_names, fight_end_ms=fight_dur_ms, damage_events=damage_events)
            avoidable        = analyze_avoidable_damage(damage_events, actor_lookup, fight_start, ability_names, player_max_hp)
            dmg_taken        = aggregate_damage_taken(damage_events, actor_lookup)
            uptime_map       = fetch_uptime_table(token, report_code, fid)
            healing_map      = fetch_healing_table(token, report_code, fid)
            rankings_map     = fetch_rankings(token, report_code, fid)
            interrupts       = analyze_interrupts(interrupt_events, actor_lookup)
            mech_defs        = BOSS_MECHANICS.get(fname, [])
            mechanics_data   = analyze_boss_mechanics(damage_events, actor_lookup, mech_defs)
            # Collect all player IDs: from casts, damage taken, uptime, and deaths
            all_pids = set()
            for sdata in split_data.values():
                for fd in sdata:
                    if fd.get("fight_id") == fid:
                        all_pids.update(fd.get("player_casts", {}).keys())
            all_pids.update(dmg_taken.keys())
            all_pids.update(uptime_map.keys())
            all_pids.update(deaths.keys())
            # Only keep actual player actors
            all_pids = {pid for pid in all_pids if pid in actor_lookup}
            if boss_key not in boss_data:
                boss_data[boss_key] = []
            boss_data[boss_key].append({
                "fight_id": fid,
                "fight_dur_ms": fight_dur_ms,
                "deaths": deaths,
                "all_player_ids": all_pids,
                "avoidable_damage": avoidable,
                "dmg_taken": dmg_taken,
                "uptime_map": uptime_map,
                "healing_map": healing_map,
                "rankings_map": rankings_map,
                "interrupts": interrupts,
                "mechanics_data": mechanics_data,
            })
            print(f"  [OK] {fname}: {len(deaths)} death(s), {len(avoidable)} player(s), {sum(interrupts.values())} interrupts.")
        except Exception as e:
            print(f"  [WARN] Could not fetch events for fight {fid}: {e}")

    # Output
    html_path = "craft_audit.html"
    write_html(records, report_info, report_code, html_path, split_data=split_data, actors=actors, boss_data=boss_data)

    xlsx_path = "craft_audit.xlsx"
    write_xlsx(records, report_info, report_code, xlsx_path, split_data=split_data, boss_data=boss_data, actors=actors)
    
    print(f"\nDone! Refresh craft_audit.html in your browser, or open craft_audit.xlsx.\n")


if __name__ == "__main__":
    main()