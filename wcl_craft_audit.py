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
import os
import json
import requests
from datetime import datetime, timezone
from html import escape
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from boss_mechanics import BOSS_MECHANICS, BOSS_HAS_INTERRUPTS
from tracked_spells import (HEALTHSTONE_IDS, HEALTH_POT_IDS, COMBAT_POT_IDS,
                             CLASS_DEFENSIVE_IDS, SPELL_NAMES)

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
        msgs = [e.get("message", "") for e in data["errors"]]
        raise RuntimeError(f"GraphQL error: {'; '.join(msgs)}")
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
                    bossPercentage
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


def analyze_frontal_failures(damage_events: list, actor_lookup: dict, mechanics: list,
                              fight_start_ms: int = 0, window_ms: int = 2000,
                              cast_events: list = None) -> list:
    """Detect frontal failures: 2+ friendly players hit by the same frontal cast.
    Uses cast event targetID (if available) to identify who had the ability.
    Only checks mechanics with type='frontal'.
    Returns list of {time_str, label, primary_player, others, hit_count}.
    """
    frontal_mechs = [m for m in mechanics if m.get("type") == "frontal"]
    if not frontal_mechs:
        return []

    spell_to_label = {}
    for m in frontal_mechs:
        for sid in m["spell_ids"]:
            spell_to_label[sid] = m["label"]

    # Build cast-target lookup: spell_id -> sorted list of (timestamp, targetID)
    cast_targets = {}  # spell_id -> [(ts, targetID)]
    for event in (cast_events or []):
        aid = event.get("abilityGameID", 0)
        if aid not in spell_to_label:
            continue
        tid = event.get("targetID")
        if tid and tid in actor_lookup:
            cast_targets.setdefault(aid, []).append((event.get("timestamp", 0), tid))
    for v in cast_targets.values():
        v.sort()

    friendly_ids = set(actor_lookup.keys())
    hits = []
    for event in damage_events:
        if event.get("type") != "damage":
            continue
        aid = event.get("abilityGameID", 0)
        if aid not in spell_to_label:
            continue
        pid = event.get("targetID")
        if pid not in friendly_ids:
            continue
        hits.append({"ts": event.get("timestamp", 0), "pid": pid, "label": spell_to_label[aid], "aid": aid})

    if not hits:
        return []

    hits.sort(key=lambda h: h["ts"])

    def _cast_target_at(aid, ts):
        """Find the cast targetID closest before ts for this spell."""
        candidates = cast_targets.get(aid, [])
        result = None
        for c_ts, c_tid in candidates:
            if c_ts <= ts + window_ms:
                result = c_tid
            else:
                break
        return result

    failures = []
    i = 0
    while i < len(hits):
        win_start = hits[i]["ts"]
        label     = hits[i]["label"]
        aid       = hits[i]["aid"]
        pids_hit  = set()
        j = i
        while j < len(hits) and hits[j]["ts"] - win_start <= window_ms and hits[j]["label"] == label:
            pids_hit.add(hits[j]["pid"])
            j += 1
        if len(pids_hit) >= 2:
            elapsed = (win_start - fight_start_ms) / 1000
            m, s    = divmod(int(elapsed), 60)
            time_str = f"{m}:{s:02d}"
            # Prefer cast targetID as primary; fall back to first hit
            primary_pid  = _cast_target_at(aid, win_start) or hits[i]["pid"]
            primary_name = actor_lookup.get(primary_pid, {}).get("name", f"ID-{primary_pid}")
            others = sorted(actor_lookup.get(pid, {}).get("name", f"ID-{pid}")
                            for pid in pids_hit if pid != primary_pid)
            failures.append({"time_str": time_str, "label": label,
                             "primary_player": primary_name, "others": others,
                             "hit_count": len(pids_hit)})
        i = j if j > i else i + 1

    return failures


def fetch_boss_cast_events(token: str, report_code: str, fight_id: int) -> list:
    """Fetch enemy cast events for a fight (boss casts), with pagination."""
    query = """
    query ($code: String!, $fightID: Int!, $startTime: Float) {
        reportData {
            report(code: $code) {
                events(dataType: Casts, fightIDs: [$fightID], hostilityType: Enemies,
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
    entries = table.get("data", {}).get("entries", []) if isinstance(table, dict) else []
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
    entries = table.get("data", {}).get("entries", []) if isinstance(table, dict) else []
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
    for fight_entry in rankings.get("data", []):
        for role_data in fight_entry.get("roles", {}).values():
            for char in role_data.get("characters", []):
                name = char.get("name", "").lower()
                pct = char.get("rankPercent", 0)
                if name and pct:
                    out[name] = round(pct)
    return out


def fetch_fight_graph(token: str, report_code: str, fight_id: int) -> dict:
    """Fetch DamageDone, DamageTaken, and Healing time-series graphs for a fight in one call.
    Returns {dps: [(t_s, val), ...], taken: [...], heal: [...]}.
    Each point is (time_seconds_from_fight_start, total_value_in_5s_bucket).
    Returns empty series on any error.
    """
    query = """
    query ($code: String!, $fightID: Int!) {
        reportData {
            report(code: $code) {
                dps:   graph(dataType: DamageDone,  fightIDs: [$fightID], hostilityType: Friendlies)
                taken: graph(dataType: DamageTaken, fightIDs: [$fightID], hostilityType: Friendlies)
                heal:  graph(dataType: Healing,     fightIDs: [$fightID], hostilityType: Friendlies)
            }
        }
    }
    """
    try:
        result = query_wcl(token, query, {"code": report_code, "fightID": fight_id})
        rep    = result["reportData"]["report"]
    except Exception:
        return {"dps": [], "taken": [], "heal": []}

    def _sum_series(graph_raw) -> list:
        """Sum all player entries per time bucket → [(t_s, total), ...], time 0-based."""
        if not isinstance(graph_raw, dict):
            return []
        inner = graph_raw.get("data", graph_raw)
        if isinstance(inner, dict):
            entries = inner.get("data", [])
        elif isinstance(inner, list):
            entries = inner
        else:
            return []
        bucket: dict = {}
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            for point in entry.get("data", []):
                if isinstance(point, (list, tuple)) and len(point) >= 2:
                    t_ms = int(point[0])
                    val  = point[1] or 0
                    bucket[t_ms] = bucket.get(t_ms, 0) + val
        if not bucket:
            return []
        t_min = min(bucket)
        return sorted((round((t - t_min) / 1000, 1), v) for t, v in bucket.items())

    return {
        "dps":   _sum_series(rep.get("dps")),
        "taken": _sum_series(rep.get("taken")),
        "heal":  _sum_series(rep.get("heal")),
    }


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
                             player_max_hp: dict = None, player_roles: dict = None) -> dict:
    """Count enemy-source damage-taken hits per player (proxy for avoidable damage).
    Returns {pid: {"hits": int, "big_hits": int, "details": [{"ability", "amount_k", "time"}]}}
    big_hits threshold: 50% HP for Tanks, 10% HP for everyone else.
    Melee auto-attacks (abilityGameID == 1) are excluded from details.
    """
    ability_names = ability_names or {}
    player_max_hp  = player_max_hp  or {}
    player_roles   = player_roles   or {}
    friendly_ids   = set(actor_lookup.keys())
    results = {}
    for event in sorted(damage_events, key=lambda e: e.get("timestamp", 0)):
        if event.get("type") != "damage":
            continue
        pid = event.get("targetID")
        if pid is None or pid not in actor_lookup:
            continue
        if event.get("sourceID") in friendly_ids:
            continue
        ability_id = event.get("abilityGameID", 0)
        if ability_id == 1:  # skip melee auto-attacks
            continue
        total_hit = event.get("unmitigatedAmount", 0) or event.get("amount", 0)
        if total_hit == 0:
            continue
        ability_name = ability_names.get(ability_id, f"#{ability_id}")
        ts = event.get("timestamp", 0)
        elapsed_ms = ts - fight_start_ms
        time_str = f"{int(elapsed_ms // 60000)}:{int((elapsed_ms % 60000) // 1000):02d}"
        amount_k = f"{total_hit / 1000:.1f}k"
        entry = results.setdefault(pid, {"hits": 0, "big_hits": 0, "details": []})
        entry["hits"] += 1
        max_hp = player_max_hp.get(pid, 0)
        role = player_roles.get(pid, "DPS")
        threshold = 0.50 if role == "Tank" else 0.10
        if max_hp > 0 and total_hit >= max_hp * threshold:
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


# ─── Tracked Spells — loaded from tracked_spells.py ─────────────────────────

# (imported at top of file from tracked_spells.py)


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

WOWHEAD_SCRIPTS = """\
<script>const whTooltips = {colorLinks: true, iconizeLinks: true, renameLinks: true};</script>
<script src="https://wow.zamimg.com/js/tooltips.js"></script>"""


def _ilvl_color(ilvl: int) -> str:
    if ilvl >= 259: return "#ffd700"   # max crafted — gold
    if ilvl >= 246: return "#a335ee"   # upper tier — epic purple
    if ilvl >= 239: return "#1eff00"   # mid tier — uncommon green
    return "#9d9d9d"                   # low — gray


def _build_gear_html(records: list):
    """Build gear table content (header, rows, stats) with mains/alts separation."""
    chars: dict = {}
    for rec in records:
        p = rec["player"]
        if p not in chars:
            _, main_alt = lookup_roster(p)
            chars[p] = {"class": rec["class"], "items": [], "is_alt": main_alt == "Alt"}
        if rec["item_id"] != 0:
            chars[p]["items"].append(rec)

    max_items     = max((len(c["items"]) for c in chars.values()), default=2)
    max_items     = max(max_items, 2)
    total_players = len(chars)
    total_crafted = sum(len(c["items"]) for c in chars.values())
    no_craft      = sum(1 for c in chars.values() if not c["items"])
    total_cols    = 1 + max_items * 5

    def _row(pname, pdata):
        cls_color = CLASS_COLORS.get(pdata["class"], "#ccc")
        row_class = "no-craft" if not pdata["items"] else ""
        row = f'<tr class="{row_class}"><td class="player-cell" style="color:{cls_color}">{escape(pname)}</td>'
        for i in range(max_items):
            div = " divider" if i > 0 else ""
            if i < len(pdata["items"]):
                item      = pdata["items"][i]
                iid       = item["item_id"]
                ilvl      = item["item_level"]
                ic        = _ilvl_color(ilvl)
                spark_cls = "spark-yes" if "Yes" in str(item["spark_used"]) else ""
                spark_txt = "Yes" if "Yes" in str(item["spark_used"]) else "No"
                item_link = f'<a href="https://www.wowhead.com/item={iid}" data-wowhead="item={iid}" target="_blank">#{iid}</a>'
                row += f'<td class="item-cell{div}">{item_link}</td>'
                row += f'<td>{escape(item["slot"])}</td>'
                row += f'<td class="center" style="color:{ic};font-weight:600">{ilvl}</td>'
                row += f'<td>{escape(item["craft_rank"])}</td>'
                row += f'<td class="center {spark_cls}">{spark_txt}</td>'
            else:
                row += f'<td class="empty{div}">—</td>' + '<td class="empty">—</td>' * 4
        row += "</tr>"
        return row

    sort_key = lambda x: (-len(x[1]["items"]), x[0].lower())
    mains = {p: d for p, d in chars.items() if not d["is_alt"]}
    alts  = {p: d for p, d in chars.items() if d["is_alt"]}

    rows = "".join(_row(p, d) for p, d in sorted(mains.items(), key=sort_key))
    if alts:
        rows += f'<tr class="section-sep"><td colspan="{total_cols}">── Alts ──</td></tr>'
        rows += "".join(_row(p, d) for p, d in sorted(alts.items(), key=sort_key))

    gear_header = '<th class="player-header">Player</th>'
    for i in range(max_items):
        div = " divider" if i > 0 else ""
        gear_header += f'<th class="{div}">Item {i+1}</th><th>Slot</th><th>Ilvl</th><th>Rank</th><th>Spark?</th>'

    return gear_header, rows, max_items, total_players, total_crafted, no_craft

CLASS_COLORS = {
    "DeathKnight": "#C41E3A", "DemonHunter": "#A330C9", "Druid": "#FF7C0A",
    "Evoker": "#33937F", "Hunter": "#AAD372", "Mage": "#3FC7EB",
    "Monk": "#00FF98", "Paladin": "#F48CBA", "Priest": "#FFFFFF",
    "Rogue": "#FFF468", "Shaman": "#0070DD", "Warlock": "#8788EE",
    "Warrior": "#C69B6D",
}

# WoW spec IDs → role  (used to detect per-fight spec swaps for accurate table grouping)
SPEC_ROLES: dict = {
    # Death Knight
    250: "Tank",   251: "DPS",    252: "DPS",
    # Demon Hunter
    577: "DPS",    581: "Tank",
    # Druid
    102: "DPS",    103: "DPS",    104: "Tank",   105: "Healer",
    # Evoker
    1467: "DPS",   1468: "Healer", 1473: "DPS",
    # Hunter
    253: "DPS",    254: "DPS",    255: "DPS",
    # Mage
    62: "DPS",     63: "DPS",     64: "DPS",
    # Monk
    268: "Tank",   269: "DPS",    270: "Healer",
    # Paladin
    65: "Healer",  66: "Tank",    70: "DPS",
    # Priest
    256: "Healer", 257: "Healer", 258: "DPS",
    # Rogue
    259: "DPS",    260: "DPS",    261: "DPS",
    # Shaman
    262: "DPS",    263: "DPS",    264: "Healer",
    # Warlock
    265: "DPS",    266: "DPS",    267: "DPS",
    # Warrior
    71: "DPS",     72: "DPS",     73: "Tank",
}


def _build_boss_html(boss_data: dict, actor_lookup: dict, id_prefix: str = "0", wipe_data: dict = None) -> dict:
    """Build HTML for each boss tab. Returns {boss_name: html_string}."""
    ROLE_ORDER = {"Tank": 0, "Healer": 1, "DPS": 2}
    wipe_data  = wipe_data or {}
    results      = {}

    for boss_idx, (boss_name, fights) in enumerate(boss_data.items()):
        table_id       = f"boss-tbl-{id_prefix}-{boss_idx}"
        boss_name_base = boss_name.rsplit(" (", 1)[0]
        mech_defs      = BOSS_MECHANICS.get(boss_name_base, [])
        has_interrupts = boss_name_base in BOSS_HAS_INTERRUPTS
        total_cols     = 8 + len(mech_defs) + (1 if has_interrupts else 0)

        all_pids_set = set()
        for fight in fights:
            all_pids_set.update(fight.get("deaths", {}).keys())
            all_pids_set.update(fight.get("all_player_ids", set()))

        def _get_role(pid, fight=None):
            """Return the role for a player in a specific fight, falling back to roster."""
            char_name = actor_lookup.get(pid, {}).get("name", "")
            pname, _  = lookup_roster(char_name)
            if fight is not None:
                spec_role = fight.get("spec_roles", {}).get(pid)
                if spec_role:
                    return spec_role
            return PLAYER_ROLES.get(pname, "DPS")

        def pid_sort(pid):
            char_name = actor_lookup.get(pid, {}).get("name", "")
            pname, _  = lookup_roster(char_name)
            # Use spec from first available fight for stable sort order
            all_fight_list = list(fights) + list(wipe_data.get(boss_name, []))
            spec_role = next(
                (f.get("spec_roles", {}).get(pid) for f in all_fight_list if f.get("spec_roles", {}).get(pid)),
                None
            )
            role = spec_role or PLAYER_ROLES.get(pname, "DPS")
            return (ROLE_ORDER.get(role, 2), pname.lower())

        sorted_pids = sorted(all_pids_set, key=pid_sort)

        def _hfmt(n):
            return (f"{n/1_000_000_000:.1f}B" if n >= 1_000_000_000
                    else (f"{n/1_000_000:.1f}M" if n >= 1_000_000 else f"{n//1000}k"))

        def _overview_row(label, fight):
            sp_dmg   = sum(v.get("total", 0) if isinstance(v, dict) else 0 for v in fight.get("uptime_map", {}).values())
            sp_heal  = sum(v for v in fight.get("healing_map", {}).values() if isinstance(v, int))
            sp_taken = sum(v for v in fight.get("dmg_taken", {}).values() if isinstance(v, int))
            sp_dur_s = fight.get("fight_dur_ms", 0) / 1000 or 1
            sp_dps   = sp_dmg / sp_dur_s
            sp_m, sp_s = divmod(int(sp_dur_s), 60)
            inner = (f'<span class="sov-title">{label}</span>'
                     f'<span class="sov-dur">{sp_m}:{sp_s:02d}</span>'
                     f'<span class="sov-item"><span class="sov-lbl">⚡ Raid DPS</span> <b>{_hfmt(sp_dps)}</b></span>'
                     f'<span class="sov-item"><span class="sov-lbl">⚔ Dmg Done</span> <b>{_hfmt(sp_dmg) if sp_dmg else "—"}</b></span>'
                     f'<span class="sov-item"><span class="sov-lbl">💚 Healing</span> <b>{_hfmt(sp_heal) if sp_heal else "—"}</b></span>'
                     f'<span class="sov-item"><span class="sov-lbl">🛡 Dmg Taken</span> <b>{_hfmt(sp_taken) if sp_taken else "—"}</b></span>')
            return f'<tr class="split-ov-row"><td colspan="{total_cols}"><div class="split-ov-bar">{inner}</div></td></tr>'

        # ── Per-pull banner builder (frontal + avoid for any list of fights) ──
        import json as _json

        def _build_banners(fights_list: list) -> str:
            """Build Frontal Failures + Avoidable Hits HTML for a given list of fights."""
            out = ""
            # Frontal failures
            all_frontal = [f for fight in fights_list for f in fight.get("frontal_failures", [])]
            if all_frontal:
                rows = ""
                for f in all_frontal:
                    others_str = ", ".join(escape(p) for p in f["others"])
                    rows += (f'<div class="ff-item">'
                             f'<span class="ff-time">{escape(f["time_str"])}</span> '
                             f'<span class="ff-label">{escape(f["label"])}</span> → '
                             f'<span class="ff-primary">{escape(f["primary_player"])}</span>'
                             + (f' <span class="ff-sep">||</span> <span class="ff-players">{others_str}</span>' if others_str else '')
                             + '</div>')
                out += (f'<div class="frontal-failures">'
                        f'<div class="ff-header" onclick="this.classList.toggle(\'open\');this.nextElementSibling.classList.toggle(\'open\')">'
                        f'<span class="ff-title">⚠ Frontal Failures</span><span class="ff-chevron">▼</span></div>'
                        f'<div class="ff-body open">{rows}</div></div>')
            # Avoidable hits
            avoid_defs_local = [m for m in mech_defs if m.get("type") == "avoid"]
            if avoid_defs_local:
                avoid_labels = {m["label"] for m in avoid_defs_local}
                label_hits: dict = {}
                for fight in fights_list:
                    for pid, label_counts in fight.get("mechanics_data", {}).items():
                        for lbl, cnt in label_counts.items():
                            if lbl in avoid_labels:
                                label_hits.setdefault(lbl, {})[pid] = label_hits.setdefault(lbl, {}).get(pid, 0) + cnt
                if label_hits:
                    rows = ""
                    for m in avoid_defs_local:
                        lbl = m["label"]
                        if lbl not in label_hits:
                            continue
                        parts = " · ".join(
                            f'<span class="av-player">{escape(actor_lookup.get(pid, {}).get("name", f"?{pid}"))}</span>'
                            f'<span class="av-count"> ×{cnt}</span>'
                            for pid, cnt in sorted(label_hits[lbl].items(), key=lambda x: -x[1])
                        )
                        rows += f'<div class="av-item"><span class="av-label">{escape(lbl)}</span> → {parts}</div>'
                    if rows:
                        out += (f'<div class="avoid-failures">'
                                f'<div class="av-header" onclick="this.classList.toggle(\'open\');this.nextElementSibling.classList.toggle(\'open\')">'
                                f'<span class="av-title">⚡ Avoidable Hits</span><span class="av-chevron">▼</span></div>'
                                f'<div class="av-body open">{rows}</div></div>')
            return out

        def _build_chart(fight: dict, chart_id: str) -> str:
            """Build a 3-line canvas timeline chart for a single fight."""
            td = fight.get("timeline_data") if fight else None
            if not td or not any(td.get(k) for k in ("dps", "taken", "heal")):
                return ""
            dps_j   = _json.dumps(td.get("dps",   []))
            taken_j = _json.dumps(td.get("taken", []))
            heal_j  = _json.dumps(td.get("heal",  []))
            return (
                f'<div class="timeline-wrap">'
                f'<canvas class="timeline-canvas" id="tc-{chart_id}" height="140"'
                f' data-dps=\'{dps_j}\' data-taken=\'{taken_j}\' data-heal=\'{heal_j}\'></canvas>'
                f'<div class="timeline-legend">'
                f'<span class="tl-dot" style="background:#ffb74d"></span><span class="tl-lbl">DPS Done</span>'
                f'<span class="tl-dot" style="background:#e57373"></span><span class="tl-lbl">Dmg Taken</span>'
                f'<span class="tl-dot" style="background:#81c784"></span><span class="tl-lbl">Healing</span>'
                f'</div>'
                f'</div>'
            )

        # ── Shared table renderer (kills + wipes) ──
        def _render_table(fights_list, tbl_id, pids=None):
            if pids is None:
                pids = sorted_pids

            def _sth(col_idx, lbl, css=""):
                return (f'<th class="{css}" data-sortable="1" style="cursor:pointer;user-select:none"'
                        f' onclick="sortBossTable(\'{tbl_id}\',{col_idx},this)">'
                        f'{lbl} <span class="sort-arrow">▼</span></th>')

            t  = f'<div class="table-wrap"><table id="{tbl_id}" class="detail-col-hidden"><thead><tr>'
            t += '<th style="width:10px"></th>'
            t += _sth(1, 'Player', 'player-header')
            t += _sth(2, 'Char')
            t += _sth(3, 'Deaths', 'death-h')
            t += '<th class="death-h">Killed by (time)</th>'
            t += _sth(5, 'Dmg Taken', 'dmg-h')
            t += _sth(6, 'Uptime %', 'uptime-h')
            if has_interrupts:
                t += _sth(7, 'Interrupts', 'interrupt-h')
            for mi, m in enumerate(mech_defs):
                mci = (8 if has_interrupts else 7) + mi
                css = "mech-soak-h" if m["type"] == "soak" else "mech-bad-h"
                tip = escape(m.get("name", m["label"])).replace("'", "&#39;")
                t += (f'<th class="{css}" data-sortable="1" style="cursor:pointer;user-select:none"'
                      f' data-htip=\'{tip}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()"'
                      f' onclick="sortBossTable(\'{tbl_id}\',{mci},this)">'
                      f'{escape(m["label"])} <span class="sort-arrow">▼</span></th>')
            t += '<th>Notes</th></tr></thead><tbody>'

            for fi, fight in enumerate(fights_list, 1):
                row_lbl = fight.get("_row_label") or f"Split {fi}"
                t += _overview_row(row_lbl, fight)
                current_role = None
                for pid in pids:
                    deaths_map = fight.get("deaths", {})
                    if pid not in fight.get("all_player_ids", set()) and pid not in deaths_map:
                        continue
                    actor     = actor_lookup.get(pid, {})
                    char_name = actor.get("name", f"ID-{pid}")
                    cls       = actor.get("subType", "Unknown")
                    cls_color = CLASS_COLORS.get(cls, "#ccc")
                    pname, _  = lookup_roster(char_name)
                    role      = _get_role(pid, fight)

                    if role != current_role:
                        current_role = role
                        t += f'<tr class="role-sep" data-role="{current_role}"><td colspan="{total_cols}">── {current_role}s ──</td></tr>'

                    dmg_taken_map = fight.get("dmg_taken", {})
                    uptime_map    = fight.get("uptime_map", {})
                    interrupts    = fight.get("interrupts", {})
                    fight_dur     = fight.get("fight_dur_ms", 0)
                    fight_mechs   = fight.get("mechanics_data", {})

                    death_list = deaths_map.get(pid, [])
                    d_raw      = dmg_taken_map.get(pid, 0)
                    dmg_str    = (f"{d_raw/1_000_000:.1f}M" if d_raw >= 1_000_000
                                  else (f"{d_raw/1000:.0f}k" if d_raw > 0 else "—"))
                    active     = (uptime_map.get(pid, {}).get("activeTime", 0)
                                  if isinstance(uptime_map.get(pid), dict) else 0)
                    uptime_str = (f"{min(active / fight_dur * 100, 100):.0f}%"
                                  if fight_dur > 0 and active > 0 else "—")
                    int_count  = interrupts.get(pid, 0)
                    int_str    = str(int_count) if int_count > 0 else "—"
                    int_style  = (' style="background:#1A3A1A"' if int_count > 0
                                  else (' style="background:#5D1A1A"' if has_interrupts else ""))

                    death_count  = len(death_list)
                    killed_str   = ("<br>".join(
                        f'{escape(d["ability"])} @ {escape(d["time"])} ({d.get("fight_pct", 0)}%)'
                        for d in death_list) or "—")

                    death_tip_html = ""
                    if death_list:
                        parts = []
                        for d in death_list:
                            tl    = d.get("timeline", [])
                            block = f"<b style='color:#e57373'>☠ {escape(d['ability'])}</b> @ {escape(d['time'])} ({d.get('fight_pct',0)}%)"
                            if tl:
                                trows = []
                                for tt in tl:
                                    ik    = abs(tt["sec_before"]) < 0.05
                                    col   = "#ff6b6b" if ik else "#aaa"
                                    lbl   = " ← killing blow" if ik else f"-{tt['sec_before']:.1f}s"
                                    trows.append(f"<span style='color:{col}'>{escape(tt['ability'])}: {escape(tt['amount_k'])} &nbsp;<span style='color:#666;font-size:11px'>{lbl}</span></span>")
                                block += "<br>" + "<br>".join(trows)
                            parts.append(block)
                        death_tip_html = ("<hr style='border-color:#333;margin:6px 0'>").join(parts)

                    row_cls = "boss-death-row" if death_count > 0 else ""
                    t += f'<tr class="{row_cls}" data-role="{role}" data-class="{cls}" data-player="{escape(pname.lower())}">'
                    t += f'<td style="background:{cls_color};width:4px;padding:0;min-width:4px"></td>'
                    t += f'<td class="player-cell"><span class="pname">{escape(pname)}</span></td>'
                    t += f'<td><span class="cname" style="color:{cls_color}">{escape(char_name)}</span></td>'
                    if death_count > 0 and death_tip_html:
                        tip_attr = death_tip_html.replace("'", "&#39;")
                        t += f'<td class="center death-h" style="cursor:help" data-htip=\'{tip_attr}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()"><span class="death-num">{death_count}</span></td>'
                        t += f'<td class="death-h" data-htip=\'{tip_attr}\' onmouseenter="showHTip(this)" onmouseleave="hideHTip()" style="cursor:help">{killed_str}</td>'
                    else:
                        t += f'<td class="center death-h">{"<span class=death-num>" + str(death_count) + "</span>" if death_count > 0 else "—"}</td>'
                        t += f'<td class="death-h">{killed_str}</td>'
                    t += f'<td class="center dmg-h">{dmg_str}</td>'
                    t += f'<td class="center uptime-h">{uptime_str}</td>'
                    if has_interrupts:
                        t += f'<td class="center interrupt-h"{int_style}>{int_str}</td>'
                    pid_mechs = fight_mechs.get(pid, {})
                    for m in mech_defs:
                        cnt = pid_mechs.get(m["label"], 0)
                        if cnt:
                            bg = "#1A3D1A" if m["type"] == "soak" else "#5D1A1A"
                            t += f'<td class="center" style="background:{bg}">{cnt}</td>'
                        else:
                            t += '<td class="center">—</td>'
                    t += '<td></td></tr>'

            t += '</tbody></table></div>'
            return t

        # ── Pull selector + panes ──
        boss_wipes = wipe_data.get(boss_name, [])  # oldest → newest

        # Build kill pane content: banners + table + chart (use first kill for chart)
        kill_banners = _build_banners(fights)
        kill_table   = _render_table(fights, table_id)
        kill_chart   = _build_chart(fights[0] if fights else {}, f"{table_id}-kill")
        kill_content = kill_banners + kill_table + kill_chart

        if boss_wipes:
            total_wipes = len(boss_wipes)
            # Annotate wipes with display labels (oldest = Wipe 1)
            for wi, w in enumerate(boss_wipes):
                bpct  = w.get("boss_pct", 0)
                dur_s = w.get("fight_dur_ms", 0) // 1000
                wm, ws = divmod(dur_s, 60)
                w["_row_label"] = f"💀 Wipe {wi + 1} — {bpct}% boss HP · {wm}:{ws:02d}"

            kill_dur_s = fights[0].get("fight_dur_ms", 0) // 1000 if fights else 0
            km, ks     = divmod(kill_dur_s, 60)
            pull_btns  = f'<button class="pull-btn active" onclick="switchPull(this,\'{table_id}\')">⚔ Kill · {km}:{ks:02d}</button>'

            wipe_panes = ""
            for wi, w in enumerate(reversed(boss_wipes)):  # newest first in selector
                actual_wipe_num = total_wipes - wi          # 1-based, newest = total_wipes
                bpct    = w.get("boss_pct", 0)
                dur_s   = w.get("fight_dur_ms", 0) // 1000
                wm, ws  = divmod(dur_s, 60)
                wtbl_id = f"wipe-{table_id}-{wi}"
                pull_btns += f'<button class="pull-btn wipe" onclick="switchPull(this,\'{wtbl_id}\')">Wipe {actual_wipe_num} · {bpct}% · {wm}:{ws:02d}</button>'
                wipe_pids = sorted(
                    all_pids_set | {p for p in (set(w.get("all_player_ids", [])) | set(w.get("deaths", {}).keys())) if p in actor_lookup},
                    key=pid_sort
                )
                wipe_banners = _build_banners([w])
                wipe_table   = _render_table([w], wtbl_id, pids=wipe_pids)
                wipe_chart   = _build_chart(w, wtbl_id)
                wipe_panes += f'<div class="pull-pane" id="pane-{wtbl_id}">{wipe_banners}{wipe_table}{wipe_chart}</div>'

            pull_selector = f'<div class="pull-selector">{pull_btns}</div>'
            kill_pane     = f'<div class="pull-pane active" id="pane-{table_id}">{kill_content}</div>'
            results[boss_name] = pull_selector + kill_pane + wipe_panes
        else:
            results[boss_name] = kill_content

    return results


_FILTER_BAR_CSS = """/* ── Filter bar ── */
.filter-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 16px; padding: 10px 14px; background: #0d1525; border-radius: 8px; border: 1px solid #2a2a4a; }
.filter-label { font-size: 11px; color: #555; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; margin-right: 2px; }
.filter-tag { background: #16213e; color: #888; border: 1px solid #2a2a4a; padding: 3px 10px; border-radius: 20px; font-size: 12px; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
.filter-tag:hover { color: #bbb; border-color: #4a5568; }
.filter-tag.active { background: #2a3a6a; color: #7289DA; border-color: #7289DA; font-weight: 600; }
.filter-divider { width: 1px; height: 20px; background: #2a2a4a; margin: 0 4px; flex-shrink: 0; }
.chip-input-area { position: relative; }
.chip-input-wrap { display: inline-flex; align-items: center; flex-wrap: wrap; gap: 4px; background: #16213e; border: 1px solid #2a2a4a; border-radius: 6px; padding: 4px 8px; min-width: 180px; cursor: text; }
.chip-input-wrap:focus-within { border-color: #7289DA; }
.chip { display: inline-flex; align-items: center; gap: 4px; background: #2a3a6a; color: #a0b4ff; border-radius: 20px; padding: 2px 8px 2px 10px; font-size: 12px; }
.chip-remove { background: none; border: none; color: #7289DA; cursor: pointer; font-size: 13px; padding: 0; line-height: 1; }
.chip-remove:hover { color: #e57373; }
.chip-text-input { background: none; border: none; color: #e0e0e0; font-size: 12px; outline: none; min-width: 90px; padding: 1px 0; }
.chip-text-input::placeholder { color: #555; }
.chip-dropdown { position: absolute; top: 100%; left: 0; min-width: 180px; background: #1a2233; border: 1px solid #4a5568; border-radius: 6px; z-index: 100; margin-top: 2px; box-shadow: 0 4px 12px #0008; }
.chip-dd-item { padding: 6px 12px; font-size: 12px; color: #e0e0e0; cursor: pointer; }
.chip-dd-item:hover { background: #2a3a6a; color: #a0b4ff; }
.filter-clear { background: none; border: 1px solid #3a2a2a; color: #666; padding: 3px 10px; border-radius: 20px; font-size: 11px; cursor: pointer; white-space: nowrap; }
.filter-clear:hover { color: #e57373; border-color: #e57373; }
"""

_FILTER_BAR_JS = """
const _CLASS_COLORS_FB = {DeathKnight:'#C41E3A',DemonHunter:'#A330C9',Druid:'#FF7C0A',Evoker:'#33937F',Hunter:'#AAD372',Mage:'#3FC7EB',Monk:'#00FF98',Paladin:'#F48CBA',Priest:'#FFFFFF',Rogue:'#FFF468',Shaman:'#0070DD',Warlock:'#8788EE',Warrior:'#C69B3A'};
function buildClassTags(splitId) {
  const tab = document.getElementById('tab-' + splitId);
  const ct = document.getElementById('class-tags-' + splitId);
  if (!tab || !ct || ct.children.length) return;
  const classes = [...new Set([...tab.querySelectorAll('tr[data-class]')].map(r => r.dataset.class))].sort();
  classes.forEach(cls => {
    const btn = document.createElement('button');
    btn.className = 'filter-tag'; btn.dataset.ftype = 'class'; btn.dataset.fval = cls;
    btn.style.cssText = 'border-left:3px solid ' + (_CLASS_COLORS_FB[cls] || '#aaa') + ';';
    btn.textContent = cls;
    btn.onclick = () => { btn.classList.toggle('active'); applyFilters(splitId); };
    ct.appendChild(btn);
  });
}
function toggleTag(btn, splitId) { btn.classList.toggle('active'); applyFilters(splitId); }
function addChip(splitId, value) {
  value = value.trim().toLowerCase();
  if (!value) return;
  const ct = document.getElementById('chips-' + splitId);
  if (!ct || [...ct.querySelectorAll('.chip')].some(c => c.dataset.val === value)) return;
  const chip = document.createElement('span');
  chip.className = 'chip'; chip.dataset.val = value;
  chip.textContent = value;
  const rm = document.createElement('button');
  rm.className = 'chip-remove'; rm.title = 'Remove'; rm.textContent = '\u00d7';
  rm.onclick = () => { chip.remove(); applyFilters(splitId); };
  chip.appendChild(rm);
  ct.appendChild(chip); applyFilters(splitId);
}
function chipKey(e, splitId) {
  const inp = e.target;
  if ((e.key === 'Enter' || e.key === ',') && inp.value.trim()) {
    e.preventDefault(); addChip(splitId, inp.value); inp.value = ''; hideDropdown(splitId);
  } else if (e.key === 'Backspace' && !inp.value) {
    const ct = document.getElementById('chips-' + splitId);
    if (ct && ct.lastElementChild) { ct.lastElementChild.remove(); applyFilters(splitId); }
  } else if (e.key === 'Escape') { hideDropdown(splitId); }
}
function chipSuggest(inp, splitId) {
  const val = inp.value.toLowerCase().trim();
  const dd = document.getElementById('chip-dd-' + splitId);
  if (!val || !dd) { if (dd) dd.style.display = 'none'; return; }
  const tab = document.getElementById('tab-' + splitId);
  if (!tab) return;
  const names = [...new Set([...tab.querySelectorAll('tr[data-player]')].map(r => r.dataset.player))].filter(n => n.includes(val)).slice(0, 8);
  if (!names.length) { dd.style.display = 'none'; return; }
  dd.innerHTML = '';
  names.forEach(n => {
    const item = document.createElement('div');
    item.className = 'chip-dd-item'; item.textContent = n;
    item.onmousedown = (ev) => { ev.preventDefault(); addChip(splitId, n); inp.value = ''; hideDropdown(splitId); };
    dd.appendChild(item);
  });
  dd.style.display = 'block';
}
function hideDropdown(splitId) { const dd = document.getElementById('chip-dd-' + splitId); if (dd) dd.style.display = 'none'; }
function clearFilters(splitId) {
  const fb = document.getElementById('fb-' + splitId);
  if (fb) fb.querySelectorAll('.filter-tag.active').forEach(b => b.classList.remove('active'));
  const ct = document.getElementById('chips-' + splitId);
  if (ct) ct.innerHTML = '';
  const inp = document.getElementById('chip-in-' + splitId);
  if (inp) inp.value = '';
  hideDropdown(splitId); applyFilters(splitId);
}
function applyFilters(splitId) {
  const fb = document.getElementById('fb-' + splitId);
  const tab = document.getElementById('tab-' + splitId);
  if (!fb || !tab) return;
  const activeRoles   = new Set([...fb.querySelectorAll('[data-ftype="role"].active')].map(b => b.dataset.fval));
  const activeClasses = new Set([...fb.querySelectorAll('[data-ftype="class"].active')].map(b => b.dataset.fval));
  const ct = document.getElementById('chips-' + splitId);
  const activeNames = ct ? [...ct.querySelectorAll('.chip')].map(c => c.dataset.val) : [];
  const hasFilters = activeRoles.size || activeClasses.size || activeNames.length;
  tab.querySelectorAll('table').forEach(tbl => {
    const tbody = tbl.tBodies[0]; if (!tbody) return;
    const roleVis = {};
    [...tbody.rows].forEach(row => {
      if (row.classList.contains('split-ov-row') || row.classList.contains('section-sep') || row.classList.contains('role-sep')) return;
      const role = row.dataset.role; if (!role) return;
      let show = true;
      if (hasFilters) {
        const roleOk  = !activeRoles.size   || activeRoles.has(role);
        const classOk = !activeClasses.size || activeClasses.has(row.dataset.class);
        const nameOk  = !activeNames.length || activeNames.some(n => (row.dataset.player || '').includes(n));
        show = roleOk && classOk && nameOk;
      }
      row.style.display = show ? '' : 'none';
      if (show) roleVis[role] = true;
    });
    [...tbody.rows].forEach(row => {
      if (row.classList.contains('role-sep'))
        row.style.display = (!hasFilters || roleVis[row.dataset.role]) ? '' : 'none';
    });
  });
}
"""


def _filter_bar_html(split_id: str) -> str:
    si = split_id.replace("'", "\\'")
    return (
        f'<div class="filter-bar" id="fb-{split_id}">'
        f'<span class="filter-label">Role</span>'
        f'<button class="filter-tag" data-ftype="role" data-fval="Tank" onclick="toggleTag(this,\'{si}\')">🛡 Tank</button>'
        f'<button class="filter-tag" data-ftype="role" data-fval="Healer" onclick="toggleTag(this,\'{si}\')">💚 Healer</button>'
        f'<button class="filter-tag" data-ftype="role" data-fval="DPS" onclick="toggleTag(this,\'{si}\')">⚔ DPS</button>'
        f'<div class="filter-divider"></div>'
        f'<span class="filter-label">Class</span>'
        f'<span class="class-tags" id="class-tags-{split_id}"></span>'
        f'<div class="filter-divider"></div>'
        f'<div class="chip-input-area">'
        f'<div class="chip-input-wrap" onclick="this.querySelector(\'.chip-text-input\').focus()">'
        f'<span class="chips-container" id="chips-{split_id}"></span>'
        f'<input class="chip-text-input" id="chip-in-{split_id}" placeholder="Player name\u2026"'
        f' onkeydown="chipKey(event,\'{si}\')" oninput="chipSuggest(this,\'{si}\')"'
        f' onblur="setTimeout(()=>hideDropdown(\'{si}\'),150)">'
        f'</div>'
        f'<div class="chip-dropdown" id="chip-dd-{split_id}" style="display:none"></div>'
        f'</div>'
        f'<button class="filter-clear" onclick="clearFilters(\'{si}\')">✕ Clear all</button>'
        f'</div>'
    )


def write_raid_html(day_data: dict, output_path: str) -> None:
    """Write a single raid day as a standalone HTML page (Gear + Split tabs)."""
    ri         = day_data["report_info"]
    rc         = day_data["report_code"]
    date_str   = datetime.fromtimestamp(ri["startTime"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d") if ri.get("startTime") else ""
    diff_label = day_data.get("difficulty", "")
    title      = ri.get("title", rc) + (f" — {diff_label}" if diff_label else "")

    # ── Gear section (Normal or no-diff only) ──
    # Heroic/Mythic: no gear. Player-group split pages: no gear (lives in gear_normal.html).
    show_gear = diff_label in ("Normal", "") and not day_data.get("player_split")

    gear_tab_html = ""
    gear_tab_btn  = ""

    if show_gear:
        gear_header, gear_rows, _, total_players, total_crafted, no_craft = _build_gear_html(day_data["records"])
        gear_inner = f"""
  <div class="stats">
    <div class="stat-box"><div class="num">{total_players}</div><div class="label">Players</div></div>
    <div class="stat-box"><div class="num">{total_crafted}</div><div class="label">Crafted Items</div></div>
    <div class="stat-box warn"><div class="num">{no_craft}</div><div class="label">No Crafted Gear</div></div>
  </div>
  <div class="search-box"><input type="text" placeholder="Search player..." onkeyup="filterGear(this)"></div>
  <div class="table-wrap">
    <table id="gear-table"><thead><tr>{gear_header}</tr></thead><tbody>{gear_rows}</tbody></table>
  </div>"""
        gear_tab_btn  = '<button class="tab-btn" onclick="switchTab(\'gear\',this)">⚙ Gear</button>\n'
        gear_tab_html = f'<div id="tab-gear" class="tab-content">{gear_inner}\n</div>'

    # ── Build Boss tabs (one tab per boss) ──
    actor_lookup = {a["id"]: a for a in day_data["actors"]}

    # Merge all fights per boss (split_num is preserved inside each fight dict)
    merged_boss_data: dict = {}
    for boss_name, fights in day_data["boss_data"].items():
        merged_boss_data.setdefault(boss_name, []).extend(fights)

    boss_htmls = _build_boss_html(merged_boss_data, actor_lookup, id_prefix="b",
                                  wipe_data=day_data.get("wipe_data", {}))

    tab_buttons = ""
    split_divs  = ""
    for bi, (boss_name, boss_html) in enumerate(boss_htmls.items()):
        boss_display = boss_name.rsplit(" (", 1)[0]  # strip difficulty suffix
        tab_id  = f"boss-{bi}"
        active  = "active" if bi == 0 else ""
        fb_html = _filter_bar_html(tab_id)
        tab_buttons += f'<button class="tab-btn {active}" onclick="switchTab(\'{tab_id}\',this)">{escape(boss_display)}</button>\n'
        split_divs  += f'<div id="tab-{tab_id}" class="tab-content {active}">{fb_html}{boss_html}</div>\n'

    tab_buttons += gear_tab_btn  # gear always last

    wcl_link = f'<a href="https://www.warcraftlogs.com/reports/{rc}" target="_blank" class="wcl-link">View on WarcraftLogs ↗</a>'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{escape(title)} — {date_str} — Raid Audit</title>
{WOWHEAD_SCRIPTS}
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #1a1a2e; color: #e0e0e0; font-family: -apple-system, 'Segoe UI', sans-serif; padding: 20px; }}
h1 {{ color: #7289DA; font-size: 22px; margin-bottom: 4px; }}
.breadcrumb {{ font-size: 12px; color: #555; margin-bottom: 14px; }}
.breadcrumb a {{ color: #7289DA; text-decoration: none; }}
.breadcrumb a:hover {{ text-decoration: underline; }}
.raid-meta {{ display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }}
.raid-date {{ color: #666; font-size: 13px; }}
.wcl-link {{ color: #7289DA; font-size: 12px; text-decoration: none; }}
.wcl-link:hover {{ text-decoration: underline; }}
/* ── Tabs ── */
.tab-bar {{ display: flex; gap: 4px; margin-bottom: 20px; border-bottom: 2px solid #2a2a4a; flex-wrap: wrap; }}
.tab-btn {{ background: #16213e; color: #888; border: none; padding: 10px 20px; border-radius: 8px 8px 0 0; font-size: 13px; font-weight: 600; cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -2px; transition: all 0.15s; }}
.tab-btn:hover {{ color: #bbb; background: #1e2d50; }}
.tab-btn.active {{ color: #7289DA; background: #1a1a2e; border-bottom: 2px solid #7289DA; }}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}
/* ── Boss section ── */
.boss-section {{ margin-bottom: 4px; }}
.boss-section-title {{ color: #a0b4ff; font-size: 14px; font-weight: 700; margin-bottom: 0; padding: 7px 12px; background: #111827; border-radius: 4px; border-left: 3px solid #7289DA; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; }}
.boss-section-title:hover {{ background: #1a2236; }}
.boss-toggle-arrow {{ font-size: 12px; transition: transform 0.2s; }}
.boss-section-title.collapsed .boss-toggle-arrow {{ transform: rotate(-90deg); }}
.boss-section-body {{ padding-top: 10px; margin-bottom: 20px; }}
.boss-section-body.collapsed {{ display: none; }}
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
/* ── Tables ── */
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
td, th {{ border-left: 1px solid rgba(255,255,255,0.07); }}
td:first-child, th:first-child {{ border-left: none; }}
/* ── Gear ── */
.no-craft {{ background: rgba(255,160,0,0.05); }}
.no-craft .player-cell {{ background: rgba(80,50,0,0.4); }}
.spark-yes {{ color: #4caf50; font-weight: bold; }}
.item-cell a {{ color: #a48cff; text-decoration: none; }}
tr.section-sep td {{ color: #555; font-size: 11px; padding: 4px 10px; background: #0d1117; letter-spacing: 0.5px; }}
.item-cell a:hover {{ text-decoration: underline; }}
/* ── Roles ── */
.role-sep td {{ background: #111827; color: #7289DA; font-size: 11px; font-weight: bold; padding: 4px 10px; }}
.pname {{ font-weight: bold; }}
.cname {{ color: #888; font-size: 11px; }}
/* ── Split overview row ── */
.split-ov-row td {{ padding: 0 !important; border: none; background: #0d1525; }}
.split-ov-bar {{ display: flex; align-items: center; flex-wrap: wrap; gap: 16px; padding: 10px 14px; border-top: 2px solid #2a3a6a; border-bottom: 1px solid #1e2a4a; }}
.sov-title {{ color: #a0b4ff; font-weight: 700; font-size: 14px; min-width: 60px; }}
.sov-dur {{ color: #666; font-size: 12px; }}
.sov-item {{ font-size: 12px; color: #ccc; }}
.sov-lbl {{ color: #666; margin-right: 4px; }}
/* ── Boss columns ── */
.death-h {{ color: #e57373; }}
.dmg-h {{ color: #ffb74d; }}
.heal-h {{ color: #81c784; }}
.parse-h {{ color: #E5CC80; }}
.uptime-h {{ color: #81c784; }}
.interrupt-h {{ color: #64b5f6; }}
.avoid-h {{ color: #ce93d8; }}
.detail-h {{ color: #7289DA; white-space: normal; min-width: 30px; cursor: pointer; user-select: none; }}
.detail-h:hover {{ color: #a0b4ff; }}
.mech-bad-h {{ color: #ff7070; }}
.mech-soak-h {{ color: #66bb6a; }}
.detail-cell {{ white-space: normal; min-width: 180px; font-size: 12px; color: #ccc; }}
.detail-col-hidden .detail-h span.detail-content,
.detail-col-hidden .detail-cell {{ display: none; }}
.detail-col-hidden .detail-h {{ min-width: 0; }}
.boss-death-row td {{ background: rgba(61,16,16,0.4); }}
.boss-death-row .player-cell {{ background: rgba(80,10,10,0.6); }}
.death-num {{ color: #e57373; font-weight: bold; }}
.bighit-row {{ color: #ff8a65; font-weight: bold; }}
/* ── Frontal failures ── */
.frontal-failures {{ margin-bottom: 12px; background: rgba(229,115,115,0.08); border: 1px solid rgba(229,115,115,0.3); border-radius: 8px; overflow: hidden; }}
.ff-header {{ display: flex; align-items: center; justify-content: space-between; padding: 8px 14px; cursor: pointer; user-select: none; }}
.ff-header:hover {{ background: rgba(229,115,115,0.12); }}
.ff-title {{ color: #e57373; font-weight: 700; font-size: 13px; }}
.ff-chevron {{ color: #e57373; font-size: 11px; transition: transform 0.2s; }}
.ff-header.open .ff-chevron {{ transform: rotate(180deg); }}
.ff-body {{ display: none; padding: 4px 0 8px; }}
.ff-body.open {{ display: block; }}
.ff-item {{ padding: 3px 14px; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.04); }}
.ff-split {{ color: #888; margin-right: 6px; }}
.ff-time {{ color: #a0b4ff; font-weight: 600; margin-right: 6px; }}
.ff-label {{ color: #e57373; font-weight: 600; margin-right: 4px; }}
.ff-primary {{ color: #f4a742; font-weight: 700; }}
.ff-sep {{ color: #555; margin: 0 6px; }}
.ff-players {{ color: #e06c6c; }}
/* ── Avoidable Hits ── */
.avoid-failures {{ margin-bottom: 10px; background: rgba(244,167,66,0.07); border: 1px solid rgba(244,167,66,0.3); border-radius: 8px; overflow: hidden; }}
.av-header {{ display: flex; align-items: center; justify-content: space-between; padding: 8px 14px; cursor: pointer; user-select: none; }}
.av-header:hover {{ background: rgba(244,167,66,0.12); }}
.av-title {{ color: #f4a742; font-weight: 700; font-size: 13px; }}
.av-chevron {{ color: #f4a742; font-size: 11px; transition: transform 0.2s; }}
.av-header.open .av-chevron {{ transform: rotate(180deg); }}
.av-body {{ display: none; padding: 4px 0 8px; }}
.av-body.open {{ display: block; }}
.av-item {{ padding: 3px 14px; font-size: 12px; border-top: 1px solid rgba(255,255,255,0.04); }}
.av-label {{ color: #f4a742; font-weight: 600; margin-right: 4px; }}
.av-player {{ color: #e0e0e0; }}
.av-count {{ color: #e06c6c; font-weight: 700; margin-right: 8px; }}
/* ── Sort arrows ── */
.sort-arrow {{ opacity: 0.6; font-size: 11px; margin-left: 4px; }}
th[data-sortable]:hover .sort-arrow {{ opacity: 1; }}
/* ── Pull selector ── */
.pull-selector {{ display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 10px; }}
.pull-btn {{ background: #16213e; color: #888; border: 1px solid #2a2a4a; padding: 4px 12px; border-radius: 4px; font-size: 12px; cursor: pointer; transition: all 0.15s; }}
.pull-btn:hover {{ color: #bbb; }}
.pull-btn.active {{ background: #1a3a2a; color: #4caf50; border-color: #4caf50; font-weight: 600; }}
.pull-btn.wipe.active {{ background: #3a1a1a; color: #e57373; border-color: #e57373; }}
.pull-pane {{ display: none; }}
.pull-pane.active {{ display: block; }}
{_FILTER_BAR_CSS}
/* ── Timeline chart ── */
.timeline-wrap {{ margin: 12px 0 4px; }}
.timeline-canvas {{ display: block; width: 100%; height: 140px; background: #0d1525; border-radius: 6px; cursor: default; }}
.timeline-legend {{ display: flex; align-items: center; gap: 14px; padding: 4px 2px; font-size: 11px; color: #888; }}
.tl-dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; }}
.tl-lbl {{ margin-left: -8px; }}
</style>
<script>const whTooltips = {{colorLinks: true, iconizeLinks: true, iconSize: 'small'}};</script>
<script src="https://wow.zamimg.com/js/tooltips.js"></script>
</head>
<body>
<div class="breadcrumb"><a href="index.html">← Overview</a> / {escape(title)}</div>
<h1>{escape(title)}</h1>
<div class="raid-meta">
  <span class="raid-date">{date_str}</span>
  {wcl_link}
</div>
<div class="tab-bar">{tab_buttons}</div>

{gear_tab_html}
<!-- ── SPLIT TABS ── -->
{split_divs}
<script>
function switchTabByName(name) {{
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  const el = document.getElementById('tab-' + name);
  if (!el) return;
  el.classList.add('active');
  const btn = [...document.querySelectorAll('.tab-btn')].find(b => (b.getAttribute('onclick') || '').includes("'" + name + "'"));
  if (btn) btn.classList.add('active');
}}
function switchTab(name, btn) {{ switchTabByName(name); }}
function renderTimeline(canvas) {{
  const raw = (key) => {{ try {{ return JSON.parse(canvas.dataset[key] || '[]'); }} catch(e) {{ return []; }} }};
  const dps = raw('dps'), taken = raw('taken'), heal = raw('heal');
  const all = dps.concat(taken, heal);
  if (!all.length) return;
  const W = canvas.offsetWidth || canvas.parentElement?.offsetWidth || 700;
  const H = 140;
  canvas.width  = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  const PAD = {{top:14, right:12, bottom:22, left:54}};
  const cW = W - PAD.left - PAD.right;
  const cH = H - PAD.top  - PAD.bottom;
  const maxT = Math.max(...all.map(p => p[0]));
  const maxV = Math.max(...all.map(p => p[1]));
  if (!maxT || !maxV) return;
  const xS = t => PAD.left + (t / maxT) * cW;
  const yS = v => PAD.top  + cH - (v / maxV) * cH;
  ctx.fillStyle = '#0d1525';
  ctx.fillRect(0, 0, W, H);
  ctx.strokeStyle = '#1a2540'; ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {{
    const y = PAD.top + (cH / 4) * i;
    ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(W - PAD.right, y); ctx.stroke();
  }}
  const fmt = v => v >= 1e9 ? (v/1e9).toFixed(1)+'B' : v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v >= 1e3 ? Math.round(v/1e3)+'k' : String(Math.round(v));
  ctx.fillStyle = '#455'; ctx.font = '10px sans-serif'; ctx.textAlign = 'right';
  for (let i = 0; i <= 2; i++) {{
    const y = PAD.top + (cH / 2) * i;
    ctx.fillText(fmt(maxV * (1 - i/2)), PAD.left - 4, y + 4);
  }}
  ctx.textAlign = 'center'; ctx.fillStyle = '#445';
  const steps = Math.min(6, Math.floor(maxT / 30));
  for (let i = 0; i <= steps; i++) {{
    const t = maxT * i / steps;
    const m = Math.floor(t / 60), s = Math.round(t % 60);
    ctx.fillText(`${{m}}:${{s.toString().padStart(2,'0')}}`, xS(t), H - 4);
  }}
  const series = [{{d: dps, c: '#ffb74d'}}, {{d: taken, c: '#e57373'}}, {{d: heal, c: '#81c784'}}];
  series.forEach(({{d, c}}) => {{
    if (!d.length) return;
    ctx.strokeStyle = c; ctx.lineWidth = 1.5; ctx.lineJoin = 'round';
    ctx.beginPath();
    d.forEach(([t, v], i) => {{ const x = xS(t), y = yS(v); i ? ctx.lineTo(x, y) : ctx.moveTo(x, y); }});
    ctx.stroke();
  }});
}}
function renderPaneCharts(paneEl) {{
  if (!paneEl) return;
  paneEl.querySelectorAll('.timeline-canvas').forEach(c => {{ if (c.offsetWidth > 0) renderTimeline(c); }});
}}
window.addEventListener('DOMContentLoaded', () => {{
  const h = location.hash.replace('#', '');
  if (h) switchTabByName(h);
  document.querySelectorAll('[id^="class-tags-"]').forEach(el => buildClassTags(el.id.replace('class-tags-', '')));
  // Render charts in initially-active pull panes
  document.querySelectorAll('.pull-pane.active').forEach(renderPaneCharts);
}});
function switchPull(btn, paneId) {{
  const body = btn.closest('.boss-section-body') || btn.closest('.pull-selector').parentElement;
  body.querySelectorAll('.pull-btn').forEach(b => b.classList.remove('active'));
  body.querySelectorAll('.pull-pane').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  const pane = document.getElementById('pane-' + paneId);
  if (pane) {{ pane.classList.add('active'); renderPaneCharts(pane); }}
}}
function toggleBoss(titleEl) {{
  titleEl.classList.toggle('collapsed');
  titleEl.nextElementSibling.classList.toggle('collapsed');
}}
function filterGear(input) {{
  const filter = input.value.toLowerCase();
  for (let row of document.getElementById('gear-table').tBodies[0].rows)
    row.style.display = row.cells[0].textContent.toLowerCase().includes(filter) ? '' : 'none';
}}
function toggleDetails(tableId, btn) {{
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  const hidden = tbl.classList.toggle('detail-col-hidden');
  btn.textContent = hidden ? '▶ Details' : '▼ Details';
}}
const _htip = document.createElement('div');
_htip.id = 'htip';
Object.assign(_htip.style, {{ position:'fixed', display:'none', pointerEvents:'none', zIndex:'9999',
  background:'#1a2233', border:'1px solid #4a5568', borderRadius:'8px',
  padding:'10px 14px', fontSize:'12px', color:'#e0e0e0', maxWidth:'340px', lineHeight:'1.6', boxShadow:'0 4px 16px #0008' }});
document.body.appendChild(_htip);
document.addEventListener('mousemove', e => {{
  if (_htip.style.display !== 'none') {{
    const x = e.clientX + 16, y = e.clientY - 10;
    _htip.style.left = (x + _htip.offsetWidth > window.innerWidth ? e.clientX - _htip.offsetWidth - 8 : x) + 'px';
    _htip.style.top  = (y + _htip.offsetHeight > window.innerHeight ? e.clientY - _htip.offsetHeight - 8 : y) + 'px';
  }}
}});
function showHTip(el) {{ const d = el.getAttribute('data-htip'); if (!d) return; _htip.innerHTML = d; _htip.style.display = 'block'; }}
function hideHTip() {{ _htip.style.display = 'none'; }}
function sortBossTable(tableId, colIdx, thEl) {{
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  const dir = thEl.dataset.sortDir === 'desc' ? 'asc' : 'desc';
  tbl.querySelectorAll('th[data-sortable]').forEach(th => {{ th.dataset.sortDir = ''; const a = th.querySelector('.sort-arrow'); if (a) a.textContent = '▼'; }});
  thEl.dataset.sortDir = dir;
  const arrow = thEl.querySelector('.sort-arrow');
  if (arrow) arrow.textContent = dir === 'desc' ? '▼' : '▲';
  const tbody = tbl.tBodies[0];
  const rows = Array.from(tbody.rows);
  const segments = []; let cur = null;
  for (const row of rows) {{
    if (row.classList.contains('role-sep')) {{ if (cur) segments.push(cur); cur = {{ sep: row, rows: [] }}; }}
    else {{ if (!cur) cur = {{ sep: null, rows: [] }}; cur.rows.push(row); }}
  }}
  if (cur) segments.push(cur);
  function parseVal(cell) {{
    if (!cell) return null;
    const text = cell.textContent.trim().replace(/⇅|▲|▼/g, '').trim();
    if (!text || text === '—') return null;
    const m = text.replace('%','').match(/^([\d.]+)\s*([MkKBbG]?)$/);
    if (m) {{ let n = parseFloat(m[1]); const s = m[2].toUpperCase();
      if (s==='B') n*=1e9; else if (s==='M') n*=1e6; else if (s==='K') n*=1e3; return n; }}
    return text.toLowerCase();
  }}
  for (const seg of segments) {{
    seg.rows.sort((a, b) => {{
      const va = parseVal(a.cells[colIdx]), vb = parseVal(b.cells[colIdx]);
      if (va===null && vb===null) return 0; if (va===null) return 1; if (vb===null) return -1;
      if (typeof va==='number' && typeof vb==='number') return dir==='asc' ? va-vb : vb-va;
      if (va<vb) return dir==='asc'?-1:1; if (va>vb) return dir==='asc'?1:-1; return 0;
    }});
  }}
  for (const seg of segments) {{ if (seg.sep) tbody.appendChild(seg.sep); for (const row of seg.rows) tbody.appendChild(row); }}
}}
{_FILTER_BAR_JS}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[OK] Raid page saved: {output_path}")


def write_index_html(days_data: list, output_path: str, guild_name: str = "") -> None:
    """Write the overview index.html.
    Normal splits are expanded into individual cards (numbered Run 1, Run 2, …).
    Heroic/Mythic get one card each.
    """
    diff_cls_map = {"Normal": "badge-normal", "Heroic": "badge-heroic", "Mythic": "badge-mythic"}

    # ── Build card descriptors (oldest→newest for NM numbering) ──
    sorted_asc = sorted(days_data, key=lambda d: d["report_info"].get("startTime", 0))
    nm_cards   = []   # Normal-split cards only, in chronological order
    all_cards  = []   # every card

    for day_data in sorted_asc:
        ri   = day_data["report_info"]
        rc   = day_data["report_code"]
        bd   = day_data["boss_data"]
        diff = day_data.get("difficulty", "")
        date_str = datetime.fromtimestamp(ri["startTime"] / 1000, tz=timezone.utc).strftime("%B %d, %Y") if ri.get("startTime") else ""
        title    = ri.get("title", rc)
        filename = _raid_filename(day_data)
        if diff in ("Heroic", "Mythic"):
            label = "HC" if diff == "Heroic" else "Mythic"
            all_cards.append({"diff": diff, "label": label, "date_str": date_str, "title": title,
                               "filename": filename, "boss_count": len(bd), "nm_idx": None, "day_data": day_data})
        else:
            card = {"diff": "Normal", "label": "NM", "date_str": date_str, "title": title,
                    "filename": filename, "boss_count": len(bd), "nm_idx": None, "day_data": day_data}
            all_cards.append(card)
            nm_cards.append(card)

    # Sequential NM numbering ordered by date then player_split
    total_nm = len(nm_cards)
    nm_ordered = sorted(nm_cards, key=lambda c: (
        c["day_data"]["report_info"].get("startTime", 0),
        c["day_data"].get("player_split", 0)
    ))
    for i, c in enumerate(nm_ordered):
        c["nm_idx"] = i + 1
        c["label"]  = f"NM · Run {i + 1}"

    # Display order: newest first; within same date Heroic > Normal; higher split before lower
    diff_order = {"Mythic": 0, "Heroic": 1, "Normal": 2}
    all_cards.sort(key=lambda c: (-c["day_data"]["report_info"].get("startTime", 0),
                                   diff_order.get(c["diff"], 3),
                                   -c["day_data"].get("player_split", 0)))

    def _card(c: dict, is_first: bool) -> str:
        diff_cls     = diff_cls_map.get(c["diff"], "badge-normal")
        recent_badge = '<span class="badge-recent">Most Recent</span>' if is_first else ""
        sub = f'<span>Run {c["nm_idx"]} of {total_nm}</span>' if c["nm_idx"] else ""
        return (
            f'<a class="raid-card" href="{c["filename"]}">'
            f'  <div class="card-top">'
            f'    <div class="card-date">{escape(c["date_str"])}</div>'
            f'    <div class="card-badges"><span class="badge-diff {diff_cls}">{escape(c["label"])}</span>{recent_badge}</div>'
            f'  </div>'
            f'  <div class="card-title">{escape(c["title"])}</div>'
            f'  <div class="card-stats">'
            f'    <span>&#9876; {c["boss_count"]} boss{"es" if c["boss_count"] != 1 else ""} killed</span>'
            f'    {sub}'
            f'  </div>'
            f'  <div class="card-link">View Report &#8594;</div>'
            f'</a>'
        )

    cards_html  = "\n".join(_card(c, i == 0) for i, c in enumerate(all_cards))
    total_cards = len(all_cards)
    gear_link   = '<a href="gear_normal.html" class="gear-link" title="Crafted Gear — All Normal Runs">⚙</a>' if nm_cards else ""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Raid Audit — {escape(guild_name)}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #1a1a2e; color: #e0e0e0; font-family: -apple-system, 'Segoe UI', sans-serif; padding: 28px 24px; }}
.page-header {{ display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }}
h1 {{ color: #7289DA; font-size: 24px; }}
.gear-link {{ color: #4caf50; font-size: 20px; text-decoration: none; line-height: 1; }}
.gear-link:hover {{ color: #81c784; }}
.subtitle {{ color: #555; font-size: 13px; margin-bottom: 28px; }}
.raid-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 18px; }}
.raid-card {{ background: #16213e; border: 1px solid #2a2a4a; border-radius: 12px; padding: 20px; text-decoration: none; color: inherit; display: block; transition: border-color 0.15s, transform 0.15s; }}
.raid-card:hover {{ border-color: #7289DA; transform: translateY(-2px); }}
.card-top {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }}
.card-date {{ color: #7289DA; font-size: 12px; font-weight: 600; }}
.card-badges {{ display: flex; gap: 6px; flex-wrap: wrap; justify-content: flex-end; }}
.badge-diff {{ display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; }}
.badge-normal  {{ background: #1a3a1a; color: #4caf50; }}
.badge-heroic  {{ background: #1a2a4a; color: #64b5f6; }}
.badge-mythic  {{ background: #2a1a3a; color: #c77dff; }}
.badge-recent  {{ background: #2a2416; color: #E5CC80; display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; }}
.card-title {{ color: #e0e0e0; font-size: 16px; font-weight: 700; margin-bottom: 12px; }}
.card-stats {{ display: flex; gap: 14px; color: #888; font-size: 12px; margin-bottom: 14px; flex-wrap: wrap; }}
.card-link {{ color: #7289DA; font-size: 13px; font-weight: 600; }}
.raid-card:hover .card-link {{ text-decoration: underline; }}
</style>
</head>
<body>
<div class="page-header"><h1>Raid Audit — {escape(guild_name)}</h1>{gear_link}</div>
<div class="subtitle">{total_cards} entr{"ies" if total_cards != 1 else "y"} tracked</div>
<div class="raid-grid">
{cards_html}
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[OK] Index page saved: {output_path}")


def write_gear_html(days_data: list, output_path: str, guild_name: str = "") -> None:
    """Write a consolidated gear page for all Normal runs."""
    normal_days = [d for d in days_data if d.get("difficulty", "") in ("Normal", "")]
    if not normal_days:
        return

    # Merge + deduplicate records across all normal days
    seen: set = set()
    merged: list = []
    for day_data in sorted(normal_days, key=lambda d: d["report_info"].get("startTime", 0)):
        for rec in day_data["records"]:
            key = (rec["player"], rec["item_id"])
            if key not in seen:
                seen.add(key)
                merged.append(rec)

    gear_header, gear_rows, _, total_players, total_crafted, no_craft = _build_gear_html(merged)
    nm_runs = len(normal_days)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Crafted Gear — {escape(guild_name)}</title>
{WOWHEAD_SCRIPTS}
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #1a1a2e; color: #e0e0e0; font-family: -apple-system, 'Segoe UI', sans-serif; padding: 28px 24px; }}
.page-header {{ display: flex; align-items: center; gap: 16px; margin-bottom: 4px; }}
h1 {{ color: #4caf50; font-size: 22px; }}
.back-link {{ color: #7289DA; font-size: 13px; text-decoration: none; }}
.back-link:hover {{ text-decoration: underline; }}
.subtitle {{ color: #555; font-size: 13px; margin-bottom: 24px; }}
.stats {{ display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }}
.stat-box {{ background: #16213e; border-radius: 8px; padding: 10px 18px; }}
.stat-box .num {{ font-size: 24px; font-weight: bold; color: #4caf50; }}
.stat-box .label {{ font-size: 11px; color: #888; text-transform: uppercase; }}
.stat-box.warn .num {{ color: #ffc107; }}
.search-box {{ margin: 0 0 12px; }}
.search-box input {{ background: #16213e; border: 1px solid #2a2a4a; color: #e0e0e0; padding: 7px 14px; border-radius: 6px; width: 280px; font-size: 13px; }}
.search-box input::placeholder {{ color: #555; }}
.table-wrap {{ overflow-x: auto; border-radius: 8px; }}
table {{ border-collapse: collapse; min-width: 100%; }}
th {{ background: #16213e; color: #4caf50; padding: 8px 10px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; position: sticky; top: 0; z-index: 1; }}
th.player-header {{ position: sticky; left: 0; z-index: 2; background: #16213e; min-width: 140px; }}
td {{ padding: 6px 10px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; white-space: nowrap; }}
td.player-cell {{ position: sticky; left: 0; background: #0d1117; font-weight: 600; z-index: 1; }}
tr:hover td {{ background: rgba(255,255,255,0.04); }}
tr.no-craft td {{ opacity: 0.45; }}
td.empty {{ color: #333; }}
td.item-cell a {{ color: #a0b4ff; text-decoration: none; }}
td.item-cell a:hover {{ text-decoration: underline; }}
td.divider {{ border-left: 1px solid #2a2a4a; }}
td.center {{ text-align: center; }}
.spark-yes {{ color: #ffc107; font-weight: 700; }}
tr.section-sep td {{ color: #555; font-size: 11px; padding: 4px 10px; background: #0d1117; letter-spacing: 0.5px; }}
</style>
</head>
<body>
<div class="page-header">
  <h1>⚙ Crafted Gear — All Normal Runs</h1>
  <a href="index.html" class="back-link">← Back to Index</a>
</div>
<div class="subtitle">Compiled across {nm_runs} Normal run{"s" if nm_runs != 1 else ""}</div>
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
<script>
function filterGear(input) {{
  const filter = input.value.toLowerCase();
  for (let row of document.getElementById('gear-table').tBodies[0].rows)
    row.style.display = row.cells[0].textContent.toLowerCase().includes(filter) ? '' : 'none';
}}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[OK] Gear overview saved: {output_path}")


# ─── Roster Mapping (The Rupture) ────────────────────────────────────────────

# Maps character name (lowercase) -> (Player name, "Main" or "Alt")
ROSTER = {}
PLAYER_ROLES = {}  # player_name -> "Tank"/"Healer"/"DPS"


def load_roster(path="roster.txt"):
    """Load player roster from roster.txt. Format: PlayerName | Role | CharName:Main/Alt, ..."""
    global ROSTER, PLAYER_ROLES
    ROSTER = {}
    PLAYER_ROLES = {}
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [p.strip() for p in line.split("|")]
                if len(parts) < 3:
                    continue
                player_name, role, chars_raw = parts[0], parts[1], parts[2]
                PLAYER_ROLES[player_name] = role
                for char_entry in chars_raw.split(","):
                    char_entry = char_entry.strip()
                    if not char_entry:
                        continue
                    if ":" in char_entry:
                        char_name, main_alt = char_entry.rsplit(":", 1)
                    else:
                        char_name, main_alt = char_entry, "Main"
                    ROSTER[char_name.strip().lower()] = (player_name, main_alt.strip())
    except FileNotFoundError:
        print("[WARN] roster.txt not found — player role/name lookup will be unavailable.")


load_roster()


def lookup_roster(char_name: str):
    """Look up a character in the roster. Returns (player_name, 'Main'/'Alt') or (char_name, 'Unknown')."""
    return ROSTER.get(char_name.lower(), (char_name, "Unknown"))


# ─── Boss Mechanic Definitions — loaded from boss_mechanics.py ───────────────

# (imported at top of file from boss_mechanics.py)

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
    """Load credentials from a config file.
    Supports multiple REPORT_URL= lines (same key, repeated).
    Inline comments are stripped: REPORT_URL=https://... # my comment
    """
    config = {}
    report_urls = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, val = line.split("=", 1)
                key = key.strip()
                # Strip inline comments (space + #)
                val = val.split(" #")[0].strip()
                if not val:
                    continue
                if key.startswith("REPORT_URL"):
                    report_urls.append(val)
                else:
                    config[key] = val
    except FileNotFoundError:
        pass
    config["REPORT_URLS"] = report_urls
    return config


def _extract_report_code(url: str) -> str:
    if "warcraftlogs.com" in url:
        parts = url.rstrip("/").split("/")
        return parts[-1].split("#")[0].split("?")[0]
    return url


def _cache_path(report_code: str) -> str:
    os.makedirs("data", exist_ok=True)
    return os.path.join("data", f"raid_{report_code}.json")


def _fix_int_keys(boss_data: dict) -> dict:
    """After JSON load, restore integer actor-ID keys in fields that use actor IDs as keys.
    rankings_map uses char-name strings as keys — excluded intentionally.
    """
    int_key_fields = (
        "uptime_map", "dmg_taken", "healing_map", "interrupts",
        "avoidable_damage", "deaths", "mechanics_data",
    )
    for fights in boss_data.values():
        for fight in fights:
            for field in int_key_fields:
                if field in fight and isinstance(fight[field], dict):
                    fight[field] = {int(k): v for k, v in fight[field].items()}
    return boss_data


def process_report(token: str, report_code: str, fight_input: str = "all") -> dict:
    """Fetch and process one WCL report. Loads from cache if available."""
    cache = _cache_path(report_code)
    if os.path.exists(cache):
        print(f"\n[CACHE] Loading {report_code} from cache (delete data/raid_{report_code}.json to re-fetch).")
        with open(cache, encoding="utf-8") as f:
            result = json.load(f)
        result["boss_data"] = _fix_int_keys(result["boss_data"])
        return result

    print(f"\n[...] Fetching report info for: {report_code}")
    report_info = fetch_report_info(token, report_code)
    guild_name = report_info.get("guild", {}).get("name", "") if report_info.get("guild") else ""
    print(f"[OK] Report: {report_info.get('title', 'Untitled')} ({guild_name})")

    all_raw_fights  = report_info.get("fights", [])
    all_fights      = [f for f in all_raw_fights if f.get("encounterID", 0) > 0 and f.get("kill")]
    all_wipe_fights = [f for f in all_raw_fights if f.get("encounterID", 0) > 0 and not f.get("kill")]
    selected_fights = all_fights

    if all_fights:
        print(f"\nBoss kills found ({len(all_fights)}):")
        for f in all_fights:
            diff = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(f.get("difficulty", 0), f"D{f.get('difficulty','?')}")
            print(f"  [{f['id']:>3}] {f['name']} — {diff} (Kill)")
        if fight_input == "interactive":
            fight_input = input("\nFight IDs to analyze (comma-separated, or 'all'): ").strip()
        if fight_input.lower() != "all" and fight_input != "":
            selected_ids = {int(x.strip()) for x in fight_input.split(",") if x.strip().isdigit()}
            selected_fights = [f for f in all_fights if f["id"] in selected_ids]

    actors = report_info.get("masterData", {}).get("actors", [])
    print(f"[OK] Found {len(actors)} player actors in report.")
    ability_names = {ab["gameID"]: ab["name"] for ab in report_info.get("masterData", {}).get("abilities", [])}
    actor_lookup = {a["id"]: a for a in actors}

    # Gear
    all_combatant_events = []
    fight_spec_roles: dict = {}   # fid → {pid: role}
    for fight in selected_fights:
        fid, fname = fight["id"], fight["name"]
        print(f"  [...] Fetching gear from fight {fid}: {fname}...")
        try:
            events = fetch_combatant_info_events(token, report_code, fid)
            all_combatant_events.extend(events)
            fight_spec_roles[fid] = {
                e["sourceID"]: SPEC_ROLES.get(e.get("specID"), "DPS")
                for e in events if e.get("sourceID") and e.get("specID")
            }
            print(f"         Got {len(events)} players' gear.")
        except Exception as e:
            print(f"    [WARN] Gear fetch failed for fight {fid}: {e}")

    player_max_hp = {}
    for ce in all_combatant_events:
        pid, stamina = ce.get("sourceID"), ce.get("stamina", 0)
        if pid and stamina:
            player_max_hp[pid] = max(player_max_hp.get(pid, 0), stamina * 20)

    print("\n[...] Analyzing crafted gear...")
    records = analyze_players(actors, all_combatant_events)
    crafted_count = sum(1 for r in records if r["item_id"] != 0)
    player_count  = len(set(r["player"] for r in records))
    no_craft_count= sum(1 for r in records if r["item_id"] == 0)
    print(f"[OK] Found {crafted_count} crafted items across {player_count} players.")
    if no_craft_count > 0:
        print(f"[!!] {no_craft_count} player(s) have NO detected crafted gear.")

    # Split detection — group by encounter, order of kills = split number
    encounter_kills = {}
    for fight in selected_fights:
        eid = fight.get("encounterID", 0)
        encounter_kills.setdefault(eid, []).append(fight)
    max_splits = max((len(v) for v in encounter_kills.values()), default=1)
    split_fights = {i + 1: [] for i in range(max_splits)}
    for kills in encounter_kills.values():
        for i, kill in enumerate(sorted(kills, key=lambda f: f["id"])):
            split_fights[i + 1].append(kill)

    print(f"\n[...] Detected {max_splits} split(s).")

    # Cast events per split
    split_data = {}
    for split_num, sfl in split_fights.items():
        split_name = f"Split {split_num}"
        print(f"\n[...] Fetching spell usage for {split_name}...")
        split_fight_data = []
        for fight in sfl:
            fid   = fight["id"]
            fname = fight["name"]
            diff  = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(fight.get("difficulty", 0), "")
            print(f"  [...] Fetching casts for {fname} ({diff})...")
            try:
                cast_events  = fetch_cast_events(token, report_code, fid)
                player_casts = analyze_fight_casts(cast_events, fight.get("startTime", 0), actor_lookup)
                total        = sum(len(v) for v in player_casts.values())
                print(f"         Found {total} tracked spell uses across {len(player_casts)} players.")
                split_fight_data.append({"fight_name": f"{fname} ({diff})", "fight_id": fid, "player_casts": player_casts})
            except Exception as e:
                print(f"    [WARN] Casts fetch failed for fight {fid}: {e}")
        split_data[split_name] = split_fight_data

    # Boss analysis — one fight dict per fight, tagged with split_num
    boss_data = {}
    encounter_kill_count = {}
    print(f"\n[...] Fetching death + damage events per boss...")
    player_roles_map = {pid: PLAYER_ROLES.get(lookup_roster(actor_lookup[pid]["name"])[0], "DPS")
                        for pid in actor_lookup if "name" in actor_lookup[pid]}

    for fight in selected_fights:
        fid  = fight["id"]
        fname= fight["name"]
        eid  = fight.get("encounterID", 0)
        diff = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(fight.get("difficulty", 0), "")
        boss_key = f"{fname} ({diff})"
        encounter_kill_count[eid] = encounter_kill_count.get(eid, 0) + 1
        split_num = encounter_kill_count[eid]
        try:
            death_events     = fetch_death_events(token, report_code, fid)
            damage_events    = fetch_damage_taken_events(token, report_code, fid)
            interrupt_events = fetch_interrupt_events(token, report_code, fid)
            cast_events      = fetch_boss_cast_events(token, report_code, fid)
            fight_start      = fight.get("startTime", 0)
            fight_dur_ms     = fight.get("endTime", 0) - fight_start
            deaths           = analyze_deaths(death_events, fight_start, ability_names,
                                              fight_end_ms=fight_dur_ms, damage_events=damage_events)
            avoidable        = analyze_avoidable_damage(damage_events, actor_lookup, fight_start,
                                                        ability_names, player_max_hp, player_roles_map)
            dmg_taken        = aggregate_damage_taken(damage_events, actor_lookup)
            uptime_map       = fetch_uptime_table(token, report_code, fid)
            healing_map      = fetch_healing_table(token, report_code, fid)
            rankings_map     = fetch_rankings(token, report_code, fid)
            interrupts       = analyze_interrupts(interrupt_events, actor_lookup)
            mech_defs        = BOSS_MECHANICS.get(fname, [])
            mechanics_data   = analyze_boss_mechanics(damage_events, actor_lookup, mech_defs)
            frontal_failures = analyze_frontal_failures(damage_events, actor_lookup, mech_defs, fight_start,
                                                        cast_events=cast_events)
            try:
                timeline_data = fetch_fight_graph(token, report_code, fid)
            except Exception:
                timeline_data = {}
            all_pids = set()
            for sdata in split_data.values():
                for fd in sdata:
                    if fd.get("fight_id") == fid:
                        all_pids.update(fd.get("player_casts", {}).keys())
            all_pids.update(dmg_taken.keys())
            all_pids.update(uptime_map.keys())
            all_pids.update(deaths.keys())
            all_pids = {pid for pid in all_pids if pid in actor_lookup}
            boss_data.setdefault(boss_key, []).append({
                "fight_id": fid, "fight_dur_ms": fight_dur_ms, "split_num": split_num,
                "deaths": deaths, "all_player_ids": all_pids,
                "avoidable_damage": avoidable, "dmg_taken": dmg_taken,
                "uptime_map": uptime_map, "healing_map": healing_map,
                "rankings_map": rankings_map, "interrupts": interrupts,
                "mechanics_data": mechanics_data, "frontal_failures": frontal_failures,
                "timeline_data": timeline_data,
                "spec_roles": fight_spec_roles.get(fid, {}),
            })
            print(f"  [OK] {fname} (Split {split_num}): {len(deaths)} death(s), {sum(interrupts.values())} interrupts.")
        except Exception as e:
            print(f"  [WARN] Could not fetch events for fight {fid}: {e}")

    # ── Wipe analysis ──
    wipe_data: dict = {}
    if all_wipe_fights:
        print(f"\n[...] Fetching wipe events ({len(all_wipe_fights)} wipe(s))...")
        encounter_wipe_count: dict = {}
        for fight in sorted(all_wipe_fights, key=lambda f: f["id"]):
            fid   = fight["id"]
            fname = fight["name"]
            eid   = fight.get("encounterID", 0)
            diff  = {3: "Normal", 4: "Heroic", 5: "Mythic"}.get(fight.get("difficulty", 0), "")
            boss_key = f"{fname} ({diff})"
            encounter_wipe_count[eid] = encounter_wipe_count.get(eid, 0) + 1
            wipe_num     = encounter_wipe_count[eid]
            boss_pct     = fight.get("bossPercentage", 0)
            fight_start  = fight.get("startTime", 0)
            fight_dur_ms = fight.get("endTime", 0) - fight_start
            try:
                death_events     = fetch_death_events(token, report_code, fid)
                damage_events    = fetch_damage_taken_events(token, report_code, fid)
                interrupt_events = fetch_interrupt_events(token, report_code, fid)
                cast_events      = fetch_boss_cast_events(token, report_code, fid)
                deaths           = analyze_deaths(death_events, fight_start, ability_names,
                                                  fight_end_ms=fight_dur_ms, damage_events=damage_events)
                avoidable        = analyze_avoidable_damage(damage_events, actor_lookup, fight_start,
                                                            ability_names, player_max_hp, player_roles_map)
                dmg_taken        = aggregate_damage_taken(damage_events, actor_lookup)
                uptime_map       = fetch_uptime_table(token, report_code, fid)
                healing_map      = fetch_healing_table(token, report_code, fid)
                interrupts       = analyze_interrupts(interrupt_events, actor_lookup)
                mech_defs        = BOSS_MECHANICS.get(fname, [])
                mechanics_data   = analyze_boss_mechanics(damage_events, actor_lookup, mech_defs)
                frontal_failures = analyze_frontal_failures(damage_events, actor_lookup, mech_defs,
                                                            fight_start, cast_events=cast_events)
                try:
                    timeline_data = fetch_fight_graph(token, report_code, fid)
                except Exception:
                    timeline_data = {}
                try:
                    wipe_ci = fetch_combatant_info_events(token, report_code, fid)
                    wipe_spec_roles = {
                        e["sourceID"]: SPEC_ROLES.get(e.get("specID"), "DPS")
                        for e in wipe_ci if e.get("sourceID") and e.get("specID")
                    }
                except Exception:
                    wipe_spec_roles = {}
                all_pids = {pid for pid in (set(dmg_taken) | set(uptime_map) | set(deaths)) if pid in actor_lookup}
                wipe_data.setdefault(boss_key, []).append({
                    "fight_id": fid, "fight_dur_ms": fight_dur_ms,
                    "wipe_num": wipe_num, "boss_pct": boss_pct,
                    "deaths": deaths, "all_player_ids": all_pids,
                    "avoidable_damage": avoidable, "dmg_taken": dmg_taken,
                    "uptime_map": uptime_map, "healing_map": healing_map,
                    "rankings_map": {},  # wipes have no rankings
                    "interrupts": interrupts, "mechanics_data": mechanics_data,
                    "frontal_failures": frontal_failures,
                    "timeline_data": timeline_data,
                    "spec_roles": wipe_spec_roles,
                })
                dur_s = fight_dur_ms // 1000
                print(f"  [OK] {fname} Wipe {wipe_num} ({boss_pct}% boss HP, {dur_s//60}:{dur_s%60:02d}): {len(deaths)} death(s).")
            except Exception as e:
                print(f"  [WARN] Could not fetch wipe events for fight {fid}: {e}")

    result = {
        "report_code": report_code, "report_info": report_info,
        "records": records, "split_data": split_data,
        "boss_data": boss_data, "wipe_data": wipe_data, "actors": actors,
        "max_splits": max_splits,
    }
    cache = _cache_path(report_code)
    with open(cache, "w", encoding="utf-8") as f:
        json.dump(result, f, default=list)
    print(f"[CACHE] Saved to {cache}")
    return result


def split_report_by_difficulty(day_data: dict) -> list:
    """If a report has multiple difficulties, return one day_data per difficulty."""
    diffs_present = [d for d in ("Normal", "Heroic", "Mythic")
                     if any(f"({d})" in k for k in day_data.get("boss_data", {}))]
    if len(diffs_present) <= 1:
        return [day_data]
    results = []
    for diff in diffs_present:
        filtered_boss  = {k: [dict(f, split_num=1) for f in v]
                          for k, v in day_data["boss_data"].items() if f"({diff})" in k}
        filtered_wipes = {k: v for k, v in day_data.get("wipe_data", {}).items() if f"({diff})" in k}
        results.append({**day_data, "boss_data": filtered_boss, "wipe_data": filtered_wipes, "difficulty": diff})
    return results


def split_report_by_player_group(day_data: dict) -> list:
    """If a report has multiple player-group splits, return one day_data per split."""
    bd = day_data.get("boss_data", {})
    split_nums = {f.get("split_num", 1) for fights in bd.values() for f in fights}
    if len(split_nums) <= 1:
        return [day_data]
    results = []
    for snum in sorted(split_nums):
        filtered_boss = {k: [f for f in v if f.get("split_num", 1) == snum]
                         for k, v in bd.items()}
        filtered_boss = {k: v for k, v in filtered_boss.items() if v}
        results.append({**day_data, "boss_data": filtered_boss,
                        "wipe_data": day_data.get("wipe_data", {}), "player_split": snum})
    return results


def _raid_filename(day_data: dict) -> str:
    rc     = day_data["report_code"]
    diff   = day_data.get("difficulty", "")
    psplit = day_data.get("player_split")
    return f"raid_{rc}{'_' + diff.lower() if diff else ''}{'_split' + str(psplit) if psplit else ''}.html"


def main():
    print("=" * 60)
    print("  WarcraftLogs Crafted Gear Audit — Midnight Season 1")
    print("=" * 60)

    config = load_config()

    if config.get("CLIENT_ID") and config.get("CLIENT_SECRET"):
        client_id, client_secret = config["CLIENT_ID"], config["CLIENT_SECRET"]
        print(f"\n[OK] Loaded credentials from wcl_config.txt (Client ID: {client_id[:8]}...)")
    else:
        client_id     = input("\nWarcraftLogs Client ID: ").strip()
        client_secret = input("WarcraftLogs Client Secret: ").strip()

    if not client_id or not client_secret:
        print("[ERROR] Both Client ID and Client Secret are required.")
        sys.exit(1)

    print("\n[...] Authenticating with WarcraftLogs...")
    token = get_access_token(client_id, client_secret)
    print("[OK] Authenticated successfully.")

    report_urls = config.get("REPORT_URLS", [])
    if not report_urls:
        report_urls = [input("\nReport code or URL: ").strip()]
    else:
        print(f"[OK] Loaded {len(report_urls)} report URL(s) from config.")

    report_codes = [_extract_report_code(u) for u in report_urls]
    fight_mode   = "interactive" if len(report_codes) == 1 else "all"

    days_data = []
    for code in report_codes:
        try:
            result = process_report(token, code, fight_input=fight_mode)
            for r in split_report_by_difficulty(result):
                days_data.extend(split_report_by_player_group(r))
        except Exception as e:
            print(f"[ERROR] Failed to process report {code}: {e}")

    if not days_data:
        print("[ERROR] No reports processed successfully.")
        sys.exit(1)

    # Most recent report first
    days_data.sort(key=lambda d: d["report_info"].get("startTime", 0), reverse=True)

    # Guild name (from most recent report)
    ri_first   = days_data[0]["report_info"]
    guild_name = ri_first.get("guild", {}).get("name", "The Rupture") if ri_first.get("guild") else "The Rupture"

    # Per-raid pages
    for day_data in days_data:
        write_raid_html(day_data, _raid_filename(day_data))

    # Overview index
    write_index_html(days_data, "index.html", guild_name=guild_name)
    write_gear_html(days_data, "gear_normal.html", guild_name=guild_name)

    # XLSX: most recent day only
    first = days_data[0]
    write_xlsx(first["records"], first["report_info"], first["report_code"],
               "craft_audit.xlsx", split_data=first["split_data"],
               boss_data=first["boss_data"], actors=first["actors"])

    print(f"\nDone! Open index.html in your browser, or open craft_audit.xlsx.\n")


if __name__ == "__main__":
    main()