# boss_mechanics.py
# ─────────────────────────────────────────────────────────────────────────────
# Edit this file to add new boss abilities or update spell IDs after a patch.
# The main script imports BOSS_MECHANICS and BOSS_HAS_INTERRUPTS from here.
#
# Mechanic types:
#   "frontal"  — detect failure: 2+ friendly players hit in the same cast window
#                shows WHO + WHEN in a red callout per boss
#   "avoid"    — any hit = personal failure (standing in bad, failing to dodge)
#                shows WHO + how many times in an orange callout per boss
#   "dmg_hits" — just track hit counts per player (shown in table columns)
#   "soak"     — hits are good, more = better (shown in green in table)
#
# Optional: "difficulties": ["Heroic", "Mythic"]
#   If omitted, the mechanic applies to all difficulties.
# ─────────────────────────────────────────────────────────────────────────────

BOSS_MECHANICS = {
    "Imperator Averzian": [
        {
            "label": "Shad.Advance",
            "name": "Shadow's Advance: Averzian summons Abyssal Voidshapers. They inflict Shadow damage to players within 10 yards and knock them away.",
            "spell_ids": {1253691},
            "type": "dmg_hits",  # proximity-based, partially unavoidable
        },
        {
            "label": "Void Fall",
            "name": "Void Fall: Averzian knocks back players and rains destruction at several destinations — step out of the impact zones.",
            "spell_ids": {1258883, 1269160},
            "type": "avoid",
        },
        {
            "label": "Obliv.Wrath",
            "name": "Oblivion's Wrath: Averzian launches void lances outward — dodge the lances.",
            "spell_ids": {1260718},
            "type": "avoid",
        },
        {
            "label": "Umbral Col.",
            "name": "Umbral Collapse (SOAK): Averzian collapses void energy around his target. Damage is reduced by the number of players within 10 yards. Stack up to soak! Higher = better.",
            "spell_ids": {1249262},
            "type": "soak",
        },
        {
            "label": "Gnash.Void",
            "name": "Gnashing Void: The Voidmaw's melee attacks inflict Shadow damage every 1 sec for 10 sec. Tank mechanic.",
            "spell_ids": {1255683},
            "type": "dmg_hits",  # tank mechanic
        },
        {
            "label": "Shad.Phalanx",
            "name": "Shadow Phalanx: A column of Averzian's troops march across the field — step out of the march path.",
            "spell_ids": {1284786},
            "type": "avoid",
        },
    ],

    "Vorasius": [
        {
            "label": "Blisterburst",
            "name": "Blisterburst: Exploding Blistercreep — don't stand near it when it dies.",
            "spell_ids": {1259186},
            "type": "avoid",
        },
        {
            "label": "Claw Slam",
            "name": "Shadowclaw Slam: Vorasius slams a giant claw — someone must be in the centre or the whole raid takes massive damage.",
            "spell_ids": {1241808, 1281954, 1281906, 1272328},
            "type": "dmg_hits",  # complex soak mechanic
        },
        {
            "label": "Parasite Exp.",
            "name": "Parasite Expulsion: Globs of dark ichor land on the ground — dodge the impact zones.",
            "spell_ids": {1275558, 1275556},
            "type": "avoid",
        },
        {
            "label": "Void Breath",
            "name": "Void Breath: Vorasius sweeps a deadly frontal beam — stand behind the boss.",
            "spell_ids": {1257607},
            "type": "dmg_hits",
        },
    ],

    "Fallen-King Salhadaar": [
        {
            "label": "Tort.Extract",
            "name": "Torturous Extract: Lingering void energy pool — move out of it.",
            "spell_ids": {1245592},
            "type": "dmg_hits",  # zone left behind, partially unavoidable
        },
        {
            "label": "Umbral Beams",
            "name": "Umbral Beams: Beams of void radiate from Salhadaar — step out of the beams.",
            "spell_ids": {1260030},
            "type": "avoid",
        },
        {
            "label": "Void Exposure",
            "name": "Void Exposure: Triggered by touching a Void Orb — don't touch the orbs.",
            "spell_ids": {1250828},
            "type": "avoid",
        },
        {
            "label": "Twilight Spk.",
            "name": "Twilight Spikes: Dark energy erupts from the ground — step out of the spikes.",
            "spell_ids": {1251213},
            "type": "avoid",
        },
    ],

    "Vaelgor & Ezzorak": [
        {
            "label": "Impale",
            "name": "Impale: Ezzorak slams a 35 yard rear cone — don't stand behind Ezzorak.",
            "spell_ids": {1265152},
            "type": "frontal",
        },
        {
            "label": "Dread Breath",
            "name": "Dread Breath: Vaelgor roars a massive frontal cone toward a targeted player — the target must face it away from the raid.",
            "spell_ids": {1244225, 1255979},
            "type": "frontal",
        },
        {
            "label": "Gloomfield",
            "name": "Gloomfield: Galactic darkness engulfs a location — don't stand in it.",
            "spell_ids": {1245421},
            "type": "avoid",
        },
        {
            "label": "Tail Lash",
            "name": "Tail Lash: Vaelgor knocks away players in a 35 yard rear cone — don't stand behind Vaelgor.",
            "spell_ids": {1264467},
            "type": "frontal",
        },
        {
            "label": "Nullbeam",
            "name": "Nullbeam (SOAK): Vaelgor expels crystalline spacetime in a frontal direction. Stack up in front to weaken the pull! Higher = better.",
            "spell_ids": {1283856, 1262688},
            "type": "soak",
        },
    ],

    "Lightblinded Vanguard": [
        {
            "label": "Final Verdict",
            "name": "Final Verdict: Devastating strike on the current tank target.",
            "spell_ids": {1251812},
            "type": "dmg_hits",  # tank mechanic
        },
        {
            "label": "Divine Toll",
            "name": "Divine Toll: Volley of holy shields — dodge them.",
            "spell_ids": {1248652},
            "type": "avoid",
        },
        {
            "label": "Exec.Sentence",
            "name": "Execution Sentence (SOAK): Commander attempts to execute his target — stack within 8 yards to split the damage. Higher = better.",
            "spell_ids": {1249024},
            "type": "soak",
        },
        {
            "label": "Trampled",
            "name": "Trampled: Senn charges forward on her elekk — get out of the charge path.",
            "spell_ids": {1249135},
            "type": "frontal",
        },
        {
            "label": "Div.Hammer",
            "name": "Divine Hammer: Holy hammers spiral outward from Execution Sentence — dodge them.",
            "spell_ids": {1249047},
            "type": "avoid",
        },
    ],

    "Crown of the Cosmos": [
        {
            "label": "Silverstrike",
            "name": "Silverstrike Arrow/Ricochet: Alleria marks a player and fires an arrow — being hit without a Void effect is avoidable.",
            "spell_ids": {1233649, 1237729},
            "type": "dmg_hits",  # hit events
        },
        {
            "label": "Void Stacks",
            "name": "Void Stacks (1233602): Debuff applied by Silverstrike — grants ability to remove add shields. Track stack count.",
            "spell_ids": {1233602},
            "type": "dmg_hits",  # debuff application — used for stack tracking
        },
        {
            "label": "P3 Circle",
            "name": "Void Circle (P3): 3 random players (ranged → melee → tank) each carry a circle and must leave in correct order.",
            "spell_ids": {1233887},
            "type": "avoid",  # leaving out of order = wipe
        },
        {
            "label": "Brstng Empty.",
            "name": "Bursting Emptiness: Backlash of magic from each obelisk.",
            "spell_ids": {1255378},
            "type": "dmg_hits",  # semi-unavoidable raid damage
        },
        {
            "label": "Void Remnants",
            "name": "Void Remnants: Celestial body crashes near a player — get away from the impact zone.",
            "spell_ids": {1233826, 1242553},
            "type": "dmg_hits",  # targeted at player, hard to fully avoid
        },
        {
            "label": "Singularity",
            "name": "Singularity Eruption: Gravity pockets surge from Alleria — dodge the impact zones.",
            "spell_ids": {1235631},
            "type": "avoid",
        },
        {
            "label": "Dev.Cosmos",
            "name": "Devouring Cosmos: Alleria unleashes the cosmos — get out of the zone immediately.",
            "spell_ids": {1238882},
            "type": "avoid",
        },
        {
            "label": "Grav.Collapse",
            "name": "Gravity Collapse: Knocks the target upwards and increases their Physical damage taken by 300% for 12 sec.",
            "spell_ids": {1239095},
            "type": "dmg_hits",  # targeted mechanic
        },
    ],

    "Chimaerus, the Undreamt God": [
        {
            "label": "Alndust Ess.",
            "name": "Alndust Essence: Nature damage pool on the ground — don't stand in it.",
            "spell_ids": {1245919},
            "type": "avoid",
        },
        {
            "label": "Alndust Uph.",
            "name": "Alndust Upheaval (SOAK): Chimaerus tears a hole in Reality — stack within 10 yards to split the damage. Higher = better.",
            "spell_ids": {1262305, 1246827},
            "type": "soak",
        },
        {
            "label": "Disc.Roar",
            "name": "Discordant Roar: Unavoidable raid-wide physical damage — track for healing reference.",
            "spell_ids": {1249207},
            "type": "dmg_hits",  # unavoidable raid damage
        },
        {
            "label": "Rift Emerg.",
            "name": "Rift Emergence: Unavoidable raid-wide nature damage — track for healing reference.",
            "spell_ids": {1258610},
            "type": "dmg_hits",  # unavoidable raid damage
        },
    ],
}

# Bosses where 0 interrupts = red (boss has meaningful interruptible abilities)
BOSS_HAS_INTERRUPTS = {
    "Imperator Averzian",
    "Fallen-King Salhadaar",
    "Vaelgor & Ezzorak",
    "Lightblinded Vanguard",
    "Crown of the Cosmos",
    "Chimaerus, the Undreamt God",
}
