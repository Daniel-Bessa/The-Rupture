# boss_mechanics.py
# ─────────────────────────────────────────────────────────────────────────────
# Edit this file to add new boss abilities or update spell IDs after a patch.
# The main script imports BOSS_MECHANICS and BOSS_HAS_INTERRUPTS from here.
#
# Mechanic types:
#   "frontal"  — detect failure: 2+ friendly players hit in the same cast window
#   "dmg_hits" — track how many times each player was hit (bad, shown in red)
#   "soak"     — track how many times each player soaked (good, shown in green)
#
# Optional: "difficulties": ["Heroic", "Mythic"]
#   If omitted, the mechanic applies to all difficulties.
#   Use this to add Heroic/Mythic-only abilities.
# ─────────────────────────────────────────────────────────────────────────────

BOSS_MECHANICS = {
    "Imperator Averzian": [
        {
            "label": "Shad.Advance",
            "name": "Shadow's Advance: Averzian summons Abyssal Voidshapers. They inflict Shadow damage to players within 10 yards and knock them away.",
            "spell_ids": {1253691},
            "type": "dmg_hits",
        },
        {
            "label": "Void Fall",
            "name": "Void Fall: Averzian knocks back players and rains destruction at several destinations, inflicting Shadow damage to players within 7 yards of the impact locations.",
            "spell_ids": {1258883, 1269160},
            "type": "dmg_hits",
        },
        {
            "label": "Obliv.Wrath",
            "name": "Oblivion's Wrath: Averzian launches void lances outward, inflicting Shadow damage to players in their path and knocking them back.",
            "spell_ids": {1260718},
            "type": "dmg_hits",
        },
        {
            "label": "Umbral Col.",
            "name": "Umbral Collapse (SOAK): Averzian collapses void energy around his target. Damage is reduced by the number of players within 10 yards. Stack up to soak! Higher = better.",
            "spell_ids": {1249262},
            "type": "soak",
        },
        {
            "label": "Gnash.Void",
            "name": "Gnashing Void: The Voidmaw's melee attacks inflict Shadow damage every 1 sec for 10 sec. This effect stacks.",
            "spell_ids": {1255683},
            "type": "dmg_hits",
        },
        {
            "label": "Shad.Phalanx",
            "name": "Shadow Phalanx: A column of Averzian's troops march across the field, inflicting Shadow damage every 1 sec to players in their path.",
            "spell_ids": {1284786},
            "type": "dmg_hits",
        },
    ],

    "Vorasius": [
        {
            "label": "Blisterburst",
            "name": "Blisterburst: The Blistercreep explodes upon death, inflicting Shadow damage to players within 8 yards, increasing their damage taken by 100% for 30 sec and knocking them away.",
            "spell_ids": {1259186},
            "type": "dmg_hits",
        },
        {
            "label": "Claw Slam",
            "name": "Shadowclaw Slam: Vorasius slams a giant claw into the ground. If the central impact fails to hit at least 1 player, Vorasius inflicts massive Shadow damage to all players instead.",
            "spell_ids": {1241808, 1281954, 1281906, 1272328},
            "type": "dmg_hits",
        },
        {
            "label": "Parasite Exp.",
            "name": "Parasite Expulsion: Vorasius shakes off parasitic Blistercreeps, spraying the battlefield with globs of dark ichor that inflict Shadow damage to players within 3 yards upon impact.",
            "spell_ids": {1275558, 1275556},
            "type": "dmg_hits",
        },
        {
            "label": "Void Breath",
            "name": "Void Breath: Vorasius sweeps a deadly beam across the battlefield, inflicting massive Shadow damage to all players in front of him and Shadow damage every 1 sec to players caught in its path.",
            "spell_ids": {1257607},
            "type": "dmg_hits",
        },
    ],

    "Fallen-King Salhadaar": [
        {
            "label": "Tort.Extract",
            "name": "Torturous Extract: Lingering void energy that inflicts Shadow damage every 1 sec to players within it.",
            "spell_ids": {1245592},
            "type": "dmg_hits",
        },
        {
            "label": "Umbral Beams",
            "name": "Umbral Beams: Beams of pure void radiate from Salhadaar, inflicting Shadow damage every 0.3 sec to players within them.",
            "spell_ids": {1260030},
            "type": "dmg_hits",
        },
        {
            "label": "Void Exposure",
            "name": "Void Exposure: Exposes players to an excessive amount of void energy, inflicting Shadow damage every 1 sec to players within it. Triggered by touching a Void Orb.",
            "spell_ids": {1250828},
            "type": "dmg_hits",
        },
        {
            "label": "Twilight Spk.",
            "name": "Twilight Spikes: Dark energy erupts from the ground, inflicting Shadow damage every 2 sec to players within it.",
            "spell_ids": {1251213},
            "type": "dmg_hits",
        },
    ],

    "Vaelgor & Ezzorak": [
        {
            "label": "Impale",
            "name": "Impale: Ezzorak slams targets within a 35 yard rear cone, bleeding for Physical damage plus additional Physical damage every 1 sec and stunning for 3 sec. Occurs immediately after Rakfang.",
            "spell_ids": {1265152},
            "type": "dmg_hits",
        },
        {
            "label": "Dread Breath",
            "name": "Dread Breath: Vaelgor roars toward a targeted player, fearing players in a massive frontal cone. Inflicts Shadow damage and an additional Shadow damage every 3 sec, fearing them for 21 sec.",
            "spell_ids": {1244225, 1255979},
            "type": "frontal",
        },
        {
            "label": "Gloomfield",
            "name": "Gloomfield: Galactic emptiness engulfs a massive location in darkness for 2.5 min, inflicting Shadow damage every 0.5 sec and reducing movement speed by 75%.",
            "spell_ids": {1245421},
            "type": "dmg_hits",
        },
        {
            "label": "Tail Lash",
            "name": "Tail Lash: Vaelgor knocks away players within a 35 yard rear cone, bleeding for Physical damage plus additional Physical damage every 0.5 sec for 4 sec.",
            "spell_ids": {1264467},
            "type": "dmg_hits",
        },
        {
            "label": "Nullbeam",
            "name": "Nullbeam (SOAK): Vaelgor expels crystalline spacetime in a frontal direction. Nullzone's pull magnitude weakens as Nullbeam stacks, up to 12 times. Higher = better.",
            "spell_ids": {1283856, 1262688},
            "type": "soak",
        },
    ],

    "Lightblinded Vanguard": [
        {
            "label": "Final Verdict",
            "name": "Final Verdict: Lightblood unleashes a devastating strike on his current target that inflicts Holy damage.",
            "spell_ids": {1251812},
            "type": "dmg_hits",
        },
        {
            "label": "Divine Toll",
            "name": "Divine Toll: Bellamy unleashes a volley of holy shields every 2 sec for 18 sec. Shields that hit players inflict Holy damage and silence them for 6 sec.",
            "spell_ids": {1248652},
            "type": "dmg_hits",
        },
        {
            "label": "Exec.Sentence",
            "name": "Execution Sentence (SOAK): Commander Venel Lightblood attempts to execute his target — damage is split evenly between players within 8 yards. Stack up! Higher = better.",
            "spell_ids": {1249024},
            "type": "soak",
        },
        {
            "label": "Trampled",
            "name": "Trampled: Senn charges forward on her mighty elekk, inflicting Physical damage to players in her path.",
            "spell_ids": {1249135},
            "type": "dmg_hits",
        },
        {
            "label": "Div.Hammer",
            "name": "Divine Hammer: Holy hammers spiral outward from Execution Sentence, inflicting Holy damage to players in their path.",
            "spell_ids": {1249047},
            "type": "dmg_hits",
        },
    ],

    "Crown of the Cosmos": [
        {
            "label": "Silverstrike",
            "name": "Silverstrike Arrow/Ricochet: Alleria marks a player, firing a silver-lined arrow that inflicts Arcane damage in a line and removes Void effects from players struck. Being hit without a Void effect is avoidable.",
            "spell_ids": {1233649, 1237729},
            "type": "dmg_hits",
        },
        {
            "label": "Brstng Empty.",
            "name": "Bursting Emptiness: A backlash of magic blasts out from each obelisk, inflicting Shadow damage to all players struck.",
            "spell_ids": {1255378},
            "type": "dmg_hits",
        },
        {
            "label": "Void Remnants",
            "name": "Void Remnants: Alleria calls down a celestial body near a player. This energy crashes into the ground and inflicts Shadow damage to all players.",
            "spell_ids": {1233826, 1242553},
            "type": "dmg_hits",
        },
        {
            "label": "Singularity",
            "name": "Singularity Eruption: Wild pockets of gravity surge from Alleria, inflicting Shadow damage to players within 6 yards of each impact and knocking them away.",
            "spell_ids": {1235631},
            "type": "dmg_hits",
        },
        {
            "label": "Dev.Cosmos",
            "name": "Devouring Cosmos: Alleria calls upon the unending cosmos to consume her foes. Players caught within suffer Shadow damage every 1 sec and receive 99% reduced healing and absorbs.",
            "spell_ids": {1238882},
            "type": "dmg_hits",
        },
        {
            "label": "Grav.Collapse",
            "name": "Gravity Collapse: Reverberating darkness knocks the target upwards and increases their Physical damage taken by 300% for 12 sec. Also inflicts Shadow damage to all other players.",
            "spell_ids": {1239095},
            "type": "dmg_hits",
        },
    ],

    "Chimaerus, the Undreamt God": [
        {
            "label": "Alndust Ess.",
            "name": "Alndust Essence: A pool of essence that inflicts Nature damage every 1 sec to players within it and reduces movement speed by 50%.",
            "spell_ids": {1245919},
            "type": "dmg_hits",
        },
        {
            "label": "Alndust Uph.",
            "name": "Alndust Upheaval (SOAK): Chimaerus tears a hole in Reality, inflicting massive Nature damage split evenly between players within 10 yards of the impact. Stack up! Higher = better.",
            "spell_ids": {1262305, 1246827},
            "type": "soak",
        },
        {
            "label": "Disc.Roar",
            "name": "Discordant Roar: The Colossal Horror unleashes a dreadful roar, inflicting Physical damage to all players. Ignores Armor.",
            "spell_ids": {1249207},
            "type": "dmg_hits",
        },
        {
            "label": "Rift Emerg.",
            "name": "Rift Emergence: Chimaerus unleashes an unearthly roar, inflicting Nature damage to all players and causing Manifestations to emerge throughout the Rift.",
            "spell_ids": {1258610},
            "type": "dmg_hits",
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
