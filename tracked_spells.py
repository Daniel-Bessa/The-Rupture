# tracked_spells.py
# ─────────────────────────────────────────────────────────────────────────────
# Add new spell IDs here when Blizzard adds new potions or defensives in a patch.
# The main script imports everything from this file.
# ─────────────────────────────────────────────────────────────────────────────

HEALTHSTONE_IDS = {
    5512,    # Healthstone
    6262,    # Healthstone (Midnight)
}

HEALTH_POT_IDS = {
    432112,  # Algari Healing Potion (TWW)
    431924,  # Algari Healing Potion (TWW alternate)
    241304,  # Silvermoon Health Potion (Midnight)
    241305,  # Silvermoon Health Potion (Midnight alternate)
    258138,  # Silvermoon Health Potion (Midnight alternate)
    1234768, # Silvermoon Health Potion (Midnight)
}

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

# ── Personal defensive cooldowns ─────────────────────────────────────────────
# Tracked as casts (or aura-applied for procs — see CLASS_DEFENSIVE_PROC_IDS).
# Excludes spammed abilities (Ironfur, Ignore Pain, etc.) and purely passive effects.

CLASS_DEFENSIVE_IDS = {
    # Death Knight
    48792,   # Icebound Fortitude
    48707,   # Anti-Magic Shell
    55233,   # Vampiric Blood
    49028,   # Dancing Rune Weapon
    49039,   # Lichborne
    # Demon Hunter
    198589,  # Blur
    196718,  # Darkness
    196555,  # Netherwalk
    187827,  # Metamorphosis (Vengeance)
    # Druid
    22812,   # Barkskin
    61336,   # Survival Instincts
    # Evoker
    363916,  # Obsidian Scales
    374348,  # Renewing Blaze
    # Hunter
    186265,  # Aspect of the Turtle
    109304,  # Exhilaration
    233526,  # Fortitude of the Bear
    264735,  # Survival of the Fittest
    # Mage
    45438,   # Ice Block
    55342,   # Mirror Image
    108978,  # Alter Time
    110959,  # Greater Invisibility
    86949,   # Cauterize (passive proc — tracked via aura applied)
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
    498,     # Divine Protection
    # Priest
    19236,   # Desperate Prayer
    47585,   # Dispersion
    586,     # Fade
    # 15286 moved to CLASS_EXTERNAL_IDS (Vampiric Embrace — raid heal)
    # Rogue
    31224,   # Cloak of Shadows
    5277,    # Evasion
    1966,    # Feint
    1856,    # Vanish
    45182,   # Cheat Death (passive proc — tracked via aura applied)
    # Shaman
    108271,  # Astral Shift
    # Warlock
    104773,  # Unending Resolve
    108416,  # Dark Pact
    6789,    # Mortal Coil
    # Warrior
    871,     # Shield Wall
    118038,  # Die by the Sword
    184364,  # Enraged Regeneration
    23920,   # Spell Reflection
}

# Passive procs that fire as aura-applied events rather than casts.
# Currently tracked alongside regular defensives via applybuff events.
CLASS_DEFENSIVE_PROC_IDS = {
    86949,   # Cauterize (Mage) — procs when hit below 20% HP
    45182,   # Cheat Death (Rogue) — procs when hit would kill
}

# ── Externals / raid cooldowns ────────────────────────────────────────────────
# Shown in a separate column — spells cast ON other players or as raid-wide CDs.

CLASS_EXTERNAL_IDS = {
    # Death Knight
    51052,   # Anti-Magic Zone
    # Druid
    102342,  # Ironbark
    # Evoker
    374227,  # Zephyr
    # Monk
    116849,  # Life Cocoon
    # Paladin
    1022,    # Blessing of Protection
    6940,    # Blessing of Sacrifice
    # Priest
    33206,   # Pain Suppression
    47788,   # Guardian Spirit
    # Shaman
    98008,   # Spirit Link Totem
    207399,  # Ancestral Protection Totem
    # Warrior
    97462,   # Rallying Cry
    # Priest
    15286,   # Vampiric Embrace (Shadow Priest — raid heal)
}

ALL_TRACKED_IDS = (HEALTHSTONE_IDS | HEALTH_POT_IDS | COMBAT_POT_IDS
                   | CLASS_DEFENSIVE_IDS | CLASS_EXTERNAL_IDS)

SPELL_NAMES = {
    # Consumables
    5512: "Healthstone",        6262: "Healthstone",
    432112: "Health Potion",    431924: "Health Potion",
    431932: "Tempered Potion",  432098: "Potion of Unwavering Focus",
    431945: "Light's Potential", 432106: "Void-Shrouded Tincture",
    1236616: "Light's Potential", 245898: "Light's Potential",
    245897: "Light's Potential",  241308: "Light's Potential",
    241309: "Light's Potential",
    241304: "Health Potion",    241305: "Health Potion",
    258138: "Health Potion",    1234768: "Health Potion",
    # Personal defensives — Death Knight
    48792: "Icebound Fortitude",  48707: "Anti-Magic Shell",
    55233: "Vampiric Blood",      49028: "Dancing Rune Weapon",
    49039: "Lichborne",
    # Personal defensives — Demon Hunter
    198589: "Blur",               196718: "Darkness",
    196555: "Netherwalk",         187827: "Metamorphosis",
    # Personal defensives — Druid
    22812: "Barkskin",            61336: "Survival Instincts",
    # Personal defensives — Evoker
    363916: "Obsidian Scales",    374348: "Renewing Blaze",
    # Personal defensives — Hunter
    186265: "Aspect of the Turtle", 109304: "Exhilaration",
    233526: "Fortitude of the Bear",  264735: "Survival of the Fittest",
    # Personal defensives — Mage
    45438: "Ice Block",           55342: "Mirror Image",
    108978: "Alter Time",         110959: "Greater Invisibility",
    86949: "Cauterize",
    # Personal defensives — Monk
    115203: "Fortifying Brew",    122783: "Diffuse Magic",
    131645: "Zen Meditation",     122278: "Dampen Harm",
    # Personal defensives — Paladin
    642: "Divine Shield",         31850: "Ardent Defender",
    86659: "Guardian of Ancient Kings", 184662: "Shield of Vengeance",
    498: "Divine Protection",
    # Personal defensives — Priest
    19236: "Desperate Prayer",    47585: "Dispersion",
    586: "Fade",
    # Personal defensives — Rogue
    31224: "Cloak of Shadows",   5277: "Evasion",
    1966: "Feint",               1856: "Vanish",
    45182: "Cheat Death",
    # Personal defensives — Shaman
    108271: "Astral Shift",
    # Personal defensives — Warlock
    104773: "Unending Resolve",  108416: "Dark Pact",
    6789: "Mortal Coil",
    # Personal defensives — Warrior
    871: "Shield Wall",          118038: "Die by the Sword",
    184364: "Enraged Regeneration", 23920: "Spell Reflection",
    # Externals / raid CDs
    51052: "Anti-Magic Zone",
    102342: "Ironbark",
    374227: "Zephyr",
    116849: "Life Cocoon",
    1022: "Blessing of Protection",
    6940: "Blessing of Sacrifice",
    33206: "Pain Suppression",
    47788: "Guardian Spirit",
    98008: "Spirit Link Totem",
    207399: "Ancestral Protection Totem",
    97462: "Rallying Cry",
    15286: "Vampiric Embrace",
}
