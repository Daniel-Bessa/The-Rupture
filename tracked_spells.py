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

SPELL_NAMES = {
    5512: "Healthstone",        6262: "Healthstone",
    432112: "Health Potion",    431924: "Health Potion",
    431932: "Tempered Potion",  432098: "Potion of Unwavering Focus",
    431945: "Light's Potential", 432106: "Void-Shrouded Tincture",
    1236616: "Light's Potential", 245898: "Light's Potential",
    245897: "Light's Potential",  241308: "Light's Potential",
    241309: "Light's Potential",
    241304: "Health Potion",    241305: "Health Potion",
    258138: "Health Potion",    1234768: "Health Potion",
    48792: "Icebound Fortitude",  48707: "Anti-Magic Shell",
    55233: "Vampiric Blood",      49028: "Dancing Rune Weapon",
    198589: "Blur",               196718: "Darkness",
    196555: "Netherwalk",         187827: "Metamorphosis",
    22812: "Barkskin",            61336: "Survival Instincts",
    102342: "Ironbark",
    363916: "Obsidian Scales",    374348: "Renewing Blaze",
    186265: "Aspect of the Turtle", 109304: "Exhilaration",
    264735: "Survival of the Fittest",
    45438: "Ice Block",           55342: "Mirror Image",
    108978: "Alter Time",         110959: "Greater Invisibility",
    115203: "Fortifying Brew",    122783: "Diffuse Magic",
    131645: "Zen Meditation",     122278: "Dampen Harm",
    642: "Divine Shield",         31850: "Ardent Defender",
    86659: "Guardian of Ancient Kings", 184662: "Shield of Vengeance",
    19236: "Desperate Prayer",    47585: "Dispersion",
    586: "Fade",                  15286: "Vampiric Embrace",
    31224: "Cloak of Shadows",   5277: "Evasion",
    1966: "Feint",               1856: "Vanish",
    108271: "Astral Shift",      98008: "Spirit Link Totem",
    104773: "Unending Resolve",  108416: "Dark Pact",
    6789: "Mortal Coil",
    871: "Shield Wall",          118038: "Die by the Sword",
    184364: "Enraged Regeneration", 97462: "Rallying Cry",
    23920: "Spell Reflection",
}
