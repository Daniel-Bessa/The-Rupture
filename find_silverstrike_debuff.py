"""Quick script to find the Silverstrike targeting debuff spell ID in WCL."""
import sys, requests

REPORT  = "dNyhT2cZwBbt6VnK"
FIGHT   = 3   # pick any Crown fight

# Known Silverstrike damage spell
DAMAGE_ID = 1233649

def load_config(path="wcl_config.txt"):
    cfg = {}
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"): continue
            if "=" in line:
                k, v = line.split("=", 1)
                cfg[k.strip()] = v.split("#")[0].strip()
    return cfg

def get_token(cid, sec):
    r = requests.post("https://www.warcraftlogs.com/oauth/token",
                      data={"grant_type": "client_credentials"},
                      auth=(cid, sec))
    return r.json()["access_token"]

def query(token, q, v):
    r = requests.post("https://www.warcraftlogs.com/api/v2/client",
                      json={"query": q, "variables": v},
                      headers={"Authorization": f"Bearer {token}"})
    return r.json()["data"]

cfg   = load_config()
token = get_token(cfg["CLIENT_ID"], cfg["CLIENT_SECRET"])

# Step 1: get fight start time
q_meta = """
query($code: String!) {
  reportData { report(code: $code) {
    fights { id startTime endTime name }
  }}
}"""
meta = query(token, q_meta, {"code": REPORT})
fights = meta["reportData"]["report"]["fights"]
fight  = next(f for f in fights if f["id"] == FIGHT)
start  = fight["startTime"]
print(f"Fight {FIGHT}: {fight['name']}  start={start}")

# Step 2: fetch damage events for Silverstrike to get hit timestamps
q_dmg = """
query($code: String!, $fid: Int!) {
  reportData { report(code: $code) {
    events(dataType: DamageDone, fightIDs: [$fid],
           filterExpression: "ability.id = 1233649", limit: 200) { data }
  }}
}"""
dmg_data = query(token, q_dmg, {"code": REPORT, "fid": FIGHT})
hits = dmg_data["reportData"]["report"]["events"].get("data", [])
print(f"\nSilverstrike damage hits: {len(hits)}")

# Collect hit timestamps relative to fight start (first 3 rounds)
round_times = []
prev = None
for e in hits:
    t = e["timestamp"] - start
    if prev is None or t - prev > 5000:
        round_times.append(t)
    prev = t
print("Round start times (ms):", round_times[:6])

# Step 3: fetch ALL debuff events on friendly players for this fight
# Look for debuffs applied within 3s BEFORE each Silverstrike round
q_debuffs = """
query($code: String!, $fid: Int!) {
  reportData { report(code: $code) {
    events(dataType: Debuffs, fightIDs: [$fid],
           hostilityType: Friendlies, limit: 1000) { data }
  }}
}"""
debuff_data = query(token, q_debuffs, {"code": REPORT, "fid": FIGHT})
debuffs = debuff_data["reportData"]["report"]["events"].get("data", [])
print(f"\nTotal friendly debuffs: {len(debuffs)}")

# Find debuffs applied within 3s before any Silverstrike round
print("\n── Debuffs applied 0-3s BEFORE Silverstrike rounds ──")
nearby = {}
for e in debuffs:
    if e.get("type") != "applydebuff":
        continue
    t = e["timestamp"] - start
    for rt in round_times[:4]:
        if 0 <= rt - t <= 3000:
            aid = e.get("abilityGameID")
            nearby[aid] = nearby.get(aid, 0) + 1

for aid, cnt in sorted(nearby.items(), key=lambda x: -x[1]):
    print(f"  spell {aid}: applied {cnt}x before strike rounds")
