# TODO

## Immediate / Missing Tracking
- [ ] Last boss of Voidspire — mechanics not yet tracked
- [ ] 2 bosses from the March raid — mechanics not yet tracked

## Features
- [ ] **Shareable multi-guild web tool** — a proper hosted web app. Flow:
  1. **Landing page** — register/login with an account; create a guild with a name → gets a unique URL slug (e.g. `thisapp.com/the-rupture`)
  2. **Guild page** (`/guildname`) — shows accumulated raid history (same layout as current `index.html`), plus a URL input box to submit a new WarcraftLogs report URL → runs the audit server-side and appends the new raid card
  - **Auth/security**: account system so only the guild owner (or invited members) can submit reports to their page; public visitors can view. Login options:
    - Email/password
    - Battle.net OAuth (SSO via Blizzard's API — ideal since the target audience already has Battle.net accounts)
  - **Daniel's WCL API credentials** baked in server-side — users never see or need them
  - Needs: a real backend (Flask/FastAPI), a database (guild accounts, report history), and hosting


- [ ] Boss abilities tracking (per-boss mechanic hit detection, already started for some bosses)
- [ ] Defensives list (track who used what defensives and when, per pull)
- [ ] Overview page for all bosses (summary of all progress across the tier)
- [ ] Public logs / Scout mode — pull Mythic logs from other guilds to improve mechanic tracking and see what spells to track
- [ ] Get info on what everyone did in the first few weeks of the tier
  - Number of vault slots opened
  - Crests farmed (Weathered / Carved / Runed / Gilded)
  - Weeklies completed

## Accessibility / No-Code Distribution
- [ ] **Zero technical knowledge required** — the tool should be fully self-service for end users:
  - No coding, no terminal, no AI needed
  - Paste a WarcraftLogs URL → get a report. That's the entire interaction
  - All complexity (API calls, parsing, mechanic detection) happens invisibly on the server
  - Error messages should be human-readable ("This report is private, make it public on WarcraftLogs first") not stack traces
- [ ] Think through the distribution model — how guilds discover and sign up for the tool (word of mouth, Discord, WoW community forums?)

## UI / Visual Overhaul
- [ ] **Replace raw tables with a proper visual design** — current tables work but aren't engaging. Ideas to explore:
  - Charts and graphs (damage taken per mechanic, attendance over time, performance trends)
  - Player cards instead of table rows
  - Color-coded indicators (green/yellow/red for performance thresholds)
  - Boss timeline visualizations (when mechanics hit, who died and when)
  - Overall look should feel closer to WarcraftLogs / Raider.IO than a spreadsheet

## Research
- [ ] How does WoW Audit figure stuff out from guild players? (API, addon, how data is sourced)
- [ ] **Study WipeFest** — understand how they identify relevant mechanics per fight, what they choose to show vs. ignore, and how they scale across multiple raids/tiers. Use this to build a faster pipeline for onboarding new bosses/raids
- [ ] **Fight onboarding process** — need a systematic way to analyze a new raid tier quickly:
  - Identify which mechanics are trackable via WarcraftLogs events
  - Decide what's worth showing (avoidable damage, interrupts, debuff handling) vs. noise
  - Goal: go from a new raid release to full tracking faster than we did for current tier
