════════════════════════════════════════════════════════════
  HOW THIS IS ORGANIZED
════════════════════════════════════════════════════════════

ONENOTE STRUCTURE
───────────────────────────────────

  Notebook: Client - Engagement
  │
  ├── Section: _Overview
  │     Page: Scope & Stack         ← fill at kickoff, rarely touched again
  │     Page: Findings              ← one row per confirmed finding, all apps
  │
  ├── Section Group: App Name
  │     Page: Surface               ← your daily driver
  │     Page: SQLi                  ← create only when you have real indicators
  │     Page: XSS
  │     Page: (etc.)
  │
  └── Section Group: App 2
        Page: Surface
        Page: ...


THE THREE RULES
───────────────────────────────────

  1. Surface is where you live.
     Drop every endpoint here as you find it. Tag vuln classes.
     Update the status symbol inline as you test.
     The daily log at the bottom is your "where I left off."

  2. Create a vuln class page only when you have indicators.
     If you see a DB-backed param throw an error on a quote, create SQLi.
     Don't pre-create all 8 pages for every app. Half will stay blank
     and that blank space becomes noise.

  3. The moment something is confirmed, it goes to Findings.
     Don't let findings live only in the vuln class page.
     Findings is the cross-app rollup you'll use to write the report.


STATUS SYMBOLS (use in Surface endpoint map)
───────────────────────────────────
  ?   untested
  >   active / in progress right now
  ✓   confirmed finding — added to Findings page
  -   tested, nothing found
  x   not applicable


STARTING A NEW APP IN AN EXISTING ENGAGEMENT
───────────────────────────────────
  1. Right-click the notebook → Add Section Group → name it after the app
  2. Create a Surface page inside it (paste from Surface template)
  3. Fill in App name at the top
  4. Start enumerating — drop endpoints into the map as you go
  5. Create vuln class pages as indicators appear


STARTING A BRAND NEW ENGAGEMENT
───────────────────────────────────
  1. New notebook → name it "ClientName - EngagementType - Year"
  2. Create _Overview section → paste Scope & Stack and Findings pages
  3. Fill in Scope & Stack
  4. Create a section group per app in scope
  5. Paste Surface page into each section group
  6. Start testing
