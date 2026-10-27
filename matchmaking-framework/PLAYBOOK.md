# B2B Matchmaking Framework — Event Playbook

Reusable process for exhibitor/VIP (hosted buyer) matchmaking at trade shows.
Proven at: Egypt Energy & FireX 2026, Nigeria Energy 2026.
To run it for a new show: give Claude the matchmaking Excel file and say
"run the matchmaking framework" — the steps, rules and templates below apply.

## The 4 phases

1. **Matchmaking** — build a conflict-free meeting schedule from the picks file
2. **Confirmation emails** — one personalized email per exhibitor and per VIP
3. **Bulk send** — through classic Outlook with the VBA macros in `outlook-macros/`
4. **Daily reminders** — every event morning, remind that day's VIPs of their meetings

---

## Phase 1 — Matchmaking

**Input file convention** (one sheet):
- Columns A–E: VIP info — Company, Job Title, Contact Name, Email, Phone
- Optional: Attendance column (e.g. "day 1" = only available that day), Tier/Category column
- Remaining columns: one per exhibitor, stand number in the header (e.g. "Skipper 1.E10")
- Cell **highlight** = a meeting request. Color legend (keep a legend block in the sheet):
  - Green = Exhibitor wants VIP · Grey/other = VIP wants exhibitor · Orange = mutual
- A pair picked by both sides = ONE meeting (dedupe), marked "Mutual"

**Scheduling rules (the heart of the framework):**
- 30-minute slots within the day's meeting hours (varies per event — confirm hours per day)
- Nobody double-booked: not a VIP, not an exhibitor, not across co-located shows
  (identify people by email — many VIPs appear in several lists)
- **One-trip rule:** each VIP gets ALL meetings on ONE day in ONE consecutive
  back-to-back block, no gaps. Only split across days if they have more meetings
  than a day has slots.
- Priority when capacity is tight: 1) Mutual  2) Exhibitor-pick  3) VIP-request;
  within a class, drop from the busiest VIPs / lowest tier first → goes to a Waitlist tab
- Respect per-VIP day constraints from the Attendance column
- Front-load earlier days if attendees fly out at the end (e.g. Nigeria: pack Days 1–2)
- Exhibitor daily capacity = number of slots that day; over-capacity requests
  beyond total slots are waitlisted, never silently dropped

**Output workbook tabs:** Read Me (rules + legend) · Master Schedule (filterable:
Day, Time, Exhibitor, Stand, VIP, Company, Job, Priority, Email, Phone) ·
Exhibitor Grids (per-day board: rows = times, columns = stands) ·
VIP Itineraries (per-person run sheet with their block) · Waitlist.
Always verify: zero double-bookings, zero gaps in blocks, all constraints honored.
Also offer a PRINTABLE pack: per-day master lists, one page per exhibitor,
one itinerary card per VIP, VIP check-in list (A4, page breaks preset).

## Phase 2 — Confirmation emails

One workbook, mail-merge layout — each row is one email: **To | (info columns) | Subject | Email Body**.
- Tab 1: Exhibitor emails — full day-by-day run sheet of their stand (needs POC name+email per exhibitor!)
- Tab 2: VIP emails — their one-trip block itinerary with stands
- Templates: `templates/EMAIL_TEMPLATES.md` (final approved wording, Nigeria 2026)
- No placeholders left in anything that will be sent; signature baked in
- Sending order: exhibitors first (stands prepared), then VIPs

## Phase 3 — Bulk send (Outlook)

Macros in `outlook-macros/` (import into the email workbook: Alt+F11 → File → Import File):
- `SendMatchmakingEmails.bas` — **CreateDrafts** (preview run, fills Drafts folder) and
  **SendForReal** (sends one email per row of the ACTIVE tab, 2s pause between sends).
  Skips any row whose To contains "[" (unfilled placeholder) — safe by design.
- `SendAllDrafts.bas` — **SendAllNigeriaDrafts**: sends existing drafts whose subject
  matches the event prefix (edit the InStr filter per event). Use after verifying drafts.
- `WhereAreMyDrafts.bas` — diagnostic: shows which mailbox/folder drafts landed in.

**Hard-won lessons (do not relearn these the painful way):**
- The macros ONLY work with **classic Outlook** (new Outlook has no automation). Keep
  classic Outlook OPEN during sends. Both apps show the same mailbox; sync can lag minutes.
- "This method can't be used with an inline response mail item" error = a draft was open
  in Outlook's preview pane. Click onto Inbox/Sent Items, close email windows, re-run —
  it resumes safely (sent emails leave Drafts, so re-runs never duplicate).
- Workflow: CreateDrafts → spot-check 2–3 drafts → delete all drafts → fix anything in
  Excel (fix ONCE in Excel, never edit 91 drafts) → SendForReal (or re-draft + SendAllDrafts).
- Downloaded .bas may need right-click → Properties → Unblock before importing.
- M365 limits (~30 msgs/min) are far above our 2s-pause pace.

## Phase 4 — Daily reminder emails

Workbook with one tab per event day; rows = only the VIPs with meetings that day.
Subject: "Reminder: Your <Event> Meetings TODAY (Day N)".
Body: good-morning line, today's block and meeting list with stands, arrive 15–20 min
early, "reply if you cannot attend" line, signature.
**Each event morning:** open the file, stand on that day's tab, Alt+F8 → SendForReal.
(Optional: same thing for exhibitors with that day's stand run-sheet.)

**Deferred-delivery gotcha:** when classic Outlook restarts, queued deferred emails can lose
their send registration (Outbox "Sent" column shows "None") and will silently never leave.
Fix/prevention: every show morning run FlushDueReminders (FlushReminders.bas) - it sends
anything overdue immediately and re-registers future batches with timers intact.

## WhatsApp channel (click-to-send)

Build a workbook from the Master Schedule: tabs per day + full itinerary, one row per VIP with
phone, message preview, and a https://wa.me/<digits>?text=<url-encoded> SEND hyperlink.
Rules learned: NO emojis in the URLs (they corrupt to �); WhatsApp bold = *text*; clean phones to
digits-only with country code; first click needs browser "Always allow" then it goes straight
to WhatsApp Desktop; never use unofficial bulk-sender tools (number bans). True one-click bulk
requires WhatsApp Business API (Twilio/Wati/360dialog, pre-approved templates, days of setup).

## Standing conventions
- Signature (2026 events): Mehfooz Ali · VIP & Hosted Buyer Manager · <Event> — B2B
  Matchmaking Team · +971 522136124
- Colors in deliverables: navy headers, Arial; blue/orange = co-located show split;
  green/yellow/orange = kept / moved / proposed (or priority classes)
- Day labels "Day 1/2/3" until real dates are confirmed, then find-and-replace everywhere
