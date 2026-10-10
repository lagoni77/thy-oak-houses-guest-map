---
title: Coming-guests repair — rollout and existing-data repair
type: runbook
status: active
wing: thy_oak_houses
room: coming_guests
hall: runbook
created: 2026-10-10
last_reviewed: 2026-10-10
refinery_outcome: implement
source_ref:
  - "02 Mainframe/99_SYSTEM/Handoffs/TOH_Dashboard_Season_Review_Handoff_2026-10-07.md#coming-guests-review--2026-10-10"
  - "LeonidOS branch fix/coming-guests; guest-map branch fix/coming-guests-lifecycle"
---
# Coming-guests repair — rollout and existing-data repair (DEPLOYED)

Status 2026-10-10: deployed under Brian's explicit “Ready to deploy” authorization for the full necessary rollout. The procedure below is retained for future rollouts.

Release: LeonidOS clean master `092bc3f` (668 tests /121 syntax checks); pinned GAS version22, all four uploaded files byte-matched (105 tests). Live confirmation prose exposed a heading ambiguity, fixed in lifecycle commit `9e933ba` by accepting standalone stay-date headings; regression and local real-message replay passed.

Applied70 date repairs,40 cancellations and15 accepted changes with before-images logged. Lifecycle mode is `all`; existing hourly triggers reconcile the whole ledger. Final audit:0 wouldFix and0 missing/unparseable. All16 upcoming arrivals match Airbnb calendar, with0 date mismatches or same-day stays. Signal House collector healthy with0 parse issues; public/private feed boundary verified. Existing iCloud widget updated; native phone rendering/sync remains unverified. Optional calendar-to-Sheet rewrite was not performed.

Private rollback copies: `/tmp/coming-guests-rollout/rollback` and `/tmp/coming-guests-rollout/widget-before.js`.

## What changes

| Repo | Change |
|---|---|
| guest-map `Lifecycle.gs` (new) | Danish stay-date parsing for stacked **and** adjacent-column layouts. Impossible calendar dates rejected. Year inference picks the earliest check-in year whose stay has not ended before the email (bookings months ahead, New Year, changes during a stay). Rejection reasons: `dates-missing`, `dates-invalid`, `dates-inverted`, `dates-implausible`, `nights-mismatch`, `parse-failed`, `parse-error`, `read-failed`, `code-missing`. Also: reservation lifecycle ledger and fold; parse-issue ledger that holds the scan bookmark; write-ahead change logs with rollback; rollout tools. |
| guest-map `Heartbeat.gs` | Adds `openIssues` to the heartbeat: the collector ran, but N emails could not be read. |
| guest-map `code-gs-coming-guests.patch` | 62 lines against **live v20 `Code.js`**. It records parse issues instead of silently skipping. It holds `LAST_SCAN_DATE`, parses only confirmation messages inside threads, ledgers confirmations, runs the lifecycle scan and publishes `openIssues`. It also emits `DatesUnverifiedSince` in the keyed feed only. |
| LeonidOS `fix/coming-guests` | Arrival cutoff is 16:00 **Europe/Copenhagen**, applied on every fresh/cached/stale read. The booking feed is accepted only with HTTP OK, no application error and a valid `guests` array; otherwise it serves the last good report (stale, dated) or a coded error, and the cache is never replaced. Notice-only accepted changes are verified against Airbnb's own calendar by reservation code. `openIssues` is rendered as incompleteness. The phone widget shows source errors, stale data, collector health and unverified dates. |

The deployed `Code.js` embeds a credential and is **never committed**. Only the patch and new files are in git. Apply the patch to a fresh private clone of the live script.

### New private sheets and column (never in the public feed)

- `ParseIssues`: EmailID, Collector, Reason, EmailDay, FirstSeenAt, LastAttemptAt, Attempts, ResolvedAt (or `dismissed …`).
- `ReservationEvents`: EmailID, Code, Kind, EventAt, CheckIn, CheckOut, RecordedAt. Notice-only changes have empty dates.
- `LifecycleLog` and `RepairLog`: At, Row, Code, Field, Old, New, State.
  - Each before-image is appended as `pending` **before** the cell changes, then marked `applied`.
  - Other states: `failed`, `not-applied`, `conflict`, `rolling-back`, `rolled-back`, `rollback-conflict`.
- GuestData column `DatesUnverifiedSince`: created on first use. Set when an accepted change arrived without dates; cleared by a later dated change.

None of them holds names, messages or email bodies.

### Lifecycle switch: Script Property `LIFECYCLE_APPLY`

Every hourly run reconciles the **whole ledger** under the current value. Events ledgered while the switch was `off` are therefore applied once it is switched on. `applyLifecycleNow()` does the same on demand.

| Value | Effect |
|---|---|
| unset / `off` (default) | Cancellations and accepted changes are **ledgered only**. GuestData is unchanged. |
| `cancellations` | Sets `Status = Cancelled` on the matching row, found by the reservation code in BookingUrl. Never deletes. `doGet` already hides cancelled rows from both feeds. |
| `all` | Also applies accepted changes. A dated change writes CheckIn, CheckOut and Nights. A notice-only change keeps the last known dates and sets `DatesUnverifiedSince`. Cancellation stays terminal. |

Verified subjects:
- cancellations: `Annulleret: Reservation <CODE> …`, and the English `Cancelled`/`Canceled`;
- accepted change: `Reservationen blev opdateret`.

Change **requests** are ignored. Other change wordings in `CHANGE_SUBJECTS` are unverified.

### Notice-only accepted changes — how dates get verified

The accepted-change notice may carry only a reservation-detail link and no dates. The repair handles it like this:

1. **GAS:** the change is ledgered without dates; dates are never inferred from their absence. The row keeps its last known dates and gets `DatesUnverifiedSince`, which appears in the keyed feed only.
2. **LeonidOS:** before applying the arrival cutoff, each flagged row is looked up by reservation code in Airbnb's own iCal.
   - **Found:** iCal dates are used, labelled `DatesSource: airbnb-calendar`.
   - **Not found, or the calendar is unavailable:** the row keeps its dates with `DatesUnverified`, shown on the phone and dashboard as "datoer ændret — ikke bekræftet".
   - **Calendar down after an earlier successful check:** the earlier calendar dates are kept (`DatesStale`, with the original retrieval time in `DatesVerifiedAt`) instead of falling back to the obsolete Sheet dates — but **only if that check was retrieved after the newest accepted-change notice** (`DatesUnverifiedSince`). Evidence from before a newer change, or with a missing or invalid timestamp on either side, takes the unknown-dates path.
   - **What counts as a calendar read:** a complete iCalendar — closed `VCALENDAR` envelope, balanced events, valid dates on every event. An HTTP 200 HTML page, or an empty or truncated body, is an `invalid-calendar` house failure: outage warning, earlier evidence kept. A complete calendar with no events is a valid read. Live feeds were checked against this rule read-only on 2026-10-10: both houses pass.
   - **Calendar down with no earlier evidence:** the row stays listed as unverified even when its obsolete Sheet dates have passed the arrival cutoff.
   - **House feed read but the code is absent:** the reservation is not current or future in Airbnb (past, or cancelled ahead of the Sheet). Earlier evidence is dropped and the normal cutoff applies.
   - Calendar evidence lives in node memory; a restart starts without it.
   - **Disclosure:** the phone and dashboard show the calendar outage (by house), stale calendar dates with their retrieval time, and the unverified reason.
   - Read-only shape check 2026-10-10 (counts only): every `Reserved` event in both house feeds carries a reservation code plus DTSTART/DTEND — Tiny 9/9, Ocean 8/8.
   - iCal covers current and future reservations only, which is exactly the arrivals window.
   - `DatesVerifiedAt` is the node's **retrieval** time for the calendar. Airbnb publishes no observation time for the feed.
3. **Not yet done:** the sheet itself is not rewritten from iCal, so the Revenue and Guest Map history keeps the last known dates with the flag. A later step could write calendar-verified dates back through the logged path; that needs its own approval.

## Rollout order

0. **Review:** Codex reviews LeonidOS `fix/coming-guests` and guest-map `fix/coming-guests-lifecycle`.

1. **LeonidOS first.** It is backward compatible: it handles missing `openIssues`, `DatesUnverifiedSince` and other new fields.
   - Merge, then run `npm run check` on the integrated tree, then `npm run deploy`.
   - Verify `/api/coming-bookings` (`checkedAt`, no `error`) and that dashboard Arrivals renders.

2. **GAS, in a private rollout folder outside git and Dropbox.**
   1. `clasp clone` the live script. Confirm `Code.js` and `Heartbeat.js` still byte-match v20; if not, stop and reconcile.
   2. Apply the patch to `Code.js`. Add `Lifecycle.gs` as `Lifecycle.js`. Replace `Heartbeat.js` with the branch `Heartbeat.gs`.
   3. Run `GUEST_MAP_CODE_GS=<patched Code.js> node --test 'tests/*.test.mjs'`. Expect 104/104.
   4. **Leave `LIFECYCLE_APPLY` unset.**
   5. Run `clasp push`.
      ⚠ **Time-driven triggers run HEAD, so this activates the new collector immediately.** The pinned web-app deployment does not change until step 6.
   6. `clasp create-version`, then `clasp redeploy` the pinned deployment so `doGet` exposes `openIssues` and `DatesUnverifiedSince`.

   **Rollback:**
   - Push the preserved v20 `Code.js`/`Heartbeat.js` (removes `Lifecycle.js`).
   - Redeploy the pinned deployment to v20.
   - The new sheets can stay; nothing reads them.

3. **After the first hourly run (read-only checks):**
   - `/api/coming-bookings` → `producerHeartbeats.bookings.openIssues`.
   - The `ParseIssues` rows (codes only).
   - `LAST_SCAN_DATE` held at the oldest open issue, if any.
   - For an email that will never parse (a non-booking template): `dismissParseIssue(emailId, why)`.

4. **Classify templates.** `classifyRecentAirbnbSubjects()` is read-only and logs redacted subject templates and counts. Confirm that accepted changes appear as `changed: Reservationen blev opdateret`, and that no change *request* is classified `changed`.

5. **Repair existing stay dates.**
   1. `auditStayDates()` is a dry run. It re-parses each row's own confirmation email and lists candidates by row number and reservation code. Kinds: `checkout-equals-checkin`, `year`, `other`. Rows with zero Nights are re-derived only from the email, never from the zero.
   2. Rows whose reservation has **any accepted change** in the ledger are skipped and counted as `lifecycleOwned`; the change is the newer evidence.
   3. Spot-check candidates against Airbnb, for example the ones in Codex's review.
   4. `repairStayDatesConfirmed()` writes each changed field through the write-ahead `RepairLog`. It is idempotent: a second run fixes 0.
   5. **Rollback:** `rollbackStayDateRepairs('<ISO time of the repair run>')`. It restores before-images where the cell still holds the repaired value, flags human edits as `rollback-conflict`, and is idempotent.

6. **Cancellations.**
   1. The ledger already holds the backlog since 2024-09-01, because the first lifecycle scan starts there.
   2. `previewLifecycle('cancellations')` is a dry run: it shows wouldApply and rowNotFound.
   3. Spot-check, then set `LIFECYCLE_APPLY=cancellations`. The next hourly run, or `applyLifecycleNow()`, applies the **already-ledgered** backlog through `LifecycleLog`.
   4. **Rollback:** set `LIFECYCLE_APPLY=off`, then run `rollbackLifecycleChanges('<ISO time>')`. It refuses while the switch is on, so the next run cannot re-apply.

7. **Accepted changes.**
   1. `previewLifecycle('all')`.
   2. Verify the reproduced stay from Codex's review: either its dated change lands on the Airbnb dates, or its notice-only change shows `DatesUnverifiedSince`, and LeonidOS shows the calendar dates.
   3. Set `LIFECYCLE_APPLY=all`, then run `applyLifecycleNow()` or wait for the next run.
   4. Dated changes that fail validation stay open `ParseIssues` (for example `dates-invalid` or `nights-mismatch`), visible as `openIssues`.

8. **Phone.** Copy `widgets/coming-bookings-widget.scriptable.js` into Scriptable (same script name). Check:
   - **Preview:** the guest card, with no notes when healthy.
   - **Error case:** with the node stopped it shows "Connection … severed". A feed failure shows "Kommende gæster ukendt", never "Ingen kommende gæster".
   - **Unverified dates** show "⚠ Datoer ændret i Airbnb — ikke bekræftet".

## Interrupted runs

- **A crash between a cell write and its log update** leaves a `pending` entry. The next run, or a rollback, resolves it from the cell's actual value: `applied`, `not-applied` or `conflict`. Nothing is written twice.
- **A failing cell write** marks its entry `failed` and is retried on the next run under a new entry.
- **If the log append itself fails,** no cell is changed.

## What the booking heartbeat does and does not cover

The `toh.bookings` heartbeat (Signal House) covers the **GAS booking collector** only: whether it ran, and how many emails it could not read (`openIssues`). It says nothing about:
- **the Airbnb calendar:** freshness or availability. The calendar has no producer heartbeat; its retrieval state is disclosed in `dateReconciliation`, on the phone and on the dashboard.
- **the lifecycle switch:** whether cancellations or changes have been applied.
- **the bookings themselves:** whether they are complete or correct.

## Not changed (policy preserved)

- Signal House 3 h warning / 6 h Telegram age policy. `openIssues`, unverified dates and calendar outages are shown but do **not** alert; data-quality alerting needs its own approved contract.
- A collector heartbeat is shown as "ran on time" only, never as proof that bookings are complete or correct.
- `auditMissingBookings` (daily email) is unchanged. It still skips messages it cannot parse; those are now visible as `ParseIssues` from the sync.
- Public feed: `{guests, payouts}` only. `heartbeats` and `DatesUnverifiedSince` are keyed/private. Issue, ledger and log sheets are never served.
