# Coming-guests repair — rollout and existing-data repair (STAGED, NOT DEPLOYED)

Status 2026-10-10: built and tested locally. **No GAS push, no Sheet change, no Airbnb change, no notification.**
Every step below that touches production needs Brian's explicit approval at that step.

## What changes

| Repo | Change |
|---|---|
| guest-map `Lifecycle.gs` (new) | Danish stay-date parsing for stacked **and** adjacent-column layouts; Copenhagen-anchored year inference; rejection reasons (`dates-missing`, `dates-inverted`, `dates-implausible`, `nights-mismatch`, `parse-failed`, `parse-error`, `read-failed`, `code-missing`); reservation lifecycle ledger and fold; parse-issue ledger that holds the scan bookmark; rollout tools. |
| guest-map `Heartbeat.gs` | Adds `openIssues` to the heartbeat: the collector ran, but N emails could not be read. |
| guest-map `code-gs-coming-guests.patch` | 60 lines against **live v20 `Code.js`**. It records parse issues instead of silently skipping. It holds `LAST_SCAN_DATE`, parses only confirmation messages inside threads, ledgers confirmations, runs the lifecycle scan and publishes `openIssues`. |
| LeonidOS `fix/coming-guests` | Arrival cutoff is 16:00 **Europe/Copenhagen**, applied on every fresh/cached/stale read. Coded feed errors. `openIssues` rendered as incompleteness. The phone widget shows source errors, stale data and collector health. |

The deployed `Code.js` embeds a credential and is **never committed**. Only the patch and new files are in git. Apply the patch to a fresh private clone of the live script.

### New private sheets (same spreadsheet; never read by `doGet`)

- `ParseIssues`: EmailID, Collector, Reason, EmailDay, FirstSeenAt, LastAttemptAt, Attempts, ResolvedAt (or `dismissed …`).
- `ReservationEvents`: EmailID, Code, Kind, EventAt, CheckIn, CheckOut, RecordedAt.
- `LifecycleLog` and `RepairLog`: every applied write, with old and new values.

None of them holds names, messages or email bodies.

### Lifecycle switch: Script Property `LIFECYCLE_APPLY`

| Value | Effect |
|---|---|
| unset / `off` (default) | Cancellations and accepted changes are **ledgered only**. GuestData is unchanged. |
| `cancellations` | Sets `Status = Cancelled` on the matching row, found by the reservation code in BookingUrl. Never deletes. `doGet` already hides cancelled rows from both feeds. |
| `all` | Also writes dates and Nights from the latest **accepted** change. Cancellation stays terminal. |

Verified subjects:
- cancellations: `Annulleret: Reservation <CODE> …`, and the English `Cancelled`/`Canceled`;
- accepted change: `Reservationen blev opdateret`.

Change **requests** are ignored. Other change wordings in `CHANGE_SUBJECTS` are unverified.

## Rollout order

0. **Review:** Codex reviews LeonidOS `fix/coming-guests` and guest-map `fix/coming-guests-lifecycle`.

1. **LeonidOS first.** It is backward compatible: it handles missing `openIssues` and payloads without the new fields.
   - Merge, then run `npm run check` on the integrated tree, then `npm run deploy`.
   - Verify that `/api/coming-bookings` shows `checkedAt` and that dashboard Arrivals renders.

2. **GAS, in a private rollout folder outside git and Dropbox.**
   1. `clasp clone` the live script. Confirm `Code.js` and `Heartbeat.js` still byte-match v20; if not, stop and reconcile.
   2. Apply the patch to `Code.js`. Add `Lifecycle.gs` as `Lifecycle.js`. Replace `Heartbeat.js` with the branch `Heartbeat.gs`.
   3. Run `GUEST_MAP_CODE_GS=<patched Code.js> node --test 'tests/*.test.mjs'`. Expect 76/76.
   4. **Leave `LIFECYCLE_APPLY` unset.**
   5. Run `clasp push`.
      ⚠ **Time-driven triggers run HEAD, so this activates the new collector immediately.** The pinned web-app deployment does not change until step 6.
   6. `clasp create-version`, then `clasp redeploy` the pinned deployment so `doGet` exposes `openIssues`.

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
   2. Spot-check candidates against Airbnb, for example the ones in Codex's review.
   3. `repairStayDatesConfirmed()` writes CheckIn, CheckOut and Nights, logging old values to `RepairLog`. It is idempotent: a second run fixes 0.
   4. **Rollback:** write `RepairLog` Old* values back.

6. **Cancellations.**
   1. The ledger already holds the backlog since 2024-09-01, because the first lifecycle scan starts there.
   2. `previewLifecycle('cancellations')` is a dry run: it shows wouldApply and rowNotFound.
   3. Spot-check, then set `LIFECYCLE_APPLY=cancellations`. The next hourly run applies it, logged to `LifecycleLog`.
   4. **Rollback:** set `off`, then restore the logged Old `Status` values.

7. **Accepted changes.**
   1. `previewLifecycle('all')`.
   2. Verify the reproduced stay from Codex's review lands on the Airbnb dates.
   3. Set `LIFECYCLE_APPLY=all`.
   4. Changes whose dates cannot be parsed stay open `ParseIssues` (`dates-missing`), visible as `openIssues`.

8. **Phone.** Copy `widgets/coming-bookings-widget.scriptable.js` into Scriptable (same script name). Check:
   - **Preview:** the guest card, with no notes when healthy.
   - **Error case:** with the node stopped it shows "Connection … severed". A feed failure shows "Kommende gæster ukendt", never "Ingen kommende gæster".

## Not changed (policy preserved)

- Signal House 3 h warning / 6 h Telegram age policy. `openIssues` is shown but does **not** alert; data-quality alerting needs its own approved contract.
- A collector heartbeat is shown as "ran on time" only, never as proof that bookings are complete or correct.
- `auditMissingBookings` (daily email) is unchanged. It still skips messages it cannot parse; those are now visible as `ParseIssues` from the sync.
- Public feed: `{guests, payouts}` only. `heartbeats` remain keyed/private. Issue and ledger sheets are never served.
