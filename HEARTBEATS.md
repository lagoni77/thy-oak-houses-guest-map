# Producer heartbeats (Signal House) — DEPLOYED, awaiting first scheduled run

Status 2026-10-10: deployed to the existing pinned guest-map feed at version20 after Brian's approval.10/10 tests (none skipped), two-axis review0findings, immutable cloud version20 byte-match and live public/private response checks pass. Both existing triggers remain hourly. Live collector age stays `unknown` until each collector completes its first genuine scheduled run.

Verified project: `1hZQHBIdk_Qoi_PgDQWjsxSdNgT62m16g4d34a2c_Obf322GftwaVx65m`.
Pinned deployment: `AKfycbyJ594unqrBaK5F55GUyglinuySLZWm6Z_NGsrdIOJnc1epHK250DY5TZe4IV81HwivJA`.
Prior feed version19 remains available for rollback. Credential-bearing source was preserved/applied outside Git; original served Code.gs is still pre-hooks and must not be blindly pushed. Canonical evidence/rollback pointers: Mainframe `99_SYSTEM/Handoffs/Signal_House_Claude_Handoff_2026-10-09.md`.

## What it adds

- `Heartbeat.gs` (new file): `withHeartbeat_(collector, run)` records for
  `bookings` (syncAirbnbEmails) and `payouts` (syncPayoutEmails) the producer
  id, `lastAttemptAt`, `lastSuccessAt`, `lastErrorCode` (fixed codes only) and
  `lastNewRows` in Script Properties `HEARTBEAT_bookings` / `HEARTBEAT_payouts`.
  A run that finds no new emails is a success.
- `code-gs-heartbeat-hooks.patch`: 26 changed lines against the **served
  working-tree** `Code.gs` (Dropbox `guest-map/Code.gs`, uncommitted there):
  - wraps both triggers' entry points;
  - turns the two quiet `return`s on Gmail failure and the missing payout sheet
    into recorded heartbeat failures (still no throw);
  - returns the new-row count;
  - adds `heartbeats` to `doGet` **only in the keyed private view**.
  The served Code.gs is not committed here because it embeds a credential.

LeonidOS reads `heartbeats` through `/api/toh-kpis` (`producerHeartbeats`) and
classifies them: warning after 3 h, Telegram after 6 h (approved 2026-10-10).

## Tests

    node --test 'tests/*.test.mjs'                             # Heartbeat.gs unit tests; hook tests skip
    GUEST_MAP_CODE_GS=/path/to/patched/Code.gs node --test 'tests/*.test.mjs'   # + hook integration tests

## Historical deployment checklist (approved and completed except final scheduled-run verification)

1. Apply the patch to an isolated exact live clone; never commit credential-bearing Code.gs. Heartbeat.gs is committed separately.
2. Clone the live script and diff the live `Code.js` against the served `Code.gs`.
   Follow the clasp flow in Brian's memory notes; the guest-map GAS project is container-bound.
3. `clasp push` both `Code.gs` and `Heartbeat.gs`. Then `create-version`.
4. `clasp redeploy` the pinned deployment that LeonidOS reads. Pushing alone is invisible to it.
5. Confirm both hourly triggers still point at `syncAirbnbEmails` / `syncPayoutEmails`.
6. After the next hourly run, confirm `/api/toh-kpis` shows `producerHeartbeats`.
