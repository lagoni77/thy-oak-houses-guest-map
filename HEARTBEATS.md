# Producer heartbeats (Signal House) — STAGED, NOT DEPLOYED

Status 2026-10-10: built and tested locally; **no GAS push, version or redeploy**.
Live TOH collector age stays `unknown` in LeonidOS until this is deployed.

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

    node --test tests/                                         # Heartbeat.gs unit tests; hook tests skip
    GUEST_MAP_CODE_GS=/path/to/patched/Code.gs node --test tests/   # + hook integration tests

## Deployment prerequisites (each needs Brian's explicit approval)

1. Apply the patch to the served `Code.gs`. Commit Heartbeat.gs alongside it.
2. Clone the live script and diff the live `Code.js` against the served `Code.gs`.
   Follow the clasp flow in Brian's memory notes; the guest-map GAS project is container-bound.
3. `clasp push` both `Code.gs` and `Heartbeat.gs`. Then `create-version`.
4. `clasp redeploy` the pinned deployment that LeonidOS reads. Pushing alone is invisible to it.
5. Confirm both hourly triggers still point at `syncAirbnbEmails` / `syncPayoutEmails`.
6. After the next hourly run, confirm `/api/toh-kpis` shows `producerHeartbeats`.
