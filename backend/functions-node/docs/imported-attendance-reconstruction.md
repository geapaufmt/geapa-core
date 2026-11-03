# Imported attendance reconstruction: local candidate

`activities/v2/206_act_import_reconstruction.gs` is an offline, pure planner. It is deliberately absent from `core-sources.json` and from the HTTP command registry. It does not authorize a caller, create an import endpoint, change normal finalization, or execute writes.

The caller supplies a closed intent and a complete fingerprinted read set. The plan recognizes documentary facts, preserves original participant IDs and results, produces current recognition revisions and provenance, and uses the existing model validators. Historical decision records remain unchanged. A pending justification may be linked to an explicitly new recognized revision; it is never approved by reconciliation.

Closure priority is mandatory:

1. Existing canonical finalization: refuse replacement.
2. Sealed documentary finalization, prior recognition or administrative recovery: preserve its recorded timestamp and label its actual origin.
3. Proven absence of an equivalent record: the explicitly authorized administrative 21:00 America/Cuiaba convention, only when compatible with the activity schedule.

A later reopening, saved draft, contradictory simultaneous event or incomplete audit prevents automatic recognition. The current reconciliation timestamp is distinct from the original or conventional closure. Generic migration/update timestamps are not closure evidence.

## Application gate

Local proposals and synthetic transaction tests do **not** authorize institutional data changes. Before a real application, validate historical policy applicability (current template equality alone does not prove historical validity), obtain approval of exact before/after values and source hashes, complete the authenticated DEV proof, and integrate the isolated transaction with canonical ownership/drain/lease/journal handling. No current test adapter is a production executor.

Per-activity application must atomically verify the complete read/query sets, exact revisions and update times, commit immutable originals and revisions, receipt/audit and shared projection fences, and reject a changed or ambiguous state. Recover a lost response by operationId and digest. A rollback is a compensating revision conditional on unchanged written heads, epochs and absence of later dependent operations; never a blind snapshot restore.

## Tests

Run with Node 22 from `backend/functions-node`. The directed suite is `test/import-reconstruction*.test.cjs`, `test/attendance-import-consultation.test.cjs`, `test/frequency-query.test.cjs`, `test/frequency-cycle.test.cjs`, and `test/agenda-profile-matrix.test.cjs`.

The emulator test requires `GEAPA_ATTENDANCE_EMULATOR_TEST=1` and `FIRESTORE_EMULATOR_HOST=127.0.0.1:8187`. It accepts only the demo project and synthetic namespaced fixtures. It exercises CAS, concurrent intents, response loss, recovery, preservation and conditional rollback. It does not establish Firebase Auth/App Check or real ROLE authorization.

Private inventories, documentary source rows, nominal manifests and runtime evidence must remain outside the repository.
