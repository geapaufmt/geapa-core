# People directory reader

`members/100_people_directory.gs` is the canonical reader for `/admin/pessoas` and `/admin/pessoas/detalhe`. Node only supplies the existing Firestore/Auth/App Check adapter. No writer or private profile grant is introduced.

Authorization uses a distinct `PEOPLE_DIRECTORY_READ` capability in ROLE policies, canonical binding/current membership/current function, explicit person/resource scope and field masks. `MEMBERS_CASE_READ` remains unchanged. Technical administrator fallback is intentionally absent.

The reader derives membership and functions through the existing aggregate/head/model contracts. Supplemental documented non-membership relations have a separate closed schema, `memberInstitutionalRelations/{personId}` (`people-relations-v1`), with source reference, approval reference, revision, optional documented dates and typed records. These relations never create membership, Auth bindings, capabilities or role eligibility. Institutional imports require a separately approved manifest and CAS; there is no public write endpoint for them.

Current limit: 250 explicitly approved persons and 2,500 general dependency paths. Batched reads load the bounded approved general cohort, then filter/sort/page on the server. This is response pagination, not a Firestore indexed paginated query. It deliberately does not fetch target private profiles/identifiers. Cost must be measured before expansion beyond the current population; a derived indexed reader may be warranted later.

The response snapshot binds actor, source policies, authority, aggregate revision, returned general facts and the operations fence. Changed pages fail closed. Node rechecks sources and token revocation before returning. Read-only Firestore transactions avoid unnecessary locks while preserving final consistency checks.

General masks: idPessoa, displayLabel, state, categories, functions, relations, history. No CPF, birth, contact or private academic values. Institutional history needs its own explicit mask. Private correction workflows remain separate.

Tests: people-directory.test.cjs; opt-in loopback people-emulator.test.cjs. Emulator Auth/App Check are test doubles and do not prove real cloud authentication. No production policy is activated by building or deploying the reader.
