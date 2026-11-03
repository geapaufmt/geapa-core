# Source synchronization — 2026-11-03

This integration preserves the modern canonical sources and the previously
published checkpoints. It does not deploy a service or authorize a parameter.

| Environment | Effective Function revision | Functional source |
| --- | --- | --- |
| DEV | geapacommanddevpilot-00080-xos | 4fe2b7562539d96d04af5ccc252c4ea850dcffdb |
| PROD, unchanged | geapacommand-00030-cep | da7097460431 |

DEV package checkpoint: `8d44b786bd1d`. Release-manifest SHA-256:
`cfbac10f408c3a3882d56492d0206f7d794c83438b78e46fb4f7edf13f50545f`.
The official-main reconciliation commit `ec7ec1bdc79c3794d0c6b2f05d3dcf42d7406da2`
preserved that checkpoint's tree. Subsequent documentation commits do not
change the deployed artifact.

Validation: 316 passing Node tests, seven emulator-dependent skips in that
run; the focused bootstrap/navigation/parameter emulator gate separately
passed 36 tests/groups. Revocation, lack of ACT authority and response fences
were covered. Private execution evidence and deployment overlays stay outside
the repository. The full physical two-person parameter reversal remains a
separate release gate until recorded; merging source does not claim that proof.

CI performs clean dependency installation, build, tests and local release
packaging. It has read-only repository permissions and no deployment step.
Production parameter admissions remain closed. Deploy only an explicitly
approved, identified artifact; never deploy main merely because it advanced.
