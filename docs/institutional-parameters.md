# Institutional parameter contracts

`institutional/204_institutional_parameters.gs` is the canonical operational
parameter authority. The Node adapter transports the same rules; the Portal
only renders the authorized view and submits explicit intentions.

Normative evidence stays within an instrument's related parameters, through
`205_institutional_parameter_evidence.gs`. Documentary values never activate
runtime rules. Unknown consumers and obsolete references remain unresolved.

The first editable resource is only
`institutionalOperationalParameters/ATIVIDADES_DESTACAR_PROXIMA`. Its boolean
value controls the featured section, not the agenda list, visibility or facts.
Future preload and attendance-window settings remain non-editable.

The three supported modes are DIRECT, AWARENESS and APPROVAL. In AWARENESS,
PROPOSE records the intention; another authorized person records ACKNOWLEDGE;
the proposer subsequently ACTIVATEs the same version. Acknowledgement is not
approval. In APPROVAL, a distinct decision capability is required. The initial
policy proposal only includes Presidency and Administration with PROPOSE and
ACKNOWLEDGE; it grants no DECIDE capability.

Each preview and commit verifies current institutional authority, resource and
field scope. Proposal activation revalidates both actors and the exact policy
revisions. Short-lived previews, parameter/proposal revisions, control epoch,
atomic CAS, hourly admission limits and operation digest prevent stale or
duplicate effects. History, ownership, receipt and audit are committed together.
No outbox is needed because this contract has no external side effect.

The client stores only an opaque recovery selector, bound to project and user.
Lost responses must be reconciled before another intention is admitted. Closing
admissions is separate from recovery; recovery still requires valid authority.

Limits: one pilot parameter, at most 16 nominal actors and 40 admissions/hour,
500 history events (fails closed), 30 recent proposals in the consultation.
Pagination must be introduced before sustained operation reaches the history
limit. Anonymous/public consumption is not enabled by this contract.

Deployment does not grant editing rights. A separate canonical control and
explicit policies must authorize each environment. Production stays closed
until the authenticated two-actor DEV flow, activation and reverse operation
are proven and the activation manifest approved. Preserve prior operations and
use a new audited reversal; never restore a snapshot over newer facts.

## Institutional-only bootstrap

Portal navigation does not require an unrelated ACT grant. When the ordinary
ACT context is denied, an enabled institutional directory control and a valid
ROLE summary policy can supply a minimal institutional context. It grants no
ACT flags, facts or commands. The actual directory and parameter endpoints
still revalidate their own permissions. Projection-basis requests cannot use
this fallback. Source revisions are rechecked and the Node response fence
rejects a concurrent revocation before returning the context.
