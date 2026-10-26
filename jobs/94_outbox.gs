// F4A: native synchronous port of F3.5 jobs/outbox/index.cjs. No external module loading.
function f4Outbox_() {
  if (f4Outbox_.value_) return f4Outbox_.value_;
  return (f4Outbox_.value_ = (function () {
'use strict';
const C=f4Core_();
function intent(operationId){C.need(C.op(operationId),'COMMAND_PAYLOAD_INVALID');return {kind:'LOCAL_DECISION_OBSERVED',state:'PENDING',idempotencyKey:operationId,attempts:0}}
return {intent}; // No dispatcher or external executor exists in F3.

  })());
}
