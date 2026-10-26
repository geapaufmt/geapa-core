// F4A: native synchronous port of F3.5 members/index.cjs. No external module loading.
function f4Members_() {
  if (f4Members_.value_) return f4Members_.value_;
  return (f4Members_.value_ = (function () {
'use strict';
const C=f4Core_(),{authorize}=f4Auth_();
function history(get,uid,occurredAt,now,projection){const {c}=authorize(get,uid,'SELF','EGRESS_HISTORY_READ',now);C.need(c.state==='EGRESSO'&&Number.isFinite(C.ms(c.egressEndsAtExclusive))&&C.ms(occurredAt)<C.ms(c.egressEndsAtExclusive),'ACCESS_REVOKED');if(projection)C.need(C.ms(projection.egressEndsAtExclusive)===C.ms(c.egressEndsAtExclusive),'ACCESS_REVOKED');return true}
return {history};

  })());
}
