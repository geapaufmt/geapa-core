// M63-A schema only. Does not read configuration, activate modes or authorize any writer.
function f4MemberAuthority_() {
  function validate(v) {
    const K=f4MemberContracts_(), need=K.need;
    K.exact(v,['schemaVersion'].concat(K.controlFields));
    need(v.schemaVersion==='member-authority-control-v13'&&v.id==='MEMBERS');
    for(const k of ['revision','epoch','policyEpoch','readAuthzEpoch']) need(K.positive(v[k]));
    need(['LEGACY','PAUSED','CANONICAL'].includes(v.mode)&&['NONE','M13'].includes(v.activeWriterVersion)&&K.id(v.approvalRef));
    K.timestamp(v.updatedAt);
    const a=v.nullableFunctionApprovals;
    need(a && Object.getPrototypeOf(a)===Object.prototype && Object.keys(a).length<=3);
    const seen=new Set();
    for(const [id,p] of Object.entries(a)) {
      need(K.id(id));K.exact(p,['caseRef','approvalRef','evidenceHash']);
      need(K.cases.includes(p.caseRef)&&p.approvalRef==='M621-D1'&&typeof p.evidenceHash==='string'&&/^[a-f0-9]{64}$/.test(p.evidenceHash)&&!seen.has(p.caseRef));
      seen.add(p.caseRef);
    }
    const p=v.selfReadPolicy;
    K.exact(p,['policyId','policyRevision','capability','eligibleStates','fieldMasks','issuedAt','expiresAt']);
    need(K.id(p.policyId)&&K.positive(p.policyRevision)&&p.capability==='MEMBERS_SELF_READ');
    need(Array.isArray(p.eligibleStates)&&p.eligibleStates.length>0&&new Set(p.eligibleStates).size===p.eligibleStates.length&&p.eligibleStates.every(s=>['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','EGRESSO'].includes(s)));
    K.exact(p.fieldMasks,['MEMBERS_SELF_READ']);
    const mask=p.fieldMasks.MEMBERS_SELF_READ;
    need(Array.isArray(mask)&&mask.length>0&&new Set(mask).size===mask.length&&mask.every(x=>K.dtoFields.concat(['displayLabel','state','egressEndsAtExclusive','contactEmail','profileDetails']).includes(x)));
    need(K.timestamp(p.issuedAt)<K.timestamp(p.expiresAt));
    return v;
  }
  return {validate};
}
