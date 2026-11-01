// Shared closed transaction plans: tracked read set, budgets, technical effects, one write phase.
function f4TransactionPlans_(){
 if(f4TransactionPlans_.value_)return f4TransactionPlans_.value_;
 const C=f4Core_(),F=f4Firestore_(),N=(ok,code)=>{if(!ok){const e=new Error(code);e.code=code;throw e}},seals=new WeakMap();
 const clone=v=>v===undefined?undefined:F.decode(F.encode(v));
 const bytes=v=>unescape(encodeURIComponent(JSON.stringify(v))).length;
 const hash=v=>C.hash(F.encode(v));
 const limit={writes:400,reads:200,requestBytes:4*1024*1024,documentBytes:512*1024};
 function assertBudget(f,plan){const attendance=plan?.kind==='ACT'&&['ACT.ATTENDANCE_SAVE_DRAFT','ACT.ATTENDANCE_FINALIZE','ACT.ATTENDANCE_ROSTER_APPLY'].includes(plan.commandType);N(f.writes<=limit.writes&&f.reads<=(attendance?360:limit.reads)&&f.requestBytes<=limit.requestBytes&&f.documentBytes<=limit.documentBytes,'AGGREGATE_LIMIT_EXCEEDED');return f}
 function effects(p){return [...p.sourceWrites,...p.headWrites,...p.claimWrites,...p.projectionWrites,...p.technicalWrites,...p.deletes]}
 function finalize(tx,p,result,now,deadline){
  const seen=new Set();for(const w of effects(p)){N(!seen.has(w.path),'M13_PRODUCER_COLLISION');seen.add(w.path);tx.get(w.path);}
  const state=F.inspectTransaction(tx);p.readSet=state.readSet.sort((a,b)=>a.path.localeCompare(b.path));
  const revisionPaths=new Map([...p.sourceWrites,...p.headWrites,...p.technicalWrites].filter(w=>Number.isSafeInteger(w.data?.revision)).map(w=>[w.path,w.data.revision]));
  for(const r of p.readSet)if(r.path==='memberAggregateHeads/'+p.aggregateId&&!revisionPaths.has(r.path)&&r.revision!==null)revisionPaths.set(r.path,r.revision);
  p.revisionBefore=[...revisionPaths].map(([path])=>({path,revision:p.readSet.find(r=>r.path===path).revision??0})).sort((a,b)=>a.path.localeCompare(b.path));
  p.revisionAfter=[...revisionPaths].map(([path,revision])=>({path,revision})).sort((a,b)=>a.path.localeCompare(b.path));
  const encoded=F.previewTransaction(tx,effects(p));
  p.footprint=assertBudget({writes:encoded.writes.length,reads:p.readSet.length,requestBytes:bytes(encoded),documentBytes:Math.max(0,...encoded.writes.map(w=>w.update?bytes(w.update):0))},p);
  const immutable=clone(p);seals.set(immutable,{tx,hash:hash(immutable),result:clone(result),now,deadline});return immutable;
 }
 function appendTechnical(tx,p,writes,validate){const s=seals.get(p);N(s&&s.tx===tx&&s.hash===hash(p)&&typeof validate==='function','M13_PLAN_INVALID');F.inspectTransaction(tx);N(Array.isArray(writes),'M13_PLAN_INVALID');
  for(const w of writes){C.exact(w,['path','data','kind']);N(w.kind==='set'&&/^((commandLedger|commandReceipts|commandTombstones|commandLeases|jobOutbox|memberAudit|memberOnboardingRuns)\/[a-f0-9]{32})$/.test(w.path),'M13_TECHNICAL_PATH_INVALID');N(validate(w)===true,'M13_TECHNICAL_INVALID');}
  const next=clone(p);next.technicalWrites.push(...clone(writes));const result=finalize(tx,next,s.result,s.now,s.deadline);seals.delete(p);return result;
 }
 function applyMutationPlan(tx,p,now){const s=seals.get(p);N(s&&s.tx===tx&&s.hash===hash(p),'M13_PLAN_INVALID');N(Number.isSafeInteger(now)&&now>=s.now&&now<s.deadline,'M13_DEADLINE_EXCEEDED');
  const state=F.inspectTransaction(tx);N(hash(state.readSet.sort((a,b)=>a.path.localeCompare(b.path)))===hash(p.readSet),'M13_READ_SET_CHANGED');
  assertBudget(p.footprint,p);seals.delete(p);for(const w of effects(p)){if(w.kind==='delete')tx.delete(w.path);else tx.set(w.path,w.data)}return clone(s.result);
 }
 return f4TransactionPlans_.value_={finalize,appendTechnical,applyMutationPlan,assertBudget,limits:Object.freeze(limit)};
}

