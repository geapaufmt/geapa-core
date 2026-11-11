// Occupations are explicit identities over canonical term segments, not grants.
// Legacy segments remain independent unless a canonical procedure links them.
function f4MemberOccupations_(){
 const C=f4Core_(),N=C.need,clone=v=>f4Firestore_().decode(f4Firestore_().encode(v));
 const modern=f=>f?.schemaVersion==='m1-memberFunctions-v3';
 const id=f=>modern(f)?f.occupationId:f.id;
 const active=f=>!modern(f)||!f.cancelled;
 const end=f=>f.endDateReal||f.endDatePlanned;
 function members(functions,f){return functions.filter(x=>id(x)===id(f)).sort((a,b)=>a.startDate.localeCompare(b.startDate)||a.id.localeCompare(b.id));}
 function validate(functions){
  for(const f of functions.filter(modern)){
   N(C.id(f.occupationId)&&typeof f.cancelled==='boolean','M1_SCHEMA_INVALID');
   const group=members(functions,f),anchor=group.find(x=>x.id===f.occupationId);
   N(anchor&&modern(anchor)&&!anchor.cancelled,'M1_OCCUPATION_INVALID');
   N(group.every(x=>modern(x)&&['idPessoa','idVinculo','idCargo','scopeKey','idDiretoria'].every(k=>x[k]===anchor[k])),'M1_OCCUPATION_INVALID');
   const spans=group.filter(active);N(spans[0].id===anchor.id,'M1_OCCUPATION_INVALID');
   for(let i=1;i<spans.length;i++)N(C.nextDay(end(spans[i-1]))===spans[i].startDate&&spans[i-1].endDateReal===null,'M1_OCCUPATION_INVALID');
   N(group.filter(x=>x.cancelled).every(x=>x.startDate>end(spans.at(-1))),'M1_OCCUPATION_INVALID');
  }
 }
 function periods(get,boardId,start,finish){
  const board=get('memberGovernanceBoards/'+boardId),head=get('memberBoardHeads/'+boardId);
  N(board&&head&&start>=board.startDate&&finish<=board.endDate&&start<=finish,'TERM_OUTSIDE_BOARD');
  const rels=head.paths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).map(get).filter(r=>r&&r.idDiretoria===boardId&&r.endDate>=start&&r.startDate<=finish).sort((a,b)=>a.startDate.localeCompare(b.startDate));
  let cursor=start;const out=[];
  for(const r of rels){const term=get('memberGovernanceTerms/'+r.idTermo);N(term&&r.startDate>=term.startDate&&r.endDate<=term.endDate,'M1_GOVERNANCE_INTERVAL');const a=r.startDate<start?start:r.startDate,z=r.endDate>finish?finish:r.endDate;N(a===cursor,'M1_GOVERNANCE_AMBIGUOUS');out.push({termId:r.idTermo,startDate:a,endDate:z});cursor=C.nextDay(z);}
  N(cursor===C.nextDay(finish),'M1_GOVERNANCE_AMBIGUOUS');return out;
 }
 function extend(get,functions,source,finish){
  validate(functions);const group=members(functions,source),live=group.filter(active),tail=live.at(-1);
  N(live.every(f=>f.endDateReal===null)&&finish>end(tail),'EXTENSION_INVALID');
  const occupationId=id(source),changed=[];
  // The existing anchor is preserved, including its ID and historical interval.
  for(const f of live)if(!modern(f))changed.push({...clone(f),schemaVersion:'m1-memberFunctions-v3',occupationId,cancelled:false,revision:f.revision+1});
  for(const p of periods(get,source.idDiretoria,C.nextDay(end(tail)),finish)){
   const prior=live.find(f=>f.idTermo===p.termId);
   if(prior){N(prior.id===tail.id,'M1_OCCUPATION_INVALID');const value={...clone(prior),schemaVersion:'m1-memberFunctions-v3',occupationId,cancelled:false,revision:prior.revision+1,endDatePlanned:p.endDate};const index=changed.findIndex(f=>f.id===prior.id);if(index<0)changed.push(value);else changed[index]=value;}
   else{const nextId='VIG_'+C.hash({occupationId,termId:p.termId,startDate:p.startDate}).slice(0,32);N(!get('memberFunctions/'+nextId),'REVISION_CONFLICT');changed.push({...clone(source),schemaVersion:'m1-memberFunctions-v3',id:nextId,origin:'BOARD_PROCEDURE',revision:1,occupationId,cancelled:false,idTermo:p.termId,startDate:p.startDate,endDatePlanned:p.endDate,endDateReal:null});}
  }
  return changed;
 }
 function close(functions,source,finish){
  const group=members(functions,source).filter(active);N(finish>=group[0].startDate&&finish<=end(group.at(-1)),'TERM_OUTSIDE_BOARD');
  return group.filter(f=>end(f)>finish||end(f)===finish&&f.endDateReal===null).map(f=>({...clone(f),revision:f.revision+1,...(f.startDate>finish?{cancelled:true}:{endDateReal:finish})}));
 }
 function summaries(functions){validate(functions);const groups=new Map();for(const f of functions){if(!groups.has(id(f)))groups.set(id(f),members(functions,f));}return [...groups.entries()].map(([occupationId,segments])=>({occupationId,explicit:segments.every(modern),startDate:segments[0].startDate,endDate:end(segments.filter(active).at(-1)),segmentIds:segments.map(f=>f.id)}));}
 return {modern,id,active,end,members,validate,periods,extend,close,summaries};
}
