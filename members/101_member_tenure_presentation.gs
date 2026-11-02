// Read-only presentation over an already authorized, validated aggregate.
// No grant, date extension, source mutation or permanence calculation lives here.
function f4MemberTenurePresentation_(d,id,derived,now){
 const C=f4Core_(),rows=col=>Object.entries(d).filter(([p,v])=>p.startsWith(col+'/')&&(!v.idPessoa||v.idPessoa===id)).map(([,v])=>v);
 const phase=(start,end)=>now<C.civilStart(start)?'FUTURE':end&&now>=C.civilStart(C.nextDay(end))?'ENDED':'CURRENT';
 const current=new Set(derived.currentFunctions.map(f=>f.idVigencia));
 const relations=rows('memberGovernanceBoardTerms');
 // The verified importer partitions one VIGENCIAS_FUNCOES source row as
 // source.ID_VIGENCIA + '_' + term.ID_SEMESTRE_DIRETORIA. Accept that lineage
 // only with exact canonical references and a unique documented term relation.
 function lineage(f){
  const suffix='_'+f.idTermo,t=d['memberGovernanceTerms/'+f.idTermo],b=d['memberGovernanceBoards/'+f.idDiretoria],l=d['memberLinks/'+f.idVinculo];
  const r=relations.filter(r=>r.idDiretoria===f.idDiretoria&&r.idTermo===f.idTermo);
  const end=f.endDateReal||f.endDatePlanned;
  if(f.origin!=='VIGENCIAS_FUNCOES'||!/^VIG-LEG-DIRETORIA-\d+_SD_\d{4}_[12]$/.test(f.id)||!f.id.endsWith(suffix)||!t||!b||!l||l.idPessoa!==id||r.length!==1||!end||f.startDate<t.startDate||end>t.endDate||f.startDate<b.startDate||end>b.endDate||f.startDate<r[0].startDate||end>r[0].endDate)return null;
  return f.id.slice(0,-suffix.length);
 }
 const buckets=new Map();
 for(const f of rows('memberFunctions')){
  const source=lineage(f),key=JSON.stringify([source||f.id,f.idPessoa,f.idVinculo,f.idCargo,f.scopeKey,f.idDiretoria]);
  if(!buckets.has(key))buckets.set(key,[]);buckets.get(key).push(f);
 }
 const tenures=[];
 function emit(group){
  const first=group[0],last=group[group.length-1],role=d['memberRoleCatalog/'+first.idCargo],board=d['memberGovernanceBoards/'+first.idDiretoria],end=last.endDateReal||last.endDatePlanned||null;
  const periodPhase=phase(first.startDate,end),status=group.some(f=>current.has(f.id))?'CURRENT':periodPhase==='CURRENT'?'INACTIVE':periodPhase;
  tenures.push({id:first.id,idVinculo:first.idVinculo,roleKey:role.roleKey,group:role.group,idDiretoria:first.idDiretoria,boardName:board.name,boardStartDate:board.startDate,boardEndDate:board.endDate,startDate:first.startDate,endDate:end,status,sourceKind:group.length>1?'DOCUMENTED_SOURCE_PARTITION':'INDIVIDUAL_RECORD',segments:group.map(f=>({idVigencia:f.id,idTermo:f.idTermo,startDate:f.startDate,endDate:f.endDateReal||f.endDatePlanned||null,status:phase(f.startDate,f.endDateReal||f.endDatePlanned),revision:f.revision}))});
 }
 for(const group of buckets.values()){
  group.sort((a,b)=>a.startDate.localeCompare(b.startDate)||a.id.localeCompare(b.id));let run=[];
  for(const f of group){const prev=run[run.length-1];
   const consecutive=prev&&!prev.endDateReal&&prev.idTermo!==f.idTermo&&prev.endDatePlanned===d['memberGovernanceTerms/'+prev.idTermo]?.endDate&&f.startDate===d['memberGovernanceTerms/'+f.idTermo]?.startDate&&C.nextDay(prev.endDatePlanned)===f.startDate;
   if(prev&&!consecutive){emit(run);run=[];}run.push(f);
  }if(run.length)emit(run);
 }
 tenures.sort((a,b)=>b.startDate.localeCompare(a.startDate)||a.id.localeCompare(b.id));
 const links=rows('memberLinks').map(v=>{const periodStatus=phase(v.startDate,v.endDate);return {id:v.id,type:v.type,startDate:v.startDate,endDate:v.endDate,status:periodStatus,state:periodStatus==='CURRENT'?derived.state.state:null};}).sort((a,b)=>b.startDate.localeCompare(a.startDate)||a.id.localeCompare(b.id));
 return {schemaVersion:'member-tenure-presentation-v1',asOf:new Date(now).toISOString(),links,tenures};
}
