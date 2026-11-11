// Historical read model. Source labels are preserved; current office display
// aliases never rewrite historical appointments. No documentary source is a grant.
function f4BoardHistory_(get,o){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),db=F.connect(o.config),now=o.clock(),control=get('institutionalDirectoryControl/current');
 N(control?.enabled&&control.boardHistoryEnabled===true,'PILOT_DISABLED');
 const auth=f4InstitutionalPolicies_(get,o.identity,now,control.policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/governance');
 N(auth.mask.includes('fields')&&auth.mask.includes('canonical'),'AUTHORIZATION_DENIED');
 const query=(collection,filters=[])=>{const result=[];let cursor;for(let page=0;page<10;page++){
  const rows=db.query({collection,filters,limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});result.push(...rows.map(r=>get(r.path)));if(rows.length<20)return result;
  N(page<9,'SCOPE_TOO_LARGE');cursor=new F.DocumentReference('projects/'+o.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+rows.at(-1).path);
 }return result;};
 const boards=query('memberGovernanceBoards'),records=boards.map(board=>{
  const head=get('memberBoardHeads/'+board.id);N(head&&Array.isArray(head.paths),'SCHEMA_INVALID');
  const terms=head.paths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).map(p=>get(p)).map(r=>({associationId:r.id,startDate:r.startDate,endDate:r.endDate,term:get('memberGovernanceTerms/'+r.idTermo)})).map(r=>({associationId:r.associationId,startDate:r.startDate,endDate:r.endDate,id:r.term?.id,name:r.term?.name||r.term?.id}));
  const functions=query('memberFunctions',[{field:'idDiretoria',op:'EQUAL',value:board.id}]).map(f=>{
   const role=get('memberRoleCatalog/'+f.idCargo),person=get('memberIdentities/'+f.idPessoa);N(role&&person,'SCHEMA_INVALID');
   return {personId:f.idPessoa,name:person.displayName,functionId:f.id,linkId:f.idVinculo,roleId:f.idCargo,roleName:role.name||f.idCargo,roleKey:role.roleKey,group:role.group,startDate:f.startDate,endDate:f.endDateReal||f.endDatePlanned,revision:f.revision};
  });
  const metadata=get('memberBoardDocumentaryMetadata/'+board.id);
  // Optional enrichment is accepted only as an explicitly verified source-bound
  // record. It cannot create boards, appointments, dates, votes or permissions.
  const verified=metadata?.schemaVersion==='board-documentary-metadata-v1'&&metadata.boardId===board.id&&metadata.boardRevision===board.revision&&metadata.verification==='VERIFIED'&&/^[a-f0-9]{64}$/.test(metadata.sourceHash||'');
  const start=C.civilStart(board.startDate),end=C.civilStart(C.nextDay(board.endDate));
  return {id:board.id,name:board.name,startDate:board.startDate,endDate:board.endDate,revision:board.revision,status:now<start?'PLANNED':now>=end?'ENDED':'CURRENT',slogan:verified?metadata.slogan||null:null,documents:verified?(metadata.documents||[]).filter(d=>['PLAN','GOALS','INSTITUTIONAL_FILE'].includes(d.kind)&&/^https:\/\/drive.google.com\/file\/d\/[A-Za-z0-9_-]+\/view$/.test(d.url)&&/^[a-f0-9]{64}$/.test(d.sha256)):[],directors: new Set(functions.filter(f=>f.group==='DIRETORIA').map(f=>f.personId)).size,advisors:new Set(functions.filter(f=>f.group==='ASSESSORIA').map(f=>f.personId)).size,presidents:[...new Set(functions.filter(f=>f.roleKey==='PRESIDENTE').map(f=>f.name))],functions,terms,issues:verified?[]:['Enriquecimento documental ainda não vinculado à gestão canônica.']};
 }).sort((a,b)=>a.startDate.localeCompare(b.startDate)||a.id.localeCompare(b.id));
 return {schemaVersion:'board-history-v1',validUntil:auth.validUntil,records,notice:'Somente gestões canônicas cadastradas. Ausência documental não equivale a inexistência histórica.'};
}
