// Small adapters for existing commands. No normative policy is created here.
function f4ActivitiesGovernanceCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function all(collection){let cursor,rows=[];for(let page=0;page<30;page++){const part=db.query({collection,filters:[],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});rows.push(...part.map(r=>r.data));if(part.length<20)return rows;cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+part[part.length-1].path);}N(false,'SCOPE_TOO_LARGE');}
 function read(route){
  const auth=ports.readAuthority(['activityAttendanceRegistry/current'],false);let data;
  if(route==='/admin/ciclos-semestres'){
   data={cycles:all('presentationCycles').map(c=>({id:c.id,revision:c.revision,label:c.cycleLabel,civilStart:c.civilStart,civilEnd:c.civilEnd,status:c.status,academicSemesterIds:(c.academicSemesterRefs||[]).map(r=>r.id),configRevision:c.currentConfigRef?.id||null})),semesters:all('academicSemesters').map(s=>({id:s.id,revision:s.revision,label:s.label,civilStart:s.civilStart,civilEnd:s.civilEnd,status:s.status,cycleId:s.cycleId||null})),policies:all('presentationPolicyReferences').filter(p=>p.usage==='CYCLE'&&p.approved&&p.active).map(p=>({id:p.id,label:p.label||p.id})),canSaveSemester:ports.commandAvailable('ACTIVITY_SEMESTER_SAVE'),canCreateCycle:ports.commandAvailable('CYCLE_CREATE','PRESENTATIONS_PLAN')};
  }else{
   N(route==='/admin/justificativas/janelas','VALIDATION_ERROR');const activities=all('canonicalActivities').filter(a=>a.recordMode!=='HISTORICAL_PARTIAL'&&a.typeSnapshot?.allowsJustification&&a.typeSnapshot.countsAbsence===true&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(a.status));
   const collections=db.getMany(activities.map(a=>'activityAttendanceCollections/'+a.id));
   data={activities:activities.filter((a,i)=>collections[i]?.lastFinalizedAt).map(a=>({id:a.id,label:a.title,civilDate:a.civilDate})),windows:all('activityExceptionalJustificationWindows').map(w=>({id:w.id,revision:w.revision,status:w.status==='OPEN'&&M.ms(w.newDeadline)<=clock()?'EXPIRED':w.status,deadline:M.ms(w.newDeadline),activityIds:w.activityRefs.map(r=>r.id),reasonCode:w.reasonCode,reason:w.reason})),canOpen:ports.commandAvailable('ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN'),canClose:ports.commandAvailable('ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE')};
  }
  ports.recheckAuthority(auth);return {ok:true,data};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload']);const p=input.payload;let commandType,payload,expected;
  if(input.route==='/admin/ciclos-semestres/semestre'){
   C.exact(p,['id','revision','label','civilStart','civilEnd','status','reason']);const semesterRef=M.ref('academicSemester',p.id||'sem_'+input.operationId);expected={ref:semesterRef,revision:p.revision};
   commandType='ACTIVITY_SEMESTER_SAVE';payload={semesterRef,label:p.label,civilStart:p.civilStart,civilEnd:p.civilEnd,status:p.status,reason:p.reason};
  }else if(input.route==='/admin/ciclos-semestres/ciclo'){
   C.exact(p,['id','revision','label','civilStart','civilEnd','semesterIds','policyId','reason']);N(Array.isArray(p.semesterIds),'VALIDATION_ERROR');
   if(p.id){commandType='CYCLE_GOVERNANCE_EDIT';expected={ref:M.ref('cycle',p.id),revision:p.revision};payload={cycleRef:expected.ref,label:p.label,civilStart:p.civilStart,civilEnd:p.civilEnd,academicSemesterRefs:p.semesterIds.map(id=>M.ref('academicSemester',id)),confirmAssociations:true,reason:p.reason};}
   else{N(p.revision===0,'VALIDATION_ERROR');const policy=db.get('presentationPolicyReferences/'+p.policyId);N(policy?.usage==='CYCLE'&&policy.approved&&policy.active,'CONFIG_CAPABILITY_DISABLED');N(typeof p.label==='string'&&p.label.trim()&&p.label.length<=120&&typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST');commandType='CYCLE_CREATE';payload={policyRef:M.ref('policy',p.policyId),cycleLabel:p.label,civilStart:p.civilStart,civilEnd:p.civilEnd,timezone:'America/Cuiaba',academicSemesterRefs:p.semesterIds.map(id=>M.ref('academicSemester',id)),confirmAssociations:true,decisionBasis:{reasonCode:'OPERATIONAL_CYCLE_DATES',reason:p.reason}};}
  }else if(input.route==='/admin/justificativas/janelas/abrir'){
   C.exact(p,['activityIds','deadline','reasonCode','reason']);N(Array.isArray(p.activityIds),'VALIDATION_ERROR');commandType='ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN';payload={activityRefs:p.activityIds.map(id=>M.ref('activity',id)),personRefs:[],newDeadline:p.deadline,reasonCode:p.reasonCode,reason:p.reason};
  }else{
   N(input.route==='/admin/justificativas/janelas/encerrar','VALIDATION_ERROR');C.exact(p,['id','revision','action','reason']);commandType='ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE';expected={ref:M.ref('exceptionalWindow',p.id),revision:p.revision};payload={windowRef:expected.ref,action:p.action,reason:p.reason};
  }
  if(expected){N(Number.isSafeInteger(expected.revision)&&expected.revision>=0,'VALIDATION_ERROR');N((db.get(M.path(expected.ref))?.revision||0)===expected.revision,'STALE_REVISION');}
  const prepared=ports.prepare({operationId:input.operationId,commandType:'ACT.'+commandType,payload});
  if(expected)N(prepared.target.expected.some(e=>M.path({kind:e.kind,id:e.id,...(e.parent?{parent:e.parent}:{})})===M.path(expected.ref)&&(expected.revision?e.revision===expected.revision:e.exists===false)),'STALE_REVISION');
  return prepared;
 }
 return {read,prepare};
}
