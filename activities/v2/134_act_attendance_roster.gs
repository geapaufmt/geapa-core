// Core-only roster preparation reads canonical Membros at the activity date.
// No identity, authorization, email, CPF or display-name copy is persisted by ACT.
function f4ActivitiesAttendanceRoster_(options){
 const M=f4ActivitiesModel_(),N=M.need,F=f4Firestore_(),C=f4Core_(),db=F.connect(options.config),clock=options.clock||Date.now;
 function prepare(activityRef){
  const activityPath=M.path(activityRef),collectionPath='activityAttendanceCollections/'+activityRef.id,auth=options.ports.readAuthority([activityPath,collectionPath],false),activity=db.get(activityPath),collection=db.get(collectionPath),registry=db.get('presentationCapacityHeads/registry');
  N(activity&&collection&&registry,'UNRESOLVED_DEPENDENCY','ATTENDANCE_CONFIGURATION_REQUIRED');N(['DRAFT','OPEN','REOPENED'].includes(collection.state),'INVALID_STATE_TRANSITION');
  const sourceHeads=[],records={},at=M.ms(activity.startsAt),memberModel=f4MemberModel_(true);let cursor,complete=false;
  // Bounded pages preserve completeness; reaching a technical bound never truncates a roster.
  for(let page=0;page<10;page++){
   const rows=db.query({collection:'memberAggregateHeads',filters:[],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});
   // Read canonical sources in bounded batches for this page, preserving each head fence.
   const pageSources=new Map(),pagePaths=[...new Set(rows.flatMap(({path,data:head})=>{N(head.kind==='PERSON'&&head.id===path.split('/')[1]&&Array.isArray(head.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');return head.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));}))];
   for(let i=0;i<pagePaths.length;i+=180){const paths=pagePaths.slice(i,i+180),values=db.getMany(paths);paths.forEach((p,j)=>pageSources.set(p,values[j]));}
   for(const {path,data:head} of rows){N(head.kind==='PERSON'&&head.id===path.split('/')[1]&&Array.isArray(head.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');sourceHeads.push({path,revision:head.revision});
    const paths=head.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.includes('memberIdentities/'+head.id)&&paths.length<=180,'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCES');
    const values=paths.map(p=>pageSources.get(p)),docs=Object.fromEntries(paths.map((p,i)=>[p,values[i]]));N(values.every(v=>v&&v.idPessoa===head.id),'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCE_MISSING');
    const person=docs['memberIdentities/'+head.id];if(person.technicalAccount)continue;
    const derived=memberModel.derive(docs,head.id,head.revision,at).state;
    if(!['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','SUSPENSO'].includes(derived.state))continue;
    const links=memberModel.rows(docs,'memberLinks').filter(l=>{const range=memberModel.interval(l.startDate,l.endDate);return range.a<=at&&at<range.z});N(links.length===1,'UNRESOLVED_DEPENDENCY','MEMBERS_LINK');
    const personRef=M.ref('person',head.id),id='a_'+M.hash({activity:activityRef.id,person:head.id}).slice(0,32),applicable=derived.state!=='SUSPENSO';
    records[id]={personRef,membershipRef:M.ref('membership',links[0].id),personRevision:head.revision,participantType:'MEMBRO',applicableAtDate:applicable,countsAttendance:applicable&&activity.typeSnapshot.countsAttendance,countsAbsence:applicable&&activity.typeSnapshot.countsAbsence!==false,notApplicableReason:applicable?'':'SITUACAO_NAO_CONTABILIZA_CHAMADA'};
   }
   if(rows.length<20){complete=true;break}cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+rows[rows.length-1].path);
  }
  N(complete&&sourceHeads.length<=180,'SCOPE_TOO_LARGE','ROSTER_DISCOVERY_LIMIT');
  // Invitations carry participation context only. Names and identifiers remain in Membros.
  const guests=activity.guestParticipants||[];N(Array.isArray(guests)&&guests.length<=25,'SCOPE_TOO_LARGE','GUEST_ROSTER_LIMIT');
  for(const guest of guests){
   const personId=guest.personRef?.id,head=sourceHeads.find(h=>h.path==='memberAggregateHeads/'+personId),person=personId&&db.get('memberIdentities/'+personId);
   N(guest.personRef?.kind==='person'&&head&&person?.idPessoa===personId&&!person.technicalAccount&&['PROFESSOR','CONVIDADO','EXTERNO','EGRESSO'].includes(guest.participantType),'UNRESOLVED_DEPENDENCY','GUEST_CORE_IDENTITY_REQUIRED');
   const id='a_'+M.hash({activity:activityRef.id,person:personId}).slice(0,32);
   // A current member retains the membership-derived attendance obligation.
   if(records[id])continue;
   records[id]={personRef:guest.personRef,membershipRef:null,personRevision:head.revision,participantType:guest.participantType,institutionLabel:guest.institutionLabel||'',applicableAtDate:true,countsAttendance:activity.typeSnapshot.countsAttendance,countsAbsence:false,notApplicableReason:''};
  }
  N(Object.keys(records).length>0&&Object.keys(records).length<=50,'SCOPE_TOO_LARGE','ATTENDANCE_ROSTER_LIMIT');
  const id=M.hash({activityRef,activityRevision:activity.revision,membersEpoch:registry.membersEpoch,...(sourceHeads.length<=64?{sourceHeads}:{sourceHeadsDigest:f4ActivitiesCapacityDigest_(sourceHeads)})}).slice(0,32),ref=M.ref('attendanceRoster',id),path=M.path(ref);
  return db.transaction(tx=>{const get=p=>tx.get(p).data();f4SharedAuthz_().recheck(get,auth.q,auth.proof,clock());N(get(activityPath)?.revision===activity.revision&&get(collectionPath)?.revision===collection.revision&&get('presentationCapacityHeads/registry')?.membersEpoch===registry.membersEpoch,'STALE_REVISION','ROSTER_SOURCE_CHANGED');
   const heads=tx.getAll(sourceHeads.map(h=>h.path));sourceHeads.forEach((h,i)=>N(heads[i].data()?.revision===h.revision,'STALE_REVISION','MEMBERS_CHANGED'));
   const old=get(path),document={id,revision:1,schemaVersion:'act-attendance-roster-v1',state:'SEALED',complete:true,activityRef,activityRevision:activity.revision,membersEpoch:registry.membersEpoch,sourceHeads,records,recordedAt:M.stamp(clock())};
   if(old)N(M.hash(old.records)===M.hash(records),'CONFLICT','ROSTER_DIGEST');else tx.set(path,document);return ref;
  });
 }
 return {prepare};
}
