// Trusted synthetic producer. This is NOT an installed command handler or Firebase verifier.
function m13DFixture_(){
 const x=m13CFixture_(),C=f4Core_(),F=f4Firestore_(),K=f4MemberCommandCatalog_(),T=F.Timestamp.fromMillis,id='d'.repeat(32),ticketId='SYN_TICKET',type='members.self.update';
 const read=m13BFixture_('SELF','MEMBERS_SELF_READ');read.d['memberAuthzPolicies/SYN_POLICY'].id='READ_POLICY';read.d['memberAuthzPolicies/SYN_POLICY'].targets.allowedResources=['memberCurrentStates/PES-000001'];
 x.d['memberAuthzPolicies/READ_POLICY']=read.d['memberAuthzPolicies/SYN_POLICY'];x.d['memberAuthorityControl/MEMBERS'].selfReadPolicy={...read.d['memberAuthorityControl/MEMBERS'].selfReadPolicy,policyId:'READ_POLICY'};
 x.req.admissionExpiresAt=T(x.now+120000);x.req.tokenExpiresAt=T(x.now+3600000);x.context.authorization={request:x.req,proof:f4MemberAuthz_().authorize(x.get,x.req,x.now).proof};
 const r={schemaVersion:'geapa-members-command-v13',operationId:id,ownerUid:'SYN_UID',commandType:type,target:{personId:'PES-000001'},expectedRevisions:[{resourceType:'memberAggregateHeads',id:'PES-000001',revision:1},{resourceType:'memberPrivateProfiles',id:'PES-000001',revision:1}],payload:{patch:{contactEmail:'updated@example.invalid'}},evidenceRefs:[],decisionRef:null,ticketId,digest:'a'.repeat(64),receivedAt:T(x.now),state:'RECEIVED'};r.digest=K.digest(fixture.config.projectId,r);
 x.d['commandRequests/'+id]=r;
 x.d['memberAdmissionTickets/'+ticketId]={schemaVersion:'member-admission-ticket-v13',id:ticketId,operationId:id,ownerUid:r.ownerUid,idPessoa:'PES-000001',requestEnvelope:K.template(r),digest:r.digest,authProofs:[x.context.authorization.proof],issuedAt:T(x.now),expiresAt:T(x.now+120000),tokenExpiresAt:x.req.tokenExpiresAt,authorityEpoch:1,status:'ISSUED'};
 x.d['memberAdmissionUses/'+ticketId]={schemaVersion:'member-admission-use-v13',ticketId,operationId:id,ownerUid:r.ownerUid,reservedAt:r.receivedAt};
 x.d['commandProvenance/'+id]={schemaVersion:'member-command-provenance-v13',operationId:id,ownerUid:r.ownerUid,ticketId,digest:r.digest,authRequests:[x.req]};
 x.d['commandLeases/'+id]={schemaVersion:'member-command-lease-v13',operationId:id,domain:'MEMBERS',workerId:'SYN_WORKER',fencingToken:1,deadline:T(x.now+120000)};
 x.key={operationId:id,digest:r.digest,commandType:type};x.handle={workerId:'SYN_WORKER',fencingToken:1};x.request=r;x.clock=x.now;x.resolveCalls=0;
 x.producers={[type]:{
  validate:r=>{C.exact(r.target,['personId']);C.exact(r.payload,['patch']);C.exact(r.payload.patch,['contactEmail']);C.need(C.id(r.target.personId)&&typeof r.payload.patch.contactEmail==='string'&&/^[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+$/.test(r.payload.patch.contactEmail)&&r.payload.patch.contactEmail.length<=180&&r.evidenceRefs.length===0&&r.decisionRef===null,'COMMAND_PAYLOAD_INVALID');return true},
  resolve:(get,r,n)=>{x.resolveCalls++;const old=get('memberPrivateProfiles/'+r.target.personId);return {...x.context,now:n,deadline:n+60000,changes:[{path:'memberPrivateProfiles/'+r.target.personId,expectedRevision:r.expectedRevisions[1].revision,data:{...old,revision:r.expectedRevisions[1].revision+1,contactEmail:r.payload.patch.contactEmail}}]};}
 }};
 x.make=()=>f4Commands_().createMembersLocal({config:fixture.config,clock:()=>x.clock,verifyIdentity:(token,n)=>{C.need(token==='SYNTHETIC_ONLY','AUTH_BINDING_INVALID');return {uid:'SYN_UID',expiresAt:n+3600000}},producers:x.producers});
 return x;
}

// Another synthetic producer only, to prove event idempotency without implementing E handlers.
function m13DEventFixture_(){
 const z=m13DFixture_(),x=m13BFixture_('EXPLICIT','MEMBERS_DISCIPLINE_EXECUTE'),C=f4Core_(),F=f4Firestore_(),T=F.Timestamp.fromMillis,K=f4MemberCommandCatalog_(),id=z.key.operationId,type='members.discipline.dismiss',path='memberEvents/EV_D';
 const readPolicy=z.d['memberAuthzPolicies/READ_POLICY'],self=z.d['memberAuthorityControl/MEMBERS'].selfReadPolicy;
 x.d['memberAuthzPolicies/READ_POLICY']=readPolicy;x.d['memberAuthorityControl/MEMBERS'].selfReadPolicy=self;x.d['memberAuthorityControl/MEMBERS'].activeWriterVersion='M13';
 x.req.target.resources=['memberLinks/VINC-002',path];x.d[x.grantPath].targets.allowedResources=[...x.req.target.resources];x.d['memberEvidenceAttestations/SYN_EVIDENCE'].targetResourceIds=[...x.req.target.resources];x.req.tokenExpiresAt=T(x.now+3600000);x.req.admissionExpiresAt=T(x.now+120000);
 const proof=f4MemberAuthz_().authorize(x.get,x.req,x.now).proof;
 const r={...z.request,commandType:type,target:{personId:'PES-000002',linkId:'VINC-002'},expectedRevisions:[{resourceType:'memberAggregateHeads',id:'PES-000002',revision:1},{resourceType:'memberEvents',id:'EV_D',revision:0}],payload:{effectiveDate:x.req.target.effectiveDate},evidenceRefs:['SYN_EVIDENCE'],decisionRef:'SYN_APPROVAL'};r.digest=K.digest(fixture.config.projectId,r);
 for(const col of ['memberAdmissionUses','commandLeases']){const key=col==='memberAdmissionUses'?'SYN_TICKET':id;x.d[col+'/'+key]=z.d[col+'/'+key];}
 x.d['commandRequests/'+id]=r;x.d['memberAdmissionTickets/SYN_TICKET']={...z.d['memberAdmissionTickets/SYN_TICKET'],requestEnvelope:K.template(r),digest:r.digest,authProofs:[proof]};x.d['commandProvenance/'+id]={...z.d['commandProvenance/'+id],digest:r.digest,authRequests:[x.req]};
 z.d=x.d;z.request=r;z.key={operationId:id,commandType:type,digest:r.digest};
 z.producers={[type]:{validate:r=>{C.exact(r.target,['personId','linkId']);C.exact(r.payload,['effectiveDate']);C.civilStart(r.payload.effectiveDate);return r.target.personId==='PES-000002'&&r.target.linkId==='VINC-002'&&r.decisionRef==='SYN_APPROVAL'&&r.evidenceRefs.length===1&&r.evidenceRefs[0]==='SYN_EVIDENCE'},resolve:(get,r,n)=>{z.resolveCalls++;return {id:r.target.personId,expected:1,changes:[{path,expectedRevision:0,data:fixture.entity('memberEvents','EV_D',{idPessoa:r.target.personId,idVinculo:r.target.linkId,type:'DESLIGAMENTO',effect:'DESLIGAMENTO',effectiveDate:r.payload.effectiveDate,endDate:null,supersedes:null,correctionPath:null,prior:null})}],now:n,boardExpected:{},mode:'M13',authorization:{request:x.req,proof},deadline:n+60000}}}};
 return z;
}
