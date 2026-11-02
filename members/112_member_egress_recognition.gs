// Closed documentary reconciliation planner, not a Portal command or new grant.
// The caller must pin an approved nominal manifest and commit its read set with CAS.
function f4MemberEgressRecognition_(get,entry,now){
 const C=f4Core_(),N=C.need,control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),M=f4MemberModel_(true,control),R=f4MemberRepository_(control),id=entry.personId;
 N(C.id(id)&&C.id(entry.approvalRef)&&Number.isSafeInteger(now),'SCHEMA_INVALID');
 const h=R.planningHelpers().head(get('memberAggregateHeads/'+id),id,'PERSON');N(h.revision===entry.expectedRevision,'REVISION_CONFLICT');
 const docs={};for(const path of h.paths.filter(p=>/^member(Identities|Links|Events|Functions)\//.test(p))){docs[path]=get(path);N(docs[path]?.idPessoa===id,'RECONCILIATION_REQUIRED');}
 const dep=R.planningHelpers().dependencies(docs,get),before=M.derive(docs,id,h.revision,now),prior=docs['memberLinks/'+entry.priorLinkId],person=docs['memberIdentities/'+id];
 N(person?.active&&!person.technicalAccount&&['SEM_VINCULO','DESLIGADO'].includes(before.state.state)&&before.currentFunctions.length===0,'DOMAIN_STATE_INVALID');
 N(prior?.type==='MEMBRO_EFETIVO'&&prior.endDate===entry.exitDate&&prior.endDate!==null&&C.civilStart(C.nextDay(prior.endDate))<=now,'EGRESS_EXIT_REQUIRED');
 N(M.rows(docs,'memberLinks').every(l=>l.endDate!==null&&l.endDate<=prior.endDate),'EGRESS_LATER_LINK');
 const newId='VIN-EGRESS-'+id.replace(/^PES-/,''),eventId='EGR-'+entry.approvalRef+'-'+id.replace(/^PES-/,''),start=C.nextDay(prior.endDate);
 N(!get('memberLinks/'+newId)&&!get('memberEvents/'+eventId),'ALREADY_RECOGNIZED');
 const link={schemaVersion:'m1-memberLinks-v1',id:newId,idPessoa:id,revision:1,origin:entry.approvalRef,type:'EGRESSO',startDate:start,endDate:null,exitDate:prior.endDate};
 const event={schemaVersion:'m1-memberEvents-v1',id:eventId,idPessoa:id,revision:1,origin:entry.approvalRef,idVinculo:newId,type:'EGRESSO',effect:'EGRESSO',effectiveDate:start,endDate:null,supersedes:null,correctionPath:null,prior:null};
 docs['memberLinks/'+newId]=link;docs['memberEvents/'+eventId]=event;
 const after=M.derive(docs,id,h.revision+1,now);N(after.state.state==='EGRESSO'&&after.currentFunctions.length===0&&M.capabilities(docs,after,'ADMINISTRACAO').length===0,'EGRESS_ACCESS_INVALID');
 const documents={['memberLinks/'+newId]:link,['memberEvents/'+eventId]:event,['memberAggregateHeads/'+id]:{...h,revision:h.revision+1,paths:[...h.paths,'memberLinks/'+newId,'memberEvents/'+eventId].sort()},...R.projections(docs,id,h.revision+1,now,dep)};
 f4CoreDomainInvalidation_(get,documents,id,now);
 const fence=get('portalReaderResourceFences/members_operations');documents['portalReaderResourceFences/members_operations']={schemaVersion:'portal-reader-resource-fence-v1',id:'members_operations',revision:(fence?.revision||0)+1,updatedAt:f4Firestore_().Timestamp.fromMillis(now)};
 return {documents,stateBefore:before.state.state,stateAfter:after.state.state,priorLinkId:prior.id,newLinkId:newId,selfCapabilities:M.capabilities(docs,after,'SELF')};
}
