// F4A: canonical editorial domain, commands, projections and read view.
function f4Editorial_() {
  if (f4Editorial_.value_) return f4Editorial_.value_;
  return (f4Editorial_.value_ = (function () {
'use strict';
const C=f4Core_(),P=f4Projections_();const {Timestamp}=f4Firestore_();
function decide(get,r,auth,now){C.need(auth.p.technicalAccount===false&&auth.account.human===true,'CAPABILITY_REQUIRED');const path='canonicalEditorial/'+r.payload.pageId,old=get(path);C.need(old&&old.draft===true&&old.page.id===r.payload.pageId,'DOMAIN_STATE_INVALID');C.expectedRevision(r.expectedRevision,old.revision);const page=old.page;C.need(page.revision===old.revision,'DOMAIN_STATE_INVALID');const checked=P.publicPages({sourceRevision:old.revision,data:page}).dto;const candidate={id:checked.id,schemaVersion:checked.schemaVersion,published:true,sanitized:true,title:checked.title,subtitle:checked.subtitle,sections:checked.sections,sourceHash:checked.sourceHash,sourceUpdatedAt:checked.sourceUpdatedAt,updatedAt:Timestamp.fromMillis(now),revision:old.revision+1};const projection=P.publicPages({sourceRevision:candidate.revision,data:candidate});return {path,data:{draft:false,revision:candidate.revision,page:projection.dto},before:old.revision,after:candidate.revision,projectionPath:'publicPages/'+candidate.id,projection}}
return {decide};
  })());
}

// Runtime Gen2 editorial authority: canonical drafts, published projections,
// timeline source, immutable version history, CAS and idempotency receipts.
function f4EditorialCommands_(){
 const C=f4Core_(),F=f4Firestore_(),N=f4MemberAuthzContracts_().need,P=f4Projections_(),A=f4Public_(),T=f4TransactionPlans_();
 const CONTROL='editorialControl/current',TYPE='editorial.public.manage',CAP='PUBLIC_CONTENT_MANAGE';
 const PAGE_IDS=['home','about'],SECTION_TYPES=['TEXT','HERO','CTA','IMAGE_TEXT','DOCUMENT_LIST'];
 const ACTIONS=['PAGE_SAVE_DRAFT','PAGE_PUBLISH','PAGE_UNPUBLISH','TIMELINE_SAVE_DRAFT','TIMELINE_PUBLISH','TIMELINE_UNPUBLISH','TIMELINE_REORDER'];
 const hash=v=>C.hash(F.encode(v)),hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v);
 const pagePath=id=>'canonicalEditorial/'+id,timelineIndexPath='canonicalEditorial/timeline',timelineItemPath=id=>'canonicalEditorial/timeline_'+id,historyPath=k=>'canonicalEditorialHistory/'+k;
 const tsMillis=v=>v==null?null:typeof v==='number'?v:typeof v?.toMillis==='function'?v.toMillis():Number.isFinite(Date.parse(String(v)))?Date.parse(String(v)):null;

 function control(get,config,identity,now,admit,resourcePath,inspection=false){
  C.destination(config);const c=get(CONTROL),actual=f4CommandExecutorId_();
  const expected=inspection&&actual==='CANONICAL_READ_ONLY'?(f4ExecutionEnvironment_().production?'CF_GEN2_PROD_V1':'CF_GEN2_DEV_EDIT_V1'):actual;
  N(c?.schemaVersion==='editorial-control-v1'&&c.projectId===config.projectId&&c.executor===expected&&c.commandType===TYPE&&C.rev(c.revision)&&c.revision>0&&C.rev(c.epoch)&&c.epoch>0&&Array.isArray(c.actorUids)&&c.actorUids.length>0&&c.actorUids.length<=32&&c.actorUids.includes(identity.uid),'ACCESS_REVOKED');
  N(c.rateLimit?.windowMs===3600000&&Number.isInteger(c.rateLimit.maxAdmissions)&&c.rateLimit.maxAdmissions>0&&c.rateLimit.maxAdmissions<=120,'ACCESS_REVOKED');
  if(admit)N(c.enabled===true&&f4ServiceAvailability_().live(c,now),'PILOT_DISABLED');
  const auth=f4InstitutionalPolicies_(get,identity,now,c.policyIds,CAP,resourcePath);
  const requiredMask=resourcePath==='canonicalEditorial/timeline'?'timeline':'pages';
  N(auth.mask.includes(requiredMask),'AUTHORIZATION_DENIED');
  return {c,auth};
 }

 function sanitizeSectionInput(raw,index){
  N(raw&&typeof raw==='object'&&!Array.isArray(raw),'COMMAND_PAYLOAD_INVALID');
  C.exact(raw,['id','type','title','subtitle','text','imageUrl','buttonText','buttonUrl','order','published']);
  const type=String(raw.type||'TEXT');
  N(SECTION_TYPES.includes(type)&&typeof raw.published==='boolean','COMMAND_PAYLOAD_INVALID');
  const id=A.text(raw.id||('sec_'+(index+1)),100);
  N(/^[A-Za-z0-9_-]{1,100}$/.test(id),'COMMAND_PAYLOAD_INVALID');
  const rawImg=String(raw.imageUrl??'').trim(),imageUrl=rawImg?A.url(rawImg,true):'';
  if(rawImg)N(!!imageUrl,'COMMAND_PAYLOAD_INVALID');
  const rawBtn=String(raw.buttonUrl??'').trim(),buttonUrl=rawBtn?A.url(rawBtn,false):'';
  if(rawBtn)N(!!buttonUrl,'COMMAND_PAYLOAD_INVALID');
  const order=Number.isSafeInteger(Number(raw.order))?Number(raw.order):(index+1);
  return {id,type,title:A.text(raw.title,240),subtitle:A.text(raw.subtitle,600),text:type==='DOCUMENT_LIST'?'':A.text(raw.text,8000),imageUrl,buttonText:A.text(raw.buttonText,100),buttonUrl,order,published:raw.published===true};
 }

 function sanitizePageDraft(pageId,raw){
  N(PAGE_IDS.includes(pageId)&&raw&&typeof raw==='object'&&!Array.isArray(raw),'COMMAND_PAYLOAD_INVALID');
  C.exact(raw,['title','subtitle','sections']);
  N(typeof raw.title==='string'&&typeof raw.subtitle==='string'&&Array.isArray(raw.sections),'COMMAND_PAYLOAD_INVALID');
  N(raw.sections.length>0&&raw.sections.length<=30,'COMMAND_PAYLOAD_INVALID');
  const title=A.text(raw.title,240);N(title.length>0,'COMMAND_PAYLOAD_INVALID');
  const subtitle=A.text(raw.subtitle,600);
  const seen=new Set();
  const sections=raw.sections.map((s,i)=>{const clean=sanitizeSectionInput(s,i);N(!seen.has(clean.id),'COMMAND_PAYLOAD_INVALID');seen.add(clean.id);return {...clean,order:i+1};});
  return {title,subtitle,sections};
 }

 function sanitizeTimelineDraft(raw,fallbackOrder){
  N(raw&&typeof raw==='object'&&!Array.isArray(raw),'COMMAND_PAYLOAD_INVALID');
  C.exact(raw,['year','date','title','description','imageUrl','order']);
  const year=Number(raw.year);N(Number.isInteger(year)&&year>=1900&&year<=2200,'COMMAND_PAYLOAD_INVALID');
  const rawDate=String(raw.date??'').trim(),date=rawDate?A.civilDate(rawDate):'';
  if(rawDate)N(!!date,'COMMAND_PAYLOAD_INVALID');
  const title=A.text(raw.title,240);N(title.length>0,'COMMAND_PAYLOAD_INVALID');
  const description=A.text(raw.description,8000);N(description.length>0,'COMMAND_PAYLOAD_INVALID');
  const rawImg=String(raw.imageUrl??'').trim(),imageUrl=rawImg?A.url(rawImg,true):'';
  if(rawImg)N(!!imageUrl,'COMMAND_PAYLOAD_INVALID');
  const order=Number(raw.order??fallbackOrder);N(Number.isInteger(order)&&order>=0&&order<=10000,'COMMAND_PAYLOAD_INVALID');
  return {year,date,title,description,imageUrl,order};
 }

 function emptyPage(id){
  return {schemaVersion:'canonical-editorial-page-v1',id,revision:0,draftRevision:0,publishedRevision:0,hasDraft:false,published:false,draft:null,publishedSnapshot:null,updatedAt:null,updatedByPersonId:null};
 }
 function emptyTimelineIndex(){
  return {schemaVersion:'canonical-editorial-timeline-index-v1',revision:0,itemIds:[],updatedAt:null,updatedByPersonId:null};
 }

 function validatePageCanonical(v,id){
  if(!v)return emptyPage(id);
  // Support legacy canonicalEditorial/{id} envelope if seeded before Fase 2.
  if(Object.hasOwn(v,'page')&&Object.hasOwn(v,'draft')){
   const p=v.page||{};const isDraft=v.draft===true;
   return {schemaVersion:'canonical-editorial-page-v1',id,revision:Number(v.revision||0),draftRevision:isDraft?Number(v.revision||0):0,publishedRevision:p.published===true?Number(v.revision||0):0,hasDraft:isDraft,published:p.published===true,draft:isDraft?{title:A.text(p.title,240),subtitle:A.text(p.subtitle,600),sections:Array.isArray(p.sections)?p.sections:[]}:null,publishedSnapshot:p.published===true?p:null,updatedAt:tsMillis(p.updatedAt),updatedByPersonId:null};
  }
  N(v.schemaVersion==='canonical-editorial-page-v1'&&v.id===id&&C.rev(v.revision),'SCHEMA_INVALID');
  return v;
 }

 function validateTimelineIndex(v){
  if(!v)return emptyTimelineIndex();
  N(v.schemaVersion==='canonical-editorial-timeline-index-v1'&&C.rev(v.revision)&&Array.isArray(v.itemIds)&&v.itemIds.length<=100&&v.itemIds.every(C.id)&&new Set(v.itemIds).size===v.itemIds.length,'SCHEMA_INVALID');
  return v;
 }

 function validateTimelineItemCanonical(v,id){
  if(!v)return null;
  N(v.schemaVersion==='canonical-editorial-timeline-item-v1'&&v.id===id&&C.rev(v.revision),'SCHEMA_INVALID');
  return v;
 }

 function preview(tx,o,input,operationId=null){
  N(input&&typeof input==='object'&&!Array.isArray(input)&&ACTIONS.includes(input.action),'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data(),now=o.clock();
  const isTimeline=input.action.startsWith('TIMELINE_');
  const resourcePath=isTimeline?'canonicalEditorial/timeline':'canonicalEditorial/pages';
  const {c,auth}=control(get,o.config,o.identity,now,true,resourcePath);
  const validUntil=Math.min(auth.validUntil,f4ServiceAvailability_().bound(c,now+60000));

  if(!isTimeline){
   C.exact(input,['action','pageId','expectedRevision',...(input.action==='PAGE_SAVE_DRAFT'?['content']:[])]);
   N(PAGE_IDS.includes(input.pageId)&&C.rev(input.expectedRevision),'COMMAND_PAYLOAD_INVALID');
   const doc=validatePageCanonical(get(pagePath(input.pageId)),input.pageId);
   N(doc.revision===input.expectedRevision,'REVISION_CONFLICT');
   let summary='',draft=doc.draft;
   if(input.action==='PAGE_SAVE_DRAFT'){
    draft=sanitizePageDraft(input.pageId,input.content);
    summary='Rascunho da página '+input.pageId+' salvo ('+draft.sections.length+' blocos).';
   }else if(input.action==='PAGE_PUBLISH'){
    const candidate=doc.draft||doc.publishedSnapshot;
    const pubSections=candidate&&Array.isArray(candidate.sections)?candidate.sections.filter(s=>s.published===true):[];
    N(candidate&&pubSections.length>0,'DOMAIN_STATE_INVALID');
    // Pre-validate publicPages DTO before commit
    const nextRev=doc.revision+1,ts=F.Timestamp.fromMillis(now);
    const pubData={id:input.pageId,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,title:candidate.title,subtitle:candidate.subtitle||'',sections:pubSections,sourceHash:hash(candidate),sourceUpdatedAt:ts,updatedAt:ts,revision:nextRev};
    P.publicPages({sourceRevision:nextRev,data:pubData});
    summary='Publicação da página '+input.pageId+' para revisão '+nextRev+'.';
   }else{
    N(doc.published===true,'DOMAIN_STATE_INVALID');
    summary='Despublicação da página '+input.pageId+'.';
   }
   return {action:input.action,targetId:input.pageId,resourcePath,beforeRevision:doc.revision,afterRevision:doc.revision+1,summary,draft,documentHash:hash({control:{revision:c.revision,epoch:c.epoch},doc}),epoch:c.epoch,controlRevision:c.revision,validUntil,publicUpdate:input.action!=='PAGE_SAVE_DRAFT'};
  }

  const idx=validateTimelineIndex(get(timelineIndexPath));
  if(input.action==='TIMELINE_REORDER'){
   C.exact(input,['action','expectedIndexRevision','items']);
   N(C.rev(input.expectedIndexRevision)&&Array.isArray(input.items)&&input.items.length===idx.itemIds.length&&idx.itemIds.length>0,'COMMAND_PAYLOAD_INVALID');
   N(idx.revision===input.expectedIndexRevision,'REVISION_CONFLICT');
   const seen=new Set(),items=[];
   if(typeof tx.getAll==='function')tx.getAll(idx.itemIds.map(timelineItemPath));
   for(const entry of input.items){
    C.exact(entry,['id','expectedRevision']);N(C.id(entry.id)&&C.rev(entry.expectedRevision)&&!seen.has(entry.id)&&idx.itemIds.includes(entry.id),'COMMAND_PAYLOAD_INVALID');
    seen.add(entry.id);
    const item=validateTimelineItemCanonical(get(timelineItemPath(entry.id)),entry.id);N(item,'RESOURCE_NOT_FOUND');
    N(item.revision===entry.expectedRevision,'REVISION_CONFLICT');
    items.push(item);
   }
   return {action:input.action,targetId:'timeline',resourcePath,beforeRevision:idx.revision,afterRevision:idx.revision+1,summary:'Reordenação de '+items.length+' marco(s) históricos.',itemIds:input.items.map(x=>x.id),documentHash:hash({control:{revision:c.revision,epoch:c.epoch},idx,items}),epoch:c.epoch,controlRevision:c.revision,validUntil,publicUpdate:items.some(i=>i.published===true)};
  }

  C.exact(input,['action','itemId','expectedRevision',...(input.action==='TIMELINE_SAVE_DRAFT'?['expectedIndexRevision','content']:[])]);
  N(C.rev(input.expectedRevision),'COMMAND_PAYLOAD_INVALID');
  const creating=input.action==='TIMELINE_SAVE_DRAFT'&&(input.itemId===null||input.itemId==='');
  if(creating){
   N(C.rev(input.expectedIndexRevision)&&idx.itemIds.length<100,'COMMAND_PAYLOAD_INVALID');
   N(idx.revision===input.expectedIndexRevision,'REVISION_CONFLICT');
   N(input.expectedRevision===0,'REVISION_CONFLICT');
  }else{
   N(C.id(input.itemId),'COMMAND_PAYLOAD_INVALID');
   if(input.action==='TIMELINE_SAVE_DRAFT')N(C.rev(input.expectedIndexRevision)&&idx.revision===input.expectedIndexRevision,'REVISION_CONFLICT');
  }
  const targetId=creating?(operationId?'marco_'+operationId.slice(0,12):'marco_preview'):input.itemId;
  const rawExisting=get(timelineItemPath(targetId));
  if(creating)N(!rawExisting,'OPERATION_CONFLICT');
  const oldItem=creating?null:validateTimelineItemCanonical(rawExisting,targetId);
  if(!creating)N(oldItem,'RESOURCE_NOT_FOUND');
  const currentRev=oldItem?oldItem.revision:0;
  N(currentRev===input.expectedRevision,'REVISION_CONFLICT');

  let summary='',draft=oldItem?.draft||null;
  if(input.action==='TIMELINE_SAVE_DRAFT'){
   draft=sanitizeTimelineDraft(input.content,idx.itemIds.length+1);
   summary=(creating?'Criação':'Atualização')+' de rascunho do marco '+draft.year+' — '+draft.title+'.';
  }else if(input.action==='TIMELINE_PUBLISH'){
   const candidate=oldItem.draft||oldItem.publishedSnapshot;N(candidate,'DOMAIN_STATE_INVALID');
   const nextRev=oldItem.revision+1,ts=F.Timestamp.fromMillis(now);
   P.publicTimeline({sourceRevision:nextRev,data:{id:targetId,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,year:candidate.year,date:candidate.date||'',title:candidate.title,description:candidate.description,imageUrl:candidate.imageUrl||'',order:candidate.order,sourceHash:hash(candidate),updatedAt:ts,revision:nextRev}});
   summary='Publicação do marco histórico '+candidate.year+' — '+candidate.title+'.';
  }else{
   N(oldItem.published===true,'DOMAIN_STATE_INVALID');
   summary='Despublicação do marco histórico '+targetId+'.';
  }
  return {action:input.action,targetId,creating,resourcePath,beforeRevision:currentRev,afterRevision:currentRev+1,summary,draft,documentHash:hash({control:{revision:c.revision,epoch:c.epoch},idx,item:oldItem}),epoch:c.epoch,controlRevision:c.revision,validUntil,publicUpdate:input.action!=='TIMELINE_SAVE_DRAFT'};
 }

 function digest(config,uid,r){return hash({projectId:config.projectId,ownerUid:uid,commandType:TYPE,...r});}

 function status(tx,o,key){
  C.exact(key,['operationId','digest']);N(C.op(key.operationId)&&hex(key.digest),'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data(),r=get('editorialReceipts/'+key.operationId),owner=get('commandExecutorOwners/'+key.operationId);
  if(!r){
   let authorized=false;
   for(const resPath of ['canonicalEditorial/pages','canonicalEditorial/timeline']){
    try{control(get,o.config,o.identity,o.clock(),false,resPath);authorized=true;break;}catch(e){if(!['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
   }
   N(authorized,'AUTHORIZATION_DENIED');N(!owner,'RECONCILIATION_REQUIRED');
   return {state:'NOT_ADMITTED',operationId:key.operationId};
  }
  const resPath=r.resourcePath||'canonicalEditorial/pages';
  const {c}=control(get,o.config,o.identity,o.clock(),false,resPath);
  N(r.ownerUid===o.identity.uid&&r.digest===key.digest&&r.projectId===o.config.projectId&&r.executor===c.executor&&r.state==='COMPLETED'&&owner?.digest===r.digest&&owner.ownerUid===r.ownerUid&&owner.executor===r.executor&&owner.commandType===TYPE,'ACCESS_REVOKED');
  return {state:'COMPLETED',operationId:r.operationId,action:r.action,targetId:r.targetId,revision:r.revisionAfter,published:r.published,hasDraft:r.hasDraft,publicUpdate:r.publicUpdate};
 }

 function appendHistory(get,targetKey,event){
  const versionPath=id=>historyPath('version_'+id);
  const h=get(historyPath(targetKey))||{schemaVersion:'canonical-editorial-history-v1',targetId:targetKey,events:[]};
  N(h.schemaVersion==='canonical-editorial-history-v1'&&Array.isArray(h.events),'SCHEMA_INVALID');
  const events=[...h.events,event];
  const inline=events.filter(e=>Object.hasOwn(e,'snapshot'));
  const writes=inline.map(e=>{
   N(C.op(e.operationId)&&hash(e.snapshot)===e.snapshotHash,'SCHEMA_INVALID');
   const path=versionPath(e.operationId);
   N(!get(path),'OPERATION_CONFLICT');
   return {path,kind:'set',data:{schemaVersion:'canonical-editorial-version-v1',historyTarget:targetKey,event:e}};
  });
  const summaries=events.slice(-100).map(({snapshot,...summary})=>summary);
  writes.push({path:historyPath(targetKey),kind:'set',data:{schemaVersion:'canonical-editorial-history-v1',targetId:targetKey,events:summaries}});
  return writes;
 }

 function execute(tx,o,r){
  C.exact(r,['operationId','input','documentHash','epoch','controlRevision','validUntil','digest']);
  N(C.op(r.operationId)&&hex(r.documentHash)&&hex(r.digest),'COMMAND_PAYLOAD_INVALID');
  const {digest:given,...semantic}=r;N(given===digest(o.config,o.identity.uid,semantic),'OPERATION_CONFLICT');
  const get=p=>tx.get(p).data(),now=o.clock();
  if(get('editorialReceipts/'+r.operationId))return status(tx,o,{operationId:r.operationId,digest:r.digest});
  const q=preview(tx,o,r.input,r.operationId),{c,auth}=control(get,o.config,o.identity,now,true,q.resourcePath);
  N(q.documentHash===r.documentHash&&c.epoch===r.epoch&&c.revision===r.controlRevision,'REVISION_CONFLICT');
  N(Number.isSafeInteger(r.validUntil)&&now<r.validUntil&&r.validUntil<=now+60000,'ACCESS_REVOKED');
  for(const prefix of ['commandExecutorOwners','commandRequests','commandReceipts','commandLedger','editorialAudit'])N(!get(prefix+'/'+r.operationId),'OPERATION_CONFLICT');
  const bucket='editorialRate/'+C.hash(o.identity.uid).slice(0,32),rate=get(bucket),start=Math.floor(now/3600000)*3600000,count=rate?.windowStart===start?rate.count:0;
  N(Number.isInteger(count)&&count>=0&&count<c.rateLimit.maxAdmissions,'PILOT_LIMIT_REACHED');

  const actorPersonId=auth.actor.actor.h.id,stamp=F.Timestamp.fromMillis(now);
  const sourceWrites=[],projectionWrites=[],deletes=[];
  let resultPublished=false,resultHasDraft=false,snapshotForHistory=null;

  if(!q.action.startsWith('TIMELINE_')){
   const old=validatePageCanonical(get(pagePath(q.targetId)),q.targetId),nextRev=old.revision+1;
   let nextDoc;
   if(q.action==='PAGE_SAVE_DRAFT'){
    const draftObj={...q.draft,sourceHash:hash(q.draft),updatedAt:now,updatedByPersonId:actorPersonId};
    nextDoc={schemaVersion:'canonical-editorial-page-v1',id:q.targetId,revision:nextRev,draftRevision:nextRev,publishedRevision:old.publishedRevision||0,hasDraft:true,published:old.published===true,draft:draftObj,publishedSnapshot:old.publishedSnapshot||null,updatedAt:now,updatedByPersonId:actorPersonId};
    resultPublished=nextDoc.published;resultHasDraft=true;snapshotForHistory=draftObj;
   }else if(q.action==='PAGE_PUBLISH'){
    const base=old.draft||old.publishedSnapshot;
    const pubSections=base.sections.filter(s=>s.published===true);
    const pubSource={id:q.targetId,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,title:base.title,subtitle:base.subtitle||'',sections:pubSections,sourceHash:hash(base),sourceUpdatedAt:stamp,updatedAt:stamp,revision:nextRev};
    const proj=P.publicPages({sourceRevision:nextRev,data:pubSource});
    const pubSnapshot={title:proj.dto.title,subtitle:proj.dto.subtitle,sections:proj.dto.sections,sourceHash:proj.dto.sourceHash,revision:nextRev,updatedAt:now,updatedByPersonId:actorPersonId};
    nextDoc={schemaVersion:'canonical-editorial-page-v1',id:q.targetId,revision:nextRev,draftRevision:old.draftRevision||nextRev,publishedRevision:nextRev,hasDraft:false,published:true,draft:old.draft||pubSnapshot,publishedSnapshot:pubSnapshot,updatedAt:now,updatedByPersonId:actorPersonId};
    projectionWrites.push({path:'publicPages/'+q.targetId,kind:'set',data:proj.dto});
    resultPublished=true;resultHasDraft=false;snapshotForHistory=pubSnapshot;
   }else{
    nextDoc={schemaVersion:'canonical-editorial-page-v1',id:q.targetId,revision:nextRev,draftRevision:old.draftRevision||old.revision,publishedRevision:0,hasDraft:true,published:false,draft:old.draft||old.publishedSnapshot,publishedSnapshot:null,updatedAt:now,updatedByPersonId:actorPersonId};
    deletes.push({path:'publicPages/'+q.targetId,kind:'delete',data:null});
    resultPublished=false;resultHasDraft=true;snapshotForHistory={unpublishedRevision:old.publishedRevision};
   }
   sourceWrites.push({path:pagePath(q.targetId),kind:'set',data:nextDoc});
   sourceWrites.push(...appendHistory(get,q.targetId,{operationId:r.operationId,action:q.action,targetId:q.targetId,revisionBefore:old.revision,revisionAfter:nextRev,actorPersonId,actorUid:o.identity.uid,at:now,summary:q.summary,snapshotHash:hash(snapshotForHistory),snapshot:snapshotForHistory}));
  }else if(q.action==='TIMELINE_REORDER'){
   const idx=validateTimelineIndex(get(timelineIndexPath)),nextIdxRev=idx.revision+1;
   q.itemIds.forEach((id,position)=>{
    const item=validateTimelineItemCanonical(get(timelineItemPath(id)),id),nextItemRev=item.revision+1,order=position+1;
    const nextDraft=item.draft?{...item.draft,order,updatedAt:now,updatedByPersonId:actorPersonId}:null;
    let nextPub=item.publishedSnapshot?{...item.publishedSnapshot,order,revision:nextItemRev,updatedAt:now,updatedByPersonId:actorPersonId}:null;
    if(item.published===true&&nextPub){
     const proj=P.publicTimeline({sourceRevision:nextItemRev,data:{id,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,year:nextPub.year,date:nextPub.date||'',title:nextPub.title,description:nextPub.description,imageUrl:nextPub.imageUrl||'',order,sourceHash:nextPub.sourceHash||hash(nextPub),updatedAt:stamp,revision:nextItemRev}});
     projectionWrites.push({path:'publicTimeline/'+id,kind:'set',data:proj.dto});
     nextPub={...nextPub,order,revision:nextItemRev};
    }
    sourceWrites.push({path:timelineItemPath(id),kind:'set',data:{...item,revision:nextItemRev,publishedRevision:item.published===true?nextItemRev:item.publishedRevision,draft:nextDraft,publishedSnapshot:nextPub,updatedAt:now,updatedByPersonId:actorPersonId}});
   });
   sourceWrites.push({path:timelineIndexPath,kind:'set',data:{schemaVersion:'canonical-editorial-timeline-index-v1',revision:nextIdxRev,itemIds:q.itemIds,updatedAt:now,updatedByPersonId:actorPersonId}});
   snapshotForHistory={itemIds:q.itemIds};
   sourceWrites.push(...appendHistory(get,'timeline',{operationId:r.operationId,action:q.action,targetId:'timeline',revisionBefore:idx.revision,revisionAfter:nextIdxRev,actorPersonId,actorUid:o.identity.uid,at:now,summary:q.summary,snapshotHash:hash(snapshotForHistory),snapshot:snapshotForHistory}));
  }else{
   const idx=validateTimelineIndex(get(timelineIndexPath));
   const old=q.creating?null:validateTimelineItemCanonical(get(timelineItemPath(q.targetId)),q.targetId);
   const oldRev=old?old.revision:0,nextRev=oldRev+1;
   let nextDoc;
   if(q.action==='TIMELINE_SAVE_DRAFT'){
    const draftObj={...q.draft,sourceHash:hash(q.draft),updatedAt:now,updatedByPersonId:actorPersonId};
    nextDoc={schemaVersion:'canonical-editorial-timeline-item-v1',id:q.targetId,revision:nextRev,draftRevision:nextRev,publishedRevision:old?.publishedRevision||0,hasDraft:true,published:old?.published===true,draft:draftObj,publishedSnapshot:old?.publishedSnapshot||null,updatedAt:now,updatedByPersonId:actorPersonId};
    if(q.creating){
     sourceWrites.push({path:timelineIndexPath,kind:'set',data:{schemaVersion:'canonical-editorial-timeline-index-v1',revision:idx.revision+1,itemIds:[...idx.itemIds,q.targetId],updatedAt:now,updatedByPersonId:actorPersonId}});
    }
    resultPublished=nextDoc.published;resultHasDraft=true;snapshotForHistory=draftObj;
   }else if(q.action==='TIMELINE_PUBLISH'){
    const base=old.draft||old.publishedSnapshot;
    const pubSource={id:q.targetId,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,year:base.year,date:base.date||'',title:base.title,description:base.description,imageUrl:base.imageUrl||'',order:base.order,sourceHash:hash(base),updatedAt:stamp,revision:nextRev};
    const proj=P.publicTimeline({sourceRevision:nextRev,data:pubSource});
    const pubSnapshot={year:proj.dto.year,date:proj.dto.date,title:proj.dto.title,description:proj.dto.description,imageUrl:proj.dto.imageUrl,order:proj.dto.order,sourceHash:proj.dto.sourceHash,revision:nextRev,updatedAt:now,updatedByPersonId:actorPersonId};
    nextDoc={schemaVersion:'canonical-editorial-timeline-item-v1',id:q.targetId,revision:nextRev,draftRevision:old.draftRevision||nextRev,publishedRevision:nextRev,hasDraft:false,published:true,draft:old.draft||pubSnapshot,publishedSnapshot:pubSnapshot,updatedAt:now,updatedByPersonId:actorPersonId};
    projectionWrites.push({path:'publicTimeline/'+q.targetId,kind:'set',data:proj.dto});
    resultPublished=true;resultHasDraft=false;snapshotForHistory=pubSnapshot;
   }else{
    nextDoc={schemaVersion:'canonical-editorial-timeline-item-v1',id:q.targetId,revision:nextRev,draftRevision:old.draftRevision||old.revision,publishedRevision:0,hasDraft:true,published:false,draft:old.draft||old.publishedSnapshot,publishedSnapshot:null,updatedAt:now,updatedByPersonId:actorPersonId};
    deletes.push({path:'publicTimeline/'+q.targetId,kind:'delete',data:null});
    resultPublished=false;resultHasDraft=true;snapshotForHistory={unpublishedRevision:old.publishedRevision};
   }
   sourceWrites.push({path:timelineItemPath(q.targetId),kind:'set',data:nextDoc});
   sourceWrites.push(...appendHistory(get,'timeline',{operationId:r.operationId,action:q.action,targetId:q.targetId,revisionBefore:oldRev,revisionAfter:nextRev,actorPersonId,actorUid:o.identity.uid,at:now,summary:q.summary,snapshotHash:hash(snapshotForHistory),snapshot:snapshotForHistory}));
  }

  const receipt={schemaVersion:'editorial-receipt-v1',projectId:o.config.projectId,operationId:r.operationId,ownerUid:o.identity.uid,actorPersonId,executor:c.executor,epoch:c.epoch,commandType:TYPE,action:q.action,resourcePath:q.resourcePath,targetId:q.targetId,digest:r.digest,revisionBefore:q.beforeRevision,revisionAfter:q.afterRevision,published:resultPublished,hasDraft:resultHasDraft,publicUpdate:q.publicUpdate,state:'COMPLETED',at:now};
  const technicalWrites=[
   {path:bucket,kind:'set',data:{windowStart:start,count:count+1}},
   {path:'commandExecutorOwners/'+r.operationId,kind:'set',data:{schemaVersion:'institutional-command-owner-v1',routeId:'PUBLIC_CONTENT_MANAGE',operationId:r.operationId,commandType:TYPE,activityId:q.resourcePath,ownerUid:o.identity.uid,digest:r.digest,executor:c.executor,epoch:c.epoch,createdAt:stamp}},
   {path:'editorialReceipts/'+r.operationId,kind:'set',data:receipt},
   {path:'editorialAudit/'+r.operationId,kind:'set',data:{...receipt,summary:q.summary,policyRefs:auth.policies.map(p=>({id:p.id,revision:p.revision}))}}
  ];
  const result={state:'COMPLETED',operationId:r.operationId,action:q.action,targetId:q.targetId,revision:q.afterRevision,published:resultPublished,hasDraft:resultHasDraft,publicUpdate:q.publicUpdate};
  const plan={kind:'PUBLIC_EDITORIAL',aggregateId:q.targetId,sourceWrites,headWrites:[],claimWrites:[],projectionWrites,technicalWrites,deletes};
  return T.applyMutationPlan(tx,T.finalize(tx,plan,result,now,Math.min(auth.validUntil,r.validUntil)),o.clock());
 }

 function readView(options,params){
  N(params&&typeof params==='object'&&!Array.isArray(params),'COMMAND_PAYLOAD_INVALID');
  const version=Object.hasOwn(params,'operationId');
  C.exact(params,version?['operationId','historyTarget']:[]);
  if(version)N(C.op(params.operationId)&&['home','about','timeline'].includes(params.historyTarget),'COMMAND_PAYLOAD_INVALID');
  return F.connect(options.config).transaction(tx=>{
   const get=p=>tx.get(p).data(),now=(options.clock||Date.now)();
   if(version){
    const target=params.historyTarget,resource=target==='timeline'?'timeline':'pages';
    control(get,options.config,options.identity,now,true,'canonicalEditorial/'+resource,true);
    const saved=get(historyPath('version_'+params.operationId));
    const event=saved?.historyTarget===target?saved.event:get(historyPath(target))?.events.find(e=>e.operationId===params.operationId);
    N(event?.operationId===params.operationId&&Object.hasOwn(event,'snapshot'),'RESOURCE_NOT_FOUND');
    return {ok:true,data:{snapshot:event.snapshot}};
   }
   let authPages=null,authTimeline=null,c=null;
   try{const r=control(get,options.config,options.identity,now,true,'canonicalEditorial/pages',true);authPages=r.auth;c=r.c;}catch(e){if(!['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
   try{const r=control(get,options.config,options.identity,now,true,'canonicalEditorial/timeline',true);authTimeline=r.auth;c=c||r.c;}catch(e){if(!['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
   N(authPages||authTimeline,'AUTHORIZATION_DENIED');
   const canManagePages=!!authPages,canManageTimeline=!!authTimeline;
   const validUntil=Math.min(f4ServiceAvailability_().bound(c,now+60000),options.identity.expiresAt,authPages?.validUntil??Infinity,authTimeline?.validUntil??Infinity);
   const pages={};
   if(canManagePages){
    for(const id of PAGE_IDS)pages[id]=validatePageCanonical(get(pagePath(id)),id);
   }
   let timeline={revision:0,items:[]};
   if(canManageTimeline){
    const idx=validateTimelineIndex(get(timelineIndexPath));
    if(idx.itemIds.length&&typeof tx.getAll==='function')tx.getAll(idx.itemIds.map(timelineItemPath));
    const items=idx.itemIds.map(id=>validateTimelineItemCanonical(get(timelineItemPath(id)),id)).filter(Boolean);
    timeline={revision:idx.revision,items};
   }
   const history={home:[],about:[],timeline:[]};
   if(canManagePages){
    for(const id of PAGE_IDS){const h=get(historyPath(id));if(h){N(h.schemaVersion==='canonical-editorial-history-v1'&&Array.isArray(h.events),'SCHEMA_INVALID');history[id]=[...h.events].reverse();}}
   }
   if(canManageTimeline){
    const h=get(historyPath('timeline'));if(h){N(h.schemaVersion==='canonical-editorial-history-v1'&&Array.isArray(h.events),'SCHEMA_INVALID');history.timeline=[...h.events].reverse();}
   }
   return {ok:true,data:{schemaVersion:'editorial-admin-view-v1',controlRevision:c.revision,epoch:c.epoch,canManagePages,canManageTimeline,pages,timeline,history,validUntil}};
  });
 }

 return {control,preview,execute,status,digest,readView,CONTROL,TYPE,CAP,PAGE_IDS,ACTIONS};
}

function f4EditorialRead_(options,params){
 return f4EditorialCommands_().readView(options,params);
}
