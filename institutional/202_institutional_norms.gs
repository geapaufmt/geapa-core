// Derived documentary index. It never changes normative authority or public publication.
function f4InstitutionalNorms_(){
 const C=f4Core_(),N=f4MemberAuthzContracts_().need,M=f4InstitutionalDirectoryModel_();
 const field=(r,k)=>r.fields.find(f=>f.label===k)?.value||'';
 function index(sections,get,now){
  const records=Object.entries(sections).flatMap(([section,v])=>{M.validate(v);return v.records.filter(r=>r.visibility==='BOARD').map(r=>({...r,section}));});
  const docs=records.filter(r=>r.section==='norms'&&r.kind==='DOCUMENTOS_NORMATIVOS'),byId=new Map(docs.map(r=>[r.id,r])),all=new Map(records.map(r=>[r.id,r]));
  const old=byId.get('REG-2025'),current=byId.get('REG-2026');
  // The two documentary rows explicitly identify the same Regimento and succession.
  const reg=!!old&&!!current&&field(old,'TIPO_DOCUMENTO')==='REGIMENTO'&&field(current,'TIPO_DOCUMENTO')==='REGIMENTO'&&field(old,'TITULO_COMPLETO')===field(current,'TITULO_COMPLETO')&&field(old,'TITULO_COMPLETO').length>10&&old.status==='SUBSTITUIDO'&&current.canonical?.path==='institutionalNormVersions/REGIMENTO_2026_OFICIAL_20260514';
  const versions=docs.map(r=>{const v=r.canonical?.path?.startsWith('institutionalNormVersions/')?get(r.canonical.path):null,a=v?.approvalRef?get('institutionalNormApprovals/'+v.approvalRef):null;
   const isCurrent=!!v&&v.status==='CURRENT'&&a?.active===true&&a.normId===v.id&&a.version===v.version&&a.immutableVersionRef===v.immutableVersionRef&&/^\d{4}-\d{2}-\d{2}$/.test(v.effectiveFrom)&&Date.parse(v.effectiveFrom+'T00:00:00Z')<=now;
   // Only already public, canonical source references are exposed as public links.
   const files=v?.provenance?.readAccess==='PUBLIC'&&/^https:\/\/drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\//.test(v.pdfUrl||'')?[{url:v.pdfUrl,label:'Arquivo canônico',reference:v.versionId||r.id}]:[];
   return {id:r.id,instrumentId:reg&&['REG-2025','REG-2026'].includes(r.id)?'REGIMENTO_GEAPA':'INSTR_'+r.id,title:r.title,description:field(r,'ASSUNTO_MACRO'),version:field(r,'VERSAO'),documentaryStatus:r.status,current:isCurrent,effectiveFrom:isCurrent?v.effectiveFrom:null,files,recordId:r.id};});
  const versionMap=new Map(versions.map(v=>[v.id,v])),relations=[];
  function targets(r){
   const direct=['DOC_ID','DOC_ALVO','DOCUMENTO_RELACIONADO_ID'].map(k=>({key:k,value:field(r,k)})).filter(x=>x.value);
   const base=field(r,'BASE_LEGAL'),device=field(r,'DISPOSITIVO_ID');
   if(base)direct.push({key:'BASE_LEGAL',value:base});if(device)direct.push({key:'DISPOSITIVO_ID',value:device});
   return direct.flatMap(x=>{if(byId.has(x.value))return [{versionId:x.value,evidence:x.key+' = '+x.value}];const linked=all.get(x.value),doc=linked&&field(linked,'DOC_ID');return linked?.kind==='DISPOSITIVOS_NORMATIVOS'&&byId.has(doc)?[{versionId:doc,evidence:x.key+' = '+x.value+' → DOC_ID = '+doc}]:[];});
  }
  for(const r of records){let links=r.kind==='DOCUMENTOS_NORMATIVOS'?[{versionId:r.id,evidence:'Identidade documental DOC_ID'}]:targets(r);links=links.filter((x,i,a)=>a.findIndex(y=>y.versionId===x.versionId)===i);
   const explicit=field(r,'DOC_ID')||field(r,'DOC_ALVO');if(explicit&&!byId.has(explicit))links=[];
   const group=r.section==='parameters'?'parameters':r.section==='governance'?'governance':r.kind==='DOCUMENTOS_NORMATIVOS'?'versions':r.kind==='DISPOSITIVOS_NORMATIVOS'?'provisions':r.kind==='PROPOSTAS_E_EMENDAS'?'proposals':field(r,'CATEGORIA')==='FAQ'?'faq':'explanations';
   relations.push({section:r.section,recordId:r.id,group,links:links.map(x=>({...x,instrumentId:versionMap.get(x.versionId).instrumentId,historical:versionMap.get(x.versionId).documentaryStatus==='SUBSTITUIDO',current:versionMap.get(x.versionId).current})),reviewRequired:!links.length||group==='faq'||r.issues.length>0});
  }
  const instruments=[...new Set(versions.map(v=>v.instrumentId))].map(id=>{const rows=versions.filter(v=>v.instrumentId===id);return {id,title:id==='REGIMENTO_GEAPA'?'Regimento Interno do GEAPA':rows[0].title,description:rows[0].description,versionIds:rows.map(v=>v.id),currentVersionIds:rows.filter(v=>v.current).map(v=>v.id)};});
  return {schemaVersion:'institutional-norm-index-v1',sourceHash:C.hash(Object.entries(sections).map(([k,v])=>({section:k,revision:v.revision,hash:C.hash(v)}))),instruments,versions,relations};
 }
 function eligible(r){return r.visibility==='BOARD'&&!r.canonical&&((r.kind==='PROPOSTAS_E_EMENDAS'&&['RASCUNHO','EM_PROPOSTA'].includes(r.status))||r.kind==='PORTAL_GEAPA'||(r.kind==='Ata / deliberação'&&['RASCUNHO','EM_REDACAO'].includes(r.status)&&!field(r,'DOCUMENTO_RELACIONADO_ID')&&!['APROVADA','APROVADO'].includes(field(r,'RESULTADO'))));}
 function overlay(v,section){if(!v)return {schemaVersion:'institutional-presentation-v1',section,revision:0,entries:[]};C.exact(v,['schemaVersion','section','revision','entries']);N(v.schemaVersion==='institutional-presentation-v1'&&v.section===section&&['norms','governance'].includes(section)&&C.rev(v.revision)&&Array.isArray(v.entries)&&v.entries.length<=150,'SCHEMA_INVALID');N(new Set(v.entries.map(e=>e.recordId)).size===v.entries.length,'SCHEMA_INVALID');for(const e of v.entries){C.exact(e,['recordId','hidden','reason','actorPersonId','at','operationId']);N(C.id(e.recordId)&&typeof e.hidden==='boolean'&&typeof e.reason==='string'&&e.reason.length>=10&&e.reason.length<=1000&&C.id(e.actorPersonId)&&C.op(e.operationId)&&Number.isSafeInteger(e.at),'SCHEMA_INVALID');}return v;}
 return {index,field,eligible,overlay};
}

function f4InstitutionalNormRead_(options,params){
 const C=f4Core_(),N=f4MemberAuthzContracts_().need,M=f4InstitutionalDirectoryModel_(),I=f4InstitutionalNorms_();
 C.exact(params,['section','view','instrument','query','visibility','page']);N(['norms','governance'].includes(params.section)&&params.view==='instruments'&&typeof params.instrument==='string'&&params.instrument.length<=80&&typeof params.query==='string'&&params.query.length<=80&&['VISIBLE','HIDDEN','ALL'].includes(params.visibility)&&Number.isInteger(params.page)&&params.page>=1&&params.page<=50,'COMMAND_PAYLOAD_INVALID');
 return f4Firestore_().connect(options.config).transaction(tx=>{
  const get=p=>tx.get(p).data(),now=(options.clock||Date.now)(),control=get('institutionalDirectoryControl/current');N(control?.enabled===true&&control.schemaVersion==='institutional-directory-control-v1','FEATURE_DISABLED');
  const resolve=s=>f4InstitutionalPolicies_(get,options.identity,now,control.policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/'+s),auth=resolve(params.section);N(auth.mask.includes('summary')&&auth.mask.includes('fields'),'AUTHORIZATION_DENIED');
  const sections={},masks={},auths=[auth],unavailable=[];
  for(const s of params.section==='norms'?['norms','parameters','governance']:['governance','norms']){let a;try{a=s===params.section?auth:resolve(s);N(a.mask.includes('summary')&&a.mask.includes('fields'),'AUTHORIZATION_DENIED');}catch(e){if(e.code==='AUTHORIZATION_DENIED'){unavailable.push(s);continue;}throw e;}const v=get('institutionalDirectory/'+s);if(!v){if(s===params.section)N(false,'PROJECTION_NOT_READY');unavailable.push(s);continue;}M.validate(v);sections[s]=v;masks[s]=a.mask;auths.push(a);}
  const idx=I.index(sections,get,now),visibility={};for(const s of Object.keys(sections))if(['norms','governance'].includes(s))visibility[s]=I.overlay(get('institutionalPresentation/'+s),s);
  const histories={};for(const s of Object.keys(visibility)){const h=get('institutionalPresentationHistory/'+s);N(!h||(h.schemaVersion==='institutional-presentation-history-v1'&&Array.isArray(h.events)&&h.events.length<=500),'SCHEMA_INVALID');histories[s]=h?.events||[];}
  const actions={};for(const s of Object.keys(visibility)){actions[s]=[];try{const a=f4InstitutionalVisibilityCommands_().control(get,options.config,options.identity,now,true,s,true);actions[s]=sections[s].records.filter(r=>I.eligible(r)&&a.c.recordIds[s]?.includes(r.id)&&a.auth.mask.includes('presentation')).map(r=>r.id);}catch(e){if(!['AUTHORIZATION_DENIED','ACCESS_REVOKED','PILOT_DISABLED','FEATURE_DISABLED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}}
  const normalize=s=>s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  let rows=Object.entries(sections).flatMap(([s,v])=>M.dto(v,masks[s]).map(r=>{const rel=idx.relations.find(x=>x.section===s&&x.recordId===r.id),presentation=I.eligible(v.records.find(x=>x.id===r.id))?(visibility[s]?.entries.find(e=>e.recordId===r.id)||null):null;return {...r,...(s==='parameters'&&masks[s].includes('canonical')?{parameterEvidence:f4InstitutionalParameterEvidence_(r,get,rel.links)}:{}),section:s,relation:rel,presentation,presentationHistory:(histories[s]||[]).filter(e=>e.recordId===r.id),canChangePresentation:actions[s]?.includes(r.id)||false,sourceRevision:v.revision,presentationRevision:visibility[s]?.revision||0};}));
  if(params.section==='governance')rows=rows.filter(r=>r.section==='governance');
  else if(params.instrument==='UNRESOLVED')rows=rows.filter(r=>!r.relation.links.length);else if(params.instrument)rows=rows.filter(r=>r.relation.links.some(l=>l.instrumentId===params.instrument));
  const hiddenCount=rows.filter(r=>r.presentation?.hidden===true).length;
  rows=rows.filter(r=>params.visibility==='ALL'||(params.visibility==='HIDDEN')===(r.presentation?.hidden===true)).filter(r=>!params.query||normalize(r.id+' '+r.title+' '+r.fields.map(f=>f.value).join(' ')).includes(normalize(params.query)));
  const order=['versions','provisions','proposals','faq','explanations','parameters','governance'];rows.sort((a,b)=>order.indexOf(a.relation.group)-order.indexOf(b.relation.group)||(a.relation.group==='versions'?Number(b.relation.links.some(l=>l.current))-Number(a.relation.links.some(l=>l.current)):0)||a.id.localeCompare(b.id));
  // Do not expose relationships/contents from sections without a read policy or fields mask.
  const safeIndex={...idx,relations:undefined};delete safeIndex.relations;
  return {ok:true,data:{schemaVersion:'institutional-instruments-view-v1',section:params.section,index:safeIndex,instrument:params.instrument,items:params.section==='norms'&&!params.instrument?[]:rows.slice((params.page-1)*40,params.page*40),total:rows.length,page:params.page,pages:Math.max(1,Math.ceil(rows.length/40)),hiddenCount,unavailableSections:unavailable,revision:sections[params.section].revision,validUntil:Math.min(...auths.map(a=>a.validUntil)),publicationEnabled:false}};
 });
}

