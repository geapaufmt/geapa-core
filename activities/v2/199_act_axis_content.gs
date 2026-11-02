// Orientative content extends existing axis references; it never authorizes a theme.
function f4AxisContent_(){
 const C=f4Core_(),N=C.need;
 function validate(v){
  N(v?.schemaVersion==='axis-content-v1'&&v.nature==='ORIENTATIVE'&&v.revision>0&&/^[a-f0-9]{64}$/.test(v.contentHash)&&Array.isArray(v.items)&&v.items.length===8,'SCHEMA_INVALID');
  v.items.forEach((x,i)=>{N(x.code==='EIXO_0'+(i+1)&&x.order===i+1&&x.active===true&&x.reference?.path==='presentationPolicyReferences/'+x.code&&x.reference.revision>0,'SCHEMA_INVALID');for(const k of ['name','shortName','formLabel','summary','roman','documentedUpdatedAt'])N(typeof x[k]==='string'&&x[k].length>0&&x[k].length<=1000,'SCHEMA_INVALID');N(typeof x.guide?.scope==='string'&&x.guide.scope.length>0&&x.guide.scope.length<=8000,'SCHEMA_INVALID');for(const y of [x,x.guide])for(const k of ['examples','keywords','disciplines'])N(Array.isArray(y[k])&&y[k].length>0&&y[k].length<=40&&y[k].every(s=>typeof s==='string'&&s.length>0&&s.length<=1000),'SCHEMA_INVALID');});return v;
 }
 // Explicit allowlist: notes, policy contents, operation IDs and source ACLs do not leave this boundary.
 function publicItem(x){return {code:x.code,order:x.order,roman:x.roman,name:x.name,shortName:x.shortName,formLabel:x.formLabel,summary:x.summary,active:x.active,examples:[...x.examples],keywords:[...x.keywords],disciplines:[...x.disciplines],documentedUpdatedAt:x.documentedUpdatedAt,guide:{scope:x.guide.scope,examples:[...x.guide.examples],keywords:[...x.guide.keywords],disciplines:[...x.guide.disciplines]}};}
 function publish(v,approval,at){
  validate(v);C.exact(approval,['contentHash','revision']);N(approval.contentHash===v.contentHash&&approval.revision===v.revision,'REVISION_CONFLICT');N(Number.isFinite(Date.parse(at)),'SCHEMA_INVALID');
  const section=(id,title,text,order,subtitle='')=>({id,type:'TEXT',title,subtitle,text,imageUrl:'',buttonText:'',buttonUrl:'',order,published:true});
  const page=(id,title,subtitle,sections)=>f4Projections_().publicPages({sourceRevision:v.revision,data:{id,schemaVersion:'geapa-public-v2.1',published:true,sanitized:true,title,subtitle,sections,sourceHash:v.contentHash,sourceUpdatedAt:v.documentedUpdatedAt,updatedAt:at,revision:v.revision}}).dto;
  const pages=v.items.map(x=>{const y=publicItem(x),sections=[section('identity','Identificação',[y.code,y.roman,y.shortName,y.formLabel,y.active?'Ativo':'Inativo',y.documentedUpdatedAt].join('\n'),0),section('scope','O que abrange',y.guide.scope,1),section('guide-examples','Exemplos do guia',y.guide.examples.join('\n'),2),section('examples','Exemplos da planilha',y.examples.join('\n'),3),section('keywords','Palavras-chave da planilha',y.keywords.join('\n'),4),section('guide-keywords','Palavras-chave do guia',y.guide.keywords.join('\n'),5),section('disciplines','Disciplinas da planilha',y.disciplines.join('\n'),6),section('guide-disciplines','Disciplinas do guia',y.guide.disciplines.join('\n'),7),section('provenance','Proveniência','Guia de Eixos Temáticos GEAPA (2025); planilha EIXOS_TEMATICOS_OFICIAIS. Conteúdo orientativo, sem valor normativo. '+v.guideReference,8)];return page(v.id+'_'+x.code,y.name,y.summary,sections);});
  const index=page('eixos-tematicos','Eixos temáticos','Guia orientativo, sem valor normativo.',v.items.map((x,i)=>section(pages[i].id,x.roman+' — '+x.name,x.summary,x.order,x.code)));
  return {index,pages};
 }
 return {validate,publicItem,publish};
}
function f4AxisContentRead_(options){
 const N=f4Core_().need,now=(options.clock||Date.now)();
 return options.ports.readWithAuthoritySnapshot((check,get)=>{
  const proof=options.ports.readAuthority(['activityAttendanceRegistry/current'],false,'ACTIVITY_BOARD_READ');
  const pointer=get('canonicalEditorial/EIXOS_CURRENT');N(pointer?.schemaVersion==='axis-publication-v1'&&/^EIXOS_[0-9]{8}$/.test(pointer.contentId),'PROJECTION_NOT_READY');
  const v=f4AxisContent_().validate(get('canonicalEditorial/'+pointer.contentId));N(v.contentHash===pointer.contentHash&&v.revision===pointer.revision,'PROJECTION_NOT_READY');
  const items=v.items.map(x=>{const ref=get(x.reference.path);N(ref?.id===x.code&&ref.usage==='THEMATIC_AXIS'&&ref.approved&&ref.active&&ref.revision===x.reference.revision&&ref.name===x.name&&ref.description===x.summary,'PROJECTION_NOT_READY');return {...f4AxisContent_().publicItem(x),administrativeNotes:x.administrativeNotes,reference:{...x.reference}};});
  check([proof]);return {ok:true,data:{schemaVersion:'axis-directory-v1',revision:v.revision,contentHash:v.contentHash,nature:v.nature,items,sources:v.sources.map(s=>({title:s.title,modifiedTime:s.modifiedTime,textSha256:s.textSha256})),validUntil:Math.min(now+60000,options.identity.expiresAt)}};
 });
}
