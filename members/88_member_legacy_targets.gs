// Protected destination registry for a local legacy package. Unknown destinations deny.
function f4MemberLegacyTargets_(){
 const C=f4Core_(),N=C.need;
 function createLocal(options){
  C.exact(options,['config','sheets','collections','urls','resources']);C.destination(options.config);N((options.config.environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(options.config)),'DESTINATION_INVALID');
  N(Array.isArray(options.sheets)&&Array.isArray(options.collections)&&Array.isArray(options.urls),'SCHEMA_INVALID');const sheets=new Map(),collections=new Map(),urls=new Map();
  for(const row of options.sheets){C.exact(row,['spreadsheetId','sheetName','domain']);N(typeof row.spreadsheetId==='string'&&row.spreadsheetId&&typeof row.sheetName==='string'&&row.sheetName&&['MEMBERS','OTHER_DOMAIN','SHARED_MAIL'].includes(row.domain),'SCHEMA_INVALID');const key=row.spreadsheetId+'/'+row.sheetName;N(!sheets.has(key),'SCHEMA_INVALID');sheets.set(key,row.domain);}
  for(const row of options.collections){C.exact(row,['name','domain']);N(C.id(row.name)&&['MEMBERS','OTHER_DOMAIN'].includes(row.domain)&&!collections.has(row.name),'SCHEMA_INVALID');collections.set(row.name,row.domain);}
  for(const row of options.urls){C.exact(row,['url','method','domain']);N(typeof row.url==='string'&&row.url.startsWith('https://')&&!row.url.includes('*')&&['get','post','patch','delete'].includes(row.method)&&['MEMBERS','OTHER_DOMAIN','READ_ONLY'].includes(row.domain),'SCHEMA_INVALID');const key=row.method+' '+row.url;N(!urls.has(key),'SCHEMA_INVALID');urls.set(key,row.domain);}
  const resources=new Map();N(Array.isArray(options.resources),'SCHEMA_INVALID');
  for(const row of options.resources){C.exact(row,['kind','id','domain']);N(['DRIVE','GMAIL','EMAIL'].includes(row.kind)&&typeof row.id==='string'&&row.id&&!row.id.includes('*')&&['MEMBERS','OTHER_DOMAIN'].includes(row.domain),'SCHEMA_INVALID');const key=row.kind+'/'+row.id;N(!resources.has(key),'SCHEMA_INVALID');resources.set(key,row.domain);}
  function classify(site,receiver,method,args){
   if(method==='sendEmail'){
    const envelope=typeof args[0]==='string'?{to:args[0],...(args[3]||{})}:args[0];
    if(!envelope||typeof envelope.to!=='string')return 'UNKNOWN';
    const recipients=[envelope.to,envelope.cc,envelope.bcc].filter(Boolean).flatMap(v=>typeof v==='string'?v.split(',').map(x=>x.trim().toLowerCase()):['']);
    const domains=recipients.map(id=>resources.get('EMAIL/'+id));if(!domains.length||domains.some(x=>!x))return 'UNKNOWN';return domains.includes('MEMBERS')?'MEMBERS':'OTHER_DOMAIN';
   }
   if(['createFolder','addFile','removeFile','moveTo','setTrashed'].includes(method)){
    const ids=[receiver,...args.filter(v=>v&&typeof v.getId==='function')].map(v=>typeof v?.getId==='function'?v.getId():null);
    const domains=ids.map(id=>id&&resources.get('DRIVE/'+id));if(domains.some(v=>!v))return 'UNKNOWN';return domains.includes('MEMBERS')?'MEMBERS':'OTHER_DOMAIN';
   }
   if(['createLabel','addLabel','removeLabel','markRead'].includes(method)){
    const id=typeof receiver.getId==='function'?receiver.getId():typeof args[0]==='string'?args[0]:null;
    return id&&resources.get('GMAIL/'+id)||'UNKNOWN';
   }
   if(method==='fetch'){
    const url=String(args[0]),opts=args[1]||{},verb=String(opts.method||'get').toLowerCase();
    if(!url.startsWith('https://firestore.googleapis.com/v1/projects/'+options.config.projectId+'/databases/(default)/documents'))return urls.get(verb+' '+url)||'UNKNOWN';
    if(verb==='get')return 'READ_ONLY';
    if([':beginTransaction',':rollback',':batchGet',':runQuery'].some(s=>url.endsWith(s)))return 'READ_ONLY';
    const documentPrefix='https://firestore.googleapis.com/v1/projects/'+options.config.projectId+'/databases/(default)/documents/';
    if(['patch','delete','post'].includes(verb)&&url.startsWith(documentPrefix)&&!url.slice(documentPrefix.length).split('?')[0].includes(':'))return collections.get(decodeURIComponent(url.slice(documentPrefix.length).split(/[/?]/)[0]))||'UNKNOWN';
    let body;try{body=typeof opts.payload==='string'?JSON.parse(opts.payload):opts.payload;}catch{return 'UNKNOWN';}
    if(!body||!Array.isArray(body.writes)||!body.writes.length)return 'UNKNOWN';
    let member=false;for(const w of body.writes){const name=w.update?.name||w.delete||w.transform?.document;if(typeof name!=='string')return 'UNKNOWN';const prefix='projects/'+options.config.projectId+'/databases/(default)/documents/';if(!name.startsWith(prefix))return 'UNKNOWN';const col=name.slice(prefix.length).split('/')[0],d=collections.get(col);if(!d)return 'UNKNOWN';if(d==='MEMBERS')member=true;}
    return member?'MEMBERS':'OTHER_DOMAIN';
   }
   let sheet=typeof receiver.getSheet==='function'?receiver.getSheet():receiver;
   if(method==='insertSheet'&&typeof receiver.getId==='function')return sheets.get(receiver.getId()+'/'+args[0])||'UNKNOWN';
   if(typeof sheet?.getName!=='function'||typeof sheet?.getParent!=='function')return 'UNKNOWN';
   const domain=sheets.get(sheet.getParent().getId()+'/'+sheet.getName());if(domain!=='SHARED_MAIL')return domain||'UNKNOWN';
   const columns=sheet.getLastColumn(),headers=sheet.getRange(1,1,1,columns).getValues()[0],col=headers.indexOf('Modulo Dono');
   if(col<0)return 'UNKNOWN';
   let rows=[];
   if(method==='appendRow')rows=[args[0]];
   else if(typeof receiver.getRow==='function'&&receiver.getRow()>1){
    rows=sheet.getRange(receiver.getRow(),1,receiver.getNumRows(),columns).getValues();
    // A bulk write must classify both old and new owners; never relabel a Members row to escape.
    if(method==='setValues'&&receiver.getColumn()===1)rows=rows.concat(args[0]);
   }else return 'UNKNOWN';
   if(!rows.length||rows.some(r=>!Array.isArray(r)||typeof r[col]!=='string'||!r[col]))return 'UNKNOWN';
   if(rows.some(r=>!['MEMBROS','DESLIGAMENTOS','ATIVIDADES','APRESENTACOES','SELETIVO'].includes(r[col])))return 'UNKNOWN';
   return rows.some(r=>['MEMBROS','DESLIGAMENTOS'].includes(r[col]))?'MEMBERS':'OTHER_DOMAIN';
  }
  return Object.freeze({classify});
 }
 return Object.freeze({createLocal});
}
