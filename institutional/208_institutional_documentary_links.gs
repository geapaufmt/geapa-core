// Documentary evidence only. A link never votes, approves or enacts a norm.
function f4InstitutionalDocumentaryLinks_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,hash=v=>C.hash(F.encode(v));
 const modalities=['MEETING_MINUTES','ELECTRONIC_DELIBERATION_TERM'];
 const purposes=['APPROVAL','AMENDMENT','REVOCATION'];
 const text=(v,max,min=1)=>typeof v==='string'&&v===v.trim()&&v.length>=min&&v.length<=max&&!/[\x00-\x1f]/.test(v);
 function normalize(get,input,old){
  if(!Object.hasOwn(input,'documentaryEvidence'))return old?.documentaryEvidence||null;
  const d=input.documentaryEvidence;
  if(d===null){N(!old?.documentaryEvidence,'DOCUMENTARY_EVIDENCE_REQUIRED');return null;}
  C.exact(d,['modality','documentDate','documentDatePrecision','documentAuthors','sourceDescription','sourceSha256','relations']);
  N(modalities.includes(d.modality)&&(!input.minutes||d.modality==='MEETING_MINUTES'),'DOCUMENTARY_MODALITY_INVALID');
  N(['DAY','MONTH','YEAR','UNKNOWN'].includes(d.documentDatePrecision),'COMMAND_PAYLOAD_INVALID');
  const pattern={DAY:/^\d{4}-\d{2}-\d{2}$/,MONTH:/^\d{4}-(0[1-9]|1[0-2])$/,YEAR:/^\d{4}$/};
  N(d.documentDatePrecision==='UNKNOWN'?d.documentDate===null:typeof d.documentDate==='string'&&pattern[d.documentDatePrecision].test(d.documentDate),'DOCUMENTARY_DATE_INVALID');
  if(d.documentDatePrecision==='DAY')N(Number.isFinite(Date.parse(d.documentDate+'T12:00:00Z'))&&new Date(d.documentDate+'T12:00:00Z').toISOString().slice(0,10)===d.documentDate,'DOCUMENTARY_DATE_INVALID');
  N(Array.isArray(d.documentAuthors)&&d.documentAuthors.length>0&&d.documentAuthors.length<=20&&d.documentAuthors.every(a=>text(a,160))&&new Set(d.documentAuthors).size===d.documentAuthors.length,'COMMAND_PAYLOAD_INVALID');
  N(text(d.sourceDescription,1500,10)&&/^[a-f0-9]{64}$/.test(d.sourceSha256)&&Array.isArray(d.relations)&&d.relations.length<=10,'COMMAND_PAYLOAD_INVALID');
  N(new Set(d.relations.map(r=>r.normVersionId+'|'+r.relation)).size===d.relations.length,'DOCUMENTARY_RELATION_DUPLICATE');
  const relations=d.relations.map(r=>{
   C.exact(r,['normVersionId','expectedNormRevision','expectedNormHash','relation','evidence']);
   N(C.id(r.normVersionId)&&C.rev(r.expectedNormRevision)&&/^[a-f0-9]{64}$/.test(r.expectedNormHash)&&purposes.includes(r.relation)&&text(r.evidence,1500,10),'COMMAND_PAYLOAD_INVALID');
   const norm=get('institutionalNormVersions/'+r.normVersionId);
   N(norm&&norm.id===r.normVersionId&&norm.revision===r.expectedNormRevision&&hash(norm)===r.expectedNormHash,'NORMATIVE_VERSION_CONFLICT');
   // Hash and revision pin the precise normative text, not its current status.
   // Author names and source declarations are evidence, not verified votes.
   return {...r,normVersion:norm.version,verification:'PENDING_VERIFICATION'};
  });
  return {...d,relations,approvalMeaning:'DOCUMENTARY_EVIDENCE_ONLY',votingImplemented:false};
 }
 function input(d){if(!d)return null;return Object.fromEntries(['modality','documentDate','documentDatePrecision','documentAuthors','sourceDescription','sourceSha256'].map(k=>[k,d[k]]).concat([['relations',d.relations.map(r=>Object.fromEntries(['normVersionId','expectedNormRevision','expectedNormHash','relation','evidence'].map(k=>[k,r[k]])))]]));}
 function page(records,q){
  if(!q)return {records};
  C.exact(q,['query','meetingType','boardId','state','from','to','normVersionId','page']);
  for(const k of ['query','meetingType','boardId','state','normVersionId'])N(typeof q[k]==='string'&&q[k].length<=100,'COMMAND_PAYLOAD_INVALID');
  for(const k of ['from','to'])N(q[k]===''||/^\d{4}-\d{2}-\d{2}$/.test(q[k])&&Number.isFinite(Date.parse(q[k]+'T12:00:00Z')),'COMMAND_PAYLOAD_INVALID');
  N((!q.from||!q.to||q.from<=q.to)&&Number.isInteger(q.page)&&q.page>=1&&q.page<=100,'COMMAND_PAYLOAD_INVALID');
  const norm=s=>s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const rows=records.filter(r=>(!q.query||norm([r.title,r.body,r.minutes?.number,r.minutes?.agenda].join(' ')).includes(norm(q.query)))&&(!q.meetingType||norm(r.minutes?.meetingType||'').includes(norm(q.meetingType)))&&(!q.boardId||r.minutes?.boardId===q.boardId)&&(!q.state||(r.minutes?.state||r.documentaryState)===q.state)&&(!q.from||r.meetingDate&&r.meetingDate>=q.from)&&(!q.to||r.meetingDate&&r.meetingDate<=q.to)&&(!q.normVersionId||r.documentaryEvidence?.relations.some(l=>l.normVersionId===q.normVersionId))).sort((a,b)=>(b.meetingDate||'').localeCompare(a.meetingDate||'')||a.id.localeCompare(b.id));
  return {records:rows.slice((q.page-1)*10,q.page*10),pagination:{page:q.page,pages:Math.max(1,Math.ceil(rows.length/10)),total:rows.length,pageSize:10},filters:q};
 }
 return {normalize,input,hash,modalities,purposes,page};
}
