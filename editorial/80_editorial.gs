// F4A: native synchronous port of F3.5 editorial/index.cjs. No external module loading.
function f4Editorial_() {
  if (f4Editorial_.value_) return f4Editorial_.value_;
  return (f4Editorial_.value_ = (function () {
'use strict';
const C=f4Core_(),P=f4Projections_();const {Timestamp}=f4Firestore_();
function decide(get,r,auth,now){C.need(auth.p.technicalAccount===false&&auth.account.human===true,'CAPABILITY_REQUIRED');const path='canonicalEditorial/'+r.payload.pageId,old=get(path);C.need(old&&old.draft===true&&old.page.id===r.payload.pageId,'DOMAIN_STATE_INVALID');C.expectedRevision(r.expectedRevision,old.revision);const page=old.page;C.need(page.revision===old.revision,'DOMAIN_STATE_INVALID');const checked=P.publicPages({sourceRevision:old.revision,data:page}).dto;const candidate={id:checked.id,schemaVersion:checked.schemaVersion,published:true,sanitized:true,title:checked.title,subtitle:checked.subtitle,sections:checked.sections,sourceHash:checked.sourceHash,sourceUpdatedAt:checked.sourceUpdatedAt,updatedAt:Timestamp.fromMillis(now),revision:old.revision+1};const projection=P.publicPages({sourceRevision:candidate.revision,data:candidate});return {path,data:{draft:false,revision:candidate.revision,page:projection.dto},before:old.revision,after:candidate.revision,projectionPath:'publicPages/'+candidate.id,projection}}
return {decide};



  })());
}
