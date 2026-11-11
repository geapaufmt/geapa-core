'use strict';
// Temporary DEV admission only; every request still resolves canonical policies.
// Never admits projection recovery, data mutation or another module's reader.
function admitted(control,uid,route,now,production){
 if(production||!['/act/context','/admin/gestao-institucional'].includes(route))return false;
 const rows=control?.specializedReaders?.temporaryInstitutionalActors;
 if(!Array.isArray(rows)||rows.length>4)return false;
 const matches=rows.filter(x=>x?.uid===uid);
 if(matches.length!==1)return false;
 const x=matches[0];return Object.keys(x).sort().join(',')==='expiresAt,startsAt,uid'&&Number.isSafeInteger(x.startsAt)&&Number.isSafeInteger(x.expiresAt)&&x.startsAt<=now&&now<x.expiresAt&&x.expiresAt-x.startsAt<=4*3600000;
}
module.exports={admitted};
