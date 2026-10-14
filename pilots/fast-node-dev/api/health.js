'use strict';
const {randomUUID}=require('node:crypto');
const instance=randomUUID();let requests=0;
// This preview-only pilot has no Firestore client, secret, command, or mutation.
module.exports=function health(req,res){
 res.setHeader('Cache-Control','no-store');
 res.setHeader('Content-Type','application/json; charset=utf-8');
 if(process.env.VERCEL_ENV==='production'){res.statusCode=403;return res.end(JSON.stringify({error:'PREVIEW_ONLY'}));}
 if(!['GET','HEAD'].includes(req.method)){res.statusCode=405;res.setHeader('Allow','GET, HEAD');return res.end(JSON.stringify({error:'METHOD_NOT_ALLOWED'}));}
 const firstInvocation=requests++===0;
 const body={ok:true,environment:'DEV_PILOT',privilegedAccess:false,commandsEnabled:false,region:process.env.VERCEL_REGION||'local',instance,firstInvocation};
 res.statusCode=200;return res.end(req.method==='HEAD'?'':JSON.stringify(body));
};
