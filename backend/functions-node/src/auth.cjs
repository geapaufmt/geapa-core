'use strict';
const {fence,fail}=require('./environment.cjs');
async function authenticate(auth,token,config=fence(),now=Date.now()){
 if(typeof token!=='string'||token.length>16000)fail('AUTH_BINDING_INVALID');
 if(!config.emulator){let h;try{h=JSON.parse(Buffer.from(token.split('.')[0],'base64url').toString());}catch{fail('AUTH_BINDING_INVALID');}if(h.alg!=='RS256'||typeof h.kid!=='string'||!h.kid)fail('AUTH_BINDING_INVALID');}
 let v;try{v=await auth.verifyIdToken(token,true);}catch{fail('AUTH_BINDING_INVALID');}
 if(typeof v.uid!=='string'||!v.uid||v.uid.length>128||v.sub!==v.uid||v.aud!==config.projectId||v.iss!=='https://securetoken.google.com/'+config.projectId||!Number.isInteger(v.exp)||v.exp*1000<=now||!Number.isInteger(v.iat)||v.iat*1000>now+30000||!Number.isInteger(v.auth_time)||v.auth_time*1000>now+30000||v.firebase?.tenant)fail('AUTH_BINDING_INVALID');
 // checkRevoked=true also checks disabled users in the official Admin SDK.
 return {uid:v.uid,expiresAt:v.exp*1000};
}
async function appCheckGate(service,token,control,origin){
 const spec=control.appCheck;
 if(!spec||!['OBSERVE','REQUIRED'].includes(spec.mode)||!Array.isArray(spec.appIds))fail('PILOT_DISABLED');
 if(origin&&(control.browserEnabled!==true||spec.mode!=='REQUIRED'))fail('APP_CHECK_REQUIRED');
 if(!token){if(spec.mode==='REQUIRED'||origin&&control.browserEnabled!==true)fail('APP_CHECK_REQUIRED');return 'MISSING_OBSERVED';}
 let result;try{result=await service.verifyToken(token);}catch{fail('APP_CHECK_INVALID');}
 if(!spec.appIds.includes(result.appId))fail('APP_CHECK_INVALID');return 'VERIFIED';
}
module.exports={authenticate,appCheckGate};
