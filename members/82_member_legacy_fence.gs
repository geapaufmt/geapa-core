// G server-side guard. No ambient mode/default and no frontend flag is consulted.
function f4MemberLegacyFence_(){
 if(f4MemberLegacyFence_.instance_)return f4MemberLegacyFence_.instance_;
 let local=null,executionEpoch=null;const stack=[];
 function denied(){const e=new Error('MEMBERS_LEGACY_FENCED');e.code='MEMBERS_LEGACY_FENCED';throw e;}
 function check(){
  if(!local)denied();const c=f4MemberAuthority_().validate(local.readControl());
  if(c.mode!=='LEGACY'||c.activeWriterVersion!=='NONE'||c.epoch!==executionEpoch)denied();return c.epoch;
 }
 function installLocal(options){
  const C=f4Core_();C.exact(options,['config','readControl','classify']);C.destination(options.config);
  C.need((options.config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(options.config))&&typeof options.readControl==='function'&&typeof options.classify==='function','DESTINATION_INVALID');
  if(local)denied();executionEpoch=f4MemberAuthority_().validate(options.readControl()).epoch;local=Object.freeze({...options});
 }
 function enter(site){if(typeof site!=='string')denied();return check();}
 function begin(site){enter(site);const token=Object.freeze({});stack.push(token);return token;}
 function end(token){if(stack[stack.length-1]!==token)denied();stack.pop();}
 function sharedEffect(site){if(stack.length)enter(site);}
 function fromModule(moduleName,fn){
  if(typeof moduleName!=='string'||!moduleName||typeof fn!=='function')denied();
  // Literal legacy Mail Hub module IDs from core/18_core_mail_adapters.js, not capability aliases.
  if(['ATIVIDADES','APRESENTACOES','SELETIVO'].includes(moduleName))return fn();
  if(!['MEMBROS','DESLIGAMENTOS'].includes(moduleName))denied();
  const token=begin('SHARED_MAIL_MEMBROS');try{return fn();}finally{end(token);}
 }
 function invoke(site,receiver,method,args){
  if(typeof site!=='string'||!receiver||typeof method!=='string'||!Array.isArray(args)||typeof receiver[method]!=='function')denied();
  if(!local)denied();const classification=local.classify(site,receiver,method,args);
  if(classification==='MEMBERS'||stack.length&&classification!=='READ_ONLY')check();else if(classification!=='OTHER_DOMAIN'&&classification!=='READ_ONLY')denied();
  return Reflect.apply(receiver[method],receiver,args);
 }
 const api=Object.freeze({installLocal,enter,begin,end,sharedEffect,fromModule,invoke});f4MemberLegacyFence_.instance_=api;return api;
}
