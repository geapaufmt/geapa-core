'use strict';
// Request-isolated shared source. No B0 Worker, eval of client code, or external I/O.
const fs=require('node:fs'),p=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto');
const script=new vm.Script(fs.readFileSync(p.join(__dirname,'../generated/domain.gs'),'utf8'),{filename:'current-core.gs'});
const facade=fs.readFileSync(p.join(__dirname,'domain-port.js'),'utf8');
function createDomain(io){
 const deny=()=>{throw Object.assign(Error('EXTERNAL_SERVICE_DISABLED'),{code:'EXTERNAL_SERVICE_DISABLED'});};
 const utilities={DigestAlgorithm:{SHA_256:'sha256',SHA_512:'sha512',MD5:'md5'},Charset:{UTF_8:'utf8'},
 computeDigest:(alg,v)=>[...crypto.createHash(alg).update(typeof v==='string'?v:Buffer.from(v.map(x=>x&255))).digest()].map(x=>x>127?x-256:x),
 newBlob:v=>({getBytes:()=>[...Buffer.from(v)],getDataAsString:()=>Buffer.from(v).toString('utf8')}),
 base64Encode:v=>Buffer.from(v.map(x=>x&255)).toString('base64'),base64Decode:v=>[...Buffer.from(v,'base64')],
 formatDate:(date,zone,pattern)=>{const q=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:zone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(date).map(x=>[x.type,x.value]));const values={'yyyy-MM-dd':`${q.year}-${q.month}-${q.day}`,'HH:mm':`${q.hour}:${q.minute}`,"yyyy-MM-dd'T'HH:mm:ss":`${q.year}-${q.month}-${q.day}T${q.hour}:${q.minute}:${q.second}`,'dd/MM/yyyy HH:mm':`${q.day}/${q.month}/${q.year} ${q.hour}:${q.minute}`};if(!values[pattern])deny();return values[pattern];}};
 const ctx=vm.createContext({Utilities:utilities,__io:io,__profileJson:JSON.stringify(require('./runtime-profile.cjs').profile),console:{log:()=>{},warn:()=>{},error:()=>{}}},{codeGeneration:{strings:false,wasm:false}});
 for(const k of ['UrlFetchApp','ScriptApp','PropertiesService','Session','SpreadsheetApp','DriveApp','GmailApp','MailApp','LockService','CacheService'])ctx[k]=new Proxy({}, {get:()=>deny});
 script.runInContext(ctx,{timeout:10000});vm.runInContext(facade,ctx,{timeout:10000});
 return {run(code,input){ctx.__input=JSON.stringify(input??null);return vm.runInContext(`JSON.stringify((()=>{const input=JSON.parse(__input);return (${code})(input);})())`,ctx,{timeout:10000});},context:ctx};
}
module.exports={createDomain};
