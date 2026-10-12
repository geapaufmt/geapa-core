'use strict';
const fs=require('node:fs'),p=require('node:path'),vm=require('node:vm'),Module=require('node:module');
const root=p.resolve(__dirname,'../..'),source=p.resolve(__dirname,'..');
process.env.F4_DEV_PACKAGE=p.join(root,'base/core');
const file=p.join(source,'f4/runtime.cjs'),mod=new Module(file,module);mod.filename=file;mod.paths=Module._nodeModulePaths(p.dirname(file));
let text=fs.readFileSync(file,'utf8');
text=text.replace("sleep: () => {},","newBlob: s=>({getBytes:()=>[...Buffer.from(s)]}), sleep: () => {},");
text=text.replace('format !== "yyyy-MM-dd\'T\'HH:mm:ss"','!["yyyy-MM-dd\'T\'HH:mm:ss","yyyy-MM-dd"].includes(format)');
text=text.replace('return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}`;', 'return `${p.year}-${p.month}-${p.day}`+(format.includes("HH")?`T${p.hour}:${p.minute}:${p.second}`:"");');
// The old double checked write preconditions only. Model Firestore's entire transaction read set.
text=text.replace("if (suffix !== ':commit')", "if (body.transaction) { const snapshot=transactions.get(body.transaction); for(const [key,v] of snapshot||[]) if(documents.get(key)?.updateTime!==v.updateTime)return response(409,{error:{status:'ABORTED'}}); }\n      if (suffix !== ':commit')");
mod._compile(text,file);
function runtime(){const r=mod.exports.createRuntime();for(const f of ['fixtures.gs','b/fixtures.gs','c/fixtures.gs','d/fixtures.gs','e/fixtures.gs','e/saga-fixtures.gs'])vm.runInContext(fs.readFileSync(p.join(source,'m13',f),'utf8'),r.context,{filename:f});return r}
function seed(r,docs){for(const [key,fields]of Object.entries(docs))r.documents.set(key,{name:'projects/demo-geapa-m04-local/databases/(default)/documents/'+key,fields,updateTime:'2026-01-01T00:00:00.001Z'})}
module.exports={root,source,runtime,seed};
