'use strict';
const fs=require('node:fs'),p=require('node:path'),cp=require('node:child_process'),crypto=require('node:crypto'),a=require('node:assert/strict');
const root=p.resolve(__dirname,'..'),repo=p.resolve(root,'../..');
const ref='fac44ff26436e11d0b91424f0a7aa1081b151ff9',hash=b=>crypto.createHash('sha256').update(b).digest('hex');
function materialize(){
 const read=file=>cp.execFileSync('git',['show',ref+':'+file],{cwd:repo,maxBuffer:20*1024*1024});
 const m=JSON.parse(read('docs/baselines/core-dev-v43.json')),files=[];
 for(const f of m.files){const b=read(f.path);a.equal(hash(b),f.sha256,f.path);const dest=p.resolve(root,'base/core',f.name+(f.type==='JSON'?'.json':'.gs'));a(dest.startsWith(p.resolve(root,'base/core')+p.sep));fs.mkdirSync(p.dirname(dest),{recursive:true});fs.writeFileSync(dest,b);files.push({name:f.name,type:f.type,source:b.toString('utf8')});}
 files.sort((a,b)=>a.name.localeCompare(b.name));a.equal(hash(JSON.stringify(files)),m.sourceHash);return {ref,files:files.length,hash:m.sourceHash,remote:false};
}
if(require.main===module)console.log(materialize());module.exports={materialize,root,repo,hash};
