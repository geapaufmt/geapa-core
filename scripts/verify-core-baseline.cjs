'use strict';
const fs=require('node:fs'),p=require('node:path'),crypto=require('node:crypto'),vm=require('node:vm'),cp=require('node:child_process'),a=require('node:assert/strict');
const root=p.resolve(__dirname,'..'),hash=v=>crypto.createHash('sha256').update(v).digest('hex');
function verify(ref){
 const read=file=>ref?cp.execFileSync('git',['show',ref+':'+file],{cwd:root,maxBuffer:20*1024*1024}):fs.readFileSync(p.join(root,file));
 const manifest=JSON.parse(read('docs/baselines/core-dev-v43.json'));
 const files=manifest.files.map(f=>{const raw=read(f.path),source=raw.toString('utf8');a.equal(hash(raw),f.sha256,f.path);if(f.type==='SERVER_JS')new vm.Script(source,{filename:f.path});else JSON.parse(source);return {name:f.name,type:f.type,source};}).sort((a,b)=>a.name.localeCompare(b.name));
 a.equal(hash(JSON.stringify(files)),manifest.sourceHash);return {version:manifest.version,sourceHash:manifest.sourceHash,files:files.length,ref:ref||'working-tree',syntax:true};
}
if(require.main===module)console.log(verify(process.argv[2]==='--ref'?process.argv[3]:undefined));
module.exports={verify};
