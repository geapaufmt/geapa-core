'use strict';
// Local-only. Reconstruct the immutable baseline from Git, never from HEAD/WIP.
const fs=require('node:fs'),p=require('node:path'),a=require('node:assert/strict'),{materialize,root,repo,hash}=require('./materialize-baseline.cjs');
function localManifest(){const m=require(p.join(repo,'docs/baselines/core-dev-v43.json'));return Object.fromEntries(m.files.map(f=>[f.path,hash(fs.readFileSync(p.join(root,'base/core',f.name+(f.type==='JSON'?'.json':'.gs'))))]));}
function main(){const baseline=materialize();fs.mkdirSync(p.join(root,'evidence'),{recursive:true});const file=p.join(root,'evidence/local-before.json'),files=localManifest();if(fs.existsSync(file))a.deepEqual(JSON.parse(fs.readFileSync(file)).files,files,'BASELINE_CHANGED');else fs.writeFileSync(file,JSON.stringify({files},null,2));console.log(baseline);}
if(require.main===module)main();module.exports={localManifest,root,workspace:repo,hash};
