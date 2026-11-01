'use strict';
const fs=require('node:fs'),p=require('node:path'),cp=require('node:child_process'),crypto=require('node:crypto');
const root=p.resolve(__dirname,'../../..'),out=p.resolve(__dirname,'../generated');
const files=require('../core-sources.json').paths;
const git=(...args)=>cp.execFileSync('git',['-C',root,...args]);
const hashes={};for(const f of files){if(!/^[a-zA-Z0-9_/-]+\.gs$/.test(f)||f.includes('..'))throw Error('SOURCE_PATH_INVALID');const local=fs.readFileSync(p.join(root,f)),tracked=git('show','HEAD:'+f);if(!local.equals(tracked))throw Error('UNCOMMITTED_CORE_SOURCE: '+f);hashes[f]=crypto.createHash('sha256').update(local).digest('hex');}
const commit=git('log','-1','--format=%H','--',...files).toString().trim();
fs.mkdirSync(out,{recursive:true});
const bundle=files.map(f=>'\n// SOURCE: '+f+'\n'+fs.readFileSync(p.join(root,f),'utf8')).join('\n;\n');
fs.writeFileSync(p.join(out,'domain.gs'),bundle);
fs.writeFileSync(p.join(out,'workspace-drive.gs'),fs.readFileSync(p.join(root,'activities/v2/140_act_workspace_drive.gs'),'utf8'));
const sha=v=>crypto.createHash('sha256').update(v).digest('hex');
fs.writeFileSync(p.join(out,'source-manifest.json'),JSON.stringify({schemaVersion:1,sourceCommit:commit,sourceRoot:'repository root',hashes,sourceDigest:sha(JSON.stringify(hashes)),bundleSha256:sha(bundle)},null,2)+'\n');
console.log(JSON.stringify({sourceCommit:commit,sourceFiles:files.length,sourceDigest:sha(JSON.stringify(hashes))}));
