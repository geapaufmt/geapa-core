'use strict';
const fs=require('node:fs'),p=require('node:path'),crypto=require('node:crypto');
const root=p.resolve(__dirname,'..'),manifest=JSON.parse(fs.readFileSync(p.join(root,'release-manifest.json'),'utf8'));
const sha=v=>crypto.createHash('sha256').update(v).digest('hex');
for(const row of manifest.files){
 if(!/^(src\/[a-z-]+\.(cjs|js)|generated\/(domain\.gs|workspace-drive\.gs|source-manifest\.json)|scripts\/verify-release\.cjs|package(-lock)?\.json|\.gcloudignore)$/.test(row.path)||row.path.includes('..'))throw Error('RELEASE_PATH_INVALID');
 const file=p.join(root,row.path);if(fs.lstatSync(file).isSymbolicLink()||sha(fs.readFileSync(file))!==row.sha256)throw Error('RELEASE_HASH_MISMATCH: '+row.path);
}
const actual=fs.readdirSync(root,{recursive:true}).map(x=>x.replaceAll('\\','/')).filter(x=>!x.startsWith('node_modules/')&&x!=='node_modules'&&fs.lstatSync(p.join(root,x)).isFile()).sort();
const expected=[...manifest.files.map(x=>x.path),'release-manifest.json'].sort();
if(JSON.stringify(actual)!==JSON.stringify(expected))throw Error('UNEXPECTED_RELEASE_FILE');
const source=JSON.parse(fs.readFileSync(p.join(root,'generated/source-manifest.json')));
if(sha(fs.readFileSync(p.join(root,'generated/domain.gs')))!==source.bundleSha256)throw Error('DOMAIN_BUNDLE_MISMATCH');
 if(sha(fs.readFileSync(p.join(root,'generated/workspace-drive.gs')))!==source.hashes['activities/v2/140_act_workspace_drive.gs'])throw Error('WORKSPACE_PROVIDER_SOURCE_MISMATCH');
if(!({DEV:'geapa-convergencia-dev',PROD:'portal-geapa'}[manifest.environment||'DEV']===manifest.projectId)||manifest.remoteActionsAuthorized!==false)throw Error('RELEASE_DESTINATION_INVALID');
console.log(JSON.stringify({verified:manifest.files.length,sourceCommit:source.sourceCommit,inventoryDigest:sha(JSON.stringify(manifest.files))}));
