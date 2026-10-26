'use strict';
// Local packaging only. Deliberately has no deploy/API/IAM commands.
const fs=require('node:fs'),p=require('node:path'),cp=require('node:child_process'),crypto=require('node:crypto');
const root=p.resolve(__dirname,'..'),repo=p.resolve(root,'../..');
const sha=v=>crypto.createHash('sha256').update(v).digest('hex');
const git=(...args)=>cp.execFileSync('git',['-C',repo,...args],{encoding:'utf8'}).trim();
if(git('status','--porcelain','--','backend/functions-node',...require('../core-sources.json').paths))throw Error('CHECKPOINT_REQUIRED');
cp.execFileSync(process.execPath,[p.join(__dirname,'build-domain.cjs')],{stdio:'inherit'});
const target=process.argv.includes('--prod')?'PROD':'DEV';const profile=require('../src/runtime-profile.cjs').select({GEAPA_TARGET_ENVIRONMENT:target});
const commit=git('rev-parse','HEAD'),source=require('../generated/source-manifest.json');
const out=p.join(root,'release',commit.slice(0,12)+(target==='PROD'?'-prod':''));fs.mkdirSync(out,{recursive:true});
const files=[...fs.readdirSync(p.join(root,'src')).filter(f=>f!=='local-network.cjs').map(f=>'src/'+f),'generated/domain.gs','generated/workspace-drive.gs','generated/source-manifest.json','scripts/verify-release.cjs','package.json','package-lock.json'].sort();
for(const file of files){const dest=p.join(out,file);fs.mkdirSync(p.dirname(dest),{recursive:true});fs.copyFileSync(p.join(root,file),dest);}
// Prevent buildpack's default npm build (which needs the repository). Verification
// runs instead, on the exact prebuilt artifact. The original lockfile is retained.
const pkg=JSON.parse(fs.readFileSync(p.join(out,'package.json')));delete pkg.devDependencies;
pkg.scripts={build:'node scripts/verify-release.cjs','gcp-build':'node scripts/verify-release.cjs',start:'functions-framework --target='+profile.functionName,deploy:'node -e "throw Error(\'DEPLOY_NOT_AUTHORIZED\')"'};
fs.writeFileSync(p.join(out,'package.json'),JSON.stringify(pkg,null,2)+'\n');
// npm's cached install normalizes devOptional to optional after dev dependencies
// are pruned. Use the npm-normalized, versioned runtime graph, retaining every pin.
const fullLock=JSON.parse(fs.readFileSync(p.join(root,'package-lock.json')));
const lock=JSON.parse(fs.readFileSync(p.join(root,'package-lock.runtime.json')));
if(JSON.stringify(lock.packages[''].dependencies)!==JSON.stringify(pkg.dependencies))throw Error('RUNTIME_LOCK_ROOT_DRIFT');
for(const [key,value]of Object.entries(lock.packages)){const original=fullLock.packages[key];if(!original||value.version!==original.version||value.resolved!==original.resolved||value.integrity!==original.integrity)throw Error('RUNTIME_LOCK_PIN_DRIFT: '+key);}
fs.writeFileSync(p.join(out,'package-lock.json'),JSON.stringify(lock,null,2)+'\n');
fs.writeFileSync(p.join(out,'.gcloudignore'),'node_modules/\n.git/\n');files.push('.gcloudignore');files.sort();
const inventory=files.map(file=>{const b=fs.readFileSync(p.join(out,file));return {path:file,bytes:b.length,sha256:sha(b)};});
const manifest={schemaVersion:1,environment:target,projectId:profile.projectId,remoteActionsAuthorized:false,sourceCommit:source.sourceCommit,candidateCommit:commit,sourceDigest:source.sourceDigest,files:inventory,inventoryDigest:sha(JSON.stringify(inventory))};
fs.writeFileSync(p.join(out,'release-manifest.json'),JSON.stringify(manifest,null,2)+'\n');
cp.execFileSync(process.execPath,[p.join(out,'scripts/verify-release.cjs')],{stdio:'inherit'});
console.log(JSON.stringify({release:out,manifestSha256:sha(fs.readFileSync(p.join(out,'release-manifest.json')))}));
