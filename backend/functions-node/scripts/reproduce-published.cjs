'use strict';
// Rebuild code from this checkout, then replay verified historical metadata.
// This is an integrity check, never a new release or deployment authorization.
const fs=require('node:fs'),p=require('node:path'),cp=require('node:child_process'),crypto=require('node:crypto'),a=require('node:assert/strict');
const root=p.resolve(__dirname,'..'),hash=b=>crypto.createHash('sha256').update(b).digest('hex');
cp.execFileSync(process.execPath,[p.join(__dirname,'build-domain.cjs')],{stdio:'inherit'});
const current=JSON.parse(fs.readFileSync(p.join(root,'generated/source-manifest.json')));
for(const env of ['dev','prod']){
 const record=require('../releases/published-'+env+'-20261101.json');
 a.deepEqual(current.hashes,record.sources.hashes,'canonical source bytes changed');
 a.equal(current.bundleSha256,record.sources.bundleSha256);
 const out=p.join(root,'release','reproduced-'+env);fs.mkdirSync(out,{recursive:true});
 for(const row of record.release.files){
  const file=row.path;if(file.includes('..')||p.isAbsolute(file))throw Error('RELEASE_PATH_INVALID');
  let b;
  if(file==='generated/source-manifest.json')b=Buffer.from(JSON.stringify(record.sources,null,2)+'\n');
  else if(file==='package.json')b=Buffer.from(JSON.stringify(record.package,null,2)+'\n');
  else if(file==='package-lock.json')b=Buffer.from(JSON.stringify(require('../package-lock.runtime.json'),null,2)+'\n');
  else if(file==='.gcloudignore')b=Buffer.from('node_modules/\n.git/\n');
  else b=fs.readFileSync(p.join(root,file));
  a.equal(hash(b),row.sha256,'published byte mismatch: '+file);
  const dest=p.join(out,file);fs.mkdirSync(p.dirname(dest),{recursive:true});fs.writeFileSync(dest,b);
 }
 fs.writeFileSync(p.join(out,'release-manifest.json'),JSON.stringify(record.release,null,2)+'\n');
 cp.execFileSync(process.execPath,[p.join(out,'scripts/verify-release.cjs')],{stdio:'inherit'});
 console.log(JSON.stringify({environment:env,revision:record.revision,files:record.release.files.length,result:'EXACT_PUBLISHED_REPRODUCTION',metadataReplay:true,deploy:false}));
}
