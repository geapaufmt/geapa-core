'use strict';
// Preservation verification only. No Google API/credential discovery here.
const a=require('node:assert/strict'),fs=require('node:fs'),p=require('node:path'),{root,localManifest}=require('./preflight.cjs');
a.deepEqual(localManifest(),JSON.parse(fs.readFileSync(p.join(root,'evidence/local-before.json'))).files);
console.log({baselineUnchanged:true,remoteChecks:false,remoteWrites:0});
