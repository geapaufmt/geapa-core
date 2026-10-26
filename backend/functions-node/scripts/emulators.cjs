'use strict';
const p=require('node:path'),fs=require('node:fs'),cp=require('node:child_process');
const root=p.resolve(__dirname,'..'),node=p.join(root,'node_modules/node-win-x64/bin/node.exe');
const cycleSettingsTest=process.argv[2]==='cycle-settings';if(cycleSettingsTest)process.argv[2]='remaining';
const preagendaTest=process.argv[2]==='preagenda';if(preagendaTest)process.argv[2]='remaining';
const continuityTest=process.argv[2]==='presentation-continuity';if(continuityTest)process.argv[2]='remaining';
const promotionTest=process.argv[2]==='activity-promotion';if(promotionTest)process.argv[2]='remaining';
const reviewTest=process.argv[2]==='administrative-review';if(reviewTest)process.argv[2]='remaining';
const productionMembers=process.argv[2]?.startsWith('production-');if(productionMembers)process.argv[2]=process.argv[2].slice(11);
const storageTest=['storage-notice','storage-download','storage-historical'].includes(process.argv[2])?process.argv[2]:null;if(storageTest)process.argv[2]='storage';
const env={...process.env};for(const k of Object.keys(env))if(/TOKEN|SECRET|PASSWORD|CREDENTIAL|PRIVATE_KEY|API_KEY|^GOOGLE_|^GCLOUD_|^CLOUDSDK_|^FIREBASE_/i.test(k))delete env[k];
Object.assign(env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197',NO_GCE_CHECK:'true',FIREBASE_CLI_DISABLE_MOTD:'1',CI:'true',XDG_CONFIG_HOME:p.join(root,'.config'),PATH:p.dirname(node)+p.delimiter+process.env.PATH});
if(['members','bootstrap','storage','workspace','remaining','real-scope','readers','flows','shared','shared-expanded','recovery-shared','recovery-flows','activities'].includes(process.argv[2]))env.GEAPA_PILOT_MODE='SHARED_DEV';
if(['storage','remaining','real-scope','flows','activities','shared-expanded'].includes(process.argv[2]))env.GEAPA_ACTIVITY_COMMANDS_ENABLED='true';
if(['storage','remaining','readers','flows','recovery-flows'].includes(process.argv[2]))env.GEAPA_ACT_FLOWS_ENABLED='true';
fs.mkdirSync(p.join(root,'evidence'),{recursive:true});
env.GEAPA_TEST_EVIDENCE_DIR=p.join(root,'evidence','local-runs',(process.argv[2]||'integration')+'-'+new Date().toISOString().replace(/[:.]/g,'-'));
fs.mkdirSync(env.GEAPA_TEST_EVIDENCE_DIR,{recursive:true});
const cli=p.join(root,'node_modules/firebase-tools/lib/bin/firebase.js');
if(process.argv[2]==='members')env.GEAPA_MEMBERS_ENABLED='true';
if(process.argv[2]==='readers')env.GEAPA_ACT_READERS_ENABLED='true';
if(process.argv[2]==='remaining')env.GEAPA_ACT_EXTENSIONS_ENABLED='true';
if(process.argv[2]==='workspace')env.GEAPA_ACT_WORKSPACE_ENABLED='true';
if(process.argv[2]==='storage'){env.GEAPA_ACT_STORAGE_ENABLED='true';env.GEAPA_ACT_EXTENSIONS_ENABLED='true';}
if(productionMembers){env.GEAPA_TARGET_ENVIRONMENT='PROD';delete env.GEAPA_PILOT_MODE;}
const testFile=process.argv[2]==='members'?'members-integration.cjs':process.argv[2]==='bootstrap'?'bootstrap-integration.cjs':process.argv[2]==='storage'?'storage-integration.cjs':process.argv[2]==='workspace'?'workspace-download-integration.cjs':process.argv[2]==='remaining'?'remaining-integration.cjs':process.argv[2]==='real-scope'?'real-scope-integration.cjs':process.argv[2]==='list-directed'?'list-directed-integration.cjs':process.argv[2]==='readers'?'readers-integration.cjs':process.argv[2]==='flows'?'flows-integration.cjs':process.argv[2]==='activities'?'activities-integration.cjs':process.argv[2]?.startsWith('recovery')?'recovery-integration.cjs':process.argv[2]==='list'?'list-integration.cjs':['shared','shared-expanded'].includes(process.argv[2])?'shared-integration.cjs':'integration.cjs';
const child=cp.spawn(node,[cli,'emulators:exec','--only','functions,auth,firestore','--project','demo-geapa-functions-local','--config','firebase.emulator.json','--non-interactive','"'+node+'" test/'+(cycleSettingsTest?'cycle-settings-integration.cjs':preagendaTest?'preagenda-integration.cjs':continuityTest?'presentation-continuity-integration.cjs':promotionTest?'activity-promotion-integration.cjs':reviewTest?'administrative-review-integration.cjs':storageTest?storageTest+'-integration.cjs':testFile)],{cwd:root,env,windowsHide:true,stdio:['ignore','pipe','pipe']});
const out=fs.createWriteStream(p.join(env.GEAPA_TEST_EVIDENCE_DIR,'emulators.log'));for(const s of [child.stdout,child.stderr])s.on('data',d=>{out.write(d);process.stdout.write(d);});child.on('error',e=>{console.error(e.message);process.exitCode=1;});child.on('close',code=>{out.end();console.log('Evidence: '+env.GEAPA_TEST_EVIDENCE_DIR);process.exitCode=code||0;});
