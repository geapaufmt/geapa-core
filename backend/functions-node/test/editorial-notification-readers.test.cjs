'use strict';
const {test}=require('node:test'),a=require('node:assert/strict');
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {control}=require('../src/act-readers.cjs'),{ROOT,PROJECT}=require('../src/environment.cjs');

test('editorial and notification readers coexist at the bounded rollout capacity',()=>{
 const routes=[...Array.from({length:48},(_,i)=>'/synthetic/reader-'+i),'/admin/gestao-institucional','/admin/ciclos-semestres/parametros','/admin/frequencia-rh','/admin/frequencia-rh/providencias','/admin/conteudo-publico','/meu-vinculo/notificacoes','/meu-vinculo/notificacoes/leitura'];
 const c={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:Date.now()-1000,expiresAt:Date.now()+60000,readModelCommands:[],policyIds:['SYNTHETIC_POLICY'],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYNTHETIC_APP']},specializedReaders:{enabled:true,actorUids:['SYNTHETIC_ACTOR'],routes}};
 for(const route of routes.slice(-3))a.equal(control(c,'SYNTHETIC_ACTOR',route),c);
 const withRoutes=routes=>({...c,specializedReaders:{...c.specializedReaders,routes}});
 a.throws(()=>control(withRoutes([...routes,'/synthetic/overflow']),'SYNTHETIC_ACTOR',routes.at(-1)),/PILOT_DISABLED/);
 a.throws(()=>control(withRoutes([...routes.slice(1),routes.at(-1)]),'SYNTHETIC_ACTOR',routes.at(-1)),/PILOT_DISABLED/);
 a.throws(()=>control(c,'OTHER_ACTOR',routes.at(-1)),/PILOT_DISABLED/);
 a.throws(()=>control({...c,enabled:false},'SYNTHETIC_ACTOR',routes.at(-1)),/PILOT_DISABLED/);
});
