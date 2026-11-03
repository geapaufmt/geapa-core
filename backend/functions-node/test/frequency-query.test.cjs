require('../src/local-network.cjs');
const {test}=require('node:test'),a=require('assert/strict');
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {queryPlan}=require('../src/query.cjs'),{encode,decode}=require('../src/firestore.cjs');
const db={collection(){return {where(){return this},orderBy(){return this},limit(){return this},startAfter(){return this}}},doc:p=>p};
test('derived frequency producer can read bounded canonical semester catalog, no arbitrary filters or private collection',()=>{const q={collection:'academicSemesters',filters:[],limit:20,order:[{field:'__name__',direction:'ASCENDING'}]};a(queryPlan(db,q,decode));for(const patch of [{limit:21},{collection:'memberPrivateProfiles'},{filters:[{field:'cycleId',op:'NOT_EQUAL',value:encode('C')}]},{order:[{field:'private',direction:'ASCENDING'}]},{cursor:[{referenceValue:'projects/portal-geapa/databases/(default)/documents/academicSemesters/S'}]}])a.throws(()=>queryPlan(db,{...q,...patch},decode));});
