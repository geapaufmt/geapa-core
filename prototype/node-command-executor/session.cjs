'use strict';
const fs=require('node:fs'),p=require('node:path'),assert=require('node:assert/strict'),S=require('./sdk.cjs');
const save=(n,v)=>fs.writeFileSync(p.join(__dirname,'evidence',n),JSON.stringify(v,null,2)+'\n');
async function session(){
 const opened=await S.open(),{initializeApp,deleteApp}=require('firebase-admin/app'),{getAuth}=require('firebase-admin/auth');
 const app=initializeApp({projectId:S.PROJECT,credential:opened.credential},'B0_LOCAL_'+process.pid),auth=getAuth(app);
 const P=require('./private-session.cjs'),secret=P.secrets();assert.equal(secret.project,S.PROJECT);const tokens={};
 async function login(kind){if(tokens[kind])return tokens[kind];const a=secret.accounts[kind];assert(a);const v=(await P.authApi('signInWithPassword',{email:a.email,password:a.password,returnSecureToken:true})).value;assert.equal(v.localId,a.uid);await auth.verifyIdToken(v.idToken,true);return tokens[kind]={uid:a.uid,token:v.idToken};}
 return {...opened,verify:t=>auth.verifyIdToken(t,true),login,async close(hosts=[]){const file=p.join(__dirname,'evidence/seed.json'),j=JSON.parse(fs.readFileSync(file));j.writtenPaths=[...new Set([...(j.writtenPaths||[]),...hosts.flatMap(h=>[...h.touchedAll])])];save('seed.json',j);await opened.db.terminate();await deleteApp(app);}};
}
module.exports={session,save};
