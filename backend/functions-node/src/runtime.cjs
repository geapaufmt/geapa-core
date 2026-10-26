'use strict';
const {fence,fail,metadataIdentity}=require('./environment.cjs');
let pending;
async function runtime(){
 const env=fence();
 if(!pending)pending=(async()=>{
  const {initializeApp}=require('firebase-admin/app');let credential,serverAuth;
  if(!env.emulator){
   await metadataIdentity();
   const {GoogleAuth,Compute}=require('google-auth-library');
   const auth=new GoogleAuth({projectId:env.projectId,scopes:['https://www.googleapis.com/auth/cloud-platform']});
   const client=await auth.getClient();if(!(client instanceof Compute))fail('DESTINATION_INVALID');
   serverAuth=auth;
   credential={getAccessToken:async()=>{const result=await client.getAccessToken();if(!result.token)fail('DESTINATION_INVALID');return {access_token:result.token,expires_in:Math.max(1,Math.floor(((client.credentials.expiry_date||0)-Date.now())/1000))};}};
  }
  const app=initializeApp({projectId:env.projectId,...(credential?{credential}:{})},'geapa-functions-pilot');
  const db=env.emulator?require('firebase-admin/firestore').getFirestore(app):new (require('@google-cloud/firestore').Firestore)({projectId:env.projectId,databaseId:'(default)',auth:serverAuth});
  return {db,auth:require('firebase-admin/auth').getAuth(app),appCheck:require('firebase-admin/app-check').getAppCheck(app)};
 })().catch(e=>{pending=null;throw e;});return pending;
}
module.exports={runtime};
