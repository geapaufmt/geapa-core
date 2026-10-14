'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),health=require('./api/health.js');
function call(method='GET'){const headers={};let body;const res={setHeader:(k,v)=>headers[k]=v,end:v=>body=v};health({method,body:{secret:'DO_NOT_ECHO'}},res);return {status:res.statusCode,body,headers};}
test('health is sanitized and has no privileged access',()=>{const r=call();a.equal(r.status,200);const b=JSON.parse(r.body);a.equal(b.privilegedAccess,false);a.equal(b.commandsEnabled,false);a(!r.body.includes('DO_NOT_ECHO'));a.equal(r.headers['Cache-Control'],'no-store');});
test('successive requests show same instance without claiming cloud cold-start',()=>{const one=JSON.parse(call().body),two=JSON.parse(call().body);a.equal(one.instance,two.instance);a.equal(two.firstInvocation,false);});
test('HEAD has no body',()=>a.equal(call('HEAD').body,''));
test('POST cannot execute a command',()=>a.equal(call('POST').status,405));
test('production fails closed',()=>{const old=process.env.VERCEL_ENV;process.env.VERCEL_ENV='production';try{a.equal(call().status,403);}finally{if(old===undefined)delete process.env.VERCEL_ENV;else process.env.VERCEL_ENV=old;}});
