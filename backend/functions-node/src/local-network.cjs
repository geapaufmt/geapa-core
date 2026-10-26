'use strict';
// Loaded only in emulator mode, before any Admin SDK or domain module.
const deny=()=>{throw Object.assign(Error('EXTERNAL_NETWORK_DENIED'),{code:'DESTINATION_INVALID'});};
const allowed=(host,port)=>{if(!['127.0.0.1','localhost','::1'].includes(String(host||'localhost').replace(/^\[|\]$/g,''))||!Number.isInteger(Number(port)))deny();};
const net=require('node:net'),connect=net.Socket.prototype.connect;
net.Socket.prototype.connect=function(...args){let x=args[0];if(Array.isArray(x))x=x[0];if(typeof x==='object'){if(x.path)return connect.apply(this,args);allowed(x.host,x.port);}else allowed(typeof args[1]==='string'?args[1]:'localhost',x);return connect.apply(this,args);};
const original=globalThis.fetch;globalThis.fetch=(input,...args)=>{const u=new URL(typeof input==='string'?input:input.url||input);allowed(u.hostname,u.port||80);return original(input,...args);};
require('node:tls').connect=deny;
require('node:dgram').createSocket=deny;
