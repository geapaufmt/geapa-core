'use strict';
// Opt-in acceptance fixture. Serves only on loopback, against a disposable demo project.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const {spawn} = require('node:child_process');
const express = require('express');
const {initializeApp, deleteApp} = require('firebase-admin/app');
const {getAuth} = require('firebase-admin/auth');
const {handler, ORIGIN} = require('../src/http.cjs');
const {fence, DEMO} = require('../src/environment.cjs');
const {readAct} = require('../src/act-readers.cjs');
const {manageEditorial} = require('../src/editorial-commands.cjs');

async function browserAcceptance(db, seed) {
  assert.ok(fence().emulator);
  assert.equal(db.projectId, DEMO);
  const dist = process.env.GEAPA_EDITORIAL_PORTAL_DIST;
  const chrome = process.env.GEAPA_EDITORIAL_CHROME;
  assert.ok(dist && fs.existsSync(path.join(dist, 'assets/act-ui/public-admin-editor.mjs')), 'Build Portal DEV first');
  assert.ok(chrome && fs.existsSync(chrome), 'GEAPA_EDITORIAL_CHROME must name the local Chromium executable');
  const cssLinks = [...fs.readFileSync(path.join(dist,'index.html'),'utf8').matchAll(/<link\b[^>]*rel="stylesheet"[^>]*href="([^"]+\.css)"[^>]*>/gi)]
    .map(match=>'/'+match[1].replace(/^\//,''));
  assert.ok(cssLinks.length && cssLinks.every(url=>/^\/[\w./-]+\.css$/.test(url) && !url.includes('..')));
  const demoApp = initializeApp({projectId:DEMO},'editorial-browser-test');
  const auth = getAuth(demoApp);
  const password = crypto.randomBytes(24).toString('hex');
  const email = 'editorial-' + crypto.randomUUID() + '@example.invalid';
  const appCheck = {async verifyToken(token) {
    assert.equal(token,'editorial-test-app'); return {appId:'SYN_APP'};
  }};
  let user;
  try {
    user = await auth.createUser({uid:'CF_ADMIN',email,password,emailVerified:true});
    for (const width of [1280,390]) {
      await seed();
      const app = express();
      const nonce = crypto.randomUUID();
      let base;
      let receive, rejectResult;
      const completed = new Promise((ok,fail)=>{receive=ok; rejectResult=fail;});
      const port = '/'+nonce;
      app.disable('x-powered-by');
      app.use((req,res,next)=>{
        if (req.headers.host !== new URL(base).host || (req.method === 'POST' && req.headers.origin !== base)) return res.sendStatus(403);
        res.set('Cache-Control','no-store'); next();
      });
      for(const url of cssLinks) app.get(url,(_req,res)=>res.sendFile(path.join(dist,url)));
      app.use('/assets/act-ui',express.static(path.join(dist,'assets/act-ui'),{index:false}));
      app.get(port,(_req,res)=>res.type('html').send('<!DOCTYPE html><style>*{margin:0;border:0}iframe{display:block}</style><iframe width="'+width+'" height="810" src="'+port+'/frame"></iframe>'));
      app.get('/assets/act-command/public-pages.mjs',(_req,res)=>res.type('application/javascript').send(`
        function decode(v) {
          if(v.mapValue) return Object.fromEntries(Object.entries(v.mapValue.fields||{}).map(([k,x])=>[k,decode(x)]));
          if(v.arrayValue) return (v.arrayValue.values||[]).map(decode);
          if (v.integerValue !== undefined) return Number(v.integerValue);
          return Object.values(v)[0];
        }
        export async function readPublicPageOptional(id) {
          if(id!=='home') throw Error('Unexpected fixture page');
          const r=await fetch('http://127.0.0.1:8187/v1/projects/demo-geapa-functions-local/databases/(default)/documents/publicPages/home');
          if(r.status===404) return null;
          if(!r.ok) throw Error('Public client: '+r.status);
          return decode({mapValue:await r.json()});
        }
      `));
      app.get(port+'/frame',(_req,res)=>res.type('html').send('<!DOCTYPE html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+cssLinks.map(url=>'<link rel="stylesheet" href="'+url+'">').join('')+'<main id="public-admin-content"></main><div id="home-editorial-content" hidden></div><script type="module" src="'+port+'/test.mjs"></script></html>'));
      app.get(port+'/test.mjs',(_req,res)=>res.type('application/javascript').send(
        'const fixture='+JSON.stringify({email,password,port,width,project:DEMO})+';\n'+fs.readFileSync(path.join(__dirname,'editorial-browser.mjs'),'utf8')
      ));
      app.use(express.json({limit:'100kb'}));
      app.post(port+'/result',(req,res)=>{res.sendStatus(204); receive(req.body);});
      const commands = handler({runtime:async()=>({db,auth,appCheck}),log:()=>{},
        actions:['actRead','editorialPreview','editorialExecute','editorialStatus'],
        execute:options=>options.body.action==='actRead'?readAct(options):manageEditorial(options)});
      app.post(port+'/command',(req,res)=>{
        // The fixture has already checked its exact loopback origin; model Hosting at the HTTP boundary.
        req.headers.origin = ORIGIN;
        return commands(req,res);
      });
      const server=await new Promise((ok,fail)=>{
        const srv=app.listen(0,'127.0.0.1',()=>ok(srv)); srv.once('error',fail);
      });
      base='http://127.0.0.1:'+server.address().port;
      const profile = fs.mkdtempSync(path.join(os.tmpdir(),'geapa-editorial-browser-'));
      const child = spawn(chrome,['--headless=new','--no-first-run','--disable-background-networking',
        '--disable-extensions','--disable-sync','--user-data-dir='+profile,'--window-size='+width+',850',base+port],
        {windowsHide:true,stdio:'ignore'});
      child.once('error',rejectResult);
      child.once('exit',code=>rejectResult(Error('Browser closed: '+code)));
      const timeout=setTimeout(()=>rejectResult(Error('Browser acceptance timed out at '+width)),90000);
      try {
        const result=await completed;
        assert.equal(result.ok,true,JSON.stringify(result));
        assert.ok(result.viewport>0 && result.overflow===0,JSON.stringify(result));
        console.log('Editorial browser:',JSON.stringify(result));
      } finally {
        clearTimeout(timeout);
        if (child.exitCode===null && !child.killed) {
          const closed=new Promise(ok=>child.once('close',ok)); child.kill(); await closed;
        }
        server.closeAllConnections(); await new Promise(ok=>server.close(ok));
        const resolvedProfile=path.resolve(profile);
        assert.equal(path.dirname(resolvedProfile),path.resolve(os.tmpdir()));
        assert.ok(path.basename(resolvedProfile).startsWith('geapa-editorial-browser-'));
        fs.rmSync(resolvedProfile,{recursive:true,force:true,maxRetries:5,retryDelay:100});
      }
    }
  } finally {
    try {if(user) await auth.deleteUser(user.uid);} finally {await deleteApp(demoApp);}
  }
}
module.exports={browserAcceptance};
