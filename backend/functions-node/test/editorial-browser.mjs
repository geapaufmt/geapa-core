// Loaded only by the local browser acceptance fixture.
import {createInstitutionalEditor} from '/assets/act-ui/institutional-editor-client.mjs';
const {port, email, password, width, project} = fixture;
const endpoint=location.origin+port+'/command';
let step='auth';
const check=(ok,message)=>{if(!ok) throw Error(message);};
async function waitFor(condition) {
  const until=Date.now()+10000;
  while (!condition()) {
    if (Date.now()>until) throw Error('Timed out: '+step);
    await new Promise(ok=>setTimeout(ok,30));
  }
}
const button=action=>document.querySelector('[data-editor-action="'+action+'"]');
async function run(){
  const signin=await fetch('http://127.0.0.1:9197/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=demo',{
    method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})
  });
  const {idToken,localId}=await signin.json();
  check(signin.ok && localId==='CF_ADMIN','Synthetic Auth sign-in failed');
  const headers={'Content-Type':'application/json',Authorization:'Bearer '+idToken,'X-Firebase-AppCheck':'editorial-test-app'};
  const send=async(action,request,custom=headers)=>{
    const response=await fetch(endpoint,{method:'POST',headers:custom,body:JSON.stringify({action,request})});
    return {response,value:await response.json()};
  };
  step='authorized reader';
  const read=async(route,params={})=>{
    const {response,value}=await send('actRead',{route,params});
    check(response.ok,JSON.stringify(value.error));return value.result;
  };
  const context=await read('/act/context');
  check(context.navigation.routes['admin-conteudo-publico']===true,'Missing server navigation hint');
  let route='admin-conteudo-publico',loseResponse=false,lost=false;
  window.PortalGeapaNavigation={getRotaAtual:()=>route};
  window.PortalDevContext={authenticated:true,act:{available:true,navigation:context.navigation}};
  const ticket={uid:localId,generation:1};
  const client=createInstitutionalEditor({
    auth:{capture:()=>ticket,assert:t=>check(t===ticket,'Session changed'),token:async()=>idToken},
    config:{projectId:project,endpoint,siteKey:'emulator'},projectId:project,endpoint,storage:localStorage,
    mode:'editorial',attest:async()=>'editorial-test-app',
    fetchImpl:async(url,opts)=>{
      const response=await fetch(url,opts);
      if(loseResponse && response.ok && JSON.parse(opts.body).action==='editorialExecute'){
        loseResponse=false;lost=true;throw Object.assign(Error('Lost response'),{code:'REQUEST_TIMEOUT'});
      }
      return response;
    }
  });
  window.PortalGeapaApi={apiGet:read,apiPost:async(_route,{payload})=>{
    const q=JSON.parse(payload);
    try {return {ok:true,data:await client.run(q.phase,q.payload)};}
    catch(e){return {ok:false,code:e.code,message:e.message};}
  }};
  const publicReader=await import('/assets/act-command/public-pages.mjs');
  await import('/assets/act-ui/public-admin-editor.mjs');
  const title='Home de ensaio '+width;
  await waitFor(()=>document.querySelector('[data-field="title"]'));
  step='draft';
  const editTitle=async(value)=>{
    const field=document.querySelector('[data-field="title"]'); field.value=value;
    field.dispatchEvent(new Event('input',{bubbles:true}));button('save').click();
    await waitFor(()=>!button('save')?.disabled && button('preview') && !button('preview').disabled);
  };
  await editTitle(title+' v1'); await editTitle(title);
  check(await publicReader.readPublicPageOptional('home')===null,'Draft became public');
  step='private preview';button('preview').click();
  check(!document.querySelector('#public-editor-preview').hidden,'Private preview unavailable');
  step='CAS';
  const competing=await client.run('preview',{action:'PAGE_PUBLISH',pageId:'home',expectedRevision:2});
  step='publish with lost response';loseResponse=true;button('publish').click();
  await waitFor(()=>lost && !button('save')?.disabled);
  button('recover').click();
  await waitFor(()=>document.querySelector('.editor-message')?.textContent.includes('confirmada') && !button('save')?.disabled);
  const page=await publicReader.readPublicPageOptional('home');
  check(page?.revision===3 && page.title===title,'Wrong public revision');
  const stale=await send('editorialExecute',competing.request);
  check(stale.response.status===409 && stale.value.error.code==='REVISION_CONFLICT','CAS conflict missing');
  step='public Home';
  route='inicio';document.dispatchEvent(new CustomEvent('portal:navigationchange'));
  await waitFor(()=>document.querySelector('#home-editorial-content').dataset.homeEditorialPublished==='true');
  route='admin-conteudo-publico';document.dispatchEvent(new CustomEvent('portal:navigationchange'));
  await waitFor(()=>button('unpublish') && !button('unpublish').disabled);
  step='version history';
  document.querySelector('[data-editor-tab="history"]').click();
  await waitFor(()=>button('version'));
  button('version').click();
  await waitFor(()=>document.querySelector('pre')?.textContent.includes(title));
  document.querySelector('[data-editor-tab="home"]').click();
  const overflow=Math.max(0,document.documentElement.scrollWidth-document.documentElement.clientWidth);
  step='unpublish';button('unpublish').click();
  await waitFor(()=>document.querySelector('.editor-message')?.textContent.includes('despublicada') && !button('save')?.disabled);
  check(await publicReader.readPublicPageOptional('home')===null,'Projection still visible');
  step='unauthorized read';
  check(!(await send('actRead',{route:'/admin/conteudo-publico',params:{}},{'Content-Type':'application/json'})).response.ok,'Anonymous admin reader');
  const privateURL='http://127.0.0.1:8187/v1/projects/'+project+'/databases/(default)/documents/canonicalEditorial/home';
  check((await fetch(privateURL)).status===403,'Anonymous canonical read');
  document.dispatchEvent(new CustomEvent('portal:act-session-suspended'));
  check(document.getElementById('public-admin-content').children.length===0,'Private DOM retained');
  check(innerWidth===width,'Unexpected viewport: '+innerWidth);
  return {ok:true,viewport:innerWidth,overflow,publicRevision:page.revision,recovered:lost,cas:stale.response.status};
}
let result;
try {result=await run();} catch(e){result={ok:false,step,message:String(e.message)};}
await fetch(port+'/result',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(result)});
