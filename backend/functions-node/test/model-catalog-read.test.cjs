const {test}=require('node:test'),a=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
function reader(allowed=true){
 const docs=Array.from({length:24},(_,i)=>({id:'SYN_'+String(i).padStart(2,'0'),revision:1,name:'Synthetic '+i,status:i<4?'ACTIVE':'INACTIVE',...(i<4?{}:{governance:{state:'DRAFT',definitionHash:'a'.repeat(64)}})}));let queries=0,rechecks=0;
 const need=(v,code)=>{if(!v)throw Error(code)},db={get:()=>null,query(q){queries++;a.deepEqual(Array.from(q.filters),[]);const cursor=q.cursor?.[0]?.name.split('/').at(-1);return docs.filter(d=>!cursor||d.id>cursor).slice(0,q.limit).map(data=>({data}))}};
 const ctx={f4ActivitiesModel_:()=>({need,id:v=>/^SYN_\d+$/.test(v)}),f4Core_:()=>({exact(){}}),f4Firestore_:()=>({connect:()=>db,DocumentReference:class{constructor(name){this.name=name}}}),f4DevBoundary_:()=>({root:''}),f4SharedAuthz_:()=>({recheck(){rechecks++}})};
 vm.createContext(ctx);vm.runInContext(fs.readFileSync(path.join(__dirname,'../../../activities/v2/149_act_activity_config_compat.gs'),'utf8'),ctx);
 const r=ctx.f4ActivitiesActivityConfigCompat_({config:{projectId:'demo'},ports:{readAuthority(){need(allowed,'AUTHORIZATION_DENIED');return{}},commandAvailable:()=>false}});
 return {read:q=>JSON.parse(JSON.stringify(r.read(q))),stats:()=>({queries,rechecks})};
}
test('administrative catalog includes draft/inactive originals across pages and a fresh reload',()=>{
 const r=reader(),first=r.read({kind:'models'}).data,second=r.read({kind:'models',cursor:first.nextCursor}).data;
 a.equal(first.items.length+second.items.length,24);a.equal(second.items[3].governance.state,'DRAFT');a.equal(second.items[3].status,'INACTIVE');a.equal(first.items.filter(v=>v.status==='ACTIVE').length,4);a.equal(second.nextCursor,null);
 a.deepEqual(reader().read({kind:'models'}),r.read({kind:'models'}));a.equal(r.stats().rechecks,3);
});
test('catalog refuses unauthorized read before querying model data',()=>{const r=reader(false);a.throws(()=>r.read({kind:'models'}),/AUTHORIZATION_DENIED/);a.equal(r.stats().queries,0)});
