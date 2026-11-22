const {test}=require('node:test'),a=require('node:assert/strict'),{admitted,routes}=require('../src/dev-linked-reader-scope.cjs');
test('synthetic linked reader requires live grant and matching binding; never grants administration or PROD',()=>{
 const person='DEVADM_0123456789ab_1',g={schemaVersion:'core-dev-linked-reader-grant-v1',projectId:'geapa-convergencia-dev',uid:'synthetic',personId:person,bindingVersion:2,revision:1,active:true,startsAt:100,expiresAt:1000},b={active:true,idPessoa:person,bindingVersion:2},p={active:true,idPessoa:person,origin:'DEV_ADMIN_SYNTHETIC_V1'},docs={'coreDevLinkedReaderGrants/synthetic':g,'memberAuthBindings/synthetic':b,['memberIdentities/'+person]:p},read=k=>docs[k];
 for(const route of routes)a.equal(admitted(read,'synthetic',route,500,false),true);
 for(const route of ['/admin/pessoas','/admin/gestao-institucional','/atividades/chamada','/meu-perfil/salvar'])a.equal(admitted(read,'synthetic',route,500,false),false);
 a.equal(admitted(()=>{throw Error('PROD must not read DEV grants')},'synthetic','/members/own',500,true),false);
 a.equal(admitted(read,'other','/members/own',500,false),false);
 for(const now of [99,1000])a.equal(admitted(read,'synthetic','/members/own',now,false),false);
 for(const [o,k,v] of [[g,'active',false],[g,'expiresAt',14400101],[g,'personId','PES-000001'],[b,'bindingVersion',3],[b,'active',false],[p,'origin','REAL']]){const old=o[k];o[k]=v;a.equal(admitted(read,'synthetic','/members/own',500,false),false);o[k]=old;}
});
