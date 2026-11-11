const {test}=require('node:test'),a=require('node:assert/strict'),{admitted}=require('../src/dev-institutional-reader-scope.cjs');
test('temporary DEV actor has only two read routes and a bounded window, never production',()=>{
 const row={uid:'synthetic',startsAt:100,expiresAt:1000},c={specializedReaders:{temporaryInstitutionalActors:[row]}};
 for(const route of ['/act/context','/admin/gestao-institucional'])a.equal(admitted(c,'synthetic',route,500,false),true);
 for(const route of ['/members/own','/v2/minha-frequencia','/admin/pessoas'])a.equal(admitted(c,'synthetic',route,500,false),false);
 for(const [uid,now,prod]of [['other',500,false],['synthetic',99,false],['synthetic',1000,false],['synthetic',500,true]])a.equal(admitted(c,uid,'/act/context',now,prod),false);
 row.expiresAt=100+4*3600000+1;a.equal(admitted(c,'synthetic','/act/context',500,false),false);
 row.expiresAt=1000;c.specializedReaders.temporaryInstitutionalActors.push(row);a.equal(admitted(c,'synthetic','/act/context',500,false),false);
});
