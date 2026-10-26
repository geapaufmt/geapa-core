'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),{createDomain}=require('../src/domain.cjs');
const boundary={schemaVersion:'act-cycle-boundary-v1',cycleIds:['CYCLE_A'],readOnlyPaths:['activityTypes/meeting'],commands:['ACT.CALENDAR_RESERVE','ACT.ATTENDANCE_FINALIZE','ACT.THEME_SUBMIT']};
const docs={
 'presentationCycles/CYCLE_A':{id:'CYCLE_A',recordMode:'NEW_V2'},'presentationCycles/CYCLE_B':{id:'CYCLE_B',recordMode:'NEW_V2'},
 'canonicalActivities/a':{id:'a',cycleRef:{kind:'cycle',id:'CYCLE_A'},recordMode:'NEW_V2'},
 'canonicalActivities/b':{id:'b',cycleRef:{kind:'cycle',id:'CYCLE_B'},recordMode:'NEW_V2'},
 'canonicalActivities/old':{id:'old',cycleRef:{kind:'cycle',id:'CYCLE_A'},recordMode:'HISTORICAL_PARTIAL'},
 'canonicalPresentations/p':{id:'p',cycleRef:{kind:'cycle',id:'CYCLE_A'},activityRef:{kind:'activity',id:'a'},recordMode:'NEW_V2'},
 'activityAttendanceCollections/a':{id:'a'},'activityAttendance/fact':{activityRef:{kind:'activity',id:'a'},personRef:{kind:'person',id:'P1'}},
 'activityJustifications/j':{attendanceRef:{kind:'attendance',id:'fact'}},
 'presentationMaterialRequirements/req':{presentationRef:{kind:'presentation',id:'p'}},
 'presentationMaterialSubmissions/sub':{requirementRef:{kind:'requirement',id:'req'}},
 'activityTypes/meeting':{id:'meeting'},'jobOutbox/file':{targetRef:{kind:'presentation',id:'p'}}
};
function check(path,{patch={},scope=boundary,readOnly=false}={}){return JSON.parse(createDomain(()=>{throw Error('UNEXPECTED_IO')}).run(x=>f4ActCycleBoundary_().contains(p=>x.docs[p],x.path,x.scope,x.readOnly),{docs:{...docs,...patch},path,scope,readOnly}));}
for(const path of ['presentationCycles/CYCLE_A','canonicalActivities/a','canonicalActivities/a/revisions/1','canonicalPresentations/p','activityAttendanceCollections/a','activityAttendance/fact','activityJustifications/j','presentationMaterialRequirements/req','presentationMaterialSubmissions/sub','jobOutbox/file'])test('cycle membership through canonical lineage: '+path,()=>a(check(path)));
for(const path of ['canonicalActivities/b','canonicalActivities/old','canonicalActivities/missing','memberPrivateProfiles/P1','activityTypes/meeting'])test('denied without approved operational lineage: '+path,()=>a.equal(check(path),false));
test('catalog exact read only, never a writable root',()=>{a(check('activityTypes/meeting',{readOnly:true}));a.equal(check('activityTypes/other',{readOnly:true}),false);});
test('conflicting cycle and activity lineage denied',()=>a.equal(check('canonicalPresentations/p',{patch:{'canonicalPresentations/p':{cycleRef:{kind:'cycle',id:'CYCLE_A'},activityRef:{kind:'activity',id:'b'}}}}),false));
test('cycle missing or historical cannot authorize',()=>{a.equal(check('canonicalActivities/a',{patch:{'presentationCycles/CYCLE_A':null}}),false);a.equal(check('canonicalActivities/a',{patch:{'presentationCycles/CYCLE_A':{id:'CYCLE_A',recordMode:'HISTORICAL_PARTIAL'}}}),false);});
test('loops and unresolved parents fail closed',()=>a.equal(check('canonicalPresentations/p',{patch:{'canonicalPresentations/p':{cycleRef:{kind:'cycle',id:'CYCLE_A'},presentationRef:{kind:'presentation',id:'p'}}}}),false));
test('read catalog cannot smuggle a Members path',()=>a.throws(()=>check('canonicalActivities/a',{scope:{...boundary,readOnlyPaths:['memberPrivateProfiles/P1']}})));
test('promotion/review remain separately approved exact resources',()=>{for(const command of ['ACT.ACTIVITY_PROMOTE_OPERATIONAL','ACT.ACTIVITY_ADMIN_REVIEW','ACT.HISTORICAL_PREAGENDA_RECOGNIZE'])a.throws(()=>check('canonicalActivities/a',{scope:{...boundary,commands:[command]}}));});
test('bounded scope and unsupported extra fields rejected',()=>{a.throws(()=>check('canonicalActivities/a',{scope:{...boundary,cycleIds:[]}}));a.throws(()=>check('canonicalActivities/a',{scope:{...boundary,allowAll:true}}));});
test('derived history inherits original subject, not arbitrary id',()=>{a(check('activityHistory/h',{patch:{'activityHistory/h':{snapshot:{subjectRef:{kind:'activity',id:'a'}}}}}));a.equal(check('activityHistory/h',{patch:{'activityHistory/h':{snapshot:{subjectRef:{kind:'activity',id:'b'}}}}}),false);});
test('policy cycle gate preserves SELF ownership check',()=>{
 const policy={ownPerson:true,allowedTargetScopes:['SELF'],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:[],resourceScope:{kind:'OWN_ATTENDANCE',cycleIds:['CYCLE_A']}};
 const run=(person,patch={})=>createDomain(()=>{throw Error('IO')}).run(x=>{f4ActivitiesAuthzDomain_().match(p=>x.docs[p],{personId:x.person,scopeKey:'SELF',resources:['activityAttendance/fact'],command:null},x.policy,x.person);return true;},{docs:{...docs,'activityAttendance/fact':{activityRef:{kind:'activity',id:'a'},personRef:{kind:'person',id:'P1'},ownerPersonRef:'P1'},...patch},person,policy});
 a.equal(run('P1'),'true');a.throws(()=>run('P2'));a.throws(()=>run('P1',{'canonicalActivities/a':docs['canonicalActivities/b']}));
});

test('publication uses protected visibility limit, never browser authority',()=>{const d=createDomain(()=>{throw Error('IO')});for(const v of ['BOARD','MEMBERS'])a.equal(d.run(x=>{f4ActivityExecutorRoutes_().publication(['BOARD','MEMBERS'],x);return true;},v),'true');for(const v of ['PUBLIC','UNKNOWN',null])a.throws(()=>d.run(x=>f4ActivityExecutorRoutes_().publication(['BOARD','MEMBERS'],x),v));});

test('multi-activity exceptional window requires every activity in the approved cycle',()=>{const w={activityRefs:[{kind:'activity',id:'a'}]};a(check('activityExceptionalJustificationWindows/w',{patch:{'activityExceptionalJustificationWindows/w':w}}));w.activityRefs.push({kind:'activity',id:'b'});a.equal(check('activityExceptionalJustificationWindows/w',{patch:{'activityExceptionalJustificationWindows/w':w}}),false);});

test('external-effect reconciliation inherits only the referenced canonical outbox',()=>{a(check('commandReconciliation/r',{patch:{'commandReconciliation/r':{effectId:'file'}}}));a.equal(check('commandReconciliation/r',{patch:{'commandReconciliation/r':{effectId:'missing'}}}),false);a.equal(check('commandReconciliation/r',{patch:{'commandReconciliation/r':{effectId:'file'},'jobOutbox/file':{targetRef:{kind:'activity',id:'b'}}}}),false);});

test('reconciliation cannot contradict its external effect cycle',()=>a.equal(check('commandReconciliation/r',{patch:{'commandReconciliation/r':{effectId:'file',cycleRef:{kind:'cycle',id:'CYCLE_B'}}}}),false));
