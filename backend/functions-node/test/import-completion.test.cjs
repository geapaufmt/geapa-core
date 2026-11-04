const {test}=require('node:test'),assert=require('node:assert/strict'),H=require('./fixtures/import-reconstruction.cjs');
const {fixture,intentFor,setup,hash,ref,plain}=H;
function adjustment(d,patch={}){
 const source=d['migrationEvidence/EA'];d['migrationEvidence/SYN_DECISION']={id:'SYN_DECISION',schemaVersion:'synthetic-decision-v1',decision:'Current documented recognition, not retrospective approval'};
 d['presentationPolicyReferences/SYN_ADJUST']={id:'SYN_ADJUST',revision:1,usage:'IMPORTED_ACTIVITY_DOCUMENTARY_ADJUSTMENT',active:true,approved:true,activityId:'SYN_IMPORT_ACTIVITY',sourceHash:hash(source),historicalApprovalClaimed:false,reason:'Synthetic current decision',decisionEvidenceRef:ref('evidence','SYN_DECISION'),decisionEvidenceHash:hash(d['migrationEvidence/SYN_DECISION']),...patch};
 const i=intentFor(d);i.documentaryAdjustmentRef=ref('policy','SYN_ADJUST');return i;
}
test('current timing decision preserves the literal source and actual finalization',()=>{
 const d=fixture(),before=plain(d),i=adjustment(d,{times:{start:'18:45',end:'20:30',durationMinutes:105}}),p=setup(d).plan(i);
 assert.deepEqual(d['migrationEvidence/EA'],before['migrationEvidence/EA']);assert.equal(p.writes[0].data.typeSnapshot.durationMinutes,105);
 assert.equal(p.writes[0].data.documentaryAdjustment.sourceRef.id,'EA');assert.equal(p.writes[1].data.completionRecognition.kind,'ADMINISTRATIVE_CONVENTION');
});
test('two documented speakers have equal roles, no invented primary and no private profile fields',()=>{
 const d=fixture();d['activityTypes/SYN_TYPE'].portalModel={requiresPrincipal:true,memberPresentation:false};
 for(const id of ['SYN_SPEAKER_A','SYN_SPEAKER_B'])d['memberIdentities/'+id]={idPessoa:id,displayName:id,technicalAccount:false};
 const i=adjustment(d,{principals:['SYN_SPEAKER_A','SYN_SPEAKER_B'].map(id=>({personRef:ref('person',id),participationRole:'PALESTRANTE'}))}),p=setup(d).plan(i),v=p.writes[0].data;
 assert.equal(v.personRef,null);assert.equal(v.modelContent.principal,undefined);assert.equal(v.modelContent.principals.length,2);assert.equal(d['memberIdentities/SYN_SPEAKER_A'].revision,undefined);
 delete d['memberIdentities/SYN_SPEAKER_B'];assert.throws(()=>setup(d).plan(i));
});
test('assembly exception has no absence effect and does not change the global template or individual results',()=>{
 const d=fixture(21);d['activityJustifications/SYN_J'].state='APPROVED';d['migrationEvidence/EA'].literal.CONTA_FALTA='NAO';d['migrationEvidence/EA'].literal.PERMITE_JUSTIFICATIVA='NAO';
 for(const [p,f]of Object.entries(d))if(p.startsWith('activityAttendance/')){f.countsAbsence=false;d['migrationEvidence/'+f.provenanceHead.id].literal.CONTA_FALTA='NAO';}
 const template=hash(d['activityTypes/SYN_TYPE']),i=adjustment(d,{snapshotOverrides:{countsAbsence:false,allowsJustification:false,portalModel:{requiresPrincipal:false,requiresMaterial:false}},times:{start:'18:45',end:'20:30',durationMinutes:105}}),p=setup(d).plan(i);
 assert.equal(hash(d['activityTypes/SYN_TYPE']),template);assert.equal(p.writes[0].data.typeSnapshot.countsAbsence,false);
 for(const w of p.writes.filter(w=>w.path.startsWith('activityAttendance/'))){assert.equal(w.data.officialCountBasis.countsAbsence,false);assert.equal(w.data.status,d[w.path].status);}
});
test('dynamic 105 minutes and no certificate are activity-specific; protected facts remain unchanged',()=>{
 const d=fixture(21),before=plain(d);const i=adjustment(d,{times:{start:'18:45',end:'20:30',durationMinutes:105},snapshotOverrides:{portalModel:{requiresAxis:false,generatesCertificate:false}}}),p=setup(d).plan(i);
 assert.equal(p.writes[0].data.typeSnapshot.portalModel.generatesCertificate,false);assert.equal(p.writes[0].data.typeSnapshot.durationMinutes,105);
 for(const w of p.writes.filter(w=>w.path.startsWith('activityAttendance/')))for(const k of ['id','status','personRef','countsAttendance','countsAbsence','applicableAtDate','provenanceHead'])assert.deepEqual(w.data[k],before[w.path][k]);
});
for(const bad of [{times:{start:'18:45',end:'20:30',durationMinutes:120}},{snapshotOverrides:{visibility:'PUBLIC'}},{snapshotOverrides:{portalModel:{memberPresentation:false}}}])test('rejects unapproved or inconsistent adjustment '+JSON.stringify(bad),()=>{const d=fixture(),i=adjustment(d,bad);assert.throws(()=>setup(d).plan(i));});
function reopened(){
 const d=fixture(),e=d['migrationEvidence/SYN_CLOSURE'];e.equivalentRecordAbsent=false;e.events=[{activityId:'SYN_IMPORT_ACTIVITY',type:'CHAMADA_FINALIZADA',at:'2026-09-11T00:30:00Z',active:true,success:true,basis:'LEGACY_PORTAL_FINALIZATION',sourceRef:'SYN_FINAL',sourceHash:hash('final')},{activityId:'SYN_IMPORT_ACTIVITY',type:'CHAMADA_REABERTA',at:'2026-09-12T00:00:00Z',active:true,success:true,basis:'LEGACY_PORTAL_FINALIZATION',sourceRef:'SYN_REOPEN',sourceHash:hash('reopen')}];
 const material=Object.entries(d).filter(([p])=>p.startsWith('activityAttendance/')).map(([,f])=>({id:f.id,personRef:f.personRef,status:f.status,countsAttendance:f.countsAttendance,countsAbsence:f.countsAbsence,applicableAtDate:f.applicableAtDate}));
 d['migrationEvidence/SYN_COMPARE']={id:'SYN_COMPARE',materialBefore:material,materialAfter:material};d['presentationPolicyReferences/SYN_REOPEN']={id:'SYN_REOPEN',usage:'IMPORTED_REOPENING_RECOGNITION',active:true,approved:true,activityId:'SYN_IMPORT_ACTIVITY',closureEvidenceHash:hash(e),beforeMaterialHash:hash(material),afterMaterialHash:hash(material),noMaterialChange:true,reason:'No material effect, keep original closure and audit',comparisonEvidenceRef:ref('evidence','SYN_COMPARE'),comparisonEvidenceHash:hash(d['migrationEvidence/SYN_COMPARE'])};const i=intentFor(d);i.reopeningRecognitionRef=ref('policy','SYN_REOPEN');return {d,i};
}
test('reopening with sealed unchanged facts retains the real closure and pending late submission',()=>{
 const {d}=reopened();d['activityJustifications/SYN_J'].submittedAt=Date.parse('2026-10-01T16:00Z');const i=intentFor(d);i.reopeningRecognitionRef=ref('policy','SYN_REOPEN');const p=setup(d).plan(i),c=p.writes[1].data,j=p.writes.find(w=>w.path==='activityJustifications/SYN_J').data;
 assert.equal(c.lastFinalizedAt,Date.parse('2026-09-11T00:30Z'));assert.equal(c.completionRecognition.reopeningRecognition.originalEventPreserved,true);assert.equal(d['migrationEvidence/SYN_CLOSURE'].events.length,2);assert.equal(j.state,'PENDING');assert.equal(j.basisRecognition.submissionWasLate,true);assert.equal(j.basisRecognition.lateSubmissionAccepted,false);assert.equal(c.acceptLateJustification,false);
});
test('changed material comparison refuses reopening recognition',()=>{const {d}=reopened();d['presentationPolicyReferences/SYN_REOPEN'].beforeMaterialHash=hash('different');const i=intentFor(d);i.reopeningRecognitionRef=ref('policy','SYN_REOPEN');assert.throws(()=>setup(d).plan(i),/REOPENING_MATERIAL_CHANGE/);});
function noAttendance(variant){const d=fixture();for(const p of Object.keys(d))if(p.startsWith('activityAttendance/')||p.startsWith('activityAttendanceCollections/')||p.startsWith('activityJustifications/'))delete d[p];const i=adjustment(d);i.variant=variant;i.factIds=[];i.justificationIds=[];const dec=d['presentationPolicyReferences/SYN_ADJUST'];dec.usage=variant;return {d,i,dec};}
test('no-call recognition changes only documentary classification; orphan justification preserved',()=>{const {d}=noAttendance('HISTORICAL_NO_CALL');d['activityJustifications/ORPHAN']={id:'ORPHAN',state:'APPROVED',historical:true};const i=intentFor(d);i.variant='HISTORICAL_NO_CALL';i.factIds=[];i.justificationIds=[];i.documentaryAdjustmentRef=ref('policy','SYN_ADJUST');const p=setup(d).plan(i);assert.equal(p.writes.length,1);assert.equal(p.writes[0].data.attendanceRecognition.state,'NO_CALL');assert.equal(p.writes[0].data.recordMode,'HISTORICAL_PARTIAL');assert.equal(d['activityJustifications/ORPHAN'].state,'APPROVED');});
test('historical reschedule preserves old source and presentation identity without generating a call',()=>{
 const {d,dec}=noAttendance('HISTORICAL_RESCHEDULE'),act=d['canonicalActivities/SYN_IMPORT_ACTIVITY'];act.legacyStatus='PLANEJADA';d['migrationEvidence/EA'].literal.STATUS_OPERACIONAL='PLANEJADA';dec.sourceHash=hash(d['migrationEvidence/EA']);Object.assign(dec,{civilDate:'2026-11-26',academicSemesterId:'SYN_SEM',start:'18:45',end:'20:30'});
 d['canonicalPresentations/SYN_PRESENTATION']={id:'SYN_PRESENTATION',revision:1,recordMode:'HISTORICAL_PARTIAL',provenanceHead:ref('evidence','EA'),activityRef:{resolution:'RESOLVED',ref:ref('activity',act.id)},legacyStatus:'PLANEJADA'};
 const i=intentFor(d);Object.assign(i,{variant:'HISTORICAL_RESCHEDULE',factIds:[],justificationIds:[],calendarIds:[act.id],calendarFenceRevision:0,calendarSnapshots:[{id:act.id,hash:hash(setup(d).ctx.f4ImportedCalendarFields_(act))}],documentaryAdjustmentRef:ref('policy','SYN_ADJUST')});const p=setup(d).plan(i);assert.equal(p.writes.length,2);assert.equal(p.writes[0].data.civilDate,'2026-11-26');assert.equal(p.writes[1].data.id,'SYN_PRESENTATION');assert.equal(p.writes[0].data.legacyStatus,'PLANEJADA');assert.equal(d['migrationEvidence/EA'].literal.STATUS_OPERACIONAL,'PLANEJADA');assert(!p.writes.some(w=>w.path.startsWith('activityAttendance')));
 d['portalReaderResourceFences/activities']={revision:1};assert.throws(()=>setup(d).plan(i),/CALENDAR_CHANGED/);delete d['portalReaderResourceFences/activities'];act.civilDate='2026-09-09';assert.throws(()=>setup(d).plan(i));
});
