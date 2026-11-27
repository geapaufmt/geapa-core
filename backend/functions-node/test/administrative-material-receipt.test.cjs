const {test}=require('node:test'),a=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const root=path.resolve(__dirname,'../../..');
function scenario(institutional=true,late=false){
 const ref=(kind,id)=>({kind,id}),now=100000,deadline=90000,received=late?95000:80000,docs=new Map();
 const M={path:r=>r.kind+'/'+r.id,stamp:n=>n,ms:n=>n,ref},put=(r,d)=>docs.set(M.path(r),structuredClone(d));
 const pr=ref('presentation','P'),rq=ref('requirement','M'),sub=ref('submission','S'),review=ref('review','R'),coverage=ref('policy','C');
 put(pr,{personRef:ref('person','SYNTHETIC'),materialRequirementRef:rq});
 put(rq,{currentRevisionRef:{kind:'requirementRevision',id:'M',number:1},currentSubmissionRef:sub,currentReviewRef:review,finalEndExclusive:deadline,materialStateRevision:1});
 put(sub,{presentationRef:pr,submittedAt:now,receivedBySystemAt:now,...(institutional?{administrativeRegistration:{institutionalEvidence:{receivedAt:received,attestedAt:now,reference:'SYNTHETIC_PROTOCOL'}}}:{})});
 put(review,{outcome:'VALID',submissionRef:sub,coveredItemsManifestRef:coverage});
 put(coverage,{digest:'SYNTHETIC',requirementRevisionRef:{id:'M',number:1},completeCoverage:true});
 const x={M,N:(ok,code,reasonCode)=>{if(!ok)throw Object.assign(Error(reasonCode||code),{code,reasonCode});},p:{submissionRef:sub,validReviewRef:review,requirementRevisionRef:{id:'M',number:1},coverageManifestRef:coverage},now,cycleRef:ref('cycle','SYNTHETIC'),load:r=>structuredClone(docs.get(M.path(r))),save:put,revision:put,manifest:r=>docs.get(M.path(r)),event:(kind,d)=>{const r=ref(kind,'NEW');put(r,d);return r;}};
 const c={};vm.createContext(c);for(const file of ['111_act_domain.gs','114_act_materials.gs'])vm.runInContext(fs.readFileSync(path.join(root,'activities/v2',file),'utf8'),c);
 return {run:()=>c.f4ActivitiesMaterials_(x).MATERIAL_RECEIVE(),docs,now,received};
}
test('valid institutional receipt uses attested date without backdating Portal submission or completing annual obligation',()=>{
 const s=scenario(),r=s.run(),receipt=s.docs.get('receipt/NEW');
 a.equal(r.deadlineComplianceBasis.kind,'ATTESTED_INSTITUTIONAL_RECEIPT');a.equal(receipt.deadlineComplianceBasis.institutionalReceivedAt,s.received);
 a.equal(receipt.deadlineComplianceBasis.registeredAt,s.now);a.equal(receipt.receivedAt,s.now);a.equal(s.docs.get('submission/S').submittedAt,s.now);
 a(![...s.docs.keys()].some(k=>/^(obligation|realization|regularization)\//.test(k)));
});
test('late institutional receipt and ordinary late upload still require their canonical exception',()=>{
 for(const s of [scenario(true,true),scenario(false)])a.throws(s.run,{reasonCode:'LATE_SUBMISSION'});
});
