// Physical historical contract only. No import, resolution or migration command is exposed.
function f4ActivitiesEvidence_(){
 const M=f4ActivitiesModel_(),N=M.need;
 const categories=['IDENTITY_UNRESOLVED','TEMPORAL_POLICY_UNRESOLVED','SOURCE_AUTHORITY_CONFLICT','REALIZATION_SEMANTICS_UNRESOLVED','MATERIAL_SEMANTICS_UNRESOLVED','SESSION_COMPOSITION_UNRESOLVED','THEMATIC_SEMANTICS_UNRESOLVED','INBOX_ASSOCIATION_UNRESOLVED','DUPLICATE_OR_COLLISION','OTHER_EVIDENCE_GAP'];
 function reference(r){N(r&&['RESOLVED','UNRESOLVED','NOT_RECONSTRUCTED','NOT_APPLICABLE'].includes(r.resolution));if(r.resolution==='RESOLVED')M.path(r.ref);else N(r.evidenceRef&&Array.isArray(r.dependencies)&&r.dependencies.length<=32,'UNRESOLVED_DEPENDENCY');return r;}
 function assertion(a){N(a&&M.id(a.id)&&a.sourceId&&a.captureRef&&a.sourceContext&&a.evidenceScope&&a.fieldPath&&a.ruleVersion,'UNRESOLVED_DEPENDENCY','P0_BASIS');N(('literal'in a)!==('partsManifest'in a),'VALIDATION_ERROR','P0_REPRESENTATION');if(a.partsManifest)N(a.partsManifest.state==='SEALED'&&M.rev(a.partsManifest.count)&&a.partsManifest.contentHash,'UNRESOLVED_DEPENDENCY');N(JSON.stringify(a).length<=65536,'SCOPE_TOO_LARGE');return a;}
 function quarantine(q){N(categories.includes(q.category)&&q.assertionRef&&Array.isArray(q.blockedEffect)&&q.blockedEffect.length>0&&q.authorityNeeded&&['OPEN','RESOLVED','REOPENED'].includes(q.status));if(q.status==='RESOLVED')N(q.currentResolutionRef&&q.resolvedAt,'UNRESOLVED_DEPENDENCY');return q;}
 function partial(d){N(d.recordMode==='HISTORICAL_PARTIAL'&&d.provenanceHead&&Array.isArray(d.evidenceScope)&&d.evidenceScope.length>0,'UNRESOLVED_DEPENDENCY');for(const k of ['personRef','cycleRef','policyRef'])if(d[k]?.resolution)reference(d[k]);N(!d.legacyStatus||!d.realizationRecognizedByLegacyLabel,'INVALID_STATE_TRANSITION');return d;}
 function admitMigration(){M.fail('AUTHORIZATION_DENIED','OP11_MIGRATION_NOT_ENABLED');}
 return {categories,reference,assertion,quarantine,partial,admitMigration};
}
