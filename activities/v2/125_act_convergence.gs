// This combined candidate is local only. Its shared pipeline owns authority and writer fences.
// No remote activation flag, installed scheduler or deployment binding is supplied here.
function f4ActivitiesExecutionGate_(options) {
 if (options.executionGate) {
  // Explicit trusted interface for the offline harness. Never supplied by a request envelope.
  f4ActivitiesModel_().need(typeof options.executionGate === 'function','VALIDATION_ERROR');
  return options.executionGate;
 }
 return function () {
  f4ActivitiesModel_().need(f4Firestore_().isSynthetic(options.db),
   'UNRESOLVED_DEPENDENCY','LOCAL_CANDIDATE_ONLY');
 };
}
