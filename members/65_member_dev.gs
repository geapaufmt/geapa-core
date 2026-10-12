// DEV homologation only. In-memory capability, never a caller-supplied path/config flag.
// Default repository entrypoints stay fenced; no production or top-level data route.
function f4MemberDev_() {
 if (f4MemberDev_.value_) return f4MemberDev_.value_;
 const sessions = new WeakMap(), C = f4Core_();
 function physical() {
  C.need(ScriptApp.getScriptId() === '1M3GpuWwyRkV24nbHnopzj1q4J-Ea7zHv5eCG-ngOw0qR5lWDMvvv8UfD', 'DESTINATION_INVALID');
  C.need(PropertiesService.getScriptProperties().getProperty('GEAPA_ENV') === 'DEV', 'DESTINATION_INVALID');
 }
 function namespace(config) {
  const run = sessions.get(config);
  if (!run) return null;
  C.destination(config); physical();
  C.need(config.environment === 'DEV', 'DESTINATION_INVALID');
  return 'memberDevHomologation/' + run;
 }
 function requireSession(config) {
  if (!namespace(config)) { const e = new Error('M1_LOCAL_ONLY'); e.code = e.message; throw e; }
 }
 function open(runId) {
  physical(); C.need(/^M3_[a-f0-9]{24}$/.test(runId), 'DESTINATION_INVALID');
  const config = f4aTarget_(); sessions.set(config, runId); return config;
 }
 return (f4MemberDev_.value_ = Object.freeze({open, namespace, requireSession, close: config => sessions.delete(config)}));
}
