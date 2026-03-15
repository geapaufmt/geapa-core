/***************************************
 * 90_core_triggers.gs
 *
 * Instala triggers do core.
 ***************************************/

function core_installTriggers() {
  const all = ScriptApp.getProjectTriggers();

  all.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "coreSyncMembersCurrentDerivedFields") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("coreSyncMembersCurrentDerivedFields")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log("Trigger instalado: coreSyncMembersCurrentDerivedFields (diário às 6h).");
}

function core_uninstallTriggers() {
  const all = ScriptApp.getProjectTriggers();

  all.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "coreSyncMembersCurrentDerivedFields") {
      ScriptApp.deleteTrigger(t);
    }
  });

  Logger.log("Triggers do core removidos.");
}