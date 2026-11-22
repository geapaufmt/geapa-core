/***************************************
 * 90_core_triggers.gs
 *
 * Instalacao e auditoria de triggers do core.
 ***************************************/

function core_getManagedTriggerSpecs_() {
  return Object.freeze([
    Object.freeze({
      handler: 'coreSyncMembersCurrentDerivedFields',
      label: 'Sincronizacao de derivados em MEMBERS_ATUAIS',
      scheduleType: 'DAILY',
      expectedSummary: 'Todo dia as 6h'
    }),
    Object.freeze({
      handler: 'coreMailProcessOutbox',
      label: 'Processador da MAIL_SAIDA',
      scheduleType: 'HOURLY',
      expectedSummary: 'A cada 1 hora'
    }),
    Object.freeze({
      handler: 'coreMailIngestInbox',
      label: 'Ingestao automatica da caixa de entrada',
      scheduleType: 'HOURLY',
      expectedSummary: 'A cada 1 hora'
    }),
    Object.freeze({
      handler: 'coreMailCleanupNoiseEvents',
      label: 'Higiene de eventos ignorados do Mail Hub',
      scheduleType: 'DAILY',
      expectedSummary: 'Todo dia as 5h'
    })
  ]);
}

function core_getManagedTriggerHandlers_() {
  return core_getManagedTriggerSpecs_().map(function(spec) {
    return spec.handler;
  });
}

function core_listManagedTriggers_() {
  var handlers = core_getManagedTriggerHandlers_();
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    return handlers.indexOf(trigger.getHandlerFunction()) >= 0;
  });
}

function core_deleteManagedTriggers_() {
  var triggers = core_listManagedTriggers_();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  return triggers.length;
}

function core_createManagedTriggers_() {
  ScriptApp.newTrigger('coreSyncMembersCurrentDerivedFields')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  ScriptApp.newTrigger('coreMailProcessOutbox')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('coreMailIngestInbox')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('coreMailCleanupNoiseEvents')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();
}

function core_installTriggers() {
  var deleted = core_deleteManagedTriggers_();
  core_createManagedTriggers_();
  return Object.freeze({
    ok: true,
    deleted: deleted,
    created: core_getManagedTriggerSpecs_().length,
    managedHandlers: core_getManagedTriggerHandlers_()
  });
}

function core_reinstallTriggers() {
  return core_installTriggers();
}

function core_uninstallTriggers() {
  var deleted = core_deleteManagedTriggers_();
  return Object.freeze({
    ok: true,
    deleted: deleted,
    managedHandlers: core_getManagedTriggerHandlers_()
  });
}

function core_listTriggers() {
  var specs = core_getManagedTriggerSpecs_();
  var triggers = core_listManagedTriggers_();
  var byHandler = {};

  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (!byHandler[handler]) byHandler[handler] = [];
    byHandler[handler].push(trigger);
  });

  return Object.freeze({
    ok: true,
    note: 'A API do Apps Script permite auditar presenca e duplicidade dos handlers, mas nao expoe todos os detalhes finos da agenda configurada.',
    managed: specs.map(function(spec) {
      var installed = byHandler[spec.handler] || [];
      return Object.freeze({
        handler: spec.handler,
        label: spec.label,
        scheduleType: spec.scheduleType,
        expectedSummary: spec.expectedSummary,
        installedCount: installed.length,
        triggerSource: installed.length ? String(installed[0].getTriggerSource()) : '',
        eventType: installed.length ? String(installed[0].getEventType()) : '',
        uniqueIds: installed.map(function(item) { return item.getUniqueId(); })
      });
    })
  });
}

function core_validateTriggers() {
  var listing = core_listTriggers();
  var missing = [];
  var duplicates = [];

  listing.managed.forEach(function(item) {
    if (item.installedCount === 0) missing.push(item.handler);
    if (item.installedCount > 1) duplicates.push(item.handler);
  });

  return Object.freeze({
    ok: missing.length === 0 && duplicates.length === 0,
    missingHandlers: missing,
    duplicateHandlers: duplicates,
    managed: listing.managed,
    note: listing.note
  });
}
