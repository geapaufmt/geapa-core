/**
 * ------------------------------------------------------------
 * CORE Health Check (verificação básica do ambiente).
 * ------------------------------------------------------------
 *
 * Objetivo:
 * - Validar rapidamente se o Core está funcionando.
 * - Confirmar permissões básicas do projeto.
 * - Verificar ambiente antes de integrar módulos.
 *
 * Quando usar:
 * - Após publicar nova versão da Library.
 * - Após alterar permissões.
 * - Para diagnosticar problemas iniciais.
 *
 * O que verifica:
 * - Geração de runId.
 * - Sistema de logs.
 * - Fuso horário do projeto.
 * - Cota diária restante de envio de e-mail.
 *
 * Observação:
 * - Não depende de planilhas.
 * - Não depende de IDs institucionais.
 * - Seguro para executar a qualquer momento.
 *
 * Exemplo de uso:
 *   Executar manualmente no editor do Core.
 */
function core_healthCheck() {
  const runId = core_runId_();
  const startedAt = new Date();

  core_logInfo_(runId, 'CORE healthCheck: INÍCIO');

  // Informações básicas do ambiente
  core_logInfo_(runId, 'Timezone do projeto', Session.getScriptTimeZone());
  core_logInfo_(runId, 'Mail quota', MailApp.getRemainingDailyQuota());

  core_logSummarize_(runId, 'CORE healthCheck', startedAt, {
    ok: true
  });
}