/**
 * ------------------------------------------------------------
 * Gera identificador curto de execução (runId).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - No início de qualquer job automático.
 * - Para correlacionar logs da mesma execução.
 *
 * Funcionamento:
 * - Gera UUID.
 * - Retorna apenas os primeiros 8 caracteres.
 *
 * Por que curto?
 * - Facilita leitura visual no Logger.
 * - Já é suficientemente único para execuções concorrentes.
 *
 * Exemplo:
 *   const runId = core_runId_();
 */
function core_runId_() {
  return Utilities.getUuid().slice(0, 8);
}


/**
 * ------------------------------------------------------------
 * Log nível INFO.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Passos normais da execução.
 * - Início/fim de rotina.
 * - Eventos esperados.
 *
 * Padrão:
 *   [INFO] [runId] mensagem
 */
function core_logInfo_(runId, msg, obj) {
  Logger.log(core_logLine_('INFO', runId, msg, obj));
}


/**
 * ------------------------------------------------------------
 * Log nível WARN.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Situação inesperada, mas não fatal.
 * - Dados ausentes opcionais.
 */
function core_logWarn_(runId, msg, obj) {
  Logger.log(core_logLine_('WARN', runId, msg, obj));
}


/**
 * ------------------------------------------------------------
 * Log nível ERROR.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Erros críticos.
 * - Falhas de validação.
 * - Exceções capturadas.
 */
function core_logError_(runId, msg, obj) {
  Logger.log(core_logLine_('ERROR', runId, msg, obj));
}


/**
 * ------------------------------------------------------------
 * Formata linha de log padronizada.
 * ------------------------------------------------------------
 *
 * Estrutura:
 *   [LEVEL] [runId] mensagem | {json opcional}
 *
 * Se obj for fornecido:
 * - Serializa como JSON.
 * - Facilita debug estruturado.
 *
 * Exemplo:
 *   core_logInfo_(id, "Processando", { linha: 3 });
 */
function core_logLine_(level, runId, msg, obj) {
  const base = `[${level}] [${runId}] ${msg}`;

  if (obj === undefined) {
    return base;
  }

  return `${base} | ${JSON.stringify(obj)}`;
}


/**
 * ------------------------------------------------------------
 * Log de encerramento padrão (resumo de execução).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - No final de qualquer job automático.
 *
 * Parâmetros:
 * - runId: identificador da execução
 * - title: nome do job
 * - startedAt: data/hora inicial
 * - counters: objeto com contadores (enviados, pulos, erros...)
 *
 * Funcionamento:
 * - Calcula duração em milissegundos.
 * - Loga como INFO.
 *
 * Exemplo:
 *   core_logSummarize_(runId, "ANIVERSARIOS", inicio, {
 *     enviados: 3,
 *     pulos: 2,
 *     erros: 0
 *   });
 */
function core_logSummarize_(runId, title, startedAt, counters) {
  const dur = startedAt
    ? (new Date().getTime() - new Date(startedAt).getTime())
    : null;

  core_logInfo_(runId, `${title} | FIM`, {
    durMs: dur,
    ...counters
  });
}