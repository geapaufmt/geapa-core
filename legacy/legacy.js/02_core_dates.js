/**
 * ------------------------------------------------------------
 * Retorna a data/hora atual.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Sempre que precisar do "agora".
 *
 * Por que existe:
 * - Centraliza a origem do tempo do sistema.
 * - Facilita futura padronização (ex.: mock em testes).
 */
function core_now_() {
  return new Date();
}


/**
 * ------------------------------------------------------------
 * Retorna a data no início do dia (00:00:00.000).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Para comparar datas ignorando horário.
 * - Para validar "é hoje?" sem erro por diferença de hora.
 *
 * Exemplo:
 * - 21/02/2026 18:30 -> 21/02/2026 00:00
 *
 * Observação:
 * - Mantém o mesmo fuso da Date original.
 */
function core_startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}


/**
 * ------------------------------------------------------------
 * Soma (ou subtrai) dias em uma data.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Calcular "ontem", "amanhã", "daqui 4 dias", etc.
 *
 * Parâmetros:
 * - date: data base
 * - days: número de dias (positivo ou negativo)
 *
 * Observação:
 * - Converte days para Number.
 * - Se days for null/undefined, assume 0.
 */
function core_addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + Number(days || 0));
  return d;
}


/**
 * ------------------------------------------------------------
 * Verifica se duas datas caem no mesmo dia.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Checar se uma data é "hoje".
 * - Comparar datas de planilha com data atual.
 *
 * Funcionamento:
 * - Trunca ambas para início do dia.
 * - Compara timestamp (getTime).
 *
 * Por que é melhor assim?
 * - Evita erro por diferença de hora/minuto/segundo.
 */
function core_isSameDay_(d1, d2) {
  return core_startOfDay_(d1).getTime() ===
         core_startOfDay_(d2).getTime();
}


/**
 * ------------------------------------------------------------
 * Verifica se uma data está dentro da janela [start, end).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Trabalhar com intervalos:
 *   - "entre ontem 00:00 e hoje 00:00"
 *   - "entre hoje e amanhã"
 *
 * Regra:
 * - Inclusivo no início (>= start)
 * - Exclusivo no fim (< end)
 *
 * Exemplo:
 *   janela = [20/02 00:00, 21/02 00:00)
 *   -> pega todos os eventos do dia 20.
 */
function core_inWindowDay_(date, windowStart, windowEnd) {
  const t = new Date(date).getTime();
  return t >= new Date(windowStart).getTime() &&
         t <  new Date(windowEnd).getTime();
}


/**
 * ------------------------------------------------------------
 * Formata data em string.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Exibir data em e-mail.
 * - Gravar data formatada em planilha.
 *
 * Parâmetros:
 * - date: Date ou valor compatível
 * - tz: fuso horário (opcional)
 * - pattern: padrão (ex.: 'dd/MM/yyyy')
 *
 * Padrões:
 * - tz padrão: fuso do script
 * - pattern padrão: 'dd/MM/yyyy'
 *
 * Exemplo:
 *   core_formatDate_(new Date(), null, 'dd/MM/yyyy HH:mm')
 */
function core_formatDate_(date, tz, pattern) {
  return Utilities.formatDate(
    new Date(date),
    tz || Session.getScriptTimeZone(),
    pattern || 'dd/MM/yyyy'
  );
}