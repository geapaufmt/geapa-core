/**
 * ------------------------------------------------------------
 * Validação obrigatória de parâmetro.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - No início de qualquer função que exige parâmetro obrigatório.
 * - Para evitar execução com valor undefined/null/vazio.
 *
 * O que considera inválido:
 * - undefined
 * - null
 * - string vazia ('')
 *
 * Comportamento:
 * - Lança erro imediatamente com mensagem clara.
 *
 * Por que falhar cedo?
 * - Evita erro mais obscuro depois.
 * - Facilita debug.
 * - Deixa explícito qual parâmetro faltou.
 *
 * Exemplo:
 *   core_assertRequired_(spreadsheetId, 'Spreadsheet ID');
 *
 * Resultado em erro:
 *   Error: Obrigatório: Spreadsheet ID
 */
function core_assertRequired_(value, label) {
  if (value === undefined || value === null || value === '') {
    throw new Error(`Obrigatório: ${label}`);
  }
}