/**
 * ------------------------------------------------------------
 * Retorna uma label do Gmail pelo nome.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Verificar se uma label já existe.
 * - Associar threads a uma label específica.
 *
 * Retorna:
 * - GmailLabel se existir
 * - null se não existir
 *
 * Observação:
 * - Não cria label automaticamente.
 * - Apenas consulta.
 */
function core_getLabel_(name) {
  if (!name) return null;
  return GmailApp.getUserLabelByName(name);
}


/**
 * ------------------------------------------------------------
 * Garante que uma label exista.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Antes de aplicar uma label em threads.
 * - Durante inicialização de módulo (install/setup).
 *
 * Comportamento:
 * - Se já existir, não faz nada.
 * - Se não existir, cria a label.
 *
 * Observação:
 * - Não retorna a label.
 * - Apenas garante a existência.
 *
 * Por que separar de getLabel?
 * - Dá mais controle:
 *   - getLabel -> leitura
 *   - ensureLabel -> criação preventiva
 */
function core_ensureLabel_(name) {
  if (!name) return;

  const l = GmailApp.getUserLabelByName(name);

  if (!l) {
    GmailApp.createLabel(name);
  }
}