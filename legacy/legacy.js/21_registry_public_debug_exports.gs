/**
 * Exports públicos adicionais do Registry para apoio a módulos,
 * diagnóstico e manutenção.
 *
 * Observação:
 * - Mantidos em arquivo separado para não conflitar com a API pública atual.
 * - Seguem a convenção de wrappers globais exportáveis pela Library.
 */

/**
 * Alias explícito de debug/inspeção.
 * Mantido separado para deixar claro que é uma função de leitura diagnóstica.
 */
function coreGetRegistrySnapshot() {
  return core_getRegistry_();
}
