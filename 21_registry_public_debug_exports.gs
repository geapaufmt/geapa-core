/**
 * Exports públicos adicionais do Registry para apoio a módulos,
 * diagnóstico e manutenção.
 *
 * Observação:
 * - Mantidos em arquivo separado para não conflitar com a API pública atual.
 * - Seguem a convenção de wrappers globais exportáveis pela Library.
 */

/** Retorna o ambiente atual do Core (DEV ou PROD). */
function coreGetCurrentEnv() {
  return core_getCurrentEnv_();
}

/**
 * Retorna os metadados completos de uma KEY do Registry.
 * Útil para debug, auditoria e validação estrutural.
 */
function coreGetRegistryMetaByKey(key) {
  return core_getRegistryMetaByKey_(key);
}

/** Limpa o cache do Registry. Útil em manutenção e debug. */
function coreClearRegistryCache() {
  return core_registryCacheClear_();
}

/**
 * Alias explícito de debug/inspeção.
 * Mantido separado para deixar claro que é uma função de leitura diagnóstica.
 */
function coreGetRegistrySnapshot() {
  return core_getRegistry_();
}
