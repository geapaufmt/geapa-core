/**
 * ------------------------------------------------------------
 * Registry de assets oficiais do GEAPA (imagens/arquivos do Drive).
 * ------------------------------------------------------------
 *
 * Objetivo:
 * - Centralizar IDs (ou links) dos arquivos visuais oficiais do grupo,
 *   como logo, banners e assinaturas, para uso em vários módulos.
 *
 * Por que existe:
 * - Evita duplicar IDs em múltiplos scripts/módulos.
 * - Facilita trocar uma imagem (basta alterar aqui).
 * - Mantém padronização visual dos e-mails institucionais.
 *
 * Regras:
 * - Preferir ID puro do Drive quando possível.
 * - Pode aceitar link do Drive, desde que seja extraível (ver coreAssetId_).
 *
 * Estrutura sugerida:
 * - BRAND: elementos de marca (logo etc.)
 * - EMAIL: elementos específicos de e-mail (banner/assinatura etc.)
 *
 * Observação:
 * - Este objeto é infraestrutura e deve ser imutável (Object.freeze).
 * - Módulos NÃO devem acessar GEAPA_ASSETS diretamente via Library;
 *   em vez disso, devem usar funções (ex.: coreGetAssetBlob).
 */
const GEAPA_ASSETS = Object.freeze({
  BRAND: {
    // Logo oficial do GEAPA (Drive File ID)
    LOGO_GEAPA: "1Md2YlFNXo4qD_5D1TwD7Ej3QqqvgitPV",
  },
  EMAIL: {
    // Assets opcionais (caso existam futuramente):
    // BANNER_PADRAO: "ID_OU_LINK",
    // ASSINATURA_DIRETORIA: "ID_OU_LINK",
  }
});