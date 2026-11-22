/**
 * ------------------------------------------------------------
 * Extrai ID de asset (Drive File ID) a partir de ID puro ou URL.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Sempre que um asset puder ser informado como:
 *   - ID puro do Drive
 *   - link contendo /d/<id>/
 *   - link contendo ?id=<id>
 *
 * O que aceita:
 * 1) ID puro:
 *    - padrão: [a-zA-Z0-9_-]{20,}
 * 2) Link com /d/<id>/
 * 3) Link com ?id=<id> (ou &id=<id>)
 *
 * Por que existe:
 * - Facilita colar links diretamente do Drive sem “limpar”.
 * - Reduz erro humano na configuração de assets.
 *
 * Comportamento:
 * - Retorna o ID extraído.
 * - Lança erro com mensagem clara se não conseguir extrair.
 */
function coreAssetId_(value) {
  const str = String(value || "").trim();

  // Caso 1: ID puro do Drive
  if (/^[a-zA-Z0-9_-]{20,}$/.test(str)) return str;

  // Caso 2: link no formato /d/<id>/
  let m = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // Caso 3: link com querystring id=<id>
  m = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  throw new Error(
    "Asset inválido. Use ID puro do Drive ou link com /d/<id> ou ?id=<id>."
  );
}


/**
 * ------------------------------------------------------------
 * Retorna o Blob de um arquivo do Drive a partir de ID ou URL.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Para obter blob de imagens (logo/banner) e enviar em inlineImages.
 * - Para anexar arquivos em e-mails.
 *
 * Parâmetro:
 * - assetIdOrUrl: ID puro ou URL do Drive (ver coreAssetId_).
 *
 * Funcionamento:
 * 1) Extrai ID com coreAssetId_.
 * 2) Busca o arquivo no Drive.
 * 3) Retorna o Blob (conteúdo binário).
 *
 * Observação:
 * - Se o arquivo não existir ou não houver permissão, DriveApp lança erro.
 */
function coreGetAssetBlob_(assetIdOrUrl) {
  const id = coreAssetId_(assetIdOrUrl);
  return DriveApp.getFileById(id).getBlob();
}