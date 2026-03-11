/**************************************
 * 08_core_drive.gs — helpers de Drive
 **************************************/


/**
 * ------------------------------------------------------------
 * Retorna uma pasta do Drive pelo ID.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Acessar pasta institucional (ex.: pasta de apresentações).
 *
 * Comportamento:
 * - Valida folderId obrigatório.
 * - Lança erro se folderId inválido.
 *
 * Retorna:
 * - Objeto Folder (DriveApp)
 */
function core_driveGetFolderById_(folderId) {
  core_assertRequired_(folderId, 'Drive folderId');
  return DriveApp.getFolderById(folderId);
}


/**
 * ------------------------------------------------------------
 * Retorna um arquivo do Drive pelo ID.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Processar anexos recebidos.
 * - Mover arquivos para pasta institucional.
 *
 * Retorna:
 * - Objeto File (DriveApp)
 */
function core_driveGetFileById_(fileId) {
  core_assertRequired_(fileId, 'Drive fileId');
  return DriveApp.getFileById(fileId);
}


/**
 * ------------------------------------------------------------
 * Garante existência de subpasta dentro de uma pasta pai.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Organizar arquivos por semestre.
 * - Organizar por data.
 * - Criar estrutura hierárquica automática.
 *
 * Funcionamento:
 * - Procura subpasta com mesmo nome.
 * - Se existir, retorna a primeira encontrada.
 * - Se não existir, cria nova pasta.
 *
 * Observação:
 * - Não trata múltiplas pastas com mesmo nome.
 * - Mantém comportamento simples e previsível.
 */
function core_driveEnsureFolder_(parentFolderId, childFolderName) {
  core_assertRequired_(parentFolderId, 'parentFolderId');
  core_assertRequired_(childFolderName, 'childFolderName');

  const parent = DriveApp.getFolderById(parentFolderId);
  const it = parent.getFoldersByName(childFolderName);

  if (it.hasNext()) {
    return it.next();
  }

  return parent.createFolder(childFolderName);
}


/**
 * ------------------------------------------------------------
 * Lista arquivos de uma pasta (metadados básicos).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Auditoria de pasta.
 * - Debug.
 * - Verificação de arquivos recebidos.
 *
 * Parâmetros:
 * - folderId: pasta alvo
 * - max: limite máximo de arquivos (padrão: 200)
 *
 * Retorna:
 * - Array de objetos:
 *   [{
 *     id,
 *     name,
 *     mimeType,
 *     url,
 *     lastUpdated
 *   }]
 *
 * Observação:
 * - Não retorna conteúdo do arquivo.
 * - Limita por segurança/performance.
 */
function core_driveListFiles_(folderId, max = 200) {
  core_assertRequired_(folderId, 'folderId');

  const folder = DriveApp.getFolderById(folderId);
  const out = [];
  const files = folder.getFiles();

  while (files.hasNext() && out.length < max) {
    const f = files.next();

    out.push({
      id: f.getId(),
      name: f.getName(),
      mimeType: f.getMimeType(),
      url: f.getUrl(),
      lastUpdated: f.getLastUpdated(),
    });
  }

  return out;
}


/**
 * ------------------------------------------------------------
 * Move arquivo para pasta destino.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Organizar anexos recebidos.
 * - Transferir arquivos para pasta institucional.
 *
 * Observação importante:
 * - Apps Script NÃO possui método "move" direto.
 * - Implementação faz:
 *     1) addFile() no destino
 *     2) removeFile() de todas as pastas anteriores
 *
 * Segurança:
 * - Não remove da pasta destino.
 *
 * Retorna:
 * - Objeto simples com:
 *   { id, url, name }
 */
function core_driveMoveFileToFolder_(fileId, targetFolderId) {
  core_assertRequired_(fileId, 'fileId');
  core_assertRequired_(targetFolderId, 'targetFolderId');

  const file = DriveApp.getFileById(fileId);
  const target = DriveApp.getFolderById(targetFolderId);

  // Adiciona no destino
  target.addFile(file);

  // Remove de todas as pastas anteriores (exceto o destino)
  const parents = file.getParents();

  while (parents.hasNext()) {
    const p = parents.next();

    if (p.getId() !== target.getId()) {
      p.removeFile(file);
    }
  }

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
  };
}