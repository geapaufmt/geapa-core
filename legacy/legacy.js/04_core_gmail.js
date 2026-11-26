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

function core_getOrCreateLabel_(name) {
  if (!name) return null;
  core_ensureLabel_(name);
  return core_getLabel_(name);
}

function core_threadHasLabel_(thread, labelName) {
  core_assertRequired_(thread, 'thread');
  var target = String(labelName || '').trim();
  if (!target) return false;

  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    if (String(labels[i].getName() || '').trim() === target) {
      return true;
    }
  }

  return false;
}

function core_searchThreads_(query, start, max) {
  core_assertRequired_(query, 'query');
  return GmailApp.search(
    String(query),
    Number(start || 0),
    Number(max || 50)
  );
}

function core_markThread_(thread, labelIn, labelOut) {
  core_assertRequired_(thread, 'thread');

  var removeLabel = typeof labelIn === 'string' ? core_getLabel_(labelIn) : labelIn;
  var addLabel = typeof labelOut === 'string' ? core_getOrCreateLabel_(labelOut) : labelOut;

  if (removeLabel) thread.removeLabel(removeLabel);
  if (addLabel) thread.addLabel(addLabel);
  thread.markRead();
}

function core_replyThreadHtml_(thread, subject, htmlBody, opts) {
  core_assertRequired_(thread, 'thread');
  opts = opts || {};

  thread.reply(opts.body || '', {
    subject: subject || '',
    htmlBody: htmlBody || '',
    noReply: opts.noReply === true
  });
}
