/**
 * ------------------------------------------------------------
 * Normalizacao simples de e-mail.
 * ------------------------------------------------------------
 */
function core_normalizeEmail_(value) {
  return core_normalizeText_(value, {
    collapseWhitespace: true,
    caseMode: 'lower'
  });
}

function core_extractEmailAddress_(value) {
  var text = String(value || '').trim();
  if (!text) return '';

  var match = text.match(/<([^>]+)>/);
  return core_normalizeEmail_(match ? match[1] : text);
}

function core_extractDisplayName_(value) {
  var text = String(value || '').trim();
  if (!text) return '';

  var match = text.match(/^"?([^"<]+)"?\s*</);
  var name = String(match ? match[1] : text).trim();
  return name.indexOf('@') !== -1 ? '' : name;
}

function core_uniqueEmails_(values) {
  var input = Array.isArray(values)
    ? values
    : String(values || '').split(/[;,]/);

  var seen = Object.create(null);
  var out = [];

  input.forEach(function(item) {
    var email = core_extractEmailAddress_(item);
    if (!email || seen[email]) return;

    seen[email] = true;
    out.push(email);
  });

  return Object.freeze(out);
}

function core_resolveEmailRecipients_(value) {
  var list = core_uniqueEmails_(value);
  list.forEach(function(email) {
    if (!core_isValidEmail_(email)) {
      throw new Error('E-mail invalido: ' + email);
    }
  });
  return list;
}

/**
 * ------------------------------------------------------------
 * Validacao simples de e-mail.
 * ------------------------------------------------------------
 */
function core_isValidEmail_(email) {
  var e = core_extractEmailAddress_(email);
  return !!e && e.includes('@') && !/\s/.test(e);
}

/**
 * ------------------------------------------------------------
 * Envia e-mail em TEXTO simples.
 * ------------------------------------------------------------
 */
function core_sendEmailText_(opts) {
  core_assertRequired_(opts, 'Email opts');

  var toList = core_resolveEmailRecipients_(opts.to);
  if (!toList.length) {
    throw new Error('core_sendEmailText_: destinatario ausente.');
  }

  var ccList = opts.cc ? core_resolveEmailRecipients_(opts.cc) : [];
  var bccList = opts.bcc ? core_resolveEmailRecipients_(opts.bcc) : [];

  MailApp.sendEmail({
    to: toList.join(','),
    cc: ccList.join(','),
    bcc: bccList.join(','),
    subject: opts.subject || '',
    body: opts.body || '',
    name: opts.name || 'GEAPA',
    attachments: opts.attachments || undefined,
    replyTo: opts.replyTo || undefined,
    noReply: opts.noReply === true
  });
}

/**
 * ------------------------------------------------------------
 * Envia e-mail em HTML (com suporte a inlineImages).
 * ------------------------------------------------------------
 */
function core_sendHtmlEmail_(opts) {
  core_assertRequired_(opts, 'Email opts');

  var toList = core_resolveEmailRecipients_(opts.to);
  if (!toList.length) {
    throw new Error('core_sendHtmlEmail_: destinatario ausente.');
  }

  var ccList = opts.cc ? core_resolveEmailRecipients_(opts.cc) : [];
  var bccList = opts.bcc ? core_resolveEmailRecipients_(opts.bcc) : [];

  MailApp.sendEmail({
    to: toList.join(','),
    cc: ccList.join(','),
    bcc: bccList.join(','),
    subject: opts.subject || '',
    body: opts.body || 'Mensagem em HTML',
    htmlBody: opts.htmlBody || '',
    attachments: opts.attachments || undefined,
    inlineImages: opts.inlineImages || undefined,
    replyTo: opts.replyTo || undefined,
    noReply: opts.noReply === true,
    name: opts.name || 'GEAPA'
  });
}

function core_sendTrackedEmail_(params) {
  core_assertRequired_(params, 'Email params');

  var toList = core_resolveEmailRecipients_(params.to);
  if (toList.length !== 1) {
    throw new Error('core_sendTrackedEmail_: informe exatamente um destinatario em "to".');
  }

  var to = toList[0];
  var subject = String(params.subject || '').trim();
  var htmlBody = String(params.htmlBody || '');
  var body = String(params.body || '');
  var newerThanDays = Number(params.newerThanDays || 7);
  var maxThreads = Number(params.maxThreads || 10);
  var sleepMs = Number(params.sleepMs || 1500);

  if (!subject) {
    throw new Error('core_sendTrackedEmail_: parametro obrigatorio ausente (subject).');
  }

  if (htmlBody) {
    core_sendHtmlEmail_(Object.assign({}, params, {
      to: to,
      subject: subject,
      body: body || 'Mensagem em HTML',
      htmlBody: htmlBody
    }));
  } else {
    core_sendEmailText_(Object.assign({}, params, {
      to: to,
      subject: subject,
      body: body
    }));
  }

  if (sleepMs > 0) {
    Utilities.sleep(sleepMs);
  }

  var safeSubject = subject.replace(/"/g, '\\"');
  var query = 'to:' + to + ' subject:"' + safeSubject + '" newer_than:' + newerThanDays + 'd';
  var threads = GmailApp.search(query, 0, maxThreads);

  var threadId = '';
  var messageId = '';

  if (threads && threads.length) {
    var thread = threads[0];
    threadId = String(thread.getId() || '');

    var messages = thread.getMessages();
    if (messages && messages.length) {
      messageId = String(messages[messages.length - 1].getId() || '');
    }
  }

  return Object.freeze({
    threadId: threadId,
    messageId: messageId
  });
}

/**
 * ------------------------------------------------------------
 * Retorna inlineImages padrao do GEAPA (ex.: logo).
 * ------------------------------------------------------------
 */
function core_inlineImagesDefault_() {
  var logoId =
    (typeof GEAPA_ASSETS !== 'undefined' &&
     GEAPA_ASSETS.BRAND &&
     GEAPA_ASSETS.BRAND.LOGO_GEAPA)
      ? GEAPA_ASSETS.BRAND.LOGO_GEAPA
      : '';

  if (!logoId) return {};

  try {
    return {
      geapa_logo: coreGetAssetBlob(logoId)
    };
  } catch (err) {
    var message = err && err.message ? err.message : String(err || '');
    if (message.indexOf('Access denied') !== -1 || message.indexOf('DriveApp') !== -1) {
      return {};
    }
    throw err;
  }
}

/**
 * ------------------------------------------------------------
 * Envia HTML mesclando inlineImages padrao + inlineImages do modulo.
 * ------------------------------------------------------------
 */
function core_sendEmailHtmlWithDefaultInline_(opts) {
  core_assertRequired_(opts, 'Email opts');
  var baseInline = core_inlineImagesDefault_();
  var extraInline = opts.inlineImages || {};
  var mergedInline = Object.assign({}, baseInline, extraInline);

  return core_sendHtmlEmail_(Object.assign({}, opts, {
    inlineImages: Object.keys(mergedInline).length ? mergedInline : undefined
  }));
}

/**
 * Alias de compatibilidade.
 */
function core_sendEmailHtml_(opts) {
  return core_sendHtmlEmail_(opts);
}

/**
 * Alias de compatibilidade.
 */
function coreSendHtmlEmail_(opts) {
  return core_sendHtmlEmail_(opts);
}
