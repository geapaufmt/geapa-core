/**
 * ------------------------------------------------------------
 * Validação simples de e-mail.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Antes de enviar qualquer e-mail.
 *
 * O que valida:
 * - Não vazio
 * - Contém "@"
 * - Não possui espaços
 *
 * Observação:
 * - Não é validação RFC completa.
 * - É uma validação leve para evitar erros óbvios.
 *
 * Por que é simples?
 * - Evita regex complexa desnecessária.
 * - O MailApp já falhará em casos extremos.
 */
function core_isValidEmail_(email) {
  const e = String(email || '').trim();
  return !!e && e.includes('@') && !/\s/.test(e);
}


/**
 * ------------------------------------------------------------
 * Envia e-mail em TEXTO simples.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Mensagens administrativas simples.
 * - Logs ou notificações internas.
 *
 * Parâmetros (opts):
 * {
 *   to: obrigatório,
 *   subject: opcional,
 *   body: opcional,
 *   bcc: opcional,
 *   name: nome do remetente (padrão: "GEAPA")
 * }
 *
 * Comportamento:
 * - Valida se opts existe.
 * - Valida se e-mail é válido.
 * - Lança erro se inválido (falha explícita).
 *
 * Por que falhar?
 * - Melhor interromper o job do que enviar para endereço errado.
 */
function core_sendEmailText_(opts) {
  core_assertRequired_(opts, 'Email opts');

  if (!core_isValidEmail_(opts.to)) {
    throw new Error('E-mail inválido: ' + opts.to);
  }

  MailApp.sendEmail({
    to: opts.to,
    bcc: opts.bcc || '',
    subject: opts.subject || '',
    body: opts.body || '',
    name: opts.name || 'GEAPA',
  });
}


/**
 * ------------------------------------------------------------
 * Envia e-mail em HTML (com suporte a inlineImages).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - E-mails institucionais.
 * - Mensagens com logo do GEAPA.
 * - Templates estilizados.
 *
 * Parâmetros (opts):
 * {
 *   to: obrigatório,
 *   subject: opcional,
 *   htmlBody: conteúdo HTML,
 *   body: fallback texto simples,
 *   bcc: opcional,
 *   name: nome do remetente,
 *   inlineImages: objeto { cidKey: Blob }
 * }
 *
 * Comportamento:
 * - Valida parâmetros obrigatórios.
 * - Permite fallback de texto (caso cliente não renderize HTML).
 * - inlineImages é opcional.
 *
 * Observação importante:
 * - inlineImages deve casar com <img src="cid:chave">
 * - Se não informado, não envia imagens.
 */
function core_sendHtmlEmail_(opts) {
  core_assertRequired_(opts, 'Email opts');

  if (!core_isValidEmail_(opts.to)) {
    throw new Error('E-mail inválido: ' + opts.to);
  }

  MailApp.sendEmail({
    to: opts.to,
    bcc: opts.bcc || '',
    subject: opts.subject || '',
    body: opts.body || 'Mensagem em HTML',
    htmlBody: opts.htmlBody || '',
    inlineImages: opts.inlineImages || undefined,
    name: opts.name || 'GEAPA',
  });
}

/**
 * ------------------------------------------------------------
 * Retorna inlineImages padrão do GEAPA (ex.: logo).
 * ------------------------------------------------------------
 *
 * Por que existe:
 * - Muitos módulos querem sempre a logo no topo sem repetir código.
 *
 * Retorna:
 * - {} se não houver logo configurada
 * - ou { geapa_logo: Blob } se houver
 *
 * Observação:
 * - O template deve usar: <img src="cid:geapa_logo">
 */
function core_inlineImagesDefault_() {
  // Se você tiver um registry de assets no core (GEAPA_ASSETS), usa ele:
  const logoId =
    (typeof GEAPA_ASSETS !== 'undefined' &&
     GEAPA_ASSETS.BRAND &&
     GEAPA_ASSETS.BRAND.LOGO_GEAPA)
      ? GEAPA_ASSETS.BRAND.LOGO_GEAPA
      : '';

  if (!logoId) return {};

  // coreGetAssetBlob / core_getAssetBlob_ depende de como você nomeou.
  // Vou usar coreGetAssetBlob() porque você já tem essa função hoje.
  return {
    geapa_logo: coreGetAssetBlob(logoId),
  };
}

/**
 * ------------------------------------------------------------
 * Envia HTML mesclando inlineImages padrão + inlineImages do módulo.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Em módulos que sempre devem ter logo (ou outros assets padrão).
 *
 * Como funciona:
 * - inlineImagesFinal = defaultInline + opts.inlineImages
 * - Se houver conflito de CID, o do módulo (opts.inlineImages) vence.
 *
 * Reaproveita:
 * - core_sendEmailHtml_ (função já existente e testada)
 */
function core_sendEmailHtmlWithDefaultInline_(opts) {
  core_assertRequired_(opts, 'Email opts');
  const baseInline = core_inlineImagesDefault_();
  const extraInline = opts.inlineImages || {};

  // Merge: extra sobrescreve base se repetir a chave/cid
  const mergedInline = Object.assign({}, baseInline, extraInline);

  return core_sendHtmlEmail_(Object.assign({}, opts, {
    inlineImages: Object.keys(mergedInline).length ? mergedInline : undefined,
  }));
}

/**
 * Alias de compatibilidade:
 * Alguns módulos chamam core_sendEmailHtml_ (nome antigo).
 * O nome "oficial" atual é core_sendHtmlEmail_.
 */
function core_sendEmailHtml_(opts) {
  return core_sendHtmlEmail_(opts);
}

/**
 * Alias de compatibilidade:
 * Alguns módulos chamam coreSendHtmlEmail_ (nome usado em assets/email).
 * O nome "oficial" atual é core_sendHtmlEmail_.
 */
function coreSendHtmlEmail_(opts) {
  return core_sendHtmlEmail_(opts);
}