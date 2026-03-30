/**
 * ============================================================
 * 19_core_mail_renderer.js
 * ============================================================
 *
 * Renderer institucional de e-mails do GEAPA.
 *
 * Objetivo:
 * - manter o layout visual centralizado no geapa-core;
 * - deixar os modulos donos apenas do conteudo de negocio;
 * - preparar drafts compativeis com a futura fila de saida.
 */

var CORE_MAIL_RENDERER_BRAND = Object.freeze({
  orgName: 'Grupo de Estudos de Apoio a Producao Agricola (GEAPA)',
  shortName: 'GEAPA',
  quoteFallback: '"Cultivar o Conhecimento Para Colher Sabedoria"',
  background: '#f5f8f5',
  cardBackground: '#ffffff',
  textColor: '#163226',
  mutedTextColor: '#587062',
  borderColor: '#167d0a',
  dividerColor: '#d6e5d3',
  buttonTextColor: '#ffffff',
  shadowColor: 'rgba(22,125,10,0.08)'
});

var CORE_MAIL_RENDERER_TEMPLATES = Object.freeze({
  GEAPA_COMEMORATIVO: Object.freeze({
    templateKey: 'GEAPA_COMEMORATIVO',
    displayName: 'GEAPA Comemorativo',
    eyebrow: 'Mensagem Comemorativa',
    accent: '#167d0a',
    accentSoft: '#ecf8e8',
    accentSurface: '#f4fbf1'
  }),
  GEAPA_OPERACIONAL: Object.freeze({
    templateKey: 'GEAPA_OPERACIONAL',
    displayName: 'GEAPA Operacional',
    eyebrow: 'Comunicado Operacional',
    accent: '#0f766e',
    accentSoft: '#e9f8f6',
    accentSurface: '#f3fbfa'
  }),
  GEAPA_CONVITE: Object.freeze({
    templateKey: 'GEAPA_CONVITE',
    displayName: 'GEAPA Convite',
    eyebrow: 'Convite Institucional',
    accent: '#9a6700',
    accentSoft: '#fff3df',
    accentSurface: '#fffaf2'
  })
});

function coreMailRendererNormalizeTemplateKey_(templateKey) {
  return core_normalizeText_(templateKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailRendererGetTemplate_(templateKey) {
  var key = coreMailRendererNormalizeTemplateKey_(templateKey);
  var template = CORE_MAIL_RENDERER_TEMPLATES[key];

  if (!template) {
    throw new Error('Template institucional de e-mail nao encontrado: ' + templateKey);
  }

  return template;
}

function coreMailRendererGetInstitutionalQuote_(refDate) {
  try {
    var currentSlogan = typeof core_getCurrentBoardSlogan_ === 'function'
      ? String(core_getCurrentBoardSlogan_(refDate) || '').trim()
      : '';
    return currentSlogan || CORE_MAIL_RENDERER_BRAND.quoteFallback;
  } catch (err) {
    return CORE_MAIL_RENDERER_BRAND.quoteFallback;
  }
}

function coreMailRendererEscapeHtml_(value) {
  return String(value || '').replace(/[<>&"]/g, function(ch) {
    return {
      '<': '&lt;',
      '>': '&gt;',
      '&': '&amp;',
      '"': '&quot;'
    }[ch];
  });
}

function coreMailRendererNormalizeWhitespace_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function coreMailRendererStripHtml_(value) {
  return coreMailRendererNormalizeWhitespace_(
    String(value || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<li>/gi, '- ')
      .replace(/<\/li>/gi, '\n')
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
  );
}

function coreMailRendererTextToHtml_(text) {
  var paragraphs = String(text || '')
    .replace(/\r\n/g, '\n')
    .split(/\n\s*\n+/)
    .map(function(part) {
      return String(part || '').trim();
    })
    .filter(function(part) {
      return !!part;
    });

  if (!paragraphs.length) return '';

  return paragraphs.map(function(part) {
    return '<p style="margin:0 0 12px 0;font-size:14px;line-height:1.7;color:' +
      CORE_MAIL_RENDERER_BRAND.textColor +
      ';">' +
      coreMailRendererEscapeHtml_(part).replace(/\n/g, '<br>') +
      '</p>';
  }).join('');
}

function coreMailRendererRenderHiddenPreview_(text) {
  var preview = coreMailRendererEscapeHtml_(coreMailRendererNormalizeWhitespace_(text)).slice(0, 180);
  if (!preview) return '';

  return '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;">' +
    preview +
    '</div>';
}

function coreMailRendererNormalizeItems_(items) {
  if (!Array.isArray(items)) return [];

  return items.map(function(item) {
    if (item == null) return null;

    if (typeof item === 'string' || typeof item === 'number' || typeof item === 'boolean') {
      return {
        label: '',
        value: String(item),
        hint: ''
      };
    }

    return {
      label: String(item.label || item.line1 || '').trim(),
      value: String(item.value || item.line2 || item.text || '').trim(),
      hint: String(item.hint || '').trim()
    };
  }).filter(function(item) {
    return item && (item.label || item.value || item.hint);
  });
}

function coreMailRendererNormalizeBlocks_(payload) {
  var sourceBlocks = Array.isArray(payload.blocks)
    ? payload.blocks
    : (Array.isArray(payload.sections) ? payload.sections : []);

  if (!sourceBlocks.length && Array.isArray(payload.items) && payload.items.length) {
    sourceBlocks = [{
      title: String(payload.itemsTitle || '').trim(),
      items: payload.items
    }];
  }

  return sourceBlocks.map(function(block, index) {
    block = block || {};
    var items = coreMailRendererNormalizeItems_(block.items);
    var cta = block.cta || null;

    return {
      id: String(block.id || 'BLOCK_' + String(index + 1)).trim(),
      title: String(block.title || '').trim(),
      text: String(block.text || '').trim(),
      html: String(block.html || '').trim(),
      tone: coreMailRendererNormalizeTemplateKey_(block.tone || 'DEFAULT'),
      items: items,
      cta: cta ? {
        label: String(cta.label || '').trim(),
        url: String(cta.url || '').trim(),
        helper: String(cta.helper || '').trim()
      } : null
    };
  }).filter(function(block) {
    return block.title || block.text || block.html || block.items.length || (block.cta && block.cta.label);
  });
}

function coreMailRendererNormalizePayload_(subjectHuman, payload) {
  payload = payload || {};

  var title = String(payload.title || subjectHuman || '').trim();
  var subtitle = String(payload.subtitle || '').trim();
  var eyebrow = String(payload.eyebrow || '').trim();
  var introText = String(payload.introText || payload.introduction || '').trim();
  var introHtml = String(payload.introHtml || '').trim();
  var footerNote = String(payload.footerNote || '').trim();
  var preheader = String(payload.preheader || payload.previewText || subtitle || subjectHuman || '').trim();
  var refDate = payload.refDate || null;
  var blocks = coreMailRendererNormalizeBlocks_(payload);
  var cta = payload.cta ? {
    label: String(payload.cta.label || '').trim(),
    url: String(payload.cta.url || '').trim(),
    helper: String(payload.cta.helper || '').trim()
  } : null;

  return {
    title: title,
    subtitle: subtitle,
    eyebrow: eyebrow,
    introText: introText,
    introHtml: introHtml,
    footerNote: footerNote,
    preheader: preheader,
    refDate: refDate,
    blocks: blocks,
    cta: cta
  };
}

function coreMailRendererRenderEyebrow_(text, template) {
  var label = String(text || template.eyebrow || '').trim();
  if (!label) return '';

  return '<div style="display:inline-block;padding:7px 12px;border-radius:999px;background:' +
    template.accentSoft +
    ';color:' +
    template.accent +
    ';font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:16px;">' +
    coreMailRendererEscapeHtml_(label) +
    '</div>';
}

function coreMailRendererRenderItems_(items, template) {
  if (!items.length) return '';

  var content = items.map(function(item) {
    var line1 = item.label
      ? '<div style="font-size:14px;font-weight:700;color:' + CORE_MAIL_RENDERER_BRAND.textColor + ';">' +
          coreMailRendererEscapeHtml_(item.label) +
        '</div>'
      : '';
    var line2 = item.value
      ? '<div style="font-size:13px;line-height:1.6;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';margin-top:2px;">' +
          coreMailRendererEscapeHtml_(item.value) +
        '</div>'
      : '';
    var hint = item.hint
      ? '<div style="font-size:12px;line-height:1.5;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';margin-top:4px;">' +
          coreMailRendererEscapeHtml_(item.hint) +
        '</div>'
      : '';

    return '<div style="padding:12px 0;border-bottom:1px solid ' +
      CORE_MAIL_RENDERER_BRAND.dividerColor +
      ';">' +
      line1 +
      line2 +
      hint +
      '</div>';
  }).join('');

  return '<div style="margin-top:12px;">' + content + '</div>';
}

function coreMailRendererRenderCta_(cta, template) {
  if (!cta || !cta.label) return '';

  var helper = cta.helper
    ? '<div style="margin-top:10px;font-size:12px;line-height:1.5;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(cta.helper) +
      '</div>'
    : '';

  if (!cta.url) {
    return '<div style="margin-top:16px;padding:12px 14px;border-radius:12px;background:' +
      template.accentSoft +
      ';color:' +
      template.accent +
      ';font-size:13px;font-weight:700;">' +
      coreMailRendererEscapeHtml_(cta.label) +
      '</div>' +
      helper;
  }

  return '<div style="margin-top:18px;">' +
    '<a href="' + coreMailRendererEscapeHtml_(cta.url) + '" ' +
    'style="display:inline-block;padding:12px 18px;border-radius:10px;background:' +
    template.accent +
    ';color:' +
    CORE_MAIL_RENDERER_BRAND.buttonTextColor +
    ';font-size:14px;font-weight:700;text-decoration:none;">' +
    coreMailRendererEscapeHtml_(cta.label) +
    '</a>' +
    helper +
    '</div>';
}

function coreMailRendererRenderBlock_(block, template) {
  var blockBackground = block.tone === 'SOFT' ? template.accentSoft : template.accentSurface;
  var headerHtml = block.title
    ? '<div style="font-size:16px;font-weight:700;color:' + CORE_MAIL_RENDERER_BRAND.textColor + ';margin-bottom:10px;">' +
        coreMailRendererEscapeHtml_(block.title) +
      '</div>'
    : '';
  var textHtml = block.text ? coreMailRendererTextToHtml_(block.text) : '';
  var rawHtml = block.html ? block.html : '';
  var itemsHtml = coreMailRendererRenderItems_(block.items, template);
  var ctaHtml = coreMailRendererRenderCta_(block.cta, template);

  return '<div style="margin-top:16px;padding:18px 18px 6px 18px;border:1px solid ' +
    CORE_MAIL_RENDERER_BRAND.dividerColor +
    ';border-radius:14px;background:' +
    blockBackground +
    ';">' +
    headerHtml +
    textHtml +
    rawHtml +
    itemsHtml +
    ctaHtml +
    '</div>';
}

function coreMailRendererRenderHeader_(subjectHuman, payload, template) {
  var subtitleHtml = payload.subtitle
    ? '<div style="font-size:13px;line-height:1.6;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';margin:6px 0 0 0;">' +
        coreMailRendererEscapeHtml_(payload.subtitle) +
      '</div>'
    : '';
  var introHtml = payload.introHtml || coreMailRendererTextToHtml_(payload.introText);

  return '<div style="padding:28px 28px 10px 28px;">' +
    '<div style="text-align:center;margin-bottom:14px;">' +
      '<img src="cid:geapa_logo" alt="GEAPA" style="max-width:380px;width:100%;height:auto;">' +
    '</div>' +
    '<div style="height:4px;background:' + template.accent + ';margin-bottom:22px;border-radius:999px;"></div>' +
    coreMailRendererRenderEyebrow_(payload.eyebrow, template) +
    '<div style="font-size:28px;line-height:1.2;font-weight:700;color:' + CORE_MAIL_RENDERER_BRAND.textColor + ';margin:0;">' +
      coreMailRendererEscapeHtml_(payload.title || subjectHuman) +
    '</div>' +
    subtitleHtml +
    (introHtml ? '<div style="margin-top:18px;">' + introHtml + '</div>' : '') +
  '</div>';
}

function coreMailRendererRenderFooter_(payload, template) {
  var institutionalQuote = coreMailRendererGetInstitutionalQuote_(payload.refDate || null);
  var footerNoteHtml = payload.footerNote
    ? '<div style="margin-top:12px;font-size:12px;line-height:1.6;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(payload.footerNote) +
      '</div>'
    : '';

  return '<div style="padding:22px 28px 28px 28px;border-top:1px solid ' +
    CORE_MAIL_RENDERER_BRAND.dividerColor +
    ';margin-top:24px;background:' +
    template.accentSoft +
    ';">' +
    '<div style="font-size:14px;line-height:1.7;color:' + CORE_MAIL_RENDERER_BRAND.textColor + ';">' +
      'Atenciosamente,<br>' +
      '<strong>' + coreMailRendererEscapeHtml_(CORE_MAIL_RENDERER_BRAND.orgName) + '</strong>' +
    '</div>' +
    '<div style="margin-top:10px;font-size:12px;line-height:1.6;color:' + CORE_MAIL_RENDERER_BRAND.mutedTextColor + ';">' +
      coreMailRendererEscapeHtml_(institutionalQuote) +
    '</div>' +
    footerNoteHtml +
  '</div>';
}

function coreMailRendererBuildTextBody_(template, subjectHuman, payload) {
  var institutionalQuote = coreMailRendererGetInstitutionalQuote_(payload.refDate || null);
  var lines = [];
  lines.push(CORE_MAIL_RENDERER_BRAND.shortName + ' - ' + template.displayName);
  lines.push(subjectHuman);

  if (payload.subtitle) lines.push(payload.subtitle);
  if (payload.introText) lines.push(payload.introText);
  if (payload.introHtml) lines.push(coreMailRendererStripHtml_(payload.introHtml));

  payload.blocks.forEach(function(block) {
    if (block.title) lines.push(block.title);
    if (block.text) lines.push(block.text);
    if (block.html) lines.push(coreMailRendererStripHtml_(block.html));

    block.items.forEach(function(item) {
      if (item.label && item.value) {
        lines.push(item.label + ': ' + item.value);
      } else if (item.label) {
        lines.push(item.label);
      } else if (item.value) {
        lines.push(item.value);
      }

      if (item.hint) lines.push(item.hint);
    });

    if (block.cta && block.cta.label) {
      lines.push(block.cta.label + (block.cta.url ? ' - ' + block.cta.url : ''));
      if (block.cta.helper) lines.push(block.cta.helper);
    }
  });

  if (payload.cta && payload.cta.label) {
    lines.push(payload.cta.label + (payload.cta.url ? ' - ' + payload.cta.url : ''));
    if (payload.cta.helper) lines.push(payload.cta.helper);
  }

  lines.push(CORE_MAIL_RENDERER_BRAND.orgName);
  lines.push(institutionalQuote);

  if (payload.footerNote) lines.push(payload.footerNote);

  return lines.filter(function(line) {
    return !!coreMailRendererNormalizeWhitespace_(line);
  }).join('\n\n');
}

function coreMailBuildFinalSubject_(subjectHuman, correlationKey) {
  core_assertRequired_(subjectHuman, 'subjectHuman');
  core_assertRequired_(correlationKey, 'correlationKey');

  var cleanSubject = String(subjectHuman || '').trim();
  var cleanKey = String(correlationKey || '').trim().toUpperCase();
  var tag = '[GEAPA][' + cleanKey + ']';

  if (/\[GEAPA\]\[[^\]]+\]/i.test(cleanSubject)) {
    return cleanSubject.replace(/\[GEAPA\]\[[^\]]+\]/i, tag).trim();
  }

  return (tag + ' ' + cleanSubject).trim();
}

function coreMailRenderEmailTemplate_(templateKey, subjectHuman, payload) {
  core_assertRequired_(templateKey, 'templateKey');
  core_assertRequired_(subjectHuman, 'subjectHuman');

  var template = coreMailRendererGetTemplate_(templateKey);
  var normalizedPayload = coreMailRendererNormalizePayload_(subjectHuman, payload || {});
  var blocksHtml = normalizedPayload.blocks.map(function(block) {
    return coreMailRendererRenderBlock_(block, template);
  }).join('');
  var rootCtaHtml = normalizedPayload.cta ? coreMailRendererRenderCta_(normalizedPayload.cta, template) : '';

  var htmlBody = ''
    + coreMailRendererRenderHiddenPreview_(normalizedPayload.preheader)
    + '<div style="margin:0;padding:24px;background:' + CORE_MAIL_RENDERER_BRAND.background + ';">'
    + '  <div style="max-width:700px;margin:0 auto;">'
    + '    <div style="background:' + CORE_MAIL_RENDERER_BRAND.cardBackground + ';border:2px solid ' + template.accent + ';border-radius:18px;overflow:hidden;box-shadow:0 10px 30px ' + CORE_MAIL_RENDERER_BRAND.shadowColor + ';">'
    +         coreMailRendererRenderHeader_(subjectHuman, normalizedPayload, template)
    + '      <div style="padding:0 28px 0 28px;">'
    +            blocksHtml
    +            rootCtaHtml
    + '      </div>'
    +         coreMailRendererRenderFooter_(normalizedPayload, template)
    + '    </div>'
    + '  </div>'
    + '</div>';

  return Object.freeze({
    templateKey: template.templateKey,
    templateName: template.displayName,
    subjectHuman: String(subjectHuman || '').trim(),
    htmlBody: htmlBody,
    bodyText: coreMailRendererBuildTextBody_(template, subjectHuman, normalizedPayload),
    inlineImagesMode: 'DEFAULT_GEAPA',
    meta: Object.freeze({
      preheader: normalizedPayload.preheader,
      blockCount: normalizedPayload.blocks.length
    })
  });
}

function coreMailRendererNormalizeRecipients_(value, required) {
  if (value === '' || value == null) {
    return required ? core_resolveEmailRecipients_(value) : [];
  }
  return core_resolveEmailRecipients_(value);
}

function coreMailBuildOutgoingDraft_(contract) {
  core_assertRequired_(contract, 'contract');
  core_assertRequired_(contract.moduleName, 'contract.moduleName');
  core_assertRequired_(contract.templateKey, 'contract.templateKey');
  core_assertRequired_(contract.correlationKey, 'contract.correlationKey');
  core_assertRequired_(contract.subjectHuman, 'contract.subjectHuman');
  core_assertRequired_(contract.to, 'contract.to');

  var toList = coreMailRendererNormalizeRecipients_(contract.to, true);
  var ccList = coreMailRendererNormalizeRecipients_(contract.cc, false);
  var bccList = coreMailRendererNormalizeRecipients_(contract.bcc, false);
  var correlationKey = String(contract.correlationKey || '').trim().toUpperCase();
  var renderResult = coreMailRenderEmailTemplate_(
    contract.templateKey,
    contract.subjectHuman,
    contract.payload || {}
  );
  var finalSubject = coreMailBuildFinalSubject_(contract.subjectHuman, correlationKey);

  return Object.freeze({
    moduleName: String(contract.moduleName || '').trim(),
    templateKey: renderResult.templateKey,
    correlationKey: correlationKey,
    to: toList.join(','),
    cc: ccList.join(','),
    bcc: bccList.join(','),
    toList: Object.freeze(toList.slice()),
    ccList: Object.freeze(ccList.slice()),
    bccList: Object.freeze(bccList.slice()),
    subjectHuman: String(contract.subjectHuman || '').trim(),
    subject: finalSubject,
    payload: contract.payload || {},
    htmlBody: renderResult.htmlBody,
    bodyText: renderResult.bodyText,
    inlineImagesMode: renderResult.inlineImagesMode,
    emailOptions: Object.freeze({
      to: toList.join(','),
      cc: ccList.join(','),
      bcc: bccList.join(','),
      subject: finalSubject,
      body: renderResult.bodyText,
      htmlBody: renderResult.htmlBody,
      useDefaultInlineImages: true,
      name: CORE_MAIL_RENDERER_BRAND.shortName
    })
  });
}
