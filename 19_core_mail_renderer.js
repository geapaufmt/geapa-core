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
  orgName: 'Grupo de Estudos de Apoio \u00e0 Produ\u00e7\u00e3o Agr\u00edcola',
  shortName: 'GEAPA',
  quoteFallback: '"Cultivar o Conhecimento Para Colher Sabedoria"',
  officialEmail: '',
  background: '#f5f8f5',
  cardBackground: '#ffffff',
  textColor: '#163226',
  mutedTextColor: '#587062',
  borderColor: '#167d0a',
  dividerColor: '#d6e5d3',
  buttonTextColor: '#ffffff',
  shadowColor: 'rgba(22,125,10,0.08)'
});

var CORE_MAIL_RENDERER_OFFICIAL_DATA_KEY = 'DADOS_OFICIAIS_GEAPA';

var CORE_MAIL_RENDERER_TEMPLATES = Object.freeze({
  GEAPA_COMEMORATIVO: Object.freeze({
    templateKey: 'GEAPA_COMEMORATIVO',
    displayName: 'GEAPA Comemorativo',
    eyebrow: 'Mensagem Comemorativa',
    accentRole: 'GREEN'
  }),
  GEAPA_OPERACIONAL: Object.freeze({
    templateKey: 'GEAPA_OPERACIONAL',
    displayName: 'GEAPA Operacional',
    eyebrow: 'Comunicado Operacional',
    accentRole: 'GREEN'
  }),
  GEAPA_CONVITE: Object.freeze({
    templateKey: 'GEAPA_CONVITE',
    displayName: 'GEAPA Convite',
    eyebrow: 'Convite Institucional',
    accentRole: 'BROWN'
  }),
  GEAPA_CLASSICO: Object.freeze({
    templateKey: 'GEAPA_CLASSICO',
    displayName: 'GEAPA Classico',
    eyebrow: '',
    accentRole: 'GREEN'
  })
});

function coreMailRendererNormalizeTemplateKey_(templateKey) {
  return core_normalizeText_(templateKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailRendererNormalizeHexColor_(value, fallback) {
  var text = String(value || '').trim();
  if (/^#([0-9a-fA-F]{6})$/.test(text)) return text;
  if (/^#([0-9a-fA-F]{3})$/.test(text)) {
    return '#' + text.substring(1).split('').map(function(ch) {
      return ch + ch;
    }).join('');
  }
  return fallback;
}

function coreMailRendererHexToRgb_(hex) {
  var normalized = coreMailRendererNormalizeHexColor_(hex, '#000000');
  return {
    r: parseInt(normalized.substring(1, 3), 16),
    g: parseInt(normalized.substring(3, 5), 16),
    b: parseInt(normalized.substring(5, 7), 16)
  };
}

function coreMailRendererRgbToHex_(rgb) {
  function clamp(value) {
    return Math.max(0, Math.min(255, Math.round(Number(value || 0))));
  }

  function toHex(value) {
    var out = clamp(value).toString(16);
    return out.length === 1 ? '0' + out : out;
  }

  return '#' + toHex(rgb.r) + toHex(rgb.g) + toHex(rgb.b);
}

function coreMailRendererMixHexColors_(baseHex, mixHex, ratio) {
  var base = coreMailRendererHexToRgb_(baseHex);
  var mix = coreMailRendererHexToRgb_(mixHex);
  var weight = Math.max(0, Math.min(1, Number(ratio || 0)));

  return coreMailRendererRgbToHex_({
    r: base.r * (1 - weight) + mix.r * weight,
    g: base.g * (1 - weight) + mix.g * weight,
    b: base.b * (1 - weight) + mix.b * weight
  });
}

function coreMailRendererBuildInstitutionLine_(officialData) {
  var parts = [
    String(officialData.courseName || '').trim(),
    String(officialData.instituteShortName || officialData.instituteName || '').trim(),
    String(officialData.universityShortName || officialData.universityName || '').trim()
  ].filter(function(part) {
    return !!part;
  });

  return parts.join(' - ');
}

function coreMailRendererShouldShowShortName_(brand) {
  var shortName = String((brand && brand.shortName) || '').trim();
  var orgName = String((brand && brand.orgName) || '').trim();
  if (!shortName || !orgName) return false;

  return orgName.toUpperCase().indexOf(shortName.toUpperCase()) === -1;
}

function coreMailRendererGetOfficialGroupData_() {
  try {
    var records = core_readRecordsByKey_(CORE_MAIL_RENDERER_OFFICIAL_DATA_KEY, {
      skipBlankRows: true
    });
    var record = records && records.length ? records[0] : {};

    return {
      orgName: String(record.NOME_OFICIAL_GRUPO || CORE_MAIL_RENDERER_BRAND.orgName || '').trim() || CORE_MAIL_RENDERER_BRAND.orgName,
      shortName: String(record.SIGLA_OFICIAL_GRUPO || CORE_MAIL_RENDERER_BRAND.shortName || '').trim() || CORE_MAIL_RENDERER_BRAND.shortName,
      officialEmail: String(record.EMAIL_OFICIAL || '').trim(),
      greenColor: coreMailRendererNormalizeHexColor_(record.VERDE_OFICIAL, CORE_MAIL_RENDERER_BRAND.borderColor),
      brownColor: coreMailRendererNormalizeHexColor_(record.MARROM_OFICIAL, '#9a6700'),
      blackColor: coreMailRendererNormalizeHexColor_(record.PRETO_OFICIAL, CORE_MAIL_RENDERER_BRAND.textColor),
      universityName: String(record.UNIVERSIDADE_MAE || '').trim(),
      universityShortName: String(record.SIGLA_UNIVERSIDADE_MAE || '').trim(),
      instituteName: String(record['INSTITUTO MAE'] || '').trim(),
      instituteShortName: String(record['SIGLA INSTITUTO MAE'] || '').trim(),
      courseName: String(record['CURSO MAE'] || '').trim()
    };
  } catch (err) {
    return {
      orgName: CORE_MAIL_RENDERER_BRAND.orgName,
      shortName: CORE_MAIL_RENDERER_BRAND.shortName,
      officialEmail: CORE_MAIL_RENDERER_BRAND.officialEmail,
      greenColor: CORE_MAIL_RENDERER_BRAND.borderColor,
      brownColor: '#9a6700',
      blackColor: CORE_MAIL_RENDERER_BRAND.textColor,
      universityName: '',
      universityShortName: '',
      instituteName: '',
      instituteShortName: '',
      courseName: ''
    };
  }
}

function coreMailRendererGetBrandProfile_(templateKey, refDate) {
  var officialData = coreMailRendererGetOfficialGroupData_();
  var quote = coreMailRendererGetInstitutionalQuote_(refDate);
  var accentSource = coreMailRendererNormalizeTemplateKey_(templateKey) === 'GEAPA_CONVITE'
    ? officialData.brownColor
    : officialData.greenColor;
  var background = coreMailRendererNormalizeTemplateKey_(templateKey) === 'GEAPA_CLASSICO'
    ? '#ffffff'
    : coreMailRendererMixHexColors_(officialData.greenColor, '#ffffff', 0.94);
  var dividerColor = coreMailRendererMixHexColors_(accentSource, '#ffffff', 0.78);

  return {
    orgName: officialData.orgName,
    shortName: officialData.shortName,
    quote: quote,
    officialEmail: officialData.officialEmail,
    institutionLine: coreMailRendererBuildInstitutionLine_(officialData),
    background: background,
    cardBackground: CORE_MAIL_RENDERER_BRAND.cardBackground,
    textColor: officialData.blackColor,
    mutedTextColor: coreMailRendererMixHexColors_(officialData.blackColor, '#ffffff', 0.35),
    borderColor: officialData.greenColor,
    dividerColor: dividerColor,
    buttonTextColor: CORE_MAIL_RENDERER_BRAND.buttonTextColor,
    shadowColor: CORE_MAIL_RENDERER_BRAND.shadowColor,
    accent: accentSource,
    accentSoft: coreMailRendererMixHexColors_(accentSource, '#ffffff', 0.87),
    accentSurface: coreMailRendererMixHexColors_(accentSource, '#ffffff', 0.93)
  };
}

function coreMailRendererGetTemplate_(templateKey, brand) {
  var key = coreMailRendererNormalizeTemplateKey_(templateKey);
  var template = CORE_MAIL_RENDERER_TEMPLATES[key];

  if (!template) {
    throw new Error('Template institucional de e-mail nao encontrado: ' + templateKey);
  }

  brand = brand || coreMailRendererGetBrandProfile_(key, null);

  return {
    templateKey: template.templateKey,
    displayName: template.displayName,
    eyebrow: template.eyebrow,
    accentRole: template.accentRole,
    accent: brand.accent,
    accentSoft: brand.accentSoft,
    accentSurface: brand.accentSurface
  };
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

function coreMailRendererTextToHtml_(text, brand) {
  brand = brand || CORE_MAIL_RENDERER_BRAND;
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
      brand.textColor +
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

function coreMailRendererRenderItems_(items, template, brand) {
  brand = brand || CORE_MAIL_RENDERER_BRAND;
  if (!items.length) return '';

  var content = items.map(function(item) {
    var line1 = item.label
      ? '<div style="font-size:14px;font-weight:700;color:' + brand.textColor + ';">' +
          coreMailRendererEscapeHtml_(item.label) +
        '</div>'
      : '';
    var line2 = item.value
      ? '<div style="font-size:13px;line-height:1.6;color:' + brand.mutedTextColor + ';margin-top:2px;">' +
          coreMailRendererEscapeHtml_(item.value) +
        '</div>'
      : '';
    var hint = item.hint
      ? '<div style="font-size:12px;line-height:1.5;color:' + brand.mutedTextColor + ';margin-top:4px;">' +
          coreMailRendererEscapeHtml_(item.hint) +
        '</div>'
      : '';

    return '<div style="padding:12px 0;border-bottom:1px solid ' +
      brand.dividerColor +
      ';">' +
      line1 +
      line2 +
      hint +
      '</div>';
  }).join('');

  return '<div style="margin-top:12px;">' + content + '</div>';
}

function coreMailRendererRenderCta_(cta, template, brand) {
  brand = brand || CORE_MAIL_RENDERER_BRAND;
  if (!cta || !cta.label) return '';

  var helper = cta.helper
    ? '<div style="margin-top:10px;font-size:12px;line-height:1.5;color:' + brand.mutedTextColor + ';">' +
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
    brand.buttonTextColor +
    ';font-size:14px;font-weight:700;text-decoration:none;">' +
    coreMailRendererEscapeHtml_(cta.label) +
    '</a>' +
    helper +
    '</div>';
}

function coreMailRendererRenderBlock_(block, template, brand) {
  brand = brand || CORE_MAIL_RENDERER_BRAND;
  var blockBackground = block.tone === 'SOFT' ? template.accentSoft : template.accentSurface;
  var headerHtml = block.title
    ? '<div style="font-size:16px;font-weight:700;color:' + brand.textColor + ';margin-bottom:10px;">' +
        coreMailRendererEscapeHtml_(block.title) +
      '</div>'
    : '';
  var textHtml = block.text ? coreMailRendererTextToHtml_(block.text, brand) : '';
  var rawHtml = block.html ? block.html : '';
  var itemsHtml = coreMailRendererRenderItems_(block.items, template, brand);
  var ctaHtml = coreMailRendererRenderCta_(block.cta, template, brand);

  return '<div style="margin-top:16px;padding:18px 18px 6px 18px;border:1px solid ' +
    brand.dividerColor +
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

function coreMailRendererRenderHeader_(subjectHuman, payload, template, brand) {
  brand = brand || CORE_MAIL_RENDERER_BRAND;
  var subtitleHtml = payload.subtitle
    ? '<div style="font-size:13px;line-height:1.6;color:' + brand.mutedTextColor + ';margin:6px 0 0 0;">' +
        coreMailRendererEscapeHtml_(payload.subtitle) +
      '</div>'
    : '';
  var introHtml = payload.introHtml || coreMailRendererTextToHtml_(payload.introText, brand);

  return '<div style="padding:28px 28px 10px 28px;">' +
    '<div style="text-align:center;margin-bottom:14px;">' +
      '<img src="cid:geapa_logo" alt="' + coreMailRendererEscapeHtml_(brand.shortName || 'GEAPA') + '" style="max-width:380px;width:100%;height:auto;">' +
    '</div>' +
    '<div style="height:4px;background:' + template.accent + ';margin-bottom:22px;border-radius:999px;"></div>' +
    coreMailRendererRenderEyebrow_(payload.eyebrow, template) +
    '<div style="font-size:28px;line-height:1.2;font-weight:700;color:' + brand.textColor + ';margin:0;">' +
      coreMailRendererEscapeHtml_(payload.title || subjectHuman) +
    '</div>' +
    subtitleHtml +
    (introHtml ? '<div style="margin-top:18px;">' + introHtml + '</div>' : '') +
  '</div>';
}

function coreMailRendererFlattenClassicRows_(payload) {
  var rows = [];

  payload.blocks.forEach(function(block) {
    if (block.title && !block.text && !block.html && !block.items.length && !(block.cta && block.cta.label)) {
      rows.push({
        line1: block.title,
        line2: ''
      });
    }

    if (block.title && (block.items.length || block.text || block.html || (block.cta && block.cta.label))) {
      rows.push({
        line1: block.title,
        line2: ''
      });
    }

    if (block.text) {
      rows.push({
        line1: '',
        line2: block.text
      });
    }

    if (block.html) {
      rows.push({
        line1: '',
        line2: coreMailRendererStripHtml_(block.html)
      });
    }

    block.items.forEach(function(item) {
      rows.push({
        line1: item.label || item.value || '',
        line2: item.label && item.value ? item.value : (item.hint || '')
      });
    });

    if (block.cta && block.cta.label) {
      rows.push({
        line1: block.cta.label,
        line2: block.cta.url || block.cta.helper || ''
      });
    }
  });

  if (payload.cta && payload.cta.label) {
    rows.push({
      line1: payload.cta.label,
      line2: payload.cta.url || payload.cta.helper || ''
    });
  }

  return rows.filter(function(row) {
    return !!String(row.line1 || '').trim() || !!String(row.line2 || '').trim();
  });
}

function coreMailRendererRenderClassicRows_(rows, brand) {
  if (!rows.length) {
    return '<div style="color:#555;">(sem itens)</div>';
  }

  return rows.map(function(row) {
    var line1 = String(row.line1 || '').trim();
    var line2 = String(row.line2 || '').trim();

    return '<div style="padding:10px 0;border-bottom:1px solid #eee;">' +
      (line1
        ? '<div style="font-size:16px;font-weight:600;color:' + brand.textColor + ';">' +
            coreMailRendererEscapeHtml_(line1) +
          '</div>'
        : '') +
      (line2
        ? '<div style="font-size:13px;color:#555;margin-top:2px;line-height:1.55;">' +
            coreMailRendererEscapeHtml_(line2) +
          '</div>'
        : '') +
      '</div>';
  }).join('');
}

function coreMailRendererRenderClassicFooter_(payload, brand) {
  var institutionalQuote = brand.quote || coreMailRendererGetInstitutionalQuote_(payload.refDate || null);
  var shortNameHtml = coreMailRendererShouldShowShortName_(brand)
    ? '<div style="margin-top:4px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';font-weight:700;">' +
        coreMailRendererEscapeHtml_(brand.shortName) +
      '</div>'
    : '';
  var contactHtml = brand.officialEmail
    ? '<div style="margin-top:6px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(brand.officialEmail) +
      '</div>'
    : '';
  var institutionHtml = brand.institutionLine
    ? '<div style="margin-top:4px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(brand.institutionLine) +
      '</div>'
    : '';
  var quoteHtml = institutionalQuote
    ? '<div style="margin-top:10px;font-size:12px;line-height:1.6;color:#666;"><em>' +
        coreMailRendererEscapeHtml_(institutionalQuote) +
      '</em></div>'
    : '';
  var footerNoteHtml = payload.footerNote
    ? '<div style="margin-top:10px;font-size:12px;line-height:1.6;color:#666;">' +
        coreMailRendererEscapeHtml_(payload.footerNote) +
      '</div>'
    : '';

  return '<div style="margin-top:16px;font-size:13px;line-height:1.6;color:' + brand.textColor + ';">' +
      'Atenciosamente,<br>' +
      '<strong>' + coreMailRendererEscapeHtml_(brand.orgName) + '</strong>' +
    '</div>' +
    shortNameHtml +
    contactHtml +
    institutionHtml +
    quoteHtml +
    footerNoteHtml;
}

function coreMailRendererRenderClassicTemplate_(subjectHuman, payload, template, brand) {
  var subtitleHtml = payload.subtitle
    ? '<div style="font-size:13px;color:#555;margin-bottom:14px;">' +
        coreMailRendererEscapeHtml_(payload.subtitle) +
      '</div>'
    : '';
  var introHtml = payload.introHtml || coreMailRendererTextToHtml_(payload.introText, brand);
  var rowsHtml = coreMailRendererRenderClassicRows_(coreMailRendererFlattenClassicRows_(payload), brand);
  var logoHtml = '<div style="text-align:center;margin-bottom:15px;">' +
      '<img src="cid:geapa_logo" alt="' + coreMailRendererEscapeHtml_(brand.shortName || 'GEAPA') + '" style="max-width:380px;width:100%;height:auto;">' +
    '</div>' +
    '<div style="height:4px;background:' + template.accent + ';margin-bottom:20px;border-radius:2px;"></div>';

  return ''
    + coreMailRendererRenderHiddenPreview_(payload.preheader)
    + '<div style="font-family:Arial,sans-serif;background:#ffffff;color:' + brand.textColor + ';padding:16px;">'
    + '  <div style="max-width:700px;margin:0 auto;border:2px solid ' + template.accent + ';border-radius:12px;padding:16px;background:#ffffff;">'
    +       logoHtml
    + '    <div style="font-size:20px;font-weight:700;margin-bottom:4px;color:' + brand.textColor + ';">' +
              coreMailRendererEscapeHtml_(payload.title || subjectHuman) +
            '</div>'
    +       subtitleHtml
    +       (introHtml ? '<div style="margin-bottom:18px;font-size:14px;line-height:1.6;">' + introHtml + '</div>' : '')
    +       rowsHtml
    +       coreMailRendererRenderClassicFooter_(payload, brand)
    + '  </div>'
    + '</div>';
}

function coreMailRendererRenderFooter_(payload, template) {
  var brand = payload.brand || CORE_MAIL_RENDERER_BRAND;
  var institutionalQuote = brand.quote || coreMailRendererGetInstitutionalQuote_(payload.refDate || null);
  var contactHtml = brand.officialEmail
    ? '<div style="margin-top:8px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
        'Contato oficial: ' + coreMailRendererEscapeHtml_(brand.officialEmail) +
      '</div>'
    : '';
  var institutionHtml = brand.institutionLine
    ? '<div style="margin-top:6px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(brand.institutionLine) +
      '</div>'
    : '';
  var shortNameHtml = coreMailRendererShouldShowShortName_(brand)
    ? '<div style="margin-top:4px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';font-weight:700;">' +
        coreMailRendererEscapeHtml_(brand.shortName) +
      '</div>'
    : '';
  var footerNoteHtml = payload.footerNote
    ? '<div style="margin-top:12px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
        coreMailRendererEscapeHtml_(payload.footerNote) +
      '</div>'
    : '';

  return '<div style="padding:22px 28px 28px 28px;border-top:1px solid ' +
    brand.dividerColor +
    ';margin-top:24px;background:' +
    template.accentSoft +
    ';">' +
    '<div style="font-size:14px;line-height:1.7;color:' + brand.textColor + ';">' +
      'Atenciosamente,<br>' +
      '<strong>' + coreMailRendererEscapeHtml_(brand.orgName) + '</strong>' +
    '</div>' +
    shortNameHtml +
    contactHtml +
     institutionHtml +
     '<div style="margin-top:10px;font-size:12px;line-height:1.6;color:' + brand.mutedTextColor + ';">' +
      '<em>' + coreMailRendererEscapeHtml_(institutionalQuote) + '</em>' +
     '</div>' +
     footerNoteHtml +
   '</div>';
}

function coreMailRendererBuildTextBody_(template, subjectHuman, payload) {
  var brand = payload.brand || CORE_MAIL_RENDERER_BRAND;
  var institutionalQuote = brand.quote || coreMailRendererGetInstitutionalQuote_(payload.refDate || null);
  var lines = [];
  lines.push(brand.shortName + ' - ' + template.displayName);
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

  lines.push(brand.orgName);
  if (coreMailRendererShouldShowShortName_(brand)) lines.push(brand.shortName);
  if (brand.officialEmail) lines.push('Contato oficial: ' + brand.officialEmail);
  if (brand.institutionLine) lines.push(brand.institutionLine);
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

  var normalizedPayload = coreMailRendererNormalizePayload_(subjectHuman, payload || {});
  var brand = coreMailRendererGetBrandProfile_(templateKey, normalizedPayload.refDate || null);
  normalizedPayload.brand = brand;
  var template = coreMailRendererGetTemplate_(templateKey, brand);
  var htmlBody = '';

  if (template.templateKey === 'GEAPA_CLASSICO') {
    htmlBody = coreMailRendererRenderClassicTemplate_(subjectHuman, normalizedPayload, template, brand);
  } else {
    var blocksHtml = normalizedPayload.blocks.map(function(block) {
      return coreMailRendererRenderBlock_(block, template, brand);
    }).join('');
    var rootCtaHtml = normalizedPayload.cta ? coreMailRendererRenderCta_(normalizedPayload.cta, template, brand) : '';

    htmlBody = ''
      + coreMailRendererRenderHiddenPreview_(normalizedPayload.preheader)
      + '<div style="margin:0;padding:24px;background:' + brand.background + ';">'
      + '  <div style="max-width:700px;margin:0 auto;">'
      + '    <div style="background:' + brand.cardBackground + ';border:2px solid ' + template.accent + ';border-radius:18px;overflow:hidden;box-shadow:0 10px 30px ' + brand.shadowColor + ';">'
      +         coreMailRendererRenderHeader_(subjectHuman, normalizedPayload, template, brand)
      + '      <div style="padding:0 28px 0 28px;">'
      +            blocksHtml
      +            rootCtaHtml
      + '      </div>'
      +         coreMailRendererRenderFooter_(normalizedPayload, template)
      + '    </div>'
      + '  </div>'
      + '</div>';
  }

  return Object.freeze({
    templateKey: template.templateKey,
    templateName: template.displayName,
    subjectHuman: String(subjectHuman || '').trim(),
    htmlBody: htmlBody,
    bodyText: coreMailRendererBuildTextBody_(template, subjectHuman, normalizedPayload),
    inlineImagesMode: 'DEFAULT_GEAPA',
    meta: Object.freeze({
      preheader: normalizedPayload.preheader,
      blockCount: normalizedPayload.blocks.length,
      orgName: brand.orgName,
      shortName: brand.shortName,
      officialEmail: brand.officialEmail
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
      replyTo: contract.replyTo || undefined,
      useDefaultInlineImages: true,
      name: renderResult.meta.shortName || CORE_MAIL_RENDERER_BRAND.shortName
    })
  });
}
