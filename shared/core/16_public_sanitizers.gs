// F4A: native synchronous port of F3.5 shared/core/public-adapters.cjs. No external module loading.
function f4Public_() {
  if (f4Public_.value_) return f4Public_.value_;
  return (f4Public_.value_ = (function () {
'use strict';
// Exact sanitizers from preserved public-data-v2.js; no network/cache code copied.
const SCHEMA='geapa-public-v2.1';
  function text(value, max) {
    return String(value == null ? '' : value).replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<[^>]*>/g, '')
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '').trim().slice(0, max || 8000);
  }
  function url(value, imageOnly) {
    var valueUrl = String(value || '').trim();
    if (/[\u0000-\u0020\u007f<>"'\\]/.test(valueUrl)) return '';
    if (!imageOnly && /^#\/[a-z0-9/-]+(?:\?[a-z0-9=&_-]+)?$/i.test(valueUrl)) return valueUrl;
    return /^https:\/\/[a-z0-9.-]+(?::443)?(?:[/?#]|$)/i.test(valueUrl) && !/^https:\/\/[^/]*@/i.test(valueUrl)
      ? valueUrl.slice(0, 2000) : '';
  }
  function sort(items) { return items.sort(function(a, b) { return a.order - b.order || a.id.localeCompare(b.id); }); }
  function section(value) {
    var s = value || {}, type = String(s.type || 'TEXT');
    if (s.published !== true || ['TEXT', 'HERO', 'CTA', 'IMAGE_TEXT', 'DOCUMENT_LIST'].indexOf(type) < 0) return null;
    if (!s.id || !isFinite(Number(s.order))) return null;
    // DOCUMENT_LIST e um slot do adapter, nunca texto normativo embutido.
    return { id: text(s.id, 100), type: type, title: text(s.title, 240), subtitle: text(s.subtitle, 600),
      text: type === 'DOCUMENT_LIST' ? '' : text(s.text), imageUrl: url(s.imageUrl, true),
      buttonText: text(s.buttonText, 100), buttonUrl: url(s.buttonUrl), order: Number(s.order), published: true };
  }
  function page(value, id) {
    var p = value || {};
    if (p.id !== id || p.published !== true || p.sanitized !== true || p.schemaVersion !== SCHEMA || !Array.isArray(p.sections)) {
      throw new Error('PUBLIC_V2_INVALID_PAGE');
    }
    var sections = sort(p.sections.slice(0, 30).map(section).filter(Boolean));
    if (!sections.length || p.sections.length > 30) throw new Error('PUBLIC_V2_EMPTY_OR_OVERSIZED_PAGE');
    return { id: id, schemaVersion: SCHEMA, sanitized: true, published: true, title: text(p.title, 240),
      subtitle: text(p.subtitle, 600), sections: sections, sourceHash: text(p.sourceHash, 100),
      sourceUpdatedAt: text(p.sourceUpdatedAt, 40), updatedAt: text(p.updatedAt, 40), revision: Number(p.revision || 0) };
  }
return {page};

  })());
}
