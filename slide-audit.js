// Runs inside the captured page. Sizes are measured after CSS transforms and final frame fitting.
document.addEventListener('DOMContentLoaded', async () => {
  // Let all DOMContentLoaded handlers (including local KaTeX) finish before inspecting.
  await new Promise(resolve => setTimeout(resolve, 0));
  await document.fonts.ready;
  await Promise.all([...document.images].map(im => im.complete ? Promise.resolve() : new Promise(resolve => {
    im.addEventListener('load', resolve, { once: true }); im.addEventListener('error', resolve, { once: true });
  })));
  const config = window.biimAuditConfig;
  const issues = [], warnings = [], texts = [], metrics = [];
  const add = (code, text) => issues.push({ code, text: String(text).slice(0, 100) });
  const outside = (r, b) => r.left < b.left - 1 || r.top < b.top - 1 || r.right > b.right + 1 || r.bottom > b.bottom + 1;
  const canvas = { left: 0, top: 0, right: 1280, bottom: 720 };
  function effectiveScale(el) {
    let scale = 1;
    for (let p = el; p; p = p.parentElement) {
      const s = getComputedStyle(p);
      if (s.transform !== 'none') {
        const m = new DOMMatrixReadOnly(s.transform);
        // Smallest singular value catches rotated/non-uniform shrinking too.
        const sum = m.a*m.a + m.b*m.b + m.c*m.c + m.d*m.d, det = m.a*m.d - m.b*m.c;
        scale *= Math.sqrt(Math.max(0, (sum - Math.sqrt(Math.max(0, sum*sum - 4*det*det))) / 2));
      }
      scale *= parseFloat(s.zoom) || 1;
    }
    return scale;
  }
  const rgba = value => {
    const a = value.match(/[\d.]+/g);
    return a && a.length >= 3 ? [Number(a[0]), Number(a[1]), Number(a[2]), a.length > 3 ? Number(a[3]) : 1] : null;
  };
  const luminance = c => c.slice(0,3).map(x => { x /= 255; return x <= .04045 ? x/12.92 : ((x+.055)/1.055)**2.4; })
    .reduce((v, x, i) => v + x * [.2126,.7152,.0722][i], 0);
  function contrast(el) {
    const color = rgba(getComputedStyle(el).color);
    if (!color) return null;
    for (let p = el; p; p = p.parentElement) {
      const s = getComputedStyle(p);
      if (s.backgroundImage !== 'none') return null; // Raster/gradient backgrounds need visual review.
      const bg = rgba(s.backgroundColor);
      if (bg && bg[3] >= .99) {
        const fg = color.map((v,i) => i < 3 ? v*color[3] + bg[i]*(1-color[3]) : 1);
        const a = luminance(fg), b = luminance(bg);
        return (Math.max(a,b)+.05)/(Math.min(a,b)+.05);
      }
      if (bg && bg[3] > 0) return null;
    }
    return null;
  }
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let t;
  while ((t = walker.nextNode())) {
    const el = t.parentElement, text = t.textContent.trim();
    if (!text || el.closest('script,style,.katex-mathml')) continue;
    if (el.getBoundingClientRect().width === 0 || getComputedStyle(el).visibility !== 'visible') continue;
    const range = document.createRange(); range.selectNodeContents(t);
    const rects = [...range.getClientRects()].filter(r => r.width > 0 && r.height > 0);
    if (!rects.length) continue;
    const pixels = parseFloat(getComputedStyle(el).fontSize) * effectiveScale(el) * config.slideScale;
    if (pixels < config.minTextPixels - .1) add('small-text', `${pixels.toFixed(1)} output px: ${text}`);
    const ratio = contrast(el);
    if (ratio !== null && ratio < config.minContrast) add('low-contrast', `${ratio.toFixed(2)}: ${text}`);
    const region = el.closest('[data-diagram],.art');
    for (const r of rects) {
      if (outside(r, canvas)) add('outside-slide', text);
      if (region && outside(r, region.getBoundingClientRect())) add('outside-diagram', text);
      for (let p = el; p && p !== document.body; p = p.parentElement) {
        const s = getComputedStyle(p), b = p.getBoundingClientRect();
        if ((s.overflowX !== 'visible' && (r.left < b.left-1 || r.right > b.right+1)) ||
            (s.overflowY !== 'visible' && (r.top < b.top-1 || r.bottom > b.bottom+1))) add('clipped-text', text);
      }
      texts.push({ rect: r, el, text });
    }
    metrics.push({ text: text.slice(0,80), outputPixels: Number(pixels.toFixed(1)), contrast: ratio === null ? null : Number(ratio.toFixed(2)) });
  }
  for (let i=0; i<texts.length; i++) for (let j=i+1; j<texts.length; j++) {
    const a=texts[i], b=texts[j];
    if (a.el.closest('.katex') || b.el.closest('.katex')) continue;
    const w=Math.min(a.rect.right,b.rect.right)-Math.max(a.rect.left,b.rect.left);
    const h=Math.min(a.rect.bottom,b.rect.bottom)-Math.max(a.rect.top,b.rect.top);
    if (w > 2 && h > 2) add('text-overlap', `${a.text} / ${b.text}`);
  }
  for (const region of document.querySelectorAll('[data-diagram],.art')) {
    const kind = region.dataset.diagram;
    const message = region.dataset.message || region.querySelector('.diagram-caption')?.textContent.trim();
    if (config.requireDiagramDescription && !message) add('missing-diagram-message', 'Explain the conclusion conveyed by the diagram.');
    if (!kind) { warnings.push('Custom diagram: visually review its meaning and shapes.'); continue; }
    const items = [...region.querySelectorAll('[data-item]')];
    if (['comparison','flow'].includes(kind) && items.length < 2) add('diagram-items', kind + ' requires at least two labeled items.');
    for (const item of items) {
      if (!item.textContent.trim() && !item.querySelector('img[alt]:not([alt=""])')) add('unlabeled-item', kind);
    }
    for (let i=0;i<items.length;i++) for(let j=i+1;j<items.length;j++) {
      const a=items[i].getBoundingClientRect(), b=items[j].getBoundingClientRect();
      if (Math.min(a.right,b.right)-Math.max(a.left,b.left)>1 && Math.min(a.bottom,b.bottom)-Math.max(a.top,b.top)>1) add('diagram-overlap', kind);
    }
    if (kind === 'fraction' && (!region.querySelector('[data-numerator]')?.textContent.trim() || !region.querySelector('[data-denominator]')?.textContent.trim())) add('fraction-labels', 'Name both numerator and denominator.');
    if (kind === 'proportion') {
      const part=Number(region.dataset.part), total=Number(region.dataset.total);
      const track=region.querySelector('[data-track]'), fill=region.querySelector('[data-fill]');
      if (!region.hasAttribute('data-part') || !region.hasAttribute('data-total') || !Number.isFinite(part) || !Number.isFinite(total) || total<=0 || part<0 || part>total) add('proportion-values', 'Require 0 <= part <= total and total > 0.');
      else if (!track || !fill) add('proportion-bar', 'Missing track or fill.');
      else {
        const ratio=fill.getBoundingClientRect().width/track.getBoundingClientRect().width;
        if (Math.abs(ratio-part/total) > .005) add('proportion-scale', 'Bar width does not match part / total.');
      }
    }
  }
  for (const im of document.images) {
    if (!im.naturalWidth) add('missing-image', im.getAttribute('src'));
    else warnings.push('Image text/meaning needs visual review: ' + (im.getAttribute('alt') || im.getAttribute('src')));
  }
  for (const svg of document.querySelectorAll('svg')) {
    if (config.requireDiagramDescription && !svg.closest('[data-diagram],[data-decorative="true"],.katex')) add('unexplained-svg', 'Place the diagram in a described data-diagram region.');
  }
  const report = { status: issues.length ? 'failed' : 'passed', slideScale: config.slideScale, issues: [...new Map(issues.map(x=>[JSON.stringify(x),x])).values()], warnings: [...new Set(warnings)], text: metrics };
  document.body.dataset.slideAudit = JSON.stringify(report);
  if (issues.length) document.body.dataset.slideIssues = JSON.stringify(report.issues);
  document.body.dataset.ready = 'true';
});
