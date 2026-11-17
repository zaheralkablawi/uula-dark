// UULA Dark — XML engine: split a slide at red separator lines, and convert a slide to dark mode.
// Works on the XML of a single-slide package (from slide.exportAsBase64). No Office.js here.
const UD = (() => {
'use strict';
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships'
};
const PT = 12700;
const MARGIN = 28 * PT;
const GAP = 8 * PT;

// ---------- DOM helpers ----------
const kids = (el, ln, ns) => el ? [...el.children].filter(c => c.localName === ln && (!ns || c.namespaceURI === NS[ns])) : [];
const kid = (el, ln, ns) => kids(el, ln, ns)[0] || null;
const desc = (el, ns, ln) => el ? [...el.getElementsByTagNameNS(NS[ns], ln)] : [];
const mk = (doc, ns, ln) => doc.createElementNS(NS[ns], ns + ':' + ln);
const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');
const serialize = d => new XMLSerializer().serializeToString(d);
function frag(doc, xml) {
  const w = parseXml('<w xmlns:a="' + NS.a + '" xmlns:p="' + NS.p + '" xmlns:m="' + NS.m + '">' + xml + '</w>');
  return doc.importNode(w.documentElement.firstElementChild, true);
}
function placeChild(parent, el, order) {
  const idx = order.indexOf(el.localName);
  for (const c of [...parent.children]) {
    if (c === el) continue;
    const ci = order.indexOf(c.localName);
    if (ci > idx) { parent.insertBefore(el, c); return el; }
  }
  parent.appendChild(el); return el;
}
const removeKids = (el, names) => { for (const c of [...el.children]) if (names.includes(c.localName)) el.removeChild(c); };
const hasAncestor = (el, test) => { for (let e = el.parentNode; e && e.nodeType === 1; e = e.parentNode) if (test(e)) return true; return false; };
const inMath = el => hasAncestor(el, e => e.namespaceURI === NS.m);
const inFallback = el => hasAncestor(el, e => e.localName === 'Fallback');

const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
const RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
const PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'];
const SPPR_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
const LN_ORDER = ['noFill', 'solidFill', 'gradFill', 'pattFill', 'prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd', 'extLst'];
const TCPR_ORDER = ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr', 'cell3D', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'headers', 'extLst'];
const BODYPR_ORDER = ['prstTxWarp', 'noAutofit', 'normAutofit', 'spAutoFit', 'scene3d', 'sp3d', 'flatTx', 'extLst'];

// ---------- colors ----------
const PRST = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', darkBlue: '00008B', navy: '000080', gray: '808080', grey: '808080', darkGray: 'A9A9A9', dkGray: '808080', yellow: 'FFFF00', green: '008000' };
const rgb = h => [0, 2, 4].map(i => parseInt(h.substr(i, 2), 16));
const isRed = ([r, g, b]) => r > 170 && g < 100 && b < 100;
const isBlue = ([r, g, b]) => b > 140 && r < 110 && b - r > 60;
const isDark = ([r, g, b]) => Math.max(r, g, b) - Math.min(r, g, b) < 40 && Math.max(r, g, b) < 200;
const isLight = h => { const [r, g, b] = rgb(h); return 0.299 * r + 0.587 * g + 0.114 * b > 200; };

function readTheme(files) {
  const scheme = {};
  if (!files.theme) return scheme;
  const d = parseXml(files.theme);
  const cs = desc(d, 'a', 'clrScheme')[0];
  if (!cs) return scheme;
  for (const c of cs.children) {
    const s = kid(c, 'srgbClr'), y = kid(c, 'sysClr');
    scheme[c.localName] = s ? s.getAttribute('val').toUpperCase() : y ? (y.getAttribute('lastClr') || '000000').toUpperCase() : null;
  }
  scheme.tx1 = scheme.dk1; scheme.bg1 = scheme.lt1; scheme.tx2 = scheme.dk2; scheme.bg2 = scheme.lt2;
  return scheme;
}
// hex of a color element container (solidFill, lnRef, ...)
function colorOf(container, ctx) {
  for (const c of container.children) {
    const v = c.getAttribute('val');
    if (c.localName === 'srgbClr') return (v || '').toUpperCase();
    if (c.localName === 'schemeClr') return ctx.scheme[v] || null;
    if (c.localName === 'prstClr') return PRST[v] || null;
    if (c.localName === 'sysClr') return (c.getAttribute('lastClr') || (v === 'windowText' ? '000000' : 'FFFFFF')).toUpperCase();
    if (c.localName === 'scrgbClr' || c.localName === 'hslClr') return null;
  }
  return null;
}
function mapColor(hex, ctx) {
  if (!hex || !/^[0-9A-F]{6}$/.test(hex)) return null;
  const c = rgb(hex);
  if (isRed(c)) return ctx.text;
  if (isBlue(c)) return ctx.accent;
  if (isDark(c)) return ctx.text;
  return null;
}
function solid(doc, hex) { const f = mk(doc, 'a', 'solidFill'), s = mk(doc, 'a', 'srgbClr'); s.setAttribute('val', hex); f.appendChild(s); return f; }
function setSolid(parent, hex, order) { removeKids(parent, FILLS); return placeChild(parent, solid(parent.ownerDocument, hex), order); }
function recolorFill(parent, ctx, order, fallbackHex) {
  const sf = kid(parent, 'solidFill', 'a');
  if (sf) {
    const m = mapColor(colorOf(sf, ctx), ctx);
    if (m) setSolid(parent, m, order);
    return;
  }
  if (fallbackHex && !FILLS.some(n => kid(parent, n, 'a'))) setSolid(parent, fallbackHex, order);
}
// swap the color element inside a container (lnRef/fontRef) if it maps
function recolorRef(ref, ctx) {
  if (!ref) return;
  const m = mapColor(colorOf(ref, ctx), ctx);
  if (!m) return;
  for (const c of [...ref.children]) ref.removeChild(c);
  const s = mk(ref.ownerDocument, 'a', 'srgbClr'); s.setAttribute('val', m); ref.appendChild(s);
}

// ---------- package ----------
async function openPackage(b64) {
  const zip = await JSZip.loadAsync(b64, { base64: true });
  const names = Object.keys(zip.files);
  const slidePath = names.find(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const files = { slide: await zip.file(slidePath).async('string') };
  const pres = zip.file('ppt/presentation.xml');
  let slideW = 12192000, slideH = 6858000;
  if (pres) {
    const sz = desc(parseXml(await pres.async('string')), 'p', 'sldSz')[0];
    if (sz) { slideW = +sz.getAttribute('cx'); slideH = +sz.getAttribute('cy'); }
  }
  const th = names.find(n => /^ppt\/theme\/theme\d+\.xml$/.test(n));
  if (th) files.theme = await zip.file(th).async('string');
  // layout + master of this slide, for placeholders without their own position
  const relPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
  const layouts = [];
  try {
    const rels = parseXml(await zip.file(relPath).async('string'));
    const lr = desc(rels, 'rel', 'Relationship').find(r => /slideLayout$/.test(r.getAttribute('Type')));
    if (lr) {
      const lp = 'ppt/slideLayouts/' + lr.getAttribute('Target').split('/').pop();
      layouts.push(await zip.file(lp).async('string'));
      const lrels = parseXml(await zip.file(lp.replace('slideLayouts/', 'slideLayouts/_rels/') + '.rels').async('string'));
      const mr = desc(lrels, 'rel', 'Relationship').find(r => /slideMaster$/.test(r.getAttribute('Type')));
      if (mr) layouts.push(await zip.file('ppt/slideMasters/' + mr.getAttribute('Target').split('/').pop()).async('string'));
    }
  } catch (e) { /* no layout info */ }
  return { zip, slidePath, files, slideW, slideH, layouts: layouts.map(parseXml) };
}
async function buildPackage(pkg, slideXml) {
  pkg.zip.file(pkg.slidePath, slideXml);
  return pkg.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}

// ---------- shapes ----------
const SHAPE_TAGS = ['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp', 'AlternateContent', 'contentPart'];
const spTreeOf = doc => desc(doc, 'p', 'spTree')[0];
const topItems = doc => [...spTreeOf(doc).children].filter(c => SHAPE_TAGS.includes(c.localName));
// the real shape element of a top item (AlternateContent -> its Choice shape)
function core(el) {
  if (el.localName !== 'AlternateContent') return el;
  const ch = kid(el, 'Choice', 'mc') || kid(el, 'Fallback', 'mc');
  return ch ? ch.firstElementChild : el;
}
function xfrmOf(el) {
  el = core(el);
  if (!el) return null;
  if (el.localName === 'graphicFrame') return kid(el, 'xfrm', 'p');
  if (el.localName === 'grpSp') return kid(kid(el, 'grpSpPr', 'p'), 'xfrm', 'a');
  return kid(kid(el, 'spPr', 'p'), 'xfrm', 'a');
}
function phOf(el) {
  el = core(el);
  const nv = el && el.firstElementChild;
  return nv ? desc(nv, 'p', 'ph')[0] || null : null;
}
function geomFromXfrm(x) {
  if (!x) return null;
  const off = kid(x, 'off', 'a'), ext = kid(x, 'ext', 'a');
  if (!off || !ext) return null;
  return { x: +off.getAttribute('x'), y: +off.getAttribute('y'), w: +ext.getAttribute('cx'), h: +ext.getAttribute('cy') };
}
function geomOf(el, pkg) {
  const g = geomFromXfrm(xfrmOf(el));
  if (g) return g;
  const ph = phOf(el);
  if (ph) {
    const type = ph.getAttribute('type') || 'body', idx = ph.getAttribute('idx');
    for (const L of pkg.layouts) {
      for (const sp of desc(L, 'p', 'sp')) {
        const lph = phOf(sp); if (!lph) continue;
        const lt = lph.getAttribute('type') || 'body';
        if ((idx != null && lph.getAttribute('idx') === idx) || lt === type) {
          const lg = geomFromXfrm(xfrmOf(sp)); if (lg) return lg;
        }
      }
    }
  }
  return { x: 0, y: 0, w: 0, h: 0 };
}
function setGeom(el, g) {
  const c = core(el), doc = c.ownerDocument;
  let x = xfrmOf(el);
  if (!x) {
    if (c.localName === 'graphicFrame') { x = mk(doc, 'p', 'xfrm'); c.insertBefore(x, kid(c, 'graphic', 'a')); }
    else {
      const sp = kid(c, c.localName === 'grpSp' ? 'grpSpPr' : 'spPr', 'p'); if (!sp) return;
      x = mk(doc, 'a', 'xfrm'); sp.insertBefore(x, sp.firstChild);
    }
    x.appendChild(mk(doc, 'a', 'off')); x.appendChild(mk(doc, 'a', 'ext'));
  }
  const off = kid(x, 'off', 'a'), ext = kid(x, 'ext', 'a');
  // groups: keep children in place by moving chOff together with off (sizes unchanged)
  off.setAttribute('x', Math.round(g.x)); off.setAttribute('y', Math.round(g.y));
  ext.setAttribute('cx', Math.round(g.w)); ext.setAttribute('cy', Math.round(g.h));
  // also move the Fallback copy of an AlternateContent so both agree
  if (el.localName === 'AlternateContent') {
    const fb = kid(el, 'Fallback', 'mc'), fs = fb && fb.firstElementChild;
    const fx = fs && kid(kid(fs, 'spPr', 'p'), 'xfrm', 'a');
    if (fx) { kid(fx, 'off', 'a').setAttribute('x', Math.round(g.x)); kid(fx, 'off', 'a').setAttribute('y', Math.round(g.y)); }
  }
}
const textOf = el => desc(el, 'a', 't').map(t => t.textContent).join('') + desc(el, 'm', 't').map(t => t.textContent).join('');
const hasText = el => textOf(el).trim().length > 0;
const isLineGeom = c => {
  if (c.localName === 'cxnSp') return true;
  const pg = desc(kid(c, 'spPr', 'p'), 'a', 'prstGeom')[0];
  return !!pg && /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(pg.getAttribute('prst'));
};
function lineColorOf(c, ctx) {
  const ln = kid(kid(c, 'spPr', 'p'), 'ln', 'a');
  const sf = ln && kid(ln, 'solidFill', 'a');
  if (sf) return colorOf(sf, ctx);
  const ref = desc(kid(c, 'style', 'p'), 'a', 'lnRef')[0];
  return ref ? colorOf(ref, ctx) : null;
}
function fillColorOf(c, ctx) {
  const sf = kid(kid(c, 'spPr', 'p'), 'solidFill', 'a');
  if (sf) return colorOf(sf, ctx);
  if (kid(kid(c, 'spPr', 'p'), 'noFill', 'a')) return null;
  const ref = desc(kid(c, 'style', 'p'), 'a', 'fillRef')[0];
  return ref && +ref.getAttribute('idx') > 0 ? colorOf(ref, ctx) : null;
}

// classify every top-level item
function analyze(doc, pkg, ctx) {
  const W = pkg.slideW, H = pkg.slideH;
  return topItems(doc).map((el, i) => {
    const c = core(el), g = geomOf(el, pkg), ph = phOf(el);
    const it = { i, el, g, tag: c ? c.localName : el.localName, kind: 'other', drop: null };
    if (!c) return it;
    const phType = ph && ph.getAttribute('type');
    const text = textOf(c).trim();
    if (it.tag === 'pic') it.kind = 'pic';
    else if (it.tag === 'graphicFrame') it.kind = desc(c, 'a', 'tbl').length ? 'table' : 'other';
    else if ((it.tag === 'sp' || it.tag === 'cxnSp') && isLineGeom(c)) it.kind = 'line';
    else if (it.tag === 'sp' && text) it.kind = 'text';
    // red separator line (or a thin red bar)
    if (g.w >= W * 0.6 && g.h <= Math.max(H * 0.02, 6 * PT)) {
      const col = it.kind === 'line' ? lineColorOf(c, ctx) : (it.tag === 'sp' && !text ? (fillColorOf(c, ctx) || lineColorOf(c, ctx)) : null);
      if (col && isRed(rgb(col))) { it.kind = 'sep'; it.drop = 'sep'; return it; }
    }
    if (['sldNum', 'ftr', 'dt'].includes(phType)) { it.drop = 'pagenum'; return it; }
    // page number: small box with digits only, low on the slide
    if (it.kind === 'text' && /^[0-9٠-٩]{1,4}$/.test(text) && g.w < W * 0.2 && g.y > H * 0.75) { it.drop = 'pagenum'; return it; }
    // QR code: small picture, linked or near the top or square
    if (it.kind === 'pic' && Math.max(g.w, g.h) < 80 * PT) {
      const linked = !!desc(c, 'a', 'hlinkClick').length;
      const square = g.h > 0 && Math.abs(g.w / g.h - 1) < 0.15;
      if (linked || square || g.y < H * 0.25) { it.drop = 'qr'; return it; }
    }
    // light decorative card behind the text (a filled shape without text)
    if (it.tag === 'sp' && !text && it.kind !== 'line') {
      const f = fillColorOf(c, ctx);
      if (f && isLight(f) && g.w > 80 * PT && g.h > 40 * PT) { it.drop = 'card'; return it; }
    }
    return it;
  });
}

// ---------- split ----------
// returns [{ids:[item index...]}] in top-to-bottom order; one entry = no separators
function sections(items) {
  const seps = items.filter(t => t.kind === 'sep').map(t => t.g.y + t.g.h / 2).sort((a, b) => a - b);
  const keep = items.filter(t => !t.drop);
  const out = Array.from({ length: seps.length + 1 }, () => []);
  for (const t of keep) {
    const cy = t.g.y + t.g.h / 2;
    let k = 0; while (k < seps.length && cy > seps[k]) k++;
    out[k].push(t.i);
  }
  return { count: seps.length, parts: out.filter(p => p.length) };
}
// a copy of the slide holding only the given items, moved up to the top margin
function sectionXml(slideXml, pkg, ctx, ids) {
  const doc = parseXml(slideXml);
  const items = analyze(doc, pkg, ctx);
  const set = new Set(ids);
  let top = Infinity;
  for (const t of items) if (set.has(t.i)) top = Math.min(top, t.g.y);
  for (const t of items) {
    if (!set.has(t.i)) { t.el.parentNode.removeChild(t.el); continue; }
    if (isFinite(top) && top > MARGIN) setGeom(t.el, { ...t.g, y: t.g.y - top + MARGIN });
  }
  return serialize(doc);
}

// ---------- dark ----------
function styleRPr(rPr, ctx, opt) {
  const math = inMath(rPr);
  recolorFill(rPr, ctx, RPR_ORDER, ctx.text);
  if (opt.size) {
    const s = rPr.getAttribute('sz');
    rPr.setAttribute('sz', s && +s > 1500 ? ctx.title : ctx.body);
  }
  if (!math && opt.font) {
    for (const ln of ['latin', 'cs']) {
      let f = kid(rPr, ln, 'a');
      if (!f) { f = placeChild(rPr, mk(rPr.ownerDocument, 'a', ln), RPR_ORDER); f.setAttribute('typeface', ctx.font); continue; }
      const tf = f.getAttribute('typeface') || '';
      if (tf === 'UULA Sans' || tf === '' || /^\+(mn|mj)-/.test(tf)) f.setAttribute('typeface', ctx.font);
    }
  }
}
function ensure(parent, ns, ln, order, first) {
  let e = kid(parent, ln, ns);
  if (!e) { e = mk(parent.ownerDocument, ns, ln); if (order) placeChild(parent, e, order); else if (first) parent.insertBefore(e, parent.firstChild); else parent.appendChild(e); }
  return e;
}
function styleTxBody(tx, ctx, inTable) {
  const doc = tx.ownerDocument;
  const arabic = /[؀-ۿ]/.test(textOf(tx));
  const bp = kid(tx, 'bodyPr', 'a');
  if (bp && !inTable) {
    bp.setAttribute('anchor', 't'); bp.setAttribute('wrap', 'square');
    removeKids(bp, ['noAutofit', 'normAutofit', 'spAutoFit']);
    placeChild(bp, mk(doc, 'a', 'spAutoFit'), BODYPR_ORDER);
  }
  const lst = kid(tx, 'lstStyle', 'a');
  if (lst) for (const d of desc(lst, 'a', 'defRPr')) styleRPr(d, ctx, { size: true, font: true });
  for (const p of kids(tx, 'p', 'a')) {
    let pPr = kid(p, 'pPr', 'a');
    if (!pPr) { pPr = mk(doc, 'a', 'pPr'); p.insertBefore(pPr, p.firstChild); }
    // Arabic box: right aligned; an English-only line inside it stays LTR so its punctuation is right
    const arP = /[؀-ۿ]/.test(textOf(p));
    pPr.setAttribute('algn', arabic ? 'r' : 'l'); pPr.setAttribute('rtl', arabic && (arP || !textOf(p).trim()) ? '1' : '0');
    removeKids(pPr, ['lnSpc', 'spcAft']);
    placeChild(pPr, frag(doc, '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>'), PPR_ORDER);
    placeChild(pPr, frag(doc, '<a:spcAft><a:spcPts val="0"/></a:spcAft>'), PPR_ORDER);
    const d = kid(pPr, 'defRPr', 'a'); if (d) styleRPr(d, ctx, { size: true, font: true });
    // runs, fields, line breaks
    for (const r of [...kids(p, 'r', 'a'), ...kids(p, 'fld', 'a'), ...kids(p, 'br', 'a')]) {
      let rPr = kid(r, 'rPr', 'a');
      if (!rPr) { rPr = mk(doc, 'a', 'rPr'); r.insertBefore(rPr, r.firstChild); }
      styleRPr(rPr, ctx, { size: true, font: true });
    }
    let e = kid(p, 'endParaRPr', 'a');
    if (!e) { e = mk(doc, 'a', 'endParaRPr'); p.appendChild(e); }
    styleRPr(e, ctx, { size: true, font: true });
    // math inside the paragraph (a14:m): colors + sizes, keep the math font
    for (const mr of desc(p, 'm', 'r')) {
      let rPr = kid(mr, 'rPr', 'a');
      if (!rPr) { rPr = mk(doc, 'a', 'rPr'); mr.insertBefore(rPr, mr.firstChild); }
      styleRPr(rPr, ctx, { size: true, font: false });
    }
    for (const cp of desc(p, 'm', 'ctrlPr')) for (const rPr of kids(cp, 'rPr', 'a')) styleRPr(rPr, ctx, { size: true, font: false });
    if (arabic) for (const op of desc(p, 'm', 'oMathPara')) {
      let pp = kid(op, 'oMathParaPr', 'm');
      if (!pp) { pp = mk(doc, 'm', 'oMathParaPr'); op.insertBefore(pp, op.firstChild); }
      let jc = kid(pp, 'jc', 'm');
      if (!jc) { jc = mk(doc, 'm', 'jc'); pp.appendChild(jc); }
      jc.setAttributeNS(NS.m, 'm:val', 'right');
    }
  }
  return arabic;
}
function styleShapeProps(c, ctx) {
  const spPr = kid(c, 'spPr', 'p'); if (!spPr) return;
  const doc = c.ownerDocument;
  if (!isLineGeom(c) && c.localName === 'sp') { removeKids(spPr, FILLS); placeChild(spPr, mk(doc, 'a', 'noFill'), SPPR_ORDER); }
  const ln = kid(spPr, 'ln', 'a'); if (ln) recolorFill(ln, ctx, LN_ORDER);
  const st = kid(c, 'style', 'p');
  if (st) { recolorRef(kid(st, 'lnRef', 'a'), ctx); recolorRef(kid(st, 'fontRef', 'a'), ctx); if (c.localName === 'sp' && !isLineGeom(c)) { const fr = kid(st, 'fillRef', 'a'); if (fr) fr.setAttribute('idx', '0'); } }
}
function styleTable(tbl, ctx) {
  const doc = tbl.ownerDocument;
  for (const tc of desc(tbl, 'a', 'tc')) {
    const tx = kid(tc, 'txBody', 'a'); if (tx) styleTxBody(tx, ctx, true);
    let tcPr = kid(tc, 'tcPr', 'a');
    if (!tcPr) { tcPr = mk(doc, 'a', 'tcPr'); const ext = kid(tc, 'extLst', 'a'); tc.insertBefore(tcPr, ext); }
    removeKids(tcPr, ['lnL', 'lnR', 'lnT', 'lnB', ...FILLS]);
    for (const s of ['lnL', 'lnR', 'lnT', 'lnB'])
      placeChild(tcPr, frag(doc, '<a:' + s + ' w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + ctx.text + '"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:' + s + '>'), TCPR_ORDER);
    placeChild(tcPr, solid(doc, ctx.bg), TCPR_ORDER);
  }
}
function duotone(pic, ctx) {
  const blip = desc(pic, 'a', 'blip')[0]; if (!blip) return false;
  if (kid(blip, 'duotone', 'a')) return false;
  const d = frag(blip.ownerDocument, '<a:duotone><a:srgbClr val="' + ctx.text + '"/><a:srgbClr val="' + ctx.bg + '"/></a:duotone>');
  blip.insertBefore(d, kid(blip, 'extLst', 'a'));
  return true;
}

// estimated height (EMU) of a text body at a given width
function estimateText(tx, widthEmu, ctx) {
  let h = 7.2 * PT;
  for (const p of kids(tx, 'p', 'a')) {
    let sz = 0;
    for (const r of desc(p, 'a', 'rPr')) sz = Math.max(sz, +(r.getAttribute('sz') || 0));
    const e = kid(p, 'endParaRPr', 'a'); if (!sz && e) sz = +(e.getAttribute('sz') || 0);
    sz = (sz || +ctx.body) / 100; // pt
    const len = textOf(p).length;
    const cpl = Math.max(1, Math.floor((widthEmu / PT - 14.4) / (sz * 0.5)));
    const lines = Math.max(1, Math.ceil(len / cpl)) + Math.max(0, desc(p, 'm', 'f').length ? 1 : 0);
    h += lines * sz * 1.2 * 1.5 * PT;
  }
  return h;
}
function estimateTable(tbl, widthEmu, ctx) {
  const cols = desc(tbl, 'a', 'gridCol').map(c => +c.getAttribute('w'));
  const total = cols.reduce((a, b) => a + b, 0) || 1;
  let h = 0;
  for (const tr of desc(tbl, 'a', 'tr')) {
    let rh = +(tr.getAttribute('h') || 0), ci = 0;
    for (const tc of kids(tr, 'tc', 'a')) {
      const span = +(tc.getAttribute('gridSpan') || 1);
      let w = 0; for (let k = 0; k < span; k++) w += cols[ci + k] || 0; ci += span;
      const tx = kid(tc, 'txBody', 'a');
      if (tx) rh = Math.max(rh, estimateText(tx, w / total * widthEmu, ctx) + 1.6 * 2 * PT);
    }
    h += rh;
  }
  return h;
}
function resizeTable(el, newW) {
  const c = core(el), tbl = desc(c, 'a', 'tbl')[0];
  const cols = desc(tbl, 'a', 'gridCol');
  const total = cols.reduce((a, g) => a + +g.getAttribute('w'), 0) || 1;
  for (const g of cols) g.setAttribute('w', Math.round(+g.getAttribute('w') / total * newW));
}

// the whole dark conversion of one slide XML; returns {xml, plan, report}
// ctx: {bg, text, accent, body:'3600', title:'4800', font, images:true}
function darkXml(slideXml, pkg, ctxIn) {
  const doc = parseXml(slideXml);
  const ctx = { ...ctxIn, scheme: readTheme(pkg.files) };
  const items = analyze(doc, pkg, ctx);
  const report = { removed: 0, duotone: 0 };
  // remove separators, page numbers, QR codes, decorative cards
  for (const t of items) if (t.drop) { t.el.parentNode.removeChild(t.el); report.removed++; }
  const keep = items.filter(t => !t.drop);
  // background
  const cSld = kid(doc.documentElement, 'cSld', 'p');
  const oldBg = kid(cSld, 'bg', 'p'); if (oldBg) cSld.removeChild(oldBg);
  cSld.insertBefore(frag(doc, '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + ctx.bg + '"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'), kid(cSld, 'spTree', 'p'));
  // shapes, text, tables (skip the Fallback copies of equations)
  const tree = spTreeOf(doc);
  for (const c of [...desc(tree, 'p', 'sp'), ...desc(tree, 'p', 'cxnSp')]) if (!inFallback(c)) styleShapeProps(c, ctx);
  for (const tx of desc(tree, 'p', 'txBody')) if (!inFallback(tx)) styleTxBody(tx, ctx, false);
  for (const tbl of desc(tree, 'a', 'tbl')) styleTable(tbl, ctx);
  // big pictures: white lines on the dark background
  const W = pkg.slideW, H = pkg.slideH;
  if (ctx.images !== false) for (const t of keep) if (t.kind === 'pic' && Math.max(t.g.w, t.g.h) >= 80 * PT) { if (duotone(core(t.el), ctx)) report.duotone++; }
  // first layout guess in the XML (the add-in refines it with real heights)
  const plan = layout(keep, pkg, ctx);
  for (const s of plan.items) {
    const t = keep.find(k => k.i === s.i);
    if (t.kind === 'table') resizeTable(t.el, s.w);
    setGeom(t.el, { x: s.x, y: s.y, w: s.w, h: s.h });
  }
  // plan indexes refer to the order of the remaining top items
  const order = topItems(doc);
  for (const s of plan.items) s.idx = order.indexOf(keep.find(k => k.i === s.i).el);
  return { xml: serialize(doc), plan, report };
}
// layout: full-width stack, or title + picture (left) + texts/tables (right)
function layout(keep, pkg, ctx) {
  const W = pkg.slideW - 2 * MARGIN, H = pkg.slideH;
  const sorted = [...keep].sort((a, b) => a.g.y - b.g.y || b.g.x - a.g.x);
  const pics = sorted.filter(t => t.kind === 'pic' && Math.max(t.g.w, t.g.h) >= 80 * PT);
  const rest = sorted.filter(t => !pics.includes(t));
  const items = [];
  const est = (t, w) => {
    const c = core(t.el);
    if (t.kind === 'text') return estimateText(kid(c, 'txBody', 'p'), w, ctx);
    if (t.kind === 'table') return estimateTable(desc(c, 'a', 'tbl')[0], w, ctx);
    return t.g.h;
  };
  const twoCol = pics.length > 0 && rest.some(t => t.kind === 'text' || t.kind === 'table');
  if (!twoCol) {
    let y = MARGIN;
    for (const t of sorted) {
      let w, x, h;
      if (t.kind === 'text' || t.kind === 'table') { w = W; x = MARGIN; h = est(t, w); }
      else {
        const k = Math.min(1, W / (t.g.w || 1), (H - 2 * MARGIN) / (t.g.h || 1));
        w = t.g.w * k; h = t.g.h * k; x = MARGIN + (W - w) / 2;
      }
      items.push({ i: t.i, kind: t.kind, col: 'full', x, y, w, h });
      y += h + GAP;
    }
    return { mode: 'stack', items, bottom: y - GAP };
  }
  const half = (W - 12 * PT) / 2;
  let y0 = MARGIN;
  const first = rest[0];
  let title = null;
  if (first && first.kind === 'text' && first.g.y <= pics[0].g.y) {
    title = first;
    const h = est(first, W);
    items.push({ i: first.i, kind: 'text', col: 'title', x: MARGIN, y: MARGIN, w: W, h });
    y0 = MARGIN + h + GAP;
  }
  // left column: pictures (Office x grows to the right, so "left" is x = MARGIN)
  let yl = y0;
  for (const t of pics) {
    const k = Math.min(half / (t.g.w || 1), (H - MARGIN - yl) / (t.g.h || 1));
    const w = t.g.w * k, h = t.g.h * k;
    items.push({ i: t.i, kind: 'pic', col: 'left', x: MARGIN, y: yl, w, h });
    yl += h + GAP;
  }
  let yr = y0;
  for (const t of rest) {
    if (t === title) continue;
    const big = t.kind === 'text' || t.kind === 'table';
    const w = big ? half : Math.min(half, t.g.w), h = big ? est(t, w) : t.g.h * (w / (t.g.w || 1));
    items.push({ i: t.i, kind: t.kind, col: 'right', x: MARGIN + half + 12 * PT, y: yr, w, h });
    yr += h + GAP;
  }
  return { mode: 'two', items, bottom: Math.max(yl, yr) - GAP };
}

return { NS, PT, MARGIN, GAP, openPackage, buildPackage, parseXml, serialize, analyze, sections, sectionXml, darkXml, readTheme, mapColor, rgb };
})();
if (typeof module !== 'undefined') module.exports = UD;
