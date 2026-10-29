// Leitura detalhada do Word (.docx): estilos, listas, parágrafos, tabelas e espaçamento.
// Adaptado do projeto curadoria-moodle.

const kids = (el, n) => el ? [...el.children].filter(c => c.nodeName === n) : [];
const kid = (el, n) => kids(el, n)[0];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const norm = s => String(s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();

function onOff(rPr, tag) {
    const e = kid(rPr, tag);
    if (!e) return false;
    const v = e.getAttribute('w:val');
    return !(v === '0' || v === 'false' || v === 'none');
}

// espaçamento "simples" do Word para Arial/Calibri ≈ 1,15 × tamanho da fonte
const FONT_SINGLE = 1.15;

// Estado do documento que está sendo lido (reiniciado a cada .docx)
let STY = { defaults: {}, styles: {}, defaultId: null, fontPt: 12 };
let NUMB = { nums: {}, abs: {}, counters: {} };
let DOCX_RELS = {};

async function carregarFormatacaoDocx(zip, relMap) {
    const ler = async path => { try { return await zip.file(path)?.async('string'); } catch (e) { return null; } };
    loadStyles(await ler('word/styles.xml'));
    loadNumbering(await ler('word/numbering.xml'));
    DOCX_RELS = relMap;
    MATHS.length = 0;
}

/* Propriedades de parágrafo: alinhamento, entrelinha, espaço antes/depois, recuo e lista */
function pPrProps(pPr) {
    const o = {};
    if (!pPr) return o;
    const jc = kid(pPr, 'w:jc');
    if (jc) o.jc = jc.getAttribute('w:val');
    const sp = kid(pPr, 'w:spacing');
    if (sp) {
        const g = a => sp.getAttribute('w:' + a);
        if (g('line') != null) { o.line = +g('line'); o.lineRule = g('lineRule') || 'auto'; }
        if (g('before') != null) o.before = +g('before');
        if (g('after') != null) o.after = +g('after');
        if (g('beforeAutospacing') === '1') o.before = 280;
        if (g('afterAutospacing') === '1') o.after = 280;
    }
    const ind = kid(pPr, 'w:ind');
    if (ind) {
        const g = a => ind.getAttribute('w:' + a);
        if (g('firstLine') != null) { o.firstLine = +g('firstLine'); o.hanging = 0; }
        if (g('hanging') != null) { o.hanging = +g('hanging'); o.firstLine = 0; }
        const l = g('left') ?? g('start');
        if (l != null) o.left = +l;
    }
    if (kid(pPr, 'w:contextualSpacing')) o.contextual = onOff(pPr, 'w:contextualSpacing');
    const np = kid(pPr, 'w:numPr');
    if (np) {
        const id = kid(np, 'w:numId')?.getAttribute('w:val');
        const lv = kid(np, 'w:ilvl')?.getAttribute('w:val');
        if (id != null) o.numId = id;
        if (lv != null) o.ilvl = +lv;
    }
    return o;
}

/* Estilos do Word: padrões do documento + estilo de parágrafo (com herança basedOn) */
function loadStyles(xml) {
    STY = { defaults: {}, styles: {}, defaultId: null, fontPt: 12 };
    if (!xml) return;
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const dd = d.getElementsByTagName('w:pPrDefault')[0];
    if (dd) STY.defaults = pPrProps(kid(dd, 'w:pPr'));
    const rd = d.getElementsByTagName('w:rPrDefault')[0];
    const sz = rd && rd.getElementsByTagName('w:sz')[0];
    if (sz) STY.fontPt = +sz.getAttribute('w:val') / 2;
    for (const st of d.getElementsByTagName('w:style')) {
        const ty = st.getAttribute('w:type');
        if (ty !== 'paragraph' && ty !== 'numbering') continue;
        const id = st.getAttribute('w:styleId');
        STY.styles[id] = { basedOn: kid(st, 'w:basedOn')?.getAttribute('w:val'), props: pPrProps(kid(st, 'w:pPr')) };
        const ssz = kid(kid(st, 'w:rPr'), 'w:sz');
        if (ssz) STY.styles[id].fontPt = +ssz.getAttribute('w:val') / 2;
        if (ty === 'paragraph' && st.getAttribute('w:default') === '1') STY.defaultId = id;
    }
}

function styleProps(id, seen = new Set()) {
    const st = STY.styles[id];
    if (!st || seen.has(id)) return {};
    seen.add(id);
    return Object.assign({}, st.basedOn ? styleProps(st.basedOn, seen) : {}, st.props);
}

/* Listas do Word (numbering.xml): marcadores e numeração de qualquer tipo, com contadores por nível */
const BULLET_MAP = { 0xB7: '•', 0xA7: '▪', 0xFC: '✓', 0xD8: '➢', 0x76: '❖', 0x6E: '■', 0x71: '❑', 0x6C: '●', 0xA8: '○', 0xE0: '➢', 0xF0: '➢', 0x22: '∎', 0x3E: '➢', 0x2D: '–', 0x6F: '○', 0x77: '◆', 0xA0: '▪', 0x9F: '•', 0xD7: '▪', 0xFB: '✗' };

function loadNumbering(xml) {
    NUMB = { nums: {}, abs: {}, counters: {}, links: {} };
    if (!xml) return;
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    for (const a of d.getElementsByTagName('w:abstractNum')) {
        const lv = {};
        for (const l of kids(a, 'w:lvl')) {
            const il = +l.getAttribute('w:ilvl');
            const rf = kid(kid(l, 'w:rPr'), 'w:rFonts');
            const ind = kid(kid(l, 'w:pPr'), 'w:ind');
            lv[il] = {
                fmt: kid(l, 'w:numFmt')?.getAttribute('w:val') || 'decimal',
                text: kid(l, 'w:lvlText')?.getAttribute('w:val') ?? '%' + (il + 1) + '.',
                start: +(kid(l, 'w:start')?.getAttribute('w:val') || 1),
                font: (rf && (rf.getAttribute('w:ascii') || rf.getAttribute('w:hAnsi'))) || '',
                left: ind ? +(ind.getAttribute('w:left') || ind.getAttribute('w:start') || 0) : null,
                hang: ind ? +(ind.getAttribute('w:hanging') || 0) : null
            };
        }
        NUMB.abs[a.getAttribute('w:abstractNumId')] = lv;
        const link = kid(a, 'w:numStyleLink')?.getAttribute('w:val');
        if (link) NUMB.links[a.getAttribute('w:abstractNumId')] = link;
    }
    for (const n of d.getElementsByTagName('w:num')) {
        const ov = {};
        for (const o of kids(n, 'w:lvlOverride')) {
            const so = kid(o, 'w:startOverride');
            if (so) ov[+o.getAttribute('w:ilvl')] = +so.getAttribute('w:val');
        }
        NUMB.nums[n.getAttribute('w:numId')] = { abs: kid(n, 'w:abstractNumId')?.getAttribute('w:val'), ov };
    }
}

function roman(n) {
    const v = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let r = '';
    for (const [a, b] of v) while (n >= a) { r += b; n -= a; }
    return r;
}

function letters(n) {
    let r = '';
    while (n > 0) { n--; r = String.fromCharCode(97 + n % 26) + r; n = Math.floor(n / 26); }
    return r;
}

function fmtNum(n, f) {
    switch (f) {
        case 'lowerLetter': return letters(n);
        case 'upperLetter': return letters(n).toUpperCase();
        case 'lowerRoman': return roman(n);
        case 'upperRoman': return roman(n).toUpperCase();
        case 'decimalZero': return (n < 10 ? '0' : '') + n;
        case 'ordinal': return n + 'º';
        case 'none': return '';
        default: return String(n);
    }
}

function listMarker(numId, ilvl) {
    const nm = NUMB.nums[numId];
    if (!nm || numId === '0') return null;
    let lv = NUMB.abs[nm.abs];
    // lista definida por "estilo de lista": segue o vínculo até a definição real
    if ((!lv || !Object.keys(lv).length) && NUMB.links[nm.abs]) {
        const st = styleProps(NUMB.links[nm.abs]);
        const nm2 = st.numId && NUMB.nums[st.numId];
        if (nm2) lv = NUMB.abs[nm2.abs];
    }
    if (!lv || !Object.keys(lv).length) return null;
    const L = lv[ilvl] || lv[0];
    if (!L) return null;
    const key = nm.abs + ':' + numId;
    const c = NUMB.counters[key] = NUMB.counters[key] || {};
    const startOf = i => nm.ov[i] ?? (lv[i]?.start ?? 1);
    c[ilvl] = (c[ilvl] == null ? startOf(ilvl) : c[ilvl] + 1);
    for (const k of Object.keys(c)) if (+k > ilvl) delete c[k];
    let mk;
    if (L.fmt === 'bullet') {
        const ch = L.text || '•';
        const code = ch.charCodeAt(0);
        const f = L.font.toLowerCase();
        if (/symbol|wingdings/.test(f) || code >= 0xF000) mk = BULLET_MAP[code >= 0xF000 ? code - 0xF000 : code] || '•';
        else if (ch === 'o' && /courier/.test(f)) mk = '○';
        else if (ch === '') mk = '•';
        else mk = ch;
    } else {
        mk = L.text.replace(/%(\d)/g, (_, d) => {
            const i = +d - 1;
            return fmtNum(c[i] ?? startOf(i), (lv[i] || L).fmt);
        });
    }
    const left = L.left != null ? L.left / 567 : (ilvl + 1) * 0.63 + 0.64;
    const hang = L.hang != null ? L.hang / 567 : 0.63;
    // listas que o Moodle sabe mostrar sozinho (•, ○, ▪, 1., a., A., i., I.) viram <ul>/<ol>
    const OLT = { decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I' };
    let native = null;
    if (L.fmt === 'bullet') {
        const css = { '•': 'disc', '●': 'disc', '○': 'circle', '◦': 'circle', '▪': 'square', '■': 'square' }[mk];
        if (css) native = { tag: 'ul', css };
    } else if (OLT[L.fmt] && new RegExp('^%' + (ilvl + 1) + '\\.$').test(L.text)) {
        native = { tag: 'ol', type: OLT[L.fmt], start: c[ilvl] };
    }
    return { mk, left, hang, bullet: L.fmt === 'bullet', numId, ilvl, native };
}

// Só links externos viram <a href> (nada de javascript: ou caminhos locais)
function linkDoHyperlink(el) {
    const target = DOCX_RELS[el.getAttribute('r:id')];
    return target && /^(https?:|mailto:)/i.test(target) ? target : null;
}

const IGNORAR_NO_PARAGRAFO = new Set(['w:pPr', 'w:del', 'w:moveFrom', 'w:proofErr', 'w:bookmarkStart', 'w:bookmarkEnd', 'mc:Fallback']);

/* Parágrafo do Word -> texto, HTML (negrito, itálico, sobrescrito, links, símbolos, imagens, equações) e formatação */
function parsePara(p) {
    let text = '', html = '', curB = '', ole = 0;
    const imgs = [], maths = [], bold = [], colored = [];
    // percorre o parágrafo na ordem: trechos de texto, equações e links
    const runs = [];
    (function walk(el, href) {
        for (const c of el.children) {
            const n = c.nodeName;
            if (n === 'w:r' || n === 'm:oMath' || n === 'm:oMathPara') runs.push({ r: c, href });
            else if (n === 'w:hyperlink') walk(c, linkDoHyperlink(c) || href);
            else if (IGNORAR_NO_PARAGRAFO.has(n)) continue;
            else walk(c, href);
        }
    })(p, null);

    for (const { r, href } of runs) {
        let h = '';
        if (r.nodeName !== 'w:r') {
            // equação: guarda o LaTeX para virar imagem depois; no HTML fica um marcador com a versão em texto
            const o = omml(r);
            let tex = '';
            try { tex = omml2tex(r); } catch (e) { tex = ''; }
            const id = MATHS.length;
            MATHS.push({ tex, t: o.t, h: o.h, display: r.nodeName === 'm:oMathPara' });
            maths.push(id);
            text += o.t;
            h = `<span data-math="${id}">${o.h}</span>`;
        } else {
            const rPr = kid(r, 'w:rPr');
            const b = onOff(rPr, 'w:b'), i = onOff(rPr, 'w:i'), u = onOff(rPr, 'w:u');
            const va = kid(rPr, 'w:vertAlign')?.getAttribute('w:val');
            const col = kid(rPr, 'w:color')?.getAttribute('w:val');
            const fmt = s => {
                if (va === 'superscript') s = `<sup>${s}</sup>`;
                else if (va === 'subscript') s = `<sub>${s}</sub>`;
                if (u) s = `<u>${s}</u>`;
                if (i) s = `<i>${s}</i>`;
                if (b) s = `<b>${s}</b>`;
                return s;
            };
            for (const c of r.children) {
                const n = c.nodeName;
                if (n === 'w:t') {
                    const t = c.textContent;
                    text += t;
                    h += fmt(esc(t));
                    if (t.trim() && !href && col && !/^(auto|000000)$/i.test(col) && !colored.includes(t.trim())) colored.push(t.trim());
                    if (b) curB += t;
                    else if (curB) { bold.push(curB.trim()); curB = ''; }
                } else if (n === 'w:sym') {
                    const ch = symChar(c);
                    text += ch;
                    h += fmt(esc(ch));
                } else if (n === 'w:noBreakHyphen') {
                    text += '-'; h += '-';
                } else if (n === 'w:object') {
                    ole++;   // objeto antigo (Equation 3.0 etc.): não dá para converter
                } else if (n === 'w:br' || n === 'w:cr') {
                    text += '\n'; h += '<br>';
                } else if (n === 'w:tab') {
                    text += ' '; h += ' ';
                } else if (n === 'w:drawing' || n === 'mc:AlternateContent' || n === 'w:pict') {
                    const blip = c.getElementsByTagName('a:blip')[0];
                    const vimg = c.getElementsByTagName('v:imagedata')[0];
                    const rid = blip ? blip.getAttribute('r:embed') : (vimg ? vimg.getAttribute('r:id') : null);
                    if (rid) {
                        const ext = c.getElementsByTagName('wp:extent')[0];
                        let cx = ext ? +ext.getAttribute('cx') : 0;
                        if (!cx && vimg) {
                            const m = (vimg.parentNode?.getAttribute('style') || '').match(/width:\s*([\d.]+)pt/);
                            if (m) cx = +m[1] * 12700;
                        }
                        imgs.push({ rid, cx });
                        h += `<img data-rid="${esc(rid)}" data-cx="${Math.round(cx) || 0}">`;
                    }
                }
            }
        }
        if (h) html += href ? `<a href="${esc(href)}">${h}</a>` : h;
    }
    if (curB.trim()) bold.push(curB.trim());

    const pPr = kid(p, 'w:pPr');
    const styleId = kid(pPr, 'w:pStyle')?.getAttribute('w:val') || STY.defaultId;
    const P = Object.assign({}, STY.defaults, styleProps(styleId), pPrProps(pPr));
    // tamanho da fonte do parágrafo (para converter pontos em "em")
    let fpt = STY.styles[styleId]?.fontPt || STY.fontPt || 12;
    const rsz = p.getElementsByTagName('w:sz')[0];
    if (rsz) fpt = +rsz.getAttribute('w:val') / 2;
    const lineEm = P.line == null ? FONT_SINGLE : (P.lineRule === 'auto' ? (P.line / 240) * FONT_SINGLE : Math.max(FONT_SINGLE, (P.line / 20) / fpt));
    const tw2em = t => (t || 0) / 20 / fpt;
    let indentCm = P.firstLine ? P.firstLine / 567 : (P.hanging ? -P.hanging / 567 : 0);
    let leftCm = (P.left || 0) / 567, list = null;
    if (P.numId != null && P.numId !== '0' && (norm(text) || imgs.length)) {
        const lm = listMarker(P.numId, P.ilvl || 0);
        if (lm && lm.mk) {
            list = lm;
            const ownInd = pPrProps(pPr);
            leftCm = ownInd.left != null ? ownInd.left / 567 : lm.left;
            const hg = ownInd.hanging != null ? ownInd.hanging / 567 : lm.hang;
            indentCm = -hg;
        }
    }
    return {
        k: 'p', t: text, html, imgs, maths, ole, list, bold, colored,
        jc: P.jc, lineEm, beforeEm: tw2em(P.before), afterEm: tw2em(P.after),
        indentCm, leftCm, style: styleId, contextual: !!P.contextual
    };
}

/* Tabela dentro da questão -> tabela HTML (células mescladas, alinhamento, cor de fundo) */
function parseTable(tbl) {
    const rows = [], grid = [];
    const open = {};   // open[col] = célula que iniciou mesclagem vertical
    for (const tr of kids(tbl, 'w:tr')) {
        const r = [], cells = [];
        let col = 0;
        for (const tc of kids(tr, 'w:tc')) {
            const tcPr = kid(tc, 'w:tcPr');
            const span = +(kid(tcPr, 'w:gridSpan')?.getAttribute('w:val') || 1);
            const vm = kid(tcPr, 'w:vMerge');
            const vmv = vm ? (vm.getAttribute('w:val') || 'continue') : null;
            const ps = kids(tc, 'w:p').map(parsePara);
            const txt = norm(ps.map(p => p.t).join(' '));
            if (vmv === 'continue' && open[col]) { open[col].rowspan++; r.push(''); col += span; continue; }
            const al = { both: 'justify', center: 'center', right: 'right', end: 'right' }[ps.find(p => norm(p.t))?.jc || ps[0]?.jc] || 'left';
            const fill = kid(tcPr, 'w:shd')?.getAttribute('w:fill');
            const va = kid(tcPr, 'w:vAlign')?.getAttribute('w:val');
            const cell = {
                span, rowspan: 1, al,
                fill: fill && /^[0-9A-F]{6}$/i.test(fill) && !/^FFFFFF$/i.test(fill) ? '#' + fill : null,
                va: va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top',
                html: ps.map(p => (p.list ? `<span class="marcador">${esc(p.list.mk)}</span> ` : '') + p.html).filter(x => x !== '').join('<br>')
            };
            if (vmv === 'restart') open[col] = cell; else delete open[col];
            cells.push(cell);
            r.push(txt);
            col += span;
        }
        rows.push(r);
        grid.push(cells);
    }
    const html = `<table>${grid.map(cells => `<tr>${cells.map(c => `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''} style="border:1px solid #000;padding:4px 8px;text-align:${c.al};vertical-align:${c.va}${c.fill ? `;background-color:${c.fill}` : ''}">${c.html || '&nbsp;'}</td>`).join('')}</tr>`).join('')}</table>`;
    return { k: 'table', rows, html };
}

/* Conteúdo de uma célula da questão: parágrafos e tabelas, na ordem do Word */
function lerConteudo(el, out) {
    for (const ch of el.children) {
        const n = ch.nodeName;
        if (n === 'w:p') out.push(parsePara(ch));
        else if (n === 'w:tbl') out.push(parseTable(ch));
        else if (n === 'w:sdt') { const c = kid(ch, 'w:sdtContent'); if (c) lerConteudo(c, out); }
        else if (n === 'w:customXml') lerConteudo(ch, out);
    }
    return out;
}

const paragrafoVazio = e => e.k === 'p' && !norm(e.t) && !e.imgs.length && !e.maths.length;
const textoItens = items => norm(items.map(e => e.k === 'p' ? e.t : e.rows.flat().join(' ')).join(' '));
const temConteudo = items => items.some(e => e.k === 'table' || !paragrafoVazio(e));

/* Títulos e fontes de figura, itens digitados e referências bibliográficas */
const CAP_TITLE = /^(Figura|Fig\.|Gr[áa]fico|Quadro|Imagem|Tabela|Esquema|Mapa|Fotografia|Ilustra[çc][ãa]o)\s*\d*\s*[–\-—:.]/i;
const CAP_SRC = /^Fonte\s*:/i;
const IS_ITEM = /^(\(\s*\)|[IVX]{1,4}\s*[\.\)–-]|[a-eA-E]\)|\d{1,2}\s*[\.\)])\s*/;
const REF_HEAD = /^(Refer[êe]ncias?|Fonte|Bibliografia)\b/i;

function looksCitation(t) {
    return /^[A-ZÀ-Ý][A-ZÀ-Ý'’\-]{1,}(\s+(de|da|do|dos|das|e|[A-ZÀ-Ý'’\-]+))*\s*[,\.–-]/.test(t)
        || /\bdoi\b|Dispon[íi]vel em|\bv\.\s*\d|\bp\.\s*\d|\bed\.\s/i.test(t);
}

function refFlags(paras) {
    let inRef = false;
    return paras.map(p => {
        const t = norm(p.t.replace(/\n/g, ' '));
        if (!t) return inRef;
        if (REF_HEAD.test(t)) { inRef = true; return true; }
        if (looksCitation(t) && (inRef || /^[A-ZÀ-Ý][A-ZÀ-Ý'’\-]+,\s/.test(t))) { inRef = true; return true; }
        inRef = false;
        return false;
    });
}

/* Monta o HTML do enunciado com o espaçamento, a entrelinha, o recuo e o alinhamento do Word.
   Imagens e equações ficam como marcadores (data-rid / data-math) para serem resolvidos depois. */
function montarEnunciado(items, opts = {}) {
    items = items.slice();
    while (items.length && paragrafoVazio(items[0])) items.shift();
    while (items.length && paragrafoVazio(items[items.length - 1])) items.pop();

    const blocks = [];
    let pend = 0, prevAfter = 0, prevPara = null;
    for (const e of items) {
        if (paragrafoVazio(e)) {
            pend += (e.lineEm || FONT_SINGLE) + (e.beforeEm || 0) + (e.afterEm || 0);
            prevAfter = 0;
            continue;
        }
        let gap, blk;
        if (e.k === 'table') {
            gap = pend + prevAfter;
            blk = { table: true, html: e.html, text: '[TABELA]' };
        } else {
            const same = prevPara && e.contextual && prevPara.contextual && prevPara.style === e.style && !pend;
            gap = pend + (same ? 0 : prevAfter + (e.beforeEm || 0));
            // quebras de linha (Shift+Enter) no começo ou no fim do parágrafo viram espaço, como no Word
            let hh = e.html;
            const lead = (hh.match(/^(\s*<br>)+/) || [''])[0].split('<br>').length - 1;
            const trail = (hh.match(/(<br>\s*)+$/) || [''])[0].split('<br>').length - 1;
            hh = hh.replace(/^(\s*<br>)+/, '').replace(/(<br>\s*)+$/, '');
            gap += lead * (e.lineEm || FONT_SINGLE);
            const onlyMath = e.maths.length && !norm(hh.replace(/<span data-math[\s\S]*?<\/span>/g, '').replace(/<[^>]+>/g, ''));
            blk = {
                para: e, html: hh, onlyMath,
                text: norm(e.t.replace(/\n/g, ' ')) + (e.imgs.length ? ' [IMG]' : '') + (onlyMath ? ' [EQ]' : ''),
                trail: trail * (e.lineEm || FONT_SINGLE)
            };
        }
        const last = blocks[blocks.length - 1];
        if (last) {
            last.mb = gap;
            // garante espaço visível entre parágrafos de texto quando o Word não tem nenhum
            const a = last.text, b = blk.text;
            const nearObj = /\[(IMG|TABELA|EQ)\]/.test(a + ' ' + b);
            const capPair = (/\[(IMG|TABELA)\]/.test(a) && CAP_SRC.test(b)) || (CAP_TITLE.test(a) && /\[(IMG|TABELA)\]/.test(b));
            const isLabel = a.length < 60 && !/[.;!?]$/.test(a);
            const listNext = blk.para && blk.para.list;
            if (opts.minGap && !capPair && !listNext && gap < 0.5 && (nearObj || !isLabel)
                && !(IS_ITEM.test(a) && IS_ITEM.test(b)) && !(looksCitation(a) && looksCitation(b))) {
                last.mb = Math.max(gap, opts.minGap);
            }
        }
        blk.mb = 0;
        blocks.push(blk);
        pend = blk.trail || 0;
        prevAfter = e.k === 'table' ? 0 : (e.afterEm || 0);
        prevPara = e.k === 'table' ? null : e;
    }

    // título e fonte junto da imagem/tabela ficam centralizados
    blocks.forEach((b, j) => {
        if (!/\[(IMG|TABELA)\]/.test(b.text)) return;
        if (blocks[j - 1] && !blocks[j - 1].table && CAP_TITLE.test(blocks[j - 1].text)) blocks[j - 1].cap = true;
        if (blocks[j + 1] && !blocks[j + 1].table && CAP_SRC.test(blocks[j + 1].text)) blocks[j + 1].cap = true;
    });
    const wp = items.filter(e => e.k === 'p'), rf = refFlags(wp);
    const refSet = new Set(wp.filter((e, i) => rf[i]));
    const fx = n => String(Math.round(n * 100) / 100);

    // listas do Word que o Moodle sabe mostrar sozinho viram <ul>/<ol>; as demais levam o marcador do Word
    const nativeHtml = group => {
        let h = '';
        const st = [];
        const pad0 = Math.max(0.63, group[0].para.leftCm || 1.27);
        for (const b of group) {
            const L = b.para.list, lv = L.ilvl || 0;
            while (st.length && st[st.length - 1].lv > lv) h += '</li></' + st.pop().tag + '>';
            if (!st.length || st[st.length - 1].lv < lv) {
                const n = L.native;
                const pad = st.length ? 0.75 : pad0;
                h += (n.tag === 'ul'
                    ? `<ul style="list-style-type:${n.css};margin:0;padding-left:${fx(pad)}cm">`
                    : `<ol type="${n.type}" start="${n.start}" style="margin:0;padding-left:${fx(pad)}cm">`) + '<li';
                st.push({ lv, tag: n.tag });
            } else {
                h += '</li><li';
            }
            const al = { both: 'justify', center: 'center', right: 'right' }[b.para.jc] || 'left';
            h += ` style="text-align:${al};line-height:${fx(Math.max(1, b.para.lineEm || 1.725))}">` + b.html;
        }
        while (st.length) h += '</li></' + st.pop().tag + '>';
        const mb = group[group.length - 1].mb;
        return h.replace(/^(<(?:ul|ol)\b[^>]*?)margin:0;/, (m, a) => `${a}margin:0 0 ${fx(mb)}em;`);
    };
    const groups = [];
    for (let i = 0; i < blocks.length; i++) {
        const b = blocks[i];
        const L = b.para && b.para.list;
        if (L && L.native && !b.table) {
            const g = [b];
            let j = i + 1;
            while (j < blocks.length && blocks[j].para && blocks[j].para.list && blocks[j].para.list.native && blocks[j].para.list.numId === L.numId) { g.push(blocks[j]); j++; }
            groups.push({ native: g });
            i = j - 1;
        } else {
            groups.push(b);
        }
    }

    return groups.map(b => {
        if (b.native) return nativeHtml(b.native);
        if (b.table) return b.html.replace('<table>', `<table style="border-collapse:collapse;margin:0 auto ${fx(b.mb)}em auto">`);
        const pr = b.para;
        const onlyImg = (pr.imgs.length && !norm(pr.t)) || b.onlyMath;
        const al = (b.cap || onlyImg) ? 'center' : ({ both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right' }[pr.jc] || 'left');
        const lh = Math.max(1, Math.min(3.5, pr.lineEm || 1.725));
        const txt = b.text;
        const body = al === 'justify' && !refSet.has(pr) && !pr.imgs.length && !IS_ITEM.test(txt) && !REF_HEAD.test(txt) && !(txt.length < 60 && !/[.;:!?]$/.test(txt));
        let ind = al === 'center' ? 0 : (pr.indentCm || 0);
        if (!ind && opts.indent && body && !pr.list) ind = 1.25;
        const left = al === 'center' ? 0 : (pr.leftCm || 0);
        if (pr.list) {
            // marcador/numeração do Word com recuo deslocado (as linhas seguintes alinham com o texto)
            const hang = Math.max(0.4, -(pr.indentCm || -0.63));
            return `<p style="margin:0 0 ${fx(b.mb)}em;margin-left:${fx(Math.max(left, hang))}cm;text-indent:-${fx(hang)}cm;text-align:${al};line-height:${fx(lh)}"><span class="marcador" style="display:inline-block;min-width:${fx(hang)}cm;text-indent:0">${esc(pr.list.mk)}&nbsp;</span>${b.html}</p>`;
        }
        return `<p style="margin:0 0 ${fx(b.mb)}em${left > 0.05 ? `;margin-left:${fx(left)}cm` : ''};text-align:${al};line-height:${fx(lh)}${Math.abs(ind) > 0.05 ? `;text-indent:${fx(ind)}cm` : ''}">${b.html || '&nbsp;'}</p>`;
    }).join('');
}

/* Alternativa: parágrafos numa linha cada (o Moodle já coloca a letra), tabelas à parte */
function montarAlternativa(items) {
    let html = '', linhas = [];
    const fechar = () => {
        if (linhas.length) html += `<p dir="ltr" style="text-align: left;">${linhas.join('<br>')}</p>`;
        linhas = [];
    };
    for (const e of items) {
        if (e.k === 'table') {
            fechar();
            html += e.html.replace('<table>', '<table style="border-collapse:collapse">');
        } else if (!paragrafoVazio(e)) {
            linhas.push(e.html.replace(/^(\s*<br>)+/, '').replace(/(<br>\s*)+$/, ''));
        }
    }
    fechar();
    return html;
}
