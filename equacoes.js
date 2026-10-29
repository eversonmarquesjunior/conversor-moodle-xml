// Equações do Word (OMML) e símbolos da fonte Symbol.
// A equação vira LaTeX, o MathJax desenha e o canvas transforma em PNG anexado à questão.
// Adaptado do projeto curadoria-moodle.

/* Fonte Symbol do Word -> Unicode (letras gregas, operadores, setas) */
const SYMBOL_MAP = {
    0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ', 0x45: 'Ε', 0x46: 'Φ', 0x47: 'Γ', 0x48: 'Η', 0x49: 'Ι', 0x4B: 'Κ', 0x4C: 'Λ', 0x4D: 'Μ', 0x4E: 'Ν', 0x4F: 'Ο', 0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ', 0x53: 'Σ', 0x54: 'Τ', 0x55: 'Υ', 0x57: 'Ω', 0x58: 'Ξ', 0x59: 'Ψ', 0x5A: 'Ζ',
    0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε', 0x66: 'φ', 0x67: 'γ', 0x68: 'η', 0x69: 'ι', 0x6A: 'ϕ', 0x6B: 'κ', 0x6C: 'λ', 0x6D: 'μ', 0x6E: 'ν', 0x6F: 'ο', 0x70: 'π', 0x71: 'θ', 0x72: 'ρ', 0x73: 'σ', 0x74: 'τ', 0x75: 'υ', 0x76: 'ϖ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ', 0x7A: 'ζ',
    0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2D: '−', 0xA2: '′', 0xA3: '≤', 0xA5: '∞', 0xAB: '↔', 0xAC: '←', 0xAD: '↑', 0xAE: '→', 0xAF: '↓', 0xB0: '°', 0xB1: '±', 0xB2: '″', 0xB3: '≥', 0xB4: '×', 0xB6: '∂', 0xB7: '•', 0xB8: '÷', 0xB9: '≠', 0xBA: '≡', 0xBB: '≈', 0xBC: '…',
    0xC6: '∅', 0xC7: '∩', 0xC8: '∪', 0xCE: '∈', 0xCF: '∉', 0xD0: '∠', 0xD1: '∇', 0xD6: '√', 0xD7: '·', 0xDB: '⇔', 0xDE: '⇒', 0xE5: '∑', 0xF2: '∫'
};

function symChar(c) {
    let code = parseInt(c.getAttribute('w:char') || '0', 16);
    const font = (c.getAttribute('w:font') || '').toLowerCase();
    if (code >= 0xF000) code -= 0xF000;
    if (font.includes('symbol')) return SYMBOL_MAP[code] || String.fromCharCode(code);
    if (/wingdings|webdings/.test(font)) return BULLET_MAP[code] || '•';
    return code ? String.fromCharCode(code) : '';
}

// Equações encontradas no documento atual (o parágrafo guarda só o índice)
const MATHS = [];

const OMML_PROPS = /^m:(rPr|ctrlPr|fPr|radPr|dPr|naryPr|sSupPr|sSubPr|sSubSupPr|sPrePr|accPr|barPr|funcPr|limLowPr|limUppPr|mPr|eqArrPr|groupChrPr|boxPr|borderBoxPr|oMathParaPr|phantPr)$/;

/* OMML -> HTML simples com sobrescrito/subscrito (usado se a imagem falhar e como descrição da imagem) */
function omml(el) {
    const K = (e, n) => e ? [...e.children].filter(c => c.nodeName === n) : [];
    const sub = (e, n) => { const k = K(e, n)[0]; return k ? omml(k) : { h: '', t: '' }; };
    const cat = e => { let h = '', t = ''; for (const c of e.children) { const r = omml(c); h += r.h; t += r.t; } return { h, t }; };
    const n = el.nodeName;
    if (n === 'm:t' || n === 'w:t') return { h: esc(el.textContent), t: el.textContent };
    if (n === 'm:r') {
        let h = '', t = '';
        for (const c of el.children) if (c.nodeName === 'm:t' || c.nodeName === 'w:t') { h += esc(c.textContent); t += c.textContent; }
        return { h, t };
    }
    if (OMML_PROPS.test(n) || n === 'w:rPr') return { h: '', t: '' };
    if (n === 'm:sSup') { const b = sub(el, 'm:e'), u = sub(el, 'm:sup'); return { h: b.h + '<sup>' + u.h + '</sup>', t: b.t + u.t }; }
    if (n === 'm:sSub') { const b = sub(el, 'm:e'), d = sub(el, 'm:sub'); return { h: b.h + '<sub>' + d.h + '</sub>', t: b.t + d.t }; }
    if (n === 'm:sSubSup') { const b = sub(el, 'm:e'), d = sub(el, 'm:sub'), u = sub(el, 'm:sup'); return { h: b.h + '<sub>' + d.h + '</sub><sup>' + u.h + '</sup>', t: b.t + d.t + u.t }; }
    if (n === 'm:f') {
        const a = sub(el, 'm:num'), b = sub(el, 'm:den');
        const w = x => x.t.length > 1 ? '(' + x.h + ')' : x.h, wt = x => x.t.length > 1 ? '(' + x.t + ')' : x.t;
        return { h: w(a) + '/' + w(b), t: wt(a) + '/' + wt(b) };
    }
    if (n === 'm:rad') { const e = sub(el, 'm:e'), g = sub(el, 'm:deg'); return { h: (g.t ? '<sup>' + g.h + '</sup>' : '') + '√(' + e.h + ')', t: (g.t || '') + '√(' + e.t + ')' }; }
    if (n === 'm:d') {
        const pr = K(el, 'm:dPr')[0];
        const beg = pr && K(pr, 'm:begChr')[0]?.getAttribute('m:val');
        const end = pr && K(pr, 'm:endChr')[0]?.getAttribute('m:val');
        const parts = K(el, 'm:e').map(omml);
        const b = beg ?? '(', e = end ?? ')';
        return { h: esc(b) + parts.map(x => x.h).join(', ') + esc(e), t: b + parts.map(x => x.t).join(', ') + e };
    }
    if (n === 'm:nary') {
        const pr = K(el, 'm:naryPr')[0];
        const ch = (pr && K(pr, 'm:chr')[0]?.getAttribute('m:val')) || '∫';
        const d = sub(el, 'm:sub'), u = sub(el, 'm:sup'), e = sub(el, 'm:e');
        return { h: esc(ch) + (d.t ? '<sub>' + d.h + '</sub>' : '') + (u.t ? '<sup>' + u.h + '</sup>' : '') + e.h, t: ch + d.t + u.t + e.t };
    }
    return cat(el);
}

/* OMML -> LaTeX, para desenhar a equação como imagem igual ao Word */
function omml2tex(el) {
    const K = (e, n) => e ? [...e.children].filter(c => c.nodeName === n) : [];
    const sub = (e, n) => { const k = K(e, n)[0]; return k ? omml2tex(k) : ''; };
    const val = (e, pr, n) => { const p = K(e, pr)[0]; const c = p && K(p, n)[0]; return c ? c.getAttribute('m:val') : null; };
    const cat = e => [...e.children].map(omml2tex).join('');
    const on = v => v === '1' || v === 'on' || v === 'true';
    const texEsc = t => t.replace(/\\/g, '\\backslash ').replace(/([#%&_{}$])/g, '\\$1').replace(/\^/g, '\\hat{}').replace(/~/g, '\\sim ').replace(/ /g, '\\ ').replace(/\u2061/g, '');
    const n = el.nodeName;
    if (n === 'm:t' || n === 'w:t') return texEsc(el.textContent);
    if (n === 'm:r') {
        const nor = K(el, 'm:rPr')[0] && K(K(el, 'm:rPr')[0], 'm:nor').length;
        let t = '';
        for (const c of el.children) if (c.nodeName === 'm:t' || c.nodeName === 'w:t') t += c.textContent;
        if (!t) return '';
        if (/^(sen|cos|tan|tg|log|ln|exp|lim|max|min|sin|sec|cot|cotg|cossec)$/i.test(t)) return '\\operatorname{' + t + '}';
        if (nor) return '\\mathrm{' + texEsc(t) + '}';
        return texEsc(t);
    }
    if (OMML_PROPS.test(n) || n === 'w:rPr' || n === 'w:bookmarkStart' || n === 'w:bookmarkEnd') return '';
    if (n === 'm:sSup') return '{' + sub(el, 'm:e') + '}^{' + sub(el, 'm:sup') + '}';
    if (n === 'm:sSub') return '{' + sub(el, 'm:e') + '}_{' + sub(el, 'm:sub') + '}';
    if (n === 'm:sSubSup') return '{' + sub(el, 'm:e') + '}_{' + sub(el, 'm:sub') + '}^{' + sub(el, 'm:sup') + '}';
    if (n === 'm:sPre') return '{}_{' + sub(el, 'm:sub') + '}^{' + sub(el, 'm:sup') + '}{' + sub(el, 'm:e') + '}';
    if (n === 'm:f') {
        const t = val(el, 'm:fPr', 'm:type');
        const a = sub(el, 'm:num'), b = sub(el, 'm:den');
        return t === 'lin' ? a + '/' + b : t === 'noBar' ? '\\genfrac{}{}{0pt}{}{' + a + '}{' + b + '}' : '\\frac{' + a + '}{' + b + '}';
    }
    if (n === 'm:rad') {
        const d = sub(el, 'm:deg');
        return (d && !on(val(el, 'm:radPr', 'm:degHide'))) ? '\\sqrt[' + d + ']{' + sub(el, 'm:e') + '}' : '\\sqrt{' + sub(el, 'm:e') + '}';
    }
    if (n === 'm:d') {
        const map = c => c == null ? null : c === '' ? '.' : ({ '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|', '⟨': '\\langle', '⟩': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil' }[c] || c);
        const b = map(val(el, 'm:dPr', 'm:begChr')) ?? '(', e = map(val(el, 'm:dPr', 'm:endChr')) ?? ')';
        const sep = val(el, 'm:dPr', 'm:sepChr') ?? '|';
        return '\\left' + b + K(el, 'm:e').map(omml2tex).join(sep === '|' ? '\\mid ' : texEsc(sep)) + '\\right' + e;
    }
    if (n === 'm:nary') {
        const ch = val(el, 'm:naryPr', 'm:chr') || '∫';
        const op = { '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap' }[ch] || ch;
        const lo = sub(el, 'm:sub'), hi = sub(el, 'm:sup');
        return op + ((lo && !on(val(el, 'm:naryPr', 'm:subHide'))) ? '_{' + lo + '}' : '') + ((hi && !on(val(el, 'm:naryPr', 'm:supHide'))) ? '^{' + hi + '}' : '') + '{' + sub(el, 'm:e') + '}';
    }
    if (n === 'm:func') {
        const f = K(el, 'm:fName')[0];
        const fn = f ? cat(f) : '';
        return (/^[A-Za-z]+$/.test(fn) ? '\\operatorname{' + fn + '}' : fn) + '{' + sub(el, 'm:e') + '}';
    }
    if (n === 'm:limLow') {
        const e = sub(el, 'm:e');
        return /^(\\mathrm\{|\\operatorname\{)?lim\}?$/.test(e) ? '\\lim_{' + sub(el, 'm:lim') + '}' : '\\underset{' + sub(el, 'm:lim') + '}{' + e + '}';
    }
    if (n === 'm:limUpp') return '\\overset{' + sub(el, 'm:lim') + '}{' + sub(el, 'm:e') + '}';
    if (n === 'm:acc') {
        const ch = val(el, 'm:accPr', 'm:chr') || '\u0302';
        const c = { '\u0302': '\\hat', '\u0303': '\\tilde', '\u0304': '\\bar', '\u0305': '\\overline', '\u0307': '\\dot', '\u0308': '\\ddot', '\u20D7': '\\vec', '\u2192': '\\vec' }[ch] || '\\hat';
        return c + '{' + sub(el, 'm:e') + '}';
    }
    if (n === 'm:bar') return (val(el, 'm:barPr', 'm:pos') === 'bot' ? '\\underline{' : '\\overline{') + sub(el, 'm:e') + '}';
    if (n === 'm:groupChr') {
        const ch = val(el, 'm:groupChrPr', 'm:chr') || '\u23DF';
        const pos = val(el, 'm:groupChrPr', 'm:pos');
        return (ch === '\u23DE' || pos === 'top' ? '\\overbrace{' : '\\underbrace{') + sub(el, 'm:e') + '}';
    }
    if (n === 'm:eqArr') return '\\begin{aligned}' + K(el, 'm:e').map(omml2tex).join('\\\\') + '\\end{aligned}';
    if (n === 'm:m') return '\\begin{matrix}' + K(el, 'm:mr').map(r => K(r, 'm:e').map(omml2tex).join('&')).join('\\\\') + '\\end{matrix}';
    if (n === 'm:box' || n === 'm:borderBox' || n === 'm:phant') return sub(el, 'm:e');
    if (n === 'm:e' && !el.children.length) return '';
    return cat(el);
}

/* MathJax é carregado só quando o documento tem equação */
let MJ = null;
function loadMathJax() {
    if (MJ) return MJ;
    MJ = new Promise((res, rej) => {
        window.MathJax = { startup: { typeset: false }, svg: { fontCache: 'none' } };
        const sc = document.createElement('script');
        sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/mathjax/3.2.2/es5/tex-svg-full.js';
        sc.async = true;
        sc.onload = () => window.MathJax.startup.promise.then(() => res(window.MathJax), rej);
        sc.onerror = () => rej(new Error('Não foi possível carregar o desenhador de equações.'));
        document.head.appendChild(sc);
    });
    return MJ;
}

async function mathToPng(tex, display) {
    const MJx = await loadMathJax();
    const node = await MJx.tex2svgPromise(tex, { display: !!display, em: 16, ex: 8, containerWidth: 900 });
    if (node.querySelector('[data-mml-node="merror"]')) throw new Error('LaTeX inválido');
    const svg = node.querySelector('svg');
    if (!svg) throw new Error('svg');
    const ex = 8.6;
    const w = Math.max(1, Math.ceil(parseFloat(svg.getAttribute('width')) * ex));
    const h = Math.max(1, Math.ceil(parseFloat(svg.getAttribute('height')) * ex));
    svg.setAttribute('width', w);
    svg.setAttribute('height', h);
    svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    svg.setAttribute('color', '#000');
    svg.style.color = '#000';
    const src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(svg));
    const img = await new Promise((r, j) => { const im = new Image(); im.onload = () => r(im); im.onerror = j; im.src = src; });
    const pad = 4, sc = 2;
    const cv = document.createElement('canvas');
    cv.width = (w + pad * 2) * sc;
    cv.height = (h + pad * 2) * sc;
    const cx = cv.getContext('2d');
    cx.fillStyle = '#fff';
    cx.fillRect(0, 0, cv.width, cv.height);
    cx.drawImage(img, pad * sc, pad * sc, w * sc, h * sc);
    return { b64: cv.toDataURL('image/png').split(',')[1], w: w + pad * 2, h: h + pad * 2 };
}

/* Troca os marcadores <span data-math> por <img> da equação (ou pelo texto formatado, se não der para desenhar) */
async function resolverEquacoes(html, nomear, registrar) {
    const ids = [...html.matchAll(/<span data-math="(\d+)">/g)].map(m => +m[1]);
    const falhas = [];
    for (const id of ids) {
        const M = MATHS[id];
        let rep;
        try {
            if (!M.tex) throw new Error('sem LaTeX');
            const r = await mathToPng(M.tex, M.display);
            const nome = nomear();
            registrar(nome, r.b64, 'image/png');
            rep = `<img src="@@PLUGINFILE@@/${encodeURIComponent(nome)}" alt="${esc(M.t)}" width="${r.w}" height="${r.h}" class="equacao" style="vertical-align:middle">`;
        } catch (e) {
            falhas.push(M.t);
            rep = `<span class="equacao-texto">${M.h}</span>`;
        }
        html = html.replace(new RegExp(`<span data-math="${id}">[\\s\\S]*?<\\/span>`), () => rep);
    }
    return { html, falhas };
}
