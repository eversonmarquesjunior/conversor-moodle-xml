// Relatório de curadoria: confere o XML gerado com o Word e aponta problemas nas questões.
// Verificações adaptadas do projeto curadoria-moodle.

const SEV_LBL = { erro: 'Erro', atencao: 'Atenção', info: 'Info' };
const SEV_RANK = { erro: 3, atencao: 2, info: 1 };
const CATEGORIAS = ['Fidelidade', 'Fórmulas', 'Imagens', 'Tabelas', 'Estrutura', 'Duplicidade', 'Formatação'];
const LETRAS = 'abcdefghij';
const semEspacos = s => String(s || '').replace(/[\s ]+/g, '');
const curto = (s, n = 90) => { s = String(s); return s.length > n ? s.slice(0, n) + '…' : s; };
const pior = q => q.issues.reduce((m, i) => Math.max(m, SEV_RANK[i.sev]), 0);

/* ---------- leitura independente do Word, para conferir o que chegou ao XML ---------- */
function refVazia() { return { texto: '', imgs: 0, maths: 0, tables: 0, ole: 0 }; }

function somarRef(a, b) {
    a.texto += ' ' + b.texto;
    a.imgs += b.imgs; a.maths += b.maths; a.tables += b.tables; a.ole += b.ole;
}

const IGNORAR_NA_CONFERENCIA = new Set(['w:del', 'w:moveFrom', 'mc:Fallback', 'w:pPr', 'w:rPr', 'w:tcPr', 'w:sdtPr', 'w:instrText']);

function extrairTextoWord(el, ref) {
    for (const c of el.children) {
        const n = c.nodeName;
        if (n === 'w:t') ref.texto += c.textContent;
        else if (n === 'w:sym') ref.texto += symChar(c);
        else if (n === 'w:noBreakHyphen') ref.texto += '-';
        else if (n === 'w:p' || n === 'w:br' || n === 'w:tab') { ref.texto += ' '; if (n === 'w:p') extrairTextoWord(c, ref); }
        else if (n === 'm:oMath' || n === 'm:oMathPara') ref.maths++;
        else if (n === 'w:object') ref.ole++;
        else if (IGNORAR_NA_CONFERENCIA.has(n)) continue;
        else {
            if (n === 'w:tbl') ref.tables++;
            if ((n === 'w:drawing' || n === 'w:pict') && (c.getElementsByTagName('a:blip').length || c.getElementsByTagName('v:imagedata').length)) ref.imgs++;
            extrairTextoWord(c, ref);
        }
    }
    return ref;
}

/* Texto que o aluno vai ler no Moodle (sem marcadores de lista e sem as equações, que são conferidas à parte) */
function textoDoHtml(html) {
    const doc = new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html');
    doc.body.querySelectorAll('.marcador, .equacao-texto, img').forEach(el => el.remove());
    doc.body.querySelectorAll('br, p, li, td').forEach(el => el.after(' '));
    return doc.body.textContent;
}

function contarNoHtml(html) {
    return {
        imgs: (html.match(/<img\b(?![^>]*class="equacao")[^>]*>/gi) || []).length,
        eqs: (html.match(/class="equacao(-texto)?"/g) || []).length,
        tables: (html.match(/<table\b/gi) || []).length
    };
}

function lcsDiff(a, b) {
    const n = a.length, m = b.length;
    const dp = Array.from({ length: n + 1 }, () => new Int16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    const ops = [];
    let i = 0, j = 0;
    while (i < n || j < m) {
        if (i < n && j < m && a[i] === b[j]) { ops.push(['=', a[i]]); i++; j++; }
        else if (j < m && (i >= n || dp[i][j + 1] >= dp[i + 1][j])) { ops.push(['+', b[j]]); j++; }
        else { ops.push(['-', a[i]]); i++; }
    }
    return ops;
}

// trechos do Word que faltam no XML e trechos do XML que não estão no Word, palavra por palavra
function diferencaPalavras(word, xml) {
    const a = word.split(/\s+/).filter(Boolean), b = xml.split(/\s+/).filter(Boolean);
    if (a.length * b.length > 4e6) return { falta: [curto(a.join(' '), 120)], sobra: [curto(b.join(' '), 120)] };
    const falta = [], sobra = [];
    let cur = null;
    for (const [op, w] of lcsDiff(a, b)) {
        if (op === '=') { cur = null; continue; }
        if (cur && cur.op === op) { cur.txt.push(w); continue; }
        cur = { op, txt: [w] };
        (op === '-' ? falta : sobra).push(cur.txt);
    }
    return { falta: falta.map(x => x.join(' ')), sobra: sobra.map(x => x.join(' ')) };
}

/* Alternativas exatamente iguais (duplicada, triplicada…). Só os espaços são ignorados:
   pontuação e algarismos contam ("I, II e IV" é diferente de "III e IV"). */
function dupCheck(texts) {
    const out = [];
    const K = texts.map(norm);
    const lst = a => a.length === 2 ? a.join(' e ') : a.slice(0, -1).join(', ') + ' e ' + a[a.length - 1];
    const NOME = { 2: 'duplicada', 3: 'triplicada', 4: 'quadruplicada', 5: 'quintuplicada' };
    const groups = {};
    K.forEach((k, i) => { if (k) (groups[k] = groups[k] || []).push(i); });
    for (const g of Object.values(groups)) {
        if (g.length < 2) continue;
        const letras = g.map(i => LETRAS[i]);
        const nome = NOME[g.length] || `repetida ${g.length} vezes`;
        out.push({ title: `Alternativa ${nome}: ${lst(letras)} iguais`, detail: `${g.length === 2 ? 'As duas têm' : 'Todas têm'} o mesmo texto: “${curto(texts[g[0]], 110)}”.` });
    }
    return out;
}

/* ---------- verificações ---------- */
function verificarFidelidade(q, add) {
    const partes = [['Enunciado', q.ref.enun, q.enuncHtml], ...q.alternatives.map((a, i) => [`Alternativa ${LETRAS[i]}`, a.ref, a.html])];
    const tot = { imgsW: 0, imgsX: 0, tabW: 0, tabX: 0, eqW: 0, eqX: 0, ole: 0 };
    for (const [onde, ref, html] of partes) {
        const c = contarNoHtml(html);
        tot.imgsW += ref.imgs; tot.imgsX += c.imgs;
        tot.tabW += ref.tables; tot.tabX += c.tables;
        tot.eqW += ref.maths; tot.eqX += c.eqs;
        tot.ole += ref.ole;
        const xmlTxt = textoDoHtml(html);
        if (semEspacos(ref.texto) === semEspacos(xmlTxt)) continue;
        const d = diferencaPalavras(ref.texto, xmlTxt);
        const lines = [...d.falta.slice(0, 3).map(t => ['w', curto(t, 140)]), ...d.sobra.slice(0, 3).map(t => ['x', curto(t, 140)])];
        add('erro', 'Fidelidade', `${onde}: texto do XML diferente do Word`, 'Trechos que estão só no Word ou só no XML gerado:', lines);
    }
    const descartadas = q.midiaDescartada.length;
    if (tot.imgsW !== tot.imgsX + descartadas) add('atencao', 'Imagens', `${tot.imgsW} imagem(ns) no Word, ${tot.imgsX} no XML`, 'Alguma imagem não foi reconhecida (imagem vinculada a arquivo externo, dentro de caixa de texto ou em formato não suportado).');
    q.midiaDescartada.forEach(t => {
        if (t === 'emf') add('erro', 'Imagens', 'Imagem em formato EMF/WMF não foi incluída', 'Esse formato não aparece no Moodle. Salve a imagem como PNG ou JPG e insira de novo no Word.');
        else add('erro', 'Imagens', 'Imagem quebrada ou vinculada a arquivo externo', 'A imagem não está dentro do Word. Insira a imagem de novo (Inserir → Imagens → Este dispositivo).');
    });
    if (tot.tabW !== tot.tabX) add('erro', 'Tabelas', `${tot.tabW} tabela(s) no Word, ${tot.tabX} no XML`, 'Tabela dentro de outra tabela não é convertida. Confira o conteúdo da tabela no XML.');
    if (tot.eqW !== tot.eqX) add('erro', 'Fórmulas', `${tot.eqW} equação(ões) no Word, ${tot.eqX} no XML`);
    if (q.eqFalhas.length) add('atencao', 'Fórmulas', `${q.eqFalhas.length} equação(ões) ficaram como texto`, `Não foi possível desenhar a equação como imagem (sem internet ou equação muito complexa). Ex.: “${curto(q.eqFalhas[0], 60)}”`);
    if (tot.ole) add('erro', 'Fórmulas', `${tot.ole} objeto(s) antigo(s) não convertido(s) (Equation 3.0 ou objeto incorporado)`, 'Refaça a equação no editor de equações do Word (Inserir → Equação) para ela chegar ao Moodle.');
}

function verificarEstrutura(q, add, stdAlts) {
    const alts = q.alternatives;
    const txt = alts.map(a => textoItens(a.items));
    if (alts.length !== stdAlts) add('erro', 'Estrutura', `${alts.length} alternativas (o padrão do banco é ${stdAlts})`, 'Confira se falta ou sobra alguma alternativa.');
    const nOk = alts.filter(a => a.fraction > 0).length;
    if (nOk === 0) add('erro', 'Estrutura', 'Nenhuma alternativa marcada como CORRETA', 'Marque a alternativa correta no quadro da questão.');
    else if (nOk > 1) add('erro', 'Estrutura', `${nOk} alternativas marcadas como CORRETA`, 'Em questão de múltipla escolha deve haver só uma correta (pode ser mistura de duas questões).');
    alts.forEach((a, i) => {
        const L = LETRAS[i], t = txt[i];
        if (!temConteudo(a.items)) add('erro', 'Estrutura', `Alternativa ${L} vazia`);
        if (/^\s*\(?[a-eA-E]\s*[\)\.\-–]\s/.test(t)) add('atencao', 'Estrutura', `Letra digitada no início da alternativa ${L}`, `O Moodle já coloca a letra e embaralha as alternativas. Retire “${t.trim().slice(0, 3)}”.`);
    });
    const lens = txt.map(t => t.length), ci = alts.findIndex(a => a.fraction > 0);
    if (ci >= 0 && alts.length >= 3) {
        const others = lens.filter((_, i) => i !== ci);
        const avg = others.reduce((x, y) => x + y, 0) / others.length;
        if (lens[ci] > 40 && lens[ci] > avg * 1.7 && lens[ci] > Math.max(...others) * 1.3) {
            add('info', 'Estrutura', 'Alternativa correta bem mais longa que as outras', `A correta tem ${lens[ci]} caracteres; as outras, em média, ${Math.round(avg)}. Isso pode entregar a resposta.`);
        }
    }
    dupCheck(txt).forEach(d => add('erro', 'Duplicidade', d.title, d.detail));
}

function verificarFormatacao(q, add) {
    const ps = [...q.enunItems, ...q.alternatives.flatMap(a => a.items)].filter(e => e.k === 'p');
    const col = ps.flatMap(p => p.colored);
    if (col.length) add('atencao', 'Formatação', 'Texto colorido', `Pode ser marcação de revisão esquecida. Ex.: “${curto(col[0], 60)}”`);
}

function gerarRelatorio(res) {
    const R = { nome: res.nome, questoes: [], global: [], naoConvertidas: res.errors.length };
    const qs = res.questions;
    const nAlts = {};
    qs.forEach(q => { nAlts[q.alternatives.length] = (nAlts[q.alternatives.length] || 0) + 1; });
    const stdAlts = +(Object.entries(nAlts).sort((a, b) => b[1] - a[1])[0]?.[0] || 5);

    qs.forEach((q, i) => {
        const Q = { nome: `Q${String(i + 1).padStart(2, '0')}`, titulo: q.titulo ? q.titulo.texto : '', issues: [], q };
        const add = (sev, cat, title, detail = '', lines) => Q.issues.push({ sev, cat, title, detail, lines });
        verificarFidelidade(q, add);
        verificarEstrutura(q, add, stdAlts);
        verificarFormatacao(q, add);
        Q.stem = textoItens(q.enunItems);
        R.questoes.push(Q);
    });

    // questões repetidas no mesmo banco (enunciado exatamente igual)
    const Qs = R.questoes;
    for (let i = 0; i < Qs.length; i++) for (let j = i + 1; j < Qs.length; j++) {
        const a = Qs[i], b = Qs[j];
        if (a.stem.length < 40 || a.stem !== b.stem) continue;
        for (const [x, y] of [[a, b], [b, a]]) {
            x.issues.unshift({ sev: 'erro', cat: 'Duplicidade', title: `Questão repetida: enunciado igual ao da ${y.nome}`, detail: `Confira se uma delas deve ser trocada. “${curto(x.stem, 100)}”` });
        }
    }

    // numeração das questões (quando o Word tem títulos "Questão N")
    const nums = qs.filter(q => q.titulo).map(q => q.titulo.n);
    if (nums.length >= Math.max(2, qs.length / 2)) {
        // a numeração que volta para 1 começa um novo bloco (várias atividades no mesmo Word)
        const blocos = [];
        nums.forEach((n, i) => { if (i === 0 || n === 1) blocos.push([]); blocos[blocos.length - 1].push(n); });
        const fmt = n => 'Questão ' + String(n).padStart(2, '0');
        blocos.forEach((b, bi) => {
            const onde = blocos.length > 1 ? ` no ${bi + 1}º bloco de questões` : '';
            const rep = [...new Set(b.filter((n, i) => b.indexOf(n) !== i))];
            const miss = [];
            for (let k = Math.min(...b); k <= Math.max(...b); k++) if (!b.includes(k)) miss.push(k);
            if (rep.length) R.global.push(['erro', `Numeração repetida${onde}: ${rep.map(fmt).join(', ')} aparece${rep.length > 1 ? 'm' : ''} mais de uma vez.`]);
            if (miss.length) R.global.push(['erro', `Numeração pulada${onde}: falta${miss.length > 1 ? 'm' : ''} ${miss.map(fmt).join(', ')}.`]);
        });
    }

    res.errors.forEach(e => R.global.push(['erro', `Questão não convertida — ${e.reason}: “${e.snippet}”`]));
    const rv = res.doc.revisoes;
    if (rv.ins || rv.del) R.global.push(['erro', `O Word tem alterações controladas não aceitas (${rv.ins} inclusões e ${rv.del} exclusões). Aceite ou rejeite todas (Revisão → Aceitar) e converta de novo.`]);
    if (rv.comentarios) R.global.push(['atencao', `O Word tem ${rv.comentarios} comentário(s). Resolva e exclua os comentários antes da inserção.`]);

    R.questoes.forEach(Q => Q.issues.sort((a, b) => SEV_RANK[b.sev] - SEV_RANK[a.sev]));
    return R;
}

/* ---------- exibição ---------- */
function htmlComImagens(html, images) {
    let out = html;
    for (const img of images) out = out.split(`@@PLUGINFILE@@/${encodeURIComponent(img.name)}`).join(`data:${img.mime};base64,${img.b64}`);
    return out;
}

function corpoQuestao(Q) {
    const q = Q.q;
    const issues = Q.issues.length
        ? `<ul class="rel-issues">${Q.issues.map(i => `<li class="rel-issue ${i.sev}">
              <div class="rel-it"><span class="rel-cat">${i.cat}</span>${esc(i.title)}</div>
              ${i.detail ? `<div class="rel-id">${esc(i.detail)}</div>` : ''}
              ${(i.lines || []).map(([k, t]) => `<div class="rel-diff ${k}">${k === 'w' ? 'Só no Word' : 'Só no XML'}: ${esc(t)}</div>`).join('')}
           </li>`).join('')}</ul>`
        : '<div class="rel-none">Nenhum problema encontrado.</div>';
    const alts = q.alternatives.map((a, i) => `<li class="${a.fraction > 0 ? 'ok' : ''}"><span>${LETRAS[i]}.</span><div>${htmlComImagens(a.html, q.images)}${a.fraction > 0 ? ' ✓' : ''}</div></li>`).join('');
    return issues + `<div class="rel-preview-lbl">Como fica no Moodle</div>
        <div class="rel-preview">${htmlComImagens(q.enuncHtml, q.images)}<ol class="rel-alts">${alts}</ol></div>`;
}

function criarSecaoRelatorio(R) {
    const N = R.questoes.length;
    const comErro = R.questoes.filter(q => pior(q) >= 3).length;
    const comAtencao = R.questoes.filter(q => pior(q) === 2).length;
    const sec = document.createElement('section');
    sec.className = 'rel';
    sec.innerHTML = `
        <div class="rel-head"><h2>Relatório de curadoria</h2><span class="rel-file">${esc(R.nome)}</span></div>
        <div class="rel-sum">
            <span class="rel-big">${N} ${N === 1 ? 'questão' : 'questões'}</span>
            <span><i class="dot erro"></i><b>${comErro}</b> com erro</span>
            <span><i class="dot atencao"></i><b>${comAtencao}</b> só com atenção</span>
            <span><i class="dot ok"></i><b>${N - comErro - comAtencao}</b> sem problemas</span>
        </div>
        ${R.global.length ? `<ul class="rel-global">${R.global.map(([s, t]) => `<li><span class="sev ${s}">${SEV_LBL[s]}</span>${esc(t)}</li>`).join('')}</ul>` : ''}
        <div class="rel-tools">
            <label><input type="checkbox" class="rel-only"${comErro + comAtencao > 0 ? ' checked' : ''}> Só questões com erro ou atenção</label>
            <button type="button" class="rel-btn rel-expand">Expandir tudo</button>
            <button type="button" class="rel-btn rel-copy">Copiar relatório</button>
        </div>
        <div class="rel-list"></div>`;

    let expandir = false;
    const lista = sec.querySelector('.rel-list');
    const abrir = (el, Q) => {
        const body = el.querySelector('.rel-qbody');
        if (!body.innerHTML) {
            body.innerHTML = corpoQuestao(Q);
            body.querySelectorAll('a').forEach(a => { a.target = '_blank'; a.rel = 'noopener'; });
        }
        body.hidden = false;
        el.querySelector('.rel-qhead').setAttribute('aria-expanded', 'true');
    };
    const desenhar = () => {
        const so = sec.querySelector('.rel-only').checked;
        const qs = so ? R.questoes.filter(q => pior(q) >= 2) : R.questoes;
        if (!qs.length) { lista.innerHTML = '<div class="rel-none">Nenhuma questão com erro ou atenção.</div>'; return; }
        lista.innerHTML = qs.map((Q, k) => {
            const porCat = {};
            for (const i of Q.issues) if (!porCat[i.cat] || SEV_RANK[i.sev] > SEV_RANK[porCat[i.cat]]) porCat[i.cat] = i.sev;
            const badges = Q.issues.length
                ? CATEGORIAS.filter(c => porCat[c]).map(c => `<span class="sev ${porCat[c]}">${c}</span>`).join('')
                : '<span class="sev ok">OK</span>';
            return `<div class="rel-q" data-k="${k}">
                <button type="button" class="rel-qhead" aria-expanded="false"><span class="rel-qn">${Q.nome}</span><span class="rel-qt">${esc(curto(Q.titulo || Q.stem, 70))}</span><span class="rel-badges">${badges}</span></button>
                <div class="rel-qbody" hidden></div></div>`;
        }).join('');
        lista.querySelectorAll('.rel-q').forEach(el => {
            const Q = qs[+el.dataset.k];
            if (expandir) abrir(el, Q);
            el.querySelector('.rel-qhead').addEventListener('click', () => {
                const body = el.querySelector('.rel-qbody');
                if (body.hidden) abrir(el, Q);
                else { body.hidden = true; el.querySelector('.rel-qhead').setAttribute('aria-expanded', 'false'); }
            });
        });
    };
    sec.querySelector('.rel-only').addEventListener('change', desenhar);
    sec.querySelector('.rel-expand').addEventListener('click', e => {
        expandir = !expandir;
        e.target.textContent = expandir ? 'Recolher tudo' : 'Expandir tudo';
        desenhar();
    });
    sec.querySelector('.rel-copy').addEventListener('click', () => copiarRelatorio(R));
    desenhar();
    return sec;
}

function mostrarRelatorios(relatorios) {
    const box = document.getElementById('relatorio');
    box.innerHTML = '';
    relatorios.forEach(R => box.appendChild(criarSecaoRelatorio(R)));
    box.hidden = false;
    box.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function limparRelatorios() {
    const box = document.getElementById('relatorio');
    box.innerHTML = '';
    box.hidden = true;
}

/* Copia em tabela: colado no Word, Excel, e-mail ou Teams sai organizado */
async function copiarRelatorio(R) {
    const N = R.questoes.length;
    const comErro = R.questoes.filter(q => pior(q) >= 3).length, comAtencao = R.questoes.filter(q => pior(q) === 2).length;
    const head = ['Questão', 'Categoria', 'Gravidade', 'Problema', 'Detalhe'];
    const rows = [];
    for (const Q of R.questoes) for (const i of Q.issues) {
        if (i.sev === 'info') continue;
        const lines = (i.lines || []).map(([k, t]) => `${k === 'w' ? 'Só no Word' : 'Só no XML'}: ${t}`).join('\n');
        rows.push([Q.nome + (Q.titulo ? ' — ' + curto(Q.titulo, 40) : ''), i.cat, SEV_LBL[i.sev], i.title, [i.detail, lines].filter(Boolean).join('\n')]);
    }
    const resumo = `${N} questões · ${comErro} com erro · ${comAtencao} só com pontos de atenção · ${N - comErro - comAtencao} sem problemas.`;
    const cell = v => esc(String(v ?? '')).replace(/\n/g, '<br>');
    const html = `<div style="font-family:Arial,sans-serif;font-size:10pt">
        <p style="margin:0 0 4px"><b>Relatório de curadoria — ${esc(R.nome)}</b></p><p style="margin:0 0 4px">${esc(resumo)}</p>
        ${R.global.map(([s, t]) => `<p style="margin:0 0 4px">${SEV_LBL[s]}: ${esc(t)}</p>`).join('')}
        ${rows.length ? `<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:10pt"><thead><tr>${head.map(h => `<th style="background:#e1efef;text-align:left">${cell(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td style="vertical-align:top">${cell(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>` : ''}</div>`;
    const tsv = [`Relatório de curadoria — ${R.nome}`, resumo, ...R.global.map(([s, t]) => `${SEV_LBL[s]}: ${t}`),
        ...(rows.length ? [head, ...rows].map(r => r.map(c => String(c ?? '').replace(/[\t\r\n]+/g, ' ')).join('\t')) : [])].join('\n');
    try {
        if (window.ClipboardItem && navigator.clipboard?.write) {
            await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([tsv], { type: 'text/plain' }) })]);
        } else {
            await navigator.clipboard.writeText(tsv);
        }
        showToast('Relatório copiado. Cole no Word, Excel ou e-mail.', 'success');
    } catch (e) {
        showToast('Não foi possível copiar o relatório', 'error');
    }
}
