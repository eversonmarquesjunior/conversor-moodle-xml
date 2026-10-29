let filesData = [];
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      REL_NS='http://schemas.openxmlformats.org/package/2006/relationships';

const dropOverlay = document.getElementById('drop-overlay');

function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

// Drag and drop funciona na página inteira (qualquer ponto), não só no quadrado
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    window.addEventListener(eventName, preventDefaults, false);
});

let dragCounter = 0;

window.addEventListener('dragenter', () => {
    dragCounter++;
    dropOverlay.classList.add('active');
});

window.addEventListener('dragleave', () => {
    dragCounter--;
    if (dragCounter <= 0) {
        dragCounter = 0;
        dropOverlay.classList.remove('active');
    }
});

window.addEventListener('drop', (e) => {
    dragCounter = 0;
    dropOverlay.classList.remove('active');
    handleFiles(e.dataTransfer.files);
});

document.getElementById('fi').onchange = e => handleFiles(e.target.files);

function handleFiles(files) {
    const filtered = [...files].filter(f => f.name.endsWith('.docx'));
    if(filtered.length === 0) {
        showToast("Por favor, selecione arquivos .docx", "error");
        return;
    }
    let loadedCount = 0;
    filtered.forEach(f => {
        const reader = new FileReader();
        reader.onload = ev => {
            filesData.push({ name: f.name, buffer: ev.target.result });
            document.getElementById('fl').innerHTML += `<div class="file-item"><span>📄 ${escXml(f.name)}</span></div>`;
            document.getElementById('bcvt').disabled = false;
            document.getElementById('btn-clear').style.display = 'block';
            loadedCount++;
            if(loadedCount === filtered.length) showToast(`${filtered.length} arquivo(s) adicionado(s)`, "success");
        };
        reader.readAsArrayBuffer(f);
    });
}

function clearFiles() {
    filesData = [];
    document.getElementById('fl').innerHTML = '';
    document.getElementById('bcvt').disabled = true;
    document.getElementById('btn-clear').style.display = 'none';
    document.getElementById('fi').value = '';
    limparRelatorios();
    showToast("Arquivos removidos", "success");
}

function showToast(msg, type) {
    const t = document.createElement('div');
    t.className = `toast ${type}`; t.innerText = msg;
    document.getElementById('toast-container').appendChild(t);
    setTimeout(() => {
        t.style.opacity = '0';
        setTimeout(() => t.remove(), 400);
    }, 3000);
}

function escXml(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function levenshtein(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = a[i - 1] === b[j - 1]
                ? dp[i - 1][j - 1]
                : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
        }
    }
    return dp[m][n];
}

// Tolera erros comuns de digitação no rótulo "ENUNCIADO" (ex: "ENUNCIA DO", "ENUNCIDO")
function isEnunciadoLabel(label) {
    const norm = String(label).toUpperCase().replace(/\s+/g, '');
    if (norm.includes('ENUNCIADO')) return true;
    return norm.length > 0 && norm.length <= 12 && levenshtein(norm, 'ENUNCIADO') <= 1;
}

function textSnippet(text, maxLen = 90) {
    const t = norm(text);
    return t.length > maxLen ? t.slice(0, maxLen) + '…' : t;
}

// Tabela aninhada = está dentro da célula de outra tabela (ex: tabela de apoio/dados dentro do enunciado)
function isNestedTable(tbl, docXml) {
    let node = tbl.parentNode;
    while (node && node !== docXml) {
        if (node.namespaceURI === W && node.localName === 'tbl') return true;
        node = node.parentNode;
    }
    return false;
}

// Título "Questão 01 (Área)" logo antes da tabela da questão, se houver
function tituloAntes(tbl) {
    let el = tbl.previousElementSibling;
    for (let i = 0; el && i < 6; i++, el = el.previousElementSibling) {
        if (el.localName === 'tbl') return null;
        if (el.localName !== 'p') continue;
        const t = el.textContent.trim();
        if (!t) continue;
        const m = t.match(/^Quest[ãa]o\s*(\d+)\s*(?:\((.*?)\))?/i);
        return m ? { el, n: +m[1], area: (m[2] || '').trim(), texto: t } : null;
    }
    return null;
}

async function parseDOCX(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const relsText = await zip.file('word/_rels/document.xml.rels').async('text');
    const relsDoc = new DOMParser().parseFromString(relsText, 'text/xml');
    const relMap = {};
    for (const rel of relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship')) {
        relMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
    }

    const imgMap = {};
    for (const [rid, target] of Object.entries(relMap)) {
        const path = 'word/' + target.replace(/^\.?\//, '').replace(/^word\//, '');
        const f = zip.file(path);
        const ext = (path.split('.').pop() || '').toLowerCase();
        if (f && /^(png|jpe?g|gif|bmp|svg|emf|wmf)$/.test(ext)) {
            const mime = ext === 'svg' ? 'image/svg+xml' : 'image/' + (ext === 'jpg' ? 'jpeg' : ext);
            imgMap[rid] = { ext, mime, b64: await f.async('base64') };
        }
    }

    const docXml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('text'), 'text/xml');
    await carregarFormatacaoDocx(zip, relMap);

    const questions = [];
    const errors = [];

    function finalizeQuestion(q) {
        if (!q) return;
        const hasEnunc = temConteudo(q.enunItems);
        if (!hasEnunc && q.alternatives.length === 0) return;
        if (hasEnunc && q.alternatives.length > 0) {
            questions.push(q);
            return;
        }
        const reason = !hasEnunc
            ? 'alternativa(s) encontrada(s) sem enunciado (rótulo "ENUNCIADO" ausente ou não reconhecido)'
            : 'enunciado encontrado sem nenhuma alternativa "Correta"/"Incorreta"';
        const snippetSource = hasEnunc ? textoItens(q.enunItems) : textoItens(q.alternatives[0].items);
        errors.push({ reason, snippet: textSnippet(snippetSource) });
    }

    const novaQuestao = titulo => ({ enunItems: [], alternatives: [], images: [], imgRidToName: {}, ref: { enun: refVazia() }, titulo });

    let currentQ = null;
    let hasEnunciado = false;
    const titulosUsados = new Set();

    // Ignora tabelas aninhadas (de apoio, dentro de célula) e junta tabelas consecutivas
    // que pertencem à mesma questão (ex: quebra de página no Word dividiu a tabela original)
    const tables = Array.from(docXml.getElementsByTagNameNS(W, 'tbl'))
        .filter(tbl => !isNestedTable(tbl, docXml));

    for (const tbl of tables) {
        let titulo = tituloAntes(tbl);
        if (titulo && titulosUsados.has(titulo.el)) titulo = null;
        const rows = Array.from(tbl.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W && n.localName === 'tr');

        for (const row of rows) {
            const cells = Array.from(row.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W && n.localName === 'tc');
            if (cells.length < 2) continue;

            const label = cells[0].textContent.trim();

            if (isEnunciadoLabel(label)) {
                if (currentQ) finalizeQuestion(currentQ);
                currentQ = novaQuestao(titulo);
                if (titulo) titulosUsados.add(titulo.el);
                hasEnunciado = true;
            } else if (!currentQ) {
                currentQ = novaQuestao(null);
            }

            // Conteúdo da linha: parágrafos e tabelas com a formatação do Word,
            // e uma leitura à parte do texto, usada no relatório para conferir o XML
            const items = [];
            const ref = refVazia();
            for (const cell of cells.slice(1)) {
                lerConteudo(cell, items);
                extrairTextoWord(cell, ref);
            }

            if (isEnunciadoLabel(label)) {
                currentQ.enunItems = items;
                currentQ.ref.enun = ref;
            } else if (/^CORRETA/i.test(label)) {
                hasEnunciado = false;
                currentQ.alternatives.push({ items, ref, fraction: 100 });
            } else if (/^Incorreta/i.test(label)) {
                hasEnunciado = false;
                currentQ.alternatives.push({ items, ref, fraction: 0 });
            } else if (hasEnunciado) {
                // Linhas sem label reconhecido entre ENUNCIADO e CORRETA/Incorreta
                // (ex: algarismos romanos I, II, III, IV em linhas separadas, ou continuação
                // do enunciado numa tabela seguinte por causa de quebra de página)
                currentQ.enunItems.push(...items);
                somarRef(currentQ.ref.enun, ref);
            }
        }
    }
    finalizeQuestion(currentQ);

    for (const q of questions) await montarQuestao(q, imgMap);

    const revisoes = {
        ins: docXml.getElementsByTagNameNS(W, 'ins').length,
        del: docXml.getElementsByTagNameNS(W, 'del').length,
        comentarios: docXml.getElementsByTagNameNS(W, 'commentReference').length
    };
    return { questions, errors, doc: { revisoes } };
}

// Gera o HTML final da questão: espaçamento do Word, equações como imagem e imagens anexadas
async function montarQuestao(q, imgMap) {
    q.midiaDescartada = [];
    q.eqFalhas = [];
    let eqN = 0;
    const registrar = (name, b64, mime) => q.images.push({ name, b64, mime });
    const finalizar = async html => {
        const r = await resolverEquacoes(html, () => `equacao_${++eqN}.png`, registrar);
        q.eqFalhas.push(...r.falhas);
        return resolverImagens(r.html, q, imgMap);
    };
    q.enuncHtml = await finalizar(montarEnunciado(q.enunItems, { minGap: FONT_SINGLE }));
    for (const alt of q.alternatives) alt.html = await finalizar(montarAlternativa(alt.items));
}

function resolverImagens(html, q, imgMap) {
    return html.replace(/<img data-rid="([^"]*)" data-cx="(\d*)">/g, (tag, rid, cx) => {
        const imgInfo = imgMap[rid];
        if (!imgInfo || /^(emf|wmf)$/.test(imgInfo.ext)) {
            q.midiaDescartada.push(imgInfo ? 'emf' : 'ausente');
            return '';
        }
        if (!q.imgRidToName[rid]) {
            const idx = q.images.length;
            const imgName = idx === 0 ? `image.${imgInfo.ext}` : `image (${idx}).${imgInfo.ext}`;
            q.images.push({ name: imgName, b64: imgInfo.b64, mime: imgInfo.mime });
            q.imgRidToName[rid] = imgName;
        }
        const imgName = q.imgRidToName[rid];
        const width = +cx ? ` width="${Math.round(cx / 9525)}"` : '';
        return `<img src="@@PLUGINFILE@@/${encodeURIComponent(imgName)}" alt="" role="presentation" class="img-fluid"${width}>`;
    });
}

function fileTagsFor(html, images) {
    if (!images || images.length === 0) return "";
    let tags = "";
    images.forEach(img => {
        if (html.includes(`@@PLUGINFILE@@/${encodeURIComponent(img.name)}`)) {
            tags += `      <file name="${escXml(img.name)}" path="/" encoding="base64">${img.b64}</file>\n`;
        }
    });
    return tags;
}

function buildXML(questions) {
    let x = `<?xml version="1.0" encoding="UTF-8"?>\n<quiz>\n`;
    questions.forEach((q, i) => {
        const num = (i + 1).toString().padStart(2, '0');
        x += `  <question type="multichoice">\n`;
        x += `    <name>\n      <text>Q${num}</text>\n    </name>\n`;
        x += `    <questiontext format="html">\n      <text><![CDATA[${q.enuncHtml}]]></text>\n${fileTagsFor(q.enuncHtml, q.images)}    </questiontext>\n`;
        x += `    <generalfeedback format="html">\n      <text></text>\n    </generalfeedback>\n`;
        x += `    <defaultgrade>1.0000000</defaultgrade>\n`;
        x += `    <penalty>0.3333333</penalty>\n`;
        x += `    <hidden>0</hidden>\n`;
        x += `    <idnumber></idnumber>\n`;
        x += `    <single>true</single>\n`;
        x += `    <shuffleanswers>true</shuffleanswers>\n`;
        x += `    <answernumbering>abc</answernumbering>\n`;
        x += `    <showstandardinstruction>0</showstandardinstruction>\n`;
        x += `    <correctfeedback format="html">\n      <text>Sua resposta est&#225; correta.</text>\n    </correctfeedback>\n`;
        x += `    <partiallycorrectfeedback format="html">\n      <text>Sua resposta est&#225; parcialmente correta.</text>\n    </partiallycorrectfeedback>\n`;
        x += `    <incorrectfeedback format="html">\n      <text>Sua resposta est&#225; incorreta.</text>\n    </incorrectfeedback>\n`;
        x += `    <shownumcorrect/>\n`;
        q.alternatives.forEach(alt => {
            x += `    <answer fraction="${alt.fraction}" format="html">\n`;
            x += `      <text><![CDATA[${alt.html}]]></text>\n`;
            x += fileTagsFor(alt.html, q.images);
            x += `      <feedback format="html">\n        <text></text>\n      </feedback>\n`;
            x += `    </answer>\n`;
        });
        x += `  </question>\n`;
    });
    return x + `</quiz>`;
}

async function doConvert() {
    const btn = document.getElementById('bcvt');
    const btnText = document.getElementById('btn-text');
    btn.disabled = true;
    btnText.textContent = 'CONVERTENDO…';
    try {
        const resultados = [];
        for (const f of filesData) {
            const r = await parseDOCX(f.buffer);
            resultados.push({ nome: f.name, ...r, xml: "﻿" + buildXML(r.questions) });
        }
        if (resultados.length === 1) {
            const r = resultados[0];
            const blob = new Blob([r.xml], {type: 'text/xml;charset=utf-8'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = r.nome.replace('.docx', '.xml');
            a.click();
        } else {
            const zip = new JSZip();
            resultados.forEach(r => zip.file(r.nome.replace('.docx','.xml'), r.xml));
            const blob = await zip.generateAsync({type:'blob'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = "questoes_moodle.zip";
            a.click();
        }
        const total = resultados.reduce((s, r) => s + r.questions.length, 0);
        showToast(`Download de ${total} questões finalizado!`, "success");

        const relatorios = resultados.map(gerarRelatorio);
        mostrarRelatorios(relatorios);
        const comErro = relatorios.reduce((s, R) => s + R.questoes.filter(q => pior(q) >= 3).length + R.naoConvertidas, 0);
        if (comErro > 0) showToast(`${comErro} questão(ões) com erro: confira o relatório`, "error");
    } catch (e) {
        console.error(e);
        showToast("Erro na conversão", "error");
    } finally {
        btn.disabled = false;
        btnText.textContent = 'CONVERTER E BAIXAR';
    }
}
