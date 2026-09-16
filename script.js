let filesData = [];
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main', 
      A='http://schemas.openxmlformats.org/drawingml/2006/main', 
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
            document.getElementById('fl').innerHTML += `<div class="file-item"><span>📄 ${f.name}</span></div>`;
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

function textSnippet(html, maxLen = 90) {
    const text = String(html).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
    return text.length > maxLen ? text.slice(0, maxLen) + '…' : text;
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

function processLists(html) {
    // Process numbered lists (1., 2., 3. etc.) digitadas manualmente como texto
    html = html.replace(/(<span>(\d+)\.\s*([^<]*?)<\/span><br>\s*)+/g, (match) => {
        const items = match.match(/<span>(\d+)\.\s*([^<]*?)<\/span><br>/g);
        if (items && items.length > 1) {
            let ol = '<ol>';
            items.forEach(item => {
                const match = item.match(/<span>\d+\.\s*([^<]*?)<\/span><br>/);
                if (match) {
                    const text = match[1];
                    ol += `<li>${text}</li>`;
                }
            });
            ol += '</ol>';
            return ol;
        }
        return match;
    });
    return html;
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
        const path = 'word/' + target.replace(/^\//, '');
        const f = zip.file(path);
        if (f && /png|jpg|jpeg|gif/i.test(path)) {
            imgMap[rid] = { mime: 'image/'+path.split('.').pop(), b64: await f.async('base64') };
        }
    }

    const docXml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('text'), 'text/xml');
    let numberingDoc = null;
    try {
        const numberingXml = await zip.file('word/numbering.xml').async('text');
        numberingDoc = new DOMParser().parseFromString(numberingXml, 'text/xml');
    } catch (e) {
        // No numbering.xml or parse error, skip
    }
    const numMap = {};
    if (numberingDoc) {
        for (const num of numberingDoc.getElementsByTagNameNS(W, 'num')) {
            const numId = num.getAttributeNS(W, 'numId');
            const abstractNumId = num.getElementsByTagNameNS(W, 'abstractNumId')[0]?.getAttributeNS(W, 'val');
            if (abstractNumId) {
                const abstractNums = numberingDoc.getElementsByTagNameNS(W, 'abstractNum');
                for (const abs of abstractNums) {
                    if (abs.getAttributeNS(W, 'abstractNumId') === abstractNumId) {
                        const lvls = abs.getElementsByTagNameNS(W, 'lvl');
                        for (const lvl of lvls) {
                            if (lvl.getAttributeNS(W, 'ilvl') === '0') {
                                const numFmtEl = lvl.getElementsByTagNameNS(W, 'numFmt')[0];
                                const numFmt = numFmtEl?.getAttributeNS(W, 'val');
                                numMap[numId] = numFmt;
                                break;
                            }
                        }
                        break;
                    }
                }
            }
        }
    }
    const questions = [];
    const errors = [];

    function finalizeQuestion(q) {
        if (!q) return;
        const hasContent = q.enuncHtml || q.alternatives.length > 0;
        if (!hasContent) return;
        if (q.enuncHtml && q.alternatives.length > 0) {
            questions.push(q);
            return;
        }
        const reason = (!q.enuncHtml && q.alternatives.length > 0)
            ? 'alternativa(s) encontrada(s) sem enunciado (rótulo "ENUNCIADO" ausente ou não reconhecido)'
            : 'enunciado encontrado sem nenhuma alternativa "Correta"/"Incorreta"';
        const snippetSource = q.enuncHtml || (q.alternatives[0] && q.alternatives[0].html) || "";
        errors.push({ reason, snippet: textSnippet(snippetSource) });
    }

    let currentQ = null;
    let hasEnunciado = false;

    // Ignora tabelas aninhadas (de apoio, dentro de célula) e junta tabelas consecutivas
    // que pertencem à mesma questão (ex: quebra de página no Word dividiu a tabela original)
    const tables = Array.from(docXml.getElementsByTagNameNS(W, 'tbl'))
        .filter(tbl => !isNestedTable(tbl, docXml));

    for (const tbl of tables) {
        const rows = Array.from(tbl.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W && n.localName === 'tr');

        for (const row of rows) {
            const cells = Array.from(row.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W && n.localName === 'tc');
            if (cells.length < 2) continue;

            const label = cells[0].textContent.trim();

            if (isEnunciadoLabel(label)) {
                if (currentQ) finalizeQuestion(currentQ);
                currentQ = { enuncHtml: "", alternatives: [], images: [], imgRidToName: {} };
                hasEnunciado = true;
            } else if (!currentQ) {
                currentQ = { enuncHtml: "", alternatives: [], images: [], imgRidToName: {} };
            }

            let rowHtml = "";

            for (let i = 1; i < cells.length; i++) {
                const cell = cells[i];
                const paragraphs = cell.getElementsByTagNameNS(W, 'p');
                
                if (cell.textContent.trim() === "" && paragraphs.length === 0) {
                    rowHtml += "<br>";
                    continue;
                }

                let currentList = null;
                let listCounters = {};
                for (const p of paragraphs) {
                    let pText = "";
                    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
                    const spacing = pPr?.getElementsByTagNameNS(W, 'spacing')[0];
                    const hasAfterSpacing = spacing?.getAttributeNS(W, 'after');

                    for (const r of p.getElementsByTagNameNS(W, 'r')) {
                        let runHtml = "";
                        const drw = r.getElementsByTagNameNS(W, 'drawing')[0];
                        if (drw) {
                            const blip = drw.getElementsByTagNameNS(A, 'blip')[0];
                            const rid = blip?.getAttributeNS(R, 'embed');
                            if (rid && imgMap[rid]) {
                                if (!currentQ.imgRidToName[rid]) {
                                    const imgInfo = imgMap[rid];
                                    const ext = imgInfo.mime.split('/')[1];
                                    const idx = currentQ.images.length;
                                    const imgName = idx === 0 ? `image.${ext}` : `image (${idx}).${ext}`;
                                    currentQ.images.push({ name: imgName, b64: imgInfo.b64 });
                                    currentQ.imgRidToName[rid] = imgName;
                                }
                                const imgName = currentQ.imgRidToName[rid];
                                runHtml += `<img src="@@PLUGINFILE@@/${encodeURIComponent(imgName)}" alt="" role="presentation" class="img-fluid">`;
                            }
                        }
                        const bEl = r.getElementsByTagNameNS(W, 'b')[0];
                        const bold = bEl !== undefined && bEl.getAttributeNS(W, 'val') !== '0';
                        for (const child of r.childNodes) {
                            if (child.localName === 'br' && child.namespaceURI === W) {
                                runHtml += '<br>';
                            } else if (child.localName === 't' && child.namespaceURI === W) {
                                let s = escXml(child.textContent);
                                if (bold) s = `<b>${s}</b>`;
                                runHtml += s;
                            }
                        }
                        // Se o run estiver dentro de um w:hyperlink, envolve em <a href>
                        let anc = r.parentNode;
                        while (anc && anc !== p) {
                            if (anc.namespaceURI === W && anc.localName === 'hyperlink') {
                                const hRid = anc.getAttributeNS(R, 'id');
                                const target = hRid && relMap[hRid];
                                if (target && runHtml) {
                                    runHtml = `<a href="${escXml(target)}">${runHtml}</a>`;
                                }
                                break;
                            }
                            anc = anc.parentNode;
                        }
                        pText += runHtml;
                    }

                    const numPr = pPr?.getElementsByTagNameNS(W, 'numPr')[0];
                    if (numPr) {
                        const ilvl = numPr.getElementsByTagNameNS(W, 'ilvl')[0]?.getAttributeNS(W, 'val');
                        const numId = numPr.getElementsByTagNameNS(W, 'numId')[0]?.getAttributeNS(W, 'val');
                        const numFmt = numMap[numId];
                        if (ilvl === '0' && (numFmt === 'upperRoman' || numFmt === 'decimal' || numFmt === 'bullet')) {
                            if (!currentList || currentList !== numId) {
                                if (currentList) {
                                    const prevFmt = numMap[currentList];
                                    rowHtml += (prevFmt === 'bullet') ? '</ul>' : '</ol>';
                                }
                                if (!listCounters[numId]) listCounters[numId] = 0;
                                let openTag;
                                if (numFmt === 'bullet') {
                                    openTag = '<ul>';
                                } else {
                                    const listType = numFmt === 'upperRoman' ? 'I' : '1';
                                    openTag = `<ol type="${listType}" start="${listCounters[numId] + 1}">`;
                                }
                                rowHtml += openTag;
                                currentList = numId;
                            }
                            listCounters[numId] = (listCounters[numId] || 0) + 1;
                            rowHtml += `<li>${pText}</li>`;
                            continue;
                        }
                    } else {
                        if (currentList) {
                            const fmt = numMap[currentList];
                            rowHtml += (fmt === 'bullet') ? '</ul>' : '</ol>';
                            currentList = null;
                        }
                    }

                    if (pText.trim()) {
                        rowHtml += `<span>${pText}</span><br>`;
                        if (hasAfterSpacing && parseInt(hasAfterSpacing) > 0) {
                            rowHtml += "<br>";
                        }
                    } else {
                        rowHtml += "<br>";
                    }
                }
                if (currentList) {
                    const fmt = numMap[currentList];
                    const closeTag = (fmt === 'bullet') ? '</ul>' : '</ol>';
                    rowHtml += closeTag;
                }
            }

            if (isEnunciadoLabel(label)) {
                currentQ.enuncHtml = rowHtml;
            } else if (/^CORRETA/i.test(label)) {
                hasEnunciado = false;
                if (currentQ.alternatives.some(alt => alt.fraction === 100)) {
                    errors.push({
                        reason: 'questão com mais de uma alternativa "Correta" (possível mistura de duas questões)',
                        snippet: textSnippet(currentQ.enuncHtml || rowHtml)
                    });
                }
                currentQ.alternatives.push({ html: rowHtml, fraction: 100 });
            } else if (/^Incorreta/i.test(label)) {
                hasEnunciado = false;
                currentQ.alternatives.push({ html: rowHtml, fraction: 0 });
            } else if (hasEnunciado) {
                // Linhas sem label reconhecido entre ENUNCIADO e CORRETA/Incorreta
                // (ex: algarismos romanos I, II, III, IV em linhas separadas, ou continuação
                // do enunciado numa tabela seguinte por causa de quebra de página)
                currentQ.enuncHtml += rowHtml;
            }
        }
    }
    finalizeQuestion(currentQ);
    return { questions, errors };
}

function wrapAnswerHtml(html) {
    const processed = processLists(html);
    // Se já tem tags de bloco, usa direto; senão envolve em <p>
    if (/<(p|ol|ul|img)\b/i.test(processed)) return processed;
    // Converte <span>texto</span><br> → <p dir="ltr" style="text-align: left;">texto<br></p>
    return processed.replace(/<span>([\s\S]*?)<\/span><br>/g,
        '<p dir="ltr" style="text-align: left;">$1<br></p>');
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
        const enuncHtml = processLists(q.enuncHtml);
        x += `    <questiontext format="html">\n      <text><![CDATA[${enuncHtml}]]></text>\n${fileTagsFor(enuncHtml, q.images)}    </questiontext>\n`;
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
            const altHtml = wrapAnswerHtml(alt.html);
            x += `    <answer fraction="${alt.fraction}" format="html">\n`;
            x += `      <text><![CDATA[${altHtml}]]></text>\n`;
            x += fileTagsFor(altHtml, q.images);
            x += `      <feedback format="html">\n        <text></text>\n      </feedback>\n`;
            x += `    </answer>\n`;
        });
        x += `  </question>\n`;
    });
    return x + `</quiz>`;
}

function showConversionErrors(errors) {
    const multiFile = new Set(errors.map(e => e.file)).size > 1;
    const maxShown = 8;
    const shown = errors.slice(0, maxShown);
    const lines = shown.map((e, i) => {
        const prefix = multiFile ? `${escXml(e.file)} — ` : '';
        return `${i + 1}. ${prefix}${escXml(e.reason)}: "${escXml(e.snippet)}"`;
    });
    if (errors.length > maxShown) lines.push(`+ ${errors.length - maxShown} outra(s) questão(ões) com problema`);
    const t = document.createElement('div');
    t.className = 'toast error toast-list';
    t.innerHTML = `<strong>${errors.length} questão(ões) não convertida(s):</strong><br>${lines.join('<br>')}`;
    document.getElementById('toast-container').appendChild(t);
    setTimeout(() => {
        t.style.opacity = '0';
        setTimeout(() => t.remove(), 400);
    }, 12000);
}

async function doConvert() {
    const btn = document.getElementById('bcvt');
    btn.disabled = true;
    try {
        let total = 0;
        const allErrors = [];
        if (filesData.length === 1) {
            const f = filesData[0];
            const { questions, errors } = await parseDOCX(f.buffer);
            total = questions.length;
            errors.forEach(e => allErrors.push({ file: f.name, ...e }));
            const xmlString = "\ufeff" + buildXML(questions);
            const blob = new Blob([xmlString], {type: 'text/xml;charset=utf-8'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = f.name.replace('.docx', '.xml');
            a.click();
        } else {
            const zip = new JSZip();
            for (const f of filesData) {
                const { questions, errors } = await parseDOCX(f.buffer);
                total += questions.length;
                errors.forEach(e => allErrors.push({ file: f.name, ...e }));
                zip.file(f.name.replace('.docx','.xml'), "\ufeff" + buildXML(questions));
            }
            const blob = await zip.generateAsync({type:'blob'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = "questoes_moodle.zip";
            a.click();
        }
        showToast(`Download de ${total} questões finalizado!`, "success");
        if (allErrors.length > 0) showConversionErrors(allErrors);
    } catch (e) {
        console.error(e);
        showToast("Erro na conversão", "error"); 
    } finally { 
        btn.disabled = false; 
    }
}