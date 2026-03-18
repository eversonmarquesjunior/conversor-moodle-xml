let filesData = [];
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main', 
      A='http://schemas.openxmlformats.org/drawingml/2006/main', 
      R='http://schemas.openxmlformats.org/officeDocument/2006/relationships', 
      REL_NS='http://schemas.openxmlformats.org/package/2006/relationships';

const dropZone = document.getElementById('dz');

// Event Listeners para Drag and Drop
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => { 
    dropZone.addEventListener(eventName, preventDefaults, false); 
});

function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

['dragenter', 'dragover'].forEach(eventName => { 
    dropZone.addEventListener(eventName, () => dropZone.classList.add('drag-over'), false); 
});

['dragleave', 'drop'].forEach(eventName => { 
    dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'), false); 
});

dropZone.addEventListener('drop', (e) => { handleFiles(e.dataTransfer.files); }, false);
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
    const numberingXml = await zip.file('word/numbering.xml').async('text');
    const numberingDoc = new DOMParser().parseFromString(numberingXml, 'text/xml');
    const numMap = {};
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
    const questions = [];

    for (const tbl of docXml.getElementsByTagNameNS(W, 'tbl')) {
        let currentQ = { enuncHtml: "", alternatives: [] };
        const rows = tbl.getElementsByTagNameNS(W, 'tr');
        
        for (const row of rows) {
            const cells = row.getElementsByTagNameNS(W, 'tc');
            if (cells.length < 2) continue;

            const label = cells[0].textContent.trim();
            let rowHtml = "";

            for (let i = 1; i < cells.length; i++) {
                const cell = cells[i];
                const paragraphs = cell.getElementsByTagNameNS(W, 'p');
                
                if (cell.textContent.trim() === "" && paragraphs.length === 0) {
                    rowHtml += "<br>";
                    continue;
                }

                let currentList = null;
                let listItems = [];
                for (const p of paragraphs) {
                    let pText = "";
                    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
                    const spacing = pPr?.getElementsByTagNameNS(W, 'spacing')[0];
                    const hasAfterSpacing = spacing?.getAttributeNS(W, 'after');

                    for (const r of p.getElementsByTagNameNS(W, 'r')) {
                        const drw = r.getElementsByTagNameNS(W, 'drawing')[0];
                        if (drw) {
                            const blip = drw.getElementsByTagNameNS(A, 'blip')[0];
                            const rid = blip?.getAttributeNS(R, 'embed');
                            if (rid && imgMap[rid]) pText += `<img src="data:${imgMap[rid].mime};base64,${imgMap[rid].b64}" style="max-width:100%">`;
                        }
                        const t = r.getElementsByTagNameNS(W, 't')[0];
                        if (t) {
                            let s = escXml(t.textContent);
                            if (r.getElementsByTagNameNS(W, 'b').length) s = `<b>${s}</b>`;
                            pText += s;
                        }
                    }

                    const numPr = pPr?.getElementsByTagNameNS(W, 'numPr')[0];
                    if (numPr) {
                        const ilvl = numPr.getElementsByTagNameNS(W, 'ilvl')[0]?.getAttributeNS(W, 'val');
                        const numId = numPr.getElementsByTagNameNS(W, 'numId')[0]?.getAttributeNS(W, 'val');
                        const numFmt = numMap[numId];
                        if (ilvl === '0' && numFmt === 'upperRoman') {
                            if (!currentList || currentList !== numId) {
                                if (currentList) {
                                    rowHtml += `</ol>`;
                                    listItems = [];
                                }
                                rowHtml += `<ol type="I" start="1">`;
                                currentList = numId;
                            }
                            rowHtml += `<li>${pText}</li>`;
                            listItems.push(pText);
                            continue;
                        }
                    } else {
                        if (currentList) {
                            rowHtml += `</ol>`;
                            currentList = null;
                            listItems = [];
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
                    rowHtml += `</ol>`;
                }
            }

            if (/ENUNCIADO/i.test(label)) { currentQ.enuncHtml = rowHtml; } 
            else if (/^CORRETA/i.test(label)) { currentQ.alternatives.push({ html: rowHtml, fraction: 100 }); } 
            else if (/^Incorreta/i.test(label)) { currentQ.alternatives.push({ html: rowHtml, fraction: 0 }); }
        }
        if (currentQ.enuncHtml && currentQ.alternatives.length > 0) questions.push(currentQ);
    }
    return questions;
}

function buildXML(questions) {
    let x = `<?xml version="1.0" encoding="UTF-8"?>\n<quiz>\n`;
    questions.forEach((q, i) => {
        x += `<question type="multichoice">\n<name><text>Q${(i+1).toString().padStart(2,'0')}</text></name>\n`;
        x += `<questiontext format="html"><text><![CDATA[${q.enuncHtml}]]></text></questiontext>\n`;
        x += `<single>true</single><shuffleanswers>true</shuffleanswers><answernumbering>abc</answernumbering>\n`;
        q.alternatives.forEach(alt => {
            x += `<answer fraction="${alt.fraction}" format="html"><text><![CDATA[${alt.html}]]></text></answer>\n`;
        });
        x += `</question>\n`;
    });
    return x + `</quiz>`;
}

async function doConvert() {
    const btn = document.getElementById('bcvt');
    btn.disabled = true;
    try {
        let total = 0;
        if (filesData.length === 1) {
            const f = filesData[0];
            const qs = await parseDOCX(f.buffer);
            total = qs.length;
            const xmlString = "\ufeff" + buildXML(qs);
            const blob = new Blob([xmlString], {type: 'text/xml;charset=utf-8'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = f.name.replace('.docx', '.xml');
            a.click();
        } else {
            const zip = new JSZip();
            for (const f of filesData) {
                const qs = await parseDOCX(f.buffer);
                total += qs.length;
                zip.file(f.name.replace('.docx','.xml'), "\ufeff" + buildXML(qs));
            }
            const blob = await zip.generateAsync({type:'blob'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = "questoes_moodle.zip";
            a.click();
        }
        showToast(`Download de ${total} questões finalizado!`, "success");
    } catch (e) { 
        showToast("Erro na conversão", "error"); 
    } finally { 
        btn.disabled = false; 
    }
}