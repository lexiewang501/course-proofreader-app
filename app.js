/**
 * 課程資料校稿小工具 (Course Proofreader Web App)
 * 核心業務邏輯：依據嚴格欄位命名、抽取隔離與容錯規則進行 Word 與 PDF 智慧校對
 */

// Configure PDF.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.js';
}

// Exact section labels strictly matching Word & PDF document naming (禁止更名)
const KNOWN_LABELS = [
    '課程目標',
    '適合對象',
    '先修課程',
    '預備知識',
    '課程內容',
    '學會技能',
    '備註事項',
    '後續推薦課程'
];

// Global Application State
const state = {
    wordFile: null,
    wordBuffer: null,
    wordCourses: [],
    
    pdfFile: null,
    pdfBuffer: null,
    pdfCourses: [],
    
    comparisons: [],
    activeFilter: 'all',
    searchQuery: ''
};

// DOM Elements Cache
const elements = {
    wordInput: document.getElementById('wordFileInput'),
    wordDropzone: document.getElementById('wordDropzone'),
    wordEmptyView: document.getElementById('wordEmptyView'),
    wordLoadedView: document.getElementById('wordLoadedView'),
    wordFileName: document.getElementById('wordFileName'),
    wordFileMeta: document.getElementById('wordFileMeta'),

    pdfInput: document.getElementById('pdfFileInput'),
    pdfDropzone: document.getElementById('pdfDropzone'),
    pdfEmptyView: document.getElementById('pdfEmptyView'),
    pdfLoadedView: document.getElementById('pdfLoadedView'),
    pdfFileName: document.getElementById('pdfFileName'),
    pdfFileMeta: document.getElementById('pdfFileMeta'),

    btnStartCompare: document.getElementById('btnStartCompare'),
    btnSampleBlockchain: document.getElementById('btnSampleBlockchain'),
    btnSampleJianzhen: document.getElementById('btnSampleJianzhen'),

    loadingIndicator: document.getElementById('loadingIndicator'),
    loadingText: document.getElementById('loadingText'),
    resultsSection: document.getElementById('resultsSection'),
    courseCardsContainer: document.getElementById('courseCardsContainer'),
    emptyFilterView: document.getElementById('emptyFilterView'),

    // KPI Summary
    kpiTotal: document.getElementById('kpiTotal'),
    kpiMatch: document.getElementById('kpiMatch'),
    kpiError: document.getElementById('kpiError'),
    kpiWarning: document.getElementById('kpiWarning'),
    kpiMissing: document.getElementById('kpiMissing'),

    // Filter counts
    countAll: document.getElementById('countAll'),
    countRed: document.getElementById('countRed'),
    countYellow: document.getElementById('countYellow'),
    countGreen: document.getElementById('countGreen'),
    countGray: document.getElementById('countGray'),

    searchInput: document.getElementById('searchInput'),
    btnCopyReport: document.getElementById('btnCopyReport'),
    btnExportCSV: document.getElementById('btnExportCSV'),
    toast: document.getElementById('toast'),
    toastMessage: document.getElementById('toastMessage')
};

// ==========================================
// Initialization & Event Listeners
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    setupUploadHandlers();
    setupFilterHandlers();
    setupActionHandlers();
});

function setupUploadHandlers() {
    setupDropzone(elements.wordDropzone, elements.wordInput, (file) => handleWordFile(file));
    setupDropzone(elements.pdfDropzone, elements.pdfInput, (file) => handlePdfFile(file));

    elements.btnSampleBlockchain.addEventListener('click', () => loadSample('blockchain'));
    elements.btnSampleJianzhen.addEventListener('click', () => loadSample('jianzhen'));
    elements.btnStartCompare.addEventListener('click', runComparison);
}

function setupDropzone(dropzone, input, onFileSelected) {
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        });
    });

    dropzone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            onFileSelected(files[0]);
        }
    });

    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            onFileSelected(e.target.files[0]);
        }
    });
}

function setupFilterHandlers() {
    document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.filter-btn').forEach(b => {
                b.classList.remove('active', 'bg-slate-900', 'text-white');
                b.classList.add('bg-slate-100', 'text-slate-600');
            });
            btn.classList.add('active', 'bg-slate-900', 'text-white');
            btn.classList.remove('bg-slate-100', 'text-slate-600');
            state.activeFilter = btn.dataset.filter;
            renderCourseCards();
        });
    });

    elements.searchInput.addEventListener('input', (e) => {
        state.searchQuery = e.target.value.trim().toLowerCase();
        renderCourseCards();
    });
}

function setupActionHandlers() {
    elements.btnCopyReport.addEventListener('click', copyErrorReport);
    elements.btnExportCSV.addEventListener('click', exportCSVReport);
}

// ==========================================
// File Loaders
// ==========================================
async function handleWordFile(file) {
    if (!file.name.endsWith('.docx')) {
        showToast('請上傳 .docx 格式的 Word 檔案！', true);
        return;
    }
    state.wordFile = file;
    state.wordBuffer = await file.arrayBuffer();

    elements.wordEmptyView.classList.add('hidden');
    elements.wordLoadedView.classList.remove('hidden');
    elements.wordFileName.textContent = file.name;
    elements.wordFileMeta.textContent = `檔案大小：${(file.size / 1024).toFixed(1)} KB (已準備就緒)`;

    checkReadyToCompare();
    showToast(`已載入 Word 檔案：${file.name}`);
}

async function handlePdfFile(file) {
    if (!file.name.endsWith('.pdf')) {
        showToast('請上傳 .pdf 格式的 PDF 檔案！', true);
        return;
    }
    state.pdfFile = file;
    state.pdfBuffer = await file.arrayBuffer();

    elements.pdfEmptyView.classList.add('hidden');
    elements.pdfLoadedView.classList.remove('hidden');
    elements.pdfFileName.textContent = file.name;
    elements.pdfFileMeta.textContent = `檔案大小：${(file.size / 1024).toFixed(1)} KB (已準備就緒)`;

    checkReadyToCompare();
    showToast(`已載入 PDF 檔案：${file.name}`);
}

function checkReadyToCompare() {
    if (state.wordBuffer && state.pdfBuffer) {
        elements.btnStartCompare.disabled = false;
        elements.btnStartCompare.classList.add('ring-4', 'ring-blue-500/30', 'animate-pulse');
        setTimeout(() => elements.btnStartCompare.classList.remove('animate-pulse'), 1500);
    }
}

async function loadSample(type) {
    const wordPath = type === 'blockchain' ? '恆逸_區塊鏈_2027年1-6月課程.docx' : '恆逸_鑒真數位_2027年1-6月課程v1_1.docx';
    const pdfPath = type === 'blockchain' ? '區塊鏈.pdf' : '鑒真數位.pdf';

    try {
        elements.loadingIndicator.classList.remove('hidden');
        elements.loadingText.textContent = `正在讀取本地測試檔案 (${type === 'blockchain' ? '區塊鏈' : '鑒真數位'})...`;

        const [wordRes, pdfRes] = await Promise.all([
            fetch(encodeURIComponent(wordPath)),
            fetch(encodeURIComponent(pdfPath))
        ]);

        const [wordBlob, pdfBlob] = await Promise.all([
            wordRes.blob(),
            pdfRes.blob()
        ]);

        const wordFile = new File([wordBlob], wordPath, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const pdfFile = new File([pdfBlob], pdfPath, { type: 'application/pdf' });

        await handleWordFile(wordFile);
        await handlePdfFile(pdfFile);

        elements.loadingIndicator.classList.add('hidden');
        showToast(`成功載入樣本檔案！點選右側按鈕即可開始校稿`);
    } catch (err) {
        console.error(err);
        elements.loadingIndicator.classList.add('hidden');
        showToast('載入樣本失敗，請手動拖入檔案！', true);
    }
}

// ==========================================
// Parsing Engines (Word & PDF)
// ==========================================

/**
 * Parses Word (.docx) file extracting clean table data without deleted (strikethrough) items.
 */
async function parseDocx(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const xmlFile = zip.file('word/document.xml');
    if (!xmlFile) throw new Error('無效的 Word 檔案 (找不到 word/document.xml)');

    const xmlString = await xmlFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, 'application/xml');

    const tables = doc.getElementsByTagName('w:tbl');
    const courses = [];

    for (let t = 0; t < tables.length; t++) {
        const tbl = tables[t];
        const trList = tbl.getElementsByTagName('w:tr');
        const rows = [];

        for (let r = 0; r < trList.length; r++) {
            const tr = trList[r];
            const tcList = tr.getElementsByTagName('w:tc');
            const cells = [];

            for (let c = 0; c < tcList.length; c++) {
                const tc = tcList[c];
                // Extract clean text, ignoring any text inside strikethrough <w:strike> or <w:dstrike>
                const pList = tc.getElementsByTagName('w:p');
                const paragraphs = [];

                for (let pi = 0; pi < pList.length; pi++) {
                    const p = pList[pi];
                    const rList = p.getElementsByTagName('w:r');
                    let pText = '';

                    for (let ri = 0; ri < rList.length; ri++) {
                        const run = rList[ri];
                        const rPr = run.getElementsByTagName('w:rPr')[0];
                        if (rPr) {
                            if (rPr.getElementsByTagName('w:strike').length > 0 ||
                                rPr.getElementsByTagName('w:dstrike').length > 0) {
                                continue; // Skip struck-through text!
                            }
                        }
                        const tList = run.getElementsByTagName('w:t');
                        for (let ti = 0; ti < tList.length; ti++) {
                            pText += tList[ti].textContent;
                        }
                    }
                    if (pText.trim().length > 0) {
                        paragraphs.push(pText.trim());
                    }
                }

                cells.push({
                    fullText: paragraphs.join(' ').trim(),
                    paragraphs: paragraphs
                });
            }
            rows.push(cells);
        }

        const fullTableText = rows.map(r => r.map(c => c.fullText).join(' ')).join('\n');
        if (!fullTableText.includes('時數') && !fullTableText.includes('費用') && !fullTableText.includes('點數')) {
            continue;
        }

        // Exact schema mapping
        const course = {
            '課程代碼': '',
            '課程名稱': '',
            '英文名稱': '',
            '時數': '',
            '費用': '',
            '點數': '',
            '教材': '',
            '課程目標': '',
            '適合對象': '',
            '先修課程': '',
            '預備知識': '',
            '課程內容': [], // List[str]
            '學會技能': '',
            '備註事項': '',
            '後續推薦課程': '',
            source: 'Word'
        };

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowFullText = row.map(c => c.fullText).join(' ').replace(/\s+/g, ' ').trim();
            const firstCellText = row[0] ? row[0].fullText : '';

            if (i === 0) {
                // Course Code & Chinese Name
                if (row.length >= 2 && row[0].fullText.length <= 10 && row[0].fullText.length > 0) {
                    course['課程代碼'] = row[0].fullText.trim();
                    course['課程名稱'] = row.slice(1).map(c => c.fullText).join(' ').replace(/^[：:\s|]+/, '').trim();
                } else {
                    const parts = rowFullText.split(/[\s|：:]+/);
                    if (parts.length >= 2 && parts[0].length <= 10) {
                        course['課程代碼'] = parts[0].trim();
                        course['課程名稱'] = rowFullText.replace(parts[0], '').replace(/^[|：:\s]+/, '').trim();
                    } else {
                        course['課程名稱'] = rowFullText;
                    }
                }
            } else if (i === 1 && !rowFullText.includes('時數') && !rowFullText.includes('費用')) {
                course['英文名稱'] = rowFullText.replace(/^[|：:\s]+/, '').trim();
            }

            // Metadata row
            if (rowFullText.includes('時數') || rowFullText.includes('費用') || rowFullText.includes('點數')) {
                const hoursMatch = rowFullText.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
                if (hoursMatch) course['時數'] = hoursMatch[1];

                const priceMatch = rowFullText.match(/費用[：:\s]*([0-9,]+)\s*元?/);
                if (priceMatch) course['費用'] = priceMatch[1].replace(/,/g, '');

                const pointsMatch = rowFullText.match(/點數[：:\s]*([0-9.]+)\s*點?/);
                if (pointsMatch) course['點數'] = pointsMatch[1];

                const matMatch = rowFullText.match(/教材[：:\s]*([^|｜\n]+)/);
                if (matMatch) course['教材'] = matMatch[1].trim();
            }

            // Check known section labels
            for (const label of KNOWN_LABELS) {
                if (firstCellText === label || firstCellText.startsWith(label)) {
                    if (label === '課程內容') {
                        // Gather paragraphs from second cell or following cells
                        let pItems = [];
                        if (row.length >= 2) {
                            pItems = row[1].paragraphs;
                        }
                        if (pItems.length <= 1) {
                            // Split if needed
                            const fullContent = row.slice(1).map(c => c.fullText).join(' ');
                            pItems = splitOutlineItems(fullContent);
                        }
                        course['課程內容'] = pItems;
                    } else {
                        const val = row.slice(1).map(c => c.fullText).join(' ').trim();
                        course[label] = val;
                    }
                } else if (rowFullText.startsWith(label) && !course[label]) {
                    const val = rowFullText.replace(new RegExp(`^${label}[：:\\s]*`), '').trim();
                    if (label === '課程內容') {
                        course['課程內容'] = splitOutlineItems(val);
                    } else {
                        course[label] = val;
                    }
                }
            }
        }

        // Clean code if prefixed to name
        if (!course['課程代碼'] && course['課程名稱']) {
            const m = course['課程名稱'].match(/^([A-Za-z0-9_-]{2,10})\s+(.*)$/);
            if (m) {
                course['課程代碼'] = m[1];
                course['課程名稱'] = m[2];
            }
        }

        // If course code is empty and all fields empty (e.g. fully struck out course), skip
        if (course['課程名稱'] || course['課程代碼']) {
            courses.push(course);
        }
    }

    return courses;
}

/**
 * Parses PDF file with strict spatial bounding box isolation to prevent field bleeding (張冠李戴).
 */
async function parsePdf(buffer) {
    const doc = await pdfjsLib.getDocument({
        data: new Uint8Array(buffer),
        useSystemFonts: true,
        disableFontFace: true
    }).promise;

    const courses = [];

    for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const textContent = await page.getTextContent();
        const items = textContent.items.map(it => ({
            x: Math.round(it.transform[4]),
            y: Math.round(it.transform[5]),
            str: it.str.trim()
        })).filter(it => it.str.length > 0);

        // 1. Extract vector rectangles to determine exact table cell bounds
        const opList = await page.getOperatorList();
        const rects = [];
        for (let i = 0; i < opList.fnArray.length; i++) {
            if (opList.fnArray[i] === pdfjsLib.OPS.constructPath) {
                const subOps = opList.argsArray[i][0];
                const subArgs = opList.argsArray[i][1];
                let subArgIdx = 0;
                for (const subOp of subOps) {
                    if (subOp === pdfjsLib.OPS.rectangle) {
                        const rx = Math.round(subArgs[subArgIdx]);
                        const ry = Math.round(subArgs[subArgIdx + 1]);
                        const rw = Math.round(subArgs[subArgIdx + 2]);
                        const rh = Math.round(subArgs[subArgIdx + 3]);
                        const top = rh < 0 ? ry : ry + rh;
                        const bottom = rh < 0 ? ry + rh : ry;
                        rects.push({ x: rx, top, bottom, w: rw, h: Math.abs(rh) });
                        subArgIdx += 4;
                    } else if (subOp === pdfjsLib.OPS.moveTo || subOp === pdfjsLib.OPS.lineTo) {
                        subArgIdx += 2;
                    }
                }
            }
        }

        const labelRects = rects.filter(r => r.x < 100 && r.w >= 40 && r.w <= 80 && r.h >= 10);
        labelRects.sort((a, b) => b.top - a.top);

        // 2. Find metadata anchor lines ("時數：")
        const metaList = [];
        for (const it of items) {
            if (it.str.includes('時數：') || it.str === '時數：' || (it.str === '時數' && it.x > 80)) {
                if (!metaList.some(m => Math.abs(m.y - it.y) <= 5)) {
                    metaList.push({ y: it.y });
                }
            }
        }
        metaList.sort((a, b) => b.y - a.y);

        // 3. Find all label text items (x < 100)
        const labelTextItems = items.filter(it => it.x < 100 && KNOWN_LABELS.includes(it.str));
        labelTextItems.sort((a, b) => b.y - a.y);

        // Process each course block
        for (let i = 0; i < metaList.length; i++) {
            const metaY = metaList[i].y;
            const nextMetaY = i + 1 < metaList.length ? metaList[i + 1].y : 0;
            const prevMetaY = i > 0 ? metaList[i - 1].y : 9999;

            // Labels belonging to this course are strictly between metaY and nextMetaY
            const courseLabelItems = labelTextItems.filter(l => l.y < metaY - 5 && (nextMetaY === 0 || l.y > nextMetaY));
            courseLabelItems.sort((a, b) => b.y - a.y);

            // Pair each label with its exact vertical bounding box [bottom, top]
            const courseSections = [];
            for (let j = 0; j < courseLabelItems.length; j++) {
                const lItem = courseLabelItems[j];
                const matchedRect = labelRects.find(r => lItem.y >= r.bottom - 4 && lItem.y <= r.top + 4);
                let top, bottom;
                if (matchedRect) {
                    top = matchedRect.top;
                    bottom = matchedRect.bottom;
                } else {
                    const prevY = j > 0 ? courseLabelItems[j - 1].y : metaY - 5;
                    const nextY = j + 1 < courseLabelItems.length ? courseLabelItems[j + 1].y : (nextMetaY > 0 ? nextMetaY + 30 : 0);
                    top = (prevY + lItem.y) / 2;
                    bottom = (lItem.y + nextY) / 2;
                }
                courseSections.push({ label: lItem.str, top, bottom, y: lItem.y });
            }

            // Determine vertical course header range
            let courseTop = 9999;
            if (i > 0) {
                const prevLabels = labelTextItems.filter(l => l.y < prevMetaY - 5 && l.y > metaY);
                if (prevLabels.length > 0) {
                    const lastL = prevLabels[prevLabels.length - 1];
                    const r = labelRects.find(r => lastL.y >= r.bottom - 4 && lastL.y <= r.top + 4);
                    courseTop = r ? r.bottom - 2 : lastL.y - 15;
                } else {
                    courseTop = (prevMetaY + metaY) / 2;
                }
            }

            const course = {
                '課程代碼': '',
                '課程名稱': '',
                '英文名稱': '',
                '時數': '',
                '費用': '',
                '點數': '',
                '教材': '',
                '課程目標': '',
                '適合對象': '',
                '先修課程': '',
                '預備知識': '',
                '課程內容': [], // List[str]
                '學會技能': '',
                '備註事項': '',
                '後續推薦課程': '',
                source: 'PDF',
                page: p
            };

            // Metadata Line items
            const metaLineItems = items.filter(it => it.y >= metaY - 5 && it.y <= metaY + 5);
            metaLineItems.sort((a, b) => a.x - b.x);
            const metaLineText = metaLineItems.map(it => it.str).join(' ');

            const hoursMatch = metaLineText.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
            if (hoursMatch) course['時數'] = hoursMatch[1];

            const priceMatch = metaLineText.match(/費用[：:\s]*([0-9,]+)\s*元?/);
            if (priceMatch) course['費用'] = priceMatch[1].replace(/,/g, '');

            const pointsMatch = metaLineText.match(/點數[：:\s]*([0-9.]+)\s*點?/);
            if (pointsMatch) course['點數'] = pointsMatch[1];

            const matMatch = metaLineText.match(/教材[：:\s]*([^｜|\n]+)/);
            if (matMatch) course['教材'] = matMatch[1].trim();

            // Header lines (above metaY + 5 and below courseTop)
            const headerItems = items.filter(it => it.y > metaY + 5 && it.y <= courseTop);
            headerItems.sort((a, b) => b.y - a.y || a.x - b.x);

            const headerLines = [];
            let curY = null;
            let curLine = [];
            for (const it of headerItems) {
                if (curY === null || Math.abs(curY - it.y) > 4) {
                    if (curLine.length) headerLines.push(curLine.join(' '));
                    curY = it.y;
                    curLine = [it.str];
                } else {
                    curLine.push(it.str);
                }
            }
            if (curLine.length) headerLines.push(curLine.join(' '));

            for (const hLine of headerLines) {
                const clean = hLine.trim();
                if (clean.includes('課程簡介') || clean.includes('各地開課時間') || clean.match(/^\|\s*.*\s*\|$/)) continue;

                const codeMatch = clean.match(/\b([A-Za-z0-9_-]{2,10})\b/);
                if (codeMatch && !course['課程代碼'] && !['APP', 'DApp', 'Web3', 'EVM', 'Full', 'Stack', 'Course'].includes(codeMatch[1])) {
                    course['課程代碼'] = codeMatch[1];
                }

                if (/^[A-Za-z0-9\s,&.:()'-]+$/.test(clean) && clean.length > 5) {
                    course['英文名稱'] = clean.replace(course['課程代碼'], '').trim();
                } else if (/[\u4e00-\u9fa5]/.test(clean)) {
                    if (!course['課程名稱']) {
                        course['課程名稱'] = clean.replace(course['課程代碼'], '').trim();
                    } else {
                        course['課程名稱'] += ' ' + clean.replace(course['課程代碼'], '').trim();
                    }
                }
            }

            if (course['課程名稱'] && course['課程代碼']) {
                course['課程名稱'] = course['課程名稱'].replace(new RegExp(`^${course['課程代碼']}[：:\\s]*`), '').trim();
            }

            // Extract each section strictly isolated inside its vertical box [bottom, top]
            for (const sec of courseSections) {
                const secItems = items.filter(it => it.x >= 100 && it.y >= sec.bottom - 2 && it.y <= sec.top + 2);
                secItems.sort((a, b) => b.y - a.y || a.x - b.x);

                if (sec.label === '課程內容') {
                    // Extract items strictly as List[str]
                    course['課程內容'] = extractBulletList(secItems);
                } else {
                    // Strictly isolate 備註事項, 後續推薦課程, 適合對象, 預備知識, etc.
                    const text = secItems.map(it => it.str).join(' ').replace(/\s+/g, ' ').trim();
                    course[sec.label] = text;
                }
            }

            courses.push(course);
        }
    }

    return courses;
}

function extractBulletList(items) {
    if (!items || items.length === 0) return [];
    // Handle multi-column layouts by sorting column 1 then column 2
    const col1 = items.filter(it => it.x < 250).sort((a, b) => b.y - a.y || a.x - b.x);
    const col2 = items.filter(it => it.x >= 250).sort((a, b) => b.y - a.y || a.x - b.x);

    let fullText = '';
    if (col1.length > 0 && col2.length > 0 && Math.abs(col1[0].x - col2[0].x) > 100) {
        fullText = col1.map(it => it.str).join(' ') + ' ' + col2.map(it => it.str).join(' ');
    } else {
        fullText = items.map(it => it.str).join(' ');
    }

    const bulletList = [];
    const parts = fullText.split(/(?=\b\d+\.|\s*•\s*)/);
    for (const part of parts) {
        const clean = part.replace(/^[\s•\d.]+/, '').trim();
        if (clean.length > 0) {
            bulletList.push(clean);
        }
    }
    return bulletList.length > 0 ? bulletList : [fullText.trim()];
}

function splitOutlineItems(text) {
    if (!text) return [];
    const items = [];
    const parts = text.split(/(?=\b\d+\.|\s*•\s*)/);
    for (const part of parts) {
        const clean = part.replace(/^[\s•\d.]+/, '').trim();
        if (clean.length > 0) {
            items.push(clean);
        }
    }
    return items.length > 0 ? items : [text.trim()];
}

// ==========================================
// Comparison & Traffic Light Engine
// ==========================================

async function runComparison() {
    try {
        elements.loadingIndicator.classList.remove('hidden');
        elements.resultsSection.classList.add('hidden');
        elements.loadingText.textContent = '正在解析 Word 原稿表格資料...';

        state.wordCourses = await parseDocx(state.wordBuffer);

        elements.loadingText.textContent = '正在解析 PDF 排版文字與章節區塊...';
        state.pdfCourses = await parsePdf(state.pdfBuffer);

        elements.loadingText.textContent = '正在執行逐欄智慧校對與容錯比對...';
        state.comparisons = compareCourseData(state.wordCourses, state.pdfCourses);

        updateKPIDashboard();
        renderCourseCards();

        elements.loadingIndicator.classList.add('hidden');
        elements.resultsSection.classList.remove('hidden');

        elements.resultsSection.scrollIntoView({ behavior: 'smooth' });
        showToast('校稿比對完成！請檢視下方詳細報告');
    } catch (err) {
        console.error(err);
        elements.loadingIndicator.classList.add('hidden');
        alert('解析比對失敗：' + err.message);
    }
}

function compareCourseData(wordCourses, pdfCourses) {
    const results = [];
    const matchedPdfIndices = new Set();

    for (const wCourse of wordCourses) {
        let matchedPdf = null;
        let matchedIdx = -1;

        // 1. Match by Course Code
        if (wCourse['課程代碼']) {
            const idx = pdfCourses.findIndex((p, i) => !matchedPdfIndices.has(i) && p['課程代碼'] && p['課程代碼'].toLowerCase() === wCourse['課程代碼'].toLowerCase());
            if (idx !== -1) {
                matchedPdf = pdfCourses[idx];
                matchedIdx = idx;
            }
        }

        // 2. Fallback Match by Course Title similarity
        if (!matchedPdf && wCourse['課程名稱']) {
            const wClean = normalizeText(wCourse['課程名稱']);
            for (let i = 0; i < pdfCourses.length; i++) {
                if (matchedPdfIndices.has(i)) continue;
                const pClean = normalizeText(pdfCourses[i]['課程名稱']);
                if (wClean.includes(pClean) || pClean.includes(wClean) || calculateSimilarity(wClean, pClean) > 0.6) {
                    matchedPdf = pdfCourses[i];
                    matchedIdx = i;
                    break;
                }
            }
        }

        if (matchedPdf) {
            matchedPdfIndices.add(matchedIdx);
            results.push(compareSinglePair(wCourse, matchedPdf));
        } else {
            // Missing in PDF
            results.push({
                status: 'gray',
                statusText: 'PDF 排版漏排此課程',
                code: wCourse['課程代碼'] || '未標註',
                name: wCourse['課程名稱'],
                wordCourse: wCourse,
                pdfCourse: null,
                fields: {
                    '課程名稱': { status: 'gray', label: '課程名稱', word: wCourse['課程名稱'], pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    '時數': { status: 'gray', label: '時數', word: `${wCourse['時數']} 小時`, pdf: '-', desc: '缺少' },
                    '點數': { status: 'gray', label: '點數', word: `${wCourse['點數']} 點`, pdf: '-', desc: '缺少' },
                    '費用': { status: 'gray', label: '費用', word: `${Number(wCourse['費用'] || 0).toLocaleString()} 元`, pdf: '-', desc: '缺少' },
                    '教材': { status: 'gray', label: '教材', word: wCourse['教材'] || '-', pdf: '-', desc: '缺少' },
                    '課程內容': { status: 'gray', label: '課程內容', word: `${wCourse['課程內容'].length} 項內容`, pdf: '-', desc: '缺少' },
                    '備註事項': { status: 'gray', label: '備註事項', word: wCourse['備註事項'] || '-', pdf: '-', desc: '缺少' },
                    '後續推薦課程': { status: 'gray', label: '後續推薦課程', word: wCourse['後續推薦課程'] || '-', pdf: '-', desc: '缺少' }
                }
            });
        }
    }

    // Check for extra courses in PDF not in Word
    for (let i = 0; i < pdfCourses.length; i++) {
        if (!matchedPdfIndices.has(i)) {
            const pCourse = pdfCourses[i];
            results.push({
                status: 'gray',
                statusText: 'Word 原稿無此課程 (PDF 多出)',
                code: pCourse['課程代碼'] || '未標註',
                name: pCourse['課程名稱'],
                wordCourse: null,
                pdfCourse: pCourse,
                fields: {
                    '課程名稱': { status: 'gray', label: '課程名稱', word: '(未找到)', pdf: pCourse['課程名稱'], desc: 'Word 原稿未列出此課' },
                    '時數': { status: 'gray', label: '時數', word: '-', pdf: `${pCourse['時數']} 小時`, desc: '原稿無' },
                    '點數': { status: 'gray', label: '點數', word: '-', pdf: `${pCourse['點數']} 點`, desc: '原稿無' },
                    '費用': { status: 'gray', label: '費用', word: '-', pdf: `${Number(pCourse['費用'] || 0).toLocaleString()} 元`, desc: '原稿無' },
                    '教材': { status: 'gray', label: '教材', word: '-', pdf: pCourse['教材'] || '-', desc: '原稿無' },
                    '課程內容': { status: 'gray', label: '課程內容', word: '-', pdf: `${pCourse['課程內容'].length} 項內容`, desc: '原稿無' },
                    '備註事項': { status: 'gray', label: '備註事項', word: '-', pdf: pCourse['備註事項'] || '-', desc: '原稿無' },
                    '後續推薦課程': { status: 'gray', label: '後續推薦課程', word: '-', pdf: pCourse['後續推薦課程'] || '-', desc: '原稿無' }
                }
            });
        }
    }

    return results;
}

function compareSinglePair(w, p) {
    const fields = {};
    let hasRed = false;
    let hasYellow = false;

    // 1. 課程代碼
    const codeMatch = (w['課程代碼'] || '').trim().toLowerCase() === (p['課程代碼'] || '').trim().toLowerCase();
    fields['課程代碼'] = {
        label: '課程代碼',
        word: w['課程代碼'] || '-',
        pdf: p['課程代碼'] || '-',
        status: codeMatch ? 'green' : 'red',
        desc: codeMatch ? '代碼正確' : '課程代碼不相符！'
    };
    if (!codeMatch) hasRed = true;

    // 2. 課程名稱 (中文名稱)
    const wNameNorm = normalizeText(w['課程名稱']);
    const pNameNorm = normalizeText(p['課程名稱']);
    if (wNameNorm === pNameNorm) {
        fields['課程名稱'] = { label: '課程名稱', word: w['課程名稱'], pdf: p['課程名稱'], status: 'green', desc: '名稱完全相符' };
    } else if (calculateSimilarity(wNameNorm, pNameNorm) > 0.8) {
        fields['課程名稱'] = { label: '課程名稱', word: w['課程名稱'], pdf: p['課程名稱'], status: 'yellow', desc: '微小文字或空白差異' };
        hasYellow = true;
    } else {
        fields['課程名稱'] = { label: '課程名稱', word: w['課程名稱'], pdf: p['課程名稱'], status: 'red', desc: '名稱明顯不同或有錯字！' };
        hasRed = true;
    }

    // 3. 時數 (Rule 4: Mismatch is RED)
    const wHours = parseFloat(w['時數']) || 0;
    const pHours = parseFloat(p['時數']) || 0;
    if (wHours > 0 && pHours > 0 && wHours === pHours) {
        fields['時數'] = { label: '時數', word: `${w['時數']} 小時`, pdf: `${p['時數']} 小時`, status: 'green', desc: '時數相符' };
    } else {
        fields['時數'] = {
            label: '時數',
            word: `${w['時數'] || 0} 小時`,
            pdf: `${p['時數'] || 0} 小時`,
            status: 'red',
            desc: `時數不一致！Word 寫 ${w['時數']} 小時，但 PDF 寫 ${p['時數']} 小時`
        };
        hasRed = true;
    }

    // 4. 點數 (Rule 4: Mismatch is RED)
    const wPoints = parseFloat(w['點數']) || 0;
    const pPoints = parseFloat(p['點數']) || 0;
    if (wPoints > 0 && pPoints > 0 && wPoints === pPoints) {
        fields['點數'] = { label: '點數', word: `${w['點數']} 點`, pdf: `${p['點數']} 點`, status: 'green', desc: '點數相符' };
    } else if (w['點數'] || p['點數']) {
        fields['點數'] = {
            label: '點數',
            word: `${w['點數'] || 0} 點`,
            pdf: `${p['點數'] || 0} 點`,
            status: 'red',
            desc: `點數不一致！Word 為 ${w['點數']} 點，但 PDF 為 ${p['點數']} 點`
        };
        hasRed = true;
    } else {
        fields['點數'] = { label: '點數', word: '無', pdf: '無', status: 'green', desc: '雙方皆無點數' };
    }

    // 5. 費用
    const wPrice = parseInt(w['費用'], 10) || 0;
    const pPrice = parseInt(p['費用'], 10) || 0;
    if (wPrice > 0 && pPrice > 0 && wPrice === pPrice) {
        fields['費用'] = { label: '費用', word: `${wPrice.toLocaleString()} 元`, pdf: `${pPrice.toLocaleString()} 元`, status: 'green', desc: '費用相符' };
    } else if (wPrice !== pPrice) {
        fields['費用'] = {
            label: '費用',
            word: `${wPrice.toLocaleString()} 元`,
            pdf: `${pPrice.toLocaleString()} 元`,
            status: 'red',
            desc: `費用不符！Word 為 ${wPrice.toLocaleString()} 元，PDF 為 ${pPrice.toLocaleString()} 元`
        };
        hasRed = true;
    } else {
        fields['費用'] = { label: '費用', word: '-', pdf: '-', status: 'green', desc: '無費用資訊' };
    }

    // 6. 教材
    const wMat = normalizeText(w['教材'] || '');
    const pMat = normalizeText(p['教材'] || '');
    if (wMat === pMat) {
        fields['教材'] = { label: '教材', word: w['教材'] || '-', pdf: p['教材'] || '-', status: 'green', desc: '教材相符' };
    } else if (wMat.includes(pMat) || pMat.includes(wMat)) {
        fields['教材'] = { label: '教材', word: w['教材'] || '-', pdf: p['教材'] || '-', status: 'yellow', desc: '教材文字有修訂或簡寫' };
        hasYellow = true;
    } else {
        fields['教材'] = { label: '教材', word: w['教材'] || '-', pdf: p['教材'] || '-', status: 'red', desc: '教材資料不一致' };
        hasRed = true;
    }

    // 7. 課程內容 (Rule 1 & 2: 保持嚴格原名『課程內容』，List[str] 隔離比對)
    const contentDiff = compareCourseContent(w['課程內容'], p['課程內容']);
    fields['課程內容'] = {
        label: '課程內容',
        word: `${w['課程內容'].length} 個項目`,
        pdf: `${p['課程內容'].length} 個項目`,
        status: contentDiff.status,
        desc: contentDiff.desc,
        details: contentDiff.details
    };
    if (contentDiff.status === 'red') hasRed = true;
    if (contentDiff.status === 'yellow') hasYellow = true;

    // 8. 備註事項 (Rule 2 & 4: 獨立欄位隔離，不混入課程內容)
    const wNotes = normalizeText(w['備註事項'] || '');
    const pNotes = normalizeText(p['備註事項'] || '');
    if (!wNotes && !pNotes) {
        fields['備註事項'] = { label: '備註事項', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無備註' };
    } else if (wNotes === pNotes) {
        fields['備註事項'] = { label: '備註事項', word: w['備註事項'], pdf: p['備註事項'], status: 'green', desc: '備註相符' };
    } else if (calculateSimilarity(wNotes, pNotes) > 0.7) {
        fields['備註事項'] = { label: '備註事項', word: w['備註事項'], pdf: p['備註事項'], status: 'yellow', desc: '備註文字微調' };
        hasYellow = true;
    } else {
        fields['備註事項'] = { label: '備註事項', word: w['備註事項'] || '(無)', pdf: p['備註事項'] || '(漏排)', status: 'red', desc: '備註事項不一致或漏排！' };
        hasRed = true;
    }

    // 9. 後續推薦課程 (Rule 2 & 4: 獨立欄位隔離，不混入課程內容)
    const wRec = normalizeText(w['後續推薦課程'] || '');
    const pRec = normalizeText(p['後續推薦課程'] || '');
    if (!wRec && !pRec) {
        fields['後續推薦課程'] = { label: '後續推薦課程', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無推薦' };
    } else if (wRec === pRec) {
        fields['後續推薦課程'] = { label: '後續推薦課程', word: w['後續推薦課程'], pdf: p['後續推薦課程'], status: 'green', desc: '推薦課程相符' };
    } else if (calculateSimilarity(wRec, pRec) > 0.7) {
        fields['後續推薦課程'] = { label: '後續推薦課程', word: w['後續推薦課程'], pdf: p['後續推薦課程'], status: 'yellow', desc: '推薦課程文字微差' };
        hasYellow = true;
    } else {
        fields['後續推薦課程'] = { label: '後續推薦課程', word: w['後續推薦課程'] || '(無)', pdf: p['後續推薦課程'] || '(漏排)', status: 'red', desc: '推薦課程不一致或漏排！' };
        hasRed = true;
    }

    // 10. 適合對象
    const wTarget = normalizeText(w['適合對象'] || '');
    const pTarget = normalizeText(p['適合對象'] || '');
    if (!wTarget && !pTarget) {
        fields['適合對象'] = { label: '適合對象', word: '(無)', pdf: '(無)', status: 'green', desc: '無資訊' };
    } else if (wTarget === pTarget) {
        fields['適合對象'] = { label: '適合對象', word: w['適合對象'], pdf: p['適合對象'], status: 'green', desc: '相符' };
    } else if (calculateSimilarity(wTarget, pTarget) > 0.7) {
        fields['適合對象'] = { label: '適合對象', word: w['適合對象'], pdf: p['適合對象'], status: 'yellow', desc: '對象描述微差' };
        hasYellow = true;
    } else {
        fields['適合對象'] = { label: '適合對象', word: w['適合對象'] || '(無)', pdf: p['適合對象'] || '(漏排)', status: 'red', desc: '適合對象不一致' };
        hasRed = true;
    }

    // 11. 預備知識
    const wPrereq = normalizeText(w['預備知識'] || '');
    const pPrereq = normalizeText(p['預備知識'] || '');
    if (!wPrereq && !pPrereq) {
        fields['預備知識'] = { label: '預備知識', word: '(無)', pdf: '(無)', status: 'green', desc: '無預備知識' };
    } else if (wPrereq === pPrereq) {
        fields['預備知識'] = { label: '預備知識', word: w['預備知識'], pdf: p['預備知識'], status: 'green', desc: '相符' };
    } else if (calculateSimilarity(wPrereq, pPrereq) > 0.7) {
        fields['預備知識'] = { label: '預備知識', word: w['預備知識'], pdf: p['預備知識'], status: 'yellow', desc: '文字微差' };
        hasYellow = true;
    } else {
        fields['預備知識'] = { label: '預備知識', word: w['預備知識'] || '(無)', pdf: p['預備知識'] || '(漏排)', status: 'red', desc: '預備知識不一致' };
        hasRed = true;
    }

    // 12. 先修課程 (if present in Word)
    if (w['先修課程'] || p['先修課程']) {
        const wPre = normalizeText(w['先修課程'] || '');
        const pPre = normalizeText(p['先修課程'] || '');
        if (wPre === pPre) {
            fields['先修課程'] = { label: '先修課程', word: w['先修課程'], pdf: p['先修課程'], status: 'green', desc: '相符' };
        } else if (calculateSimilarity(wPre, pPre) > 0.7) {
            fields['先修課程'] = { label: '先修課程', word: w['先修課程'], pdf: p['先修課程'], status: 'yellow', desc: '文字微調' };
            hasYellow = true;
        } else {
            fields['先修課程'] = { label: '先修課程', word: w['先修課程'] || '(無)', pdf: p['先修課程'] || '(漏排)', status: 'red', desc: '先修課程不符！' };
            hasRed = true;
        }
    }

    // 13. 課程目標 (Rule 3: 特殊校對容錯規則：PDF 缺失視為正常，顯示 ⚪/🟡 略過，嚴禁亮紅燈！)
    const wObj = normalizeText(w['課程目標'] || '');
    const pObj = normalizeText(p['課程目標'] || '');
    if (!pObj) {
        fields['課程目標'] = {
            label: '課程目標',
            word: w['課程目標'] ? `${w['課程目標'].slice(0, 35)}...` : '(無)',
            pdf: '(版面精簡未排版)',
            status: 'gray', // ⚪ 灰色略過
            desc: '版面精簡未排版 / 略過 (正常)'
        };
        // Rule 3: Do NOT set hasRed or hasYellow!
    } else {
        if (wObj === pObj) {
            fields['課程目標'] = { label: '課程目標', word: w['課程目標'], pdf: p['課程目標'], status: 'green', desc: '完全相符' };
        } else if (calculateSimilarity(wObj, pObj) > 0.7) {
            fields['課程目標'] = { label: '課程目標', word: w['課程目標'], pdf: p['課程目標'], status: 'yellow', desc: '文字微調' };
            hasYellow = true;
        } else {
            fields['課程目標'] = { label: '課程目標', word: w['課程目標'], pdf: p['課程目標'], status: 'red', desc: '課程目標不一致！' };
            hasRed = true;
        }
    }

    // 14. 學會技能 (Rule 3: 特殊校對容錯規則：PDF 缺失視為正常，顯示 ⚪/🟡 略過，嚴禁亮紅燈！)
    const wSkill = normalizeText(w['學會技能'] || '');
    const pSkill = normalizeText(p['學會技能'] || '');
    if (!pSkill) {
        fields['學會技能'] = {
            label: '學會技能',
            word: w['學會技能'] ? `${w['學會技能'].slice(0, 35)}...` : '(無)',
            pdf: '(版面精簡未排版)',
            status: 'gray', // ⚪ 灰色略過
            desc: '版面精簡未排版 / 略過 (正常)'
        };
        // Rule 3: Do NOT set hasRed or hasYellow!
    } else {
        if (wSkill === pSkill) {
            fields['學會技能'] = { label: '學會技能', word: w['學會技能'], pdf: p['學會技能'], status: 'green', desc: '完全相符' };
        } else if (calculateSimilarity(wSkill, pSkill) > 0.7) {
            fields['學會技能'] = { label: '學會技能', word: w['學會技能'], pdf: p['學會技能'], status: 'yellow', desc: '文字微調' };
            hasYellow = true;
        } else {
            fields['學會技能'] = { label: '學會技能', word: w['學會技能'], pdf: p['學會技能'], status: 'red', desc: '學會技能不一致！' };
            hasRed = true;
        }
    }

    // Overall Status
    let overallStatus = 'green';
    let overallText = '檢查通過，所有重要欄位正確';
    if (hasRed) {
        overallStatus = 'red';
        overallText = '發現重大錯誤（時數、點數、費用或文字不符）';
    } else if (hasYellow) {
        overallStatus = 'yellow';
        overallText = '部分欄位有格式或文字微差提醒';
    }

    return {
        status: overallStatus,
        statusText: overallText,
        code: w['課程代碼'] || p['課程代碼'] || '未標註',
        name: w['課程名稱'] || p['課程名稱'],
        wordCourse: w,
        pdfCourse: p,
        fields
    };
}

function compareCourseContent(wItems, pItems) {
    if (!wItems.length && !pItems.length) {
        return { status: 'green', desc: '雙方皆無內容', details: [] };
    }

    const details = [];
    let matchCount = 0;

    for (let i = 0; i < Math.max(wItems.length, pItems.length); i++) {
        const wItem = wItems[i] || null;
        const pItem = pItems[i] || null;

        if (wItem && pItem) {
            const wClean = normalizeText(wItem);
            const pClean = normalizeText(pItem);
            if (wClean === pClean) {
                matchCount++;
                details.push({ index: i + 1, status: 'green', word: wItem, pdf: pItem, desc: '相符' });
            } else if (calculateSimilarity(wClean, pClean) > 0.6) {
                matchCount++;
                details.push({ index: i + 1, status: 'yellow', word: wItem, pdf: pItem, desc: '文字微差' });
            } else {
                details.push({ index: i + 1, status: 'red', word: wItem, pdf: pItem, desc: '項目內容不符' });
            }
        } else if (wItem && !pItem) {
            details.push({ index: i + 1, status: 'red', word: wItem, pdf: '(PDF 漏排)', desc: 'PDF 缺少此項目' });
        } else if (!wItem && pItem) {
            details.push({ index: i + 1, status: 'yellow', word: '(Word 無此項)', pdf: pItem, desc: 'PDF 多排此項目' });
        }
    }

    if (wItems.length === pItems.length && matchCount === wItems.length) {
        return { status: 'green', desc: `課程內容完全相符 (${wItems.length} 項)`, details };
    } else if (Math.abs(wItems.length - pItems.length) <= 1 && matchCount >= Math.min(wItems.length, pItems.length) * 0.8) {
        return { status: 'yellow', desc: `課程內容有微調 (Word: ${wItems.length}項, PDF: ${pItems.length}項)`, details };
    } else {
        return { status: 'red', desc: `課程內容項目數量或文字有缺漏 (Word: ${wItems.length}項, PDF: ${pItems.length}項)`, details };
    }
}

// ==========================================
// UI Rendering & Dashboard
// ==========================================

function updateKPIDashboard() {
    const list = state.comparisons;
    const total = list.length;
    const green = list.filter(c => c.status === 'green').length;
    const red = list.filter(c => c.status === 'red').length;
    const yellow = list.filter(c => c.status === 'yellow').length;
    const gray = list.filter(c => c.status === 'gray').length;

    elements.kpiTotal.textContent = total;
    elements.kpiMatch.textContent = green;
    elements.kpiError.textContent = red;
    elements.kpiWarning.textContent = yellow;
    elements.kpiMissing.textContent = gray;

    elements.countAll.textContent = total;
    elements.countGreen.textContent = green;
    elements.countRed.textContent = red;
    elements.countYellow.textContent = yellow;
    elements.countGray.textContent = gray;
}

function renderCourseCards() {
    const container = elements.courseCardsContainer;
    container.innerHTML = '';

    const filter = state.activeFilter;
    const query = state.searchQuery;

    const filtered = state.comparisons.filter(c => {
        if (filter !== 'all' && c.status !== filter) return false;
        if (query) {
            const codeStr = (c.code || '').toLowerCase();
            const nameStr = (c.name || '').toLowerCase();
            if (!codeStr.includes(query) && !nameStr.includes(query)) return false;
        }
        return true;
    });

    if (filtered.length === 0) {
        elements.emptyFilterView.classList.remove('hidden');
        return;
    } else {
        elements.emptyFilterView.classList.add('hidden');
    }

    filtered.forEach((item, idx) => {
        const card = createCourseCard(item, idx);
        container.appendChild(card);
    });
}

function createCourseCard(item, idx) {
    const card = document.createElement('div');
    card.className = `course-card bg-white rounded-2xl border shadow-xs transition overflow-hidden ${
        item.status === 'red' ? 'border-red-300 ring-1 ring-red-400/20' :
        item.status === 'yellow' ? 'border-amber-300' :
        item.status === 'gray' ? 'border-slate-300 bg-slate-50/50' : 'border-slate-200'
    }`;

    // Header Badge info
    let statusBadgeClass = '';
    let statusDotClass = '';
    let statusIcon = '';
    if (item.status === 'green') {
        statusBadgeClass = 'badge-green';
        statusDotClass = 'green';
        statusIcon = `<svg class="w-4 h-4 text-emerald-600 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"></path></svg>`;
    } else if (item.status === 'red') {
        statusBadgeClass = 'badge-red';
        statusDotClass = 'red';
        statusIcon = `<svg class="w-4 h-4 text-red-600 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"></path></svg>`;
    } else if (item.status === 'yellow') {
        statusBadgeClass = 'badge-yellow';
        statusDotClass = 'yellow';
        statusIcon = `<svg class="w-4 h-4 text-amber-600 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"></path></svg>`;
    } else {
        statusBadgeClass = 'badge-gray';
        statusDotClass = 'gray';
        statusIcon = `<svg class="w-4 h-4 text-slate-500 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-8-3a1 1 0 00-.867.5 1 1 0 11-1.731-1A3 3 0 0113 8a3.001 3.001 0 01-2 2.83V11a1 1 0 11-2 0v-1a1 1 0 011-1 1 1 0 100-2zm0 8a1 1 0 100-2 1 1 0 000 2z" clip-rule="evenodd"></path></svg>`;
    }

    const pdfPage = item.pdfCourse && item.pdfCourse.page ? `PDF 第 ${item.pdfCourse.page} 頁` : '';

    // Render all fields in exact order
    const fieldsOrder = [
        '時數',
        '點數',
        '課程名稱',
        '費用',
        '教材',
        '課程內容',
        '備註事項',
        '後續推薦課程',
        '適合對象',
        '預備知識',
        '先修課程',
        '課程目標',
        '學會技能'
    ];

    let rowsHtml = '';

    for (const key of fieldsOrder) {
        const field = item.fields[key];
        if (!field) continue;

        let rowBg = '';
        let badgeStyle = '';
        let dotStyle = '';
        if (field.status === 'red') {
            rowBg = 'bg-red-50/50';
            badgeStyle = 'badge-red';
            dotStyle = 'red';
        } else if (field.status === 'yellow') {
            rowBg = 'bg-amber-50/30';
            badgeStyle = 'badge-yellow';
            dotStyle = 'yellow';
        } else if (field.status === 'green') {
            badgeStyle = 'badge-green';
            dotStyle = 'green';
        } else {
            badgeStyle = 'badge-gray';
            dotStyle = 'gray';
        }

        // 課程內容 accordion button
        let contentExpandBtn = '';
        if (key === '課程內容' && field.details && field.details.length > 0) {
            contentExpandBtn = `
                <button onclick="toggleContentDetails(${idx})" class="mt-1.5 text-xs text-blue-600 hover:text-blue-800 font-semibold inline-flex items-center cursor-pointer transition">
                    <span>展開課程內容逐條對照</span>
                    <svg class="w-3.5 h-3.5 ml-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
                </button>
            `;
        }

        rowsHtml += `
            <tr class="border-b border-slate-100 last:border-none ${rowBg}">
                <td class="py-3 px-4 text-xs font-semibold text-slate-700 whitespace-nowrap align-top">
                    ${field.label}
                </td>
                <td class="py-3 px-4 text-xs text-slate-800 font-mono align-top break-words">
                    ${formatFieldValue(key, field.word, field.status === 'red')}
                </td>
                <td class="py-3 px-4 text-xs text-slate-800 font-mono align-top break-words">
                    ${formatFieldValue(key, field.pdf, field.status === 'red')}
                </td>
                <td class="py-3 px-4 text-xs align-top">
                    <span class="inline-flex items-center px-2.5 py-1 rounded text-xs font-medium ${badgeStyle}">
                        <span class="light-dot ${dotStyle} mr-1.5"></span>
                        ${field.desc}
                    </span>
                    ${contentExpandBtn}
                </td>
            </tr>
        `;
    }

    // 課程內容 expandable accordion
    let contentAccordionHtml = '';
    if (item.fields['課程內容'] && item.fields['課程內容'].details && item.fields['課程內容'].details.length > 0) {
        let itemsHtml = '';
        item.fields['課程內容'].details.forEach(d => {
            const dColor = d.status === 'red' ? 'text-red-600 font-semibold' : d.status === 'yellow' ? 'text-amber-700' : 'text-slate-700';
            itemsHtml += `
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 py-2 border-b border-slate-100 text-xs last:border-none">
                    <div class="bg-white p-2.5 rounded border border-slate-200">
                        <span class="text-xs text-slate-400 font-mono mr-1">#${d.index} Word:</span>
                        <span class="text-slate-800">${d.word || '(無)'}</span>
                    </div>
                    <div class="bg-white p-2.5 rounded border ${d.status === 'red' ? 'border-red-300 bg-red-50/30' : 'border-slate-200'}">
                        <span class="text-xs text-slate-400 font-mono mr-1">#${d.index} PDF:</span>
                        <span class="${dColor}">${d.pdf || '(漏排)'}</span>
                        ${d.status === 'red' ? `<span class="ml-2 text-2xs px-1.5 py-0.5 bg-red-100 text-red-700 rounded font-bold">項目內容不符或遺漏</span>` : ''}
                    </div>
                </div>
            `;
        });

        contentAccordionHtml = `
            <div id="contentDetails-${idx}" class="hidden p-4 bg-slate-50 border-t border-slate-200">
                <h5 class="text-xs font-bold text-slate-700 mb-2 flex items-center">
                    <svg class="w-4 h-4 mr-1 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path></svg>
                    課程內容逐條對照清單 (Word vs PDF)
                </h5>
                <div class="space-y-1">
                    ${itemsHtml}
                </div>
            </div>
        `;
    }

    card.innerHTML = `
        <div class="p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100">
            <div class="flex items-center space-x-3">
                <span class="px-2.5 py-1 rounded-lg text-xs font-bold font-mono tracking-wide ${
                    item.status === 'red' ? 'bg-red-100 text-red-800' : 'bg-blue-100 text-blue-800'
                }">
                    ${item.code}
                </span>
                <div>
                    <h3 class="text-base font-bold text-slate-900">${item.name}</h3>
                    <p class="text-xs text-slate-400">${item.wordCourse && item.wordCourse['英文名稱'] ? item.wordCourse['英文名稱'] : (item.pdfCourse && item.pdfCourse['英文名稱'] ? item.pdfCourse['英文名稱'] : '')}</p>
                </div>
            </div>

            <div class="flex items-center space-x-3">
                ${pdfPage ? `<span class="text-xs text-slate-400 font-medium">${pdfPage}</span>` : ''}
                <span class="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold ${statusBadgeClass}">
                    ${statusIcon}
                    ${item.statusText}
                </span>
            </div>
        </div>

        <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse">
                <thead>
                    <tr class="bg-slate-50/70 border-b border-slate-100 text-3xs uppercase tracking-wider text-slate-400 font-semibold">
                        <th class="py-2.5 px-4 w-28">欄位名稱</th>
                        <th class="py-2.5 px-4 w-5/12">Word 原稿內容 (.docx)</th>
                        <th class="py-2.5 px-4 w-5/12">美編排版內容 (.pdf)</th>
                        <th class="py-2.5 px-4 w-44">校對結果</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>

        ${contentAccordionHtml}
    `;

    return card;
}

function formatFieldValue(field, val, isError) {
    if (!val) return '<span class="text-slate-300">-</span>';
    if (isError && (field === '時數' || field === '點數' || field === '費用')) {
        return `<span class="diff-val-error">${val}</span>`;
    }
    return val;
}

window.toggleContentDetails = function(idx) {
    const el = document.getElementById(`contentDetails-${idx}`);
    if (el) {
        el.classList.toggle('hidden');
    }
};

// ==========================================
// Reporting & Export
// ==========================================

function copyErrorReport() {
    const redItems = state.comparisons.filter(c => c.status === 'red');
    const grayItems = state.comparisons.filter(c => c.status === 'gray');

    if (redItems.length === 0 && grayItems.length === 0) {
        showToast('恭喜！未發現任何錯誤，無需複製修改清單 🎉');
        return;
    }

    let report = `【課程資料校稿差異報告 - 待美編修正】\n`;
    report += `比對時間：${new Date().toLocaleString('zh-TW')}\n`;
    report += `來源檔案：Word [${state.wordFile ? state.wordFile.name : 'Word'}] ⇄ PDF [${state.pdfFile ? state.pdfFile.name : 'PDF'}]\n`;
    report += `說明：依排版規則，『課程目標』與『學會技能』若因精簡版面未排入，視為正常略過。\n`;
    report += `--------------------------------------------------------\n\n`;

    if (redItems.length > 0) {
        report += `🔴 【資料不一致錯誤 (${redItems.length} 門課程)】：\n`;
        redItems.forEach((item, i) => {
            report += `\n${i + 1}. 【${item.code}】${item.name} (頁數：${item.pdfCourse && item.pdfCourse.page ? 'P.' + item.pdfCourse.page : '未知'})\n`;
            for (const key of Object.keys(item.fields)) {
                const f = item.fields[key];
                if (f.status === 'red') {
                    report += `   - ${f.label}：Word 原稿寫「${f.word}」，但 PDF 排版為「${f.pdf}」 ➔ ${f.desc}\n`;
                }
            }
        });
        report += `\n`;
    }

    if (grayItems.length > 0) {
        report += `⚪ 【遺漏或多排課程 (${grayItems.length} 門課程)】：\n`;
        grayItems.forEach((item, i) => {
            report += `${i + 1}. 【${item.code}】${item.name} ➔ ${item.statusText}\n`;
        });
        report += `\n`;
    }

    report += `--------------------------------------------------------\n`;
    report += `請美編團隊協助核對並修正上述紅字項目，謝謝！\n`;

    navigator.clipboard.writeText(report).then(() => {
        showToast('已成功複製校稿報告清單至剪貼簿！可直接傳送給美編');
    }).catch(() => {
        const textarea = document.createElement('textarea');
        textarea.value = report;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        showToast('已成功複製校稿報告清單至剪貼簿！');
    });
}

function exportCSVReport() {
    if (state.comparisons.length === 0) {
        showToast('尚未進行校稿比對，無資料可匯出！', true);
        return;
    }

    const headers = [
        '課程代碼',
        '課程名稱',
        '比對狀態',
        'Word時數',
        'PDF時數',
        'Word點數',
        'PDF點數',
        'Word費用',
        'PDF費用',
        '教材比對',
        '課程內容比對',
        '備註事項比對',
        '後續推薦課程比對',
        '課程目標比對',
        '學會技能比對',
        'PDF頁碼',
        '差異說明'
    ];
    const rows = [headers];

    for (const c of state.comparisons) {
        const fields = c.fields;
        const errDescs = [];
        for (const k of Object.keys(fields)) {
            if (fields[k].status === 'red') errDescs.push(`${fields[k].label}:${fields[k].desc}`);
        }

        rows.push([
            c.code,
            `"${(c.name || '').replace(/"/g, '""')}"`,
            c.status === 'green' ? '相符' : c.status === 'red' ? '錯誤' : c.status === 'yellow' ? '提醒' : '遺漏',
            fields['時數'] ? fields['時數'].word : '',
            fields['時數'] ? fields['時數'].pdf : '',
            fields['點數'] ? fields['點數'].word : '',
            fields['點數'] ? fields['點數'].pdf : '',
            fields['費用'] ? fields['費用'].word : '',
            fields['費用'] ? fields['費用'].pdf : '',
            fields['教材'] ? fields['教材'].desc : '',
            fields['課程內容'] ? fields['課程內容'].desc : '',
            fields['備註事項'] ? fields['備註事項'].desc : '',
            fields['後續推薦課程'] ? fields['後續推薦課程'].desc : '',
            fields['課程目標'] ? fields['課程目標'].desc : '',
            fields['學會技能'] ? fields['學會技能'].desc : '',
            c.pdfCourse && c.pdfCourse.page ? c.pdfCourse.page : '',
            `"${errDescs.join('; ').replace(/"/g, '""')}"`
        ]);
    }

    const csvContent = '\uFEFF' + rows.map(r => r.join(',')).join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `課程校稿報告_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
    showToast('已成功匯出 CSV 校稿報告！');
}

// ==========================================
// Utilities
// ==========================================

function normalizeText(str) {
    if (!str) return '';
    return str
        .replace(/[\s\r\n\t\u3000]+/g, '')
        .replace(/[，,。.:：;；()（）「」『』"'-]/g, '')
        .toLowerCase();
}

function calculateSimilarity(s1, s2) {
    if (!s1 || !s2) return 0;
    if (s1 === s2) return 1;
    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;
    if (longer.length === 0) return 1.0;

    let matches = 0;
    for (let i = 0; i < shorter.length; i++) {
        if (longer.includes(shorter[i])) matches++;
    }
    return matches / longer.length;
}

function showToast(message, isError = false) {
    elements.toastMessage.textContent = message;
    elements.toast.classList.remove('translate-y-20', 'opacity-0');
    if (isError) {
        elements.toast.classList.add('bg-red-900');
        elements.toast.classList.remove('bg-slate-900');
    } else {
        elements.toast.classList.remove('bg-red-900');
        elements.toast.classList.add('bg-slate-900');
    }
    setTimeout(() => {
        elements.toast.classList.add('translate-y-20', 'opacity-0');
    }, 3500);
}
