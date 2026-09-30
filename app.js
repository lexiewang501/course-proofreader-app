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
    '課程大綱',
    '課程介紹',
    '學會技能',
    '備註事項',
    '後續推薦課程'
];

// Fields strictly formatted as List[str] / Array<string>
const LIST_FIELDS = new Set([
    '課程內容',
    '備註事項',
    '適合對象',
    '預備知識',
    '先修課程'
]);

// Universal outline numbering & bullet patterns (Chinese & Arabic numbers, bullets, brackets, outline headers, Domains, learning codes)
const BULLET_ITEM_PATTERN = /^(?:[一二三四五六七八九十百]+[、.．]|[（(][一二三四五六七八九十百\d]+[)）]|[①-⑳❶-❿㈠-㈩]|\d+[、.．]|\d+\.(?!\d)|[【\[]\d+[】\]]|[•●※·◆▪＊★☆✦✧✓✔✗✘►▶▷▸]|[-*](?:\s+|$)|(?:Lesson|Module|Chapter|Unit|Section|Topic|Domain|主題|單元|章節|階段|步驟|目標|項目|Day|Step|Phase|Part)\s*[一二三四五六七八九十\d]+|[A-Za-z]\d+(?:\.\d+)+|[A-Za-z][.、)）])/i;

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
    btnSampleComptia: document.getElementById('btnSampleComptia'),

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
    if (elements.btnSampleComptia) {
        elements.btnSampleComptia.addEventListener('click', () => loadSample('comptia'));
    }
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
    let wordPath, pdfPath, labelName;
    if (type === 'blockchain') {
        wordPath = '恆逸_區塊鏈_2027年1-6月課程.docx';
        pdfPath = '區塊鏈.pdf';
        labelName = '區塊鏈';
    } else if (type === 'jianzhen') {
        wordPath = '恆逸_鑒真數位_2027年1-6月課程v1_1.docx';
        pdfPath = '鑒真數位.pdf';
        labelName = '鑒真數位';
    } else {
        wordPath = '恆逸_CompTIA_2027年1-6月課程.docx';
        pdfPath = 'CompTIA.pdf';
        labelName = 'CompTIA';
    }

    try {
        elements.loadingIndicator.classList.remove('hidden');
        elements.loadingText.textContent = `正在讀取本地測試檔案 (${labelName})...`;

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
 * Extracts CourseIdentity (course_code, course_name_zh, course_name_en) from Word table rows.
 * Handles parentheses, newlines, and dedicated English course title fields.
 */
function extractCourseIdentityFromWord(rows, course) {
    let rawTitle = '';
    let rawCode = '';
    let rawEnTitle = '';

    if (rows.length > 0) {
        const row0 = rows[0];
        if (row0.length >= 2 && row0[0].fullText.trim().length <= 15 && row0[0].fullText.trim().length > 0) {
            rawCode = row0[0].fullText.trim();
            rawTitle = row0.slice(1).map(c => c.fullText).join(' ').trim();
        } else {
            rawTitle = row0.map(c => c.fullText).join(' ').trim();
            // In tables where row 0 cell 0 is blank (e.g. NSPA, ANSPA, CNSPA),
            // the course code is placed in row 1 cell 0 alongside the English title in row 1 cell 1.
            if (rows.length > 1 && rows[1].length >= 2) {
                const r1c0 = rows[1][0].fullText.trim();
                const KNOWN_NON_CODE_LABELS = ['英文課名', '英文名稱', '原廠課名', '英文', '時數', '費用', '點數', '教材', '課程目標', '適合對象'];
                const isNonCode = KNOWN_NON_CODE_LABELS.some(l => r1c0.startsWith(l)) || 
                                  r1c0.toLowerCase().startsWith('course') || 
                                  r1c0.toLowerCase().startsWith('english');
                if (!isNonCode && /^[A-Za-z0-9_-]{2,15}$/.test(r1c0)) {
                    rawCode = r1c0;
                    const r1Rest = rows[1].slice(1).map(c => c.fullText).join(' ').trim();
                    if (r1Rest) rawEnTitle = r1Rest;
                }
            }
        }
    }

    // Check for dedicated English course title row in table
    for (let r = 1; r < rows.length; r++) {
        const rText = rows[r].map(c => c.fullText).join(' ').trim();
        const firstCell = rows[r][0] ? rows[r][0].fullText.trim() : '';

        if (rText.includes('時數') || rText.includes('費用') || rText.includes('點數')) {
            break;
        }

        if (firstCell.startsWith('英文課名') || firstCell.startsWith('英文名稱') || 
            firstCell.startsWith('原廠課名') || firstCell.startsWith('英文') ||
            firstCell.toLowerCase().startsWith('course name') || firstCell.toLowerCase().startsWith('english name')) {
            const val = rows[r].slice(1).map(c => c.fullText).join(' ').trim() || rText.replace(/^[^\s：:]+[：:\s]*/, '').trim();
            if (val) rawEnTitle = val;
            break;
        } else if (r === 1 && !rawEnTitle && !rText.includes('時數') && !rText.includes('費用')) {
            if (rawCode && rows[r].length >= 2 && rows[r][0].fullText.trim() === rawCode) {
                rawEnTitle = rows[r].slice(1).map(c => c.fullText).join(' ').trim();
            } else {
                rawEnTitle = rText.replace(/^[|：:\s]+/, '').trim();
            }
        }
    }

    // 1. Check for parentheses containing English name: (Azure Fundamentals) or （Blockchain Developer Course）
    const parenMatch = rawTitle.match(/[(（]([^)）\u4e00-\u9fa5]*[A-Za-z]{2,}[^)）\u4e00-\u9fa5]*)[)）]/);
    if (parenMatch) {
        if (!rawEnTitle) {
            rawEnTitle = parenMatch[1].trim();
        }
        rawTitle = rawTitle.replace(parenMatch[0], ' ').trim();
    }

    // 2. Check for newlines in rawTitle
    if (rawTitle.includes('\n')) {
        const lines = rawTitle.split(/\n+/).map(l => l.trim()).filter(Boolean);
        const zhLines = [];
        for (const line of lines) {
            if (/[\u4e00-\u9fa5]/.test(line)) {
                zhLines.push(line);
            } else if (!rawEnTitle && !/[\u4e00-\u9fa5]/.test(line) && /[A-Za-z]{2,}/.test(line)) {
                rawEnTitle = line;
            }
        }
        if (zhLines.length > 0) rawTitle = zhLines.join(' ');
    }

    // 3. Extract course code from rawTitle if not already found
    if (!rawCode) {
        const codeBracketMatch = rawTitle.match(/^[\[【]([A-Za-z0-9_-]{2,15})[\]】]/);
        if (codeBracketMatch) {
            rawCode = codeBracketMatch[1];
            rawTitle = rawTitle.replace(codeBracketMatch[0], '').trim();
        } else {
            const codePrefixMatch = rawTitle.match(/^([A-Za-z0-9_-]{2,15})[：:\s]+(.*)$/);
            if (codePrefixMatch && !['APP', 'DApp', 'Web3', 'EVM', 'Full', 'Stack', 'Course'].includes(codePrefixMatch[1])) {
                rawCode = codePrefixMatch[1];
                rawTitle = codePrefixMatch[2].trim();
            }
        }
    } else {
        rawTitle = rawTitle.replace(new RegExp(`^${rawCode}[：:\\s]*`), '').trim();
    }

    // 4. Fallback: extract course code from rawEnTitle if still not found
    if (!rawCode && rawEnTitle) {
        const enCodeMatch = rawEnTitle.match(/^([A-Za-z0-9_-]{2,15})[：:\s]+([A-Za-z].*)$/);
        if (enCodeMatch) {
            const potentialCode = enCodeMatch[1];
            const COMMON_WORDS = ['THE', 'AN', 'A', 'FOR', 'AND', 'OF', 'IN', 'ON', 'WITH', 'BY', 'TO', 'AT', 'FROM', 'ALL', 'NEW', 'ADVANCED', 'BASIC', 'COURSE', 'PRACTICAL', 'MASTER', 'CYBERSECURITY'];
            if (!COMMON_WORDS.includes(potentialCode.toUpperCase()) && /^[A-Z0-9_-]+$/.test(potentialCode)) {
                rawCode = potentialCode;
                rawEnTitle = enCodeMatch[2].trim();
            }
        }
    }

    // 5. Ensure rawEnTitle is stripped of redundant course code prefix if present
    if (rawCode && rawEnTitle) {
        rawEnTitle = rawEnTitle.replace(new RegExp(`^${rawCode}[：:\\s]+`, 'i'), '').trim();
    }

    course.course_code = rawCode.trim();
    course.course_name_zh = rawTitle.replace(/^[：:\s|]+/, '').replace(/\s+/g, ' ').trim();
    course.course_name_en = rawEnTitle.replace(/^[：:\s|]+/, '').replace(/\s+/g, ' ').trim();

    // Synchronize aliases
    course['課程代碼'] = course.course_code;
    course['課程名稱'] = course.course_name_zh;
    course['中文課名'] = course.course_name_zh;
    course['英文名稱'] = course.course_name_en;
    course['英文課名'] = course.course_name_en;
}

/**
 * Extracts CourseIdentity (course_code, course_name_zh, course_name_en) from PDF header items.
 * Strictly recognizes English subtitles positioned below Chinese titles, ensuring no omission.
 * Filters out breadcrumbs / category tabs (e.g. "CompTIA SecAI+ 認證", "此課程為...證照") located above the title banner.
 */
function extractCourseIdentityFromPdf(headerItems, metaY, courseTop) {
    // 1. In promotional flyers, the course title banner is located below courseTop and within metaY + 80 headroom
    // for multi-line titles (e.g. 2-line Chinese titles like SEMGEI).
    // Out-of-table elements such as side tab markers (x < 35) are strictly excluded.
    const maxHeaderY = Math.min(courseTop, metaY + 80);
    let validItems = headerItems.filter(it => (it.x < 95 ? (it.x >= 35 && it.y >= metaY - 5) : it.y > metaY + 3) && it.y <= maxHeaderY);

    // Merge superscript symbols (® / ™ / ©) into the preceding word/acronym
    const supers = validItems.filter(it => /^[®™©]$/.test(it.str));
    if (supers.length > 0) {
        const remainingItems = [];
        for (const it of validItems) {
            if (/^[®™©]$/.test(it.str)) continue;
            const sup = supers.find(s => s.x >= it.x && (s.x - it.x) <= 45 && s.y >= it.y && (s.y - it.y) <= 15);
            if (sup) {
                remainingItems.push({
                    ...it,
                    str: it.str + sup.str
                });
            } else {
                remainingItems.push(it);
            }
        }
        validItems = remainingItems;
    }

    let code = '';
    const zhLines = [];
    const enLines = [];

    validItems.sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    let curY = null;
    let curLine = [];
    for (const it of validItems) {
        if (curY === null || Math.abs(curY - it.y) > 4) {
            if (curLine.length) lines.push(curLine);
            curY = it.y;
            curLine = [it];
        } else {
            curLine.push(it);
        }
    }
    if (curLine.length) lines.push(curLine);

    const candidateLines = [];
    for (const line of lines) {
        const lineMaxH = Math.max(...line.map(it => it.h || 0));

        // 1. Identify left-side course code badge inside table bounds (35 <= x < 95)
        // Strictly exclude external margin tabs (e.g. EPI, PMI, CompTIA), 4-digit years (2026/2027),
        // accreditation badges (PDU, CPE, OCP, ACP), and small font badges (h < 9)
        const leftBadge = line.find(it => it.x >= 35 && it.x < 95 && (it.h || 0) >= 9 && /^[A-Za-z0-9_-]{2,15}$/.test(it.str) && !/^[12]\d{3}$/.test(it.str) && !['EPI', 'PMI', 'CompTIA', 'PDU', 'CPE', 'OCP', 'ACP'].includes(it.str));
        if (leftBadge) {
            if (!code) {
                code = leftBadge.str;
            } else if (code.endsWith('-')) {
                code = code + leftBadge.str;
            }
        }

        const contentItems = line.filter(it => it !== leftBadge);
        if (contentItems.length === 0) continue;

        const lineText = contentItems.map(it => it.str).join(' ').trim();
        if (lineText.includes('課程簡介') || 
            lineText.includes('各地開課時間') || 
            lineText.includes('後續推薦課程') ||
            lineText.match(/^\|\s*.*\s*\|$/) || 
            lineText === '區塊鏈' || 
            lineText === '鑒真數位' ||
            lineText === 'EPI' ||
            lineText === 'PMI' ||
            lineText.includes('認可之資通安全專業證照') ||
            // Promotional & accreditation badges with small font (lineMaxH <= 10)
            (lineMaxH <= 10 && (
                lineText.includes('認證') || 
                lineText.includes('證照') ||
                lineText.includes('學分') ||
                lineText.includes('積分') ||
                lineText.includes('新課') ||
                lineText.includes('獨家') ||
                /^[12]\d{3}$/.test(lineText)
            ))
        ) {
            continue;
        }

        const standaloneCodeMatch = lineText.match(/^[A-Za-z0-9_-]{2,10}$/);
        if (!code && standaloneCodeMatch && !['APP', 'DApp', 'Web3', 'EVM', 'Full', 'Stack', 'Course', 'EPI', 'PMI', 'CompTIA'].includes(standaloneCodeMatch[0])) {
            code = standaloneCodeMatch[0];
            continue;
        }

        candidateLines.push({ line, contentItems, lineText, lineMaxH, y: line[0].y });
    }

    // Identify Chinese title lines by font hierarchy (title has largest font in banner)
    const zhCandidates = candidateLines.filter(l => /[\u4e00-\u9fa5]/.test(l.lineText));
    if (zhCandidates.length > 0) {
        const maxZhH = Math.max(...zhCandidates.map(l => l.lineMaxH));
        for (const l of zhCandidates) {
            // Only include lines matching the dominant title font size (filter out small font tags)
            if (maxZhH >= 12 && l.lineMaxH < maxZhH - 3) {
                continue;
            }
            zhLines.push(l.lineText);
        }
    }

    // English subtitle lines: any line in header banner without Chinese characters that contains Latin letters
    for (const l of candidateLines) {
        if (!/[\u4e00-\u9fa5]/.test(l.lineText) && /[A-Za-z]{2,}/.test(l.lineText) && l.lineText !== code) {
            enLines.push(l.lineText);
        }
    }

    let zhTitle = zhLines.join(' ').replace(/\s+/g, ' ').trim();
    if (code) {
        zhTitle = zhTitle.replace(new RegExp(`^${code}[：:\\s]*`), '').trim();
    } else {
        const prefixMatch = zhTitle.match(/^([A-Za-z0-9_-]{2,10})[：:\s]+(.*)$/);
        if (prefixMatch && !['APP', 'DApp', 'Web3', 'EVM', 'Full', 'Stack', 'Course'].includes(prefixMatch[1])) {
            code = prefixMatch[1];
            zhTitle = prefixMatch[2].trim();
        }
    }

    let enTitle = enLines.join(' ').replace(/\s+/g, ' ').trim();
    if (code) {
        enTitle = enTitle.replace(new RegExp(`^${code}(?:[：:\\s\\-]+|$)`, 'i'), '').trim();
    }

    if (!enTitle) {
        const parenMatch = zhTitle.match(/[(（]([^)）\u4e00-\u9fa5]*[A-Za-z]{2,}[^)）\u4e00-\u9fa5]*)[)）]/);
        if (parenMatch) {
            enTitle = parenMatch[1].trim();
            zhTitle = zhTitle.replace(parenMatch[0], ' ').trim();
        }
    }

    return {
        course_code: code,
        course_name_zh: zhTitle,
        course_name_en: enTitle
    };
}

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

        // Exact schema mapping with CourseIdentity
        const course = {
            course_code: '',     // 課程代碼/認證代號 (如 CCNA, AZ-900)
            course_name_zh: '',  // 中文課名
            course_name_en: '',  // 英文課名（原廠課名/官方名稱）

            // Aliases for backwards compatibility and strict field labeling
            '課程代碼': '',
            '課程名稱': '',
            '中文課名': '',
            '英文名稱': '',
            '英文課名': '',

            '時數': '',
            '費用': '',
            '點數': '',
            '教材': '',
            '課程目標': '',
            '適合對象': [], // List[str]
            '先修課程': [], // List[str]
            '預備知識': [], // List[str]
            '課程內容': [], // List[str]
            '學會技能': '',
            '備註事項': [], // List[str]
            '後續推薦課程': [], // List[str]
            source: 'Word'
        };

        // Extract CourseIdentity (course_code, course_name_zh, course_name_en)
        extractCourseIdentityFromWord(rows, course);

        let lastSectionLabel = null;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowFullText = row.map(c => c.fullText).join(' ').replace(/\s+/g, ' ').trim();
            const firstCellText = row[0] ? row[0].fullText.trim() : '';

            // Metadata row: only parse if not a known label row (防止備註事項等區塊誤認)
            const isLabelRow = KNOWN_LABELS.some(l => firstCellText === l || firstCellText.startsWith(l));
            if (!isLabelRow && (rowFullText.includes('時數') || rowFullText.includes('費用') || rowFullText.includes('點數'))) {
                const hoursMatch = rowFullText.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
                if (hoursMatch && !course['時數']) course['時數'] = hoursMatch[1];

                const priceMatch = rowFullText.match(/費用[：:\s]*([0-9,]+)\s*元?/);
                if (priceMatch && !course['費用']) course['費用'] = priceMatch[1].replace(/,/g, '');

                const pointsMatch = rowFullText.match(/點數[：:\s]*([0-9.]+)\s*點?/);
                if (pointsMatch && !course['點數']) course['點數'] = pointsMatch[1];

                const matCell = row.find(c => /(?:^|[\s|｜])教材[：:\s]/.test(c.fullText));
                if (matCell && !course['教材']) {
                    const m = matCell.fullText.match(/(?:^|[\s|｜])教材[：:\s]+([^|｜\n]+)/);
                    if (m) course['教材'] = m[1].trim();
                } else if (!course['教材']) {
                    const matMatch = rowFullText.match(/(?:^|[\s|｜])教材[：:\s]+([^|｜\n]+)/);
                    if (matMatch) course['教材'] = matMatch[1].trim();
                }
                lastSectionLabel = null;
                continue;
            }

            // Check known section labels or continuation rows
            let matchedLabel = KNOWN_LABELS.find(l => firstCellText === l || firstCellText.startsWith(l));
            if (!matchedLabel && rowFullText) {
                matchedLabel = KNOWN_LABELS.find(l => rowFullText.startsWith(l));
            }

            // Continuation row: firstCell is empty (e.g. vertically merged cell <w:vMerge> or unlabelled follow-up row like CSSLP 課程內容)
            if (!matchedLabel && !firstCellText && lastSectionLabel) {
                matchedLabel = lastSectionLabel;
            }

            if (matchedLabel) {
                if (matchedLabel === '課程大綱') matchedLabel = '課程內容';
                if (matchedLabel === '課程介紹') matchedLabel = '課程目標';
                lastSectionLabel = matchedLabel;
                const cellParas = row.slice(1).flatMap(c => c.paragraphs).map(p => p.trim()).filter(Boolean);
                const cellFull = row.slice(1).map(c => c.fullText).join(' ').trim();

                if (matchedLabel === '適合對象' || matchedLabel === '預備知識' || matchedLabel === '先修課程') {
                    if (hasBulletMarkers(cellParas) || hasBulletMarkers(cellFull) || cellParas.length > 1) {
                        const newItems = extractListItems(cellParas, cellFull);
                        course[matchedLabel] = Array.isArray(course[matchedLabel]) ? [...course[matchedLabel], ...newItems] : newItems;
                    } else {
                        const val = cellParas.length === 1 ? cellParas[0] : cellFull;
                        if (!course[matchedLabel] || (Array.isArray(course[matchedLabel]) && course[matchedLabel].length === 0)) {
                            course[matchedLabel] = val;
                        } else if (Array.isArray(course[matchedLabel])) {
                            course[matchedLabel].push(val);
                        } else {
                            course[matchedLabel] = [course[matchedLabel], val];
                        }
                    }
                } else if (matchedLabel === '課程內容' || matchedLabel === '備註事項') {
                    const newItems = extractListItems(cellParas, cellFull);
                    course[matchedLabel] = Array.isArray(course[matchedLabel]) ? [...course[matchedLabel], ...newItems] : newItems;
                } else if (matchedLabel === '後續推薦課程') {
                    const allParas = row.slice(1).flatMap(c => c.paragraphs).map(p => p.trim()).filter(Boolean);
                    const rawItems = allParas.length > 0 ? allParas : (cellFull ? [cellFull] : []);
                    const newItems = rawItems.flatMap(p => splitRecommendedCourses(p));
                    course['後續推薦課程'] = Array.isArray(course['後續推薦課程']) ? [...course['後續推薦課程'], ...newItems] : newItems;
                } else if (matchedLabel === '課程目標') {
                    const val = cellParas.length > 0 ? cellParas.join('\n') : cellFull;
                    course['課程目標'] = course['課程目標'] ? course['課程目標'] + '\n' + val : val;
                } else {
                    const val = row.slice(1).map(c => c.fullText).join(' ').trim();
                    course[matchedLabel] = course[matchedLabel] ? course[matchedLabel] + '\n' + val : val;
                }
            } else if (firstCellText) {
                lastSectionLabel = null;
            }
        }

        // If course identity exists, push to list
        if (course.course_name_zh || course.course_code || course.course_name_en) {
            courses.push(course);
        }
    }

    // Post-pass: split any concatenated recommended courses using catalog titles
    const docxTitles = courses.flatMap(c => [c.course_name_zh, c.course_name_en]).filter(t => t && t.length >= 4 && !/^[A-Za-z0-9/_-]+$/.test(t));
    for (const course of courses) {
        if (course['後續推薦課程']) {
            const rawRec = Array.isArray(course['後續推薦課程']) ? course['後續推薦課程'] : [course['後續推薦課程']];
            course['後續推薦課程'] = rawRec.flatMap(t => splitRecommendedCourses(t, docxTitles));
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
            w: Math.round(it.width * 10) / 10,
            h: Math.round((it.height || 0) * 10) / 10,
            str: it.str.trim()
        })).filter(it => it.str.length > 0 && it.x >= 35 && it.x < 565);

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

        // 2. Find metadata anchor lines: A true course metadata row has "時數"
        // AND contains "費用", "點數", or "教材", strictly excluding notes like "完整練習時數："
        const metaList = [];
        const yMapMeta = {};
        for (const it of items) {
            const key = Math.round(it.y / 4) * 4;
            if (!yMapMeta[key]) yMapMeta[key] = [];
            yMapMeta[key].push(it);
        }
        for (const [keyY, lineItems] of Object.entries(yMapMeta)) {
            const lineText = lineItems.map(i => i.str).join(' ');
            const hasHours = /(?:^|\s|\|)時數[：:\s]*\d+/.test(lineText) || lineText.startsWith('時數');
            const hasMetaContext = lineText.includes('費用') || lineText.includes('點數') || lineText.includes('教材');
            if (hasHours && hasMetaContext) {
                const avgY = lineItems.reduce((s, i) => s + i.y, 0) / lineItems.length;
                if (!metaList.some(m => Math.abs(m.y - avgY) <= 8)) {
                    metaList.push({ y: Math.round(avgY) });
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

            // Determine strict contiguous non-overlapping intervals for each section to prevent boundary bleed
            for (let s = 0; s < courseSections.length; s++) {
                const cur = courseSections[s];
                const prev = s > 0 ? courseSections[s - 1] : null;
                const next = s + 1 < courseSections.length ? courseSections[s + 1] : null;

                if (prev) {
                    cur.effectiveTop = (prev.bottom + cur.top) / 2;
                } else {
                    cur.effectiveTop = cur.top + 2;
                }

                if (next) {
                    cur.effectiveBottom = (cur.bottom + next.top) / 2;
                } else {
                    cur.effectiveBottom = cur.bottom - 2;
                }
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
                course_code: '',     // 課程代碼/認證代號 (如 CCNA, AZ-900)
                course_name_zh: '',  // 中文課名
                course_name_en: '',  // 英文課名（原廠課名/官方名稱）

                // Aliases for backwards compatibility and strict field labeling
                '課程代碼': '',
                '課程名稱': '',
                '中文課名': '',
                '英文名稱': '',
                '英文課名': '',

                '時數': '',
                '費用': '',
                '點數': '',
                '教材': '',
                '課程目標': '',
                '適合對象': [], // List[str]
                '先修課程': [], // List[str]
                '預備知識': [], // List[str]
                '課程內容': [], // List[str]
                '學會技能': '',
                '備註事項': [], // List[str]
                '後續推薦課程': [], // List[str]
                source: 'PDF',
                page: p
            };

            // Metadata Line items (strictly on the metadata row baseline, within content card bounds)
            const metaLineItems = items.filter(it => Math.abs(it.y - metaY) <= 3 && it.x >= 80 && it.x <= 550);
            metaLineItems.sort((a, b) => a.x - b.x);
            const metaLineText = metaLineItems.map(it => it.str).join(' ');

            const hoursMatch = metaLineText.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
            if (hoursMatch) course['時數'] = hoursMatch[1];

            const priceMatch = metaLineText.match(/費用[：:\s]*([0-9,]+)\s*元?/);
            if (priceMatch) course['費用'] = priceMatch[1].replace(/,/g, '');

            const pointsMatch = metaLineText.match(/點數[：:\s]*([0-9.]+)\s*點?/);
            if (pointsMatch) course['點數'] = pointsMatch[1];

            const matMatch = metaLineText.match(/教材[：:\s]*([^｜|\n]+)/);
            if (matMatch) {
                let matVal = matMatch[1].trim();
                let firstSectionTop = metaY - 30;
                if (courseSections.length > 0) {
                    firstSectionTop = courseSections[0].top;
                }
                const wrappedItems = items.filter(it => it.y < metaY - 3 && it.y >= firstSectionTop && it.x >= 280 && it.x <= 550);
                if (wrappedItems.length > 0) {
                    wrappedItems.sort((a, b) => b.y - a.y || a.x - b.x);
                    const extraMat = wrappedItems.map(it => it.str).join(' ').trim();
                    if (extraMat && (extraMat.startsWith('+') || /(?:教材|講義|書籍|電子書|環境|Lab|紙本|原廠)/.test(extraMat))) {
                        matVal = (matVal + (extraMat.startsWith('+') ? ' ' : ' ') + extraMat).trim();
                    }
                }
                course['教材'] = matVal;
            }

            // Extract CourseIdentity from PDF header items (嚴格擷取英文原廠副標題與代碼，嚴禁忽略英文副標，且限定於表格內 x >= 35 杜絕側邊頁籤)
            const headerItems = items.filter(it => (it.x < 95 ? (it.x >= 35 && it.y >= metaY - 5) : it.y > metaY + 3) && it.y <= courseTop);
            const identity = extractCourseIdentityFromPdf(headerItems, metaY, courseTop);

            course.course_code = identity.course_code;
            course.course_name_zh = identity.course_name_zh;
            course.course_name_en = identity.course_name_en;

            // Synchronize aliases
            course['課程代碼'] = course.course_code;
            course['課程名稱'] = course.course_name_zh;
            course['中文課名'] = course.course_name_zh;
            course['英文名稱'] = course.course_name_en;
            course['英文課名'] = course.course_name_en;

            // Extract each section strictly isolated inside its vertical partition (杜絕邊界文字滲透)
            for (const sec of courseSections) {
                const secItems = items.filter(it => it.x >= 88 && it.y > sec.effectiveBottom && it.y <= sec.effectiveTop);
                secItems.sort((a, b) => b.y - a.y || a.x - b.x);

                if (sec.label === '課程內容' || sec.label === '課程大綱') {
                    course['課程內容'] = extractListFromPdfSection(secItems, false);
                } else if (sec.label === '備註事項') {
                    course['備註事項'] = extractNotesFromPdfSection(secItems);
                } else if (sec.label === '後續推薦課程') {
                    const fullText = secItems.map(it => it.str).join(' ').replace(/\s+/g, ' ').trim();
                    course['後續推薦課程'] = splitRecommendedCourses(fullText);
                } else if (sec.label === '課程目標' || sec.label === '課程介紹') {
                    const lines = groupItemsIntoVisualLines(secItems);
                    const lineTexts = lines.map(l => l.text);
                    if (hasBulletMarkers(lineTexts) || getColumnSplits(secItems).length > 0) {
                        const assembled = extractListFromPdfSection(secItems, false);
                        course['課程目標'] = assembled.join('\n');
                    } else {
                        const assembled = assembleLinesIntoItems(lines, false);
                        course['課程目標'] = assembled.join('\n');
                    }
                } else if (sec.label === '適合對象' || sec.label === '預備知識' || sec.label === '先修課程') {
                    const lines = groupItemsIntoVisualLines(secItems);
                    const lineTexts = lines.map(l => l.text);
                    if (hasBulletMarkers(lineTexts) || getColumnSplits(secItems).length > 0) {
                        course[sec.label] = extractListFromPdfSection(secItems, false);
                    } else {
                        const assembled = assembleLinesIntoItems(lines, false);
                        const split = assembled.flatMap(it => splitInlineBullets(it));
                        course[sec.label] = split.length > 1 ? split : (split[0] || '');
                    }
                } else {
                    // Strictly isolate 學會技能, etc. (全欄位支援分欄與項目語意提取)
                    const lines = groupItemsIntoVisualLines(secItems);
                    const lineTexts = lines.map(l => l.text);
                    if (hasBulletMarkers(lineTexts) || getColumnSplits(secItems).length > 0) {
                        course[sec.label] = extractListFromPdfSection(secItems, false);
                    } else {
                        const assembled = assembleLinesIntoItems(lines, false);
                        const split = assembled.flatMap(it => splitInlineBullets(it));
                        course[sec.label] = split.length > 1 ? split : (split[0] || '');
                    }
                }
            }

            courses.push(course);
        }
    }

    // Post-pass: split any concatenated recommended courses using catalog titles
    const pdfTitles = courses.flatMap(c => [c.course_name_zh, c.course_name_en]).filter(t => t && t.length >= 4 && !/^[A-Za-z0-9/_-]+$/.test(t));
    for (const course of courses) {
        if (course['後續推薦課程']) {
            const rawRec = Array.isArray(course['後續推薦課程']) ? course['後續推薦課程'] : [course['後續推薦課程']];
            course['後續推薦課程'] = rawRec.flatMap(t => splitRecommendedCourses(t, pdfTitles));
        }
    }

    return courses;
}

// ==========================================
// List Field Extraction & Alignment Helpers
// ==========================================

/**
 * Universal helper to normalize list or multiline string values into clean string arrays.
 * Handles both Array and multiline String representations consistently across Word & PDF.
 */
function toCleanItemArray(val) {
    if (!val) return [];
    if (Array.isArray(val)) {
        return val.flatMap(v => String(v).split(/\r?\n+/)).map(s => s.trim()).filter(Boolean);
    }
    return String(val).split(/\r?\n+/).map(s => s.trim()).filter(Boolean);
}

function cleanItemText(str) {
    if (!str) return '';
    return str
        .replace(/^[\s•●\-\*※\d.、()（）]+/, '')
        .replace(/^(?:Domain\s*\d+|[A-Za-z]\d+(?:\.\d+)*)[.、\s]*/i, '')
        .trim();
}

/**
 * Splits text with in-line embedded bullets or numbering (e.g. "1. xxx 2. yyy" or "● xxx ● yyy").
 * Prevents horizontally merged items from staying concatenated across all fields.
 */
function splitInlineBullets(text) {
    if (!text) return [];
    const parts = text.split(/(?<=[^\s])\s+(?=(?:\d+[.、](?!\d)|[•●※·◆▪＊★☆✦✧])\s*)/);
    return parts.map(p => p.trim()).filter(Boolean);
}

/**
 * Splits text with numbered bullets (1., 2.) or symbols (•, ※) into clean array of items.
 * Does not split on hyphens inside words (e.g. Man-in-the-middle).
 */
function splitOutlineItems(text) {
    if (!text) return [];
    const parts = text.split(/(?=\b\d+[.、]|\s*[•●※·◆▪＊★☆✦✧]\s*|(?:^|[\r\n])\s*[-*]\s+|\s+[-*]\s+|(?<=^|[\s\r\n])(?:[一二三四五六七八九十百]+[、.．]|[（(][一二三四五六七八九十百\d]+[)）]|[①-⑳❶-❿㈠-㈩]))/);
    const res = [];
    for (const p of parts) {
        const trimmed = p.trim();
        if (trimmed) res.push(trimmed);
    }
    return res.length > 0 ? res : [text.trim()];
}

/**
 * Extracts list items from Word table cell paragraphs.
 * If paragraphs contain embedded bullets or numbering, splits them cleanly.
 * Also separates compound discount scheme headers (e.g. 課程優惠方案：接早鳥優惠價：) into distinct items.
 */
function extractListItems(paragraphs, fullText) {
    // Pre-process paragraphs: split compound discount headers and concatenated discount schemes
    const DISCOUNT_HEADER_SPLIT = /(?<=[^\s：:\n])\s*(?=(?:早鳥優惠|早鳥優惠價|限時優惠|專案優惠|續報優惠|學生優惠|學生優惠價|學生專屬優惠|企業優惠|原廠優惠|證照優惠|重聽服務|方案\s*\d*)[：:])/;
    if (paragraphs && paragraphs.length > 0) {
        paragraphs = paragraphs.flatMap(p => {
            const m = p.match(/^((?:\d+[.、]\s*)?課程優惠方案[：:])\s*((?:早鳥|限時|專案|續報|學生|企業|方案)[^：:\n]{0,8}[：:][\s\S]+)$/);
            if (m) return [m[1].trim(), ...m[2].trim().split(DISCOUNT_HEADER_SPLIT).map(s => s.trim()).filter(Boolean)];
            if (DISCOUNT_HEADER_SPLIT.test(p)) {
                return p.split(DISCOUNT_HEADER_SPLIT).map(s => s.trim()).filter(Boolean);
            }
            return [p];
        });
    }

    let items = [];
    if (paragraphs && paragraphs.length > 0) {
        for (const p of paragraphs) {
            const pTrimmed = p.trim();
            if (!pTrimmed) continue;
            const split = splitOutlineItems(pTrimmed);
            if (split.length > 1) {
                items.push(...split);
            } else {
                items.push(pTrimmed);
            }
        }
    }
    if (items.length === 0 && fullText && fullText.trim()) {
        const m = fullText.trim().match(/^((?:\d+[.、]\s*)?課程優惠方案[：:])\s*((?:早鳥|限時|專案|續報|學生|企業|方案)[^：:\n]{0,8}[：:][\s\S]+)$/);
        if (m) {
            items = [m[1].trim(), ...splitOutlineItems(m[2].trim())];
        } else {
            items = splitOutlineItems(fullText.trim());
        }
    }
    return items.map(it => it.trim()).filter(Boolean);
}

/**
 * Detects whether a string or array of strings contains explicit list bullet or numbering markers.
 */
function hasBulletMarkers(val) {
    if (!val) return false;
    const arr = Array.isArray(val) ? val : [val];
    return arr.some(s => BULLET_ITEM_PATTERN.test((s || '').trim()));
}

/**
 * Checks whether a text line indicates the start of a discrete note item.
 * Strictly avoids splitting on decimal numbers (e.g. 1.5小時).
 * If current item is already a numbered item (e.g. 1., 2.), ordinary continuation lines are never split.
 */
function isNotesItemStart(text, currentItem) {
    const trimmed = text.trim();
    if (!trimmed) return false;

    // Standalone symbol lines are not new items
    if (/^[®™©\s]+$/.test(trimmed)) return false;

    // 1. Explicit numbering / bullets (Chinese & Arabic & symbols)
    if (BULLET_ITEM_PATTERN.test(trimmed)) return true;

    // 2. Scheme / discount / note headers ending in colon
    if (/^(?:課程優惠方案|限時優惠|早鳥優惠|早鳥優惠價|專案優惠|續報優惠|學生優惠|學生優惠價|學生專屬優惠|重聽服務|原廠優惠|證照優惠|方案\s*\d*)[：:]/.test(trimmed)) return true;
    if (/^[^：:\n]{2,8}[：:]\s*(?:即日起|開課前|報名|原報名|續報|凡報名|參與|可享|贈送|提供|投入|完成|透過)/.test(trimmed)) return true;

    // If current item already started with an explicit number (e.g. "1.", "2."),
    // do NOT split on regular text lines! Only split on a new number or bullet!
    if (currentItem && BULLET_ITEM_PATTERN.test(currentItem.trim())) {
        return false;
    }

    // 4. In unnumbered notes, recognize independent topic starts
    if (/^(?:白天班|晚上班|週末班|假日班|上課時間|報名本課程|課程結束後|贈送\d+|免費提供|本課程與|出席率達|原報名班級|沉浸式體驗|實務應用|技能追蹤|完整練習時數)/.test(trimmed)) {
        return true;
    }

    return false;
}

/**
 * Assembles visual lines in notes into discrete items, preserving multi-line continuations.
 */
function assembleNotesLines(lines) {
    if (!lines || lines.length === 0) return [];
    const items = [];
    let curItem = '';

    for (let i = 0; i < lines.length; i++) {
        const lineText = lines[i].text.trim();
        if (!lineText) continue;
        if (/^[®™©\s]+$/.test(lineText)) continue;

        if (!curItem) {
            curItem = lineText;
        } else {
            if (isNotesItemStart(lineText, curItem)) {
                items.push(curItem);
                curItem = lineText;
            } else {
                curItem += (/[a-zA-Z0-9]$/.test(curItem) && /^[a-zA-Z0-9]/.test(lineText) ? ' ' : '') + lineText;
            }
        }
    }
    if (curItem) items.push(curItem);
    return items.map(it => it.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

/**
 * Extracts note items from PDF section in single-column layout with robust line assembly.
 */
function extractNotesFromPdfSection(secItems) {
    if (!secItems || secItems.length === 0) return [];
    const lines = groupItemsIntoVisualLines(secItems);
    return assembleNotesLines(lines);
}

/**
 * Dynamically detects column boundaries (X coordinates) in a section.
 * Returns array of split X coordinates (empty array for single column).
 */
function getColumnSplits(secItems) {
    if (!secItems || secItems.length === 0) return [];
    
    // 1. Identify true bullet / item markers at the start of a column/line
    const bullets = secItems.filter(it => {
        const trimmed = it.str.trim();
        if (!BULLET_ITEM_PATTERN.test(trimmed)) return false;
        // Bare punctuation or hyphens inside running text cannot be column bullets
        if (trimmed === '-' || trimmed === '–' || trimmed === '—' || trimmed === '·') {
            const hasPreceding = secItems.some(other => 
                other !== it && 
                Math.abs(other.y - it.y) <= 3 && 
                other.x < it.x && 
                (other.x + (other.w || other.str.length * 6)) >= it.x - 8
            );
            if (hasPreceding) return false;
        }
        return true;
    });

    if (bullets.length >= 2) {
        const bulletXs = bullets.map(b => b.x).sort((a, b) => a - b);
        const clusters = [];
        for (const x of bulletXs) {
            const last = clusters[clusters.length - 1];
            // Columns on A4 page must have significant horizontal separation (>= 60pt)
            if (!last || x - last.max > 60) {
                clusters.push({ min: x, max: x, count: 1 });
            } else {
                last.max = Math.max(last.max, x);
                last.min = Math.min(last.min, x);
                last.count++;
            }
        }
        // Column 1 is near left margin (x < 200).
        // Any subsequent column (Column 2, Column 3) MUST start at x > 200.
        const validClusters = clusters.filter((c, idx) => {
            if (idx === 0) return true;
            return c.min > 200 && (c.count >= 2 || (clusters.length <= 3 && c.count >= 1));
        });

        if (validClusters.length >= 2) {
            const splits = [];
            for (let i = 1; i < validClusters.length; i++) {
                const colStartX = validClusters[i].min;
                const prevColItems = secItems.filter(it => it.x < colStartX);
                let maxRight = 0;
                for (const it of prevColItems) {
                    const charW = /[\u4e00-\u9fa5]/.test(it.str) ? 10 : 6;
                    const right = it.x + (it.w || it.str.length * charW);
                    if (right > maxRight) maxRight = right;
                }
                let splitX;
                if (colStartX > maxRight) {
                    splitX = Math.round((maxRight + colStartX) / 2);
                } else {
                    splitX = Math.round(colStartX - 2);
                }
                splits.push(splitX);
            }
            return splits;
        }
    }

    // 2. Gutter search fallback between 220 and 380
    const intervals = secItems.map(it => {
        const charWidth = /[\u4e00-\u9fa5]/.test(it.str) ? 10 : 6;
        return { start: it.x, end: it.x + Math.max(it.str.length * charWidth, 10) };
    });

    let bestGutterStart = -1, bestGutterWidth = 0;
    let inGutter = false, currentStart = 0;
    for (let testX = 220; testX <= 380; testX += 2) {
        const hasText = intervals.some(inv => testX >= inv.start && testX <= inv.end);
        if (!hasText) {
            if (!inGutter) {
                inGutter = true;
                currentStart = testX;
            }
        } else {
            if (inGutter) {
                inGutter = false;
                const width = testX - currentStart;
                if (width > bestGutterWidth) {
                    bestGutterWidth = width;
                    bestGutterStart = currentStart;
                }
            }
        }
    }
    if (bestGutterWidth >= 16) {
        return [Math.round(bestGutterStart + bestGutterWidth / 2)];
    }
    return [];
}

/**
 * Groups items into horizontal visual lines based on vertical baseline (Y).
 */
function groupItemsIntoVisualLines(items) {
    if (!items || items.length === 0) return [];
    const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    let curY = null;
    let curLineItems = [];

    for (const it of sorted) {
        if (curY === null || Math.abs(curY - it.y) > 3.5) {
            if (curLineItems.length > 0) {
                curLineItems.sort((a, b) => a.x - b.x);
                const startX = curLineItems[0].x;
                const lastItem = curLineItems[curLineItems.length - 1];
                const charW = /[\u4e00-\u9fa5]/.test(lastItem.str) ? 10 : 6;
                const endX = lastItem.x + (lastItem.w || Math.max(lastItem.str.length * charW, 10));
                lines.push({
                    y: curY,
                    startX,
                    endX,
                    text: curLineItems.map(i => i.str).join(' ').replace(/\s+/g, ' ').trim()
                });
            }
            curY = it.y;
            curLineItems = [it];
        } else {
            curLineItems.push(it);
        }
    }
    if (curLineItems.length > 0) {
        curLineItems.sort((a, b) => a.x - b.x);
        const startX = curLineItems[0].x;
        const lastItem = curLineItems[curLineItems.length - 1];
        const charW = /[\u4e00-\u9fa5]/.test(lastItem.str) ? 10 : 6;
        const endX = lastItem.x + (lastItem.w || Math.max(lastItem.str.length * charW, 10));
        lines.push({
            y: curY,
            startX,
            endX,
            text: curLineItems.map(i => i.str).join(' ').replace(/\s+/g, ' ').trim()
        });
    }
    return lines;
}

/**
 * Assembles visual lines into discrete list items, intelligently distinguishing
 * new items from line-wrapped continuation text.
 */
function assembleLinesIntoItems(lines, isNotes = false) {
    if (!lines || lines.length === 0) return [];
    const items = [];
    let currentItem = '';

    const colMinX = Math.min(...lines.map(l => l.startX));
    const colMaxX = Math.max(...lines.map(l => l.endX));

    function isNewItemStart(line, prevLine) {
        const text = line.text;
        const prevText = prevLine ? prevLine.text : '';
        if (/^[®™©\s]+$/.test(text.trim())) return false;

        // If current item has unclosed parenthesis, ordinary line continuation belongs to it
        if (currentItem) {
            const openCount = (currentItem.match(/[(（]/g) || []).length;
            const closeCount = (currentItem.match(/[)）]/g) || []).length;
            if (openCount > closeCount && !BULLET_ITEM_PATTERN.test(text.trim())) {
                return false;
            }
        }

        // 1. Explicit bullet / numbering (Chinese, Arabic, circled, symbols, outline headers)
        if (BULLET_ITEM_PATTERN.test(text.trim())) return true;

        // 2. Colon headers (e.g. 課程優惠方案：, 限時優惠：, 續報優惠：, 學生優惠價：, 重聽服務：, 實務應用：)
        if (/^[^：:\n]{2,8}[：:]/.test(text)) return true;

        // 3. If previous line ended with colon '：'
        if (prevText && /[：:]\s*$/.test(prevText)) return true;

        // 4. In notes (備註事項), phrases starting at base margin
        if (isNotes) {
            if (/^(?:白天班|晚上班|週末班|假日班|上課時間|報名|課程結束|贈送|免費|本課程|出席率|重聽|早鳥|限時)/.test(text)) return true;
            if (line.startX <= colMinX + 4 && !/[，、(（與和或的之及]$/.test(prevText)) {
                return true;
            }
        }

        // 5. In general, if previous line ended before right margin, author finished that item
        if (prevLine && prevLine.endX < colMaxX - 35) {
            if (!/[，、(（與和或的之及]$/.test(prevText) && line.startX <= colMinX + 6) {
                return true;
            }
        }

        // 6. If previous line closed parenthesis and this line is at margin
        if (prevLine && /[)）]$/.test(prevText) && line.startX <= colMinX + 4) {
            return true;
        }

        return false;
    }

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/^[®™©\s]+$/.test(line.text.trim())) continue;
        const prevLine = i > 0 ? lines[i - 1] : null;

        if (!currentItem) {
            currentItem = line.text;
        } else {
            if (isNewItemStart(line, prevLine)) {
                if (currentItem) items.push(currentItem.trim());
                currentItem = line.text;
            } else {
                const needSpace = /[a-zA-Z0-9]$/.test(currentItem) && /^[a-zA-Z0-9]/.test(line.text);
                currentItem += (needSpace ? ' ' : '') + line.text;
            }
        }
    }
    if (currentItem) items.push(currentItem.trim());
    return items.map(it => it.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

/**
 * Extracts list items from PDF section items inside a bounding box.
 * Preserves dynamic multi-column reading order and sorts by explicit numbers if present.
 */
function extractListFromPdfSection(secItems, isNotes = false) {
    if (!secItems || secItems.length === 0) return [];
    const splits = getColumnSplits(secItems);

    if (splits.length === 0) {
        const lines = groupItemsIntoVisualLines(secItems);
        return assembleLinesIntoItems(lines, isNotes);
    }

    const colItems = [];
    for (let i = 0; i <= splits.length; i++) {
        const minX = i === 0 ? 0 : splits[i - 1];
        const maxX = i === splits.length ? 9999 : splits[i];
        const itemsInCol = secItems.filter(it => it.x >= minX && it.x < maxX);
        if (itemsInCol.length > 0) {
            const lines = groupItemsIntoVisualLines(itemsInCol);
            const assembled = assembleLinesIntoItems(lines, isNotes);
            colItems.push(...assembled);
        }
    }

    const finalItems = isNotes ? colItems : colItems.flatMap(it => splitInlineBullets(it));
    return isNotes ? finalItems : sortListItems(finalItems);
}

/**
 * Sorts array of list items by leading number (e.g. 1., 2., 3.) if present.
 */
function sortListItems(items) {
    if (!items || items.length <= 1) return items;
    const numbered = items.map(it => {
        const m = it.match(/^(\d+)(?:[.、．]|\s+)/);
        return { num: m ? parseInt(m[1], 10) : null, text: it };
    });
    const nums = numbered.map(x => x.num).filter(n => n !== null);
    if (nums.length < items.length * 0.6) return items;

    // Check for duplicate numbers (e.g. multiple "1.", "2." indicating nested/domain lists)
    const uniqueNums = new Set(nums);
    if (uniqueNums.size < nums.length * 0.8) {
        return items; // Nested or grouped list with repeated numbering: keep natural layout order!
    }

    return [...items].sort((a, b) => {
        const na = a.match(/^(\d+)(?:[.、．]|\s+)/);
        const nb = b.match(/^(\d+)(?:[.、．]|\s+)/);
        if (na && nb) return parseInt(na[1], 10) - parseInt(nb[1], 10);
        if (na) return -1;
        if (nb) return 1;
        return 0;
    });
}

function findRawSliceLength(rawText, targetNormalized) {
    if (!rawText || !targetNormalized) return 0;
    for (let i = 1; i <= rawText.length; i++) {
        if (normalizeText(rawText.slice(0, i)) === targetNormalized) {
            return i;
        }
    }
    return rawText.length;
}

function findRawPrefixIndex(rawText, targetNormalizedIndex) {
    if (!rawText) return 0;
    for (let i = 1; i <= rawText.length; i++) {
        if (normalizeText(rawText.slice(0, i)).length >= targetNormalizedIndex) {
            return i;
        }
    }
    return rawText.length;
}

function splitRecommendedCourses(text, catalogTitles = []) {
    if (!text) return [];
    const trimmed = text.trim();
    if (!trimmed) return [];

    // 1. If contains newlines
    if (trimmed.includes('\n')) {
        return trimmed.split(/\n+/).flatMap(s => splitRecommendedCourses(s, catalogTitles)).map(s => s.trim()).filter(Boolean);
    }

    // 2. If separated by semicolon
    if (trimmed.includes('；') || trimmed.includes(';')) {
        return trimmed.split(/[；;]+/).flatMap(s => splitRecommendedCourses(s, catalogTitles)).map(s => s.trim()).filter(Boolean);
    }

    // 3. Embedded course code: preceded by non-ASCII (Chinese) or whitespace, followed by code + colon
    if (/(?<=[^\x00-\x7F\s]|\s)(?=[A-Za-z0-9/_-]{2,12}\s*[：:])/.test(trimmed)) {
        return trimmed.split(/(?<=[^\x00-\x7F\s]|\s)(?=[A-Za-z0-9/_-]{2,12}\s*[：:])/).flatMap(s => splitRecommendedCourses(s, catalogTitles)).map(s => s.trim()).filter(Boolean);
    }

    // 4. If catalogTitles provided, split by matching titles
    if (catalogTitles && catalogTitles.length > 0) {
        const validTitles = catalogTitles.filter(t => t && t.length >= 4 && !/^[A-Za-z0-9/_-]+$/.test(t));
        const sorted = [...validTitles].sort((a, b) => b.length - a.length);

        const result = [];
        let remaining = trimmed;

        while (remaining.length > 0) {
            const codePrefixMatch = remaining.match(/^([A-Za-z0-9/_-]{2,12}\s*[：:\s]\s*)(.*)$/);
            let checkText = remaining;
            let codePrefix = '';
            if (codePrefixMatch) {
                codePrefix = codePrefixMatch[1];
                checkText = codePrefixMatch[2];
            }

            const normCheck = normalizeText(checkText);
            if (!normCheck) {
                result.push(remaining);
                break;
            }

            const matchedPrefix = sorted.find(t => normCheck.startsWith(normalizeText(t)));
            if (matchedPrefix) {
                const len = findRawSliceLength(checkText, normalizeText(matchedPrefix));
                result.push((codePrefix + checkText.slice(0, len)).trim());
                remaining = checkText.slice(len).trim();
                continue;
            }

            if (!codePrefix) {
                const normR = normalizeText(remaining);
                let bestIdx = -1;
                let bestTitle = null;
                for (const t of sorted) {
                    const normT = normalizeText(t);
                    const idx = normR.indexOf(normT);
                    if (idx > 0 && (bestIdx === -1 || idx < bestIdx)) {
                        bestIdx = idx;
                        bestTitle = t;
                    }
                }

                if (bestIdx > 0 && bestTitle) {
                    const splitIdx = findRawPrefixIndex(remaining, bestIdx);
                    const prefixPart = remaining.slice(0, splitIdx).trim();
                    if (prefixPart) result.push(prefixPart);
                    remaining = remaining.slice(splitIdx).trim();
                    continue;
                }
            }

            result.push(remaining);
            break;
        }
        if (result.length > 1) {
            return result.filter(Boolean);
        }
    }

    return [trimmed];
}

/**
 * Analyzes item text to determine its structural hierarchy level (main number, bullet sub-item, header).
 */
function analyzeItemHierarchy(itemText, prevH, fieldLabel) {
    if (!itemText) return { level: 1, type: 'text', badgeText: '•', indent: '', text: '' };
    const raw = itemText.trim();

    // 1. Explicit bullet dot (●, •, ※, ·, -, *)
    const bMatch = raw.match(/^([●•※\-\*·])\s*(.*)$/);
    if (bMatch) {
        return {
            level: 2,
            type: 'bullet',
            badgeText: '●',
            indent: 'md:ml-6 ml-3 pl-3 border-l-2 border-indigo-200',
            cleanText: bMatch[2] || raw
        };
    }

    // 2. Explicit numbered items (1., 2., 3., 1、, 2、, etc.) - NEVER match decimals like 1.5!
    const nMatch = raw.match(/^(\d+)[.、]\s*(.*)$/);
    if (nMatch && !/^\d+\.\d/.test(raw)) {
        return {
            level: 1,
            type: 'number',
            badgeText: nMatch[1],
            indent: '',
            cleanText: raw
        };
    }

    // 3. Section/Scheme Header ending with colon (e.g. 課程優惠方案：, 重聽服務：, 實務應用：)
    const hMatch = raw.match(/^([^：:\n]{2,10})[：:]\s*(.*)$/);
    if (hMatch && (!hMatch[2] || hMatch[2].length === 0)) {
        return {
            level: 0,
            type: 'header',
            badgeText: '方案',
            indent: '',
            cleanText: raw
        };
    }

    // 4. Sub-clauses under a header (e.g. 限時優惠：..., 續報優惠：..., 學生優惠價：...)
    if (fieldLabel === '備註事項' && prevH && (prevH.type === 'header' || (prevH.type === 'number' && /優惠方案/.test(prevH.cleanText)) || prevH.type === 'sub-clause')) {
        if (/^(?:限時|續報|學生|早鳥|企業|校園|加贈|贈送)/.test(raw)) {
            return {
                level: 2,
                type: 'sub-clause',
                badgeText: '●',
                indent: 'md:ml-6 ml-3 pl-3 border-l-2 border-indigo-200',
                cleanText: raw
            };
        }
    }

    return {
        level: 1,
        type: 'text',
        badgeText: '',
        indent: '',
        cleanText: raw
    };
}

/**
 * Needleman-Wunsch sequence alignment for Word and PDF list items.
 * Guarantees optimal horizontal alignment so missing or modified items are clearly visible side-by-side.
 */
function alignListItems(wItems, pItems, fieldLabel) {
    if (!wItems.length && !pItems.length) {
        return { status: 'green', desc: `雙方皆無${fieldLabel}`, details: [] };
    }

    const sortedP = fieldLabel === '備註事項' ? pItems : sortListItems(pItems);
    const sortedW = fieldLabel === '備註事項' ? wItems : sortListItems(wItems);

    const m = sortedW.length;
    const n = sortedP.length;

    function score(i, j) {
        const w = sortedW[i];
        const p = sortedP[j];
        const wNorm = normalizeText(cleanItemText(w));
        const pNorm = normalizeText(cleanItemText(p));
        if (!wNorm && !pNorm) return 1.0;
        if (wNorm === pNorm) return 2.0;
        const sim = calculateSimilarity(wNorm, pNorm);
        if (sim >= 0.7 || wNorm.includes(pNorm) || pNorm.includes(wNorm)) {
            return 1.0 + sim;
        }
        return -1.0;
    }

    const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
    const gap = -0.5;

    for (let i = 0; i <= m; i++) dp[i][0] = i * gap;
    for (let j = 0; j <= n; j++) dp[0][j] = j * gap;

    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            const matchScore = dp[i - 1][j - 1] + score(i - 1, j - 1);
            const delScore = dp[i - 1][j] + gap;
            const insScore = dp[i][j - 1] + gap;
            dp[i][j] = Math.max(matchScore, delScore, insScore);
        }
    }

    const aligned = [];
    let i = m, j = n;
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && Math.abs(dp[i][j] - (dp[i - 1][j - 1] + score(i - 1, j - 1))) < 1e-6) {
            aligned.unshift({ word: sortedW[i - 1], pdf: sortedP[j - 1] });
            i--;
            j--;
        } else if (i > 0 && Math.abs(dp[i][j] - (dp[i - 1][j] + gap)) < 1e-6) {
            aligned.unshift({ word: sortedW[i - 1], pdf: null });
            i--;
        } else {
            aligned.unshift({ word: null, pdf: sortedP[j - 1] });
            j--;
        }
    }

    // Consolidate N:1 and 1:N compound list items across contiguous blocks
    // Resolves cases where multiple Word paragraphs correspond to a single PDF item (e.g. iPASSE, GPTEH)
    // or one Word item corresponds to multiple PDF lines (e.g. PMP001)
    const consolidated = [];
    let cIdx = 0;
    while (cIdx < aligned.length) {
        let end = cIdx;
        let wCount = 0, pCount = 0;
        const wTexts = [], pTexts = [];

        while (end < aligned.length) {
            const cur = aligned[end];
            if (cur.word) { wCount++; wTexts.push(cur.word); }
            if (cur.pdf) { pCount++; pTexts.push(cur.pdf); }

            if (wCount >= 2 && pCount >= 2) break;

            end++;
            if ((wCount >= 2 && pCount === 1) || (wCount === 1 && pCount >= 2)) {
                const combinedW = wTexts.map(w => cleanItemText(w)).join(' ');
                const combinedP = pTexts.map(p => cleanItemText(p)).join(' ');
                const wNorm = normalizeText(combinedW);
                const pNorm = normalizeText(combinedP);
                const sim = calculateSimilarity(wNorm, pNorm);

                if (wCount >= 2 && pCount === 1 && end < aligned.length && aligned[end].word && !aligned[end].pdf) {
                    continue;
                }
                if (wCount === 1 && pCount >= 2 && end < aligned.length && !aligned[end].word && aligned[end].pdf) {
                    continue;
                }

                if (sim >= 0.7 || wNorm.includes(pNorm) || pNorm.includes(wNorm)) {
                    consolidated.push({
                        word: wTexts.join('\n'),
                        pdf: pTexts.join('\n')
                    });
                    cIdx = end;
                    wCount = -1;
                    break;
                }
            }
        }

        if (wCount !== -1) {
            consolidated.push(aligned[cIdx]);
            cIdx++;
        }
    }

    const details = [];
    let matchCount = 0;
    let hasRed = false;
    let hasYellow = false;
    let prevH = null;
    let currentMainNum = 0;
    let mainCount = 0;
    let subCount = 0;

    consolidated.forEach((pair, idx) => {
        const w = pair.word;
        const p = pair.pdf;
        const sampleText = (p && /^\d+[.、]/.test(p.trim())) ? p : (w || p || '');
        const h = analyzeItemHierarchy(sampleText, prevH, fieldLabel);

        if (h.type === 'number') {
            currentMainNum = parseInt(h.badgeText, 10);
        }
        if (h.level >= 2) subCount++;
        else mainCount++;
        prevH = h;

        let itemStatus = 'green';
        let itemDesc = '相符';

        if (w && p) {
            const wNorm = normalizeText(cleanItemText(w));
            const pNorm = normalizeText(cleanItemText(p));
            if (wNorm === pNorm) {
                matchCount++;
                itemStatus = 'green';
                itemDesc = '相符';
            } else if (calculateSimilarity(wNorm, pNorm) >= 0.7 || wNorm.includes(pNorm) || pNorm.includes(wNorm)) {
                matchCount++;
                hasYellow = true;
                itemStatus = 'yellow';
                itemDesc = '文字微差';
            } else {
                hasRed = true;
                itemStatus = 'red';
                itemDesc = '項目內容不符';
            }
        } else if (w && !p) {
            hasRed = true;
            itemStatus = 'red';
            itemDesc = 'PDF 漏排此項目';
        } else if (!w && p) {
            hasYellow = true;
            itemStatus = 'yellow';
            itemDesc = 'PDF 多排此項目';
        }

        details.push({
            index: idx + 1,
            status: itemStatus,
            word: w,
            pdf: p,
            desc: itemDesc,
            hierarchy: h
        });
    });

    let overallStatus = 'green';
    let overallDesc = `${fieldLabel}相符 (${consolidated.length} 項)`;
    if (hasRed) {
        overallStatus = 'red';
        overallDesc = `${fieldLabel}內容有缺漏或不符 (Word: ${sortedW.length}項, PDF: ${sortedP.length}項)`;
    } else if (hasYellow) {
        overallStatus = 'yellow';
        const hasMissingOrExtra = consolidated.some(c => !c.word || !c.pdf);
        if (hasMissingOrExtra) {
            overallDesc = `${fieldLabel}文字有微差或增減 (Word: ${sortedW.length}項, PDF: ${sortedP.length}項)`;
        } else {
            overallDesc = `${fieldLabel}文字有微差 (共 ${consolidated.length} 項)`;
        }
    }

    return {
        status: overallStatus,
        desc: overallDesc,
        details
    };
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
        let bestPdf = null;
        let bestIdx = -1;
        let bestScore = 0;

        const wCode = (wCourse.course_code || wCourse['課程代碼'] || '').trim().toLowerCase();
        const wZh = normalizeText(wCourse.course_name_zh || wCourse['課程名稱'] || '');
        const wEn = normalizeText(wCourse.course_name_en || wCourse['英文名稱'] || '');

        for (let i = 0; i < pdfCourses.length; i++) {
            if (matchedPdfIndices.has(i)) continue;
            const p = pdfCourses[i];

            const pCode = (p.course_code || p['課程代碼'] || '').trim().toLowerCase();
            const pZh = normalizeText(p.course_name_zh || p['課程名稱'] || '');
            const pEn = normalizeText(p.course_name_en || p['英文名稱'] || '');

            // Anchor 1: 課程代碼完全一致 (如 CCNA, AZ-104, BCIC, iFCI)
            const codeMatch = wCode && pCode && (wCode === pCode);

            // Anchor 2: 英文課名完全相符或相似度 > 85% (忽略大小寫與前後符號)
            let enSim = 0;
            let enMatch = false;
            if (wEn && pEn) {
                if (wEn === pEn) {
                    enMatch = true;
                    enSim = 1.0;
                } else {
                    enSim = calculateSimilarity(wEn, pEn);
                    if (enSim > 0.85 || wEn.includes(pEn) || pEn.includes(wEn)) {
                        enMatch = true;
                    }
                }
            }

            // Anchor 3: 中文名稱相似度 > 80%
            let zhSim = 0;
            let zhMatch = false;
            if (wZh && pZh) {
                if (wZh === pZh) {
                    zhMatch = true;
                    zhSim = 1.0;
                } else {
                    zhSim = calculateSimilarity(wZh, pZh);
                    if (zhSim > 0.80 || wZh.includes(pZh) || pZh.includes(wZh)) {
                        zhMatch = true;
                    }
                }
            }

            if (codeMatch || enMatch || zhMatch) {
                let score = 0;
                if (codeMatch) score += 100;
                if (enMatch) score += 90 + enSim * 5;
                if (zhMatch) score += 80 + zhSim * 5;

                if (score > bestScore) {
                    bestScore = score;
                    bestPdf = p;
                    bestIdx = i;
                }
            }
        }

        if (bestPdf) {
            matchedPdfIndices.add(bestIdx);
            results.push(compareSinglePair(wCourse, bestPdf));
        } else {
            // Missing in PDF
            const wNameZh = wCourse.course_name_zh || wCourse['課程名稱'] || '';
            const wNameEn = wCourse.course_name_en || wCourse['英文名稱'] || '';
            const wCodeVal = wCourse.course_code || wCourse['課程代碼'] || '未標註';

            results.push({
                status: 'gray',
                statusText: 'PDF 排版漏排此課程',
                code: wCodeVal,
                name: wNameZh || wNameEn,
                nameZh: wNameZh,
                nameEn: wNameEn,
                wordCourse: wCourse,
                pdfCourse: null,
                fields: {
                    '中文課名': { status: 'gray', label: '中文課名', word: wNameZh || '-', pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    '英文課名': { status: 'gray', label: '英文課名', word: wNameEn || '-', pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    '課程名稱': { status: 'gray', label: '中文課名', word: wNameZh || '-', pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    '英文名稱': { status: 'gray', label: '英文課名', word: wNameEn || '-', pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    '課程代碼': { status: 'gray', label: '課程代碼', word: wCodeVal, pdf: '-', desc: '缺少' },
                    '時數': { status: 'gray', label: '時數', word: `${wCourse['時數']} 小時`, pdf: '-', desc: '缺少' },
                    '點數': { status: 'gray', label: '點數', word: `${wCourse['點數']} 點`, pdf: '-', desc: '缺少' },
                    '費用': { status: 'gray', label: '費用', word: `${Number(wCourse['費用'] || 0).toLocaleString()} 元`, pdf: '-', desc: '缺少' },
                    '教材': { status: 'gray', label: '教材', word: wCourse['教材'] || '-', pdf: '-', desc: '缺少' },
                    '課程內容': createMissingListField('課程內容', wCourse['課程內容'], true),
                    '備註事項': createMissingListField('備註事項', wCourse['備註事項'], true),
                    '後續推薦課程': createMissingListField('後續推薦課程', wCourse['後續推薦課程'], true),
                    '適合對象': createMissingListField('適合對象', wCourse['適合對象'], true),
                    '預備知識': createMissingListField('預備知識', wCourse['預備知識'], true),
                    '先修課程': createMissingListField('先修課程', wCourse['先修課程'], true),
                    '課程目標': { status: 'gray', label: '課程目標', word: (Array.isArray(wCourse['課程目標']) ? wCourse['課程目標'].join('\n') : wCourse['課程目標']) || '-', pdf: '-', desc: '缺少' },
                    '學會技能': { status: 'gray', label: '學會技能', word: (Array.isArray(wCourse['學會技能']) ? wCourse['學會技能'].join('\n') : wCourse['學會技能']) || '-', pdf: '-', desc: '缺少' }
                },
                layoutMode: 'none',
                layoutModeText: 'PDF 漏排此課程'
            });
        }
    }

    // Check for extra courses in PDF not in Word
    for (let i = 0; i < pdfCourses.length; i++) {
        if (!matchedPdfIndices.has(i)) {
            const pCourse = pdfCourses[i];
            const pNameZh = pCourse.course_name_zh || pCourse['課程名稱'] || '';
            const pNameEn = pCourse.course_name_en || pCourse['英文名稱'] || '';
            const pCodeVal = pCourse.course_code || pCourse['課程代碼'] || '未標註';

            results.push({
                status: 'gray',
                statusText: 'Word 原稿無此課程 (PDF 多出)',
                code: pCodeVal,
                name: pNameZh || pNameEn,
                nameZh: pNameZh,
                nameEn: pNameEn,
                wordCourse: null,
                pdfCourse: pCourse,
                fields: {
                    '中文課名': { status: 'gray', label: '中文課名', word: '(未找到)', pdf: pNameZh || '-', desc: 'Word 原稿未列出此課' },
                    '英文課名': { status: 'gray', label: '英文課名', word: '(未找到)', pdf: pNameEn || '-', desc: 'Word 原稿未列出此課' },
                    '課程名稱': { status: 'gray', label: '中文課名', word: '(未找到)', pdf: pNameZh || '-', desc: 'Word 原稿未列出此課' },
                    '英文名稱': { status: 'gray', label: '英文課名', word: '(未找到)', pdf: pNameEn || '-', desc: 'Word 原稿未列出此課' },
                    '課程代碼': { status: 'gray', label: '課程代碼', word: '-', pdf: pCodeVal, desc: '原稿無' },
                    '時數': { status: 'gray', label: '時數', word: '-', pdf: `${pCourse['時數']} 小時`, desc: '原稿無' },
                    '點數': { status: 'gray', label: '點數', word: '-', pdf: `${pCourse['點數']} 點`, desc: '原稿無' },
                    '費用': { status: 'gray', label: '費用', word: '-', pdf: `${Number(pCourse['費用'] || 0).toLocaleString()} 元`, desc: '原稿無' },
                    '教材': { status: 'gray', label: '教材', word: '-', pdf: pCourse['教材'] || '-', desc: '原稿無' },
                    '課程內容': createMissingListField('課程內容', pCourse['課程內容'], false),
                    '備註事項': createMissingListField('備註事項', pCourse['備註事項'], false),
                    '後續推薦課程': createMissingListField('後續推薦課程', pCourse['後續推薦課程'], false),
                    '適合對象': createMissingListField('適合對象', pCourse['適合對象'], false),
                    '預備知識': createMissingListField('預備知識', pCourse['預備知識'], false),
                    '先修課程': createMissingListField('先修課程', pCourse['先修課程'], false),
                    '課程目標': { status: 'gray', label: '課程目標', word: '-', pdf: (Array.isArray(pCourse['課程目標']) ? pCourse['課程目標'].join('\n') : pCourse['課程目標']) || '-', desc: '原稿無' },
                    '學會技能': { status: 'gray', label: '學會技能', word: '-', pdf: (Array.isArray(pCourse['學會技能']) ? pCourse['學會技能'].join('\n') : pCourse['學會技能']) || '-', desc: '原稿無' }
                },
                layoutMode: 'none',
                layoutModeText: 'Word 原稿無此課程'
            });
        }
    }

    return results;
}

function createMissingListField(label, rawVal, isWord) {
    const arr = Array.isArray(rawVal) ? rawVal : (rawVal ? [rawVal] : []);
    let prevH = null;
    let mainNum = 0;
    return {
        status: 'gray',
        label,
        isList: true,
        word: isWord ? `${arr.length} 個項目` : '-',
        pdf: isWord ? '-' : `${arr.length} 個項目`,
        desc: isWord ? 'PDF 缺少此課程' : 'Word 原稿無此課程',
        details: arr.map((item, i) => {
            const h = analyzeItemHierarchy(item, prevH, label);
            if (h.type === 'number') {
                mainNum = parseInt(h.badgeText, 10);
            } else if (h.type === 'text' && !/^[●•※\-\*·]/.test(item)) {
                mainNum++;
                if (!h.badgeText || h.badgeText === '•') h.badgeText = String(mainNum);
            }
            prevH = h;
            return {
                index: i + 1,
                status: 'gray',
                word: isWord ? item : null,
                pdf: isWord ? null : item,
                desc: isWord ? '未排入 PDF' : '原稿未列出',
                hierarchy: h
            };
        })
    };
}

function compareSinglePair(w, p) {
    const fields = {};
    let hasRed = false;
    let hasYellow = false;

    // 1. 課程代碼
    const wCode = (w.course_code || w['課程代碼'] || '').trim();
    const pCode = (p.course_code || p['課程代碼'] || '').trim();
    const codeMatch = wCode && pCode && (wCode.toLowerCase() === pCode.toLowerCase());
    fields['課程代碼'] = {
        label: '課程代碼',
        word: wCode || '-',
        pdf: pCode || '-',
        status: codeMatch ? 'green' : (wCode && pCode ? 'red' : 'yellow'),
        desc: codeMatch ? '代碼相符' : (wCode && pCode ? '課程代碼不相符！' : '代碼未標註')
    };
    if (wCode && pCode && !codeMatch) hasRed = true;

    // 2. 中文課名
    const wNameZh = (w.course_name_zh || w['課程名稱'] || '').trim();
    const pNameZh = (p.course_name_zh || p['課程名稱'] || '').trim();
    const wZhNorm = normalizeText(wNameZh);
    const pZhNorm = normalizeText(pNameZh);

    if (!wZhNorm && !pZhNorm) {
        fields['中文課名'] = { label: '中文課名', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無中文課名' };
    } else if (wZhNorm === pZhNorm) {
        fields['中文課名'] = { label: '中文課名', word: wNameZh, pdf: pNameZh, status: 'green', desc: '中文課名完全相符' };
    } else if (calculateSimilarity(wZhNorm, pZhNorm) > 0.85 || wZhNorm.includes(pZhNorm) || pZhNorm.includes(wZhNorm)) {
        fields['中文課名'] = { label: '中文課名', word: wNameZh, pdf: pNameZh, status: 'yellow', desc: '中文課名文字微差' };
        hasYellow = true;
    } else {
        fields['中文課名'] = { label: '中文課名', word: wNameZh || '(無)', pdf: pNameZh || '(漏排)', status: 'red', desc: '中文課名不一致或錯字！' };
        hasRed = true;
    }
    // Backward compatibility for 課程名稱
    fields['課程名稱'] = fields['中文課名'];

    // 3. 英文課名 (原廠課名/官方名稱)
    const wNameEn = (w.course_name_en || w['英文名稱'] || '').trim();
    const pNameEn = (p.course_name_en || p['英文名稱'] || '').trim();
    const wEnNorm = normalizeText(wNameEn);
    const pEnNorm = normalizeText(pNameEn);

    if (!wEnNorm && !pEnNorm) {
        fields['英文課名'] = { label: '英文課名', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無英文課名' };
    } else if (wEnNorm && !pEnNorm) {
        fields['英文課名'] = { label: '英文課名', word: wNameEn, pdf: '(PDF 漏排英文副標)', status: 'red', desc: 'PDF 漏排英文課名副標！' };
        hasRed = true;
    } else if (!wEnNorm && pEnNorm) {
        fields['英文課名'] = { label: '英文課名', word: '(無)', pdf: pNameEn, status: 'yellow', desc: 'Word 原稿無英文課名，但 PDF 有排版' };
        hasYellow = true;
    } else {
        if (wEnNorm === pEnNorm) {
            fields['英文課名'] = { label: '英文課名', word: wNameEn, pdf: pNameEn, status: 'green', desc: '英文課名完全相符' };
        } else if (calculateSimilarity(wEnNorm, pEnNorm) > 0.85 || wEnNorm.includes(pEnNorm) || pEnNorm.includes(wEnNorm)) {
            fields['英文課名'] = { label: '英文課名', word: wNameEn, pdf: pNameEn, status: 'yellow', desc: '英文課名微差 (單字大小寫或標點差異)' };
            hasYellow = true;
        } else {
            fields['英文課名'] = {
                label: '英文課名',
                word: wNameEn,
                pdf: pNameEn,
                status: 'red',
                desc: `英文課名不一致！Word 為「${wNameEn}」，但 PDF 為「${pNameEn}」`
            };
            hasRed = true;
        }
    }
    // Backward compatibility for 英文名稱
    fields['英文名稱'] = fields['英文課名'];

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

    // ==========================================
    // 排版模式判定（課程內容 vs 課程目標）
    // 業務規則：因 PDF 排版空間限制，美編有時採「課程目標」排版，有時採「課程內容」排版。
    // 以 PDF 排版結果為準校對 Word；兩者出現任一種皆視為正常，並明確標記本課程排版模式。
    // ==========================================
    const wContentArr = Array.isArray(w['課程內容']) ? w['課程內容'] : (w['課程內容'] ? [w['課程內容']] : []);
    const pContentArr = Array.isArray(p['課程內容']) ? p['課程內容'] : (p['課程內容'] ? [p['課程內容']] : []);

    const wObjStr = Array.isArray(w['課程目標']) ? w['課程目標'].join('\n') : (w['課程目標'] || '');
    const pObjStr = Array.isArray(p['課程目標']) ? p['課程目標'].join('\n') : (p['課程目標'] || '');
    const wObjNorm = normalizeText(wObjStr);
    const pObjNorm = normalizeText(pObjStr);

    const hasPContent = pContentArr.length > 0;
    const hasPObjective = Boolean(pObjNorm);

    const pContentText = pContentArr.join(' ');
    const pContentNorm = normalizeText(pContentText);
    const pContentMatchesWordObj = hasPContent && wObjNorm && (
        calculateSimilarity(pContentNorm, wObjNorm) > 0.60 ||
        (wObjNorm.length > 20 && pContentNorm.includes(wObjNorm)) ||
        (pContentNorm.length > 20 && wObjNorm.includes(pContentNorm))
    );

    const wContentText = wContentArr.join(' ');
    const wContentNorm = normalizeText(wContentText);
    const pObjMatchesWordContent = hasPObjective && wContentNorm && (
        calculateSimilarity(pObjNorm, wContentNorm) > 0.60 ||
        (wContentNorm.length > 20 && pObjNorm.includes(wContentNorm)) ||
        (pObjNorm.length > 20 && wContentNorm.includes(pObjNorm))
    );

    let layoutMode = 'content'; // 'content' | 'objective' | 'both' | 'none'
    let layoutModeText = '課程內容';

    if (hasPContent && hasPObjective) {
        layoutMode = 'both';
        layoutModeText = '課程目標與課程內容';
    } else if (hasPObjective && !hasPContent) {
        // PDF 標題為課程目標。以 PDF 排版結果為準校對 Word；若 Word 也有課程目標，排版模式為「課程目標」
        const pObjMatchesWordObj = hasPObjective && wObjNorm && (
            calculateSimilarity(pObjNorm, wObjNorm) > 0.40 ||
            wObjNorm.includes(pObjNorm) || pObjNorm.includes(wObjNorm)
        );
        if (pObjMatchesWordObj || !wContentArr.length || Boolean(wObjNorm)) {
            layoutMode = 'objective';
            layoutModeText = '課程目標';
        } else if (pObjMatchesWordContent) {
            layoutMode = 'content';
            layoutModeText = '課程內容 (PDF標為目標)';
        } else {
            layoutMode = 'objective';
            layoutModeText = '課程目標';
        }
    } else if (hasPContent && !hasPObjective) {
        if (pContentMatchesWordObj && (!wContentArr.length || calculateSimilarity(pContentNorm, wContentNorm) < 0.40)) {
            layoutMode = 'objective';
            layoutModeText = '課程目標 (PDF標為內容)';
        } else {
            layoutMode = 'content';
            layoutModeText = '課程內容';
        }
    } else {
        layoutMode = 'none';
        layoutModeText = '未排課程大綱/目標';
    }

    // 7. 課程內容 (依排版模式校對)
    if (layoutMode === 'content' || layoutMode === 'both') {
        const contentItemsToCompare = (layoutMode === 'content' && pObjMatchesWordContent)
            ? (p['課程目標_items'] || splitOutlineItems(pObjStr))
            : pContentArr;

        const contentDiff = alignListItems(
            wContentArr,
            contentItemsToCompare,
            '課程內容'
        );
        fields['課程內容'] = {
            label: '課程內容',
            isList: true,
            layoutTag: 'adopted-content',
            word: `${wContentArr.length} 個項目`,
            pdf: `${contentItemsToCompare.length} 個項目`,
            wordItems: wContentArr,
            pdfItems: contentItemsToCompare,
            status: contentDiff.status,
            desc: contentDiff.desc + (layoutMode === 'both' ? ' (同時排入目標與內容)' : ' (本課採「課程內容」排版)'),
            details: contentDiff.details
        };
        if (contentDiff.status === 'red') hasRed = true;
        if (contentDiff.status === 'yellow') hasYellow = true;
    } else {
        // layoutMode is 'objective' or 'none' -> 課程內容免排，視為正常！
        fields['課程內容'] = {
            label: '課程內容',
            isList: true,
            layoutTag: 'skipped-objective',
            word: `${wContentArr.length} 個項目`,
            pdf: '(依排版規則免排)',
            wordItems: wContentArr,
            pdfItems: [],
            status: 'gray', // ⚪ 正常略過，嚴禁亮紅燈！
            desc: `依版面排版規則略過 (美編採「${layoutModeText}」排版，免排課程內容，正常)`,
            details: wContentArr.map((item, idx) => ({
                index: idx + 1,
                status: 'gray',
                word: item,
                pdf: '(依排版規則免排)',
                desc: `本課採「${layoutModeText}」排版，此項免排 (正常)`,
                hierarchy: { level: 1, type: 'text', badgeText: String(idx + 1), indent: '', cleanText: item }
            }))
        };
    }

    // 8. 備註事項 (Rule 2 & 4: 獨立欄位隔離，List[str] 陣列與水平對齊)
    const notesDiff = alignListItems(
        Array.isArray(w['備註事項']) ? w['備註事項'] : (w['備註事項'] ? [w['備註事項']] : []),
        Array.isArray(p['備註事項']) ? p['備註事項'] : (p['備註事項'] ? [p['備註事項']] : []),
        '備註事項'
    );
    fields['備註事項'] = {
        label: '備註事項',
        isList: true,
        word: `${(w['備註事項'] || []).length} 個項目`,
        pdf: `${(p['備註事項'] || []).length} 個項目`,
        wordItems: Array.isArray(w['備註事項']) ? w['備註事項'] : [],
        pdfItems: Array.isArray(p['備註事項']) ? p['備註事項'] : [],
        status: notesDiff.status,
        desc: notesDiff.desc,
        details: notesDiff.details
    };
    if (notesDiff.status === 'red') hasRed = true;
    if (notesDiff.status === 'yellow') hasYellow = true;

    // 9. 後續推薦課程 (業務規則：PDF 只會抓取 Word 的第一個推薦課程，若有第2、第3個推薦課程未排上 PDF 視為正常)
    let wRecList = Array.isArray(w['後續推薦課程'])
        ? [...w['後續推薦課程']]
        : splitRecommendedCourses(w['後續推薦課程'] || '');
    let pRecList = Array.isArray(p['後續推薦課程'])
        ? [...p['後續推薦課程']]
        : splitRecommendedCourses(p['後續推薦課程'] || '');

    const cleanRec = s => s.replace(/^[A-Za-z0-9/_-]{2,12}\s*[：:]\s*/, '').replace(/^[A-Z0-9_-]*\d[A-Z0-9_-]*\s+(?=[\u4e00-\u9fa5])/, '').trim();

    // Re-split using partner's recommendations if concatenated in Word
    if (wRecList.length > 0 && pRecList.length > 0) {
        const pFirstClean = cleanRec(cleanItemText(pRecList[0]));
        const pFirstNorm = normalizeText(pFirstClean);
        const wFirstClean = cleanRec(cleanItemText(wRecList[0]));
        const wFirstNorm = normalizeText(wFirstClean);

        if (pFirstNorm && wFirstNorm.startsWith(pFirstNorm) && wFirstNorm.length > pFirstNorm.length) {
            const rawW = wRecList[0];
            const splitIdx = findRawSliceLength(rawW, pFirstNorm);
            const part1 = rawW.slice(0, splitIdx).trim();
            const part2 = rawW.slice(splitIdx).trim();
            if (part1 && part2) {
                wRecList = [part1, part2, ...wRecList.slice(1)];
            }
        }
    }

    const wFirstRec = wRecList.length > 0 ? wRecList[0] : '';
    const pFirstRec = pRecList.length > 0 ? pRecList[0] : '';

    const wFirstNorm = normalizeText(cleanRec(cleanItemText(wFirstRec)));
    const pFirstNorm = normalizeText(cleanRec(cleanItemText(pFirstRec)));
    const wAllNorm = normalizeText(cleanRec(cleanItemText(wRecList.join(' '))));

    if (wRecList.length === 0 && pRecList.length === 0) {
        fields['後續推薦課程'] = {
            label: '後續推薦課程',
            word: '(無)',
            pdf: '(無)',
            status: 'green',
            desc: '雙方皆無推薦課程'
        };
    } else if (wRecList.length > 0 && pRecList.length === 0) {
        fields['後續推薦課程'] = {
            label: '後續推薦課程',
            isList: true,
            status: 'red',
            desc: 'PDF 漏排首門推薦課程！',
            details: wRecList.map((c, idx) => ({
                index: idx + 1,
                word: c + (idx === 0 ? ' (首門推薦)' : ' (依規則免排PDF)'),
                pdf: idx === 0 ? null : '(依規則免排)',
                status: idx === 0 ? 'red' : 'gray',
                desc: idx === 0 ? 'PDF 漏排此首門推薦課程！' : '第2門以上未排PDF視為正常',
                hierarchy: { level: 1, type: 'number', badgeText: String(idx + 1), indent: '', cleanText: c }
            }))
        };
        hasRed = true;
    } else if (wRecList.length === 0 && pRecList.length > 0) {
        fields['後續推薦課程'] = {
            label: '後續推薦課程',
            word: '(無)',
            pdf: pRecList.join('\n'),
            status: 'yellow',
            desc: 'Word 原稿無推薦課程，但 PDF 有排版'
        };
        hasYellow = true;
    } else {
        const isMatch = (wFirstNorm === pFirstNorm) || (wAllNorm === pFirstNorm);
        const simFirst = calculateSimilarity(wFirstNorm, pFirstNorm);
        const simAll = calculateSimilarity(wAllNorm, pFirstNorm);
        const maxSim = Math.max(simFirst, simAll);

        let recStatus = 'green';
        let recDesc = '首門推薦課程相符';
        if (isMatch) {
            recStatus = 'green';
            recDesc = wRecList.length > 1
                ? `首門推薦課程相符 (Word 共 ${wRecList.length} 門，美編依規則僅排首門，正常)`
                : '推薦課程相符';
        } else if (maxSim > 0.75 || wFirstNorm.includes(pFirstNorm) || pFirstNorm.includes(wFirstNorm)) {
            recStatus = 'yellow';
            recDesc = '推薦課程文字微差 (首門推薦課程大致相符)';
            hasYellow = true;
        } else {
            recStatus = 'red';
            recDesc = `推薦課程不符！Word 首門為「${wFirstRec}」，但 PDF 為「${pFirstRec}」`;
            hasRed = true;
        }

        const recDetails = [
            {
                index: 1,
                word: wFirstRec + ' (首門推薦)',
                pdf: pFirstRec,
                status: recStatus,
                desc: isMatch ? '首門推薦課程相符' : (recStatus === 'yellow' ? '文字微差' : '首門課程不符'),
                hierarchy: { level: 1, type: 'number', badgeText: '1', indent: '', cleanText: wFirstRec }
            }
        ];
        for (let ri = 1; ri < wRecList.length; ri++) {
            recDetails.push({
                index: ri + 1,
                word: wRecList[ri] + ' (依規則免排PDF)',
                pdf: '(依規則免排)',
                status: 'gray',
                desc: '第2門以上推薦課程未排PDF視為正常',
                hierarchy: { level: 1, type: 'number', badgeText: String(ri + 1), indent: '', cleanText: wRecList[ri] }
            });
        }

        fields['後續推薦課程'] = {
            label: '後續推薦課程',
            isList: true,
            status: recStatus,
            desc: recDesc,
            details: recDetails
        };
    }

    // 10. 適合對象 (動態判定 List vs 純文字段落)
    const wTargetVal = w['適合對象'] || '';
    const pTargetVal = p['適合對象'] || '';
    const wTargetArr = toCleanItemArray(wTargetVal);
    const pTargetArr = toCleanItemArray(pTargetVal);
    const targetIsList = wTargetArr.length > 1 || pTargetArr.length > 1 || hasBulletMarkers(wTargetVal) || hasBulletMarkers(pTargetVal);

    if (targetIsList) {
        const targetDiff = alignListItems(wTargetArr, pTargetArr, '適合對象');
        fields['適合對象'] = {
            label: '適合對象',
            isList: true,
            word: `${wTargetArr.length} 個項目`,
            pdf: `${pTargetArr.length} 個項目`,
            wordItems: wTargetArr,
            pdfItems: pTargetArr,
            status: targetDiff.status,
            desc: targetDiff.desc,
            details: targetDiff.details
        };
        if (targetDiff.status === 'red') hasRed = true;
        if (targetDiff.status === 'yellow') hasYellow = true;
    } else {
        const wTargetStr = Array.isArray(wTargetVal) ? wTargetVal.join('\n') : String(wTargetVal || '');
        const pTargetStr = Array.isArray(pTargetVal) ? pTargetVal.join('\n') : String(pTargetVal || '');
        const wTNorm = normalizeText(wTargetStr);
        const pTNorm = normalizeText(pTargetStr);
        if (!wTNorm && !pTNorm) {
            fields['適合對象'] = { label: '適合對象', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無適合對象' };
        } else if (wTNorm === pTNorm) {
            fields['適合對象'] = { label: '適合對象', word: wTargetStr, pdf: pTargetStr, status: 'green', desc: '適合對象相符' };
        } else if (calculateSimilarity(wTNorm, pTNorm) > 0.70 || wTNorm.includes(pTNorm) || pTNorm.includes(wTNorm)) {
            fields['適合對象'] = { label: '適合對象', word: wTargetStr, pdf: pTargetStr, status: 'yellow', desc: '適合對象文字微差' };
            hasYellow = true;
        } else {
            fields['適合對象'] = { label: '適合對象', word: wTargetStr || '(無)', pdf: pTargetStr || '(漏排)', status: 'red', desc: '適合對象不一致！' };
            hasRed = true;
        }
    }

    // 11. 預備知識 (動態判定 List vs 純文字段落)
    const wPrereqVal = w['預備知識'] || '';
    const pPrereqVal = p['預備知識'] || '';
    const wPrereqArr = toCleanItemArray(wPrereqVal);
    const pPrereqArr = toCleanItemArray(pPrereqVal);
    const prereqIsList = wPrereqArr.length > 1 || pPrereqArr.length > 1 || hasBulletMarkers(wPrereqVal) || hasBulletMarkers(pPrereqVal);

    if (prereqIsList) {
        const prereqDiff = alignListItems(wPrereqArr, pPrereqArr, '預備知識');
        fields['預備知識'] = {
            label: '預備知識',
            isList: true,
            word: `${wPrereqArr.length} 個項目`,
            pdf: `${pPrereqArr.length} 個項目`,
            wordItems: wPrereqArr,
            pdfItems: pPrereqArr,
            status: prereqDiff.status,
            desc: prereqDiff.desc,
            details: prereqDiff.details
        };
        if (prereqDiff.status === 'red') hasRed = true;
        if (prereqDiff.status === 'yellow') hasYellow = true;
    } else {
        const wPrereqStr = Array.isArray(wPrereqVal) ? wPrereqVal.join('\n') : String(wPrereqVal || '');
        const pPrereqStr = Array.isArray(pPrereqVal) ? pPrereqVal.join('\n') : String(pPrereqVal || '');
        const wPNorm = normalizeText(wPrereqStr);
        const pPNorm = normalizeText(pPrereqStr);
        if (!wPNorm && !pPNorm) {
            fields['預備知識'] = { label: '預備知識', word: '(無)', pdf: '(無)', status: 'green', desc: '雙方皆無預備知識' };
        } else if (wPNorm === pPNorm) {
            fields['預備知識'] = { label: '預備知識', word: wPrereqStr, pdf: pPrereqStr, status: 'green', desc: '預備知識相符' };
        } else if (calculateSimilarity(wPNorm, pPNorm) > 0.70 || wPNorm.includes(pPNorm) || pPNorm.includes(wPNorm)) {
            fields['預備知識'] = { label: '預備知識', word: wPrereqStr, pdf: pPrereqStr, status: 'yellow', desc: '預備知識文字微差' };
            hasYellow = true;
        } else {
            fields['預備知識'] = { label: '預備知識', word: wPrereqStr || '(無)', pdf: pPrereqStr || '(漏排)', status: 'red', desc: '預備知識不一致！' };
            hasRed = true;
        }
    }

    // 12. 先修課程 (if present in Word or PDF)
    if ((w['先修課程'] && w['先修課程'].length > 0) || (p['先修課程'] && p['先修課程'].length > 0)) {
        const wPreVal = w['先修課程'] || '';
        const pPreVal = p['先修課程'] || '';
        const wPreArr = toCleanItemArray(wPreVal);
        const pPreArr = toCleanItemArray(pPreVal);
        const preIsList = wPreArr.length > 1 || pPreArr.length > 1 || hasBulletMarkers(wPreVal) || hasBulletMarkers(pPreVal);
        if (preIsList) {
            const preDiff = alignListItems(wPreArr, pPreArr, '先修課程');
            fields['先修課程'] = {
                label: '先修課程',
                isList: true,
                word: `${wPreArr.length} 個項目`,
                pdf: `${pPreArr.length} 個項目`,
                wordItems: wPreArr,
                pdfItems: pPreArr,
                status: preDiff.status,
                desc: preDiff.desc,
                details: preDiff.details
            };
            if (preDiff.status === 'red') hasRed = true;
            if (preDiff.status === 'yellow') hasYellow = true;
        } else {
            const wPreStr = Array.isArray(wPreVal) ? wPreVal.join('\n') : String(wPreVal || '');
            const pPreStr = Array.isArray(pPreVal) ? pPreVal.join('\n') : String(pPreVal || '');
            const wNorm = normalizeText(wPreStr);
            const pNorm = normalizeText(pPreStr);
            if (wNorm === pNorm) {
                fields['先修課程'] = { label: '先修課程', word: wPreStr, pdf: pPreStr, status: 'green', desc: '先修課程相符' };
            } else {
                fields['先修課程'] = { label: '先修課程', word: wPreStr, pdf: pPreStr, status: 'yellow', desc: '先修課程文字微差' };
                hasYellow = true;
            }
        }
    }

    // 13. 課程目標 (依排版模式校對，一律為純文字段落比對，絕不切分成 list items！)
    if (layoutMode === 'objective' || layoutMode === 'both') {
        const objTextToCompare = (layoutMode === 'objective' && pContentMatchesWordObj)
            ? pContentText
            : pObjStr;
        const normPdfObj = normalizeText(objTextToCompare);
        const normCleanWordObj = normalizeText(toCleanItemArray(wObjStr).map(cleanItemText).join(''));
        const normCleanPdfObj = normalizeText(toCleanItemArray(objTextToCompare).map(cleanItemText).join(''));

        if (wObjNorm === normPdfObj || (normCleanWordObj && normCleanWordObj === normCleanPdfObj)) {
            fields['課程目標'] = {
                label: '課程目標',
                layoutTag: 'adopted-objective',
                word: wObjStr,
                pdf: objTextToCompare,
                status: 'green',
                desc: '課程目標完全相符 (本課採「課程目標」排版)'
            };
        } else if (calculateSimilarity(wObjNorm, normPdfObj) > 0.70 ||
                   calculateSimilarity(normCleanWordObj, normCleanPdfObj) > 0.70 ||
                   wObjNorm.includes(normPdfObj) || normPdfObj.includes(wObjNorm) ||
                   normCleanWordObj.includes(normCleanPdfObj) || normCleanPdfObj.includes(normCleanWordObj)) {
            fields['課程目標'] = {
                label: '課程目標',
                layoutTag: 'adopted-objective',
                word: wObjStr,
                pdf: objTextToCompare,
                status: 'yellow',
                desc: '課程目標文字微調 (本課採「課程目標」排版)'
            };
            hasYellow = true;
        } else {
            fields['課程目標'] = {
                label: '課程目標',
                layoutTag: 'adopted-objective',
                word: wObjStr,
                pdf: objTextToCompare,
                status: 'red',
                desc: `課程目標不一致！(本課採「課程目標」排版)`
            };
            hasRed = true;
        }
    } else {
        // layoutMode is 'content' or 'none' -> 課程目標免排，視為正常略過！
        fields['課程目標'] = {
            label: '課程目標',
            layoutTag: 'skipped-content',
            word: wObjStr ? (wObjStr.length > 40 ? `${wObjStr.slice(0, 40)}...` : wObjStr) : '(無)',
            pdf: '(依排版規則免排)',
            status: 'gray', // ⚪ 正常略過，嚴禁亮紅燈！
            desc: `依版面排版規則略過 (美編採「${layoutModeText}」排版，免排課程目標，正常)`
        };
    }

    // 14. 學會技能 (業務規則：學會技能一律免排入 PDF，不用列出對比，僅需於表尾備註，視為正常略過)
    fields['學會技能'] = {
        label: '學會技能',
        word: w['學會技能'] ? `${w['學會技能'].slice(0, 35)}...` : '(無)',
        pdf: '(版面精簡未排版)',
        status: 'gray', // ⚪ 灰色略過，依規則免排
        desc: '版面精簡未排版 / 依規則免排 (正常)'
    };

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

    const finalNameZh = w.course_name_zh || p.course_name_zh || w['課程名稱'] || p['課程名稱'] || '';
    const finalNameEn = w.course_name_en || p.course_name_en || w['英文名稱'] || p['英文名稱'] || '';
    const finalCode = w.course_code || p.course_code || w['課程代碼'] || p['課程代碼'] || '未標註';

    return {
        status: overallStatus,
        statusText: overallText,
        code: finalCode,
        name: finalNameZh || finalNameEn,
        nameZh: finalNameZh,
        nameEn: finalNameEn,
        wordCourse: w,
        pdfCourse: p,
        fields,
        layoutMode,
        layoutModeText
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

    // Render all fields in exact order (從課程代碼、中英文課名、時數費用點數、課程目標/內容...直到推薦課程)
    const fieldsOrder = [
        '課程代碼',
        '中文課名',
        '英文課名',
        '時數',
        '費用',
        '點數',
        '教材',
        '課程目標',
        '課程內容',
        '適合對象',
        '先修課程',
        '預備知識',
        '備註事項',
        '後續推薦課程'
    ];

    let rowsHtml = '';

    for (const key of fieldsOrder) {
        // 1. 學會技能依規則免排，不列入對比列（改於表尾備註說明）
        if (key === '學會技能') continue;

        // 2. 課程目標 vs 課程內容 依排版模式條件式隱藏
        if (key === '課程內容' && item.layoutMode === 'objective') continue;
        if (key === '課程目標' && item.layoutMode === 'content') continue;

        const field = item.fields[key] || (key === '中文課名' ? item.fields['課程名稱'] : (key === '英文課名' ? item.fields['英文名稱'] : null));
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

        if (field.isList && field.details && field.details.length > 0) {
            let itemsHtml = '';
            field.details.forEach(d => {
                let pBorder = 'border-slate-200';
                let pBg = 'bg-white';
                let pTextColor = 'text-slate-800';
                let pBadgeBg = 'bg-slate-100 text-slate-600';

                if (d.status === 'green') {
                    pBorder = 'border-emerald-200';
                    pBg = 'bg-emerald-50/40';
                    pTextColor = 'text-emerald-950';
                    pBadgeBg = 'bg-emerald-100 text-emerald-800';
                } else if (d.status === 'yellow') {
                    pBorder = 'border-amber-200';
                    pBg = 'bg-amber-50/50';
                    pTextColor = 'text-amber-950';
                    pBadgeBg = 'bg-amber-100 text-amber-800';
                } else if (d.status === 'red') {
                    pBorder = 'border-red-300';
                    pBg = 'bg-red-50/70';
                    pTextColor = 'text-red-950 font-semibold';
                    pBadgeBg = 'bg-red-100 text-red-800 font-bold';
                } else if (d.status === 'gray') {
                    pBorder = 'border-slate-200';
                    pBg = 'bg-slate-50';
                    pTextColor = 'text-slate-400 italic';
                    pBadgeBg = 'bg-slate-100 text-slate-500';
                }

                const h = d.hierarchy || { level: 1, type: 'text', badgeText: `${d.index}`, indent: '' };

                // Badge display based on hierarchy level and type
                let wBadgeHtml = '';
                let pBadgeHtml = '';
                if (h.type === 'header') {
                    wBadgeHtml = `<span class="inline-flex items-center justify-center px-1.5 h-5 rounded text-3xs font-bold bg-indigo-50 text-indigo-700 border border-indigo-200 shrink-0 mt-0.5">方案</span>`;
                    pBadgeHtml = `<span class="inline-flex items-center justify-center px-1.5 h-5 rounded text-3xs font-bold ${pBadgeBg} shrink-0 mt-0.5">方案</span>`;
                } else if (h.type === 'bullet' || h.type === 'sub-clause') {
                    wBadgeHtml = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-xs font-bold bg-indigo-50 text-indigo-600 shrink-0 mt-0.5">●</span>`;
                    pBadgeHtml = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-xs font-bold ${pBadgeBg} shrink-0 mt-0.5">●</span>`;
                } else if (h.type === 'number' && h.badgeText) {
                    wBadgeHtml = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-3xs font-mono font-bold bg-slate-100 text-slate-700 shrink-0 mt-0.5">${h.badgeText}</span>`;
                    pBadgeHtml = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-3xs font-mono font-bold ${pBadgeBg} shrink-0 mt-0.5">${h.badgeText}</span>`;
                } else {
                    wBadgeHtml = '';
                    pBadgeHtml = '';
                }

                // Strip leading duplicate bullets if bullet badge is already shown
                const displayWord = (d.word && h.type === 'bullet') ? d.word.replace(/^[●•※\-\*·]\s*/, '') : d.word;
                const displayPdf = (d.pdf && h.type === 'bullet') ? d.pdf.replace(/^[●•※\-\*·]\s*/, '') : d.pdf;

                // Hierarchy row indentation and header styling
                const rowIndentClass = h.level >= 2 ? 'md:ml-6 ml-3 pl-2.5 border-l-2 border-indigo-200/70' : '';
                const headerBoxClass = h.type === 'header' ? 'font-semibold bg-slate-50/80' : '';

                itemsHtml += `
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3 items-stretch ${rowIndentClass}">
                        <!-- Word Item Box -->
                        <div class="p-2.5 rounded-lg border border-slate-200 ${h.type === 'header' ? headerBoxClass : 'bg-white'} flex items-start space-x-2 text-xs text-slate-800 leading-relaxed whitespace-pre-line break-words shadow-2xs">
                            ${wBadgeHtml}
                            <div class="flex-1 min-w-0">
                                ${displayWord ? displayWord : '<span class="text-slate-400 italic">(Word 無此項)</span>'}
                            </div>
                        </div>

                        <!-- PDF Item Box -->
                        <div class="p-2.5 rounded-lg border ${pBorder} ${pBg} ${h.type === 'header' ? headerBoxClass : ''} flex items-start space-x-2 text-xs ${pTextColor} leading-relaxed whitespace-pre-line break-words shadow-2xs">
                            ${pBadgeHtml}
                            <div class="flex-1 min-w-0">
                                ${displayPdf ? displayPdf : '<span class="font-bold text-red-600">❌ (PDF 漏排此項)</span>'}
                                ${d.desc && d.status === 'red' && d.pdf ? `<div class="mt-1 text-3xs text-red-600 font-normal">[${d.desc}]</div>` : ''}
                            </div>
                        </div>
                    </div>
                `;
            });

            const scrollContainerClass = field.details.length > 8 ? 'max-h-[480px] overflow-y-auto pr-1' : '';

            const subItemCount = field.details.filter(d => d.hierarchy && d.hierarchy.level >= 2).length;
            const mainItemCount = field.details.length - subItemCount;
            const isSinglePlainText = field.details.length <= 1 && (!field.details[0] || !field.details[0].hierarchy || field.details[0].hierarchy.type === 'text');
            const countBadgeHtml = isSinglePlainText
                ? ''
                : (subItemCount > 0
                    ? `<span class="mt-1 inline-block text-3xs px-2 py-0.5 rounded font-mono bg-indigo-50 text-indigo-700 border border-indigo-100 font-medium">共 ${mainItemCount} 主項 / ${subItemCount} 子項</span>`
                    : `<span class="mt-1 inline-block text-3xs px-2 py-0.5 rounded font-mono bg-slate-100 text-slate-500 font-normal">共 ${field.details.length} 項</span>`);

            let tagBadgeHtml = '';
            if (field.layoutTag) {
                if (field.layoutTag === 'adopted-content') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-blue-100 text-blue-800 font-semibold">排版採用</span>`;
                } else if (field.layoutTag === 'adopted-objective') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-purple-100 text-purple-800 font-semibold">排版採用</span>`;
                } else if (field.layoutTag === 'skipped-content') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-slate-100 text-slate-500 font-normal">免排 (採內容)</span>`;
                } else if (field.layoutTag === 'skipped-objective') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-slate-100 text-slate-500 font-normal">免排 (採目標)</span>`;
                }
            }

            rowsHtml += `
                <tr class="border-b border-slate-100 last:border-none ${rowBg}">
                    <td class="py-3 px-4 text-xs font-semibold text-slate-700 whitespace-nowrap align-top">
                        <div class="font-bold text-slate-800 flex items-center flex-wrap gap-1">${field.label}${tagBadgeHtml}</div>
                        ${countBadgeHtml}
                    </td>
                    <td colspan="2" class="py-2.5 px-4 align-top">
                        <div class="space-y-2 ${scrollContainerClass}">
                            ${itemsHtml}
                        </div>
                    </td>
                    <td class="py-3 px-4 text-xs align-top">
                        <span class="inline-flex items-center px-2.5 py-1 rounded text-xs font-medium ${badgeStyle}">
                            <span class="light-dot ${dotStyle} mr-1.5"></span>
                            ${field.desc}
                        </span>
                    </td>
                </tr>
            `;
        } else {
            let tagBadgeHtml = '';
            if (field.layoutTag) {
                if (field.layoutTag === 'adopted-content') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-blue-100 text-blue-800 font-semibold">排版採用</span>`;
                } else if (field.layoutTag === 'adopted-objective') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-purple-100 text-purple-800 font-semibold">排版採用</span>`;
                } else if (field.layoutTag === 'skipped-content') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-slate-100 text-slate-500 font-normal">免排 (採內容)</span>`;
                } else if (field.layoutTag === 'skipped-objective') {
                    tagBadgeHtml = `<span class="ml-1 text-3xs px-1.5 py-0.5 rounded font-mono bg-slate-100 text-slate-500 font-normal">免排 (採目標)</span>`;
                }
            }

            rowsHtml += `
                <tr class="border-b border-slate-100 last:border-none ${rowBg}">
                    <td class="py-3 px-4 text-xs font-semibold text-slate-700 whitespace-nowrap align-top">
                        <div class="font-bold text-slate-800 flex items-center flex-wrap gap-1">${field.label}${tagBadgeHtml}</div>
                    </td>
                    <td class="py-3 px-4 text-xs text-slate-800 font-mono align-top break-words whitespace-pre-line leading-relaxed">
                        ${formatFieldValue(key, field.word, field.status === 'red')}
                    </td>
                    <td class="py-3 px-4 text-xs text-slate-800 font-mono align-top break-words whitespace-pre-line leading-relaxed">
                        ${formatFieldValue(key, field.pdf, field.status === 'red')}
                    </td>
                    <td class="py-3 px-4 text-xs align-top">
                        <span class="inline-flex items-center px-2.5 py-1 rounded text-xs font-medium ${badgeStyle}">
                            <span class="light-dot ${dotStyle} mr-1.5"></span>
                            ${field.desc}
                        </span>
                    </td>
                </tr>
            `;
        }
    }

    // Prepare badges for 中文課名 and 英文課名
    const zhField = item.fields['中文課名'] || item.fields['課程名稱'] || { status: 'gray', desc: '-' };
    const enField = item.fields['英文課名'] || item.fields['英文名稱'] || { status: 'gray', desc: '-' };

    let zhBadgeClass = 'badge-gray';
    let zhDotClass = 'gray';
    let zhBadgeText = '中文課名未比對';
    if (zhField.status === 'green') {
        zhBadgeClass = 'badge-green';
        zhDotClass = 'green';
        zhBadgeText = '中文課名相符';
    } else if (zhField.status === 'yellow') {
        zhBadgeClass = 'badge-yellow';
        zhDotClass = 'yellow';
        zhBadgeText = '中文課名微差';
    } else if (zhField.status === 'red') {
        zhBadgeClass = 'badge-red';
        zhDotClass = 'red';
        zhBadgeText = '中文課名不一致';
    } else {
        zhBadgeText = zhField.desc || '中文課名缺少';
    }

    let enBadgeClass = 'badge-gray';
    let enDotClass = 'gray';
    let enBadgeText = '英文課名未比對';
    if (enField.status === 'green') {
        enBadgeClass = 'badge-green';
        enDotClass = 'green';
        enBadgeText = '英文課名相符';
    } else if (enField.status === 'yellow') {
        enBadgeClass = 'badge-yellow';
        enDotClass = 'yellow';
        enBadgeText = '英文課名微差';
    } else if (enField.status === 'red') {
        enBadgeClass = 'badge-red';
        enDotClass = 'red';
        enBadgeText = (enField.desc && enField.desc.includes('漏排')) ? 'PDF 漏排英文課名' : '英文課名不一致';
    } else {
        enBadgeText = enField.desc || '英文課名缺少';
    }

    const zhNameDisplay = item.nameZh || item.name || '(未提供中文課名)';
    const enNameDisplay = item.nameEn || (item.pdfCourse && item.pdfCourse.course_name_en) || (item.wordCourse && item.wordCourse.course_name_en) || '(無英文課名/未排)';

    let layoutBadgeHtml = '';
    if (item.layoutMode === 'content') {
        layoutBadgeHtml = `
            <div class="flex items-center gap-1.5 pt-0.5">
                <span class="inline-flex items-center px-2 py-0.5 rounded text-2xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                    <svg class="w-3 h-3 mr-1 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h16M4 18h16"></path></svg>
                    排版模式：以「課程內容」排版
                </span>
            </div>
        `;
    } else if (item.layoutMode === 'objective') {
        layoutBadgeHtml = `
            <div class="flex items-center gap-1.5 pt-0.5">
                <span class="inline-flex items-center px-2 py-0.5 rounded text-2xs font-semibold bg-purple-50 text-purple-700 border border-purple-200">
                    <svg class="w-3 h-3 mr-1 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path></svg>
                    排版模式：以「課程目標」排版
                </span>
            </div>
        `;
    } else if (item.layoutMode === 'both') {
        layoutBadgeHtml = `
            <div class="flex items-center gap-1.5 pt-0.5">
                <span class="inline-flex items-center px-2 py-0.5 rounded text-2xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                    <svg class="w-3 h-3 mr-1 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg>
                    排版模式：同時排入「課程目標」與「課程內容」
                </span>
            </div>
        `;
    }

    let omittedLayoutNote = '';
    if (item.layoutMode === 'content') {
        omittedLayoutNote = `<p>• <strong>課程目標 / 內容</strong>：本課程美編採<strong>「課程內容」</strong>排版，Word 原稿之「課程目標」依規則免排入 PDF，故已自動隱藏課程目標對比列（視為正常）。</p>`;
    } else if (item.layoutMode === 'objective') {
        omittedLayoutNote = `<p>• <strong>課程目標 / 內容</strong>：本課程美編採<strong>「課程目標」</strong>排版，Word 原稿之「課程內容」依規則免排入 PDF，故已自動隱藏課程內容對比列（視為正常）。</p>`;
    } else if (item.layoutMode === 'both') {
        omittedLayoutNote = `<p>• <strong>課程目標 / 內容</strong>：本課程美編版面同時排入<strong>「課程目標」</strong>與<strong>「課程內容」</strong>，上方皆已完整對比。</p>`;
    }

    const tableNotesHtml = `
        <div class="px-5 py-3.5 bg-slate-50/80 border-t border-slate-200/80 flex items-start text-xs text-slate-600 gap-2.5 leading-relaxed">
            <svg class="w-4 h-4 text-slate-400 shrink-0 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
            <div class="space-y-1">
                <div><strong class="text-slate-700">排版與校對規則備註：</strong></div>
                <div class="text-3xs text-slate-500 space-y-0.5">
                    ${omittedLayoutNote}
                    <p>• <strong>學會技能</strong>：因宣傳品版面精簡，依規則一律免排入 PDF，故不列入上方對比列（視為正常）。</p>
                    <p>• <strong>後續推薦課程</strong>：美編依版面限制僅排入首門推薦課程，只要第一門相符即判定通過（其餘免排視為正常）。</p>
                </div>
            </div>
        </div>
    `;

    card.innerHTML = `
        <div class="p-5 flex flex-col md:flex-row md:items-start justify-between gap-4 border-b border-slate-100 bg-slate-50/40">
            <div class="flex items-start space-x-3.5 flex-1 min-w-0">
                <span class="mt-0.5 px-2.5 py-1.5 rounded-lg text-xs font-bold font-mono tracking-wide flex-shrink-0 ${
                    item.status === 'red' ? 'bg-red-100 text-red-800' : 'bg-blue-100 text-blue-800'
                }">
                    ${item.code}
                </span>
                <div class="space-y-2 flex-1 min-w-0">
                    <!-- 中文課名與標籤 -->
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="text-2xs font-semibold px-2 py-0.5 bg-slate-200 text-slate-700 rounded">中文課名</span>
                        <h3 class="text-base font-bold text-slate-900 break-words">${zhNameDisplay}</h3>
                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-2xs font-semibold ${zhBadgeClass}">
                            <span class="light-dot ${zhDotClass} mr-1"></span>
                            ${zhBadgeText}
                        </span>
                    </div>
                    <!-- 英文課名與標籤 -->
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="text-2xs font-semibold px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded">英文課名</span>
                        <p class="text-xs font-medium text-slate-600 font-mono break-words">${enNameDisplay}</p>
                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-2xs font-semibold ${enBadgeClass}">
                            <span class="light-dot ${enDotClass} mr-1"></span>
                            ${enBadgeText}
                        </span>
                    </div>
                    <!-- 排版模式指示標籤 -->
                    ${layoutBadgeHtml}
                </div>
            </div>

            <div class="flex items-center space-x-3 flex-shrink-0 self-start md:self-auto">
                ${pdfPage ? `<span class="text-xs text-slate-400 font-medium">${pdfPage}</span>` : ''}
                <span class="inline-flex items-center px-3 py-1.5 rounded-full text-xs font-semibold ${statusBadgeClass}">
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
        ${tableNotesHtml}
    `;

    return card;
}

function formatFieldValue(field, val, isError) {
    if (!val) return '<span class="text-slate-300">-</span>';
    if (isError && (field === '時數' || field === '點數' || field === '費用' || field === '中文課名' || field === '英文課名' || field === '課程名稱' || field === '英文名稱' || field === '課程代碼')) {
        return `<span class="diff-val-error">${val}</span>`;
    }
    return val;
}

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
    report += `說明：依排版規則，本系統以美編 PDF 排版為準進行校對；課程不論以「課程目標」或「課程內容」排版皆視為正確，未排之對應欄位視為正常略過；『學會技能』因版面精簡未排亦視為正常；『後續推薦課程』美編依規則僅排首門。\n`;
    report += `--------------------------------------------------------\n\n`;

    if (redItems.length > 0) {
        report += `🔴 【資料不一致錯誤 (${redItems.length} 門課程)】：\n`;
        redItems.forEach((item, i) => {
            const layoutInfo = item.layoutModeText ? ` [排版模式：以「${item.layoutModeText}」排版]` : '';
            report += `\n${i + 1}. 【${item.code}】${item.name} (頁數：${item.pdfCourse && item.pdfCourse.page ? 'P.' + item.pdfCourse.page : '未知'})${layoutInfo}\n`;
            const reportFieldsOrder = [
                '課程代碼',
                '中文課名',
                '英文課名',
                '時數',
                '費用',
                '點數',
                '教材',
                '課程目標',
                '課程內容',
                '適合對象',
                '先修課程',
                '預備知識',
                '備註事項',
                '後續推薦課程'
            ];
            for (const key of reportFieldsOrder) {
                if (key === '學會技能') continue;
                if (key === '課程內容' && item.layoutMode === 'objective') continue;
                if (key === '課程目標' && item.layoutMode === 'content') continue;
                const f = item.fields[key] || (key === '中文課名' ? item.fields['課程名稱'] : (key === '英文課名' ? item.fields['英文名稱'] : null));
                if (!f || f.status !== 'red') continue;
                if (f.isList && f.details) {
                    report += `   - ${f.label}：${f.desc}\n`;
                    const errDetails = f.details.filter(d => d.status === 'red');
                    errDetails.forEach(d => {
                        const badge = d.hierarchy ? (d.hierarchy.level >= 2 ? (d.hierarchy.badgeText ? `子項 ${d.hierarchy.badgeText}` : `子項`) : (d.hierarchy.badgeText ? `第 ${d.hierarchy.badgeText} 項` : `第 ${d.index} 項`)) : `第 ${d.index} 項`;
                        report += `     • [${badge}] Word「${d.word || '無'}」⇄ PDF「${d.pdf || '漏排'}」➔ ${d.desc}\n`;
                    });
                } else {
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
        'Word中文課名',
        'PDF中文課名',
        '中文課名比對',
        'Word英文課名',
        'PDF英文課名',
        '英文課名比對',
        '比對狀態',
        '排版模式',
        'Word時數',
        'PDF時數',
        'Word費用',
        'PDF費用',
        'Word點數',
        'PDF點數',
        '教材比對',
        '課程目標比對',
        '課程內容比對',
        '適合對象比對',
        '先修課程比對',
        '預備知識比對',
        '備註事項比對',
        '後續推薦課程比對',
        'PDF頁碼',
        '差異說明'
    ];
    const rows = [headers];

    for (const c of state.comparisons) {
        const fields = c.fields;
        const errDescs = [];
        for (const k of Object.keys(fields)) {
            if (fields[k].status === 'red') {
                if (k === '課程名稱' && fields['中文課名']) continue;
                if (k === '英文名稱' && fields['英文課名']) continue;
                if (k === '學會技能') continue;
                if (k === '課程內容' && c.layoutMode === 'objective') continue;
                if (k === '課程目標' && c.layoutMode === 'content') continue;
                errDescs.push(`${fields[k].label}:${fields[k].desc}`);
            }
        }

        const zhF = fields['中文課名'] || fields['課程名稱'] || {};
        const enF = fields['英文課名'] || fields['英文名稱'] || {};

        rows.push([
            c.code,
            `"${(zhF.word || '').replace(/"/g, '""')}"`,
            `"${(zhF.pdf || '').replace(/"/g, '""')}"`,
            zhF.desc || '',
            `"${(enF.word || '').replace(/"/g, '""')}"`,
            `"${(enF.pdf || '').replace(/"/g, '""')}"`,
            enF.desc || '',
            c.status === 'green' ? '相符' : c.status === 'red' ? '錯誤' : c.status === 'yellow' ? '提醒' : '遺漏',
            c.layoutModeText || '',
            fields['時數'] ? fields['時數'].word : '',
            fields['時數'] ? fields['時數'].pdf : '',
            fields['費用'] ? fields['費用'].word : '',
            fields['費用'] ? fields['費用'].pdf : '',
            fields['點數'] ? fields['點數'].word : '',
            fields['點數'] ? fields['點數'].pdf : '',
            fields['教材'] ? fields['教材'].desc : '',
            fields['課程目標'] ? fields['課程目標'].desc : '',
            fields['課程內容'] ? fields['課程內容'].desc : '',
            fields['適合對象'] ? fields['適合對象'].desc : '',
            fields['先修課程'] ? fields['先修課程'].desc : '',
            fields['預備知識'] ? fields['預備知識'].desc : '',
            fields['備註事項'] ? fields['備註事項'].desc : '',
            fields['後續推薦課程'] ? fields['後續推薦課程'].desc : '',
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
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/[\s\r\n\t\u3000]+/g, '')
        .replace(/[，,。.:：;；()（）「」『』"'\-／/＋+\\®™©&]/g, '')
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
