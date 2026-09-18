/**
 * 課程資料校稿小工具 (Course Proofreader Web App)
 * 核心比對與解析邏輯
 */

// Configure PDF.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.js';
}

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

// DOM Elements
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

    // KPI
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
    // Word Drag & Drop
    setupDropzone(elements.wordDropzone, elements.wordInput, (file) => handleWordFile(file));
    // PDF Drag & Drop
    setupDropzone(elements.pdfDropzone, elements.pdfInput, (file) => handlePdfFile(file));

    // Sample Buttons
    elements.btnSampleBlockchain.addEventListener('click', () => loadSample('blockchain'));
    elements.btnSampleJianzhen.addEventListener('click', () => loadSample('jianzhen'));

    // Start Compare Button
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
        btn.addEventListener('click', (e) => {
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
// File Loaders & Handlers
// ==========================================
async function handleWordFile(file) {
    if (!file.name.endsWith('.docx')) {
        showToast('請上傳 .docx 格式的 Word 檔案！', true);
        return;
    }
    state.wordFile = file;
    state.wordBuffer = await file.arrayBuffer();
    
    // Update UI
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

    // Update UI
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
    let wordPath, pdfPath, wordName, pdfName;
    if (type === 'blockchain') {
        wordPath = '恆逸_區塊鏈_2027年1-6月課程.docx';
        pdfPath = '區塊鏈.pdf';
        wordName = '恆逸_區塊鏈_2027年1-6月課程.docx';
        pdfName = '區塊鏈.pdf';
    } else {
        wordPath = '恆逸_鑒真數位_2027年1-6月課程v1_1.docx';
        pdfPath = '鑒真數位.pdf';
        wordName = '恆逸_鑒真數位_2027年1-6月課程v1_1.docx';
        pdfName = '鑒真數位.pdf';
    }

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

        const wordFile = new File([wordBlob], wordName, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const pdfFile = new File([pdfBlob], pdfName, { type: 'application/pdf' });

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
 * Parses Word (.docx) file by unzipping word/document.xml and extracting table data.
 */
async function parseDocx(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const xmlFile = zip.file('word/document.xml');
    if (!xmlFile) throw new Error('無效的 Word 檔案 (找不到 document.xml)');

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
                const textNodes = tc.getElementsByTagName('w:t');
                let cellText = '';
                for (let k = 0; k < textNodes.length; k++) {
                    cellText += textNodes[k].textContent;
                }
                cells.push(cellText.trim());
            }
            rows.push(cells);
        }

        const fullTableText = rows.map(r => r.join(' ')).join('\n');
        // Check if this table is a course info table
        if (!fullTableText.includes('時數') && !fullTableText.includes('費用') && !fullTableText.includes('點數')) {
            continue;
        }

        const course = {
            code: '',
            name: '',
            nameEn: '',
            hours: '',
            price: '',
            points: '',
            material: '',
            outline: '',
            outlineItems: [],
            target: '',
            notes: '',
            source: 'Word'
        };

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rowText = row.join(' ').replace(/\s+/g, ' ').trim();

            if (i === 0) {
                // Course Code & Name
                if (row.length >= 2 && row[0].length <= 10 && row[0].trim().length > 0) {
                    course.code = row[0].trim();
                    course.name = row.slice(1).join(' ').replace(/^[：:\s|]+/, '').trim();
                } else {
                    const parts = rowText.split(/[\s|：:]+/);
                    if (parts.length >= 2 && parts[0].length <= 10) {
                        course.code = parts[0].trim();
                        course.name = rowText.replace(parts[0], '').replace(/^[|：:\s]+/, '').trim();
                    } else {
                        course.name = rowText;
                    }
                }
            } else if (i === 1 && !rowText.includes('時數') && !rowText.includes('費用')) {
                course.nameEn = rowText.replace(/^[|：:\s]+/, '').trim();
            }

            // Metadata row
            if (rowText.includes('時數') || rowText.includes('費用') || rowText.includes('點數')) {
                const hoursMatch = rowText.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
                if (hoursMatch) course.hours = hoursMatch[1];

                const priceMatch = rowText.match(/費用[：:\s]*([0-9,]+)\s*元?/);
                if (priceMatch) course.price = priceMatch[1].replace(/,/g, '');

                const pointsMatch = rowText.match(/點數[：:\s]*([0-9.]+)\s*點?/);
                if (pointsMatch) course.points = pointsMatch[1];

                const matMatch = rowText.match(/教材[：:\s]*([^|｜\n]+)/);
                if (matMatch) course.material = matMatch[1].trim();
            }

            // Content / Outline
            if (row[0] && row[0].includes('課程內容')) {
                course.outline = row.slice(1).join(' ').trim();
            } else if (rowText.includes('課程內容') && !course.outline) {
                course.outline = rowText.replace(/^.*課程內容[：:\s]*/, '').trim();
            }

            // Other fields
            if (row[0] && row[0].includes('適合對象')) course.target = row.slice(1).join(' ').trim();
            if (row[0] && row[0].includes('備註事項')) course.notes = row.slice(1).join(' ').trim();
        }

        // Clean code if prefixed to name
        if (!course.code && course.name) {
            const m = course.name.match(/^([A-Za-z0-9_-]{2,10})\s+(.*)$/);
            if (m) {
                course.code = m[1];
                course.name = m[2];
            }
        }

        if (course.outline) {
            course.outlineItems = splitOutlineItems(course.outline);
        }

        if (course.code || course.name) {
            courses.push(course);
        }
    }

    return courses;
}

/**
 * Parses PDF file page by page, groups visual lines, and identifies course sections.
 */
async function parsePdf(buffer) {
    const doc = await pdfjsLib.getDocument({
        data: new Uint8Array(buffer),
        useSystemFonts: true,
        disableFontFace: true
    }).promise;

    const allLinesWithPos = [];

    for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const textContent = await page.getTextContent();

        // Sort items by Y descending (top to bottom), then X ascending (left to right)
        const sortedItems = [...textContent.items].sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5];
            if (Math.abs(yDiff) > 3) return yDiff;
            return a.transform[4] - b.transform[4];
        });

        let curY = null;
        let curLine = [];
        for (const item of sortedItems) {
            const y = Math.round(item.transform[5]);
            if (curY === null || Math.abs(curY - y) > 4) {
                if (curLine.length > 0) {
                    allLinesWithPos.push({ page: p, y: curY, text: curLine.join('').trim() });
                }
                curY = y;
                curLine = [item.str];
            } else {
                curLine.push(item.str);
            }
        }
        if (curLine.length > 0) {
            allLinesWithPos.push({ page: p, y: curY, text: curLine.join('').trim() });
        }
    }

    // Identify course metadata lines
    const courses = [];
    const metaIndices = [];
    for (let i = 0; i < allLinesWithPos.length; i++) {
        const line = allLinesWithPos[i].text.replace(/\s+/g, ' ');
        if (line.includes('時數') && (line.includes('費用') || line.includes('點數') || line.includes('教材'))) {
            metaIndices.push(i);
        }
    }

    for (let k = 0; k < metaIndices.length; k++) {
        const metaIdx = metaIndices[k];
        const nextMetaIdx = k + 1 < metaIndices.length ? metaIndices[k + 1] : allLinesWithPos.length;
        const prevMetaIdx = k > 0 ? metaIndices[k - 1] : 0;
        const metaLine = allLinesWithPos[metaIdx].text.replace(/\s+/g, ' ');

        const course = {
            code: '',
            name: '',
            nameEn: '',
            hours: '',
            price: '',
            points: '',
            material: '',
            outline: '',
            outlineItems: [],
            target: '',
            notes: '',
            source: 'PDF',
            page: allLinesWithPos[metaIdx].page
        };

        // Extract metadata from metaLine
        const hoursMatch = metaLine.match(/時數[：:\s]*([0-9.]+)\s*小時?/);
        if (hoursMatch) course.hours = hoursMatch[1];

        const priceMatch = metaLine.match(/費用[：:\s]*([0-9,]+)\s*元?/);
        if (priceMatch) course.price = priceMatch[1].replace(/,/g, '');

        const pointsMatch = metaLine.match(/點數[：:\s]*([0-9.]+)\s*點?/);
        if (pointsMatch) course.points = pointsMatch[1];

        const matMatch = metaLine.match(/教材[：:\s]*([^｜|\n]+)/);
        if (matMatch) course.material = matMatch[1].trim();

        // Scan lines BEFORE metaLine for Course Code, Chinese Name, and English Name
        const titleLines = [];
        const scanStart = Math.max(prevMetaIdx === 0 ? 0 : prevMetaIdx + 1, metaIdx - 5);
        for (let j = metaIdx - 1; j >= scanStart; j--) {
            const t = allLinesWithPos[j].text.trim();
            if (t.includes('課程簡介') || t.includes('各地開課時間') || t.match(/^\|\s*.*\s*\|$/)) continue;
            if (t.includes('後續推薦課程') || t.includes('備註事項')) break;
            titleLines.unshift(t);
        }

        // Title and Code parsing
        for (const line of titleLines) {
            const cleanLine = line.replace(/\s+/g, ' ').trim();
            const codeMatch = cleanLine.match(/\b([A-Za-z0-9_-]{2,10})\b/);
            if (codeMatch && !course.code && !['APP', 'DApp', 'Web3', 'EVM', 'Full', 'Stack', 'Course', 'Exam'].includes(codeMatch[1])) {
                course.code = codeMatch[1];
            }

            if (/^[A-Za-z0-9\s,&.:()'-]+$/.test(cleanLine) && cleanLine.length > 5) {
                course.nameEn = cleanLine.replace(course.code, '').trim();
            } else if (/[\u4e00-\u9fa5]/.test(cleanLine)) {
                if (!course.name) {
                    course.name = cleanLine.replace(course.code, '').trim();
                } else {
                    course.name += ' ' + cleanLine.replace(course.code, '').trim();
                }
            }
        }

        // Scan lines AFTER metaLine for syllabus/outline
        let currentSection = '';
        const outlineParts = [];
        for (let j = metaIdx + 1; j < nextMetaIdx; j++) {
            const line = allLinesWithPos[j].text.trim();
            if (line.includes('課程簡介') || line.includes('各地開課時間')) continue;

            if (line.includes('課程內容')) {
                currentSection = 'outline';
                const rest = line.replace(/.*課程內容[:：\s]*/, '').trim();
                if (rest) outlineParts.push(rest);
            } else if (line.includes('適合對象')) {
                currentSection = 'target';
            } else if (line.includes('預備知識')) {
                currentSection = 'prereq';
            } else if (line.includes('備註事項')) {
                currentSection = 'notes';
            } else if (line.includes('後續推薦課程') || line.includes('先修課程')) {
                currentSection = 'other';
            } else if (currentSection === 'outline') {
                outlineParts.push(line);
            }
        }
        course.outline = outlineParts.join(' ').trim();
        if (course.outline) {
            course.outlineItems = splitOutlineItems(course.outline);
        }

        // Clean up course name
        if (course.name && course.code) {
            course.name = course.name.replace(new RegExp(`^${course.code}[：:\\s]*`), '').trim();
        }

        courses.push(course);
    }

    return courses;
}

/**
 * Splits course outline text into distinct numbered or bulleted items.
 */
function splitOutlineItems(text) {
    if (!text) return [];
    // Split on numbers like "1.", "2." or bullets "•"
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

        elements.loadingText.textContent = '正在進行逐欄比對與標示紅綠燈號...';
        state.comparisons = compareCourseData(state.wordCourses, state.pdfCourses);

        updateKPIDashboard();
        renderCourseCards();

        elements.loadingIndicator.classList.add('hidden');
        elements.resultsSection.classList.remove('hidden');

        // Scroll smoothly to results
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

    // Map each Word course to a PDF course
    for (const wCourse of wordCourses) {
        let matchedPdf = null;
        let matchedIdx = -1;

        // 1. Exact match by course code
        if (wCourse.code) {
            const idx = pdfCourses.findIndex((p, i) => !matchedPdfIndices.has(i) && p.code && p.code.toLowerCase() === wCourse.code.toLowerCase());
            if (idx !== -1) {
                matchedPdf = pdfCourses[idx];
                matchedIdx = idx;
            }
        }

        // 2. Match by course title similarity if code didn't match
        if (!matchedPdf && wCourse.name) {
            const wClean = normalizeText(wCourse.name);
            for (let i = 0; i < pdfCourses.length; i++) {
                if (matchedPdfIndices.has(i)) continue;
                const pClean = normalizeText(pdfCourses[i].name);
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
                status: 'gray', // Missing
                statusText: 'PDF 排版漏排此課程',
                code: wCourse.code || '未標註',
                name: wCourse.name,
                wordCourse: wCourse,
                pdfCourse: null,
                fields: {
                    name: { status: 'gray', label: '課程名稱', word: wCourse.name, pdf: '(未找到)', desc: 'PDF 缺少此課程' },
                    hours: { status: 'gray', label: '課程時數', word: `${wCourse.hours} 小時`, pdf: '-', desc: '缺少' },
                    points: { status: 'gray', label: '課程點數', word: `${wCourse.points} 點`, pdf: '-', desc: '缺少' },
                    price: { status: 'gray', label: '課程費用', word: `${Number(wCourse.price || 0).toLocaleString()} 元`, pdf: '-', desc: '缺少' },
                    material: { status: 'gray', label: '課程教材', word: wCourse.material || '-', pdf: '-', desc: '缺少' },
                    outline: { status: 'gray', label: '課程大綱', word: `${wCourse.outlineItems.length} 項內容`, pdf: '-', desc: '缺少' }
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
                code: pCourse.code || '未標註',
                name: pCourse.name,
                wordCourse: null,
                pdfCourse: pCourse,
                fields: {
                    name: { status: 'gray', label: '課程名稱', word: '(未找到)', pdf: pCourse.name, desc: 'Word 原稿未列出此課' },
                    hours: { status: 'gray', label: '課程時數', word: '-', pdf: `${pCourse.hours} 小時`, desc: '原稿無' },
                    points: { status: 'gray', label: '課程點數', word: '-', pdf: `${pCourse.points} 點`, desc: '原稿無' },
                    price: { status: 'gray', label: '課程費用', word: '-', pdf: `${Number(pCourse.price || 0).toLocaleString()} 元`, desc: '原稿無' },
                    material: { status: 'gray', label: '課程教材', word: '-', pdf: pCourse.material || '-', desc: '原稿無' },
                    outline: { status: 'gray', label: '課程大綱', word: '-', pdf: `${pCourse.outlineItems.length} 項內容`, desc: '原稿無' }
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

    // 1. Course Code
    const codeMatch = (w.code || '').trim().toLowerCase() === (p.code || '').trim().toLowerCase();
    fields.code = {
        label: '課程代碼',
        word: w.code || '-',
        pdf: p.code || '-',
        status: codeMatch ? 'green' : 'red',
        desc: codeMatch ? '代碼正確' : '課程代碼不相符！'
    };
    if (!codeMatch) hasRed = true;

    // 2. Course Name (Chinese)
    const wNameNorm = normalizeText(w.name);
    const pNameNorm = normalizeText(p.name);
    if (wNameNorm === pNameNorm) {
        fields.name = { label: '課程名稱', word: w.name, pdf: p.name, status: 'green', desc: '名稱完全相符' };
    } else if (calculateSimilarity(wNameNorm, pNameNorm) > 0.8) {
        fields.name = { label: '課程名稱', word: w.name, pdf: p.name, status: 'yellow', desc: '微小文字或空白差異' };
        hasYellow = true;
    } else {
        fields.name = { label: '課程名稱', word: w.name, pdf: p.name, status: 'red', desc: '名稱明顯不同或有錯字！' };
        hasRed = true;
    }

    // 3. Hours (時數) - CRITICAL FIELD
    const wHours = parseFloat(w.hours) || 0;
    const pHours = parseFloat(p.hours) || 0;
    if (wHours > 0 && pHours > 0 && wHours === pHours) {
        fields.hours = { label: '課程時數', word: `${w.hours} 小時`, pdf: `${p.hours} 小時`, status: 'green', desc: '時數相符' };
    } else {
        fields.hours = {
            label: '課程時數',
            word: `${w.hours || 0} 小時`,
            pdf: `${p.hours || 0} 小時`,
            status: 'red',
            desc: `時數不一致！Word 寫 ${w.hours} 小時，但 PDF 寫 ${p.hours} 小時`
        };
        hasRed = true;
    }

    // 4. Points (點數) - CRITICAL FIELD
    const wPoints = parseFloat(w.points) || 0;
    const pPoints = parseFloat(p.points) || 0;
    if (wPoints > 0 && pPoints > 0 && wPoints === pPoints) {
        fields.points = { label: '課程點數', word: `${w.points} 點`, pdf: `${p.points} 點`, status: 'green', desc: '點數相符' };
    } else if (w.points || p.points) {
        fields.points = {
            label: '課程點數',
            word: `${w.points || 0} 點`,
            pdf: `${p.points || 0} 點`,
            status: 'red',
            desc: `點數不一致！Word 為 ${w.points} 點，但 PDF 為 ${p.points} 點`
        };
        hasRed = true;
    } else {
        fields.points = { label: '課程點數', word: '無', pdf: '無', status: 'green', desc: '雙方皆無點數' };
    }

    // 5. Price (費用)
    const wPrice = parseInt(w.price, 10) || 0;
    const pPrice = parseInt(p.price, 10) || 0;
    if (wPrice > 0 && pPrice > 0 && wPrice === pPrice) {
        fields.price = { label: '課程費用', word: `${wPrice.toLocaleString()} 元`, pdf: `${pPrice.toLocaleString()} 元`, status: 'green', desc: '費用相符' };
    } else if (wPrice !== pPrice) {
        fields.price = {
            label: '課程費用',
            word: `${wPrice.toLocaleString()} 元`,
            pdf: `${pPrice.toLocaleString()} 元`,
            status: 'red',
            desc: `費用不符！Word 為 ${wPrice.toLocaleString()} 元，PDF 為 ${pPrice.toLocaleString()} 元`
        };
        hasRed = true;
    } else {
        fields.price = { label: '課程費用', word: '-', pdf: '-', status: 'green', desc: '無費用資訊' };
    }

    // 6. Material (教材)
    const wMat = normalizeText(w.material || '');
    const pMat = normalizeText(p.material || '');
    if (wMat === pMat) {
        fields.material = { label: '課程教材', word: w.material || '-', pdf: p.material || '-', status: 'green', desc: '教材相符' };
    } else if (wMat.includes(pMat) || pMat.includes(wMat)) {
        fields.material = { label: '課程教材', word: w.material || '-', pdf: p.material || '-', status: 'yellow', desc: '教材文字有修訂或簡寫' };
        hasYellow = true;
    } else {
        fields.material = { label: '課程教材', word: w.material || '-', pdf: p.material || '-', status: 'red', desc: '教材資料不一致' };
        hasRed = true;
    }

    // 7. Outline (大綱)
    const outlineStatus = compareOutlines(w.outlineItems, p.outlineItems);
    fields.outline = {
        label: '課程大綱',
        word: `${w.outlineItems.length} 個大綱項目`,
        pdf: `${p.outlineItems.length} 個大綱項目`,
        status: outlineStatus.status,
        desc: outlineStatus.desc,
        details: outlineStatus.details
    };
    if (outlineStatus.status === 'red') hasRed = true;
    if (outlineStatus.status === 'yellow') hasYellow = true;

    // Determine Overall Status
    let overallStatus = 'green';
    let overallText = '檢查通過，所有欄位完全正確';
    if (hasRed) {
        overallStatus = 'red';
        overallText = '發現重大錯誤（時數、點數或名稱不符）';
    } else if (hasYellow) {
        overallStatus = 'yellow';
        overallText = '部分文字或大綱排版有微調提醒';
    }

    return {
        status: overallStatus,
        statusText: overallText,
        code: w.code || p.code || '未註記',
        name: w.name || p.name,
        wordCourse: w,
        pdfCourse: p,
        fields
    };
}

function compareOutlines(wItems, pItems) {
    if (!wItems.length && !pItems.length) {
        return { status: 'green', desc: '雙方皆無大綱', details: [] };
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
        return { status: 'green', desc: `大綱完全相符 (${wItems.length} 項)`, details };
    } else if (Math.abs(wItems.length - pItems.length) <= 1 && matchCount >= Math.min(wItems.length, pItems.length) * 0.8) {
        return { status: 'yellow', desc: `大綱有微調或格式不同 (Word: ${wItems.length}項, PDF: ${pItems.length}項)`, details };
    } else {
        return { status: 'red', desc: `大綱項目數量或內容有缺漏 (Word: ${wItems.length}項, PDF: ${pItems.length}項)`, details };
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
        // Status filter
        if (filter !== 'all' && c.status !== filter) return false;
        // Search filter
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

    // Generate table rows for fields
    const fieldsToRender = ['hours', 'points', 'name', 'price', 'material', 'outline'];
    let rowsHtml = '';

    for (const key of fieldsToRender) {
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

        // Outline special button
        let outlineExpandBtn = '';
        if (key === 'outline' && field.details && field.details.length > 0) {
            outlineExpandBtn = `
                <button onclick="toggleOutlineDetails(${idx})" class="mt-1 text-xs text-blue-600 hover:text-blue-800 font-medium inline-flex items-center cursor-pointer">
                    <span>展開大綱逐條對照</span>
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
                    <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${badgeStyle}">
                        <span class="light-dot ${dotStyle} mr-1.5"></span>
                        ${field.desc}
                    </span>
                    ${outlineExpandBtn}
                </td>
            </tr>
        `;
    }

    // Outline expandable block
    let outlineAccordionHtml = '';
    if (item.fields.outline && item.fields.outline.details && item.fields.outline.details.length > 0) {
        let itemsHtml = '';
        item.fields.outline.details.forEach(d => {
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
                        ${d.status === 'red' ? `<span class="ml-2 text-2xs px-1.5 py-0.5 bg-red-100 text-red-700 rounded font-bold">項目錯誤或遺漏</span>` : ''}
                    </div>
                </div>
            `;
        });

        outlineAccordionHtml = `
            <div id="outlineDetails-${idx}" class="hidden p-4 bg-slate-50 border-t border-slate-200">
                <h5 class="text-xs font-bold text-slate-700 mb-2 flex items-center">
                    <svg class="w-4 h-4 mr-1 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path></svg>
                    大綱逐條對照清單 (Word vs PDF)
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
                    <p class="text-xs text-slate-400">${item.wordCourse && item.wordCourse.nameEn ? item.wordCourse.nameEn : (item.pdfCourse && item.pdfCourse.nameEn ? item.pdfCourse.nameEn : '')}</p>
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
                        <th class="py-2.5 px-4 w-40">校對結果</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>

        ${outlineAccordionHtml}
    `;

    return card;
}

function formatFieldValue(field, val, isError) {
    if (!val) return '<span class="text-slate-300">-</span>';
    if (isError && (field === 'hours' || field === 'points' || field === 'price')) {
        return `<span class="diff-val-error">${val}</span>`;
    }
    return val;
}

window.toggleOutlineDetails = function(idx) {
    const el = document.getElementById(`outlineDetails-${idx}`);
    if (el) {
        el.classList.toggle('hidden');
    }
};

// ==========================================
// Reporting & Export Helpers
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
    report += `--------------------------------------------------------\n\n`;

    if (redItems.length > 0) {
        report += `🔴 【資料不一致錯誤 (${redItems.length} 門課程)】：\n`;
        redItems.forEach((item, i) => {
            report += `\n${i + 1}. 【${item.code}】${item.name} (頁數：${item.pdfCourse && item.pdfCourse.page ? 'P.' + item.pdfCourse.page : '未知'})\n`;
            for (const key of Object.keys(item.fields)) {
                const f = item.fields[key];
                if (f.status === 'red') {
                    report += `   - ${f.label}：Word原稿寫「${f.word}」，但PDF排版為「${f.pdf}」 ➔ ${f.desc}\n`;
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
        // Fallback
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

    const headers = ['課程代碼', '課程名稱', '比對狀態', 'Word時數', 'PDF時數', 'Word點數', 'PDF點數', 'Word費用', 'PDF費用', '教材比對', 'PDF頁碼', '差異說明'];
    const rows = [headers];

    for (const c of state.comparisons) {
        const fields = c.fields;
        const errDescs = [];
        for (const k of Object.keys(fields)) {
            if (fields[k].status === 'red') errDescs.push(fields[k].desc);
        }

        rows.push([
            c.code,
            `"${(c.name || '').replace(/"/g, '""')}"`,
            c.status === 'green' ? '相符' : c.status === 'red' ? '錯誤' : c.status === 'yellow' ? '提醒' : '遺漏',
            fields.hours ? fields.hours.word : '',
            fields.hours ? fields.hours.pdf : '',
            fields.points ? fields.points.word : '',
            fields.points ? fields.points.pdf : '',
            fields.price ? fields.price.word : '',
            fields.price ? fields.price.pdf : '',
            fields.material ? fields.material.desc : '',
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
// Utility Functions
// ==========================================

function normalizeText(str) {
    if (!str) return '';
    return str
        .replace(/[\s\r\n\t\u3000]+/g, '') // remove whitespace and full-width space
        .replace(/[，,。.:：;；()（）「」『』"'-]/g, '') // remove punctuation
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
