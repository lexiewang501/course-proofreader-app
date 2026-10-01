const fs = require('fs');
const path = require('path');
const workspaceDir = 'c:/Users/lexiewang/OneDrive - 精誠資訊股份有限公司/桌面/校稿';

class Node { constructor(t) { this.tagName = t; this.children = []; this.textContent = ''; this.attributes = {}; } getElementsByTagName(n) { const r = []; function rec(c) { for (const ch of c.children) { if (ch.tagName === n) r.push(ch); rec(ch); } } rec(this); return r; } getAttribute(n) { return this.attributes[n] || null; } }
function parseXmlToDoc(xml) { const root = new Node('root'); const stack = [root]; const tagRegex = /<(\/)?([a-zA-Z0-9:_]+)([^>]*?)(\/)?>|([^<]+)/g; let match; while ((match = tagRegex.exec(xml)) !== null) { const [full, isClosing, tagName, attrs, isSelfClosing, text] = match; if (text) { if (stack.length > 1) { stack[stack.length - 1].textContent += text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'); } } else if (isClosing) { if (stack.length > 1 && stack[stack.length - 1].tagName === tagName) stack.pop(); } else { const node = new Node(tagName); if (attrs) { const attrRegex = /([a-zA-Z0-9:_]+)="([^"]*)"/g; let am; while ((am = attrRegex.exec(attrs)) !== null) { node.attributes[am[1]] = am[2]; } } if (isSelfClosing) { stack[stack.length - 1].children.push(node); } else { stack[stack.length - 1].children.push(node); stack.push(node); } } } return root; }
global.DOMParser = function() { return { parseFromString: (str) => parseXmlToDoc(str) }; };
global.document = { addEventListener: () => {}, getElementById: () => ({ addEventListener: () => {}, classList: { add: () => {}, remove: () => {} }, innerHTML: '', textContent: '', appendChild: () => {}, scrollIntoView: () => {} }), querySelectorAll: () => [], createElement: (tag) => ({ tagName: tag, className: '', innerHTML: '', appendChild: () => {}, classList: { add: () => {}, remove: () => {} } }) };
global.window = {};
global.pdfjsLib = require(path.join(workspaceDir, 'lib/pdf.min.js'));
global.pdfjsLib.GlobalWorkerOptions.workerSrc = './lib/pdf.worker.js';
global.JSZip = require(path.join(workspaceDir, 'lib/jszip.min.js'));

const appCode = fs.readFileSync(path.join(workspaceDir, 'app.js'), 'utf8');
eval(appCode);

async function run() {
    const exampleDir = path.join(workspaceDir, '課程範例');

    // 10 exact pairs
    const pairs = [
        {
            name: '資安課程',
            docx: '恆逸_資安課程_2027年1-6月課程.docx',
            pdf: '資安_二校.pdf'
        },
        {
            name: 'Cisco 思科',
            docx: '恆逸_Cisco_2027年01-06月課程.docx',
            pdf: 'Cisco_二校.pdf'
        },
        {
            name: 'RedHat',
            docx: '恆逸_RedHat_2027年01-06月課程.docx',
            pdf: 'RedHat_二校.pdf'
        },
        {
            name: 'Python',
            docx: '恆逸_Python_2027年1-6月課程.docx',
            pdf: 'Python_二校.pdf'
        },
        {
            name: 'MSSE',
            docx: '恆逸_MS_2027年01-06月MSSE課程.docx',
            pdf: 'MSSE_二校.pdf'
        },
        {
            name: 'Power Platform',
            docx: '恆逸_MS_2027年01-06月Power Platform課程.docx',
            pdf: 'Power Platform_二校.pdf'
        },
        {
            name: 'OffSec',
            docx: '恆逸_OffSec_2027年1-6月課程.docx',
            pdf: 'OffSec_二校.pdf'
        },
        {
            name: 'OpenSource',
            docx: '恆逸_OpenSource_2027年01-06月課程.docx',
            pdf: 'OpenSource_二校.pdf'
        },
        {
            name: 'Oracle DB',
            docx: '恆逸_Oracle_2027年1-6月Oracle DB課程.docx',
            pdf: 'Oracle DB_二校.pdf'
        },
        {
            name: 'Oracle Java',
            docx: '恆逸_Oracle_2027年1-6月Oracle Java課程.docx',
            pdf: 'Oracle Java_二校.pdf'
        }
    ];

    console.log(`Starting 10-Catalog Verification (Testing ${pairs.length} catalogs)...`);
    let totalRed = 0;
    let totalYellow = 0;
    let totalGreen = 0;
    let totalCoursesTested = 0;

    for (let pIdx = 0; pIdx < pairs.length; pIdx++) {
        const pair = pairs[pIdx];
        console.log(`\n========================================`);
        console.log(`[Catalog ${pIdx + 1}/${pairs.length}]: ${pair.name}`);
        console.log(`  Word: ${pair.docx}`);
        console.log(`  PDF:  ${pair.pdf}`);

        const docxBuf = fs.readFileSync(path.join(exampleDir, pair.docx));
        const pdfBuf = fs.readFileSync(path.join(exampleDir, pair.pdf));

        const docxCourses = await parseDocx(docxBuf.buffer);
        const pdfCourses = await parsePdf(pdfBuf.buffer);

        console.log(`  Parsed: ${docxCourses.length} docx courses, ${pdfCourses.length} pdf courses`);

        const comps = compareCourseData(docxCourses, pdfCourses);
        console.log(`  Matched Comparisons: ${comps.length} courses`);

        let catRed = 0, catYellow = 0, catGreen = 0;
        for (const comp of comps) {
            totalCoursesTested++;
            for (const f of Object.values(comp.fields)) {
                if (f.status === 'red') {
                    catRed++;
                    totalRed++;
                    console.log(`    [RED] Course ${comp.code || comp.name} -> Field: ${f.label} (${f.desc})`);
                } else if (f.status === 'yellow') {
                    catYellow++;
                    totalYellow++;
                } else if (f.status === 'green') {
                    catGreen++;
                    totalGreen++;
                }
            }
        }
        console.log(`  Result for ${pair.name}: Red=${catRed}, Yellow=${catYellow}, Green=${catGreen}`);
    }

    console.log(`\n========================================`);
    console.log(`VERIFICATION SUMMARY:`);
    console.log(`Total Catalogs Tested: ${pairs.length}`);
    console.log(`Total Courses Tested:  ${totalCoursesTested}`);
    console.log(`Total Red Fields:      ${totalRed}`);
    console.log(`Total Yellow Fields:   ${totalYellow}`);
    console.log(`Total Green Fields:    ${totalGreen}`);
    console.log(`========================================`);

    if (totalRed === 0) {
        console.log(`SUCCESS: 0 RED ISSUES ACROSS ALL 10 CATALOGS!`);
    } else {
        console.error(`FAILED: FOUND ${totalRed} RED ISSUES!`);
        process.exit(1);
    }
}
run().catch(console.error);
