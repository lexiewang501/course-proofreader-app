const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const PORT = 8080;
const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.pdf': 'application/pdf',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.ico': 'image/x-icon',
    '.svg': 'image/svg+xml'
};

const server = http.createServer((req, res) => {
    // Enable CORS just in case
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');

    let reqUrl = decodeURIComponent(req.url.split('?')[0]);
    if (reqUrl === '/') reqUrl = '/index.html';

    const filePath = path.join(__dirname, reqUrl);

    // Prevent directory traversal
    if (!filePath.startsWith(__dirname)) {
        res.writeHead(403);
        res.end('Forbidden');
        return;
    }

    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('404 Not Found: ' + reqUrl);
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';

        res.writeHead(200, {
            'Content-Type': contentType,
            'Content-Length': stats.size,
            'Cache-Control': 'no-cache'
        });

        const stream = fs.createReadStream(filePath);
        stream.pipe(res);
    });
});

server.listen(PORT, () => {
    console.log(`====================================================`);
    console.log(`  🚀 課程資料校稿小工具 伺服器已啟動！`);
    console.log(`  網址：http://localhost:${PORT}`);
    console.log(`====================================================`);

    // Open browser automatically
    const startCmd = process.platform === 'win32' ? `start http://localhost:${PORT}` : `open http://localhost:${PORT}`;
    exec(startCmd, (err) => {
        if (err) console.log('請手動在瀏覽器開啟：http://localhost:' + PORT);
    });
});
