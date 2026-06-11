const puppeteer = require('puppeteer');
const Tesseract = require('tesseract.js');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const os = require('os');
const ExcelJS = require('exceljs');
const { google } = require('googleapis');

// --- CONFIGURATION ---
const config = {
    gpsUser: process.env.GPS_USER || '',
    gpsPass: process.env.GPS_PASSWORD || '',
    emailFrom: process.env.EMAIL_FROM || '',
    emailPass: process.env.EMAIL_PASSWORD || '',
    emailTo: process.env.EMAIL_TO || '',
    downloadTimeout: 120000, // 2 minutes
    googleSheetId: '1-D9J36mYKE7vldyowMPl-xXLlVU8orBpUYS5d1OjSrE',
    googleSheetTabName: 'YAWNING_BEHAVIOR_DASHBOARD' // เปลี่ยนชื่อให้ตรงกับชื่อแท็บ (Sheet) ของคุณที่ด้านล่างสุดของเว็บ
};

const downloadPath = path.resolve(__dirname, 'downloads');
const defaultDownloadPath = path.join(os.homedir(), 'Downloads');

if (!fs.existsSync(downloadPath)) {
    fs.mkdirSync(downloadPath);
}

// ฟังก์ชันรอไฟล์โหลด
async function waitForFileToDownload(timeout) {
    return new Promise((resolve, reject) => {
        let timer;
        const checkInterval = 2000;
        let timePassed = 0;

        console.log(`   Waiting for file in:`);
        console.log(`      1. ${downloadPath}`);
        console.log(`      2. ${defaultDownloadPath}`);

        const checker = setInterval(() => {
            const dirsToCheck = [downloadPath];
            if (fs.existsSync(defaultDownloadPath)) dirsToCheck.push(defaultDownloadPath);

            let foundFile = null;
            let foundDir = null;

            for (const dir of dirsToCheck) {
                try {
                    const files = fs.readdirSync(dir);
                    if (files.length > 0) {
                        const validFiles = files.filter(f => !f.startsWith('.') && !f.endsWith('.crdownload') && !f.endsWith('.tmp'));
                        if (validFiles.length > 0) {
                            const latest = validFiles
                                .map(f => ({ name: f, path: path.join(dir, f), time: fs.statSync(path.join(dir, f)).mtime.getTime() }))
                                .sort((a, b) => b.time - a.time)[0];
                            
                            if (latest && (Date.now() - latest.time < timeout + 60000)) { 
                                foundFile = latest; foundDir = dir; break;
                            }
                        }
                    }
                } catch (e) { }
            }

            if (foundFile) {
                const filePath = foundFile.path;
                const size1 = fs.statSync(filePath).size;
                
                if (size1 > 0) {
                    console.log(`      Found potential file: ${foundFile.name} in ${foundDir}`);
                    setTimeout(() => {
                        if (fs.existsSync(filePath)) {
                            const size2 = fs.statSync(filePath).size;
                            if (size1 === size2) {
                                clearInterval(checker); clearTimeout(timer);
                                console.log(`      File confirmed: ${filePath}`);
                                resolve(filePath);
                            } else {
                                console.log(`      File still downloading (${size1} -> ${size2})...`);
                            }
                        }
                    }, 5000); 
                    return;
                }
            }

            timePassed += checkInterval;
            if (timePassed >= timeout) {
                clearInterval(checker); clearTimeout(timer);
                reject(new Error(`Download timeout (${timeout}ms). No new files found.`));
            }
        }, checkInterval);
    });
}

// ฟังก์ชันจัดรูปแบบ Sheet
function formatSheet(worksheet) {
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        if (column && column.eachCell) {
            column.eachCell({ includeEmpty: true }, function(cell) {
                const columnLength = cell.text ? cell.text.length : 10;
                if (columnLength > maxLength) maxLength = columnLength;
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });
            column.width = maxLength < 10 ? 10 : maxLength + 2;
        }
    });

    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
}

// -------------------------------------------------------------
// NEW: ตัวช่วยแปลงวันที่ (แก้ปัญหา Excel Serial Date 46180.878)
// -------------------------------------------------------------
function formatExcelDate(value) {
    let d;
    if (value instanceof Date) {
        d = value;
    } else if (typeof value === 'number' && value > 40000 && value < 60000) {
        // สูตรคำนวณวันจาก Excel Serial Number (นับจากปี 1900)
        d = new Date((value - (25567 + 2)) * 86400 * 1000);
        d = new Date(d.getTime() + (d.getTimezoneOffset() * 60000));
    } else {
        return null;
    }

    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
    
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// ฟังก์ชันทำ Pivot และแยกข้อมูล Raw Data
async function processExcelFile(filePath) {
    try {
        console.log(`   Processing Excel file (with exceljs): ${filePath}`);
        const workbook = new ExcelJS.Workbook();
        const ext = path.extname(filePath).toLowerCase();
        
        try {
            if (ext === '.csv') await workbook.csv.readFile(filePath);
            else await workbook.xlsx.readFile(filePath);
        } catch (e) {
            console.error('   Read failed, trying as CSV fallback...');
            await workbook.csv.readFile(filePath);
        }

        const worksheet = workbook.worksheets[0]; 
        if (!worksheet) {
            console.warn('   No worksheet found.');
            return { filePath, pivotData: {}, rawData: [] };
        }
        
        const firstRow = worksheet.getRow(1);
        const headers = [];
        firstRow.eachCell((cell, colNumber) => {
            headers[colNumber] = cell.value ? cell.value.toString().trim() : '';
        });

        let licensePlateIndex = -1;
        let reportTypeIndex = -1;

        headers.forEach((header, index) => {
            if (header) {
                if (header.includes('ทะเบียน') || header.includes('License') || header.includes('ชื่อรถ')) licensePlateIndex = index;
                if (header.includes('ชนิด') || header.includes('Type') || header.includes('Alarm') || header.includes('Event')) reportTypeIndex = index;
            }
        });

        if (licensePlateIndex === -1) licensePlateIndex = 1;
        if (reportTypeIndex === -1) reportTypeIndex = 2; // Column B ใน Excel (Index 2 ใน exceljs)

        // --- ดึงข้อมูล Raw Data เพื่อเอาไปใส่ Google Sheets ---
        const rawDataToUpload = [];
        const colCount = worksheet.columnCount;
        
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // ข้าม Header (บรรทัดแรก)
            
            // อ่านค่าในคอลัมน์ชนิดรายงาน (คอลัมน์ B คือ index 2)
            const typeCell = row.getCell(reportTypeIndex);
            const reportType = typeCell.text ? typeCell.text.trim() : '';
            
            // เช็คเงื่อนไข: ดึงเฉพาะที่มีชนิดรายงานเป็น "แจ้งเตือนการหาวนอน" (หรือมีคำนี้อยู่)
            if (reportType.includes('แจ้งเตือนการหาวนอน') || reportType.includes('Yawning')) {
                const rowData = [];
                for (let c = 1; c <= colCount; c++) {
                    const cell = row.getCell(c);
                    let textVal = '';

                    // ใช้ฟังก์ชันแปลงวันที่ที่เราสร้างไว้ด้านบน
                    const formattedDate = formatExcelDate(cell.value);
                    
                    if (formattedDate) {
                        textVal = formattedDate; // ได้วันที่ออกมาสวยงาม
                    } else {
                        textVal = cell.text ? cell.text.trim() : ''; // ถ้าเป็นข้อความทั่วไป
                    }
                    
                    rowData.push(textVal);
                }
                rawDataToUpload.push(rowData);
            }
        });
        // ------------------------------------------------

        formatSheet(worksheet);

        const pivotData = {};
        const allTypes = new Set();

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; 

            const plateCell = row.getCell(licensePlateIndex);
            const typeCell = row.getCell(reportTypeIndex);
            
            const plate = plateCell.value ? plateCell.value.toString() : null;
            const type = typeCell.value ? typeCell.value.toString() : null;

            if (plate && type) {
                if (!pivotData[plate]) pivotData[plate] = {};
                if (!pivotData[plate][type]) pivotData[plate][type] = 0;
                
                pivotData[plate][type]++;
                allTypes.add(type);
            }
        });

        const pivotSheetName = "Summary_Pivot";
        const oldSheet = workbook.getWorksheet(pivotSheetName);
        if (oldSheet) workbook.removeWorksheet(oldSheet.id);

        const pivotSheet = workbook.addWorksheet(pivotSheetName);

        const typeArray = Array.from(allTypes).sort();
        const pivotHeaders = ['ทะเบียนรถ', ...typeArray, 'รวมทั้งหมด'];
        pivotSheet.addRow(pivotHeaders);

        for (const plate in pivotData) {
            const rowData = [plate];
            let total = 0;
            for (const type of typeArray) {
                const count = pivotData[plate][type] || 0;
                rowData.push(count);
                total += count;
            }
            rowData.push(total);
            pivotSheet.addRow(rowData);
        }

        formatSheet(pivotSheet);

        let outputFilePath = filePath;
        if (path.extname(filePath) !== '.xlsx') {
            outputFilePath = filePath.substring(0, filePath.lastIndexOf('.')) + '.xlsx';
        }
        
        await workbook.xlsx.writeFile(outputFilePath);
        console.log(`   Excel file processed and saved to: ${outputFilePath}`);
        console.log(`   Filtered ${rawDataToUpload.length} rows of 'แจ้งเตือนการหาวนอน' for Google Sheets.`);
        
        return { filePath: outputFilePath, pivotData: pivotData, rawData: rawDataToUpload };

    } catch (error) {
        console.error('   Error processing Excel file:', error.message);
        return { filePath: filePath, pivotData: {}, rawData: [] };
    }
}

// ฟังก์ชันสร้าง PDF (DMS Dashboard)
async function generatePDFSummary(page, pivotData, dateStr) {
    console.log('   Generating PDF Summary Report...');
    
    let yawningStats = [];
    let sleepingStats = [];
    let totalYawning = 0;
    let totalSleeping = 0;

    for (const [license, types] of Object.entries(pivotData)) {
        const yawn = types['แจ้งเตือนการหาวนอน'] || types['Yawning'] || 0;
        const sleep = types['แจ้งเตือนการหลับตา'] || types['Closing eyes'] || 0;
        
        if (yawn > 0) { yawningStats.push({ license, count: yawn }); totalYawning += yawn; }
        if (sleep > 0) { sleepingStats.push({ license, count: sleep }); totalSleeping += sleep; }
    }

    yawningStats.sort((a, b) => b.count - a.count);
    sleepingStats.sort((a, b) => b.count - a.count);

    const top10Yawning = yawningStats.slice(0, 10);
    const top10Sleeping = sleepingStats.slice(0, 10);

    const maxYawnCount = top10Yawning.length > 0 ? top10Yawning[0].count : 1;
    const maxSleepCount = top10Sleeping.length > 0 ? top10Sleeping[0].count : 1;

    const html = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
        <style>
        @page { size: A4; margin: 0; }
        body { font-family: 'Noto Sans Thai', sans-serif; margin: 0; padding: 0; background: #fff; color: #333; }
        .page { width: 210mm; height: 296mm; position: relative; page-break-after: always; overflow: hidden; }
        .content { padding: 40px; }
        .header-banner { background: #1E40AF; color: white; padding: 15px 40px; font-size: 24px; font-weight: bold; margin-bottom: 30px; }
        h1 { font-size: 32px; color: #1E40AF; margin-bottom: 10px; }
        .grid-2x2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 50px; }
        .card { background: #F8FAFC; border-radius: 12px; padding: 30px; text-align: center; border: 1px solid #E2E8F0; }
        .card-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }
        .card-value { font-size: 48px; font-weight: bold; margin: 10px 0; }
        .card-sub { font-size: 14px; color: #64748B; }
        .c-blue { color: #1E40AF; }
        .c-orange { color: #F59E0B; }
        .c-red { color: #DC2626; }
        .chart-container { margin: 40px 0; }
        .bar-row { display: flex; align-items: center; margin-bottom: 15px; }
        .bar-label { width: 200px; text-align: right; padding-right: 15px; font-weight: 600; font-size: 14px; }
        .bar-track { flex-grow: 1; background: #F1F5F9; height: 30px; border-radius: 4px; overflow: hidden; }
        .bar-fill { height: 100%; display: flex; align-items: center; justify-content: flex-end; padding-right: 10px; color: white; font-size: 12px; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #1E40AF; color: white; padding: 12px; text-align: left; font-size: 14px; }
        td { padding: 10px; border-bottom: 1px solid #E2E8F0; font-size: 14px; }
        tr:nth-child(even) { background: #F8FAFC; }
        </style>
    </head>
    <body>

        <!-- Page 1: Summary -->
        <div class="page">
            <div style="text-align: center; padding-top: 60px;">
                <h1 style="font-size: 40px;">รายงานสรุปพฤติกรรมการขับขี่ (DMS)</h1>
                <div style="font-size: 24px; color: #64748B;">Driver Monitoring System Report</div>
                <div style="margin-top: 20px; font-size: 18px;">วันที่สิ้นสุด: ${dateStr} (รอบเวลา 18:00 - 06:00 น.)</div>
            </div>

            <div class="content">
                <div class="header-banner" style="margin-top: 40px; text-align: center;">บทสรุปยอดรวม (Executive Summary)</div>
                <div class="grid-2x2">
                    <div class="card">
                        <div class="card-title c-orange">แจ้งเตือนการหาวนอน (Yawning)</div>
                        <div class="card-value c-orange">${totalYawning}</div>
                        <div class="card-sub">จำนวนครั้งทั้งหมด</div>
                    </div>
                    <div class="card">
                        <div class="card-title c-red">แจ้งเตือนการหลับตา (Closing Eyes)</div>
                        <div class="card-value c-red">${totalSleeping}</div>
                        <div class="card-sub">จำนวนครั้งทั้งหมด</div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Page 2: Yawning Details -->
        <div class="page">
            <div class="header-banner">1. สถิติแจ้งเตือนการหาวนอน (Top 10)</div>
            <div class="content">
                <div class="chart-container">
                ${top10Yawning.slice(0, 5).map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.count / maxYawnCount) * 100}%; background: #F59E0B;">${item.count} ครั้ง</div>
                    </div>
                    </div>
                `).join('')}
                ${top10Yawning.length === 0 ? '<div style="text-align:center; padding: 20px; color:#888;">ไม่มีข้อมูลในรอบเวลานี้</div>' : ''}
                </div>

                <table>
                    <thead>
                        <tr><th style="width:50px;">No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้งที่แจ้งเตือน (ครั้ง)</th></tr>
                    </thead>
                    <tbody>
                        ${top10Yawning.map((item, idx) => `
                        <tr>
                            <td>${idx + 1}</td>
                            <td>${item.license}</td>
                            <td style="font-weight: bold; color: #F59E0B;">${item.count}</td>
                        </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Page 3: Sleeping Details -->
        <div class="page">
            <div class="header-banner">2. สถิติแจ้งเตือนการหลับตา (Top 10)</div>
            <div class="content">
                <div class="chart-container">
                ${top10Sleeping.slice(0, 5).map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.count / maxSleepCount) * 100}%; background: #DC2626;">${item.count} ครั้ง</div>
                    </div>
                    </div>
                `).join('')}
                ${top10Sleeping.length === 0 ? '<div style="text-align:center; padding: 20px; color:#888;">ไม่มีข้อมูลในรอบเวลานี้</div>' : ''}
                </div>

                <table>
                    <thead>
                        <tr><th style="width:50px;">No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้งที่แจ้งเตือน (ครั้ง)</th></tr>
                    </thead>
                    <tbody>
                        ${top10Sleeping.map((item, idx) => `
                        <tr>
                            <td>${idx + 1}</td>
                            <td>${item.license}</td>
                            <td style="font-weight: bold; color: #DC2626;">${item.count}</td>
                        </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>

    </body>
    </html>
    `;

    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfPath = path.join(downloadPath, `DMS_Report_Summary_Night_${dateStr}.pdf`);
    await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });
    console.log(`   ✅ PDF Generated successfully: ${pdfPath}`);
    return pdfPath;
}

// ฟังก์ชันอัปโหลดข้อมูลไปยัง Google Sheets
async function appendToGoogleSheet(dataToAppend) {
    console.log(`   Uploading ${dataToAppend.length} rows to Google Sheets...`);
    try {
        const credentialsStr = process.env.GOOGLE_CREDENTIALS;
        if (!credentialsStr) {
            console.warn('   ⚠️ Skipping Google Sheets upload: GOOGLE_CREDENTIALS not found in secrets.');
            return;
        }

        const credentials = JSON.parse(credentialsStr);
        const auth = new google.auth.GoogleAuth({
            credentials,
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });

        const sheets = google.sheets({ version: 'v4', auth });
        
        // กำหนดช่วง (Range) เช่น 'Sheet1!A:Z'
        const range = `${config.googleSheetTabName}!A:Z`;

        await sheets.spreadsheets.values.append({
            spreadsheetId: config.googleSheetId,
            range: range,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            requestBody: {
                values: dataToAppend
            }
        });
        console.log(`   ✅ Successfully appended to Google Sheets.`);
    } catch (error) {
        console.error('   ❌ Failed to upload to Google Sheets:', error.message);
    }
}

// ฟังก์ชันส่งอีเมล
async function sendEmail(subject, message, attachmentPaths = []) {
    if (!config.emailFrom || !config.emailPass) {
        console.log('Skipping email: No credentials provided.');
        return;
    }

    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: config.emailFrom, pass: config.emailPass }
    });

    const attachments = attachmentPaths.filter(p => p && fs.existsSync(p)).map(p => ({
        filename: path.basename(p), 
        path: p
    }));

    const mailOptions = {
        from: `"Thai Tracking DMS Reporter" <${config.emailFrom}>`,
        to: config.emailTo,
        subject: subject,
        text: message,
        attachments: attachments
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log('Email sent successfully.');
    } catch (err) {
        console.error('Failed to send email:', err);
    }
}

async function clickByXPath(page, xpath, description = 'Element', timeout = 10000) {
    try {
        const selector = xpath.startsWith('xpath/') ? xpath : `xpath/${xpath}`;
        await page.waitForSelector(selector, { timeout: timeout, visible: true });
        const elements = await page.$$(selector);
        if (elements.length > 0) {
            await elements[0].click();
            console.log(`   Clicked: ${description}`);
        } else {
            throw new Error(`Element not found: ${description}`);
        }
    } catch (e) {
        throw new Error(`Failed to click ${description} (${xpath}): ${e.message}`);
    }
}

(async () => {
    console.log(`--- Started GPS Report Automation (Night Shift) [${new Date().toLocaleString()}] ---`);
    
    if (!config.gpsUser || !config.gpsPass) {
        console.warn("WARNING: GPS_USER or GPS_PASSWORD is missing.");
    }

    const browser = await puppeteer.launch({
        headless: "new",
        ignoreHTTPSErrors: true, 
        args: [
            '--no-sandbox', 
            '--disable-setuid-sandbox',
            '--window-size=1920,1080',
            '--disable-popup-blocking',
            '--allow-running-insecure-content',
            '--ignore-certificate-errors',
            '--unsafely-treat-insecure-origin-as-secure=http://cctvwli.com:3001',
            '--disable-web-security', 
            '--disable-features=IsolateOrigins,site-per-process,SafeBrowsing,DownloadBubble,DownloadBubbleV2',
            '--disable-site-isolation-trials',
            '--disable-client-side-phishing-detection',
            '--safebrowsing-disable-auto-update',
            '--safebrowsing-disable-download-protection',
            '--safebrowsing-disable-extension-blacklist',
            '--no-first-run',
            '--no-default-browser-check',
            '--lang=th-TH' 
        ]
    });

    const page = await browser.newPage();
    
    await page.setExtraHTTPHeaders({
        'Accept-Language': 'th-TH,th;q=0.9,en;q=0.8'
    });
    
    try {
        const client = await page.target().createCDPSession();
        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
        await client.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath, eventsEnabled: true }); 
    } catch(e) { }

    try {
        const context = browser.defaultBrowserContext();
        await context.overridePermissions('http://cctvwli.com:3001', ['automatic-downloads']);
    } catch(e) { console.log('   Warning: Permission override failed:', e.message); }

    page.setDefaultTimeout(60000);

    try {
        let isLoggedIn = false;
        const maxRetries = 20;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                console.log(`\n>>> Login Attempt ${attempt}/${maxRetries} <<<`);
                await page.goto('https://dvrai.net/808gps/login.html', { waitUntil: 'networkidle0' });
                await page.waitForSelector('#lwm'); 
                await new Promise(r => setTimeout(r, 2000)); 
                
                const captchaElement = await page.$('#lwm');
                if (!captchaElement) throw new Error('Captcha not found');
                const captchaImage = await captchaElement.screenshot();
                
                const worker = await Tesseract.createWorker('eng');
                await worker.setParameters({ tessedit_char_whitelist: '0123456789' });
                const { data: { text } } = await worker.recognize(captchaImage);
                await worker.terminate();
                
                const captchaCode = text.trim().replace(/\s/g, '');
                console.log(`   READ CAPTCHA: "${captchaCode}"`);

                if (!captchaCode || captchaCode.length < 4) { continue; }

                await page.type('#loginAccount', config.gpsUser);
                await page.type('#loginPassword', config.gpsPass);
                await page.type('#phraseLogin', captchaCode);

                console.log('   Clicking Login...');
                await Promise.all([
                    page.click('#loginSubmit'),
                    page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 5000 }).catch(() => {})
                ]);

                if (!page.url().includes('login.html')) {
                    console.log('   SUCCESS: Login Successful!');
                    isLoggedIn = true;
                    await new Promise(r => setTimeout(r, 10000));
                    break; 
                }
            } catch (err) { }
        }

        if (!isLoggedIn) throw new Error(`Failed to login.`);

        // --- STEP 5: Report Center ---
        console.log('5. Accessing Report Center...');
        let reportPage = null;
        const initialPages = await browser.pages();
        const initialPageCount = initialPages.length;
        const startTime = Date.now();

        while (Date.now() - startTime < 60000) {
            const currentPages = await browser.pages();
            if (currentPages.length > initialPageCount) {
                reportPage = currentPages[currentPages.length - 1]; 
                break;
            }
            try {
                const jsResult = await page.evaluate(() => {
                    if (typeof showReportCenter === 'function') { showReportCenter(); return true; } 
                    else {
                        const btn = document.querySelector('div[onclick*="showReportCenter"]') || document.querySelector('#main-topPanel > div.header-nav > div:nth-child(7)');
                        if (btn) { btn.click(); return true; }
                    }
                    return false;
                });
            } catch (e) {}
            await new Promise(r => setTimeout(r, 5000));
        }

        if (!reportPage) {
            const finalPages = await browser.pages();
            if (finalPages.length > initialPageCount) reportPage = finalPages[finalPages.length - 1];
            else throw new Error("Failed to open Report Center.");
        }
        
        try { await reportPage.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }); } catch(e) {}
        try { await reportPage.waitForSelector('xpath//*[@id="root"]', { timeout: 10000 }); } catch (e) {}
        await reportPage.setViewport({ width: 1920, height: 1080 });

        try {
            const clientReport = await reportPage.target().createCDPSession();
            await clientReport.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
            await clientReport.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath, eventsEnabled: true });
        } catch (e) {}

        // --- STEP 6: Report Filters ---
        console.log('6. Configuring Report Filters...');
        let dmsClicked = false;
        
        const dmsSelectors = [
            '//*[local-name()="svg" and @data-testid="FaceIcon"]/..', 
            '//*[@id="root"]/div/div[2]/div[1]/div/button[2]', 
            '//button[contains(., "รายงาน DMS")]'
        ];

        for (const selector of dmsSelectors) {
            if (dmsClicked) break;
            try {
                const xpSelector = `xpath/${selector}`;
                await reportPage.waitForSelector(xpSelector, { visible: true, timeout: 5000 });
                const elements = await reportPage.$$(xpSelector);
                if (elements.length > 0) {
                    await elements[0].click();
                    console.log(`   Clicked DMS via XPath: ${selector}`);
                    dmsClicked = true;
                }
            } catch (e) {}
        }

        if (!dmsClicked) {
            try {
                dmsClicked = await reportPage.evaluate(() => {
                    const buttons = Array.from(document.querySelectorAll('button'));
                    const dmsBtn = buttons.find(b => b.textContent.includes('รายงาน DMS'));
                    if (dmsBtn) { dmsBtn.click(); return true; }
                    return false;
                });
                if (dmsClicked) console.log('   Clicked DMS via JS Text Search');
            } catch (e) {}
        }
        
        if (!dmsClicked) throw new Error('Could not select DMS Report button.');

        await new Promise(r => setTimeout(r, 2000)); 
        await clickByXPath(reportPage, '//div[contains(@class, "css-xn5mga")]//tr[2]//td[2]//div/div', 'Alert Type Dropdown');
        await new Promise(r => setTimeout(r, 1000));

        const selectOption = async (optionText) => {
            const selector = `xpath///div[contains(text(), '${optionText}')]`;
            const elements = await reportPage.$$(selector);
            if (elements.length > 0) await elements[0].click();
        };

        await selectOption('แจ้งเตือนการหาวนอน');
        await new Promise(r => setTimeout(r, 500));
        await selectOption('แจ้งเตือนการหลับตา');
        await reportPage.keyboard.press('Escape');

        // NEW: Night Shift Logic (18:00 Yesterday - 06:00 Today) + Timezone Asia/Bangkok
        const thaiDateObj = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
        const yyyy = thaiDateObj.getFullYear();
        const mm = String(thaiDateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(thaiDateObj.getDate()).padStart(2, '0');
        const todayStr = `${yyyy}-${mm}-${dd}`;

        const yesterdayObj = new Date(thaiDateObj);
        yesterdayObj.setDate(yesterdayObj.getDate() - 1);
        const y_yyyy = yesterdayObj.getFullYear();
        const y_mm = String(yesterdayObj.getMonth() + 1).padStart(2, '0');
        const y_dd = String(yesterdayObj.getDate()).padStart(2, '0');
        const yesterdayStr = `${y_yyyy}-${y_mm}-${y_dd}`;

        const startDateTime = `${yesterdayStr} 18:00:00`;
        const endDateTime = `${todayStr} 06:00:00`;
        
        console.log(`   Setting Time: ${startDateTime} to ${endDateTime}`);

        await clickByXPath(reportPage, '//div[contains(@class, "css-xn5mga")]//tr[3]//td[2]//input', 'Start Date');
        await reportPage.keyboard.down('Control'); await reportPage.keyboard.press('A'); await reportPage.keyboard.up('Control');
        await reportPage.keyboard.press('Backspace'); await reportPage.keyboard.type(startDateTime); await reportPage.keyboard.press('Enter');

        await clickByXPath(reportPage, '//div[contains(@class, "css-xn5mga")]//tr[3]//td[4]//input', 'End Date');
        await reportPage.keyboard.down('Control'); await reportPage.keyboard.press('A'); await reportPage.keyboard.up('Control');
        await reportPage.keyboard.press('Backspace'); await reportPage.keyboard.type(endDateTime); await reportPage.keyboard.press('Enter');

        await new Promise(r => setTimeout(r, 500)); 
        await reportPage.keyboard.press('Tab'); 
        await new Promise(r => setTimeout(r, 300));
        await reportPage.keyboard.press('Enter'); 
        
        console.log('   Waiting 120s for report generation...');
        await new Promise(r => setTimeout(r, 120000));

        console.log('   Clicking EXCEL (via Keyboard Tab+Enter)...');
        await reportPage.keyboard.press('Tab');
        await new Promise(r => setTimeout(r, 500));
        await reportPage.keyboard.press('Enter');
        
        console.log('   Waiting 60s for Save Dialog...');
        await new Promise(r => setTimeout(r, 60000)); 
        
        console.log('   Clicking SAVE (Floppy Disk)...');
        const saveXPath = `//*[@id="root"]/div/div[1]/div[2]/div[2]/div/div/div/ul/li/div/div/div/div/button/svg | //*[@data-testid="SaveOutlinedIcon"]`;
        await clickByXPath(reportPage, saveXPath, 'Save Icon', 60000).catch(() => console.log('Try JS Save'));

        console.log('7. Waiting for file download...');
        let downloadedFile = await waitForFileToDownload(config.downloadTimeout);

        const ext = path.extname(downloadedFile);
        if (!ext || (ext !== '.xls' && ext !== '.xlsx')) {
            const dir = path.dirname(downloadedFile);
            const newName = `GPS_Report_Night_${todayStr}.xlsx`;
            const newFilePath = path.join(dir, newName);
            if (fs.existsSync(newFilePath)) try { fs.unlinkSync(newFilePath); } catch(e) {}
            try { fs.renameSync(downloadedFile, newFilePath); downloadedFile = newFilePath; } catch (e) {}
        }

        // --- NEW STEP: Process Excel & Extract Pivot Data + Raw Data ---
        const processResult = await processExcelFile(downloadedFile);
        downloadedFile = processResult.filePath;
        const pivotData = processResult.pivotData;
        const rawDataToUpload = processResult.rawData;

        // --- NEW STEP: Upload to Google Sheets ---
        if (rawDataToUpload && rawDataToUpload.length > 0) {
            await appendToGoogleSheet(rawDataToUpload);
        } else {
            console.log('   ⚠️ No raw data found to upload to Google Sheets.');
        }

        // --- NEW STEP: Generate PDF ---
        const pdfFilePath = await generatePDFSummary(page, pivotData, todayStr);

        // --- STEP 8: Email ---
        console.log(`8. Sending Email...`);
        await sendEmail(
            `THAI TRACKING DMS REPORT: ${todayStr} (Night Shift)`, 
            `ถึง ผู้เกี่ยวข้อง\nรายงาน THAI TRACKING DMS REPORT รอบ 18:00 ถึง 06:00 น. และ สรุปกราฟ PDF Top 10\nhttps://script.google.com/macros/s/AKfycbzhTXXdegU_3Sx9gaRtcAJdWsIpudOWq-sUW5rxDGsqt2eX31MR3Ikk6-fhAlaV8Yd0uw/exec\nด้วยความนับถือ\nBOT REPORT`, 
            [downloadedFile, pdfFilePath] 
        );

        // --- STEP 9: Cleanup ---
        console.log('9. Cleaning up files...');
        if (fs.existsSync(downloadedFile)) fs.unlinkSync(downloadedFile);
        if (fs.existsSync(pdfFilePath)) fs.unlinkSync(pdfFilePath);
        console.log('   Cleanup done.');

    } catch (error) {
        console.error('!!! PROCESS FAILED !!!', error);
        const pages = await browser.pages();
        const activePage = pages[pages.length - 1]; 
        const errorScreenshotPath = path.resolve(__dirname, 'error_debug.png');
        await activePage.screenshot({ path: errorScreenshotPath, fullPage: true });
        await sendEmail(`GPS Automation FAILED`, `Error details: ${error.message}`, [errorScreenshotPath]);
        process.exit(1);
    } finally {
        await browser.close();
    }
})();
