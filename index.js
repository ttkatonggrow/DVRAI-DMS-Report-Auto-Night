const puppeteer = require('puppeteer');
const Tesseract = require('tesseract.js');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const os = require('os');
const ExcelJS = require('exceljs'); // ใช้ exceljs แทน xlsx

// --- CONFIGURATION ---
const config = {
    gpsUser: process.env.GPS_USER || '',
    gpsPass: process.env.GPS_PASSWORD || '',
    emailFrom: process.env.EMAIL_FROM || '',
    emailPass: process.env.EMAIL_PASSWORD || '',
    emailTo: process.env.EMAIL_TO || '',
    downloadTimeout: 40000 
};

// กำหนด Path หลักที่เราต้องการ (ใน Project folder)
const downloadPath = path.resolve(__dirname, 'downloads');
// กำหนด Path สำรอง (Default Downloads ของ User)
const defaultDownloadPath = path.join(os.homedir(), 'Downloads');

// สร้างโฟลเดอร์ download ถ้ายังไม่มี
if (!fs.existsSync(downloadPath)) {
    fs.mkdirSync(downloadPath);
}

// ฟังก์ชันรอจนกว่าไฟล์จะโหลดเสร็จ
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
            if (fs.existsSync(defaultDownloadPath)) {
                dirsToCheck.push(defaultDownloadPath);
            }

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
                                foundFile = latest;
                                foundDir = dir;
                                break;
                            }
                        }
                    }
                } catch (e) { /* Ignore access errors */ }
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
                                clearInterval(checker);
                                clearTimeout(timer);
                                console.log(`      File confirmed: ${filePath}`);
                                resolve(filePath);
                            }
                        }
                    }, 3000);
                    return;
                }
            }

            timePassed += checkInterval;
            if (timePassed >= timeout) {
                clearInterval(checker);
                clearTimeout(timer);
                reject(new Error(`Download timeout (${timeout}ms). No new files found.`));
            }
        }, checkInterval);
    });
}

// ฟังก์ชันจัดรูปแบบ Sheet (Border + Auto Width + Header Style)
function formatSheet(worksheet) {
    // 1. จัด Header (แถวแรก)
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' } // สีเทาอ่อน
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // 2. จัด Auto Width และ Border ให้ข้อมูล
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        if (column && column.eachCell) {
            column.eachCell({ includeEmpty: true }, function(cell) {
                const columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
                // ใส่เส้นขอบทุกช่อง
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
            column.width = maxLength < 10 ? 10 : maxLength + 2;
        }
    });
}

// ฟังก์ชันประมวลผล Excel (ทำ Pivot Table ด้วย exceljs + จัดรูปแบบเต็มสูบ)
async function processExcelFile(filePath) {
    try {
        console.log(`   Processing Excel file (with exceljs): ${filePath}`);
        
        const workbook = new ExcelJS.Workbook();
        
        // ตรวจสอบนามสกุลไฟล์เพื่อใช้วิธีอ่านที่ถูกต้อง
        const ext = path.extname(filePath).toLowerCase();
        if (ext === '.csv') {
            await workbook.csv.readFile(filePath);
        } else {
            await workbook.xlsx.readFile(filePath);
        }

        const worksheet = workbook.worksheets[0]; // เอา Sheet แรก
        
        // จัดรูปแบบ Sheet เดิมก่อน
        formatSheet(worksheet);

        // --- เตรียมข้อมูลทำ Pivot ---
        // อ่าน Header row (แถว 1)
        const firstRow = worksheet.getRow(1);
        const headers = [];
        firstRow.eachCell((cell, colNumber) => {
            headers[colNumber] = cell.value ? cell.value.toString().trim() : '';
        });

        console.log(`   Headers found: ${JSON.stringify(headers)}`);

        let licensePlateIndex = -1;
        let reportTypeIndex = -1;

        // หา Index ของคอลัมน์ (ExcelJS index เริ่มที่ 1)
        headers.forEach((header, index) => {
            if (header) {
                if (header.includes('ทะเบียน') || header.includes('License') || header.includes('ชื่อรถ')) licensePlateIndex = index;
                if (header.includes('ชนิด') || header.includes('Type') || header.includes('Alarm') || header.includes('Event')) reportTypeIndex = index;
            }
        });

        // Fallback Index (ถ้าหาไม่เจอ ใช้คอลัมน์ 1 และ 2)
        if (licensePlateIndex === -1) {
            console.log('   Warning: "License" header not found. Defaulting to Column 1.');
            licensePlateIndex = 1;
        }
        if (reportTypeIndex === -1) {
            console.log('   Warning: "Type" header not found. Defaulting to Column 2.');
            reportTypeIndex = 2;
        }

        const pivotData = {};
        const allTypes = new Set();

        // วนลูปข้อมูล (เริ่มแถว 2)
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // ข้าม Header

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

        // --- สร้าง Sheet ใหม่ (Summary_Pivot) ---
        const pivotSheetName = "Summary_Pivot";
        const oldSheet = workbook.getWorksheet(pivotSheetName);
        if (oldSheet) {
            workbook.removeWorksheet(oldSheet.id);
        }

        const pivotSheet = workbook.addWorksheet(pivotSheetName);

        // สร้าง Header สำหรับ Pivot
        const typeArray = Array.from(allTypes).sort();
        const pivotHeaders = ['ทะเบียนรถ', ...typeArray, 'รวมทั้งหมด'];
        pivotSheet.addRow(pivotHeaders);

        // ใส่ข้อมูล Pivot
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

        // จัดรูปแบบ Sheet ใหม่ (เส้นขอบ + Auto Width)
        formatSheet(pivotSheet);

        // บันทึกไฟล์ทับ (ต้องเป็น .xlsx)
        let outputFilePath = filePath;
        if (ext !== '.xlsx') {
            outputFilePath = filePath.replace(ext, '.xlsx');
        }
        
        await workbook.xlsx.writeFile(outputFilePath);
        console.log(`   Excel file processed and saved to: ${outputFilePath}`);
        
        return outputFilePath;

    } catch (error) {
        console.error('   Error processing Excel file with exceljs:', error.message);
        return filePath;
    }
}

// ฟังก์ชันส่งอีเมล
async function sendEmail(subject, message, attachmentPath = null) {
    if (!config.emailFrom || !config.emailPass) {
        console.log('Skipping email: No credentials provided.');
        return;
    }

    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: config.emailFrom, pass: config.emailPass }
    });

    const attachments = [];
    if (attachmentPath && fs.existsSync(attachmentPath)) {
        attachments.push({
            filename: path.basename(attachmentPath), 
            path: attachmentPath
        });
    }

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

// ฟังก์ชันช่วยคลิก Element โดยใช้ XPath
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
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath,
        });
        await client.send('Browser.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath,
            eventsEnabled: true 
        }); 
    } catch(e) { console.log('CDP Setup Warning:', e.message); }

    try {
        const context = browser.defaultBrowserContext();
        await context.overridePermissions('http://cctvwli.com:3001', ['automatic-downloads']);
    } catch(e) {}

    page.setDefaultTimeout(60000);

    try {
        // --- LOGIN LOOP ---
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

                if (!captchaCode || captchaCode.length < 4) {
                    console.warn(`   !!! Invalid Captcha. Retrying...`);
                    continue; 
                }

                await page.type('#loginAccount', config.gpsUser);
                await page.type('#loginPassword', config.gpsPass);
                await page.type('#phraseLogin', captchaCode);

                console.log('   Clicking Login...');
                await Promise.all([
                    page.click('#loginSubmit'),
                    page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 5000 }).catch(() => {})
                ]);

                if (page.url().includes('login.html')) {
                    console.warn('   !!! Login Failed. Retrying...');
                    continue; 
                } else {
                    console.log('   SUCCESS: Login Successful!');
                    isLoggedIn = true;
                    console.log('   Waiting 10s for dashboard...');
                    await new Promise(r => setTimeout(r, 10000));
                    break; 
                }
            } catch (err) {
                console.warn(`   Error during login: ${err.message}`);
            }
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
                console.log(`   >>> New tab detected! URL: ${reportPage.url()}`);
                
                const pageTitle = await reportPage.title();
                if (pageTitle.includes('Privacy') || pageTitle.includes('Security')) {
                    try {
                        const advanced = await reportPage.$('#details-button');
                        if (advanced) {
                            await advanced.click();
                            await new Promise(r => setTimeout(r, 1000));
                            const proceed = await reportPage.$('#proceed-link');
                            if (proceed) await proceed.click();
                        }
                    } catch (e) {}
                }
                break;
            }

            try {
                const jsResult = await page.evaluate(() => {
                    if (typeof showReportCenter === 'function') {
                        showReportCenter();
                        return 'Executed showReportCenter() directly';
                    } else {
                        const btn = document.querySelector('div[onclick*="showReportCenter"]') || 
                                    document.querySelector('#main-topPanel > div.header-nav > div:nth-child(7)');
                        if (btn) { btn.click(); return 'Clicked element via JS'; }
                    }
                    return null;
                });
                if (jsResult) console.log(`   Triggered: ${jsResult}`);
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
            await clientReport.send('Page.setDownloadBehavior', {
                behavior: 'allow',
                downloadPath: downloadPath,
            });
            await clientReport.send('Browser.setDownloadBehavior', { 
                behavior: 'allow', 
                downloadPath: downloadPath, 
                eventsEnabled: true 
            });
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
                    dmsClicked = true;
                }
            } catch (e) {}
        }

        if (!dmsClicked) {
             try {
                const jsClicked = await reportPage.evaluate(() => {
                    const buttons = Array.from(document.querySelectorAll('button'));
                    const dmsBtn = buttons.find(b => b.textContent.includes('รายงาน DMS'));
                    if (dmsBtn) { dmsBtn.click(); return true; }
                    return false;
                });
                if (jsClicked) dmsClicked = true;
             } catch (e) {}
        }
        
        if (!dmsClicked) throw new Error('Could not select DMS Report button.');

        console.log('   Selecting Alerts...');
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

        // --- Date Inputs (Logic: Yesterday 18:00 - Today 06:00) ---
        const todayObj = new Date();
        const yesterdayObj = new Date(todayObj);
        yesterdayObj.setDate(yesterdayObj.getDate() - 1);

        const todayStr = todayObj.toISOString().slice(0, 10);
        const yesterdayStr = yesterdayObj.toISOString().slice(0, 10);

        const startDateTime = `${yesterdayStr} 18:00:00`;
        const endDateTime = `${todayStr} 06:00:00`;
        
        console.log(`   Setting Time: ${startDateTime} to ${endDateTime}`);

        await clickByXPath(reportPage, '//div[contains(@class, "css-xn5mga")]//tr[3]//td[2]//input', 'Start Date');
        await reportPage.keyboard.down('Control'); await reportPage.keyboard.press('A'); await reportPage.keyboard.up('Control');
        await reportPage.keyboard.press('Backspace'); await reportPage.keyboard.type(startDateTime); await reportPage.keyboard.press('Enter');

        await clickByXPath(reportPage, '//div[contains(@class, "css-xn5mga")]//tr[3]//td[4]//input', 'End Date');
        await reportPage.keyboard.down('Control'); await reportPage.keyboard.press('A'); await reportPage.keyboard.up('Control');
        await reportPage.keyboard.press('Backspace'); await reportPage.keyboard.type(endDateTime); await reportPage.keyboard.press('Enter');

        // --- Tab + Enter to Search ---
        console.log('   Pressing Tab + Enter to Search...');
        await new Promise(r => setTimeout(r, 500)); 
        await reportPage.keyboard.press('Tab'); 
        await new Promise(r => setTimeout(r, 300));
        await reportPage.keyboard.press('Enter'); 
        
        console.log('   Waiting 120s for report generation...');
        await new Promise(r => setTimeout(r, 120000));

        // 6.5 กดปุ่ม EXCEL (Tab + Enter)
        console.log('   Clicking EXCEL (via Keyboard Tab+Enter)...');
        await reportPage.keyboard.press('Tab');
        await new Promise(r => setTimeout(r, 500));
        await reportPage.keyboard.press('Enter');
        console.log('   Pressed Enter on EXCEL button!');
        
        // SAVE
        console.log('   Waiting 20s for Save Dialog...');
        await new Promise(r => setTimeout(r, 20000)); 
        
        console.log('   Clicking SAVE (Floppy Disk)...');
        let saveClicked = false;
        saveClicked = await reportPage.evaluate(() => {
            const saveBtn = document.querySelector("#root > div > div.MuiBox-root.css-jbmhbb > div.ant-card.ant-card-bordered.css-y8x9xp > div.ant-card-body > div > div > div > ul > li > div > div > div > div > button");
            if (saveBtn) { saveBtn.click(); return true; }
            return false;
        });

        if (!saveClicked) {
            const saveXPath = `//*[@id="root"]/div/div[1]/div[2]/div[2]/div/div/div/ul/li/div/div/div/div/button/svg | //*[@data-testid="SaveOutlinedIcon"]`;
            await clickByXPath(reportPage, saveXPath, 'Save Icon', 60000);
        }

        // --- STEP 7: Wait for Download ---
        console.log('7. Waiting for file download...');
        let downloadedFile = await waitForFileToDownload(config.downloadTimeout);
        console.log(`   File downloaded: ${downloadedFile}`);

        // --- FIX: Rename file ---
        const ext = path.extname(downloadedFile);
        if (!ext || (ext !== '.xls' && ext !== '.xlsx')) {
            console.log(`   Renaming file to .xls...`);
            const dir = path.dirname(downloadedFile);
            const newName = `GPS_Report_${todayStr}.xls`;
            const newFilePath = path.join(dir, newName);
            if (fs.existsSync(newFilePath)) try { fs.unlinkSync(newFilePath); } catch(e) {}
            try {
                fs.renameSync(downloadedFile, newFilePath);
                downloadedFile = newFilePath;
                console.log(`   Renamed file to: ${downloadedFile}`);
            } catch (e) {}
        }

        // --- NEW STEP: Process Excel & Add Pivot Sheet (with ExcelJS) ---
        downloadedFile = await processExcelFile(downloadedFile);

        // --- STEP 8: Email ---
        console.log(`8. Sending Email...`);
        await sendEmail(
            `THAI TRACKING DMS REPORT: ${todayStr}`, 
            `ถึง ผู้เกี่ยวข้อง\nรายงาน THAI TRACKING DMS REPORT รอบ 18:00 ถึง 06:00 น.\nด้วยความนับถือ\nBOT REPORT`, 
            downloadedFile
        );

        // --- STEP 9: Cleanup ---
        console.log('9. Cleaning up...');
        if (fs.existsSync(downloadedFile)) {
            fs.unlinkSync(downloadedFile);
            console.log('   File deleted.');
        }

    } catch (error) {
        console.error('!!! PROCESS FAILED !!!', error);
        const pages = await browser.pages();
        const activePage = pages[pages.length - 1]; 
        const errorScreenshotPath = path.resolve(__dirname, 'error_debug.png');
        await activePage.screenshot({ path: errorScreenshotPath, fullPage: true });
        await sendEmail(`GPS Automation FAILED`, `Error details: ${error.message}`);
        process.exit(1);
    } finally {
        await browser.close();
    }
})();
