
import express from 'express';
import { createServer as createViteServer } from 'vite';
import axios from 'axios';
import cron from 'node-cron';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SHEET_ID = '1ylgWBQJQgv0HjRiV9Ay-GYs__ywWUk8ZTSeb0KwaPsM';
const SHEET_GID = '1144899065';
const CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SHEET_GID}`;

// Column indices based on inspection
const COLUMNS = {
    LC: 0,
    VENDOR: 3,
    SHIP_TO: 7,
    SHIPPINGLINE: 8,
    KODE_STORE: 39,
    STUFFING_DATE: 40,
    JAM: 41,
    KOLI_SIAP_KIRIM: 42, // CASE ID
    CBM_SHIPPED: 50,
    CASE_ID_SHIPPED: 51,
    PLAT_NOMOR: 58, // NOPOL
    NOMOR: 59,
    DRIVER: 60
};

async function startServer() {
    const app = express();
    const port = 3000;

    // API Routes
    app.get('/api/health', (req, res) => {
        res.json({ status: 'ok', time: new Date().toISOString() });
    });

    app.get('/api/test-email', async (req, res) => {
        try {
            console.log('Manual trigger of email automation...');
            await processEmails();
            res.json({ message: 'Email automation job triggered successfully.' });
        } catch (error: any) {
            res.status(500).json({ error: error.message });
        }
    });

    // Integrated Vite Server for development
    if (process.env.NODE_ENV !== 'production') {
        const vite = await createViteServer({
            server: { middlewareMode: true },
            appType: 'spa',
        });
        app.use(vite.middlewares);
    } else {
        const distPath = path.join(process.cwd(), 'dist');
        app.use(express.static(distPath));
        app.get('*', (req, res) => {
            res.sendFile(path.join(distPath, 'index.html'));
        });
    }

    app.listen(port, '0.0.0.0', () => {
        console.log(`Server running at http://0.0.0.0:${port}`);
    });

    // Schedule Email Task: 09:00 Daily
    cron.schedule('0 9 * * *', () => {
        console.log('Running daily email automation task at 09:00...');
        processEmails();
    }, {
        timezone: "Asia/Jakarta"
    });
}

async function processEmails() {
    try {
        console.log('Fetching data for emails...');
        const response = await axios.get(CSV_URL);
        const data = response.data;
        const rows = parseCSV(data);

        // Get H-1 Date string
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const hMinus1Str = formatDateForFilter(yesterday);
        console.log(`Filtering data for date: ${hMinus1Str}`);

        // Filter: Category Store, Date H-1, Shipped > 0, Koli Siap Kirim > 0
        const storeRows = rows.filter(row => {
            const isStore = (row[COLUMNS.SHIP_TO] || '').includes('ST AHI') || (row[COLUMNS.KODE_STORE] || '').startsWith('A');
            const dateMatch = (row[COLUMNS.STUFFING_DATE] || '') === hMinus1Str;
            const shipped = parseFloat((row[COLUMNS.CBM_SHIPPED] || '0').replace(/,/g, '')) || 0;
            const koliSiapKirim = parseFloat((row[COLUMNS.KOLI_SIAP_KIRIM] || '0').replace(/,/g, '')) || 0;

            return isStore && dateMatch && shipped > 0 && koliSiapKirim > 0;
        });

        console.log(`Found ${storeRows.length} rows matching criteria.`);

        if (storeRows.length === 0) return;

        const transporter = nodemailer.createTransport({
            host: process.env.SMTP_HOST,
            port: parseInt(process.env.SMTP_PORT || '587'),
            auth: {
                user: process.env.SMTP_USER,
                pass: process.env.SMTP_PASS
            }
        });

        for (const row of storeRows) {
            await sendFormattedEmail(transporter, row);
        }

    } catch (error) {
        console.error('Error in email automation:', error);
    }
}

async function sendFormattedEmail(transporter: any, row: any[]) {
    // Extract data
    const lc = row[COLUMNS.LC];
    const kodeStore = row[COLUMNS.KODE_STORE];
    const namaStoreFull = row[COLUMNS.SHIP_TO];
    const vendor = row[COLUMNS.VENDOR];
    const driver = row[COLUMNS.DRIVER];
    const nomor = row[COLUMNS.NOMOR];
    const platNomor = row[COLUMNS.PLAT_NOMOR];
    const jumlahKoli = row[COLUMNS.CASE_ID_SHIPPED]; // Column AZ (51) as per user request
    const shippingLine = row[COLUMNS.SHIPPINGLINE] || '';

    // Extract On Site Time from shippingLine or fallback
    let onSite = '07.00';
    const onSiteMatch = shippingLine.match(/ON SITE [^)]+/i);
    if (onSiteMatch) {
        onSite = onSiteMatch[0].replace(/ON SITE /i, '').trim();
    }

    // Determine recipient email
    // Since no email column was found, we use a placeholder or check if any column is an email
    // For now, let's use a mapping or a default to avoid failure
    let recipient = process.env.DEFAULT_RECIPIENT || 'planner.sidoarjo@kawanlama.com';
    
    // TEMPORARY: If LC matches some pattern, determine email
    // User said "kirim ke email setiap LC" - we need the email column index!
    // I'll add a check for all columns to see if one contains an '@'
    const foundEmailInRow = row.find(val => typeof val === 'string' && val.includes('@'));
    if (foundEmailInRow) recipient = foundEmailInRow;

    const subject = `Informasi Pengiriman Barang ke Store ${namaStoreFull} (ON SITE ${onSite})`;
    
    const body = `Yth. Tim Store,

Dengan hormat,

Berikut kami sampaikan informasi terkait pengiriman barang ke Store dengan detail sebagai berikut:

KODE STORE : ${kodeStore}
NAMA STORE : ${namaStoreFull}
VENDOR : ${vendor}
DRIVER : ${driver}
NOMOR : ${nomor}
PLAT NOMOR : ${platNomor}
JUMLAH KOLI: ${jumlahKoli}

Mohon untuk dapat dipersiapkan penerimaan barang sesuai dengan data di atas. Apabila terdapat pertanyaan atau kendala, silakan menghubungi pihak terkait.

Demikian informasi ini kami sampaikan. Terima kasih atas perhatian dan kerjasamanya.

Hormat kami,
GUNAWAN
PLANNER
DC SIDOARJO KAWAN LAMA`;

    try {
        await transporter.sendMail({
            from: process.env.SMTP_FROM || process.env.SMTP_USER,
            to: recipient,
            subject: subject,
            text: body
        });
        console.log(`Email sent for LC ${lc} to ${recipient}`);
    } catch (err) {
        console.error(`Failed to send email for LC ${lc}:`, err);
    }
}

function parseCSV(csv: string) {
    const rows: any[][] = [];
    let current: any[] = [];
    let field = '';
    let inQuotes = false;
    for (let i = 0; i < csv.length; i++) {
        const char = csv[i];
        const next = csv[i+1];
        if (char === '"') {
            if (inQuotes && next === '"') { field += '"'; i++; }
            else inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            current.push(field.trim());
            field = '';
        } else if ((char === '\n' || char === '\r') && !inQuotes) {
            current.push(field.trim());
            if (current.length > 0) rows.push(current);
            current = [];
            field = '';
            if (char === '\r' && next === '\n') i++;
        } else {
            field += char;
        }
    }
    return rows;
}

function formatDateForFilter(date: Date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    // Assuming CSV uses YYYY-MM-DD format based on sample data
    return `${y}-${m}-${d}`;
}

startServer();
