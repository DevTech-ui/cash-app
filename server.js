require('dotenv').config();
const crypto = require('crypto');
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs-extra');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const { google } = require('googleapis');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const ExcelJS = require('exceljs');
const { savePaymentRecord } = require('./payment-db');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data', 'emails.json');
const documentDownloads = new Map();
const DOWNLOAD_TTL_MS = 15 * 60 * 1000;

const submissionSchema = new mongoose.Schema({
  email: { type: String, lowercase: true, trim: true, index: true },
  name: { type: String, trim: true },
  amount: { type: String, default: '' },
  timestamp: { type: Date, default: Date.now },
  ip: { type: String, default: 'unknown' },
  payments: [{
    paymentId: String,
    amount: String,
    status: String,
    timestamp: { type: Date, default: Date.now },
  }],
}, { timestamps: true });

const Submission = mongoose.model('Submission', submissionSchema);

async function connectMongo() {
  if (!process.env.MONGODB_URI) {
    console.warn('[MongoDB] MONGODB_URI is not set. Skipping MongoDB storage.');
    return null;
  }

  if (mongoose.connection.readyState === 1) {
    return Submission;
  }

  try {
    await mongoose.connect(process.env.MONGODB_URI, {
      serverSelectionTimeoutMS: 5000,
    });
    console.log('[MongoDB] Connected successfully.');
    return Submission;
  } catch (err) {
    console.error('[MongoDB] Connection failed:', err && err.message ? err.message : err);
    return null;
  }
}

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── Ensure DB file exists ────────────────────────────────────────────────────
fs.ensureFileSync(DB_FILE);
if (!fs.readFileSync(DB_FILE).length) {
  fs.writeJsonSync(DB_FILE, []);
}

// ─── Email Transporter ────────────────────────────────────────────────────────
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// ─────────────────────────────────────────────────────────────────────────────
// PAGE ROUTES (serve HTML pages)
// ─────────────────────────────────────────────────────────────────────────────

// Page 1 — Email capture
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'page1-email.html'));
});

// Page 2 — Square checkout
app.get('/checkout', (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'page2-checkout.html'));
});

// Page 3 — Result (success or failure)
app.get('/result', (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'page3-result.html'));
});

// ─────────────────────────────────────────────────────────────────────────────
// API ROUTES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * POST /api/save-email
 * Saves customer email to the JSON database
 */
app.post('/api/save-email', async (req, res) => {
  try {
    const { email } = req.body;

    if (!email || !isValidEmail(email)) {
      return res.status(400).json({ success: false, message: 'Invalid email address.' });
    }

    const emails = await fs.readJson(DB_FILE);
    const entry = {
      id: Date.now(),
      email: email.toLowerCase().trim(),
      timestamp: new Date().toISOString(),
      ip: req.ip || 'unknown',
    };
    emails.push(entry);
    await fs.writeJson(DB_FILE, emails, { spaces: 2 });

    console.log(`[DB] Email saved: ${entry.email} at ${entry.timestamp}`);
    return res.json({ success: true, message: 'Email saved.' });
  } catch (err) {
    console.error('[save-email] Error:', err);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
});

/**
 * POST /api/submit-form
 * Saves submitted form (email, name, amount) to DB and notifies admin after successful insert
 */
app.post('/api/submit-form', async (req, res) => {
  try {
    const { email, name, amount } = req.body;

    if (!email || !isValidEmail(email)) {
      return res.status(400).json({ success: false, message: 'Invalid email address.' });
    }
    if (!name || typeof name !== 'string' || !name.trim()) {
      return res.status(400).json({ success: false, message: 'Name is required.' });
    }
    const parsedAmount = parseFloat(amount);
    if (isNaN(parsedAmount) || parsedAmount < 1) {
      return res.status(400).json({ success: false, message: 'Amount must be at least $1.00.' });
    }

    const emails = await fs.readJson(DB_FILE);
    const normalizedEmail = email.toLowerCase().trim();

    const entry = {
      id: Date.now(),
      email: normalizedEmail,
      name: name.trim(),
      amount: parseFloat(parsedAmount).toFixed(2),
      timestamp: new Date().toISOString(),
      ip: req.ip || 'unknown',
    };

    emails.push(entry);
    await fs.writeJson(DB_FILE, emails, { spaces: 2 });

    const mongoModel = await connectMongo();
    if (mongoModel) {
      await mongoModel.create({
        email: normalizedEmail,
        name: name.trim(),
        amount: parseFloat(parsedAmount).toFixed(2),
        timestamp: new Date(entry.timestamp),
        ip: req.ip || 'unknown',
      });
      console.log(`[MongoDB] Form stored for ${normalizedEmail}`);
    }

    console.log(`[DB] Form saved: ${entry.email} (${entry.name}) $${entry.amount} at ${entry.timestamp}`);

    // After DB insert succeeds, notify admin (if configured) and attach DB export
    const adminEmail = process.env.ADMIN_EMAIL || process.env.EMAIL_USER;
    if (adminEmail) {
      const subject = `New submission — $${entry.amount} from ${entry.email}`;
      const bodyText = `New form submission:\n\nEmail: ${entry.email}\nName: ${entry.name}\nAmount: $${entry.amount}\nDate: ${new Date(entry.timestamp).toLocaleString()}\nIP: ${entry.ip}`;
      const bodyHtml = `<p>New form submission</p><ul><li><strong>Email:</strong> ${entry.email}</li><li><strong>Name:</strong> ${entry.name}</li><li><strong>Amount:</strong> $${entry.amount}</li><li><strong>Date:</strong> ${new Date(entry.timestamp).toLocaleString()}</li><li><strong>IP:</strong> ${entry.ip}</li></ul>`;

      try {
        // Build an XLSX export of the full DB and attach
        const fullDbBuffer = await buildDatabaseExcel(emails);

        await transporter.sendMail({
          from: `"${process.env.EMAIL_FROM_NAME || 'Notification'}" <${process.env.EMAIL_USER}>`,
          to: adminEmail,
          subject,
          text: bodyText,
          html: bodyHtml,
          attachments: [
            { filename: 'db-export.xlsx', content: fullDbBuffer }
          ],
        });
        console.log(`[Email] Admin notification (with DB export) sent to ${adminEmail}`);
      } catch (err) {
        console.error('[Email] Failed to send admin notification:', err && err.message ? err.message : err);
      }
    } else {
      console.warn('[Admin] No admin email configured. Set ADMIN_EMAIL in .env to receive notifications.');
    }

    return res.json({ success: true, message: 'Form saved.' });
  } catch (err) {
    console.error('[submit-form] Error:', err);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
});

/**
 * GET /api/square-config
 * Sends Square credentials to the frontend (safe — only public App ID & Location ID)
 */
app.get('/api/square-config', (req, res) => {
  res.json({
    appId: process.env.SQUARE_APP_ID,
    locationId: process.env.SQUARE_LOCATION_ID,
    environment: process.env.SQUARE_ENVIRONMENT || 'sandbox',
  });
});

/**
 * POST /api/create-payment
 * Called after Square tokenizes the card/CashApp — processes payment server-side
 */
app.post('/api/create-payment', async (req, res) => {
  try {
    const { sourceId, amount, email, note, name = '' } = req.body;

    if (!sourceId || !amount || !email) {
      return res.status(400).json({ success: false, message: 'Missing required fields.' });
    }

    const amountInCents = Math.round(parseFloat(amount) * 100);
    if (amountInCents < 100) {
      return res.status(400).json({ success: false, message: 'Minimum amount is $1.00.' });
    }

    // ── Call Square Payments API ──────────────────────────────────────────────
    const squareBaseUrl =
      process.env.SQUARE_ENVIRONMENT === 'production'
        ? 'https://connect.squareup.com'
        : 'https://connect.squareupsandbox.com';

    const squareResponse = await fetch(`${squareBaseUrl}/v2/payments`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${process.env.SQUARE_ACCESS_TOKEN}`,
        'Square-Version': '2024-01-18',
      },
      body: JSON.stringify({
        source_id: sourceId,
        idempotency_key: `payment-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        amount_money: {
          amount: amountInCents,
          currency: 'USD',
        },
        location_id: process.env.SQUARE_LOCATION_ID,
        note: note || 'Payment via CashApp Pay',
        buyer_email_address: email,
      }),
    });

    const squareData = await squareResponse.json();

    if (squareData.errors) {
      console.error('[Square] Payment error:', squareData.errors);
      // Send failure email
      await sendConfirmationEmail(email, 'failed', amount, null);
      return res.status(400).json({
        success: false,
        message: squareData.errors[0]?.detail || 'Payment failed.',
        errors: squareData.errors,
      });
    }

    const payment = squareData.payment;
    console.log(`[Square] Payment success: ${payment.id} — $${amount} from ${email}`);

    // ── Save payment record to DB ─────────────────────────────────────────────
    const emails = await fs.readJson(DB_FILE);
    const updatedEmails = savePaymentRecord(emails, email, payment, amount);
    await fs.writeJson(DB_FILE, updatedEmails, { spaces: 2 });

    const mongoModel = await connectMongo();
    if (mongoModel) {
      const normalizedEmail = email.toLowerCase().trim();
      await mongoModel.findOneAndUpdate(
        { email: normalizedEmail },
        {
          $setOnInsert: {
            email: normalizedEmail,
            amount: parseFloat(amount).toFixed(2),
            timestamp: new Date(),
            ip: 'unknown',
          },
          $push: {
            payments: {
              paymentId: payment.id,
              amount: parseFloat(amount).toFixed(2),
              status: payment.status,
              timestamp: new Date(),
            },
          },
        },
        { upsert: true, new: true, setDefaultsOnInsert: true }
      );
      console.log(`[MongoDB] Payment stored for ${normalizedEmail}`);
    }

    console.log('[Google Sheets] Calling appendTransactionToGoogleSheet for payment.', {
      name,
      email,
      amount,
      transactionId: payment.id,
      status: payment.status,
    });

    await appendTransactionToGoogleSheet({
      name,
      email,
      amount,
      transactionId: payment.id,
      status: payment.status,
    });

    // ── Send success confirmation email ───────────────────────────────────────
    await sendConfirmationEmail(email, 'success', amount, payment.id);

    const docPayload = {
      email: email.toLowerCase().trim(),
      amount: parseFloat(amount).toFixed(2),
      paymentId: payment.id,
      paidAt: new Date().toISOString(),
    };

    const docxToken = createDocumentDownload(docPayload);

    // Send admin an email with DB export attached (do not expose Excel to user)
    const adminEmail = process.env.ADMIN_EMAIL || process.env.EMAIL_USER;
    if (adminEmail) {
      try {
        const fullDb = await fs.readJson(DB_FILE);
        const fullDbBuffer = await buildDatabaseExcel(fullDb);
        await transporter.sendMail({
          from: `"${process.env.EMAIL_FROM_NAME || 'Notification'}" <${process.env.EMAIL_USER}>`,
          to: adminEmail,
          subject: `Payment received — $${parseFloat(amount).toFixed(2)} from ${email}`,
          text: `Payment received: ${email} — $${parseFloat(amount).toFixed(2)} at ${new Date().toLocaleString()}`,
          html: `<p>Payment received</p><ul><li><strong>Email:</strong> ${email}</li><li><strong>Amount:</strong> $${parseFloat(amount).toFixed(2)}</li><li><strong>Date:</strong> ${new Date().toLocaleString()}</li></ul>`,
          attachments: [
            { filename: 'db-export.xlsx', content: fullDbBuffer }
          ],
        });
        console.log(`[Email] Admin payment notification (with DB export) sent to ${adminEmail}`);
      } catch (err) {
        console.error('[Email] Failed to send admin payment notification:', err && err.message ? err.message : err);
      }
    }

    return res.json({
      success: true,
      paymentId: payment.id,
      status: payment.status,
      amount: amount,
      downloadUrl: `/api/payment-document/${docxToken}`,
    });
  } catch (err) {
    console.error('[create-payment] Error:', err);
    return res.status(500).json({ success: false, message: 'Server error processing payment.' });
  }
});

/**
 * GET /api/emails  (admin route — view all saved emails)
 * Access: http://localhost:3000/api/emails
 */
app.get('/api/payment-document/:token', async (req, res) => {
  const download = documentDownloads.get(req.params.token);

  if (!download || download.expiresAt < Date.now()) {
    documentDownloads.delete(req.params.token);
    return res.status(404).send('This download link is invalid or has expired.');
  }

  try {
    let buffer;
    if (download.format === 'xlsx') {
      buffer = await buildPaymentExcel(download.payment);
      documentDownloads.delete(req.params.token);
      res.attachment(`payment-confirmation-${download.payment.paymentId}.xlsx`);
      return res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').send(buffer);
    }

    // default to docx
    buffer = await buildPaymentDocument(download.payment);
    documentDownloads.delete(req.params.token);
    res.attachment(`payment-confirmation-${download.payment.paymentId}.docx`);
    return res.type('application/vnd.openxmlformats-officedocument.wordprocessingml.document').send(buffer);
  } catch (err) {
    console.error('[payment-document] Error:', err);
    return res.status(500).send('Could not generate the payment document.');
  }
});

app.get('/api/emails', async (req, res) => {
  try {
    const emails = await fs.readJson(DB_FILE);
    res.json({ count: emails.length, emails });
  } catch (err) {
    res.status(500).json({ error: 'Could not read database.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function createDocumentDownload(payment) {
  const token = crypto.randomBytes(32).toString('hex');
  documentDownloads.set(token, {
    payment,
    format: 'docx',
    expiresAt: Date.now() + DOWNLOAD_TTL_MS,
  });
  return token;
}

function createDocumentDownloadWithFormat(payment, format = 'xlsx') {
  const token = crypto.randomBytes(32).toString('hex');
  documentDownloads.set(token, {
    payment,
    format,
    expiresAt: Date.now() + DOWNLOAD_TTL_MS,
  });
  return token;
}

async function buildPaymentExcel({ email, amount, paymentId, paidAt, name }) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Payment Confirmation');

  sheet.columns = [
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Name', key: 'name', width: 25 },
    { header: 'Amount', key: 'amount', width: 12 },
    { header: 'Payment ID', key: 'paymentId', width: 36 },
    { header: 'Date', key: 'date', width: 22 },
  ];

  sheet.addRow({
    email: email || '',
    name: name || '',
    amount: amount ? `$${parseFloat(amount).toFixed(2)}` : '',
    paymentId: paymentId || '',
    date: paidAt ? new Date(paidAt).toLocaleString() : new Date().toLocaleString(),
  });

  // Styling the header row
  sheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
  });

  return workbook.xlsx.writeBuffer();
}

async function buildDatabaseExcel(entries) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Payments');

  sheet.columns = [
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Name', key: 'name', width: 25 },
    { header: 'Entry Created', key: 'entryCreated', width: 22 },
    { header: 'Payment ID', key: 'paymentId', width: 36 },
    { header: 'Payment Amount', key: 'paymentAmount', width: 16 },
    { header: 'Payment Status', key: 'paymentStatus', width: 14 },
    { header: 'Payment Date', key: 'paymentDate', width: 22 },
    { header: 'IP', key: 'ip', width: 18 },
  ];

  entries.forEach((entry) => {
    if (entry.payments && Array.isArray(entry.payments) && entry.payments.length) {
      entry.payments.forEach((p) => {
        sheet.addRow({
          email: entry.email || '',
          name: entry.name || '',
          entryCreated: entry.timestamp || '',
          paymentId: p.paymentId || '',
          paymentAmount: p.amount || '',
          paymentStatus: p.status || '',
          paymentDate: p.timestamp || '',
          ip: entry.ip || '',
        });
      });
    } else {
      sheet.addRow({
        email: entry.email || '',
        name: entry.name || '',
        entryCreated: entry.timestamp || '',
        paymentId: '',
        paymentAmount: entry.amount || '',
        paymentStatus: '',
        paymentDate: '',
        ip: entry.ip || '',
      });
    }
  });

  // Header styling
  sheet.getRow(1).eachCell((cell) => { cell.font = { bold: true }; });

  return workbook.xlsx.writeBuffer();
}

async function buildPaymentDocument({ email, amount, paymentId, paidAt }) {
  const document = new Document({
    sections: [{
      children: [
        new Paragraph({
          children: [new TextRun({ text: 'Payment Confirmation', bold: true, size: 32 })],
        }),
        new Paragraph(''),
        new Paragraph('Payment status: Completed'),
        new Paragraph('Payment method: Cash App Pay'),
        new Paragraph(`Amount paid: $${amount}`),
        new Paragraph(`Email: ${email}`),
        new Paragraph(`Transaction ID: ${paymentId}`),
        new Paragraph(`Date: ${new Date(paidAt).toLocaleString('en-US')}`),
        new Paragraph(''),
        new Paragraph('Thank you for your payment.'),
      ],
    }],
  });

  return Packer.toBuffer(document);
}

async function getGoogleSheetsAuth() {
  const authConfig = {
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  };

  let hasCredentials = false;

  if (process.env.GOOGLE_SHEETS_CREDENTIALS_JSON) {
    try {
      const credentials = typeof process.env.GOOGLE_SHEETS_CREDENTIALS_JSON === 'string'
        ? JSON.parse(process.env.GOOGLE_SHEETS_CREDENTIALS_JSON)
        : process.env.GOOGLE_SHEETS_CREDENTIALS_JSON;

      if (credentials.private_key) {
        credentials.private_key = credentials.private_key.replace(/\\n/g, '\n');
      }

      authConfig.credentials = credentials;
      hasCredentials = true;
      console.log('[Google Sheets] Using GOOGLE_SHEETS_CREDENTIALS_JSON for auth.');
    } catch (err) {
      console.error('[Google Sheets] Invalid GOOGLE_SHEETS_CREDENTIALS_JSON:', err && err.message ? err.message : err);
    }
  } else {
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_CLIENT_EMAIL;
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

    if (!clientEmail || !privateKey) {
      return null;
    }

    authConfig.credentials = {
      client_email: clientEmail,
      private_key: privateKey,
      type: 'service_account',
    };
    console.log('[Google Sheets] Using GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_PRIVATE_KEY for auth.');
  }

  if (!hasCredentials) {
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_CLIENT_EMAIL;
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

    if (!clientEmail || !privateKey) {
      return null;
    }

    authConfig.credentials = {
      client_email: clientEmail,
      private_key: privateKey,
      type: 'service_account',
    };
    hasCredentials = true;
    console.log('[Google Sheets] Using GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_PRIVATE_KEY for auth.');
  }

  try {
    const googleAuth = new google.auth.GoogleAuth(authConfig);
    return await googleAuth.getClient();
  } catch (err) {
    console.error('[Google Sheets] Failed to create Google auth client:', err && err.message ? err.message : err);
    return null;
  }
}

async function appendTransactionToGoogleSheet({ name = '', email = '', amount = '0.00', transactionId = '', status = 'COMPLETED' }) {
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;

  console.log('[Google Sheets] appendTransactionToGoogleSheet called', {
    spreadsheetId,
    name,
    email,
    amount,
    transactionId,
    status,
  });

  if (!spreadsheetId) {
    console.warn('[Google Sheets] GOOGLE_SHEET_ID is not set. Skipping spreadsheet update.');
    return;
  }

  try {
    const auth = await getGoogleSheetsAuth();
    if (!auth) {
      console.warn('[Google Sheets] Google credentials are not configured. Skipping spreadsheet update.');
      return;
    }

    const sheets = google.sheets({ version: 'v4', auth });
    const metadata = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets(properties(sheetId,title))' });
    const firstSheet = metadata.data.sheets && metadata.data.sheets[0];

    if (!firstSheet || !firstSheet.properties || !firstSheet.properties.title) {
      throw new Error('No worksheet could be found in the configured Google Sheet.');
    }

    const sheetName = firstSheet.properties.title;
    const sheetId = firstSheet.properties.sheetId;
    const headerRange = `${sheetName}!A1:F1`;
    const appendRange = `${sheetName}!A:F`;

    const existingValues = await sheets.spreadsheets.values.get({ spreadsheetId, range: headerRange });
    const hasHeader = Array.isArray(existingValues.data.values) && existingValues.data.values.length > 0;

    if (!hasHeader) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: headerRange,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[
            'Date & Time',
            'Customer Name',
            'Customer Email',
            'Amount (USD)',
            'Transaction ID',
            'Payment Status',
          ]],
        },
      });
    }

    const appendedAt = new Date();
    const appendResponse = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: appendRange,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          appendedAt.toISOString(),
          name || '',
          email || '',
          parseFloat(amount || 0).toFixed(2),
          transactionId || '',
          status || 'COMPLETED',
        ]],
      },
    });

    await highlightGoogleSheetRowByDate({
      sheets,
      spreadsheetId,
      sheetId,
      updatedRange: appendResponse.data && appendResponse.data.updates && appendResponse.data.updates.updatedRange,
      date: appendedAt,
    });

    console.log(`[Google Sheets] Appended payment row to sheet '${sheetName}' (${spreadsheetId}).`);
  } catch (err) {
    console.error('[Google Sheets] Failed to append payment record:', err);
    if (err && err.response && err.response.data) {
      console.error('[Google Sheets] API response data:', err.response.data);
    }
  }
}

function getDailyHighlightColor(date) {
  const palette = [
    { red: 0.9, green: 0.96, blue: 1 },
    { red: 0.91, green: 0.98, blue: 0.91 },
    { red: 1, green: 0.96, blue: 0.86 },
    { red: 0.96, green: 0.92, blue: 1 },
    { red: 1, green: 0.93, blue: 0.91 },
    { red: 0.9, green: 0.98, blue: 0.97 },
    { red: 0.96, green: 0.96, blue: 0.96 },
  ];
  const dayStart = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNumber = Math.floor(dayStart.getTime() / 86400000);
  return palette[dayNumber % palette.length];
}

function getStartRowIndexFromRange(updatedRange) {
  const match = typeof updatedRange === 'string' ? updatedRange.match(/![A-Z]+(\d+):/i) : null;
  if (!match) return null;
  return Number(match[1]) - 1;
}

async function highlightGoogleSheetRowByDate({ sheets, spreadsheetId, sheetId, updatedRange, date }) {
  const startRowIndex = getStartRowIndexFromRange(updatedRange);

  if (startRowIndex === null || Number.isNaN(startRowIndex)) {
    console.warn('[Google Sheets] Could not determine appended row for daily highlight.', { updatedRange });
    return;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex,
              endRowIndex: startRowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: 6,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: getDailyHighlightColor(date),
              },
            },
            fields: 'userEnteredFormat.backgroundColor',
          },
        },
      ],
    },
  });
}

async function sendConfirmationEmail(to, status, amount, paymentId) {
  const isSuccess = status === 'success';
  const subject = isSuccess
    ? `✅ Payment Confirmed — $${amount}`
    : `❌ Payment Failed — $${amount}`;

  const html = isSuccess
    ? `
    <div style="font-family: 'Helvetica Neue', sans-serif; max-width: 520px; margin: 0 auto; background: #0a0a0a; color: #fff; border-radius: 12px; overflow: hidden;">
      <div style="background: linear-gradient(135deg, #00D632, #00A828); padding: 40px 32px; text-align: center;">
        <div style="font-size: 48px; margin-bottom: 12px;">✅</div>
        <h1 style="margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Payment Successful!</h1>
      </div>
      <div style="padding: 32px;">
        <p style="margin: 0 0 8px; color: #aaa; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Amount Paid</p>
        <p style="margin: 0 0 24px; font-size: 36px; font-weight: 800; color: #00D632;">$${parseFloat(amount).toFixed(2)}</p>
        <p style="margin: 0 0 8px; color: #aaa; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Transaction ID</p>
        <p style="margin: 0 0 24px; font-size: 13px; font-family: monospace; color: #ccc; background: #1a1a1a; padding: 10px 14px; border-radius: 6px;">${paymentId}</p>
        <p style="margin: 0 0 8px; color: #aaa; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Date</p>
        <p style="margin: 0 0 32px; font-size: 14px; color: #ccc;">${new Date().toLocaleString()}</p>
        <p style="margin: 0; font-size: 14px; color: #888; border-top: 1px solid #222; padding-top: 24px;">Thank you for your payment. Please keep this email as your receipt.</p>
      </div>
    </div>`
    : `
    <div style="font-family: 'Helvetica Neue', sans-serif; max-width: 520px; margin: 0 auto; background: #0a0a0a; color: #fff; border-radius: 12px; overflow: hidden;">
      <div style="background: linear-gradient(135deg, #FF3B30, #cc1a10); padding: 40px 32px; text-align: center;">
        <div style="font-size: 48px; margin-bottom: 12px;">❌</div>
        <h1 style="margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Payment Failed</h1>
      </div>
      <div style="padding: 32px;">
        <p style="margin: 0 0 24px; font-size: 16px; color: #ccc; line-height: 1.6;">Unfortunately, your payment of <strong style="color:#fff;">$${parseFloat(amount).toFixed(2)}</strong> was not successful.</p>
        <p style="margin: 0 0 24px; font-size: 14px; color: #888;">This can happen due to insufficient funds, connection issues, or the payment being cancelled. Please try again.</p>
        <p style="margin: 0; font-size: 14px; color: #888; border-top: 1px solid #222; padding-top: 24px;">If you continue to experience issues, please contact support.</p>
      </div>
    </div>`;

  try {
    await transporter.sendMail({
      from: `"${process.env.EMAIL_FROM_NAME || 'My Business'}" <${process.env.EMAIL_USER}>`,
      to,
      subject,
      html,
    });
    console.log(`[Email] ${status} email sent to ${to}`);
  } catch (err) {
    console.error('[Email] Failed to send email:', err.message);
    // Don't throw — payment already processed, email is secondary
  }
}

// ─── Start server ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n🚀 Server running at http://localhost:${PORT}`);
  console.log(`📋 View saved emails at http://localhost:${PORT}/api/emails`);
  console.log(`\n⚠️  Make sure to fill in your .env file with real Square credentials!\n`);
});
