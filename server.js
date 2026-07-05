require('dotenv').config();
const crypto = require('crypto');
const dns = require('dns');
const https = require('https');
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

dns.setDefaultResultOrder('ipv4first');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data', 'emails.json');
const documentDownloads = new Map();
const taptapupPayments = new Map();
const DOWNLOAD_TTL_MS = 15 * 60 * 1000;
const MIN_PAYMENT_AMOUNT = Number(process.env.TAPTAPUP_MIN_AMOUNT || 20);

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
  res.sendFile(path.join(__dirname, 'views', 'page1-email.html'));
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
    if (isNaN(parsedAmount) || parsedAmount < MIN_PAYMENT_AMOUNT) {
      return res.status(400).json({ success: false, message: `Amount must be at least $${MIN_PAYMENT_AMOUNT.toFixed(2)}.` });
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
    return res.json({ success: true, message: 'Form saved.' });

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
    const { amount, email, name = '' } = req.body;

    if (!amount || !email) {
      return res.status(400).json({ success: false, message: 'Missing required fields.' });
    }
    if (!isValidEmail(email)) {
      return res.status(400).json({ success: false, message: 'Invalid email address.' });
    }

    const parsedAmount = parseFloat(amount);
    if (isNaN(parsedAmount) || parsedAmount < MIN_PAYMENT_AMOUNT) {
      return res.status(400).json({ success: false, message: `Minimum amount is $${MIN_PAYMENT_AMOUNT.toFixed(2)}.` });
    }

    // Start a TapTapUp hosted payment session.
    const merchantReference = `INV-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const baseUrl = getBaseUrl(req);
    const taptapupData = await callTapTapUp('/initiate-payment', 'POST', {
      product_id: getTapTapUpProductId(),
      amount: Number(parsedAmount.toFixed(2)),
      email: email.toLowerCase().trim(),
      merchant_reference: merchantReference,
      return_url: `${baseUrl}/result?ref=${encodeURIComponent(merchantReference)}`,
      webhook_url: `${baseUrl}/api/taptapup-webhook`,
    });
    const paymentSession = unwrapTapTapUpData(taptapupData);

    if (!taptapupData.success || !paymentSession.redirect_url || !paymentSession.token) {
      console.error('[TapTapUp] Payment initiation failed:', taptapupData);
      await sendConfirmationEmail(email, 'failed', parsedAmount.toFixed(2), null);
      return res.status(400).json({
        success: false,
        message: taptapupData.message || 'Could not start TapTapUp payment.',
        details: taptapupData,
      });
    }

    taptapupPayments.set(merchantReference, {
      token: paymentSession.token,
      name,
      email: email.toLowerCase().trim(),
      amount: parsedAmount.toFixed(2),
      merchantReference,
      createdAt: Date.now(),
    });

    console.log(`[TapTapUp] Payment initiated: ${merchantReference} - $${parsedAmount.toFixed(2)} from ${email}`);

    return res.json({
      success: true,
      redirectUrl: paymentSession.redirect_url,
      merchantReference,
      expiresIn: paymentSession.expires_in,
      requestedAmount: paymentSession.requested_amount,
      finalAmount: paymentSession.final_amount,
    });

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
    return res.status(500).json({ success: false, message: err.message || 'Server error processing payment.' });
  }
});

/**
 * GET /api/emails  (admin route — view all saved emails)
 * Access: http://localhost:3000/api/emails
 */
app.get('/api/payment-status', async (req, res) => {
  try {
    const merchantReference = String(req.query.ref || '').trim();
    if (!merchantReference) {
      return res.status(400).json({ success: false, message: 'Missing payment reference.' });
    }

    const pendingPayment = taptapupPayments.get(merchantReference);
    if (!pendingPayment) {
      return res.status(404).json({ success: false, message: 'Payment reference was not found or has expired.' });
    }

    const statusData = await callTapTapUp(`/status/${encodeURIComponent(pendingPayment.token)}`, 'GET');
    const paymentStatus = unwrapTapTapUpData(statusData);
    if (!statusData.success) {
      return res.status(400).json({ success: false, message: statusData.message || 'Could not verify payment status.', details: statusData });
    }

    if (paymentStatus.status === 'completed' && !pendingPayment.completed) {
      const result = await finalizeTapTapUpPayment({
        ...pendingPayment,
        orderId: paymentStatus.order_id,
        status: paymentStatus.status,
      });
      pendingPayment.completed = true;
      pendingPayment.orderId = paymentStatus.order_id;
      pendingPayment.downloadUrl = result.downloadUrl;
    }

    return res.json({
      success: true,
      status: paymentStatus.status,
      paymentId: String(paymentStatus.order_id || pendingPayment.orderId || merchantReference),
      amount: pendingPayment.amount,
      email: pendingPayment.email,
      downloadUrl: pendingPayment.downloadUrl || '',
    });
  } catch (err) {
    console.error('[payment-status] Error:', err);
    return res.status(500).json({ success: false, message: err.message || 'Could not verify payment status.' });
  }
});

app.post('/api/taptapup-webhook', async (req, res) => {
  try {
    const payload = req.body || {};
    const merchantReference = payload.merchant_reference || req.get('X-Merchant-Reference');

    if (payload.event !== 'payment.completed' || payload.status !== 'completed') {
      return res.status(202).json({ success: true });
    }

    const pendingPayment = merchantReference ? taptapupPayments.get(merchantReference) : null;
    const result = await finalizeTapTapUpPayment({
      merchantReference,
      name: pendingPayment?.name || '',
      email: payload.email || pendingPayment?.email || '',
      amount: String(payload.amount || pendingPayment?.amount || '0.00'),
      orderId: payload.order_id || payload.order_number || merchantReference,
      status: payload.status,
    });

    if (pendingPayment) {
      pendingPayment.completed = true;
      pendingPayment.orderId = payload.order_id || payload.order_number;
      pendingPayment.downloadUrl = result.downloadUrl;
    }

    return res.json({ success: true });
  } catch (err) {
    console.error('[taptapup-webhook] Error:', err);
    return res.status(500).json({ success: false });
  }
});

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

app.get('/api/taptapup-diagnostics', async (req, res) => {
  const { baseUrls } = getTapTapUpConfig();
  const diagnostics = [];

  for (const baseUrl of baseUrls) {
    const hostname = new URL(baseUrl).hostname;
    const result = { baseUrl, hostname };

    try {
      result.addresses = await dns.promises.lookup(hostname, { all: true });
    } catch (err) {
      result.dnsError = { code: err.code, message: err.message };
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    try {
      const startedAt = Date.now();
      const response = await fetch(`${baseUrl}/status/diagnostic`, {
        method: 'GET',
        signal: controller.signal,
      });
      result.http = {
        ok: response.ok,
        status: response.status,
        elapsedMs: Date.now() - startedAt,
      };
    } catch (err) {
      result.httpError = {
        code: err?.cause?.code || err?.code || err.name,
        message: err.message,
      };
    } finally {
      clearTimeout(timeout);
    }

    diagnostics.push(result);
  }

  res.json({ success: true, diagnostics });
});

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getBaseUrl(req) {
  return (process.env.BASE_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/, '');
}

function getTapTapUpProductId() {
  const productId = Number(process.env.TAPTAPUP_PRODUCT_ID);
  if (!Number.isInteger(productId) || productId <= 0) {
    throw new Error('TAPTAPUP_PRODUCT_ID must be set to your TapTapUp product ID.');
  }
  return productId;
}

function getTapTapUpConfig() {
  const merchantId = process.env.TAPTAPUP_MERCHANT_ID;
  const sharedSecret = process.env.TAPTAPUP_SHARED_SECRET;

  if (!merchantId || !sharedSecret) {
    throw new Error('TapTapUp credentials are missing. Set TAPTAPUP_MERCHANT_ID and TAPTAPUP_SHARED_SECRET.');
  }

  return {
    merchantId,
    sharedSecret,
    baseUrls: getTapTapUpBaseUrls(process.env.TAPTAPUP_API_BASE_URL || 'https://taptapup.xyz/api/v1'),
  };
}

function normalizeTapTapUpBaseUrl(rawBaseUrl) {
  const withProtocol = /^https?:\/\//i.test(rawBaseUrl) ? rawBaseUrl : `https://${rawBaseUrl}`;
  const parsedUrl = new URL(withProtocol.replace(/\/$/, ''));
  return parsedUrl.toString().replace(/\/$/, '');
}

function getTapTapUpBaseUrls(rawBaseUrl) {
  const primary = normalizeTapTapUpBaseUrl(rawBaseUrl);
  const parsedPrimary = new URL(primary);
  const hostWithoutWww = parsedPrimary.hostname.replace(/^www\./i, '');
  const hostWithWww = hostWithoutWww.startsWith('www.') ? hostWithoutWww : `www.${hostWithoutWww}`;

  const candidates = [primary];

  for (const hostname of [hostWithoutWww, hostWithWww]) {
    const candidate = new URL(primary);
    candidate.hostname = hostname;
    candidates.push(candidate.toString().replace(/\/$/, ''));
  }

  return [...new Set(candidates)];
}

async function callTapTapUp(endpoint, method = 'GET', payload) {
  const { merchantId, sharedSecret, baseUrls } = getTapTapUpConfig();
  const timestamp = Math.floor(Date.now() / 1000).toString();
  const rawBody = payload ? JSON.stringify(payload) : '';
  const signature = crypto
    .createHmac('sha256', sharedSecret)
    .update(`${timestamp}${rawBody}`)
    .digest('hex');

  let response;
  const failures = [];

  for (const baseUrl of baseUrls) {
    try {
      response = await requestJsonOverHttps(`${baseUrl}${endpoint}`, {
        method,
        headers: {
          'Content-Type': 'application/json',
          'X-Merchant-ID': merchantId,
          'X-Timestamp': timestamp,
          'X-Signature': signature,
        },
        body: payload ? rawBody : '',
        timeoutMs: Number(process.env.TAPTAPUP_TIMEOUT_MS || 12000),
      });
      break;
    } catch (err) {
      const failure = {
        url: `${baseUrl}${endpoint}`,
        code: err?.code || err.name,
        message: err?.message,
      };
      failures.push(failure);
      console.error('[TapTapUp] API request failed:', failure);
    }
  }

  if (!response) {
    throw new Error(`Could not reach TapTapUp API. Tried: ${failures.map((failure) => failure.url).join(', ')}. Check Render outbound connectivity and TapTapUp IP allowlisting.`);
  }

  const data = response.data;

  if (!response.ok && data.success !== false) {
    data.success = false;
    data.message = data.message || `TapTapUp request failed with HTTP ${response.status}.`;
  }

  return data;
}

function requestJsonOverHttps(url, { method = 'GET', headers = {}, body = '', timeoutMs = 12000 } = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const requestBody = body || '';
    const requestHeaders = {
      ...headers,
      Accept: 'application/json',
    };

    if (requestBody) {
      requestHeaders['Content-Length'] = Buffer.byteLength(requestBody);
    }

    const request = https.request({
      protocol: parsedUrl.protocol,
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || 443,
      path: `${parsedUrl.pathname}${parsedUrl.search}`,
      method,
      headers: requestHeaders,
      family: 4,
      timeout: timeoutMs,
    }, (response) => {
      let responseText = '';

      response.setEncoding('utf8');
      response.on('data', (chunk) => {
        responseText += chunk;
      });
      response.on('end', () => {
        let data;
        try {
          data = responseText ? JSON.parse(responseText) : {};
        } catch (err) {
          data = { success: false, message: 'TapTapUp returned a non-JSON response.', raw: responseText.slice(0, 500) };
        }

        resolve({
          ok: response.statusCode >= 200 && response.statusCode < 300,
          status: response.statusCode,
          data,
        });
      });
    });

    request.on('timeout', () => {
      request.destroy(Object.assign(new Error(`TapTapUp request timed out after ${timeoutMs}ms`), { code: 'ETIMEDOUT' }));
    });
    request.on('error', reject);

    if (requestBody) {
      request.write(requestBody);
    }
    request.end();
  });
}

function unwrapTapTapUpData(response) {
  return response && typeof response.data === 'object' && response.data !== null
    ? response.data
    : response || {};
}

async function finalizeTapTapUpPayment({ name = '', email = '', amount = '0.00', orderId = '', status = 'completed' }) {
  const normalizedEmail = email.toLowerCase().trim();
  const paymentId = String(orderId || `taptapup-${Date.now()}`);
  const paymentStatus = String(status || 'completed').toUpperCase();
  const payment = { id: paymentId, status: paymentStatus };

  const emails = await fs.readJson(DB_FILE);
  const updatedEmails = savePaymentRecord(emails, normalizedEmail, payment, amount);
  await fs.writeJson(DB_FILE, updatedEmails, { spaces: 2 });

  const mongoModel = await connectMongo();
  if (mongoModel) {
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
            paymentId,
            amount: parseFloat(amount).toFixed(2),
            status: paymentStatus,
            timestamp: new Date(),
          },
        },
      },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    );
  }

  await appendTransactionToGoogleSheet({
    name,
    email: normalizedEmail,
    amount,
    transactionId: paymentId,
    status: paymentStatus,
  });

  await sendConfirmationEmail(normalizedEmail, 'success', amount, paymentId);

  const downloadUrl = `/api/payment-document/${createDocumentDownload({
    email: normalizedEmail,
    amount: parseFloat(amount).toFixed(2),
    paymentId,
    paidAt: new Date().toISOString(),
  })}`;

  await sendAdminPaymentNotification({ email: normalizedEmail, amount, paymentId });
  return { paymentId, downloadUrl };
}

async function sendAdminPaymentNotification({ email, amount, paymentId }) {
  const adminEmail = process.env.ADMIN_EMAIL || process.env.EMAIL_USER;
  if (!adminEmail) return;

  try {
    const fullDb = await fs.readJson(DB_FILE);
    const fullDbBuffer = await buildDatabaseExcel(fullDb);
    await transporter.sendMail({
      from: `"${process.env.EMAIL_FROM_NAME || 'Notification'}" <${process.env.EMAIL_USER}>`,
      to: adminEmail,
      subject: `Payment received - $${parseFloat(amount).toFixed(2)} from ${email}`,
      text: `Payment received: ${email} - $${parseFloat(amount).toFixed(2)} at ${new Date().toLocaleString()}`,
      html: `<p>Payment received</p><ul><li><strong>Email:</strong> ${email}</li><li><strong>Amount:</strong> $${parseFloat(amount).toFixed(2)}</li><li><strong>Transaction ID:</strong> ${paymentId}</li><li><strong>Date:</strong> ${new Date().toLocaleString()}</li></ul>`,
      attachments: [
        { filename: 'db-export.xlsx', content: fullDbBuffer }
      ],
    });
  } catch (err) {
    console.error('[Email] Failed to send admin payment notification:', err && err.message ? err.message : err);
  }
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
