require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs-extra');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data', 'emails.json');

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
    const { sourceId, amount, email, note } = req.body;

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
    const entryIndex = emails.findIndex(
      (e) => e.email === email.toLowerCase().trim()
    );
    const paymentRecord = {
      paymentId: payment.id,
      amount: amount,
      status: payment.status,
      timestamp: new Date().toISOString(),
    };
    if (entryIndex !== -1) {
      emails[entryIndex].payments = emails[entryIndex].payments || [];
      emails[entryIndex].payments.push(paymentRecord);
      await fs.writeJson(DB_FILE, emails, { spaces: 2 });
    }

    // ── Send success confirmation email ───────────────────────────────────────
    await sendConfirmationEmail(email, 'success', amount, payment.id);

    return res.json({
      success: true,
      paymentId: payment.id,
      status: payment.status,
      amount: amount,
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
