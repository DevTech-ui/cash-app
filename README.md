# 💳 Square Cash App Pay — 3-Page Checkout Website

A complete payment website with email capture, Square Cash App Pay checkout, and result page.

---

## 📁 Project Structure

```
square-checkout/
├── server.js               ← Main backend (Node.js + Express)
├── package.json
├── .env                    ← YOUR CREDENTIALS GO HERE
├── data/
│   └── emails.json         ← Auto-created database of customer emails
├── views/
│   ├── page1-email.html    ← Page 1: Email capture
│   ├── page2-checkout.html ← Page 2: Cash App Pay checkout
│   └── page3-result.html   ← Page 3: Success / Failure result
└── public/                 ← Static assets (CSS, JS, images)
```

---

## 🚀 Setup Instructions (Step by Step)

### Step 1 — Install Node.js
Download from: https://nodejs.org (choose LTS version)

### Step 2 — Install dependencies
Open a terminal in this folder and run:
```bash
npm install
```

### Step 3 — Get your Square credentials

1. Go to https://developer.squareup.com/apps
2. Create a new app (or use an existing one)
3. Go to **Credentials** tab → copy:
   - **Application ID** (starts with `sq0idb-...`)
   - **Access Token** (starts with `EAAAl...`)
4. Go to **Locations** tab → copy your **Location ID**
5. In your Square Dashboard → go to **Cash App Pay** and **enable** it for your location

### Step 4 — Enable Cash App Pay
- In Square Dashboard: `Settings → Payment Methods → Cash App Pay → Enable`
- Make sure your account is approved for Cash App Pay

### Step 5 — Set up email (Gmail)

To send confirmation emails:
1. Go to your Google Account → Security → **App Passwords**
2. Create an App Password for "Mail"
3. Copy the 16-character password

### Step 6 — Fill in your .env file

Open `.env` and replace ALL placeholder values:

```env
SQUARE_APP_ID=sandbox-sq0idb-YOUR_ACTUAL_APP_ID
SQUARE_LOCATION_ID=YOUR_ACTUAL_LOCATION_ID
SQUARE_ACCESS_TOKEN=EAAAl_YOUR_ACTUAL_TOKEN
SQUARE_ENVIRONMENT=sandbox       ← change to "production" when ready for real money

EMAIL_USER=yourgmail@gmail.com
EMAIL_PASS=your16charapppassword
EMAIL_FROM_NAME=Your Business Name

PORT=3000
BASE_URL=http://localhost:3000    ← change to your domain when deployed
```

### Step 7 — Start the server
```bash
npm start
```

Then open your browser to: **http://localhost:3000**

---

## 🧪 Testing with Sandbox

When `SQUARE_ENVIRONMENT=sandbox`:
- Use the Square sandbox Cash App Pay — it will show a test flow
- No real money is charged
- Sandbox App ID starts with `sandbox-sq0idb-`

When ready for real payments:
1. Change `SQUARE_ENVIRONMENT=production`
2. Use your **production** App ID and Access Token (not sandbox)

---

## 📊 Viewing Saved Emails (Your Database)

All customer emails are saved to `data/emails.json`.

**Option 1 — Browser:**
Visit: http://localhost:3000/api/emails

**Option 2 — Terminal:**
```bash
cat data/emails.json
```

The file looks like:
```json
[
  {
    "id": 1704067200000,
    "email": "customer@example.com",
    "timestamp": "2024-01-01T12:00:00.000Z",
    "ip": "::1",
    "payments": [
      {
        "paymentId": "abc123XYZ",
        "amount": "50.00",
        "status": "COMPLETED",
        "timestamp": "2024-01-01T12:05:00.000Z"
      }
    ]
  }
]
```

---

## 📱 Mobile vs Desktop Behavior

Square's Cash App Pay SDK automatically handles this:
- **Mobile** → Opens Cash App directly (deep link)
- **Desktop** → Shows a QR code to scan with your phone

---

## 🌐 Deploying to the Internet

### Option A — Railway (easiest, free tier)
1. Push this folder to GitHub
2. Go to https://railway.app
3. Connect your GitHub repo
4. Add environment variables (same as .env)
5. Deploy — Railway gives you a public URL

### Option B — Render.com (free tier)
1. Push to GitHub
2. https://render.com → New Web Service
3. Build command: `npm install`
4. Start command: `node server.js`
5. Add environment variables

### Option C — Your own VPS (DigitalOcean, etc.)
```bash
# On your server:
git clone your-repo
cd square-checkout
npm install
# Set environment variables
npm start
# Use PM2 to keep it running:
npm install -g pm2
pm2 start server.js
```

After deploying, update `.env`:
```
BASE_URL=https://yourdomain.com
```

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| Cash App Pay button doesn't appear | Check Square App ID & Location ID in .env; make sure Cash App Pay is enabled in dashboard |
| "Payment SDK failed to load" | Make sure `SQUARE_ENVIRONMENT` is exactly `sandbox` or `production` |
| Email not sending | Check Gmail App Password (not regular password); enable 2FA on Google first |
| Port already in use | Change `PORT=3001` in .env |
| Payment works but email fails | Check EMAIL_USER and EMAIL_PASS in .env; emails are sent but failures won't stop payments |

---

## 🔒 Security Notes

- `.env` is in `.gitignore` — never commit it to GitHub
- The `/api/emails` endpoint has no password — add authentication before going live
- For production, use HTTPS (Railway/Render provide this automatically)

---

## 📞 Support

If you need to add features:
- Password-protect the emails endpoint
- Add a product description field
- Support multiple payment amounts with dropdown
- Add webhook for real-time payment notifications from Square
