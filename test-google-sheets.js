require('dotenv').config();
const { google } = require('googleapis');

async function getAuth() {
  const authConfig = {
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  };

  let hasCredentials = false;

  if (process.env.GOOGLE_SHEETS_CREDENTIALS_JSON) {
    try {
      const credentials = JSON.parse(process.env.GOOGLE_SHEETS_CREDENTIALS_JSON);
      if (credentials.private_key) {
        credentials.private_key = credentials.private_key.replace(/\\n/g, '\n');
      }
      authConfig.credentials = credentials;
      hasCredentials = true;
      console.log('[Google Sheets Test] Using GOOGLE_SHEETS_CREDENTIALS_JSON.');
    } catch (err) {
      console.warn('[Google Sheets Test] Ignoring invalid GOOGLE_SHEETS_CREDENTIALS_JSON:', err.message);
    }
  }

  if (!hasCredentials) {
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_CLIENT_EMAIL;
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

    if (!clientEmail || !privateKey) {
      throw new Error('Missing Google service account email or private key.');
    }

    authConfig.credentials = {
      client_email: clientEmail,
      private_key: privateKey,
      type: 'service_account',
    };
    console.log('[Google Sheets Test] Using service account email/private key.');
  }

  return new google.auth.GoogleAuth(authConfig).getClient();
}

async function main() {
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;
  if (!spreadsheetId) {
    throw new Error('Missing GOOGLE_SHEET_ID.');
  }

  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(sheetId,title))',
  });

  const firstSheet = metadata.data.sheets && metadata.data.sheets[0];
  if (!firstSheet || !firstSheet.properties || !firstSheet.properties.title) {
    throw new Error('No worksheet found in the configured Google Sheet.');
  }

  const sheetName = firstSheet.properties.title;
  const sheetId = firstSheet.properties.sheetId;
  const headerRange = `${sheetName}!A1:F1`;
  const appendRange = `${sheetName}!A:F`;
  const existingValues = await sheets.spreadsheets.values.get({ spreadsheetId, range: headerRange });

  if (!existingValues.data.values || existingValues.data.values.length === 0) {
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
        'Local Test',
        'local-test@example.com',
        '1.00',
        `local-test-${Date.now()}`,
        'TEST',
      ]],
    },
  });

  await highlightRowByDate({
    sheets,
    spreadsheetId,
    sheetId,
    updatedRange: appendResponse.data && appendResponse.data.updates && appendResponse.data.updates.updatedRange,
    date: appendedAt,
  });

  console.log(`[Google Sheets Test] Appended and highlighted test row in "${sheetName}".`);
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

async function highlightRowByDate({ sheets, spreadsheetId, sheetId, updatedRange, date }) {
  const startRowIndex = getStartRowIndexFromRange(updatedRange);

  if (startRowIndex === null || Number.isNaN(startRowIndex)) {
    throw new Error(`Could not determine appended row from range: ${updatedRange}`);
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

main().catch((err) => {
  console.error('[Google Sheets Test] Failed:', err.message);
  process.exit(1);
});
