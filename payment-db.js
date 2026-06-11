function normalizeEmail(email) {
  return String(email || '').toLowerCase().trim();
}

function buildPaymentRecord(payment, amount) {
  return {
    paymentId: payment?.id || '',
    amount: typeof amount === 'string' ? amount : String(amount ?? ''),
    status: payment?.status || 'UNKNOWN',
    timestamp: new Date().toISOString(),
  };
}

function savePaymentRecord(emails, email, payment, amount) {
  const normalizedEmail = normalizeEmail(email);
  const paymentRecord = buildPaymentRecord(payment, amount);
  const entryIndex = emails.findIndex((entry) => normalizeEmail(entry.email) === normalizedEmail);

  if (entryIndex !== -1) {
    emails[entryIndex].payments = Array.isArray(emails[entryIndex].payments) ? emails[entryIndex].payments : [];
    emails[entryIndex].payments.push(paymentRecord);
    return emails;
  }

  emails.push({
    id: Date.now(),
    email: normalizedEmail,
    amount: typeof amount === 'string' ? amount : String(amount ?? ''),
    timestamp: new Date().toISOString(),
    payments: [paymentRecord],
  });

  return emails;
}

module.exports = {
  normalizeEmail,
  buildPaymentRecord,
  savePaymentRecord,
};
