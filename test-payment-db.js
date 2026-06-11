const assert = require('node:assert/strict');
const { savePaymentRecord } = require('./payment-db');

const existing = [
  {
    id: 1,
    email: 'user@example.com',
    name: 'User',
    amount: '10.00',
    timestamp: '2026-06-09T00:00:00.000Z',
    ip: '::1',
  },
];

const updated = savePaymentRecord(existing, 'USER@Example.com', {
  id: 'pay_123',
  status: 'COMPLETED',
}, '10.00');

assert.equal(updated.length, 1);
assert.equal(updated[0].email, 'user@example.com');
assert.ok(Array.isArray(updated[0].payments));
assert.equal(updated[0].payments.length, 1);
assert.equal(updated[0].payments[0].paymentId, 'pay_123');
assert.equal(updated[0].payments[0].amount, '10.00');
assert.equal(updated[0].payments[0].status, 'COMPLETED');

const created = savePaymentRecord([], 'new@example.com', { id: 'pay_999', status: 'COMPLETED' }, '25.50');
assert.equal(created.length, 1);
assert.equal(created[0].email, 'new@example.com');
assert.equal(created[0].payments[0].paymentId, 'pay_999');
console.log('payment-db regression check passed');
