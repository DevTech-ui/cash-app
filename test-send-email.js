require('dotenv').config();
const nodemailer = require('nodemailer');

async function main(){
  const user = process.env.EMAIL_USER;
  const pass = process.env.EMAIL_PASS;
  const admin = process.env.ADMIN_EMAIL || user;

  if (!user || !pass){
    console.error('Missing EMAIL_USER or EMAIL_PASS in .env');
    process.exit(2);
  }

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user, pass },
  });

  try {
    const info = await transporter.sendMail({
      from: `"Test" <${user}>`,
      to: admin,
      subject: 'Test email from cash-app',
      text: `This is a test email sent at ${new Date().toISOString()}`,
    });

    console.log('Email sent:', info && info.accepted ? info.accepted : info);
  } catch (err) {
    console.error('Send failed:', err && err.message ? err.message : err);
    if (err && err.response) console.error('SMTP response:', err.response);
    process.exit(1);
  }
}

main();
