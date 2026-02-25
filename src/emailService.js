const nodemailer = require('nodemailer');
const config = require('./config');

function parseRecipients(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.map((email) => String(email).trim()).filter(Boolean);
  }
  return String(value)
    .split(',')
    .map((email) => email.trim())
    .filter(Boolean);
}

function buildTransporter() {
  if (!config.email.enabled) {
    throw new Error('Email is disabled. Set SMTP_ENABLED=true in .env');
  }

  if (!config.email.host || !config.email.user || !config.email.pass) {
    throw new Error('SMTP credentials are incomplete. Check SMTP_HOST/SMTP_USER/SMTP_PASS in .env');
  }

  return nodemailer.createTransport({
    host: config.email.host,
    port: config.email.port,
    secure: config.email.secure,
    auth: {
      user: config.email.user,
      pass: config.email.pass
    }
  });
}

async function sendMomEmail({ to, cc, subject, body, attachmentPath, attachmentName }) {
  const toRecipients = parseRecipients(to);
  const ccRecipients = parseRecipients(cc);

  if (!toRecipients.length) {
    throw new Error('At least one recipient is required to send email.');
  }

  const transporter = buildTransporter();

  await transporter.sendMail({
    from: config.email.from,
    to: toRecipients.join(', '),
    cc: ccRecipients.length ? ccRecipients.join(', ') : undefined,
    subject,
    html: body,
    attachments: [
      {
        filename: attachmentName,
        path: attachmentPath,
        contentType: 'application/pdf'
      }
    ]
  });
}

module.exports = {
  sendMomEmail
};
