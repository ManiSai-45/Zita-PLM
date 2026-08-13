"use strict";

const nodemailer = require("nodemailer");

let transporter = null;

function getTransporter() {
  if (transporter) return transporter;
  transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || "smtp.gmail.com",
    port: Number(process.env.SMTP_PORT) || 465,
    secure: true,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
    tls: { rejectUnauthorized: false },
  });
  return transporter;
}

function isEnabled() {
  return process.env.EMAIL_ENABLED === "true" && !!process.env.SMTP_USER;
}

function appUrl() {
  return process.env.APP_URL || "http://localhost:" + (process.env.PORT || 3000);
}

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function send({ to, subject, html }) {
  if (!isEnabled()) {
    console.log(`[mailer] disabled — skipping "${subject}" -> ${to || "(no recipient)"}`);
    return Promise.resolve(false);
  }
  if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
    console.warn(`[mailer] no valid recipient for "${subject}"`);
    return Promise.resolve(false);
  }
  const fromAddr = process.env.EMAIL_FROM || process.env.SMTP_USER;
  const from = process.env.EMAIL_NAME ? `${process.env.EMAIL_NAME} <${fromAddr}>` : `ZITA PLM <${fromAddr}>`;
  return getTransporter()
    .sendMail({ from, to, subject, html })
    .then((info) => {
      console.log(`[mailer] sent "${subject}" -> ${to} (${info.messageId})`);
      return true;
    });
}

function wrap(inner) {
  const base = appUrl();
  return `<div style="max-width:520px;margin:0 auto;padding:28px 24px;font-family:Inter,Arial,Helvetica,sans-serif;background:#ffffff">
      <div style="text-align:center;margin-bottom:18px">
        <span style="font-size:20px;font-weight:800;color:#4338ca;letter-spacing:.5px">ZITA PLM</span>
      </div>
      <div style="background:#f8f9fd;border:1px solid #e5e9f2;border-radius:14px;padding:22px">
        ${inner}
      </div>
      <p style="font-size:11.5px;color:#94a3b8;text-align:center;margin-top:18px">This is an automated message from ZITA PLM (${base}).</p>
    </div>`;
}

module.exports = { send, appUrl, esc, wrap };