"use strict";

const crypto = require("crypto");
const { query, initSchema } = require("../db");
require("dotenv").config();

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return { salt, hash: `${salt}:${hash}` };
}

async function seedAdmin() {
  await initSchema();
  const { hash } = hashPassword(process.env.ADMIN_PASSWORD || "ADMIN123");
  await query(
    `INSERT INTO app_users (username, name, role, password_hash)
     VALUES ($1, $2, $3, $4)
     ON CONFLICT (username) DO UPDATE
       SET password_hash = EXCLUDED.password_hash, role = 'admin', name = EXCLUDED.name`,
    ["ADMIN", "Administrator", "admin", hash]
  );
  console.log("[seed] ADMIN user ensured (username=ADMIN).");
  process.exit(0);
}

seedAdmin().catch((err) => {
  console.error("[seed] Failed:", err.message);
  process.exit(1);
});