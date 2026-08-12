# ZITA PLM

Internal Project Management & Timesheet tool.

This repository contains a Node.js + Express backend and a vanilla-JS frontend for a lightweight PLM application.

Quick start

1. Install dependencies:

   ```bash
   npm install
   ```

2. Configure your `.env` with your Postgres/Supabase settings (do not commit `.env`).

3. Initialize the database (seeds admin user):

   ```bash
   npm run init-db
   ```

4. Run locally:

   ```bash
   npm start
   ```

Open http://localhost:3000 and sign in with `ADMIN` / `ADMIN123`.

Note: `.env` contains secrets and is git-ignored. Make sure `TOKEN_SECRET` is set before deploying.

---
**Repository note:** uploaded from a local workspace on 2026-08-13. Branch `readme-update` will contain this README cleanup and a PR will be opened.
