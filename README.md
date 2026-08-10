# ZITA PLM

Internal Project Management & Timesheet tool.

- **Backend:** Node.js + Express REST API connected to Supabase Postgres
- **Frontend:** Vanilla JS single-page app (no build step)
- **Auth:** username/password, tokens signed with HMAC; passwords hashed with scrypt

## Features

- **Admin role** (`ADMIN` / `ADMIN123`) — create users, delete users, reset passwords, assign tasks, full edit on every task, view everyone's timesheets (monitor mode), see everything.
- **User role** — can create tasks (auto or manual task ID) and assign them to a responsible person. On their own/assigned tasks, users can **only change status and add comments**; they cannot edit title, description, priority, assignee, due date, task type, etc. Other members' tasks are read-only.
- **Tasks tab** with two views:
  - **My work** — kanban board of tasks assigned to / created by you. Open a card → change status (card auto-moves) or add a comment.
  - **Overall** — for users this is **read-only**; admins get full edit. Filters: search, status, priority, task type, assignee, due-date range.
- **Task columns** — Task ID (auto next number, editable), Title, Task Type, Client, Status, % Complete, Priority, Assignee, Assigned By/Date (auto-populated), Due Date, Duration (auto), Dependencies, Risk/Blockers, Comments (thread with author+date stamps).
- **Excel export** — exports include every task column; timesheets and members also exportable from each screen.
- **Timesheets tab** — members log daily hours (task, description, date, hours, task type); **admins monitor all members** (no entry form, no logging).
- **Dashboard** — stat cards + charts (tasks by status, hours by member).
- **Team** — member directory; admins manage accounts.

## Set up

1. Install dependencies:

   ```
   npm install
   ```

2. Configure the database in `.env` (already filled with your Supabase connection):

   ```
   PGHOST, PGPORT, PGDATABASE, PGUSER, PGPASSWORD
   ```

3. Create the tables and seed the `ADMIN` user (idempotent, safe to rerun):

   ```
   npm run init-db
   ```

## Run

```
npm start
```

Open **http://localhost:3000** and sign in with `ADMIN` / `ADMIN123`.

> Tip: log in as ADMIN → **Team & Users** → create accounts for each member (e.g. GANESH, KISHORE, MANISAI) and share their username + password.

## API overview

| Method | Path | Access | Purpose |
| --- | --- | --- | --- |
| POST | `/api/login` | public | sign in, returns token + user |
| GET | `/api/me` | any | current user |
| GET | `/api/users` | any | list members |
| POST | `/api/users` | admin | create user |
| PATCH | `/api/users/:username` | admin | rename / change role / reset password |
| DELETE | `/api/users/:username` | admin | delete user |
| GET | `/api/tasks/next-code` | any | next available task ID |
| GET | `/api/tasks?status=&assigned_to=&domain=&priority=&from=&to=&search=` | any | list/filter tasks |
| POST | `/api/tasks` | any | create + assign task |
| PATCH | `/api/tasks/:id` | any (owner/assignee/admin) | edit task / completion status |
| DELETE | `/api/tasks/:id` | admin or creator | delete task |
| GET | `/api/timesheets?username=&from=&to=&domain=&search=` | any | non-admins see own only |
| POST/PATCH/DELETE | `/api/timesheets[/:id]` | any (owner/admin) | log / edit / delete hours |
| GET | `/api/summary` | any | dashboard stats |

## Security notes

- `.env` contains real database credentials and is git-ignored — do not commit or share it.
- Change `TOKEN_SECRET` in `.env` to a long random string before any shared deployment.
- Change the default ADMIN password after first sign-in (Team & Users → gear icon).
- Permissions are enforced server-side; the UI only shows controls the user is allowed to use.

## Deployment (optional)

The app is a standard Node/Express process — deploy it anywhere Node runs (Render, Railway, Fly.io, a VM) and point `PGHOST` etc. back at the same Supabase instance. The `public/` folder is served automatically; a reverse proxy (e.g. Caddy/nginx) can add HTTPS.