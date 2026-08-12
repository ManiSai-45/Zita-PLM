/* ============================= PLM Suite ============================= */
(() => {
  "use strict";

  const API = "";
  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

  /* ---------------- state ---------------- */
  const state = {
    user: null,
    users: [],
    tasks: [],
    timesheets: [],
    filters: {},
    tsFilters: {},
    taskView: "mine", // 'mine' | 'all'
    charts: {},
  };

  /* ---------------- helpers ---------------- */
  const STATUS = ["Pending", "In Progress", "On Hold", "Blocked", "Completed"];
  const PRIORITY = ["Low", "Medium", "High", "Critical"];

  const statusColor = (s) => {
    const map = {
      Pending: ["#6366f1", "#e0e7ff"],
      "In Progress": ["#0ea5e9", "#e0f2fe"],
      "On Hold": ["#f59e0b", "#fef3c7"],
      Blocked: ["#ef4444", "#fee2e2"],
      Completed: ["#10b981", "#d1fae5"],
    };
    return map[s] || ["#64748b", "#e2e8f0"];
  };
  const priorityColor = (p) => {
    const map = {
      Low: ["#10b981", "#d1fae5"],
      Medium: ["#f59e0b", "#fef3c7"],
      High: ["#f97316", "#ffedd5"],
      Critical: ["#ef4444", "#fee2e2"],
    };
    return map[p] || ["#64748b", "#e2e8f0"];
  };
  const atat = (d) => d + 18000; // adjust to local
  const fmtDate = (d) => {
    if (!d) return "—";
    const dt = new Date(typeof d === "number" ? d : d.includes("T") ? d : d + (d.length === 10 ? "T00:00:00" : ""));
    if (isNaN(dt)) return d;
    return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  };
  const isOverdue = (due) => {
    if (!due) return false;
    const diff = (new Date(due).getTime() - Date.now()) / 86400000;
    return diff < 0;
  };
  const uidHex = (u) => {
    let h = 0;
    for (let i = 0; i < (u || "").length; i++) h = (h * 31 + u.charCodeAt(i)) >>> 0;
    return h % 6;
  };
  const initials = (name) =>
    (name || "?")
      .split(/\s+/)
      .map((w) => w[0])
      .slice(0, 2)
      .join("")
      .toUpperCase();

  function api(path, opts = {}) {
    const token = localStorage.getItem("plm_token");
    return fetch(API + path, {
      ...opts,
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
        ...(opts.headers || {}),
      },
    }).then(async (res) => {
      const data = await res.json().catch(() => ({}));
      if (res.status === 401) {
        localStorage.clear();
        window.location.href = "index.html";
        throw new Error("Session expired");
      }
      if (!res.ok) throw new Error(data.error || "Request failed");
      return data;
    });
  }

  function toast(msg, type = "ok") {
    const wrap = $("#toastWrap");
    const el = document.createElement("div");
    el.className = "toast " + type;
    el.innerHTML =
      type === "err"
        ? `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>`
        : `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M20 6L9 17l-5-5"/></svg>`;
    el.appendChild(document.createTextNode(msg));
    wrap.appendChild(el);
    setTimeout(() => {
      el.style.transition = "opacity .4s, transform .4s";
      el.style.opacity = "0";
      el.style.transform = "translateY(10px)";
      setTimeout(() => el.remove(), 400);
    }, 3200);
  }

  let confirmCb = null;
  function confirmBox({ title, message, yes = "Confirm", danger = true, cb }) {
    confirmCb = cb;
    openModal(`
      <div class="modal-head"><h3>${title}</h3><button class="modal-x" data-close>✕</button></div>
      <div class="modal-body">
        <p style="color:#475569;line-height:1.6;font-size:14px">${message}</p>
        <div class="modal-foot">
          <button class="btn btn-ghost" data-close>Cancel</button>
          <button class="btn ${danger ? "btn-danger" : "btn-primary"}" id="confirmYes">${yes}</button>
        </div>
      </div>`);
    $("#confirmYes").addEventListener("click", async () => {
      closeModal();
      if (confirmCb) await confirmCb();
    });
  }

  function openModal(html) {
    $("#modalBackdrop").hidden = false;
    $("#modal").hidden = false;
    $("#modal").innerHTML = html;
    $$("[data-close]", $("#modal")).forEach((b) => b.addEventListener("click", closeModal));
    $("#modalBackdrop").addEventListener("click", closeModal);
    const first = $("#modal input, #modal select, #modal textarea");
    if (first) first.focus();
  }
  function closeModal() {
    $("#modalBackdrop").hidden = true;
    $("#modal").hidden = true;
    $("#modal").innerHTML = "";
  }

  const esc = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  /* ---------------- excel export ---------------- */
  function exportExcel(filename, data) {
    if (typeof XLSX === "undefined") {
      toast("Excel library not loaded (check internet connection)", "err");
      return;
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast("Excel downloaded");
  }

  function exportTasksExcel() {
    const tasks = state.exportTasks || state.tasks || [];
    const data = [["Task ID", "Title", "Client", "Task Type", "Status", "% Complete", "Priority", "Assignee", "Assigned By", "Assigned Date", "Due Date", "Duration (days)", "Dependencies", "Risk/Blockers", "Description", "Comments"]];
    tasks.forEach((t) =>
      data.push([
        t.task_code, t.title,
        t.client || "", t.task_type || "",
        t.status, t.percent_complete ?? 0,
        t.priority,
        t.assigned_name || t.assigned_to || "Unassigned",
        t.assigned_by_name || t.assigned_by || "",
        t.assigned_at ? fmtDate(t.assigned_at) : "",
        t.due_date ? fmtDate(t.due_date) : "",
        t.duration ?? "",
        t.dependencies || "",
        t.risk_blockers || "",
        t.description || "",
        t.comments || "",
      ])
    );
    exportExcel("zita_tasks", data);
  }

  function exportTimesheetsExcel() {
    const rows = state.exportTimesheets || state.timesheets || [];
    const data = [["Member", "Username", "Date", "Task", "Description", "Task Type", "Hours"]];
    rows.forEach((r) =>
      data.push([r.user_name || r.username, r.username, fmtDate(r.entry_date), r.task || "", r.description || "", r.domain || "", Number(r.hours)])
    );
    exportExcel("zita_timesheets", data);
  }

  function exportUsersExcel() {
    const data = [["Username", "Name", "Role", "Created"]];
    (state.users || []).forEach((u) => data.push([u.username, u.name, u.role, fmtDate(u.created_at)]));
    exportExcel("zita_members", data);
  }

  /* ---------------- auth boot ---------------- */
  function boot() {
    const token = localStorage.getItem("plm_token");
    const userRaw = localStorage.getItem("plm_user");
    if (!token) {
      window.location.href = "index.html";
      return;
    }
    try {
      state.user = JSON.parse(userRaw);
    } catch {
      window.location.href = "index.html";
      return;
    }
    if (state.user.must_change_password) {
      localStorage.clear();
      window.location.href = "index.html";
      return;
    }
    $("#sideUserName").textContent = state.user.name;
    $("#sideUserRole").textContent = "Admin" === state.user.role ? "Administrator" : state.user.role;
    $("#sideAvatar").textContent = initials(state.user.name);
    $("#sideAvatar").classList.add("av" + uidHex(state.user.username));
    $("#navTeamLabel").textContent = state.user.role === "admin" ? "Team & Users" : "Team";
    $("#logoutBtn").addEventListener("click", () => {
      localStorage.clear();
      window.location.href = "index.html";
    });
    $$(".nav-item").forEach((b) =>
      b.addEventListener("click", () => {
        $$(".nav-item").forEach((x) => x.classList.remove("active"));
        b.classList.add("active");
        navigate(b.dataset.view);
      })
    );
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") closeModal();
    });
    navigate("dashboard");
  }

  const PAGE_META = {
    dashboard: ["Dashboard", "A live overview of tasks, statuses and hours logged."],
    tasks: ["Tasks", "Create, assign and track tasks. Everyone can see the full picture."],
    timesheets: ["Timesheets", "Log your daily work hours per task and domain."],
    team: [state.user?.role === "admin" ? "Team & Users" : "Team", "People working in this workspace."],
  };

  function navigate(view) {
    const meta = PAGE_META[view] || ["", ""];
    $("#pageTitle").textContent = meta[0];
    $("#pageSub").textContent = meta[1];
    const actions = $("#topbarActions");
    actions.innerHTML = "";
    if (view === "dashboard") renderDashboard();
    else if (view === "tasks") renderTasks();
    else if (view === "timesheets") renderTimesheets();
    else if (view === "team") renderTeam();
  }

  /* ---------------- dashboard ---------------- */
  async function renderDashboard() {
    const view = $("#view");
    view.innerHTML = spinner();
    let summary;
    try {
      summary = await api("/api/summary?from=&to=");
      await refreshUsers();
    } catch (e) {
      view.innerHTML = `<div class="empty"><h3>Could not load dashboard</h3><p>${esc(e.message)}</p></div>`;
      return;
    }

    const statusMap = {};
    summary.tasksByStatus.forEach((s) => (statusMap[s.status] = s.c));
    const completed = statusMap["Completed"] || 0;
    const active = summary.totalTasks - completed;
    const over = state.tasks.filter((t) => t.status !== "Completed" && isOverdue(t.due_date)).length;

    view.innerHTML = `
      <div class="stat-grid">
        ${statCard(
          "Total tasks",
          summary.totalTasks,
          ["#6366f1", "#e0e7ff"],
          `<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/>`
        )}
        ${statCard(
          "Active work",
          active,
          ["#0ea5e9", "#e0f2fe"],
          `<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/>`
        )}
        ${statCard(
          "Completed",
          completed,
          ["#10b981", "#d1fae5"],
          `<path d="M20 6L9 17l-5-5"/>`
        )}
        ${statCard(
          "Overdue",
          over,
          ["#ef4444", "#fee2e2"],
          `<circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/>`
        )}
        ${statCard(
          "Hours logged",
          summary.hours.total_hours.toFixed(1),
          ["#8b5cf6", "#ede9fe"],
          `<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>`
        )}
        ${statCard(
          "Active members",
          state.users.length,
          ["#f59e0b", "#fef3c7"],
          `<path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/>`
        )}
      </div>

      <div class="chart-row">
        <div class="card chart-box">
          <h3>Tasks by status</h3>
          <div class="chart-canvas"><canvas id="chartStatus"></canvas></div>
        </div>
        <div class="card chart-box">
          <h3>Top working members</h3>
          <div class="chart-canvas"><canvas id="chartHours"></canvas></div>
        </div>
      </div>

      <div class="section-head">
        <div>
          <h2>Recently updated tasks</h2>
          <p>Latest activity across the workspace.</p>
        </div>
        <button class="btn btn-ghost btn-sm" onclick="__plm.go('tasks')">View all tasks →</button>
      </div>
      ${recentTasksTable()}
    `;

    renderCharts(summary);
    window.__plm = Object.assign(window.__plm || {}, { go: (v) => { $$(".nav-item").forEach((x) => x.classList.remove("active")); $('[data-view="' + v + '"]').classList.add("active"); navigate(v); } });
  }

  function statCard(label, num, [c, bg], svgPath) {
    return `
      <div class="card stat-card">
        <div class="stat-icon" style="background:linear-gradient(135deg,${c},${c}cc)">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${svgPath}</svg>
        </div>
        <div>
          <div class="stat-num">${num}</div>
          <div class="stat-label">${label}</div>
        </div>
        <div class="stat-spark"><svg viewBox="0 0 24 24" fill="${c}">${svgPath}</svg></div>
      </div>`;
  }

  function recentTasksTable() {
    const tasks = state.tasks
      .slice()
      .sort((a, b) => new Date(b.updated_at || 0) - new Date(a.updated_at || 0))
      .slice(0, 6);
    if (!tasks.length)
      return `<div class="empty"><h3>No tasks yet</h3><p>Create your first task to get things moving.</p></div>`;
    return `
      <div class="card table-wrap">
        <table class="data-table">
          <thead><tr><th>Task</th><th>Status</th><th>Priority</th><th>Assignee</th><th>Due date</th></tr></thead>
          <tbody>
            ${tasks
              .map(
                (t) => `<tr>
                  <td><strong>${esc(t.task_code)}</strong> · ${esc(t.title)}</td>
                  <td>${badge(t.status, statusColor(t.status))}</td>
                  <td>${badge(t.priority, priorityColor(t.priority))}</td>
                  <td>${avatarChip(t.assigned_name || t.assigned_to)}</td>
                  <td>${dueCell(t.due_date, t.status)}</td>
                </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>`;
  }

  function renderCharts(summary) {
    if (window.Chart) {
      if (state.charts.status) state.charts.status.destroy();
      if (state.charts.hours) state.charts.hours.destroy();
    }
    const elS = $("#chartStatus");
    const elH = $("#chartHours");
    if (elS && window.Chart) {
      const labels = summary.tasksByStatus.map((s) => s.status);
      const data = summary.tasksByStatus.map((s) => s.c);
      state.charts.status = new Chart(elS, {
        type: "doughnut",
        data: { labels, datasets: [{ data, backgroundColor: labels.map((l) => statusColor(l)[0]), borderWidth: 3, borderColor: "#fff" }] },
        options: { responsive: true, maintainAspectRatio: false, cutout: "66%", plugins: { legend: { position: "bottom", labels: { usePointStyle: true, pointStyle: "circle", padding: 16, font: { family: "Inter", size: 11, weight: 600 } } } } },
      });
    }
    if (elH && window.Chart) {
      const rows = state.timesheets.slice(0, 200);
      const byUser = {};
      rows.forEach((r) => {
        const k = r.user_name || r.username;
        byUser[k] = (byUser[k] || 0) + Number(r.hours);
      });
      const top = Object.entries(byUser).sort((a, b) => b[1] - a[1]).slice(0, 6);
      state.charts.hours = new Chart(elH, {
        type: "bar",
        data: { labels: top.map(([k]) => k), datasets: [{ label: "Hours", data: top.map(([, v]) => v), backgroundColor: ["#6366f1", "#8b5cf6", "#0ea5e9", "#10b981", "#f59e0b", "#f43f5e"], borderRadius: 8, maxBarThickness: 42 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: "#eef1f7" }, ticks: { precision: 0 } } } },
      });
    }
  }

  /* ---------------- tasks ---------------- */
  async function renderTasks() {
    const view = $("#view");
    view.innerHTML = spinner();
    await Promise.all([refreshTasks(), refreshUsers()]);
    const actions = $("#topbarActions");
    actions.innerHTML = `
      <div class="seg" id="taskSeg">
        <button class="${state.taskView === "mine" ? "active" : ""}" data-tv="mine">My work</button>
        <button class="${state.taskView === "all" ? "active" : ""}" data-tv="all">Overall</button>
      </div>
      <button class="btn btn-ghost" id="exportTasksBtn">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>
        Export
      </button>
      <button class="btn btn-primary" id="newTaskBtn">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>
        New task
      </button>`;

    $("#newTaskBtn").addEventListener("click", () => openTaskModal());
    $("#exportTasksBtn").addEventListener("click", exportTasksExcel);
    $$("#taskSeg [data-tv]").forEach((b) =>
      b.addEventListener("click", () => {
        state.taskView = b.dataset.tv;
        renderTasks();
      })
    );
    renderTaskViews();
  }

  function filterBar() {
    const f = state.filters;
    const types = [...new Set(state.tasks.map((t) => t.task_type).filter(Boolean))];
    return `
      <div class="filters">
        <div class="filter-item">
          <label>Search</label>
          <input class="input search" id="fSearch" placeholder="Title, code, description…" value="${esc(f.search || "")}">
        </div>
        <div class="filter-item">
          <label>Status</label>
          <select class="select" id="fStatus" style="min-width:150px">
            <option value="">All statuses</option>
            ${STATUS.map((s) => `<option ${f.status === s ? "selected" : ""}>${s}</option>`).join("")}
          </select>
        </div>
        <div class="filter-item">
          <label>Priority</label>
          <select class="select" id="fPriority" style="min-width:140px">
            <option value="">All priorities</option>
            ${PRIORITY.map((p) => `<option ${f.priority === p ? "selected" : ""}>${p}</option>`).join("")}
          </select>
        </div>
        <div class="filter-item">
          <label>Task Type</label>
          <select class="select" id="fType" style="min-width:150px">
            <option value="">All types</option>
            ${types.map((ty) => `<option ${f.task_type === ty ? "selected" : ""}>${esc(ty)}</option>`).join("")}
          </select>
        </div>
        <div class="filter-item">
          <label>Assignee</label>
          <select class="select" id="fAssignee" style="min-width:160px">
            <option value="">Everyone</option>
            ${state.users.map((u) => `<option value="${esc(u.username)}" ${f.assigned_to === u.username ? "selected" : ""}>${esc(u.name)}</option>`).join("")}
          </select>
        </div>
        <div class="filter-item">
          <label>Due from</label>
          <input type="date" class="input" id="fFrom" value="${f.from || ""}">
        </div>
        <div class="filter-item">
          <label>Due to</label>
          <input type="date" class="input" id="fTo" value="${f.to || ""}">
        </div>
        <button class="btn btn-ghost btn-sm" id="fApply" style="height:38px">Apply</button>
        <button class="btn btn-ghost btn-sm" id="fClear" style="height:38px">Reset</button>
      </div>`;
  }

  function renderTaskViews() {
    const view = $("#view");
    let tasks = state.tasks.slice();

    const f = state.filters;
    if (f.status) tasks = tasks.filter((t) => t.status === f.status);
    if (f.priority) tasks = tasks.filter((t) => t.priority === f.priority);
    if (f.task_type) tasks = tasks.filter((t) => t.task_type === f.task_type);
    if (f.assigned_to) tasks = tasks.filter((t) => t.assigned_to === f.assigned_to);
    if (f.search) {
      const q = f.search.toLowerCase();
      tasks = tasks.filter((t) => (t.title + " " + t.description + " " + t.task_code + " " + (t.dependencies || "") + " " + (t.client || "") + " " + (t.comments || "")).toLowerCase().includes(q));
    }
    if (f.from) tasks = tasks.filter((t) => t.due_date && t.due_date >= f.from);
    if (f.to) tasks = tasks.filter((t) => t.due_date && t.due_date <= f.to);

    if (state.taskView === "mine") {
      const mine = tasks.filter(
        (t) => (t.assigned_to === state.user.username) || (t.created_by === state.user.username)
      );
      state.exportTasks = mine;
      view.innerHTML =
        filterBar() +
        `<div>
          <div class="section-head"><div><h2>Assignments</h2><p>${state.user.role === "admin" ? "Tasks assigned to or created by you." : "Open a card to update its status or add a comment — the card moves automatically."}</p></div><span class="chip-count">${mine.length} tasks</span></div>
          ${mine.length ? board(mine) : empty("No assignments yet", "Create a task and assign it to someone to get started.")}
        </div>`;
      bindFilters();
      return;
    }

    state.exportTasks = tasks;
    view.innerHTML =
      filterBar() +
      `<div class="section-head"><div><h2>Overall view</h2><p>${state.user.role === "admin" ? "Every member's tasks across the team." : "Read-only view of every member's tasks."}</p></div><span class="chip-count">${tasks.length} tasks</span></div>` +
      (tasks.length ? taskTable(tasks) : empty("No tasks match", "Try changing the filters or create a new task."));
    bindFilters();
  }

  function bindFilters() {
    ["fSearch", "fStatus", "fPriority", "fAssignee", "fFrom", "fTo"].forEach((id) => {
      const el = $("#" + id);
      if (el) el.addEventListener("keydown", (e) => { if (e.key === "Enter") applyFilters(); });
    });
    const apply = $("#fApply");
    const clear = $("#fClear");
    if (apply) apply.addEventListener("click", applyFilters);
    if (clear) clear.addEventListener("click", () => { state.filters = {}; renderTasks(); });
  }

  function applyFilters() {
    state.filters = {
      search: ($("#fSearch") || {}).value || "",
      status: ($("#fStatus") || {}).value || "",
      priority: ($("#fPriority") || {}).value || "",
      task_type: ($("#fType") || {}).value || "",
      assigned_to: ($("#fAssignee") || {}).value || "",
      from: ($("#fFrom") || {}).value || "",
      to: ($("#fTo") || {}).value || "",
    };
    renderTaskViews();
  }

  function board(tasks) {
    const cols = ["Pending", "In Progress", "On Hold", "Blocked", "Completed"];
    return `<div class="board">
      ${cols
        .map((c) => {
          const [bg, soft] = statusColor(c);
          const items = tasks.filter((t) => t.status === c);
          return `
          <div class="board-col">
            <div class="board-col-head"><span class="badge" style="background:${soft};color:${bg}"><span class="dot"></span>${c}</span><span style="margin-left:auto;color:var(--muted);font-size:12px">${items.length}</span></div>
            <div class="board-col-body">
              ${items.map(taskCard).join("") || `<div class="empty" style="padding:26px 10px"><p>Nothing here</p></div>`}
            </div>
          </div>`;
        })
        .join("")}
    </div>`;
  }

  function taskCard(t) {
  return `
      <div class="task-card" onclick="__plm.openTask(${t.id})">
        <div class="task-card-top">
          <span class="task-code">${esc(t.task_code)}</span>
          <span>${badge(t.priority, priorityColor(t.priority), "no-dot")}</span>
        </div>
        <div class="task-title">${esc(t.title)}</div>
        ${t.due_date ? `<div style="font-size:11.5px;color:${isOverdue(t.due_date) && t.status !== "Completed" ? "var(--red)" : "var(--muted)"}">Due ${fmtDate(t.due_date)}${isOverdue(t.due_date) && t.status !== "Completed" ? " · overdue" : ""}</div>` : `<div style="font-size:11.5px;color:var(--muted)">No due date</div>`}
        <div class="task-meta">
          ${avatarChip(t.assigned_to || "Unassigned")}
          ${t.assigned_to === state.user.username ? `<span class="mini-chip" style="margin-left:auto;color:var(--primary-dark)">Assigned to you</span>` : ""}
        </div>
      </div>`;
  }

  function taskTable(tasks) {
    const canAnyEdit = (t) => state.user.role === "admin" || t.created_by === state.user.username;
    return `
      <div class="card table-wrap">
        <table class="data-table">
          <thead><tr>
            <th>Task</th><th>Task Type</th><th>Client</th><th>Status</th><th>% Comp</th><th>Priority</th>
            <th>Assignee</th><th>Assigned By</th><th>Assigned Date</th><th>Due Date</th><th>Duration</th>
            <th>Dependencies</th><th>Risk / Blockers</th><th>Comments</th><th></th>
          </tr></thead>
          <tbody>
            ${tasks
              .map(
                (t) => { const edit = canAnyEdit(t);
                return `<tr>
                  <td>
                    <div><strong>${esc(t.task_code)}</strong> · ${esc(t.title)}</div>
                    ${t.description ? `<div style="font-size:11.5px;color:var(--muted);margin-top:2px;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(t.description)}</div>` : ""}
                  </td>
                  <td>${t.task_type ? `<span class="mini-chip">${esc(t.task_type)}</span>` : "—"}</td>
                  <td>${t.client ? esc(t.client) : "—"}</td>
                  <td>${badge(t.status, statusColor(t.status))}</td>
                  <td><strong>${t.percent_complete ?? 0}%</strong></td>
                  <td>${badge(t.priority, priorityColor(t.priority))}</td>
                  <td>${avatarChip(t.assigned_to || "Unassigned")}</td>
                  <td>${t.assigned_by_name || t.assigned_by || "—"}</td>
                  <td>${t.assigned_at ? fmtDate(t.assigned_at) : "—"}</td>
                  <td>${dueCell(t.due_date, t.status)}</td>
                  <td>${t.duration != null ? t.duration + "d" : "—"}</td>
                  <td style="max-width:180px">${t.dependencies ? esc(t.dependencies) : "—"}</td>
                  <td>${t.risk_blockers ? `<span style="color:${t.risk_blockers.toLowerCase().includes("high") || t.risk_blockers.toLowerCase().includes("critical") ? "var(--red)" : "var(--slate)"}">${esc(t.risk_blockers)}</span>` : "—"}</td>
                  <td style="max-width:160px"><span style="color:var(--muted)">${t.comments ? esc(t.comments.slice(0, 60).replace(/\n/g, " ")) + (t.comments.length > 60 ? "…" : "") : "—"}</span></td>
                  <td>
                    <div class="row-actions">
                      <button class="icon-btn" title=${edit ? '"Edit"' : '"View"'} onclick="__plm.${edit ? "openTask" : "viewTask"}(${t.id})"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>
                    </div>
                  </td>
                </tr>`; }
              )
              .join("")}
          </tbody>
        </table>
      </div>`;
  }

  function dueCell(due, status) {
    if (!due) return '<span style="color:var(--muted)">—</span>';
    const overdue = isOverdue(due) && status !== "Completed";
    return `<span style="${overdue ? "color:var(--red);font-weight:700" : ""}">${fmtDate(due)}${overdue ? ' <span style="font-weight:700">⚠</span>' : ""}</span>`;
  }

  function avatarChip(u) {
    if (!u) return '<span class="mini-chip">Unassigned</span>';
    const match = state.users.find((x) => x.username === u || x.name === u);
    const name = match ? match.name : u;
    return `<span class="mini-chip"><span class="avatar av${uidHex(u)}" style="width:18px;height:18px;font-size:9px">${initials(name)}</span>${esc(name)}</span>`;
  }

  /* -------- helper renderers for task detail -------- */
  function taskInfoRow(t) {
    return `
      <div class="stat-row" style="margin-bottom:14px">
        <div class="mini-stat"><div class="num">${badge(t.status, statusColor(t.status))}</div><div class="lbl">Status</div></div>
        <div class="mini-stat"><div class="num">${t.percent_complete ?? 0}%</div><div class="lbl">Complete</div></div>
        <div class="mini-stat"><div class="num">${badge(t.priority, priorityColor(t.priority))}</div><div class="lbl">Priority</div></div>
        <div class="mini-stat"><div class="num">${fmtDate(t.due_date)}</div><div class="lbl">Due date</div></div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 18px;font-size:13px;margin-bottom:14px">
        ${kv("Task Type", t.task_type || "—")}
        ${kv("Client", t.client || "—")}
        ${kv("Assignee", (t.assigned_name || t.assigned_to) || "Unassigned")}
        ${kv("Assigned By", (t.assigned_by_name || t.assigned_by) || "—")}
        ${kv("Assigned Date", t.assigned_at ? fmtDate(t.assigned_at) : "—")}
        ${kv("Duration", t.duration != null ? t.duration + " days" : "—")}
        ${kv("Dependencies", t.dependencies || "—")}
        ${kv("Risk / Blockers", t.risk_blockers || "—")}
      </div>
      ${t.description ? `<div style="background:#f8f9fd;border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:14px"><strong style="font-size:12px;color:var(--muted)">DESCRIPTION</strong><p style="margin-top:5px;line-height:1.6;color:#334155;white-space:pre-wrap">${esc(t.description)}</p></div>` : ""}`;
  }
  const kv = (k, v) => `<div><span style="color:var(--muted);font-weight:600">${k}:</span> ${esc(v)}</div>`;

  function commentsHTML(t) {
    if (!t.comments) return `<div style="padding:10px 2px;color:var(--muted);font-size:13px">No comments yet.</div>`;
    return `<div style="max-height:180px;overflow:auto;border:1px solid var(--line);border-radius:10px;background:#f8f9fd;padding:12px 14px"><div style="white-space:pre-wrap;font-size:13px;line-height:1.55;color:#334155">${esc(t.comments)}</div></div>`;
  }

  function renderTaskActions(t) {
    openModal(`
      <div class="modal-head"><h3>${esc(t.task_code)} · ${esc(t.title)}</h3><button class="modal-x" data-close>✕</button></div>
      <div class="modal-body">
        <div style="font-size:12.5px;color:var(--muted);margin-bottom:12px">Change the status below — the card will move automatically.</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px">
          ${STATUS.map((s) => `<button class="btn btn-sm ${t.status === s ? "btn-primary" : "btn-ghost"}" data-qstatus="${s}">${s}</button>`).join("")}
        </div>
        ${taskInfoRow(t)}
        <div class="section-head" style="margin:0 0 10px"><div><h2 style="font-size:14px">Comments</h2></div></div>
        ${commentsHTML(t)}
        <div style="display:flex;gap:10px;margin-top:12px">
          <textarea class="textarea" id="cmtBox" placeholder="Add a comment…" style="min-height:64px"></textarea>
          <button class="btn btn-primary" id="cmtAdd" style="height:fit-content;white-space:nowrap">Add comment</button>
        </div>
      </div>`);
    $$("[data-qstatus]").forEach((b) => b.addEventListener("click", () => quickStatus(t, b.dataset.qstatus)));
    $("#cmtAdd").addEventListener("click", async () => {
      const text = $("#cmtBox").value.trim();
      if (!text) return toast("Write a comment first", "err");
      try {
        await api("/api/tasks/" + t.id, { method: "PATCH", body: JSON.stringify({ comment: text }) });
        state.tasks = await api("/api/tasks");
        const fresh = state.tasks.find((x) => x.id === t.id);
        if (fresh) renderTaskActions(fresh);
        toast("Comment added");
      } catch (e) {
        toast(e.message, "err");
      }
    });
  }

  function renderViewOnly(t) {
    openModal(`
      <div class="modal-head"><h3>${esc(t.task_code)} · ${esc(t.title)}</h3><button class="modal-x" data-close>✕</button></div>
      <div class="modal-body">
        ${taskInfoRow(t)}
        ${t.comments ? `<div class="section-head" style="margin:0 0 10px"><div><h2 style="font-size:14px">Comments</h2></div></div>${commentsHTML(t)}` : ""}
        <div class="modal-foot"><button class="btn btn-ghost" data-close>Close</button></div>
      </div>`);
  }

  /* ---------------- task modal create / admin edit ---------------- */
  async function openTaskModal(task = null) {
    const isEdit = !!task;
    let nextCode = "";
    if (!isEdit) {
      try { const r = await api("/api/tasks/next-code"); nextCode = r.next_code; } catch {}
    }
    const assignees = state.users.filter((u) => u.role !== "admin");
    const opts = assignees
      .map((u) => `<option value="${esc(u.username)}" ${task && task.assigned_to === u.username ? "selected" : ""}>${esc(u.name)} (${esc(u.username)})</option>`)
      .join("");
    const today = new Date().toISOString().slice(0, 10);
    let durVal = "";
    if (task && task.due_date && task.duration != null) durVal = task.duration + " days";
    else if (!task && false) durVal = "";

    openModal(`
      <div class="modal-head"><h3>${isEdit ? "Edit task" : "New task"}</h3><button class="modal-x" data-close>✕</button></div>
      <div class="modal-body">
        <div class="form-grid">
          <div class="form-field">
            <label>Task ID / number</label>
            <input class="input" id="tCode" value="${isEdit ? esc(task.task_code) : esc(nextCode)}" placeholder="TSK-0001">
            <span class="hint">Auto-filled with the next number · editable.</span>
          </div>
          <div class="form-field">
            <label>Status</label>
            <select class="select" id="tStatus">${STATUS.map((s) => `<option ${task && task.status === s ? "selected" : ""}>${s}</option>`).join("")}</select>
          </div>
          <div class="form-field full">
            <label>Title *</label>
            <input class="input full" id="tTitle" value="${isEdit ? esc(task.title) : ""}" placeholder="Short, clear title">
          </div>
          <div class="form-field"><label>Client</label><input class="input" id="tClient" value="${isEdit ? esc(task.client || "") : ""}" placeholder="e.g. ZITA"></div>
          <div class="form-field"><label>Task Type</label><input class="input" id="tType" value="${isEdit ? esc(task.task_type || "") : ""}" placeholder="e.g. Glovia, BODS, MDM"></div>
          <div class="form-field"><label>Priority</label><select class="select" id="tPriority">${PRIORITY.map((p) => `<option ${task && task.priority === p ? "selected" : ""}>${p}</option>`).join("")}</select></div>
          <div class="form-field">
            <label>Assign to</label>
            <select class="select" id="tAssignee"><option value="">Unassigned</option>${opts}</select>
            ${task ? `<span class="hint">${esc((task.assigned_by_name || task.assigned_by) || "Not assigned yet")} · ${task.assigned_at ? "on " + fmtDate(task.assigned_at) : ""}</span>` : ""}
          </div>
          <div class="form-field"><label>Due date</label><input type="date" class="input" id="tDue" value="${isEdit && task.due_date ? task.due_date : ""}"></div>
          <div class="form-field">
            <label>Duration</label>
            <input class="input" id="tDur" value="${isEdit ? esc(durVal) : ""}" placeholder="auto" disabled>
            <span class="hint">Auto-calculated from due date.</span>
          </div>
          <div class="form-field"><label>% Complete</label><input type="number" class="input" id="tPct" min="0" max="100" value="${isEdit ? (task.percent_complete ?? 0) : 0}"></div>
          <div class="form-field full"><label>Dependencies</label><input class="input" id="tDep" value="${isEdit ? esc(task.dependencies || "") : ""}" placeholder="Tasks or work that must happen first"></div>
          <div class="form-field full"><label>Risk / Blockers</label><input class="input" id="tRisk" value="${isEdit ? esc(task.risk_blockers || "") : ""}" placeholder="e.g. Waiting on vendor data"></div>
          <div class="form-field full"><label>Description</label><textarea class="textarea" id="tDesc" placeholder="What needs to be done?">${isEdit ? esc(task.description || "") : ""}</textarea></div>
          <div class="form-field full"><label>Comments</label><textarea class="textarea" id="tComments" placeholder="Comment thread…">${isEdit ? esc(task.comments || "") : ""}</textarea></div>
        </div>
        <div class="modal-foot">
          <button class="btn btn-ghost" data-close>Cancel</button>
          ${isEdit ? `<button class="btn btn-danger" id="tDelete">Delete</button>` : ""}
          <button class="btn btn-primary" id="tSave">${isEdit ? "Save changes" : "Create & assign"}</button>
        </div>
      </div>`);
    const del = $("#tDelete");
    if (del) del.addEventListener("click", () => confirmBox({ title: "Delete task?", message: `${esc(task.task_code)} · ${esc(task.title)} will be permanently deleted.`, yes: "Delete", cb: async () => { await api("/api/tasks/" + task.id, { method: "DELETE" }); toast("Task deleted"); renderTasks(); } }));
    $("#tSave").addEventListener("click", () => saveTask(task));
  }

  async function saveTask(task) {
    const body = {
      task_code: $("#tCode").value.trim().toUpperCase(),
      title: $("#tTitle").value.trim(),
      client: $("#tClient").value.trim(),
      task_type: $("#tType").value.trim(),
      priority: $("#tPriority").value,
      assigned_to: $("#tAssignee").value || null,
      due_date: $("#tDue").value || null,
      dependencies: $("#tDep").value.trim(),
      risk_blockers: $("#tRisk").value.trim(),
      percent_complete: Math.max(0, Math.min(100, Number($("#tPct").value || 0))),
      status: $("#tStatus").value,
      description: $("#tDesc").value.trim(),
    };
    if (task) body.comments = $("#tComments").value || "";
    if (!body.task_code) return toast("Task ID is required", "err");
    if (!body.title) return toast("Title is required", "err");
    try {
      if (task) {
        await api("/api/tasks/" + task.id, { method: "PATCH", body: JSON.stringify(body) });
        toast("Task updated");
      } else {
        await api("/api/tasks", { method: "POST", body: JSON.stringify(body) });
        toast("Task created & assigned");
      }
      closeModal();
      renderTasks();
    } catch (e) {
      toast(e.message, "err");
    }
  }

  async function quickStatus(task, status) {
    try {
      await api("/api/tasks/" + task.id, { method: "PATCH", body: JSON.stringify({ status }) });
      closeModal();
      toast(`Task marked ${status}`);
      renderTasks();
    } catch (e) {
      toast(e.message, "err");
    }
  }

  window.__plm = window.__plm || {};
  window.__plm.openTask = (id) => {
    const task = state.tasks.find((t) => t.id === id);
    if (!task) return;
    if (state.user.role === "admin" || task.created_by === state.user.username) return openTaskModal(task);
    const mine = task.assigned_to === state.user.username;
    return mine ? renderTaskActions(task) : renderViewOnly(task);
  };
  window.__plm.viewTask = (id) => {
    const task = state.tasks.find((t) => t.id === id);
    if (!task) return;
    renderViewOnly(task);
  };

  /* ---------------- timesheets ---------------- */
  async function renderTimesheets() {
    const view = $("#view");
    view.innerHTML = spinner();
    await refreshTimesheets();
    const f = state.tsFilters;

    let rows = state.timesheets.slice();
    if (f.username) rows = rows.filter((r) => r.username === f.username);
    if (f.from) rows = rows.filter((r) => r.entry_date >= f.from);
    if (f.to) rows = rows.filter((r) => r.entry_date <= f.to);
    if (f.domain) rows = rows.filter((r) => r.domain === f.domain);
    if (f.search) rows = rows.filter((r) => (r.task + " " + r.description).toLowerCase().includes(f.search.toLowerCase()));

    const totalHours = rows.reduce((a, r) => a + Number(r.hours), 0);
    const domains = [...new Set(state.timesheets.map((r) => r.domain).filter(Boolean))];
    state.exportTimesheets = rows;

    view.innerHTML = `
      ${state.user.role === "admin" ? `
        <div class="card" style="padding:20px;display:flex;gap:14px;align-items:center">
          <div class="stat-icon" style="background:linear-gradient(135deg,#6366f1,#8b5cf6);width:44px;height:44px;border-radius:12px;display:grid;place-items:center;color:#fff;flex:0 0 44px">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>
          </div>
          <div>
            <strong style="font-size:14px">Monitor mode</strong>
            <div style="color:var(--muted);font-size:12.5px;margin-top:2px">As an admin you monitor all members' timesheets below. Members log their own hours.</div>
          </div>
        </div>` : `
      <div class="card" style="padding:20px">
        <div class="section-head" style="margin:0 0 16px">
          <div><h2>Log hours</h2><p>Enter the work you did today.</p></div>
        </div>
        <div class="form-grid">
          <div class="form-field"><label>Task</label><input class="input" id="tsTask" placeholder="e.g. TSK-0001 · Glovia cleanse"></div>
          <div class="form-field"><label>Task Type</label><input class="input" id="tsDomain" list="domainList" placeholder="e.g. Glovia, BODS, MDM"><datalist id="domainList">${domains.map((d) => `<option value="${esc(d)}">`).join("")}</datalist></div>
          <div class="form-field"><label>Date *</label><input type="date" class="input" id="tsDate" value="${new Date().toISOString().slice(0, 10)}"></div>
          <div class="form-field"><label>Hours *</label><input type="number" class="input" id="tsHours" min="0.5" max="24" step="0.5" placeholder="e.g. 8"></div>
          <div class="form-field full"><label>Description</label><textarea class="textarea" id="tsDesc" placeholder="What did you accomplish?"></textarea></div>
        </div>
        <div style="display:flex;justify-content:flex-end;margin-top:16px">
          <button class="btn btn-primary" id="tsAdd"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Add entry</button>
        </div>
      </div>`}

      <div class="stat-grid" style="margin:20px 0">
        ${statCard("Hours shown", totalHours.toFixed(1), ["#8b5cf6", "#ede9fe"], `<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>`)}
        ${statCard("Entries shown", rows.length, ["#6366f1", "#e0e7ff"], `<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/>`)}
        ${statCard("Days logged", new Set(rows.map((r) => r.entry_date)).size, ["#10b981", "#d1fae5"], `<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>`)}
      </div>

      <div class="section-head">
        <div><h2>Timesheet records</h2><p>${state.user.role === "admin" ? "All members' timesheets — visible to admins." : "Your logged entries."}</p></div>
        <button class="btn btn-ghost btn-sm" id="exportTsBtn">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>
          Export Excel
        </button>
      </div>

      <div class="filters" style="margin-top:-4px">
        ${state.user.role === "admin" ? `<div class="filter-item"><label>Member</label><select class="select" id="tfUser" style="min-width:170px"><option value="">All members</option>${state.users.map((u) => `<option ${f.username === u.username ? "selected" : ""} value="${esc(u.username)}">${esc(u.name)}</option>`).join("")}</select></div>` : ""}
        <div class="filter-item"><label>From</label><input type="date" class="input" id="tfFrom" value="${f.from || ""}"></div>
        <div class="filter-item"><label>To</label><input type="date" class="input" id="tfTo" value="${f.to || ""}"></div>
        <div class="filter-item"><label>Task Type</label><select class="select" id="tfDomain" style="min-width:140px"><option value="">All types</option>${domains.map((d) => `<option ${f.domain === d ? "selected" : ""}>${esc(d)}</option>`).join("")}</select></div>
        <div class="filter-item"><label>Search</label><input class="input search" id="tfSearch" value="${esc(f.search || "")}" placeholder="Task / description…"></div>
        <button class="btn btn-ghost btn-sm" id="tfApply" style="height:38px">Apply</button>
        <button class="btn btn-ghost btn-sm" id="tfClear" style="height:38px">Reset</button>
      </div>

      ${rows.length ? timesheetTable(rows) : empty("No timesheet entries", "Log your hours above to get started.")}
    `;

    const addBtn = $("#tsAdd");
    if (addBtn) addBtn.addEventListener("click", addTimesheet);
    const apply = $("#tfApply");
    const clear = $("#tfClear");
    const expTs = $("#exportTsBtn");
    if (expTs) expTs.addEventListener("click", exportTimesheetsExcel);
    if (apply)
      apply.addEventListener("click", () => {
        Object.assign(state.tsFilters, {
          username: ($("#tfUser") || { value: "" }).value,
          from: ($("#tfFrom") || { value: "" }).value,
          to: ($("#tfTo") || { value: "" }).value,
          domain: ($("#tfDomain") || { value: "" }).value,
          search: ($("#tfSearch") || { value: "" }).value,
        });
        renderTimesheets();
      });
    if (clear)
      clear.addEventListener("click", () => {
        state.tsFilters = {};
        renderTimesheets();
      });
    ["tfFrom", "tfTo", "tfSearch"].forEach((id) => {
      const el = $("#" + id);
      if (el) el.addEventListener("keydown", (e) => { if (e.key === "Enter") apply && apply.click(); });
    });
  }

  async function addTimesheet() {
    const date = $("#tsDate").value;
    const hours = $("#tsHours").value;
    if (!date || !hours || Number(hours) <= 0) return toast("Date and a positive number of hours are required", "err");
    try {
      await api("/api/timesheets", {
        method: "POST",
        body: JSON.stringify({
          task: $("#tsTask").value.trim(),
          description: $("#tsDesc").value.trim(),
          entry_date: date,
          hours: Number(hours),
          domain: $("#tsDomain").value.trim(),
        }),
      });
      toast("Hours logged");
      renderTimesheets();
    } catch (e) {
      toast(e.message, "err");
    }
  }

  function timesheetTable(rows) {
    return `
      <div class="card table-wrap">
        <table class="data-table">
          <thead><tr>
            <th>Member</th><th>Date</th><th>Task</th><th>Description</th><th>Task Type</th><th>Hours</th><th></th>
          </tr></thead>
          <tbody>
            ${rows
              .map(
                (r) => `<tr>
                  <td>${avatarChip(r.username)}</td>
                  <td style="white-space:nowrap">${fmtDate(r.entry_date)}</td>
                  <td>${esc(r.task || "—")}</td>
                  <td style="max-width:320px;color:#475569">${esc(r.description || "—")}</td>
                  <td>${r.domain ? `<span class="mini-chip">${esc(r.domain)}</span>` : "—"}</td>
                  <td><strong>${Number(r.hours).toFixed(1)}h</strong></td>
                  <td>
                    ${r.username === state.user.username || state.user.role === "admin" ? `<div class="row-actions"><button class="icon-btn danger" title="Delete" onclick="__plm.delTs(${r.id})"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6"/></svg></button></div>` : ""}
                  </td>
                </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>`;
  }

  window.__plm.delTs = (id) => {
    const row = state.timesheets.find((r) => r.id === id);
    confirmBox({
      title: "Delete timesheet entry?",
      message: `${fmtDate(row.entry_date)} · ${esc(row.task)} · ${row.hours}h will be removed.`,
      yes: "Delete",
      cb: async () => {
        await api("/api/timesheets/" + id, { method: "DELETE" });
        toast("Entry deleted");
        renderTimesheets();
      },
    });
  };

  /* ---------------- team / users ---------------- */
  async function renderTeam() {
    const view = $("#view");
    view.innerHTML = spinner();
    await refreshUsers();
    const isAdmin = state.user.role === "admin";

    view.innerHTML = `
      ${isAdmin ? `
        <div class="section-head">
          <div><h2>Create user</h2><p>New members sign in with their Employee ID and the default password <code>Welcome</code>, then set their own password on first sign-in.</p></div>
        </div>
        <div class="card" style="padding:20px">
          <div class="form-grid">
            <div class="form-field"><label>Employee ID / Username *</label><input class="input" id="nuUser" placeholder="e.g. GANESH"></div>
            <div class="form-field"><label>Full name *</label><input class="input" id="nuName" placeholder="e.g. Ganesh Kamalapuram"></div>
            <div class="form-field"><label>Role</label><select class="select" id="nuRole"><option value="user">User</option><option value="admin">Admin</option></select></div>
          </div>
          <div style="display:flex;justify-content:flex-end;margin-top:16px">
            <button class="btn btn-primary" id="nuCreate"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Create user</button>
          </div>
        </div>` : `
        <div class="section-head"><div><h2>Team members</h2><p>People working in this workspace.</p></div></div>`}

      <div class="section-head">
        <div><h2>Members (${state.users.length})</h2><p>${isAdmin ? "Manage accounts, reset passwords or remove users." : "Read-only directory of the team."}</p></div>
        <div class="row-actions">
          <button class="btn btn-ghost btn-sm" id="expMembers"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>Members</button>
          ${isAdmin ? `<button class="btn btn-ghost btn-sm" id="expAllTasks"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>All tasks</button>` : ""}
          ${isAdmin ? `<button class="btn btn-ghost btn-sm" id="expAllTs"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>All timesheets</button>` : ""}
        </div>
      </div>
      <div class="grid-2x">
        ${state.users
          .map(
            (u) => `
          <div class="user-card">
            <div class="avatar av${uidHex(u.username)}">${initials(u.name)}</div>
            <div class="user-card-info">
              <div class="user-card-name">${esc(u.name)} ${u.username === state.user.username ? '<span style="color:var(--primary);font-size:11px">(you)</span>' : ""}</div>
              <div class="user-card-uname">${esc(u.username)} · <span style="text-transform:capitalize;color:${u.role === "admin" ? "var(--primary-dark)" : "var(--muted)"}">${u.role}</span></div>
            </div>
            ${isAdmin
              ? `<div class="row-actions">
                  <button class="icon-btn" title="Reset password" onclick="__plm.resetPw('${esc(u.username)}','${esc(u.name)}')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 17a2 2 0 100-4 2 2 0 000 4z"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 11-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 11-4 0v-.09a1.65 1.65 0 00-1-1.51 1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 11-2.83-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 110-4h.09a1.65 1.65 0 001.51-1 1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 112.83-2.83l.06.06a1.65 1.65 0 001.82.33h.09a1.65 1.65 0 001-1.51V3a2 2 0 114 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 112.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82v.09a1.65 1.65 0 001.51 1H21a2 2 0 110 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg></button>
                  <button class="icon-btn danger" title="Delete user" onclick="__plm.delUser('${esc(u.username)}','${esc(u.name)}')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6"/></svg></button>
                </div>` : ""}
          </div>`
          )
          .join("")}
      </div>
    `;

    const create = $("#nuCreate");
    if (create)
      create.addEventListener("click", async () => {
        const username = $("#nuUser").value.trim();
        const name = $("#nuName").value.trim();
        const role = $("#nuRole").value;
        if (!username || !name) return toast("Username and name are required", "err");
        try {
          await api("/api/users", { method: "POST", body: JSON.stringify({ username, name, role }) });
          toast(`User ${username} created — default password is 'Welcome'`);
          renderTeam();
        } catch (e) {
          toast(e.message, "err");
        }
      });

    const expM = $("#expMembers");
    if (expM) expM.addEventListener("click", exportUsersExcel);
    const expAT = $("#expAllTasks");
    if (expAT) expAT.addEventListener("click", () => { state.exportTasks = state.tasks; exportTasksExcel(); });
    const expATS = $("#expAllTs");
    if (expATS) expATS.addEventListener("click", () => { state.exportTimesheets = state.timesheets; exportTimesheetsExcel(); });

    window.__plm.resetPw = (username, name) =>
      confirmBox({
        title: "Reset password",
        message: `Reset <strong>${esc(name)}</strong> (${esc(username)}) to the default password <strong>Welcome</strong>? They will be asked to set their own password on their next sign-in.`,
        yes: "Reset to Welcome",
        cb: async () => {
          try {
            await api("/api/users/" + username, { method: "PATCH", body: JSON.stringify({ password: "Welcome" }) });
            toast("Password reset to 'Welcome'");
          } catch (e) {
            toast(e.message, "err");
          }
        },
      });

    window.__plm.delUser = (username, name) =>
      confirmBox({
        title: "Delete user?",
        message: `${esc(name)} (${esc(username)}) will be removed. Their tasks remain but become unassigned.`,
        yes: "Delete",
        cb: async () => {
          try {
            await api("/api/users/" + username, { method: "DELETE" });
            toast("User deleted");
            renderTeam();
          } catch (e) {
            toast(e.message, "err");
          }
        },
      });
  }

  /* ---------------- data loading ---------------- */
  async function refreshTasks() {
    state.tasks = await api("/api/tasks");
  }
  async function refreshTimesheets() {
    state.timesheets = await api("/api/timesheets");
  }
  async function refreshUsers() {
    state.users = await api("/api/users");
    return state.users;
  }

  function spinner() {
    return `<div class="empty"><div style="width:34px;height:34px;border:3px solid #e0e7ff;border-top-color:#6366f1;border-radius:50%;margin:0 auto 14px;animation:spin 1s linear infinite"></div><p>Loading…</p></div>`;
  }

  function badge(text, [c, bg], noDot = false) {
    return `<span class="badge" style="background:${bg};color:${c}">${noDot ? "" : '<span class="dot"></span>'}<span>${text}</span></span>`;
  }

  function empty(title, sub) {
    return `<div class="empty">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6"/><path d="M9 14h6M9 18h6"/></svg>
      <h3>${title}</h3><p>${sub}</p>
    </div>`;
  }

  window.addEventListener("DOMContentLoaded", boot);
})();