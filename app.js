// Basic SPA navigation
const views = document.querySelectorAll(".view");
document.querySelectorAll(".nav-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    views.forEach(v => v.classList.remove("active"));
    document.getElementById(btn.dataset.target).classList.add("active");
  });
});

// Local storage demo dataset
const LS_KEY = "campus_ideas";
function loadLocalIdeas() {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || []; }
  catch { return []; }
}
function saveLocalIdeas(list) {
  localStorage.setItem(LS_KEY, JSON.stringify(list));
}

// Microsoft auth placeholders (configure for Graph integration)
const msAuth = {
  isSignedIn: false,
  // TODO: Fill in MSAL config to enable sign-in
  signIn() {
    // Integrate MSAL here. For now, simulate.
    msAuth.isSignedIn = true;
    alert("Signed in (demo). Replace with MSAL sign-in for production.");
  },
  signOut() {
    msAuth.isSignedIn = false;
    alert("Signed out.");
  }
};

document.getElementById("btnSignIn").addEventListener("click", msAuth.signIn);
document.getElementById("btnSignOut").addEventListener("click", msAuth.signOut);

// Submit fallback local form
const ideaForm = document.getElementById("ideaForm");
if (ideaForm) {
  ideaForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const name = document.getElementById("name").value.trim();
    const email = document.getElementById("email").value.trim();
    const title = document.getElementById("title").value.trim();
    const desc = document.getElementById("desc").value.trim();
    const category = document.getElementById("category").value;
    const impact = Number(document.getElementById("impact").value || 0);
    const feasibility = Number(document.getElementById("feasibility").value || 0);
    const fileInput = document.getElementById("file");

    const idea = {
      id: crypto.randomUUID(),
      name, email, title, desc, category,
      impact, feasibility,
      submittedAt: new Date().toISOString(),
      fileName: fileInput?.files?.[0]?.name || null
    };

    // If signed-in, send to Microsoft Graph (Excel/SharePoint) — placeholder
    if (msAuth.isSignedIn) {
      await sendIdeaToGraph(idea);
    } else {
      const list = loadLocalIdeas();
      list.push(idea);
      saveLocalIdeas(list);
    }

    ideaForm.reset();
    alert("Idea submitted!");
    renderTable();
    renderSummary();
  });
}

// Placeholder for Microsoft Graph integration
async function sendIdeaToGraph(idea) {
  // TODO: Implement Graph calls:
  // - Create/append row in an Excel table in OneDrive/SharePoint
  // - Or add item to a SharePoint list with columns matching fields
  // - Optional: upload file as attachment if provided
  console.log("Graph submission (placeholder):", idea);
  // Fallback to local demo on success for now
  const list = loadLocalIdeas();
  list.push(idea);
  saveLocalIdeas(list);
}

// Admin dashboard actions
document.getElementById("btnClearLocal").addEventListener("click", () => {
  if (confirm("Clear local demo data?")) {
    localStorage.removeItem(LS_KEY);
    renderTable();
    renderSummary();
  }
});

document.getElementById("btnExportLocal").addEventListener("click", () => {
  const list = loadLocalIdeas();
  const csv = toCSV(list);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "campus_ideas.csv";
  a.click();
});

document.getElementById("btnSyncFromGraph").addEventListener("click", async () => {
  // TODO: Implement Graph read:
  // - Read Excel table rows or SharePoint list items
  // - Merge into local list (or replace)
  alert("Graph sync (placeholder). Configure Graph to read Excel/SharePoint.");
  renderTable();
  renderSummary();
});

// Render table
function renderTable() {
  const tbody = document.querySelector("#ideasTable tbody");
  const list = loadLocalIdeas();
  tbody.innerHTML = "";
  list.forEach((row, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escapeHTML(row.name)}</td>
      <td>${escapeHTML(row.email)}</td>
      <td>${escapeHTML(row.title)}</td>
      <td>${escapeHTML(row.desc)}</td>
      <td>${escapeHTML(row.category || "")}</td>
      <td>${row.impact || ""}</td>
      <td>${row.feasibility || ""}</td>
      <td>${new Date(row.submittedAt).toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  });
}
renderTable();

// Summary: top categories, word cloud, matrix
function renderSummary() {
  const list = loadLocalIdeas();

  // Top categories
  const counts = {};
  ["Energy", "Waste", "Innovation", "Social Impact"].forEach(c => counts[c] = 0);
  list.forEach(i => { if (i.category) counts[i.category] = (counts[i.category] || 0) + 1; });

  const topList = document.getElementById("topCategories");
  topList.innerHTML = "";
  Object.entries(counts)
    .sort((a,b) => b[1] - a[1])
    .forEach(([cat, n]) => {
      const li = document.createElement("li");
      li.textContent = `${cat}: ${n}`;
      topList.appendChild(li);
    });

  // Simple bar chart on canvas
  drawBarChart("categoryChart", counts);

  // Word cloud
  const text = list.map(i => `${i.title} ${i.desc}`).join(" ");
  const freq = wordFrequency(text);
  drawWordCloud("wordCloud", freq);

  // Impact-feasibility matrix
  drawMatrix("matrixChart", list);
}
renderSummary();

// Canvas charts (no external libs, Microsoft-friendly stack)
function drawBarChart(canvasId, counts) {
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  const labels = Object.keys(counts);
  const values = Object.values(counts);
  const max = Math.max(1, ...values);
  const barWidth = Math.floor((canvas.width - 80) / labels.length);
  const baseY = canvas.height - 30;

  ctx.fillStyle = "#e5e7eb";
  ctx.font = "12px Segoe UI";
  ctx.fillText("Ideas per category", 10, 16);

  labels.forEach((label, i) => {
    const h = Math.round((values[i] / max) * (canvas.height - 80));
    const x = 40 + i * barWidth;
    const y = baseY - h;
    ctx.fillStyle = "#2563eb";
    ctx.fillRect(x, y, Math.max(24, barWidth - 20), h);
    ctx.fillStyle = "#9ca3af";
    ctx.fillText(label, x, baseY + 14);
    ctx.fillStyle = "#e5e7eb";
    ctx.fillText(values[i], x, y - 6);
  });
}

function wordFrequency(text) {
  const stop = new Set(["the","and","a","an","of","for","to","in","on","with","by","from","is","are","be","this","that","it","we","our"]);
  const tokens = text.toLowerCase().match(/[a-z]+/g) || [];
  const freq = {};
  tokens.forEach(t => {
    if (!stop.has(t) && t.length > 2) freq[t] = (freq[t] || 0) + 1;
  });
  return freq;
}

function drawWordCloud(canvasId, freq) {
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  const entries = Object.entries(freq).sort((a,b) => b[1]-a[1]).slice(0, 50);
  const max = entries.length ? entries[0][1] : 1;

  // Simple scatter placement
  entries.forEach(([word, count], idx) => {
    const size = 10 + Math.round((count / max) * 26);
    ctx.font = `${size}px Segoe UI`;
    ctx.fillStyle = ["#2563eb","#10b981","#e5e7eb","#f59e0b","#ef4444"][idx % 5];
    const x = Math.floor(Math.random() * (canvas.width - 100)) + 20;
    const y = Math.floor(Math.random() * (canvas.height - 40)) + 30;
    ctx.fillText(word, x, y);
  });
}

function drawMatrix(canvasId, list) {
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  const padding = 40;
  const w = canvas.width - padding * 2;
  const h = canvas.height - padding * 2;

  // Axes
  ctx.strokeStyle = "#9ca3af";
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, padding + h);
  ctx.lineTo(padding + w, padding + h);
  ctx.stroke();

  ctx.fillStyle = "#e5e7eb";
  ctx.font = "12px Segoe UI";
  ctx.fillText("Impact (1 → 5)", padding + w/2 - 40, padding + h + 24);
  ctx.save();
  ctx.translate(12, padding + h/2 + 40);
  ctx.rotate(-Math.PI/2);
  ctx.fillText("Feasibility (1 → 5)", 0, 0);
  ctx.restore();

  // Points
  list.forEach((i, idx) => {
    const impact = clamp(i.impact || 0, 1, 5);
    const feas = clamp(i.feasibility || 0, 1, 5);
    const x = padding + (impact - 1) / 4 * w;
    const y = padding + h - (feas - 1) / 4 * h;
    ctx.fillStyle = ["#2563eb","#10b981","#f59e0b","#ef4444","#e5e7eb"][idx % 5];
    ctx.beginPath();
    ctx.arc(x, y, 6, 0, Math.PI * 2);
    ctx.fill();
  });

  // Quadrant labels
  ctx.fillStyle = "#9ca3af";
  ctx.font = "12px Segoe UI";
  ctx.fillText("High Impact / High Feasibility", padding + w - 210, padding + 16);
  ctx.fillText("Low Impact / High Feasibility", padding + 16, padding + 16);
  ctx.fillText("High Impact / Low Feasibility", padding + w - 200, padding + h - 8);
  ctx.fillText("Low Impact / Low Feasibility", padding + 16, padding + h - 8);
}

function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }
function escapeHTML(s="") {
  return s.replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
}
function toCSV(list) {
  const cols = ["name","email","title","desc","category","impact","feasibility","submittedAt","fileName"];
  const header = cols.join(",");
  const rows = list.map(i => cols.map(c => `"${String(i[c] ?? "").replace(/"/g,'""')}"`).join(","));
  return [header, ...rows].join("\n");
}

/* ============================
   Optional Microsoft integrations
   ============================ */

// 1) Microsoft Forms: Replace iframe src in index.html with your embed URL.

// 2) Microsoft Graph (Excel/SharePoint):
//    - Add MSAL config and sign-in flow to acquire token.
//    - Use Graph API to append rows to an Excel table (OneDrive/SharePoint)
//      or create items in a SharePoint list.
//    - Map 'idea' fields to table/list columns.
//    - Handle file upload via DriveItem or SharePoint attachments.

// 3) Power BI embed:
//    - Add the powerbi-client library locally.
//    - Provide embedConfig with type: 'report', id, embedUrl, accessToken.
//    - Call powerbi.embed(document.getElementById('powerBiContainer'), embedConfig).
//    - Ensure your report has visuals for categories and impact/feasibility.

// Example scaffold (non-functional placeholder):
async function embedPowerBI() {
  const config = {
    type: "report",
    id: "POWER_BI_REPORT_ID",           // TODO
    embedUrl: "POWER_BI_EMBED_URL",     // TODO
    accessToken: "POWER_BI_EMBED_TOKEN" // TODO
  };
  console.log("Power BI embed placeholder:", config);
  // powerbi.embed(document.getElementById('powerBiContainer'), config);
}
// Call this after you configure your Power BI credentials
// embedPowerBI();