// server.js
const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const path = require("path");
const session = require("express-session");
const multer = require("multer");
const xlsx = require("xlsx");

const app = express();
const PORT = 3000;

// ---------- Paths / Files ----------
const PUBLIC_DIR = path.join(__dirname, "public");

const USERS_FILE = path.join(__dirname, "users.json");
const PAYROLL_FILE = path.join(__dirname, "payroll.json"); // maps payrollNumber -> role
const AGRO_FILE = path.join(__dirname, "agronomist_data.json");
const FARM_REPORT_FILE = path.join(__dirname, "farm_report.json");
const LEGACY_PEST_DISEASE = path.join(__dirname, "pest_disease_data.json"); // keep as-is (legacy)

// ---------- Helpers ----------
function ensureFile(filePath, initial = "[]") {
  if (!fs.existsSync(filePath)) fs.writeFileSync(filePath, initial, "utf8");
}
function readJSON(filePath, fallback = []) {
  ensureFile(filePath);
  try {
    const txt = fs.readFileSync(filePath, "utf8");
    return txt.trim() ? JSON.parse(txt) : fallback;
  } catch {
    return fallback;
  }
}
function writeJSON(filePath, data) {
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), "utf8");
}

// ---------- Core middleware ----------
app.use(bodyParser.json({ limit: "10mb" }));
app.use(
  session({
    secret: "change-this-in-prod",
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false },
  })
);
app.use(express.static(PUBLIC_DIR));

// ---------- Auth helpers ----------
const ROLES = {
  Viewer: "Viewer",
  Supervisor: "Supervisor",
  Agronomist: "Agronomist",
  GeneralManager: "GeneralManager",
};

function requireLoginPage(req, res, next) {
  if (!req.session.user) return res.redirect("/login.html");
  next();
}
function requireLoginApi(req, res, next) {
  if (!req.session.user)
    return res.status(401).json({ success: false, message: "Unauthorized" });
  next();
}
function requireAnyRolePage(allowed) {
  return (req, res, next) => {
    if (!req.session.user) return res.redirect("/login.html");
    if (!allowed.includes(req.session.user.role))
      return res.status(403).send("Forbidden");
    next();
  };
}
function requireAnyRoleApi(allowed) {
  return (req, res, next) => {
    if (!req.session.user)
      return res.status(401).json({ success: false, message: "Unauthorized" });
    if (!allowed.includes(req.session.user.role))
      return res.status(403).json({ success: false, message: "Forbidden" });
    next();
  };
}

// ---------- Pages ----------
app.get("/", (req, res) => res.sendFile(path.join(PUBLIC_DIR, "index.html")));
app.get("/login.html", (req, res) =>
  res.sendFile(path.join(PUBLIC_DIR, "login.html"))
);
app.get("/register.html", (req, res) =>
  res.sendFile(path.join(PUBLIC_DIR, "register.html"))
);

// Viewer: everyone logged in can see
app.get(
  "/viewer.html",
  requireAnyRolePage([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "viewer.html"))
);

// Supervisor: Supervisor + Agronomist + GM
app.get(
  "/supervisor.html",
  requireAnyRolePage([ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "supervisor.html"))
);

// Agronomist: Agronomist + GM
app.get(
  "/agronomist.html",
  requireAnyRolePage([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "agronomist.html"))
);

// Farm Report: Agronomist + GM (others blocked)
app.get(
  "/farmreport.html",
  requireAnyRolePage([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "farmreport.html"))
);

// ---------- Auth APIs ----------
app.post("/register", (req, res) => {
  const { username, password, role, payrollNumber } = req.body || {};
  if (!username || !password || !role || !payrollNumber)
    return res.json({ success: false, message: "All fields are required" });

  if (!Object.values(ROLES).includes(role))
    return res.json({ success: false, message: "Invalid role" });

  const users = readJSON(USERS_FILE);
  if (users.find((u) => u.username === username))
    return res.json({ success: false, message: "Username already exists" });

  const payroll = readJSON(PAYROLL_FILE);
  const payrollRecord = payroll.find((p) => String(p.payrollNumber) === String(payrollNumber));
  if (!payrollRecord || payrollRecord.role !== role) {
    return res.json({
      success: false,
      message:
        "Invalid credentials: payroll number does not match the selected role.",
    });
  }

  users.push({
    id: users.length ? Math.max(...users.map((u) => u.id || 0)) + 1 : 1,
    username,
    password, // plain (you can add bcrypt later)
    role,
    payrollNumber,
    createdAt: new Date().toISOString(),
  });
  writeJSON(USERS_FILE, users);
  res.json({ success: true });
});

app.post("/login", (req, res) => {
  const { username, password } = req.body || {};
  const users = readJSON(USERS_FILE);
  const user = users.find((u) => u.username === username && u.password === password);
  if (!user) return res.json({ success: false, message: "Invalid credentials" });
  req.session.user = { username: user.username, role: user.role };
  res.json({ success: true, role: user.role, username: user.username });
});

app.post("/logout", (req, res) => req.session.destroy(() => res.json({ success: true })));

app.get("/check-session", (req, res) => {
  if (req.session.user) return res.json({ loggedIn: true, role: req.session.user.role });
  res.status(401).json({ loggedIn: false });
});

// ---------- AGRONOMIST DATA (master program) ----------
// Schema fields: farm, gh, area, crop, variety, mode, method, time,
// Day blocks (mon..sun): {rate, vol, chemical, area, mode}, plus:
// target, justification, morning, evening, preparedBy, agronomistRemarks, supervisorRemarks

app.get(
  "/agro",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.json(readJSON(AGRO_FILE))
);

app.post(
  "/agro/bulk_set",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    let arr = Array.isArray(req.body) ? req.body : [];
    // assign IDs if missing
    let nextId =
      readJSON(AGRO_FILE).reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
    arr = arr.map((r) => ({
      id: r.id && Number.isInteger(r.id) ? r.id : nextId++,
      ...r,
    }));
    writeJSON(AGRO_FILE, arr);
    res.json({ success: true, count: arr.length });
  }
);

app.post(
  "/agro/add",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const rows = readJSON(AGRO_FILE);
    const nextId = rows.reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
    const payload = { id: nextId, ...(req.body || {}) };
    rows.push(payload);
    writeJSON(AGRO_FILE, rows);
    res.json({ success: true, id: nextId });
  }
);

// Supervisor can only edit supervisorRemarks
app.post(
  "/agro/supervisor-remarks",
  requireAnyRoleApi([ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const { id, supervisorRemarks } = req.body || {};
    if (!id) return res.status(400).json({ success: false, message: "Missing id" });
    const rows = readJSON(AGRO_FILE);
    const idx = rows.findIndex((r) => r.id === Number(id));
    if (idx === -1) return res.status(404).json({ success: false, message: "Not found" });
    rows[idx].supervisorRemarks = supervisorRemarks || "";
    writeJSON(AGRO_FILE, rows);
    res.json({ success: true });
  }
);

// Excel export/import
app.get(
  "/agro/export",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const data = readJSON(AGRO_FILE);
    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "AgronomistData");
    const file = "agronomist_data.xlsx";
    xlsx.writeFile(wb, file);
    res.download(file, () => {
      try {
        fs.unlinkSync(file);
      } catch {}
    });
  }
);

const upload = multer({ dest: "uploads/" });
app.post(
  "/agro/import",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  upload.single("file"),
  (req, res) => {
    if (!req.file) return res.status(400).send("No file uploaded");
    const wb = xlsx.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(ws, { defval: "" });
    // assign IDs if missing
    let nextId = readJSON(AGRO_FILE).reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
    const mapped = data.map((r) => ({
      id: Number.isInteger(r.id) ? r.id : nextId++,
      ...r,
    }));
    writeJSON(AGRO_FILE, mapped);
    fs.unlinkSync(req.file.path);
    res.json({ success: true, count: mapped.length });
  }
);

// ---------- FARM REPORT (new) ----------
app.get(
  "/farmreport",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.json(readJSON(FARM_REPORT_FILE))
);

// Append new record (historical)
// *** MINIMAL CHANGE: include `year` when storing the record (no other logic changed) ***
app.post(
  "/farmreport",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const {
      year,                    // <-- NEW: read year from request body
      weekRange, farm, greenhouse, bed, crop, variety, pest, disease,
      pestRate, diseaseRate
    } = req.body || {};
    if (!weekRange || !farm || !greenhouse)
      return res.status(400).json({ success: false, message: "weekRange, farm, greenhouse required" });

    const rows = readJSON(FARM_REPORT_FILE);
    const nextId = rows.reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
    rows.push({
      id: nextId,
      year: String(year || ""),               // <-- NEW: store year
      weekRange: String(weekRange).trim(),
      farm: String(farm || ""),
      greenhouse: String(greenhouse || ""),
      bed: String(bed || ""),
      crop: String(crop || ""),
      variety: String(variety || ""),
      pest: String(pest || ""),
      disease: String(disease || ""),
      pestRate: Number(pestRate) || 0,
      diseaseRate: Number(diseaseRate) || 0,
      createdAt: new Date().toISOString(),
    });
    writeJSON(FARM_REPORT_FILE, rows);
    res.json({ success: true, id: nextId });
  }
);

// ---------- Start ----------
app.listen(PORT, () => {
  // create files if missing
  ensureFile(USERS_FILE, "[]");
  ensureFile(PAYROLL_FILE, "[]");       // you populate this
  ensureFile(AGRO_FILE, "[]");
  ensureFile(FARM_REPORT_FILE, "[]");
  ensureFile(LEGACY_PEST_DISEASE, "[]"); // legacy kept
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
