// server.js
const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const path = require("path");
const session = require("express-session");
const multer = require("multer");
const xlsx = require("xlsx");

const app = express();
const cors = require("cors");
app.use(cors({
  origin: "https://lathysprayc9.netlify.app", // your Netlify frontend
  credentials: true
}));

const PORT = process.env.PORT || 3000;


// ---------- Paths / Files ----------
const PUBLIC_DIR = path.join(__dirname, "public");

const USERS_FILE = path.join(__dirname, "users.json");
const PAYROLL_FILE = path.join(__dirname, "payroll.json");
const AGRO_FILE = path.join(__dirname, "agronomist_data.json");
const FARM_REPORT_FILE = path.join(__dirname, "farm_report.json");
const LEGACY_PEST_DISEASE = path.join(__dirname, "pest_disease_data.json");

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

// Ensure farm_report entries have createdAt
function ensureCreatedAtForFarmReport() {
  const rows = readJSON(FARM_REPORT_FILE);
  let changed = false;
  for (const r of rows) {
    if (!r.createdAt) {
      r.createdAt = new Date().toISOString();
      changed = true;
    }
  }
  if (changed) writeJSON(FARM_REPORT_FILE, rows);
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

// ---------- Roles ----------
const ROLES = {
  Viewer: "Viewer",
  Supervisor: "Supervisor",
  Agronomist: "Agronomist",
  GeneralManager: "GeneralManager",
};

// ---------- Auth helpers ----------
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

app.get(
  "/viewer.html",
  requireAnyRolePage([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "viewer.html"))
);

app.get(
  "/supervisor.html",
  requireAnyRolePage([ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "supervisor.html"))
);

app.get(
  "/agronomist.html",
  requireAnyRolePage([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => res.sendFile(path.join(PUBLIC_DIR, "agronomist.html"))
);

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
  if (!payrollRecord || payrollRecord.role !== role)
    return res.json({ success: false, message: "Payroll number mismatch" });

  users.push({
    id: users.length ? Math.max(...users.map((u) => u.id || 0)) + 1 : 1,
    username,
    password,
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

// ---------- AGRONOMIST DATA ----------

// Get data (with optional filters)
app.get(
  "/agro",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const q = (req.query.q || "").trim().toLowerCase();
    const farmFilter = (req.query.farm || "").trim().toLowerCase();
    const ghFilter = (req.query.gh || "").trim().toLowerCase();
    const timeFilter = (req.query.time || "").trim().toLowerCase();

    const agroData = readJSON(AGRO_FILE);

    const filtered = agroData.filter(r => {
      const values = Object.values(r).map(v => String(v).toLowerCase());
      const textMatch = !q || values.some(v => v.includes(q));
      const farmMatch = !farmFilter || (r.farm && r.farm.toLowerCase() === farmFilter);
      const ghMatch = !ghFilter || (r.gh && r.gh.toLowerCase() === ghFilter);
      const timeMatch = !timeFilter || (r.time && r.time.toLowerCase() === timeFilter);
      return textMatch && farmMatch && ghMatch && timeMatch;
    });

    res.json(filtered);
  }
);

// Add a single row
app.post(
  "/agro/add",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const rows = readJSON(AGRO_FILE);
    const newRow = {
      id: rows.length ? Math.max(...rows.map(r => r.id || 0)) + 1 : 1,
      ...req.body,
      supervisorRemarks: ""
    };
    rows.push(newRow);
    writeJSON(AGRO_FILE, rows);
    res.json({ success: true, row: newRow });
  }
);
// Safer Save (prevents overwriting full file with filtered results)
app.post(
  "/agro/save",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    try {
      const newRows = req.body; // rows from frontend
      if (!Array.isArray(newRows)) {
        return res.status(400).json({ success: false, message: "Invalid data" });
      }

      const existingRows = readJSON(AGRO_FILE);

      let added = 0;
      let duplicates = 0;

      newRows.forEach(row => {
        // ensure ID
        if (!row.id) {
          row.id = existingRows.length
            ? Math.max(...existingRows.map(r => r.id || 0)) + 1
            : 1;
        }

        const exists = existingRows.some(r =>
          r.id === row.id ||
          (
            r.farm === row.farm &&
            r.gh === row.gh &&
            r.area === row.area &&
            r.crop === row.crop &&
            r.time === row.time
          )
        );

        if (exists) {
          duplicates++;
        } else {
          existingRows.push(row);
          added++;
        }
      });

      writeJSON(AGRO_FILE, existingRows);

      res.json({
        success: true,
        added,
        duplicates,
        message: duplicates
          ? `${duplicates} rows already saved`
          : `${added} new rows saved`
      });
    } catch (err) {
      console.error("Save failed:", err);
      res.status(500).json({ success: false, message: "Save failed" });
    }
  }
);


// Bulk replace data (used for Excel import in JSON format)
// Bulk replace data (used for Excel import in JSON format)
app.post(
  "/agro/bulk_set",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    try {
      const newData = req.body;
      if (!Array.isArray(newData)) {
        return res.status(400).json({ success: false, message: "Invalid data" });
      }

      const existing = readJSON(AGRO_FILE);

      // Determine max ID in existing data
      let maxId = existing.reduce((m, r) => Math.max(m, r.id || 0), 0);

      const merged = [...existing];

      newData.forEach((row) => {
        // Assign new ID if missing
        if (!row.id) row.id = ++maxId;

        // Find existing by ID
        const idx = merged.findIndex(r => r.id === row.id);

        if (idx !== -1) {
          // Merge, preserving supervisorRemarks if not provided
          merged[idx] = {
            ...merged[idx],
            ...row,
            supervisorRemarks: row.supervisorRemarks ?? merged[idx].supervisorRemarks ?? ""
          };
        } else {
          // New row
          if (!row.supervisorRemarks) row.supervisorRemarks = "";
          merged.push(row);
        }
      });

      // Save merged data
      writeJSON(AGRO_FILE, merged);
      res.json({ success: true, count: newData.length });
    } catch (err) {
      console.error("Bulk set failed:", err);
      res.status(500).json({ success: false, message: "Server error" });
    }
  }
);



// Supervisor remarks
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

// Search with filters
app.get(
  "/agro/search",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const { farm, gh, time, q } = req.query;
    let rows = readJSON(AGRO_FILE);

    if (farm) rows = rows.filter(r => String(r.farm || "").toLowerCase().includes(farm.toLowerCase()));
    if (gh) rows = rows.filter(r => String(r.gh || "").toLowerCase().includes(gh.toLowerCase()));
    if (time) rows = rows.filter(r => String(r.time || "").toLowerCase().includes(time.toLowerCase()));
    if (q) {
      const qLower = q.toLowerCase();
      rows = rows.filter(r => Object.values(r).some(v => String(v || "").toLowerCase().includes(qLower)));
    }

    rows.sort((a, b) => (b.id || 0) - (a.id || 0));
    res.json(rows);
  }
);


// Excel export/import
const upload = multer({ dest: "uploads/" });

// Excel export (supports .xlsx and .xlsm)
app.get(
  "/agro/export",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager, ROLES.Supervisor, ROLES.Viewer]),
  (req, res) => {
    try {
      const rows = readJSON(AGRO_FILE);
      const ws = xlsx.utils.json_to_sheet(rows);
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, "AgroData");

      // Decide file extension from query (default = xlsx)
      const ext = (req.query.ext || "xlsx").toLowerCase();
      const filename = `agro_export.${ext}`;

      // Write workbook to temp file
      const tmpPath = path.join(__dirname, filename);
      xlsx.writeFile(wb, tmpPath, { bookType: ext === "xlsm" ? "xlsm" : "xlsx" });

      res.download(tmpPath, filename, err => {
        fs.unlink(tmpPath, () => {}); // cleanup
        if (err) console.error("Download error:", err);
      });
    } catch (err) {
      console.error("Export failed:", err);
      res.status(500).json({ success: false, message: "Export failed" });
    }
  }
);

// Excel import (supports .xlsx and .xlsm)
// Universal import (CSV, XLSX, XLSM, ODS, etc.)
// Universal import (CSV, XLSX, XLSM, ODS, etc.)
app.post(
  "/agro/import",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  upload.single("file"),
  (req, res) => {
    if (!req.file) return res.status(400).send("No file uploaded");

    try {
      const ext = path.extname(req.file.originalname).toLowerCase();

      let importedRaw = [];

      if (ext === ".csv" || ext === ".txt") {
        // Read CSV/TXT
        const fileContent = fs.readFileSync(req.file.path, "utf8");
        const rows = xlsx.utils.sheet_to_json(
          xlsx.utils.csv_to_sheet(fileContent),
          { defval: "" }
        );
        importedRaw = rows;
      } else {
        // Excel-like (xlsx, xlsm, ods…)
        const wb = xlsx.readFile(req.file.path, { bookVBA: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        importedRaw = xlsx.utils.sheet_to_json(ws, { defval: "" });
      }

      // Read existing rows and determine next ID
      const existingRows = readJSON(AGRO_FILE);
      let maxId = existingRows.reduce((m, r) => Math.max(m, r.id || 0), 0);

      const importedData = importedRaw.map((row) => {
        const existing = row.id
          ? existingRows.find((r) => r.id === row.id)
          : null;

        return {
          id: row.id || ++maxId,
          farm: row.farm || row.Farm || "",
          gh: row.gh || row.GH || "",
          area: row.area || row.AREA || "",
          crop: row.crop || row.CROP || "",
          variety: row.variety || row.VARIETY || "",
          mode: row.mode || row.MODE || "",
          method: row.method || row.METHOD || "",
          time: row.time || row.TIME || "",

          // --- Week fields ---
          mon: row.mon || row.Mon || row.Monday || "",
          tue: row.tue || row.Tue || row.Tuesday || "",
          wed: row.wed || row.Wed || row.Wednesday || "",
          thu: row.thu || row.Thu || row.Thursday || "",
          fri: row.fri || row.Fri || row.Friday || "",
          sat: row.sat || row.Sat || row.Saturday || "",
          sun: row.sun || row.Sun || row.Sunday || "",

          target: row.target || row.Target || "",
          justification: row.justification || row.Justification || "",
          morning: row.morning || row.Morning || "",
          evening: row.evening || row.Evening || "",
          preparedBy: row.preparedBy || row["Prepared By"] || "",
          agronomistRemarks: row.agronomistRemarks || row["Agronomist Remarks"] || "",

          // Preserve supervisor remarks
          supervisorRemarks: existing
            ? existing.supervisorRemarks
            : (row.supervisorRemarks || "")
        };
      });

      // Merge
      const merged = [...existingRows];
      importedData.forEach((r) => {
        const idx = merged.findIndex((e) => e.id === r.id);
        if (idx >= 0) merged[idx] = r;
        else merged.push(r);
      });

      writeJSON(AGRO_FILE, merged);
      fs.unlinkSync(req.file.path);

      res.json({ success: true, count: importedData.length });
    } catch (err) {
      console.error("Import failed:", err);
      res.status(500).json({ success: false, message: "Import failed" });
    }
  }
);



// ---------- AGRONOMIST SEARCH WITH FARM, GH, TIME ----------
app.get(
  "/agro/search",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const { farm, gh, time, q } = req.query;
    let rows = readJSON(AGRO_FILE);

    if (farm) rows = rows.filter(r => String(r.farm || "").toLowerCase().includes(farm.toLowerCase()));
    if (gh) rows = rows.filter(r => String(r.gh || "").toLowerCase().includes(gh.toLowerCase()));
    if (time) rows = rows.filter(r => String(r.time || "").toLowerCase().includes(time.toLowerCase()));
    if (q) {
      const qLower = q.toLowerCase();
      rows = rows.filter(r => Object.values(r).some(v => String(v || "").toLowerCase().includes(qLower)));
    }

    rows.sort((a, b) => (b.id || 0) - (a.id || 0));
    res.json(rows);
  }
);

// ---------- FARM REPORT ----------
// Keep existing /farmreport routes unchanged (omitted here for brevity)
// Include ensureCreatedAtForFarmReport() on startup
app.get(
  "/farmreport",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    // return raw rows (createdAt preserved as ISO) — client will format for display
    const rows = readJSON(FARM_REPORT_FILE);
    res.json(rows);
  }
);

app.post(
  "/farmreport",
  requireAnyRoleApi([ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const {
      year,
      weekRange, farm, greenhouse, bed, crop, variety, pest, disease,
      pestRate, diseaseRate
    } = req.body || {};
    if (!weekRange || !farm || !greenhouse)
      return res.status(400).json({ success: false, message: "weekRange, farm, greenhouse required" });

    const rows = readJSON(FARM_REPORT_FILE);
    const nextId = rows.reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
    rows.push({
      id: nextId,
      year: String(year || ""),
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

// helper to parse a "1-4" style range into {start,end}
function parseWeekRange(weekRangeStr = "") {
  const s = String(weekRangeStr || "").trim();
  const parts = s.split("-");
  if (parts.length === 2) {
    const start = parseInt(parts[0], 10);
    const end = parseInt(parts[1], 10);
    if (!isNaN(start) && !isNaN(end)) return { start, end };
  }
  const single = parseInt(s, 10);
  if (!isNaN(single)) return { start: single, end: single };
  return { start: NaN, end: NaN };
}
// helper: format createdAt for display (Kenya time, short date+time)
function formatCreatedAtForDisplay(ts) {
  if (!ts) return "";
  try {
    return new Date(ts).toLocaleString("en-KE", {
      dateStyle: "short",
      timeStyle: "short",
      timeZone: "Africa/Nairobi",
      timeZoneName: "short",
    });
  } catch (e) {
    return ts;
  }
}

// Reusable filter function for farm report queries
function filterFarmRows(query = {}) {
  const {
    year,
    weekFrom,
    weekTo,
    farm,
    greenhouse,
    bed,
    crop,
    variety,
    pest,
    disease,
    pestRateMin,
    pestRateMax,
    diseaseRateMin,
    diseaseRateMax,
  } = query;

  const wf = weekFrom !== undefined && weekFrom !== "" ? parseInt(weekFrom, 10) : null;
  const wt = weekTo !== undefined && weekTo !== "" ? parseInt(weekTo, 10) : null;

  const pMin = pestRateMin !== undefined && pestRateMin !== "" ? Number(pestRateMin) : null;
  const pMax = pestRateMax !== undefined && pestRateMax !== "" ? Number(pestRateMax) : null;
  const dMin = diseaseRateMin !== undefined && diseaseRateMin !== "" ? Number(diseaseRateMin) : null;
  const dMax = diseaseRateMax !== undefined && diseaseRateMax !== "" ? Number(diseaseRateMax) : null;

  const rows = readJSON(FARM_REPORT_FILE);

  const filtered = rows.filter(r => {
    if (year && String(r.year || "").toLowerCase() !== String(year).toLowerCase()) return false;

    if (wf !== null || wt !== null) {
      const { start, end } = parseWeekRange(r.weekRange);
      const s = isNaN(start) ? null : start;
      const e = isNaN(end) ? s : end;
      if (s === null) return false;

      if (wf !== null && s < wf) return false;
      if (wt !== null && e > wt) return false;
    }

    if (farm && String(r.farm || "").toLowerCase() !== String(farm).toLowerCase()) return false;
    if (greenhouse && String(r.greenhouse || "").toLowerCase() !== String(greenhouse).toLowerCase()) return false;
    if (bed && String(r.bed || "").toLowerCase() !== String(bed).toLowerCase()) return false;
    if (crop && String(r.crop || "").toLowerCase() !== String(crop).toLowerCase()) return false;
    if (variety && String(r.variety || "").toLowerCase() !== String(variety).toLowerCase()) return false;
    if (pest && String(r.pest || "").toLowerCase() !== String(pest).toLowerCase()) return false;
    if (disease && String(r.disease || "").toLowerCase() !== String(disease).toLowerCase()) return false;

    const pr = Number(r.pestRate) || 0;
    const dr = Number(r.diseaseRate) || 0;
    if (pMin !== null && pr < pMin) return false;
    if (pMax !== null && pr > pMax) return false;
    if (dMin !== null && dr < dMin) return false;
    if (dMax !== null && dr > dMax) return false;

    return true;
  });

  // newest first by createdAt if present; otherwise by id desc
  filtered.sort((a, b) => {
    const aT = a.createdAt ? new Date(a.createdAt).getTime() : 0;
    const bT = b.createdAt ? new Date(b.createdAt).getTime() : 0;
    if (aT !== bT) return bT - aT;
    return (b.id || 0) - (a.id || 0);
  });

  return filtered;
}

// /farmreport/search (existing filtering)
app.get(
  "/farmreport/search",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const filtered = filterFarmRows(req.query || {});
    res.json(filtered);
  }
);

// /farmreport/charts already supports pest & disease and year/farm/greenhouse
app.get(
  "/farmreport/charts",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    const { year, farm, greenhouse, pest, disease } = req.query;
    // we reuse the filter function to keep consistent behavior
    const filtered = filterFarmRows({ year, farm, greenhouse, pest, disease });
    // for charts, sort ascending by week start
    filtered.sort((a, b) => {
      const aW = parseWeekRange(a.weekRange).start || 0;
      const bW = parseWeekRange(b.weekRange).start || 0;
      return aW - bW;
    });
    res.json(filtered);
  }
);

// NEW: Export farm report to Excel (supports same query params as /farmreport/search)
// If no query params provided -> exports all rows
app.get(
  "/farmreport/export",
  requireAnyRoleApi([ROLES.Viewer, ROLES.Supervisor, ROLES.Agronomist, ROLES.GeneralManager]),
  (req, res) => {
    // re-use filtering
    const filtered = filterFarmRows(req.query || []);

    // map rows to friendly columns for Excel, with CreatedAt in Kenya time
    const excelRows = filtered.map(r => ({
      ID: r.id || "",
      Year: r.year || "",
      WeekRange: r.weekRange || "",
      Farm: r.farm || "",
      Greenhouse: r.greenhouse || "",
      Bed: r.bed || "",
      Crop: r.crop || "",
      Variety: r.variety || "",
      Pest: r.pest || "",
      Disease: r.disease || "",
      PestRate: r.pestRate || 0,
      DiseaseRate: r.diseaseRate || 0,
      CreatedAt: r.createdAt ? formatCreatedAtForDisplay(r.createdAt) : "",
      CreatedAtISO: r.createdAt || ""
    }));

    // Build workbook
    const ws = xlsx.utils.json_to_sheet(excelRows);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "FarmReport");

    const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=farmreport_export.xlsx");
    res.send(buf);
  }
);
// ---------- Start ----------
app.listen(PORT, () => {
  ensureFile(USERS_FILE, "[]");
  ensureFile(PAYROLL_FILE, "[]");
  ensureFile(AGRO_FILE, "[]");
  ensureFile(FARM_REPORT_FILE, "[]");
  ensureFile(LEGACY_PEST_DISEASE, "[]");
  ensureCreatedAtForFarmReport();
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
