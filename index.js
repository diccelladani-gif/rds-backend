require("dotenv").config();
console.log("RDS Backend v4.0 starting...");

const express     = require("express");
const cors        = require("cors");
const XLSX        = require("xlsx");
const PDFDocument = require("pdfkit");
const fs          = require("fs");
const path        = require("path");
const Groq        = require("groq-sdk");
const mammoth     = require("mammoth");
const pdfjsLib    = require("pdfjs-dist/legacy/build/pdf.js");
const AdmZip      = require("adm-zip");
const sizeOf      = require("image-size");
const { createClient } = require("@supabase/supabase-js");

const SUPABASE_URL = process.env.SUPABASE_URL || "https://czgyxziunupgypvtkbod.supabase.co";
const SUPABASE_KEY = process.env.SUPABASE_KEY || "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImN6Z3l4eml1bnVwZ3lwdnRrYm9kIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NzM2OTIxNCwiZXhwIjoyMDkyOTQ1MjE0fQ.azBkUIm_bJGXaz3wkvHNU5MnltTN4F2It8jo8ZUrD_I";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── AUDIT LOG HELPER ─────────────────────────────────────
async function logAudit({ roomId, roomCode, action, performedBy, details = {} }) {
  try {
    await supabase.from("rds_audit_log").insert({
      room_id:      String(roomId),
      room_code:    roomCode,
      action,                          // "created" | "updated" | "deleted" | "exported"
      performed_by: performedBy || "system",
      details:      JSON.stringify(details),
      created_at:   new Date().toISOString()
    });
  } catch (e) {
    console.warn("Audit log failed:", e.message); // non-blocking
  }
}

async function extractPdfText(base64) {
  const buf  = Buffer.from(base64, "base64");
  const data = new Uint8Array(buf);
  const doc  = await pdfjsLib.getDocument({ data, useWorkerFetch: false, isEvalSupported: false }).promise;
  let text = "";
  for (let i = 1; i <= doc.numPages; i++) {
    const page    = await doc.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(" ") + "\n";
  }
  return text.trim();
}

const app = express();

// ─── MIDDLEWARE ───────────────────────────────────────────
app.use(cors({
  origin: (origin, callback) => {
    const allowed = !origin
      || origin.startsWith("http://localhost")
      || origin.endsWith(".vercel.app")
      || origin.endsWith(".onrender.com");
    if (allowed) callback(null, true);
    else callback(new Error("Not allowed by CORS"));
  },
  methods: ["GET", "POST", "PUT", "DELETE"],
  allowedHeaders: ["Content-Type"]
}));
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));
app.use((req, _res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// ─── CONFIG ──────────────────────────────────────────────
const DATA_DIR   = path.join(__dirname, "data");
const FILE_PATH  = path.join(DATA_DIR, "rds-data.xlsx");
const IMAGE_DIR  = path.join(DATA_DIR, "images");
const USERS_FILE = path.join(DATA_DIR, "users.xlsx");   // ← NEW
const SHEET      = "RDS";
if (!fs.existsSync(DATA_DIR))  fs.mkdirSync(DATA_DIR,  { recursive: true });
if (!fs.existsSync(IMAGE_DIR)) fs.mkdirSync(IMAGE_DIR, { recursive: true });

// ─── FIELD LABEL FORMATTER ───────────────────────────────
function formatFieldValue(key, val) {
  if (key === "userGroups") {
    try {
      const arr = JSON.parse(val);
      if (Array.isArray(arr)) return arr.map(u => `${u.role} × ${u.qty}`).join(", ");
    } catch (_) {}
  }
  return String(val);
}

function toLabel(key) {
  return key
    .replace(/([A-Z])/g, " $1")
    .replace(/^./, s => s.toUpperCase())
    .trim();
}

// ─── SECTION DEFINITIONS ─────────────────────────────────
const SECTIONS = [
  { label: "1. Room Identity & General Information",  keys: ["project","department","roomName","roomCode","location","roomTypology","criticalityLevel","infectionRiskCategory","isolationType"] },
  { label: "2. Architectural & Spatial Requirements", keys: ["netArea","minimumDimensions","clearances","ceilingHeight","doorType","doorSize","accessibility"] },
  { label: "3. Interior Finishes & Aesthetics",       keys: ["floor","skirting","walls","ceiling","wallProtection","specialFinishes"] },
  { label: "4. Clinical Functionality & Workflow",    keys: ["roomFunction","keyActivities","userGroups","operationalScenarios","patientZone","staffZone","equipmentZone","cleanZone","dirtyZone","patientFlow","staffFlow","materialFlow","entryPoints","restrictedZones"] },
  { label: "5. Capacity & Operations",                keys: ["patientCapacity","staffRequirement","peakLoad","throughput","averageStayTime","surgeCapacity","operationalHours"] },
  { label: "6. Adjacency Matrix",                     keys: ["mustBeAdjacent","shouldBeAdjacent","avoidAdjacency"] },
  { label: "7. MEP & Engineering Systems",            keys: ["airChangesACH","pressure","temperature","humidity","filtration","airflowDirection","pandemicMode","powerLoad","normalPower","emergencyPower","ups","numberOfSockets","specialOutlets","oxygen","medicalAir","vacuum","nitrousOxide","handWash","wc","shower","plumbingSpecialSystems"] },
  { label: "8. Digital & Smart Systems",              keys: ["hisEmr","pacs","lis","rtls","nurseCall","cctv","iotSensors","aiAnalytics"] },
  { label: "9. Safety & Infection Control",           keys: ["pressureRegime","isolationLevel","radiationProtection","biohazardHandling","fireSafety","emergencySystems"] },
  { label: "10. Stakeholder Experience",              keys: ["lightingQuality","acousticControl","privacy","patientComfort","familyInteraction","visualEnvironment"] },
  { label: "11. Fittings, Fixtures & Equipment",      keys: ["airFlowmeter","oxygenFlowmeter","suctionAdapterLowFlow","suctionBottle","oxygenFlowmeterLowFlow","trolleyProcedure","blenderAirOxygen","stoolAdjustableMobile","curtainTrackSystem","ivHook","additionalFF","infusionPumpSyringe","examinationLight","physiologicMonitor","infantIncubator","phototherapyLamp","supplyUnitCeiling","infusionPumpEnteral","infusionPumpSingleChannel","ventilatorNeonatal","additionalFE"] },
];

// ─── DATA HELPERS ─────────────────────────────────────────
function safeJson(str) {
  try { return JSON.parse(str); } catch { return {}; }
}

async function readAll() {
  try {
    const { data, error } = await supabase
      .from("rds_rooms")
      .select("*")
      .order("id", { ascending: true });
    if (error) throw error;
    return (data || []).map(row => ({
      ...row,
      data: typeof row.data === "string" ? safeJson(row.data) : (row.data || {})
    }));
  } catch(e) {
    console.error("readAll error:", e.message);
    return [];
  }
}

async function writeAll(rows) {
  // writeAll is only used for bulk replace — not needed with Supabase (individual ops used)
  return true;
}

// ─── USERS HELPER ─────────────────────────────────────────  ← NEW
function readUsers() {
  return [
    { id: "EMP001", name: "Rohan", email: "rohan@rds.med", password: "123",  role: "admin",    department: "administration" },
    { id: "EMP002", name: "Ravi",  email: "ravi@rds.med",  password: "1234", role: "reviewer", department: "clinical"       },
  ];
}

// ─── PDF BUILDER ─────────────────────────────────────────
function buildPDF(rows) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDocument({ size: "A4", margin: 45, info: { Title: "Room Data Sheet", Author: "Medical College RDS System" } });
    const chunks = [];
    doc.on("data",  c  => chunks.push(c));
    doc.on("end",   () => resolve(Buffer.concat(chunks)));
    doc.on("error", e  => reject(e));

    const NAVY   = "#1e3a8a";
    const BLUE   = "#2563eb";
    const LBLUE  = "#dbeafe";
    const LGRAY  = "#f1f5f9";
    const BORDER = "#cbd5e1";
    const TEXT   = "#111827";
    const MUTED  = "#64748b";
    const WHITE  = "#ffffff";

    const PAGE_W  = doc.page.width;
    const PAGE_H  = doc.page.height;
    const MARGIN  = 45;
    const CONTENT = PAGE_W - MARGIN * 2;

    const CRIT_COLORS = {
      Critical: "#dc2626", High: "#ea580c", Medium: "#ca8a04",
      Low: "#16a34a", Ancillary: "#0369a1"
    };

    function hLine(y, color = BORDER, lw = 0.5) {
      doc.save().strokeColor(color).lineWidth(lw)
         .moveTo(MARGIN, y).lineTo(MARGIN + CONTENT, y).stroke().restore();
    }

    function fillRect(x, y, w, h, color) {
      doc.save().fillColor(color).rect(x, y, w, h).fill().restore();
    }

    function borderRect(x, y, w, h, fill, stroke = BORDER, lw = 0.5) {
      doc.save().fillColor(fill).strokeColor(stroke).lineWidth(lw)
         .rect(x, y, w, h).fillAndStroke().restore();
    }

    function ensureSpace(needed) {
      if (doc.y + needed > PAGE_H - MARGIN - 30) {
        doc.addPage();
        drawPageHeader();
      }
    }

    function drawPageHeader() {
      fillRect(0, 0, PAGE_W, 36, NAVY);
      doc.font("Helvetica-Bold").fontSize(13).fillColor(WHITE)
         .text("ROOM DATA SHEET", MARGIN, 11, { width: CONTENT / 2 });
      doc.font("Helvetica").fontSize(9).fillColor("rgba(255,255,255,0.75)")
         .text("Medical College — Facility Planning", MARGIN + CONTENT / 2, 14,
               { width: CONTENT / 2, align: "right" });
      doc.y = 50;
    }

    rows.forEach((r, rIdx) => {
      if (rIdx > 0) doc.addPage();
      const d      = r.data || {};
      const crit   = d.criticalityLevel || "";
      const critC  = CRIT_COLORS[crit] || BLUE;

      drawPageHeader();

      let y = doc.y;

      const CARD_H = 82;
      borderRect(MARGIN, y, CONTENT, CARD_H, LBLUE, BLUE, 1);

      doc.font("Helvetica-Bold").fontSize(22).fillColor(NAVY)
         .text(r.roomCode || "—", MARGIN + 14, y + 10, { width: CONTENT * 0.55, lineBreak: false });
      doc.font("Helvetica").fontSize(11).fillColor(MUTED)
         .text(d.roomName || r.roomName || "Unnamed Room", MARGIN + 14, y + 38, { width: CONTENT * 0.55 });

      const BADGE_W = 110, BADGE_H = 28;
      const bx = MARGIN + CONTENT - BADGE_W - 14;
      const by = y + 10;
      fillRect(bx, by, BADGE_W, BADGE_H, critC);
      doc.font("Helvetica-Bold").fontSize(10).fillColor(WHITE)
         .text(crit.toUpperCase() || "SUBMITTED", bx, by + 8, { width: BADGE_W, align: "center" });
      doc.font("Helvetica").fontSize(9).fillColor(MUTED)
         .text(d.roomTypology || "", bx, by + 42, { width: BADGE_W, align: "center" });

      y += CARD_H + 10;

      const META = [
        ["Department",  d.department  || r.department  || "—"],
        ["Project",     d.project     || r.project     || "—"],
        ["Location",    d.location    || "—"],
        ["Net Area",    d.netArea     ? `${d.netArea} m²` : "—"],
        ["Patient Cap.",d.patientCapacity || "—"],
        ["Created",     r.createdAt   ? new Date(r.createdAt).toLocaleDateString("en-IN", { day:"numeric", month:"short", year:"numeric" }) : "—"],
      ];

      const COLS   = 3;
      const CELL_W = CONTENT / COLS;
      const CELL_H = 36;
      META.forEach(([label, value], i) => {
        const col  = i % COLS;
        const row  = Math.floor(i / COLS);
        const cx   = MARGIN + col * CELL_W;
        const cy   = y + row * CELL_H;
        fillRect(cx, cy, CELL_W, CELL_H, i % 2 === 0 ? LGRAY : WHITE);
        doc.save().strokeColor(BORDER).lineWidth(0.4)
           .rect(cx, cy, CELL_W, CELL_H).stroke().restore();
        doc.font("Helvetica-Bold").fontSize(7.5).fillColor(MUTED)
           .text(label.toUpperCase(), cx + 8, cy + 6, { width: CELL_W - 16, lineBreak: false });
        doc.font("Helvetica").fontSize(10).fillColor(TEXT)
           .text(String(value), cx + 8, cy + 18, { width: CELL_W - 16, lineBreak: false, ellipsis: true });
      });

      y += Math.ceil(META.length / COLS) * CELL_H + 14;
      doc.y = y;

      const imagePath = d.imagePath;
      if (imagePath && fs.existsSync(imagePath)) {
        ensureSpace(140);
        y = doc.y;
        try {
          const imgBuffer = fs.readFileSync(imagePath);
          doc.font("Helvetica-Bold").fontSize(9).fillColor(NAVY)
             .text("📐 Extracted Room Layout / Floor Plan", MARGIN, y, { width: CONTENT });
          y += 16;
          const img = doc.openImage(imgBuffer);
          const maxWidth = 300;
          const scale = maxWidth / img.width;
          const imgHeight = img.height * scale;
          doc.image(imgBuffer, MARGIN, y, { width: maxWidth, height: imgHeight });
          y += imgHeight + 12;
          doc.y = y;
        } catch (e) {
          console.warn("Could not embed room image:", e.message);
          doc.font("Helvetica").fontSize(9).fillColor(MUTED)
             .text("(Image data could not be rendered)", MARGIN, y + 16);
          y += 30;
          doc.y = y;
        }
      }

      SECTIONS.forEach(sec => {
        const pairs = sec.keys
          .filter(k => d[k] != null && String(d[k]).trim() !== "")
          .map(k => [toLabel(k), formatFieldValue(k, d[k])]);
        if (!pairs.length) return;

        const needed = 24 + pairs.length * 20 + 12;
        ensureSpace(needed);

        y = doc.y;

        fillRect(MARGIN, y, CONTENT, 20, NAVY);
        doc.font("Helvetica-Bold").fontSize(9).fillColor(WHITE)
           .text(sec.label.toUpperCase(), MARGIN + 8, y + 6, { width: CONTENT - 16 });
        y += 20;

        const COL1 = CONTENT * 0.38;
        const COL2 = CONTENT * 0.62;

        pairs.forEach(([label, value], i) => {
          const valLines = Math.ceil(value.length / 68) || 1;
          const ROW_H    = Math.max(18, valLines * 13 + 6);

          ensureSpace(ROW_H + 2);
          y = doc.y;

          fillRect(MARGIN,         y, COL1, ROW_H, i % 2 === 0 ? LGRAY : "#fafbfc");
          fillRect(MARGIN + COL1,  y, COL2, ROW_H, i % 2 === 0 ? WHITE : "#fdfdfd");
          doc.save().strokeColor(BORDER).lineWidth(0.3)
             .rect(MARGIN, y, CONTENT, ROW_H).stroke().restore();
          hLine(y + ROW_H, BORDER, 0.3);

          doc.font("Helvetica-Bold").fontSize(8.5).fillColor(MUTED)
             .text(label, MARGIN + 7, y + (ROW_H - 10) / 2,
                   { width: COL1 - 14, lineBreak: false, ellipsis: true });

          doc.font("Helvetica").fontSize(9).fillColor(TEXT)
             .text(value, MARGIN + COL1 + 7, y + 5,
                   { width: COL2 - 14, lineBreak: true });

          doc.y = y + ROW_H;
        });

        doc.y += 6;
      });

      ensureSpace(20);
      hLine(doc.y, BORDER, 0.5);
      doc.font("Helvetica").fontSize(7.5).fillColor(MUTED)
         .text(
           `Generated: ${new Date().toLocaleString("en-IN")}   |   RDS ID: ${r.id}   |   Medical Infra Facility Planning`,
           MARGIN, doc.y + 4, { width: CONTENT, align: "center" }
         );
    });

    doc.end();
  });
}

// ─── EXCEL BUILDER ────────────────────────────────────────
function buildExcel(rows) {
  const wb = XLSX.utils.book_new();

  const sh = ["Room Code","Room Name","Department","Project","Location",
               "Typology","Criticality","Infection Risk","Net Area (m²)",
               "Patient Capacity","Staff Required","Op. Hours","Status","Created"];
  const sr = rows.map(r => {
    const d = r.data || {};
    return [r.roomCode,r.roomName,r.department,r.project,
            d.location,d.roomTypology,d.criticalityLevel,d.infectionRiskCategory,
            d.netArea,d.patientCapacity,d.staffRequirement,d.operationalHours,
            r.status,r.createdAt ? new Date(r.createdAt).toLocaleDateString("en-IN") : ""];
  });
  const sumWs = XLSX.utils.aoa_to_sheet([sh,...sr]);
  sumWs["!cols"] = sh.map((_,i) => ({ wch:[12,22,20,20,20,14,13,14,12,14,12,18,10,18][i]||14 }));
  XLSX.utils.book_append_sheet(wb, sumWs, "Master Summary");

  const allKeys = new Set();
  rows.forEach(r => Object.keys(r.data||{}).forEach(k => allKeys.add(k)));
  const dk = [...allKeys];
  const fh = ["ID","Room Code","Room Name","Department","Project","Status","Created",...dk];
  const fr = rows.map(r => {
    const d = r.data||{};
    return [r.id,r.roomCode,r.roomName,r.department,r.project,r.status,r.createdAt,...dk.map(k=>d[k]??"")];
  });
  const fullWs = XLSX.utils.aoa_to_sheet([fh,...fr]);
  fullWs["!cols"] = fh.map(()=>({wch:18}));
  XLSX.utils.book_append_sheet(wb, fullWs, "Full Data");

  rows.forEach(r => {
    const d   = r.data||{};
    const nm  = `${(r.roomCode||"ROOM").replace(/[^a-zA-Z0-9]/g,"").slice(0,10)}_${String(r.id).slice(-4)}`.slice(0,31);
    const aoa = [];
    aoa.push(["ROOM DATA SHEET","","Medical College — Facility Planning",""]);
    aoa.push(["Room Code:",r.roomCode||"","Room Name:",r.roomName||""]);
    aoa.push(["Department:",r.department||"","Project:",r.project||""]);
    aoa.push(["Status:",r.status||"","Created:",r.createdAt?new Date(r.createdAt).toLocaleDateString("en-IN"):""]);
    aoa.push([]);
    SECTIONS.forEach(sec => {
      aoa.push([sec.label,"",""]);
      aoa.push(["Field","Value"]);
      let any=false;
      sec.keys.forEach(k => {
        const v=d[k];
        if(v!=null && v!==""){aoa.push([toLabel(k), formatFieldValue(k, v)]);any=true;}
      });
      if(!any) aoa.push(["(No data entered)","",""]);
      aoa.push([]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"]=[{wch:34},{wch:56},{wch:20},{wch:20}];
    XLSX.utils.book_append_sheet(wb,ws,nm);
  });
  return wb;
}

// ─── ROUTES ──────────────────────────────────────────────

// Health check
app.get("/", (_req, res) => res.json({ status:"ok", version:"4.0.0", timestamp:new Date() }));

// ─── AUTH LOGIN ───────────────────────────────────────────  ← NEW
app.post("/auth/login", (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password)
      return res.status(400).json({ error: "Email and password required" });

    const users = readUsers();
    const user  = users.find(
      u => u.email?.toLowerCase() === email.toLowerCase() && u.password === password
    );

    if (!user)
      return res.status(401).json({ error: "Invalid credentials" });

    res.json({
      user: {
        id:         String(user.id),
        name:       user.name,
        email:      user.email,
        role:       user.role,
        department: user.department || "",
      }
    });
  } catch (e) {
    console.error("Login error:", e);
    res.status(500).json({ error: "Login failed" });
  }
});


// GET all records
app.get("/data", async (req, res) => {
  try {
    let rows = await readAll();
    const { search, department, roomTypology, criticalityLevel, page=1, limit=20 } = req.query;
    if (search) {
      const q = search.toLowerCase();
      rows = rows.filter(r => {
        const d = r.data||{};
        return (
          String(d.roomName   ||r.roomName   ||"").toLowerCase().includes(q)||
          String(d.roomCode   ||r.roomCode   ||"").toLowerCase().includes(q)||
          String(d.project    ||r.project    ||"").toLowerCase().includes(q)||
          String(d.department ||r.department ||"").toLowerCase().includes(q)
        );
      });
    }
    if (department)       rows=rows.filter(r=>(r.data?.department      ||r.department||"")===department);
    if (roomTypology)     rows=rows.filter(r=>(r.data?.roomTypology    ||"")===roomTypology);
    if (criticalityLevel) rows=rows.filter(r=>(r.data?.criticalityLevel||"")===criticalityLevel);
    const total=rows.length, start=(parseInt(page)-1)*parseInt(limit);
    res.json({total,page:parseInt(page),limit:parseInt(limit),rows:rows.slice(start,start+parseInt(limit))});
  } catch(e){console.error(e);res.status(500).json({error:"Failed to read data"});}
});

// GET single record
app.get("/data/:id", async (req, res) => {
  try {
    const { data, error } = await supabase.from("rds_rooms").select("*").eq("id", req.params.id).single();
    if (error || !data) return res.status(404).json({error:"Record not found"});
    data.data = typeof data.data === "string" ? safeJson(data.data) : (data.data || {});
    res.json(data);
  } catch{res.status(500).json({error:"Failed to fetch record"});}
});

// POST save
app.post("/save", async (req, res) => {
  try {
    const newData = req.body;
    console.log("Saving RDS...");
    const roomCode = newData.roomCode || newData.data?.roomCode || "";
    if (!roomCode) return res.status(400).json({ error: "roomCode is required" });

    let imagePath = null;
    if (newData.roomImage) {
      try {
        const base64Data = newData.roomImage.replace(/^data:image\/\w+;base64,/, "");
        const imgBuffer = Buffer.from(base64Data, "base64");
        const filename = `${Date.now()}.png`;
        const filepath = path.join(IMAGE_DIR, filename);
        fs.writeFileSync(filepath, imgBuffer);
        imagePath = filepath;
      } catch (e) { console.warn("Failed to save image:", e.message); }
      delete newData.roomImage;
    }
    if (imagePath) newData.imagePath = imagePath;

    const now = Date.now();
    // Check duplicate
    const { data: existing } = await supabase.from("rds_rooms")
      .select("id").eq("roomcode", roomCode)
      .gte("id", now - 30000).limit(1);
    if (existing?.length) {
      return res.status(201).json({ message: "Saved successfully", id: existing[0].id });
    }

    const newRow = {
      id:           String(now),
      roomcode:     roomCode,
      roomname:     newData.roomName   || newData.data?.roomName   || "",
      department:   newData.department || newData.data?.department || "",
      project:      newData.project    || newData.data?.project    || "",
      createdat:    new Date().toISOString(),
      updatedat:    new Date().toISOString(),
      submittedby:  newData._submittedBy || "system",
      lasteditedby: newData._submittedBy || "system",
      status:       "submitted",
      data:         JSON.stringify(newData)
    };

    const { data: saved, error } = await supabase.from("rds_rooms").insert(newRow).select().single();
    if (error) throw error;
    saved.data = safeJson(saved.data);
    console.log(`✓ Saved id=${now} roomCode=${roomCode}`);

    // Audit log
    await logAudit({
      roomId:      now,
      roomCode,
      action:      "created",
      performedBy: newData._submittedBy || "system",
      details:     { roomName: newRow.roomname, department: newRow.department, project: newRow.project }
    });

    res.status(201).json({ message: "Saved successfully", id: now, record: saved });
  } catch (e) {
    console.error("Save error:", e);
    res.status(500).json({ error: "Save failed: " + e.message });
  }
});

// PUT update
app.put("/data/:id", async (req, res) => {
  try {
    // Fetch old data first for diff
    const { data: oldRow } = await supabase.from("rds_rooms").select("*").eq("id", req.params.id).single();
    const oldData = oldRow?.data ? (typeof oldRow.data === "string" ? safeJson(oldRow.data) : oldRow.data) : {};

    const updates = {
      roomname:     req.body.roomName   || "",
      department:   req.body.department || "",
      updatedat:    new Date().toISOString(),
      lasteditedby: req.body._editedBy || req.body._submittedBy || "system",
      data:         JSON.stringify(req.body)
    };
    const { data, error } = await supabase.from("rds_rooms").update(updates).eq("id", req.params.id).select().single();
    if (error) throw error;

    // Compute field-level diff
    const newData = req.body;
    const IGNORE  = new Set(["_editedBy","_submittedBy","roomImage","imagePath"]);
    const changes = {};
    const allKeys = new Set([...Object.keys(oldData), ...Object.keys(newData)]);
    allKeys.forEach(k => {
      if (IGNORE.has(k)) return;
      const oldVal = String(oldData[k] ?? "").trim();
      const newVal = String(newData[k] ?? "").trim();
      if (oldVal !== newVal) changes[k] = { from: oldVal || "(empty)", to: newVal || "(empty)" };
    });

    await logAudit({
      roomId:      req.params.id,
      roomCode:    data.roomcode || "",
      action:      "updated",
      performedBy: req.body._editedBy || req.body._submittedBy || "system",
      details:     {
        roomName:   updates.roomname,
        department: updates.department,
        changes
      }
    });

    res.json({message:"Updated", record: data});
  } catch{res.status(500).json({error:"Failed to update"});}
});

// DELETE
app.delete("/data/:id", async (req, res) => {
  try {
    const { data: room } = await supabase.from("rds_rooms").select("roomcode,roomname").eq("id", req.params.id).single();
    const { error } = await supabase.from("rds_rooms").delete().eq("id", req.params.id);
    if (error) throw error;
    await logAudit({
      roomId:      req.params.id,
      roomCode:    room?.roomcode || "",
      action:      "deleted",
      performedBy: "system",
      details:     { roomName: room?.roomname || "" }
    });
    res.json({ message: "Deleted" });
  } catch {
    res.status(500).json({ error: "Failed to delete" });
  }
});

// GET audit log for a room
app.get("/audit/:roomId", async (req, res) => {
  try {
    const { data, error } = await supabase
      .from("rds_audit_log")
      .select("*")
      .eq("room_id", req.params.roomId)
      .order("created_at", { ascending: false });
    if (error) throw error;
    res.json({ logs: data || [] });
  } catch (e) {
    res.status(500).json({ error: "Failed to fetch audit log: " + e.message });
  }
});

// GET stats
app.get("/stats", async (_req, res) => {
  try {
    const rows = await readAll();
    const depts={}, typologies={}, criticalities={};
    rows.forEach(r=>{
      const d=r.data||{};
      if(d.department)       depts[d.department]               =(depts[d.department]              ||0)+1;
      if(d.roomTypology)     typologies[d.roomTypology]        =(typologies[d.roomTypology]       ||0)+1;
      if(d.criticalityLevel) criticalities[d.criticalityLevel] =(criticalities[d.criticalityLevel]||0)+1;
    });
    res.json({total:rows.length,byDepartment:depts,byTypology:typologies,byCriticality:criticalities,recent:rows.slice(-5).reverse()});
  } catch{res.status(500).json({error:"Failed to compute stats"});}
});

// GET filter options
app.get("/filter-options", async (_req, res) => {
  try {
    const rows = await readAll();
    res.json({
      departments:   [...new Set(rows.map(r=>r.data?.department      ||r.department).filter(Boolean))].sort(),
      typologies:    [...new Set(rows.map(r=>r.data?.roomTypology    ).filter(Boolean))].sort(),
      criticalities: [...new Set(rows.map(r=>r.data?.criticalityLevel).filter(Boolean))].sort(),
    });
  } catch{res.status(500).json({error:"Failed to fetch options"});}
});

// ─── EXCEL EXPORTS ────────────────────────────────────────

// ── Save buffer to Supabase Storage ──────────────────────
async function saveToStorage(buf, filename, mimetype) {
  try {
    const { error } = await supabase.storage
      .from("rds-exports")
      .upload(filename, buf, { contentType: mimetype, upsert: true });
    if (error) console.warn("Storage upload warn:", error.message);
    else console.log(`✓ Saved to Supabase Storage: ${filename}`);
  } catch(e) { console.warn("Storage upload failed:", e.message); }
}

app.get("/export/excel", async (req, res) => {
  try {
    const rows = await readAll();
    if(!rows.length) return res.status(404).json({error:"No records found"});
    const filename = `RDS_All_${new Date().toISOString().slice(0,10)}.xlsx`;
    const wb=buildExcel(rows), buf=XLSX.write(wb,{type:"buffer",bookType:"xlsx"});
    saveToStorage(buf, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition",`attachment; filename="${filename}"`);
    res.send(buf);
  } catch(e){console.error(e);res.status(500).json({error:"Excel failed: "+e.message});}
});

app.get("/export/excel/:id", async (req, res) => {
  try {
    const all  = await readAll();
    const rows = all.filter(r=>String(r.id)===req.params.id);
    if(!rows.length) return res.status(404).json({error:"Record not found"});
    const filename = `RDS_${rows[0].roomcode||rows[0].id}_${new Date().toISOString().slice(0,10)}.xlsx`;
    const wb=buildExcel(rows), buf=XLSX.write(wb,{type:"buffer",bookType:"xlsx"});
    saveToStorage(buf, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition",`attachment; filename="${filename}"`);
    res.send(buf);
  } catch(e){console.error(e);res.status(500).json({error:"Excel failed: "+e.message});}
});

app.get("/export/csv", (_req, res) => res.redirect("/export/excel"));

// ─── PDF EXPORTS ──────────────────────────────────────────

app.get("/export/pdf", async (req, res) => {
  try {
    const rows = await readAll();
    if (!rows.length) return res.status(404).json({ error: "No records found. Submit at least one RDS first." });
    console.log(`Generating PDF for ${rows.length} room(s)...`);
    const buf = await buildPDF(rows);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="RDS_All_${new Date().toISOString().slice(0,10)}.pdf"`);
    res.setHeader("Content-Length", buf.length);
    res.send(buf);
  } catch(e) {
    console.error("PDF error:", e);
    res.status(500).json({ error: "PDF generation failed: " + e.message });
  }
});

app.get("/export/pdf/:id", async (req, res) => {
  try {
    const all  = await readAll();
    const rows = all.filter(r => String(r.id) === req.params.id);
    if (!rows.length) return res.status(404).json({ error: "Record not found" });
    const buf = await buildPDF(rows);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="RDS_${rows[0].roomcode || rows[0].id}.pdf"`);
    res.setHeader("Content-Length", buf.length);
    res.send(buf);
  } catch(e) {
    console.error("PDF error:", e);
    res.status(500).json({ error: "PDF generation failed: " + e.message });
  }
});

// ─── AI EXTRACT ROUTE ─────────────────────────────────────
const FIELD_LIST = `project,department,roomName,roomCode,location,roomTypology,criticalityLevel,infectionRiskCategory,isolationType,roomFunction,keyActivities,userGroups,operationalScenarios,patientCapacity,staffRequirement,peakLoad,throughput,averageStayTime,surgeCapacity,operationalHours,patientZone,staffZone,equipmentZone,cleanZone,dirtyZone,patientFlow,staffFlow,materialFlow,entryPoints,restrictedZones,mustBeAdjacent,shouldBeAdjacent,avoidAdjacency,netArea,minimumDimensions,clearances,ceilingHeight,doorType,doorSize,accessibility,floor,skirting,wallFinish,ceiling,cabinetry,worktop,specialFinishes,acuCategory,ventilationRate,acuCount,pressureControl,tempRange,humidityRange,filtrationGrade,powerLoad,normalPower,emergencyPower,ups,numberOfSockets,specialOutlets,oxygen,medicalAir,vacuum,nitrousOxide,handWash,wc,shower,plumbingSpecialSystems,hisEmr,pacs,lis,rtls,nurseCall,cctv,iotSensors,aiAnalytics,pressureRegime,isolationLevel,radiationProtection,biohazardHandling,fireSafety,emergencySystems,lightingQuality,acousticControl,privacy,patientComfort,familyInteraction,visualEnvironment,airFlowmeter,oxygenFlowmeter,suctionAdapterLowFlow,suctionBottle,oxygenFlowmeterLowFlow,trolleyProcedure,blenderAirOxygen,stoolAdjustableMobile,curtainTrackSystem,ivHook,additionalFF,infusionPumpSyringe,examinationLight,physiologicMonitor,infantIncubator,phototherapyLamp,supplyUnitCeiling,infusionPumpEnteral,infusionPumpSingleChannel,ventilatorNeonatal,additionalFE`;

const SYSTEM_PROMPT = `You are an expert at extracting Room Data Sheet (RDS) data from documents.
Extract values and return ONLY a valid JSON object using ONLY these exact keys where data is found:
[${FIELD_LIST}]
Rules:
- Omit keys with no matching data (no nulls, no empty strings)
- Number fields (netArea,patientCapacity,staffRequirement,peakLoad,powerLoad,numberOfSockets,airChangesACH,ceilingHeight,oxygen,medicalAir,vacuum,nitrousOxide,handWash,wc,shower and all equipment qty fields): return numbers only
- Yes/No fields (ups,cctv,pandemicMode): return exactly "Yes" or "No"
- For SELECT fields use ONLY these exact values:
  roomTypology: "ICU"|"Ward"|"OT"|"Emergency"|"Outpatient"|"Diagnostic"|"Laboratory"|"Pharmacy"|"Administrative"|"Support"|"NICU"|"PICU"|"CCU"|"HDU"|"Isolation"|"Other"
  criticalityLevel: "Critical"|"High"|"Medium"|"Low"|"Ancillary"
  infectionRiskCategory: "Very High"|"High"|"Medium"|"Low"|"Minimal"
  isolationType: "None"|"Contact"|"Droplet"|"Airborne"|"Protective (Reverse)"|"Combined"|"Strict Isolation"
  operationalHours: "24×7"|"Scheduled (Day only)"|"Scheduled (Day & Evening)"|"On-call"|"As Required"
  doorType: "Sliding"|"Hinged (Single)"|"Hinged (Double)"|"Automatic Sliding"|"Hermetic Sealed"|"Fire-Rated"|"Other"
  accessibility: "Full Barrier-Free Compliance"|"Partial Compliance"|"Standard"|"Not Applicable"
  pressure: "Positive (+ve)"|"Negative (-ve)"|"Neutral"|"Variable"
  filtration: "HEPA H14"|"HEPA H13"|"MERV-16"|"MERV-13"|"Standard"|"ULPA"|"Other"
  pressureRegime: "Positive (+ve)"|"Negative (-ve)"|"Neutral"|"Variable / Switchable"
  isolationLevel: "Level 1 – Standard"|"Level 2 – Enhanced"|"Level 3 – Strict"|"Level 4 – Maximum / BSL-4"|"Not Applicable"
  radiationProtection: "Not Required"|"Lead Lining Required"|"Lead Glass Windows"|"Controlled Zone"|"Supervised Zone"
  biohazardHandling: "Not Applicable"|"BSL-1"|"BSL-2"|"BSL-3"|"BSL-4"
- Pay special attention to "roomFunction" which is typically at the top of the document
- Return ONLY the JSON object, no markdown, no explanation`;

function extractRoomImageFromZip(zipBuffer) {
  try {
    const zip     = new AdmZip(zipBuffer);
    const entries = zip.getEntries();

    const blockedImages = new Set();
    entries.forEach(e => {
      if (/word\/_rels\/(header|footer)\d*\.xml\.rels$/i.test(e.entryName)) {
        try {
          const xml = e.getData().toString("utf8");
          [...xml.matchAll(/Target="[^"]*media\/([^"]+)"/gi)]
            .forEach(m => blockedImages.add(path.basename(m[1]).toLowerCase()));
        } catch (_) {}
      }
    });

    const docRelsEntry = entries.find(e => e.entryName === "word/_rels/document.xml.rels");
    const docXmlEntry  = entries.find(e => e.entryName === "word/document.xml");

    if (docRelsEntry && docXmlEntry) {
      try {
        const relsXml   = docRelsEntry.getData().toString("utf8");
        const rIdToFile = {};
        [...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="[^"]*media\/([^"]+)"/gi)]
          .forEach(m => { rIdToFile[m[1]] = path.basename(m[2]).toLowerCase(); });

        const docXml = docXmlEntry.getData().toString("utf8");
        const allRows = [...docXml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)];

        if (allRows.length > 0) {
          const firstRow = allRows[0][0];
          [...firstRow.matchAll(/r:embed="([^"]+)"/gi)]
            .forEach(m => { if (rIdToFile[m[1]]) blockedImages.add(rIdToFile[m[1]]); });
        }
      } catch (_) {}
    }

    console.log(`[img] Blocked images: ${[...blockedImages].join(", ") || "none"}`);

    const imageEntries = entries.filter(e =>
      !e.isDirectory &&
      (e.entryName.includes("word/media/") || e.entryName.includes("xl/media/")) &&
      /\.(png|jpe?g|gif|bmp|webp)$/i.test(e.entryName)
    );
    if (imageEntries.length === 0) return null;

    let bestImage = null;
    let bestScore = -Infinity;

    for (const entry of imageEntries) {
      const basename = path.basename(entry.entryName).toLowerCase();

      if (blockedImages.has(basename)) {
        console.log(`[img] SKIP blocked: ${basename}`);
        continue;
      }

      const imgData    = entry.getData();
      const fileSizeKB = imgData.length / 1024;
      let width = 0, height = 0;
      try { const d = sizeOf(imgData); width = d.width || 0; height = d.height || 0; } catch (_) {}

      if (fileSizeKB < 3) { console.log(`[img] SKIP tiny file: ${basename}`); continue; }
      if (width > 0 && height > 0) {
        if (width < 80 || height < 80)    { console.log(`[img] SKIP tiny px: ${basename}`); continue; }
        if (width * height < 20000)        { console.log(`[img] SKIP small area: ${basename}`); continue; }
        const asp = width / height;
        if (asp > 5.0 || asp < 0.2)       { console.log(`[img] SKIP extreme aspect: ${basename}`); continue; }
      }

      let score = fileSizeKB * 10;
      if (width > 0 && height > 0) {
        score += Math.sqrt(width * height) / 4;
        const asp = width / height;
        if (asp >= 0.4 && asp <= 2.5) score += 60;
      }

      console.log(`[img] CANDIDATE: ${basename} ${width}x${height} ${fileSizeKB.toFixed(1)}KB score=${score.toFixed(0)}`);
      if (score > bestScore) { bestScore = score; bestImage = { data: imgData, entryName: entry.entryName }; }
    }

    if (!bestImage) {
      console.log("[img] All filtered — picking largest non-blocked image");
      for (const entry of imageEntries) {
        const basename = path.basename(entry.entryName).toLowerCase();
        if (blockedImages.has(basename)) continue;
        const data = entry.getData();
        if (!bestImage || data.length > bestImage.data.length)
          bestImage = { data, entryName: entry.entryName };
      }
    }

    if (bestImage) {
      const ext  = path.extname(bestImage.entryName).toLowerCase();
      const mime = (ext === ".jpg" || ext === ".jpeg") ? "image/jpeg"
                 : ext === ".gif" ? "image/gif" : "image/png";
      console.log(`[img] ✅ Selected: ${bestImage.entryName} (${(bestImage.data.length / 1024).toFixed(1)} KB)`);
      return `data:${mime};base64,${bestImage.data.toString("base64")}`;
    }
  } catch (e) {
    console.warn("[img] Extraction error:", e.message);
  }
  return null;
}

app.post("/extract", async (req, res) => {
  try {
    const { type, content } = req.body;
    if (!content) return res.status(400).json({ error: "No content provided" });

    const GROQ_KEY = process.env.GROQ_API_KEY;
    if (!GROQ_KEY) return res.status(500).json({ error: "GROQ_API_KEY not set" });

    const groq = new Groq({ apiKey: GROQ_KEY });

    let textContent = "";
    let imageBase64 = null;

    if (type === "pdf") {
      try {
        textContent = await extractPdfText(content);
        console.log(`PDF extracted: ${textContent.length} chars`);
      } catch(e) {
        console.error("PDF extract error:", e.message);
        return res.status(400).json({ error: "Could not read PDF: " + e.message });
      }
      if (textContent.length < 20)
        return res.status(400).json({ error: "PDF has no readable text (scanned/image PDFs not supported)" });
    } else if (type === "word") {
      try {
        const buf = Buffer.from(content, "base64");
        const htmlResult = await mammoth.convertToHtml({ buffer: buf });
        const html = htmlResult.value || "";
        textContent = html
          .replace(/<tr[^>]*>/gi, "\n")
          .replace(/<\/tr>/gi, "")
          .replace(/<td[^>]*>/gi, " | ")
          .replace(/<\/td>/gi, "")
          .replace(/<th[^>]*>/gi, " | ")
          .replace(/<\/th>/gi, "")
          .replace(/<h[1-6][^>]*>/gi, "\n## ")
          .replace(/<\/h[1-6]>/gi, "\n")
          .replace(/<p[^>]*>/gi, "\n")
          .replace(/<\/p>/gi, "")
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
          .replace(/&nbsp;/g, " ").replace(/&#\d+;/g, " ")
          .replace(/[ \t]{2,}/g, " ")
          .replace(/\n{3,}/g, "\n\n")
          .trim();
        console.log(`Word extracted (with tables): ${textContent.length} chars`);
        imageBase64 = extractRoomImageFromZip(buf);
        if (imageBase64) console.log("Word image extracted");
      } catch(e) {
        console.error("Word extract error:", e.message);
        return res.status(400).json({ error: "Could not read Word file: " + e.message });
      }
      if (textContent.trim().length < 20)
        return res.status(400).json({ error: "No readable text found in Word document." });
    } else if (type === "excel") {
      try {
        const buf = Buffer.from(content, "base64");
        const wb = XLSX.read(buf, { type: "buffer" });
        let text = "";
        wb.SheetNames.forEach(n => {
          const ws = wb.Sheets[n];
          text += `\n--- Sheet: ${n} ---\n`;
          text += XLSX.utils.sheet_to_csv(ws);
        });
        textContent = text;
        console.log(`Excel extracted: ${textContent.length} chars`);
        imageBase64 = extractRoomImageFromZip(buf);
        if (imageBase64) console.log("Excel image extracted");
      } catch(e) {
        console.error("Excel extract error:", e.message);
        return res.status(400).json({ error: "Could not read Excel file: " + e.message });
      }
      if (textContent.trim().length < 20)
        return res.status(400).json({ error: "No readable text found in Excel document." });
    } else {
      textContent = content;
    }

    const completion = await groq.chat.completions.create({
      model: "llama-3.3-70b-versatile", max_tokens: 4000, temperature: 0,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user",   content: `Extract all Room Data Sheet fields from this document. Important instructions:

1. "roomFunction" field: Extract ALL bullet points from the room function/description section at the top. Join them with semicolons into one string. Do not truncate.

2. "additionalFF" field: Extract ALL items from "Fittings and Furniture (FF)" table as: "ItemCode: Description x Quantity" joined with semicolons. Example: "FF 150: Air flowmeter x1; FF 1465: Bracket: suction bottle x1; Curtain Track System x1 set"

3. "additionalFE" field: Extract ALL items from "Fixtures, Equipment and associated Services (FE)" table as: "ItemCode: Description x Quantity" joined with semicolons. Example: "FE 31700: Light: examination ceiling x1; FE 36050: Monitor: cardiac x1"

4. For other fields, use 80-90% confidence inference — if the document says "PACU" infer roomTypology as "Other", criticalityLevel as "High", etc.

5. Extract numeric values as numbers (netArea, patientCapacity, etc.)

Document content:
${textContent.slice(0, 28000)}` }
      ]
    });

    const raw     = completion.choices[0]?.message?.content || "{}";
    const cleaned = raw.replace(/^```[a-z]*\n?/i, "").replace(/```$/, "").trim();
    const fields  = JSON.parse(cleaned);
    console.log(`✓ Extracted ${Object.keys(fields).length} fields from ${type}`);

    res.json({ fields, image: imageBase64 });
  } catch (e) {
    console.error("Extract error:", e);
    res.status(500).json({ error: "Extraction failed: " + e.message });
  }
});

// ─── 404 / ERROR ─────────────────────────────────────────
app.use((_req, res) => res.status(404).json({ error: "Route not found" }));
app.use((err, _req, res, _next) => {
  console.error("Unhandled:", err);
  res.status(500).json({ error: "Internal error" });
});

// ─── START ───────────────────────────────────────────────
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`\n✓ RDS Backend v4.0 → http://localhost:${PORT}`);
  console.log(`  POST /auth/login          → user login`);
  console.log(`  POST /save               → submit form`);
  console.log(`  GET  /data               → list / search / filter`);
  console.log(`  GET  /export/excel       → Excel all rooms`);
  console.log(`  GET  /export/excel/:id   → Excel single room`);
  console.log(`  GET  /export/pdf         → PDF all rooms`);
  console.log(`  GET  /export/pdf/:id     → PDF single room`);
  console.log(`  POST /extract            → AI extraction + image (Word/Excel)\n`);
});