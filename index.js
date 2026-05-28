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

  const allLines = [];

  for (let i = 1; i <= doc.numPages; i++) {
    const page    = await doc.getPage(i);
    const content = await page.getTextContent();
    const vp      = page.getViewport({ scale: 1 });
    const pageH   = vp.height;

    // Collect items with normalized Y (top-down) and X positions
    const items = content.items
      .filter(item => item.str && item.str.trim())
      .map(item => ({
        str: item.str.trim(),
        x:   Math.round(item.transform[4]),
        y:   Math.round(pageH - item.transform[5]), // flip Y to top-down
      }));

    // Group items into lines by Y proximity (within 6px = same line)
    const lineMap = new Map();
    items.forEach(item => {
      const lineY = [...lineMap.keys()].find(k => Math.abs(k - item.y) <= 6);
      if (lineY !== undefined) {
        lineMap.get(lineY).push(item);
      } else {
        lineMap.set(item.y, [item]);
      }
    });

    // Sort lines top-to-bottom, items left-to-right within line
    const sortedLines = [...lineMap.entries()]
      .sort(([ya], [yb]) => ya - yb)
      .map(([, lineItems]) =>
        lineItems.sort((a, b) => a.x - b.x).map(i => i.str)
      );

    // Reconstruct structured text:
    // If line has 2+ items with x-gap > 80px → treat as label:value pair
    // Otherwise join as single line
    for (const lineItems of sortedLines) {
      if (lineItems.length === 0) continue;

      if (lineItems.length >= 2) {
        // Detect label:value structure
        const joined = lineItems.join("  |  ");
        // Also emit as "Label: Value" for better AI parsing
        const labelVal = lineItems[0] + ": " + lineItems.slice(1).join(" ");
        allLines.push(labelVal);
        allLines.push(joined);
      } else {
        allLines.push(lineItems[0]);
      }
    }
    allLines.push("--- Page Break ---");
  }

  // Post-process: merge continuation lines (items that are values continuing on next line)
  const merged = [];
  for (let i = 0; i < allLines.length; i++) {
    const line = allLines[i];
    // If line looks like a heading/section, add separator
    if (/^\d+\.\s+[A-Z]/.test(line)) {
      merged.push("\n" + line);
    } else {
      merged.push(line);
    }
  }

  return merged.join("\n").replace(/\n{3,}/g, "\n\n").trim();
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
  { label: "1. Room Identity & General Information",  keys: [
      "projectName","projectCode","type","department","departmentCode","category","categoryCode",
      "roomName","roomCode","location","roomTypology","criticalityLevel","infectionRiskCategory","isolationType"
  ]},
  { label: "2. Architectural & Spatial Requirements", keys: [
      "netArea","minimumDimension","clearance","floorToSoffitHeight","floorToCeilingHeight",
      "doorType","doorSize","accessibilityCompliance",
      "hazardousStorage","radiationShielding","vibrationIsolation","magneticShielding",
      "soundInsulation","rfShielding","equipmentMountingSupport","structuralFloorDrop","otherSpecialNeeds",
      "constructionMatrix"
  ]},
  { label: "3. Interior Finishes & Aesthetics",       keys: [
      "floor","floorSpec","skirting","walls","wallsSpec","ceiling","ceilingSpec",
      "wallProtection","wallProtectionNotes","internalGlazing","hatches","specialFinishes",
      "doorConfig","windowConfig","sanitaryFittings"
  ]},
  { label: "4. Interior Lighting & Furniture",        keys: [
      "lightingControl","lightingControlNotes",
      "lightingLevelStandard","lightingLevelTreatment","lightingLevelOther",
      "lightFitT5Fluorescent","lightFitLEDTube","lightFitLEDStrip","lightFitCompactPL",
      "lightFitLEDDownlight","lightFitBiophilic","lightFitCeilingLight","lightFitOthers","lightFittingNotes",
      "ctrlOnOff","ctrlTimer","ctrlMotionSensor","ctrlPhotosensor","ctrlLMS","ctrlOthers","lightingControlDeviceNotes",
      "furnitureLaboratoryBenches","furnitureSystemFurniture","furnitureLooseChairs","furnitureModularCabin",
      "furnitureCustom","furnitureFixedBench","furnitureLockers","furnitureCoatHooks",
      "furnitureBulletinBoard","furnitureMarkerBoard","furnitureHandRail","furnitureOthers","furnitureNotes",
      "cabBuiltInIntegrated","cabMobilePedestal","cabOverheadCabinets","cabUndercountCabinets",
      "cabOpenShelvesOverhead","cabOpenShelvesUnder","cabFullHeightCabinets","cabFullHeightShelving",
      "cabInventoryStocked","cabOthers","cabinetryNotes",
      "fumeFloorVertical1200","fumeFloorVertical1500","fumeFloorVertical1800",
      "fumeWalkIn1200","fumeWalkIn1500","fumeWalkIn1800",
      "fumePortable1200","fumePortable1500","fumePortable1800","fumeNotes"
  ]},
  { label: "5. Clinical Functionality & Workflow",    keys: [
      "roomFunction","keyActivities","userGroups","operationalScenarios",
      "patientZone","staffZone","equipmentZone","cleanZone","dirtyZone",
      "patientFlow","staffFlow","materialFlow","entryPoints","restrictedZones","medicalGasMatrix"
  ]},
  { label: "6. Capacity & Operations",                keys: [
      "patientCapacity","staffRequirement","peakLoad","throughput","averageStayTime","surgeCapacity","operationalHours"
  ]},
  { label: "7. Adjacency Matrix",                     keys: ["mustBeAdjacent","shouldBeAdjacent","avoidAdjacency"] },
  { label: "8. MEP & Engineering Systems",            keys: [
      "airChangesACH","pressure","temperature","humidity","filtration",
      "providedFanInRoom","airflowDirection","naturalVentilation","mechanicalVentilation","smokeExtraction","pandemicMode",
      "powerLoad","normalPower","emergencyPower","ups","numberOfSockets","specialOutlets","ssoMatrix","isolatorMatrix",
      "equip_dedicatedCircuit","equip_upsBackup","equip_vibrationIsolation","equip_bmsInterface",
      "equip_isolatedGrounding","equip_humidityControl","equip_remoteMonitoring","equip_voltageStabilizer",
      "equip_antiStaticFlooring","equip_fireRatedEnclosure","equip_fireAlarmInterface","equip_gasDetection",
      "oxygen","medicalAir","vacuum","nitrousOxide",
      "handWash","wc","shower","plumbingSpecialSystems"
  ]},
  { label: "9. Digital & Smart Systems",              keys: [
      "hisEmr","pacs","lis","rtls","nurseCall","cctv","iotSensors","aiAnalytics","elvMatrix","itAccessories"
  ]},
  { label: "10. Safety & Infection Control",          keys: [
      "coreSafetyMatrix","infectionControlMatrix","plumbingFixturesMatrix",
      "fireLifeSafetyMatrix","electricalSafetyMatrix","physicalSecurityMatrix","chemHazardMatrix","safetyAdditionalNotes",
      "pressureRegime","isolationLevel","radiationProtection","biohazardHandling","fireSafety","emergencySystems"
  ]},
  { label: "11. Stakeholder Experience",              keys: [
      "lightingQuality","lightingNotes","acousticControl","acousticNotes",
      "thermalOdorControl","thermalOdorNotes","patientComfort","patientComfortNotes",
      "privacy","familyInteraction","familyInteractionNotes","visualEnvironment","visualEnvironmentNotes",
      "biophiliaHealingEnvironment","biophiliaHealingNotes",
      "technologyInfotainment","technologyInfotainmentNotes","infectionControlHygiene","infectionControlHygieneNotes"
  ]},
  { label: "12. Fittings, Fixtures & Equipment",      keys: [
      "airFlowmeter","oxygenFlowmeter","suctionAdapterLowFlow","suctionBottle","oxygenFlowmeterLowFlow",
      "trolleyProcedure","blenderAirOxygen","stoolAdjustableMobile","curtainTrackSystem","ivHook",
      "patientFurniture","staffVisitorFurniture","storageFurniture",
      "wallMountedDispensers","wasteBins","additionalFF",
      "infusionPumpSyringe","examinationLight","physiologicMonitor","infantIncubator","phototherapyLamp",
      "supplyUnitCeiling","infusionPumpEnteral","infusionPumpSingleChannel","ventilatorNeonatal",
      "medicalEquipments","wallMountedDiagnostics","itCommunicationHardware","nurseCallSystems","additionalFE"
  ]},
  { label: "13. Waste Management",                    keys: [
      "wmBiohazard","wmRadioactive","wmFlammableSolvent","wmChemicalWaste","wmHumanAnatomical",
      "wmMicrobiologyWaste","wmWasteSharps","wmCytotoxicDrugs",
      "wmSoiledWaste","wmSolidWaste","wmLiquidWaste","wmDiscardedContainers",
      "wmUsedOil","wmEwaste","wmConfidentialPaper","wmFoodPantryWaste",
      "wmOthers1","wmOthers2","wmNotes"
  ]},
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
    { id: "EMP001", name: "Parth", email: "parth.patel4@adani.com", password: "123",  role: "admin",    department: "administration" },
    { id: "EMP002", name: "Ravi",  email: "ravi@rds.med",  password: "1234", role: "reviewer", department: "clinical"       },
    { id: "EMP003", name: "Healthcare",  email: "healthcare@rds.med",  password: "123", role: "reviewer", department: "clinical"       },
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

// ── Clean room name for use in filename ──────────────────
function cleanName(row) {
  const name = row.roomname || row.roomName ||
    (row.data ? (typeof row.data === "string" ? JSON.parse(row.data) : row.data)?.roomName : "") || "Room";
  return name.replace(/[^a-zA-Z0-9\s\-_]/g, "").replace(/\s+/g, "_").slice(0, 40);
}

app.get("/export/excel", async (req, res) => {
  try {
    const rows = await readAll();
    if(!rows.length) return res.status(404).json({error:"No records found"});
    const date     = new Date().toISOString().slice(0,10);
    const filename = `RDS_All_Rooms_${date}.xlsx`;
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
    const date     = new Date().toISOString().slice(0,10);
    const name     = cleanName(rows[0]);
    const filename = `RDS_${name}_${date}.xlsx`;
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
    const date     = new Date().toISOString().slice(0,10);
    const filename = `RDS_All_Rooms_${date}.pdf`;
    const buf = await buildPDF(rows);
    saveToStorage(buf, filename, "application/pdf");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
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
    const date     = new Date().toISOString().slice(0,10);
    const name     = cleanName(rows[0]);
    const filename = `RDS_${name}_${date}.pdf`;
    const buf = await buildPDF(rows);
    saveToStorage(buf, filename, "application/pdf");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Length", buf.length);
    res.send(buf);
  } catch(e) {
    console.error("PDF error:", e);
    res.status(500).json({ error: "PDF generation failed: " + e.message });
  }
});

// ─── AI EXTRACT ROUTE ─────────────────────────────────────
const FIELD_LIST = `projectName,projectCode,type,department,departmentCode,category,categoryCode,roomName,roomCode,location,roomTypology,criticalityLevel,infectionRiskCategory,isolationType,netArea,minimumDimension,clearance,floorToSoffitHeight,floorToCeilingHeight,doorType,doorSize,accessibilityCompliance,hazardousStorage,radiationShielding,vibrationIsolation,magneticShielding,soundInsulation,rfShielding,equipmentMountingSupport,structuralFloorDrop,otherSpecialNeeds,constructionMatrix,floor,floorSpec,skirting,walls,wallsSpec,ceiling,ceilingSpec,wallProtection,wallProtectionNotes,internalGlazing,hatches,specialFinishes,doorConfig,windowConfig,sanitaryFittings,lightingControl,lightingControlNotes,lightingLevelStandard,lightingLevelTreatment,lightingLevelOther,lightFitT5Fluorescent,lightFitLEDTube,lightFitLEDStrip,lightFitCompactPL,lightFitLEDDownlight,lightFitBiophilic,lightFitCeilingLight,lightFitOthers,lightFittingNotes,ctrlOnOff,ctrlTimer,ctrlMotionSensor,ctrlPhotosensor,ctrlLMS,ctrlOthers,lightingControlDeviceNotes,furnitureLaboratoryBenches,furnitureSystemFurniture,furnitureLooseChairs,furnitureModularCabin,furnitureCustom,furnitureFixedBench,furnitureLockers,furnitureCoatHooks,furnitureBulletinBoard,furnitureMarkerBoard,furnitureHandRail,furnitureOthers,furnitureNotes,cabBuiltInIntegrated,cabMobilePedestal,cabOverheadCabinets,cabUndercountCabinets,cabOpenShelvesOverhead,cabOpenShelvesUnder,cabFullHeightCabinets,cabFullHeightShelving,cabInventoryStocked,cabOthers,cabinetryNotes,fumeFloorVertical1200,fumeFloorVertical1500,fumeFloorVertical1800,fumeWalkIn1200,fumeWalkIn1500,fumeWalkIn1800,fumePortable1200,fumePortable1500,fumePortable1800,fumeNotes,roomFunction,keyActivities,userGroups,operationalScenarios,patientZone,staffZone,equipmentZone,cleanZone,dirtyZone,patientFlow,staffFlow,materialFlow,entryPoints,restrictedZones,medicalGasMatrix,patientCapacity,staffRequirement,peakLoad,throughput,averageStayTime,surgeCapacity,operationalHours,mustBeAdjacent,shouldBeAdjacent,avoidAdjacency,airChangesACH,pressure,temperature,humidity,filtration,providedFanInRoom,airflowDirection,naturalVentilation,mechanicalVentilation,smokeExtraction,pandemicMode,powerLoad,normalPower,emergencyPower,ups,numberOfSockets,specialOutlets,ssoMatrix,isolatorMatrix,equip_dedicatedCircuit,equip_upsBackup,equip_vibrationIsolation,equip_bmsInterface,equip_isolatedGrounding,equip_humidityControl,equip_remoteMonitoring,equip_voltageStabilizer,equip_antiStaticFlooring,equip_fireRatedEnclosure,equip_fireAlarmInterface,equip_gasDetection,oxygen,medicalAir,vacuum,nitrousOxide,handWash,wc,shower,plumbingSpecialSystems,hisEmr,pacs,lis,rtls,nurseCall,cctv,iotSensors,aiAnalytics,elvMatrix,itAccessories,coreSafetyMatrix,infectionControlMatrix,plumbingFixturesMatrix,fireLifeSafetyMatrix,electricalSafetyMatrix,physicalSecurityMatrix,chemHazardMatrix,safetyAdditionalNotes,pressureRegime,isolationLevel,radiationProtection,biohazardHandling,fireSafety,emergencySystems,lightingQuality,lightingNotes,acousticControl,acousticNotes,thermalOdorControl,thermalOdorNotes,patientComfort,patientComfortNotes,privacy,familyInteraction,familyInteractionNotes,visualEnvironment,visualEnvironmentNotes,biophiliaHealingEnvironment,biophiliaHealingNotes,technologyInfotainment,technologyInfotainmentNotes,infectionControlHygiene,infectionControlHygieneNotes,airFlowmeter,oxygenFlowmeter,suctionAdapterLowFlow,suctionBottle,oxygenFlowmeterLowFlow,trolleyProcedure,blenderAirOxygen,stoolAdjustableMobile,curtainTrackSystem,ivHook,patientFurniture,staffVisitorFurniture,storageFurniture,wallMountedDispensers,wasteBins,additionalFF,infusionPumpSyringe,examinationLight,physiologicMonitor,infantIncubator,phototherapyLamp,supplyUnitCeiling,infusionPumpEnteral,infusionPumpSingleChannel,ventilatorNeonatal,medicalEquipments,wallMountedDiagnostics,itCommunicationHardware,nurseCallSystems,additionalFE,wmBiohazard,wmRadioactive,wmFlammableSolvent,wmChemicalWaste,wmHumanAnatomical,wmMicrobiologyWaste,wmWasteSharps,wmCytotoxicDrugs,wmSoiledWaste,wmSolidWaste,wmLiquidWaste,wmDiscardedContainers,wmUsedOil,wmEwaste,wmConfidentialPaper,wmFoodPantryWaste,wmOthers1,wmOthers2,wmNotes`;

const SYSTEM_PROMPT = `You are an expert at extracting Room Data Sheet (RDS) data from documents.
Extract values and return ONLY a valid JSON object using ONLY these exact keys where data is found:
[${FIELD_LIST}]
Rules:
- Omit keys with no matching data (no nulls, no empty strings)
- Number fields (netArea,patientCapacity,staffRequirement,peakLoad,powerLoad,numberOfSockets,airChangesACH,ceilingHeight,oxygen,medicalAir,vacuum,nitrousOxide,handWash,wc,shower and all equipment qty fields): return numbers only
- Yes/No fields return exactly "Yes" or "No": ups,cctv,pandemicMode,internalGlazing,hatches,hazardousStorage,radiationShielding,vibrationIsolation,magneticShielding,soundInsulation,rfShielding,equipmentMountingSupport,providedFanInRoom,airflowDirection,naturalVentilation,mechanicalVentilation,smokeExtraction,equip_dedicatedCircuit,equip_upsBackup,equip_vibrationIsolation,equip_bmsInterface,equip_isolatedGrounding,equip_humidityControl,equip_remoteMonitoring,equip_voltageStabilizer,equip_antiStaticFlooring,equip_fireRatedEnclosure,equip_fireAlarmInterface,equip_gasDetection,wmBiohazard,wmRadioactive,wmFlammableSolvent,wmChemicalWaste,wmHumanAnatomical,wmMicrobiologyWaste,wmWasteSharps,wmCytotoxicDrugs,wmSoiledWaste,wmSolidWaste,wmLiquidWaste,wmDiscardedContainers,wmUsedOil,wmEwaste,wmConfidentialPaper,wmFoodPantryWaste,lightFitT5Fluorescent,lightFitLEDTube,lightFitLEDStrip,lightFitCompactPL,lightFitLEDDownlight,lightFitBiophilic,lightFitCeilingLight,lightFitOthers,ctrlOnOff,ctrlTimer,ctrlMotionSensor,ctrlPhotosensor,ctrlLMS,ctrlOthers,furnitureLaboratoryBenches,furnitureSystemFurniture,furnitureLooseChairs,furnitureModularCabin,furnitureCustom,furnitureFixedBench,furnitureLockers,furnitureCoatHooks,furnitureBulletinBoard,furnitureMarkerBoard,furnitureHandRail,furnitureOthers,cabBuiltInIntegrated,cabMobilePedestal,cabOverheadCabinets,cabUndercountCabinets,cabOpenShelvesOverhead,cabOpenShelvesUnder,cabFullHeightCabinets,cabFullHeightShelving,cabInventoryStocked,cabOthers,fumeFloorVertical1200,fumeFloorVertical1500,fumeFloorVertical1800,fumeWalkIn1200,fumeWalkIn1500,fumeWalkIn1800,fumePortable1200,fumePortable1500,fumePortable1800
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
  accessibilityCompliance: "Full Barrier-Free Compliance"|"Partial Compliance"|"Standard"|"Not Applicable"
  lightingQuality: "Standard"|"Enhanced"|"Specialist"|"Biodynamic / Circadian"
  acousticControl: "Standard"|"Enhanced Acoustic Treatment"|"Full Acoustic Isolation"|"Not Applicable"
  thermalOdorControl: "Standard HVAC"|"Enhanced Odour Extraction"|"Negative Pressure Odour Control"|"Not Applicable"
  privacy: "Open Plan — No Privacy Screening"|"Partial Privacy — Curtains / Screens"|"Full Visual Privacy — Solid Partitions"|"Acoustic Privacy — Sound Attenuation"|"Full Privacy — Visual + Acoustic"
  familyInteraction: "Not Applicable"|"Waiting Area Access Only"|"Bedside Family Zone"|"Family Participation in Care"|"Dedicated Family Room / Lounge"
  visualEnvironment: "Standard — Neutral Palette, No Feature Treatment"|"Wayfinding Colour Coding — Departmental Colour Scheme"|"Artwork & Murals — Feature Wall or Ceiling"|"Nature-Inspired — Biophilic Imagery, Textures, Patterns"|"Full Sensory Environment — Colour + Art + Nature + Lighting"
  biophiliaHealingEnvironment: "Not Applicable — No Biophilic Elements Required"|"Views to Nature — External Garden, Courtyard or Sky"|"Indoor Plants — Potted or Planter Arrangements"|"Living Wall — Vertical Green Feature"|"Natural Materials — Timber Accents, Stone Features"|"Nature-Inspired Artwork & Photographic Murals"|"Circadian / Biodynamic Lighting (Human-Centric Lighting)"|"Full Biophilic Design Package — Views + Plants + Materials + Biodynamic Light"
  technologyInfotainment: "Not Required"|"Bedside Entertainment Screen — TV / Streaming"|"Patient Education & Information System"|"Digital Wayfinding Display (Corridor / Room Entry)"|"Interactive Patient Portal — Bedside Tablet / Screen"|"Smart Room Automation — Lighting, Climate & Blinds Control"|"Full Smart Room + Entertainment + Interactive Patient Portal"
  infectionControlHygiene: "Standard — ABHR Dispenser at Entry (Staff & Visitor Accessible)"|"Enhanced — ABHR at Entry + Bedside + Toilet"|"Contactless Experience — Sensor Taps, Auto Doors, Touchless Dispensers"|"Visible Hygiene Stations — Prominently Positioned ABHR + Signage"|"Antimicrobial Surface Materials (Copper, Silver-Ion Finishes)"|"Full Hygiene-by-Design — Contactless + Antimicrobial + Seamless Surfaces + Signage"
  lightingLevelStandard: "Not Specified"|"50–100 lux — Corridor / Low-activity"|"150–200 lux — General Ward / Waiting"|"300 lux — Standard Office / Admin"|"500 lux — Clinical / Examination"|"750 lux — Procedure Room"|"1000 lux — Operating / High-precision"|"Custom (specify in notes)"
  lightingLevelTreatment: "Not Applicable"|"300 lux — General Task"|"500 lux — Clinical Examination"|"750–1000 lux — Procedure / Surgical Field"|"10 000–100 000 lux — Surgical Light (Examination Luminaire)"|"Custom (specify in notes)"
- Lighting inference rules: if document mentions "dimmer" or "dimming" → lightingControl includes "Dimming Control"; "PIR" or "motion" → ctrlMotionSensor "Yes"; "photocell" or "daylight" → ctrlPhotosensor "Yes"; "BMS" or "touch panel" or "iPad" → ctrlLMS "Yes"; "T5" → lightFitT5Fluorescent "Yes"; "LED strip" → lightFitLEDStrip "Yes"; "downlight" or "down-light" → lightFitLEDDownlight "Yes"; "biophilic" or "stretch ceiling" → lightFitBiophilic "Yes"
- Furniture inference rules: if document mentions "workstation" or "workbench" → furnitureSystemFurniture "Yes"; "locker" → furnitureLockers "Yes"; "whiteboard" or "marker board" → furnitureMarkerBoard "Yes"; "handrail" or "grab rail" → furnitureHandRail "Yes"; "countertop" or "bench top" → furnitureFixedBench "Yes"; "overhead cabinet" → cabOverheadCabinets "Yes"; "undercounter" → cabUndercountCabinets "Yes"; "full height cabinet" → cabFullHeightCabinets "Yes"; "mobile pedestal" → cabMobilePedestal "Yes"
- Fume cupboard inference: if document mentions "fume hood" or "fume cupboard" → set the appropriate fumeFloorVertical/fumeWalkIn/fumePortable field to "Yes" based on size mentioned
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

1. The document uses "Label: Value" format on each line — extract accordingly.

2. "roomFunction" field: Extract ALL bullet points / activity lines. Join with semicolons. Do not truncate.

3. "patientFurniture" field: Extract all patient furniture items (beds, over-bed tables, bedside lockers, recliners) as a descriptive string.

4. "staffVisitorFurniture" field: Extract all staff and visitor furniture items (chairs, desks, workstations) as a descriptive string.

5. "storageFurniture" field: Extract all storage items (trolleys, cabinets, shelving) as a descriptive string.

6. "wallMountedDispensers" field: Extract all dispenser items (sanitiser, soap, paper towel, glove dispensers) as a descriptive string.

7. "wasteBins" field: Extract all waste bin/container items (clinical, general, sharps containers) as a descriptive string.

8. "additionalFF" field: Extract any remaining fittings/furniture items not captured above as: "Code: Description xQty" joined with semicolons.

9. "wallMountedDiagnostics" field: Extract wall-mounted diagnostic items (sphygmomanometer, pulse oximeter, ophthalmoscope) as a descriptive string.

10. "itCommunicationHardware" field: Extract IT/comms hardware (PC, monitor, scanner, printer, VOIP, network points) as a descriptive string.

11. "nurseCallSystems" field: Extract nurse call/emergency call items (bedhead unit, pull cord, annunciator) as a descriptive string.

12. "medicalEquipments" field: Extract all other medical equipment items (ECG, ultrasound, defibrillator) as a descriptive string.

13. "additionalFE" field: Extract any remaining fixture/equipment items not captured above as: "Code: Description xQty" joined with semicolons.

14. Numeric conversions — these values may appear as text like "One (1)", "1 No.", "2 Nos." — always extract as a plain number, never negative:
   netArea, patientCapacity, staffRequirement, peakLoad, powerLoad, numberOfSockets,
   oxygen, medicalAir, vacuum, nitrousOxide, handWash, wc, shower, ceilingHeight,
   airFlowmeter, oxygenFlowmeter, suctionAdapterLowFlow, suctionBottle,
   oxygenFlowmeterLowFlow, trolleyProcedure, blenderAirOxygen, stoolAdjustableMobile,
   curtainTrackSystem, ivHook, infusionPumpSyringe, examinationLight, physiologicMonitor,
   infantIncubator, phototherapyLamp, supplyUnitCeiling, infusionPumpEnteral,
   infusionPumpSingleChannel, ventilatorNeonatal

15. For select fields match closest option (e.g. "24/7" → "24×7", "negative" → "Negative (-ve)").

16. Use 80-90% inference for unlisted fields when context is clear.

17. INTERIOR LIGHTING — scan for any mention of lighting control, lux levels, dimming, switching, luminaire types, sensors, or lighting management systems. Map to:
   - lightingControl: overall control strategy
   - lightingLevelStandard / lightingLevelTreatment: lux requirements
   - lightFit* fields: luminaire types present (Yes/No each)
   - ctrl* fields: control device types present (Yes/No each)
   - lightFittingNotes / lightingControlDeviceNotes / lightingLevelOther: any free-text lighting notes

18. FURNITURE & CABINETRY — scan for any mention of benches, workstations, chairs, lockers, boards, handrails, cabinets, shelves, pedestals. Map to:
   - furniture* fields: Yes/No for each furniture/fixture type
   - cab* fields: Yes/No for each cabinetry/shelving type
   - furnitureNotes / cabinetryNotes: any free-text notes

19. FUME CUPBOARDS — scan for fume hood, fume cupboard, ductless hood mentions. Map to:
    - fumeFloorVertical* / fumeWalkIn* / fumePortable* fields: Yes/No based on type and size (1200/1500/1800mm)
    - fumeNotes: any additional fume cupboard specifications

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

// ─────────────────────────────────────────────────────────────────────────────
// AI VALIDATION ENGINE — paste this block into backend/index.js
// PLACE IT just before the "404 / ERROR" section at the bottom of index.js
// ─────────────────────────────────────────────────────────────────────────────
//
// SETUP REQUIRED:
//   1. npm install node-fetch  (if not already installed)
//   2. Add to Render env vars:  TAVILY_API_KEY=tvly-xxxxxxxxxxxxxxxx
//   3. Run this SQL in Supabase:
//        CREATE TABLE IF NOT EXISTS rds_validations (
//          id          uuid DEFAULT gen_random_uuid() PRIMARY KEY,
//          room_id     text NOT NULL,
//          room_code   text,
//          report      jsonb,
//          created_at  timestamptz DEFAULT now()
//        );
// ─────────────────────────────────────────────────────────────────────────────

const fetch = (...args) => import("node-fetch").then(({ default: f }) => f(...args));
 
// ── 13 Agent Definitions ──────────────────────────────────────────────────────
const AGENT_CONFIGS = [
  {
    id: 1,
    section: "Room Identity & General Information",
    searchQueries: (data) => [
      `hospital ${data.roomTypology || "clinical room"} criticality infection risk classification standards NABH JCI 2025`,
      `isolation type ${data.isolationType || "clinical room"} hospital infection control standards India 2025`
    ],
    systemPrompt: `You are a senior healthcare facility planner specializing in clinical classification standards (NABH, JCI, HTM, HBN, FGI).
 
IMPORTANT SCOPE RULE: This section has two subsections — "Project & Room Identification" (projectName, projectCode, type, department, departmentCode, category, categoryCode, roomName, roomCode) and "Clinical Classification" (criticalityLevel, infectionRiskCategory, isolationType).
 
You MUST ONLY validate and suggest improvements for the CLINICAL CLASSIFICATION fields: criticalityLevel, infectionRiskCategory, and isolationType. Do NOT comment on or suggest changes to any Project & Room Identification fields — those are user-defined administrative codes.
 
For the three clinical classification fields, validate whether the selected values are appropriate for the room type, check for mismatches (e.g. a critical care room with Low criticality), and suggest better options with clinical rationale. Always include "In place of X, consider Y because..." style replacement suggestions.`
  },
  {
    id: 2,
    section: "Architectural & Spatial Requirements",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} net area minimum dimensions clearance FGI guidelines NABH 2025`,
      `hospital ${data.roomTypology || "clinical"} room door type acoustic wall construction standards 2025`
    ],
    systemPrompt: `You are an expert healthcare architect specializing in clinical space planning per FGI Guidelines 2022, HTM, HBN, NABH, and NBC India.
 
Analyse ALL filled fields across ALL subsections of this section:
- Spatial Requirements: netArea, minimumDimension, clearance, floorToSoffitHeight, floorToCeilingHeight, doorType, doorSize, accessibilityCompliance
- Special Construction: hazardousStorage, radiationShielding, vibrationIsolation, magneticShielding, soundInsulation, rfShielding, equipmentMountingSupport, structuralFloorDrop, otherSpecialNeeds
- Construction Details (constructionMatrix): wall types (wall1–wall4), ceiling, floor, skirting, visionPanel, acoustics — validate materials for each element
 
For every filled field, check compliance with current standards. For every issue or outdated choice, always include: "In place of [current], consider [modern alternative] because [clinical/technical reason]." Cover the entire section, not just one part.`
  },
  {
    id: 3,
    section: "Interior Finishes & Aesthetics",
    searchQueries: (data) => [
      `best hospital flooring wall ceiling finishes ${data.roomTypology || "clinical room"} infection control antimicrobial 2025`,
      `hospital door window sanitary fittings wall protection latest standards healthcare design 2025`
    ],
    systemPrompt: `You are an expert in healthcare interior specifications — finishes, surfaces, infection control, doors, windows, and sanitary fittings.
 
Analyse ALL filled fields across ALL subsections:
- Room Finishes: floor, floorSpec, skirting, walls, wallsSpec, ceiling, ceilingSpec — validate material suitability for room type and infection risk
- Wall Protection & Surface Features: wallProtection, wallProtectionNotes, internalGlazing, hatches, specialFinishes
- Doors (doorConfig): door type, size, material, fire rating, access control hardware
- Windows (windowConfig): exterior and internal viewing windows, glazing type, provisions
- Sanitary Fittings (sanitaryFittings): all selected fittings for clinical appropriateness
 
For every field with a value, assess whether current choices meet modern healthcare standards. Always suggest: "In place of [current finish/material], consider [modern alternative] — e.g. instead of standard vinyl, specify Heterogeneous PUR-coated vinyl with welded seams and integral coved skirting for full infection control." Compare against latest global materials available in 2025.`
  },
  {
    id: 4,
    section: "Interior Lighting & Furniture",
    searchQueries: (data) => [
      `hospital ${data.roomTypology || "clinical room"} lighting lux levels LED circadian tunable white CIBSE HTM 08-03 2025`,
      `healthcare cabinetry furniture fume cupboard ergonomic clinical standards latest 2025`
    ],
    systemPrompt: `You are a healthcare lighting, furniture, and laboratory equipment specialist (CIBSE, IES, HTM 08-03, NABH).
 
Analyse ALL filled fields across ALL subsections:
- Lighting Control & Optimization: lightingControl, lightingControlNotes — is the control strategy appropriate?
- Lighting Levels: lightingLevelStandard, lightingLevelTreatment, lightingLevelOther — do lux levels match room type and clinical task?
- Lighting Fitting Type: T5 fluorescent, LED tube, LED strip, compact PL, LED downlight, biophilic/stretch ceiling, ceiling light — flag obsolete fittings
- Lighting Control Devices: on/off switches, timers, PIR/occupancy sensors, photosensors, LMS — suggest smart alternatives
- Furniture & Room Fixtures: all yes/no furniture items and furnitureNotes — validate clinical appropriateness
- Cabinetry & Shelving: all cabinet and shelf items, cabinetryNotes — check for clinical storage standards
- Fume Cupboards: all fume cupboard types and fumeNotes — validate selection for room type
 
For every obsolete fitting or inadequate spec, state: "In place of [current], consider [modern alternative] — e.g. replace T5 fluorescent with LED panels with tunable white (2700K–6500K) for circadian support and 60% energy saving." Cover all filled fields.`
  },
  {
    id: 5,
    section: "Clinical Functionality & Workflow",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} clinical workflow functional zones patient staff flow lean design 2025`,
      `medical gas outlets ${data.roomTypology || "hospital room"} HTM 02-01 NABH requirements 2025`
    ],
    systemPrompt: `You are a clinical workflow expert and medical gas systems specialist (HTM 02-01, NABH, FGI).
 
Analyse ALL filled fields across ALL subsections:
- Functionality & Workflow: roomFunction, keyActivities, userGroups, operationalScenarios — are they complete and clinically sound?
- Functional Zones: patientZone, staffZone, equipmentZone, cleanZone, dirtyZone — validate zone separation and infection control logic
- Circulation: patientFlow, staffFlow, materialFlow — check for cross-contamination risks and lean flow principles
- Access Control: entryPoints, restrictedZones — validate access control strategy
- Medical Gas Matrix (medicalGasMatrix): for each gas (vacuum, oxygen, CO₂, N₂O, medAir4, surgAir7, agss, compAir, liqN2, png, lmo, oog) — validate outlet types, quantities, locations, and mounting heights against HTM 02-01 and room type requirements
 
For every gap or outdated configuration, state: "In place of [current approach], consider [modern alternative] — e.g. replace fixed gas pendants with ceiling-mounted medical supply units (Draeger/Ondal) for flexible layout and infection control." Cover all filled fields.`
  },
  {
    id: 6,
    section: "Capacity & Operations",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} patient capacity bed ratio staffing standards NABH 2025`,
      `healthcare surge capacity throughput operational hours planning best practices 2025`
    ],
    systemPrompt: `You are a healthcare operations and capacity planning expert (NABH, JCI, NHS England benchmarks).
 
Analyse ALL filled fields:
- patientCapacity: is bed/user count appropriate for room type and area?
- staffRequirement: validate nurse-to-patient ratio per shift against current standards
- peakLoad: assess max occupancy planning
- throughput: patients/day — is it realistic for the room type?
- averageStayTime: validate against clinical benchmarks
- surgeCapacity: is expandability logic adequate?
- operationalHours: 24×7 vs scheduled — appropriate for room criticality?
 
For every field, compare against current Indian and international healthcare benchmarks. State: "In place of [current value/approach], consider [recommended standard] — e.g. for a Critical Care room, nurse-to-patient ratio should be 1:1 or 1:2 per NABH ICU standards, not 1:4." Cover all filled fields.`
  },
  {
    id: 7,
    section: "Adjacency Matrix",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} adjacency functional relationships hospital master planning FGI 2025`,
      `healthcare department adjacency infection control workflow natural light acoustic separation 2025`
    ],
    systemPrompt: `You are a hospital master planner and healthcare architect specializing in spatial relationships (FGI, HTM, HBN, NABH).
 
Analyse ALL filled fields across ALL subsections:
- Functional Adjacency: primaryFunctionalAdjacency, secondaryFunctionalAdjacency, negativeAdjacency, equipmentSharingAccess
- Departmental & Site Access: departmentalGrouping, externalAccessSite
- Environmental Requirements: naturalLightViewRequirement, acousticAdjacencySeparation, visualAdjacencyLineOfSight
- Infrastructure & Logistics: verticalCoreProximity, materialSupplyRoute
- Additional Notes: adjacencyOthers
 
For every field, validate the adjacency logic for the room type. Flag any clinically unsafe adjacencies. State: "In place of [current adjacency choice], consider [better relationship] — e.g. for an ICU, direct adjacency to OT via dedicated corridor is preferred over same-floor-only proximity." Cover all filled fields.`
  },
  {
    id: 8,
    section: "MEP & Engineering Services",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} HVAC ACH pressure temperature humidity filtration ASHRAE 170 2025`,
      `hospital electrical medical gas plumbing standards energy efficiency smart MEP 2025`
    ],
    systemPrompt: `You are a senior MEP engineer specializing in healthcare facilities (ASHRAE 170-2021, HTM 02-01, NBC India, NABH, IEC 60601).
 
Analyse ALL filled fields across ALL subsections:
- HVAC: airChangesACH, pressure, temperature, humidity, filtration, providedFanInRoom, airflowDirection, naturalVentilation, mechanicalVentilation, smokeExtraction, pandemicMode — validate all against ASHRAE 170 and room type
- Electrical: powerLoad, normalPower, emergencyPower, ups, numberOfSockets, specialOutlets, ssoMatrix (socket outlets by location/source), isolatorMatrix (isolators by location/source/rating)
- Equipment Related Provision: all 12 yes/no items (dedicatedCircuit, upsBackup, vibrationIsolation, bmsInterface, isolatedGrounding, humidityControl, remoteMonitoring, voltageStabilizer, antiStaticFlooring, fireRatedEnclosure, fireAlarmInterface, gasDetection)
- Medical Gases: oxygen outlets, medicalAir outlets, vacuum/AGSS outlets, nitrousOxide outlets
- Plumbing: handWash, wc, shower, plumbingSpecialSystems
 
For every field with a value, validate against current standards and state: "In place of [current spec], consider [modern alternative] — e.g. replace MERV-13 filtration with HEPA H14 terminal filters for a Critical Care room per ASHRAE 170-2021 Table 7.1." Cover all filled fields.`
  },
  {
    id: 9,
    section: "Digital & Smart Systems",
    searchQueries: (data) => [
      `hospital smart room IoT digital systems HIS EMR PACS RTLS ${data.roomTypology || "clinical"} 2025`,
      `healthcare ELV systems nurse call CCTV access control IT accessories latest innovations 2025`
    ],
    systemPrompt: `You are a healthcare IT and smart building specialist (HL7, FHIR, HIMSS, ISO 80001).
 
Analyse ALL filled fields across ALL subsections:
- Core Clinical Systems: hisEmr, pacs, lis, rtls, cctv, iotSensors, aiAnalytics — validate integration completeness and suggest modern platforms
- ELV Matrix (elvMatrix): all configured ELV systems (data, voice, nurse call, fire alarm, CCTV, access control, AV, IPTV, public address, etc.) — validate quantities per location against room type
- IT & Digital Accessories (itAccessories): monitorSystem, printer, vitalEquipment, barcodePrinter, laptop, kiosk, multiFunctionPrinter, scanner, highSpeedPrinter, queueManagement, tv, networkSwitch, lanHub — validate selection for room type
 
For every field, compare against 2025 healthcare digital standards and state: "In place of [current system], consider [modern alternative] — e.g. replace standalone PACS with cloud-based VNA (Vendor Neutral Archive) integrated with AI-powered diagnostic tools for faster reads and remote access." Cover all filled fields.`
  },
  {
    id: 10,
    section: "Safety & Infection Control",
    searchQueries: (data) => [
      `hospital ${data.roomTypology || "clinical room"} infection control isolation pressure HEPA ACH standards WHO CDC NABH 2025`,
      `healthcare fire safety anti-ligature nurse call electrical safety anti-static flooring standards 2025`
    ],
    systemPrompt: `You are an infection control, patient safety, and fire safety specialist (WHO, CDC, NABH, JCI, BS EN, NBC India).
 
Analyse ALL filled fields across ALL subsections:
- Core Safety Parameters (coreSafetyMatrix): pressureRegime, isolationLevel, radiationProtect, biohazard
- Infection Control & Air Quality (infectionControlMatrix): handHygiene, hepa, ach, tempHumidity, uvDisinfect, anteRoom, sharps, bmwWaste, fumigation
- Plumbing Safety Fixtures (plumbingFixturesMatrix): eyeWash, eShower, combined, kneeFix, footFix, wristBlade, sprayHose, elbowFaucet
- Fire & Life Safety (fireLifeSafetyMatrix): fireSafety, emergencyPower, sprinkler, smokeDetect, raisedFloor, cleanGas, smokeCompart, fireDoors, evacAids, gasLeak, dampers
- Electrical & Equipment Safety (electricalSafetyMatrix): antiStatic, emiShield, elecSafeCls
- Physical Safety & Security (physicalSecurityMatrix): antiLigature, slipResist, nurseCall, panicAlarm, cctv, accessCtrl, seismic, ppeStorage, spillContain
- Chemical & Hazardous Material Safety (chemHazardMatrix): coshh, cryogenic, cytotoxic
- Additional Safety Notes: safetyAdditionalNotes
 
For every field, validate against the latest guidelines and state: "In place of [current selection], consider [upgraded option] — e.g. replace Standard Terminal Clean with HPV Decontamination Ready provision for a High infection risk room per WHO IPC guidelines 2024." Cover all filled fields.`
  },
  {
    id: 11,
    section: "Stakeholder Experience",
    searchQueries: (data) => [
      `patient experience design ${data.roomTypology || "hospital room"} biophilic healing environment lighting acoustic 2025`,
      `healthcare privacy family interaction smart room infotainment hygiene by design latest 2025`
    ],
    systemPrompt: `You are a healthcare design researcher specializing in evidence-based design and patient experience (HERD Journal, Planetree, Magnet, WELL Building Standard).
 
Analyse ALL filled fields across ALL subsections:
- Sensory Comfort: lightingQuality, lightingNotes, acousticControl, acousticNotes, thermalOdorControl, thermalOdorNotes — validate comfort standards
- Patient & Family Experience: privacy, patientComfort, patientComfortNotes, familyInteraction, familyInteractionNotes — validate against current patient-centred care standards
- Visual & Healing Environment: visualEnvironment, visualEnvironmentNotes, biophiliaHealingEnvironment, biophiliaHealingNotes — validate against HERD evidence and biophilic design research
- Technology & Hygiene Experience: technologyInfotainment, technologyInfotainmentNotes, infectionControlHygiene, infectionControlHygieneNotes — validate smart room and hygiene-by-design provisions
 
For every field, state: "In place of [current choice], consider [modern approach] — e.g. replace standard curtain screening with full visual + acoustic privacy (STC-50 solid walls + vision panel) for a consultation room to meet patient confidentiality standards." Cover all filled fields.`
  },
  {
    id: 12,
    section: "Fittings, Fixtures & Equipment",
    searchQueries: (data) => [
      `${data.roomTypology || "hospital room"} medical gas fittings clinical equipment ceiling supply unit pendant 2025`,
      `hospital nurse call IT hardware diagnostic equipment furniture latest technology standards 2025`
    ],
    systemPrompt: `You are a clinical equipment planner, biomedical engineer, and healthcare facilities specialist (HTM 02-01, IEC 60601, NABH).
 
Analyse ALL filled fields across ALL subsections:
- Medical Gas & Clinical Fittings: airFlowmeter, oxygenFlowmeter, oxygenFlowmeterLowFlow, suctionAdapterLowFlow, suctionBottle, blenderAirOxygen, ivHook, curtainTrackSystem — validate quantities and types for room type
- Furniture: trolleyProcedure, stoolAdjustableMobile, patientFurniture, staffVisitorFurniture, storageFurniture — validate clinical ergonomics and quantity
- Accessories & Dispensers: wallMountedDispensers, wasteBins, additionalFF — validate infection control requirements
- Clinical Equipment: infusionPumpSyringe, infusionPumpEnteral, infusionPumpSingleChannel, physiologicMonitor, ventilatorNeonatal, infantIncubator, phototherapyLamp, examinationLight, supplyUnitCeiling, medicalEquipments — validate for room type and suggest modern alternatives
- Diagnostics, IT & Communication: wallMountedDiagnostics, itCommunicationHardware, nurseCallSystems, additionalFE
 
For every item, validate whether the selection is current best-practice and state: "In place of [current equipment], consider [modern alternative] — e.g. replace standalone infusion pumps with networked smart pump systems (Alaris/BD/B.Braun) with dose-error reduction software (DERS) integration." Cover all filled fields.`
  },
  {
    id: 13,
    section: "Waste Management",
    searchQueries: (data) => [
      `hospital biomedical waste management BMW Rules 2016 amendment ${data.roomTypology || "clinical room"} India 2025`,
      `healthcare waste segregation colour coding disposal technology sustainable practices 2025`
    ],
    systemPrompt: `You are a biomedical waste management specialist (BMW Rules 2016 & 2019 Amendment, WHO Healthcare Waste Guidelines, CPCB India).
 
Analyse ALL filled fields across ALL subsections:
- Clinical & Hazardous Waste: wmBiohazard, wmRadioactive, wmFlammableSolvent, wmChemicalWaste, wmHumanAnatomical, wmMicrobiologyWaste, wmWasteSharps, wmCytotoxicDrugs — validate bin type, colour coding, segregation and disposal route
- General & Solid Waste: wmSoiledWaste, wmSolidWaste, wmLiquidWaste, wmDiscardedContainers, wmUsedOil, wmEwaste, wmConfidentialPaper, wmFoodPantryWaste — validate compliance with BMW Rules
- Additional Waste Streams: wmOthers1, wmOthers2, wmNotes — check for any non-standard streams
 
For every enabled waste stream, validate the complete management chain (segregation → storage → transport → disposal) and state: "In place of [current approach], consider [modern/compliant alternative] — e.g. replace standard 2-bin system with full 4-bin BMW-compliant station (Yellow/Red/Blue/Black) with foot-operated lids and CPCB-authorised collector tie-up." Cover all enabled waste streams.`
  }
];
 
// ── Tavily Web Search ─────────────────────────────────────────────────────────
async function tavilySearch(query) {
  const TAVILY_KEY = process.env.TAVILY_API_KEY;
  if (!TAVILY_KEY) return [];
  try {
    const res = await fetch("https://api.tavily.com/search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        api_key: TAVILY_KEY,
        query,
        search_depth: "basic",
        max_results: 3,
        include_answer: true
      })
    });
    const json = await res.json();
    return (json.results || []).map(r => ({
      title: r.title,
      url: r.url,
      snippet: r.content?.slice(0, 300)
    }));
  } catch (e) {
    console.warn(`Tavily search failed for "${query}":`, e.message);
    return [];
  }
}
 
// ── Run one section agent ─────────────────────────────────────────────────────
async function runSectionAgent(agentConfig, sectionData, roomContext, groq) {
  // 1. Web search — run both queries in parallel
  const searchResults = await Promise.all(
    agentConfig.searchQueries(roomContext).map(q => tavilySearch(q))
  );
  const allResults = searchResults.flat().slice(0, 5);
  const searchContext = allResults.length
    ? allResults.map(r => `• ${r.title}: ${r.snippet}`).join("\n")
    : "No web search results available.";
 
  // 2. Build prompt
  const prompt = `
ROOM CONTEXT:
- Room Name: ${roomContext.roomName || "Unknown"}
- Room Type: ${roomContext.roomTypology || "Unknown"}
- Department: ${roomContext.department || "Unknown"}
- Criticality: ${roomContext.criticalityLevel || "Unknown"}
 
SECTION: ${agentConfig.section}
SECTION DATA (what the user filled):
${JSON.stringify(sectionData, null, 2)}
 
LATEST INDUSTRY RESEARCH (from web):
${searchContext}
 
TASK:
1. Validate EVERY filled field — is it appropriate for this room type, criticality, and infection risk level?
2. Identify outdated specs, missing items, non-compliant selections, or compliance gaps.
3. For EVERY issue or improvement opportunity, explicitly suggest a modern replacement in this format:
   "In place of [current value], consider [modern alternative] — because [clinical/technical reason per current standards]."
4. Cover ALL fields that have been filled, not just one part of the section.
5. Give an overall confidence/quality score 0-100 based on completeness and standard compliance.
 
Respond ONLY with this exact JSON (no markdown, no extra text):
{
  "valid": true or false,
  "confidence": <number 0-100>,
  "summary": "<2 sentence overall assessment>",
  "issues": [
    { "field": "<fieldName>", "current": "<current value>", "problem": "<what's wrong or outdated>" }
  ],
  "suggestions": [
    { "field": "<fieldName or topic>", "recommendation": "<specific suggestion>", "reason": "<clinical/technical rationale>", "priority": "High|Medium|Low" }
  ],
  "sources": [
    { "title": "<source title>", "url": "<url>" }
  ]
}`;
 
  // 3. Call Groq
  const completion = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    max_tokens: 1200,
    temperature: 0.3,
    messages: [
      { role: "system", content: agentConfig.systemPrompt },
      { role: "user", content: prompt }
    ]
  });
 
  const raw = completion.choices[0]?.message?.content || "{}";
  const cleaned = raw.replace(/^```[a-z]*\n?/i, "").replace(/```$/, "").trim();
 
  try {
    const result = JSON.parse(cleaned);
    return {
      sectionId: agentConfig.id,
      section: agentConfig.section,
      ...result,
      sources: [...(result.sources || []), ...allResults.map(r => ({ title: r.title, url: r.url }))]
        .filter((s, i, arr) => s.url && arr.findIndex(x => x.url === s.url) === i)
        .slice(0, 4)
    };
  } catch {
    return {
      sectionId: agentConfig.id,
      section: agentConfig.section,
      valid: true,
      confidence: 50,
      summary: "Validation completed with limited analysis.",
      issues: [],
      suggestions: [],
      sources: allResults.map(r => ({ title: r.title, url: r.url }))
    };
  }
}
 
// ── Extract section data from full room data ──────────────────────────────────
function extractSectionData(fullData, sectionKeys) {
  const result = {};
  sectionKeys.forEach(k => {
    if (fullData[k] !== undefined && fullData[k] !== null && fullData[k] !== "") {
      result[k] = fullData[k];
    }
  });
  return result;
}
 
// ── POST /validate-rds ────────────────────────────────────────────────────────
app.post("/validate-rds", async (req, res) => {
  try {
    const { roomId } = req.body;
    if (!roomId) return res.status(400).json({ error: "roomId is required" });
 
    const GROQ_KEY = process.env.GROQ_API_KEY;
    if (!GROQ_KEY) return res.status(500).json({ error: "GROQ_API_KEY not set" });
 
    // Fetch room from Supabase
    const { data: room, error: fetchError } = await supabase
      .from("rds_rooms")
      .select("*")
      .eq("id", String(roomId))
      .single();
 
    if (fetchError || !room) {
      return res.status(404).json({ error: "Room not found" });
    }
 
    const fullData = typeof room.data === "string" ? safeJson(room.data) : (room.data || {});
    const groq = new Groq({ apiKey: GROQ_KEY });
 
    const roomContext = {
      roomName:        fullData.roomName        || room.roomname     || "",
      roomTypology:    fullData.roomTypology     || "",
      department:      fullData.department       || room.department   || "",
      criticalityLevel: fullData.criticalityLevel || "",
      roomCode:        fullData.roomCode         || room.roomcode     || ""
    };
 
    console.log(`[Validation] Starting 13-agent validation for room ${roomContext.roomCode}...`);
 
    // Run all 13 agents in parallel
    const agentPromises = AGENT_CONFIGS.map((agentConfig, idx) => {
      const sectionDef  = SECTIONS[idx];
      const sectionData = sectionDef
        ? extractSectionData(fullData, sectionDef.keys)
        : {};
      return runSectionAgent(agentConfig, sectionData, roomContext, groq)
        .catch(err => ({
          sectionId: agentConfig.id,
          section:   agentConfig.section,
          valid:     true,
          confidence: 40,
          summary:   `Agent encountered an error: ${err.message}`,
          issues:    [],
          suggestions: [],
          sources:   []
        }));
    });
 
    const sectionResults = await Promise.all(agentPromises);
    console.log(`[Validation] All 13 agents completed for ${roomContext.roomCode}`);
 
    // Consolidate report
    const totalConfidence   = Math.round(sectionResults.reduce((s, r) => s + (r.confidence || 0), 0) / 13);
    const totalIssues       = sectionResults.reduce((s, r) => s + (r.issues?.length || 0), 0);
    const totalSuggestions  = sectionResults.reduce((s, r) => s + (r.suggestions?.length || 0), 0);
    const highPriority      = sectionResults.flatMap(r => r.suggestions || []).filter(s => s.priority === "High").length;
 
    const report = {
      roomId:       String(roomId),
      roomCode:     roomContext.roomCode,
      roomName:     roomContext.roomName,
      roomTypology: roomContext.roomTypology,
      department:   roomContext.department,
      validatedAt:  new Date().toISOString(),
      overallScore: totalConfidence,
      overallStatus: totalConfidence >= 80 ? "Excellent" : totalConfidence >= 60 ? "Good" : totalConfidence >= 40 ? "Needs Review" : "Critical",
      summary: {
        totalIssues,
        totalSuggestions,
        highPriorityCount: highPriority,
        sectionsValidated: 13
      },
      sections: sectionResults
    };
 
    // Save to Supabase rds_validations — upsert so re-runs overwrite old report
    const { error: saveErr } = await supabase
      .from("rds_validations")
      .upsert(
        {
          room_id:    String(roomId),
          room_code:  roomContext.roomCode,
          report:     report,
          created_at: new Date().toISOString()
        },
        { onConflict: "room_id" }   // update existing row if room_id already exists
      );
 
    if (saveErr) {
      // Log the full error so you can see it in Render logs
      console.error("[Validation] Supabase save FAILED:", JSON.stringify(saveErr));
    } else {
      console.log("[Validation] ✓ Report saved to rds_validations for room_id:", String(roomId));
    }
 
    console.log(`[Validation] ✓ Score=${totalConfidence} Issues=${totalIssues} Suggestions=${totalSuggestions}`);
    res.json(report);
 
  } catch (e) {
    console.error("Validation error:", e);
    res.status(500).json({ error: "Validation failed: " + e.message });
  }
});
 
// GET /validate-rds/:roomId — fetch saved validation report
app.get("/validate-rds/:roomId", async (req, res) => {
  try {
    const { data, error } = await supabase
      .from("rds_validations")
      .select("*")
      .eq("room_id", req.params.roomId)
      .order("created_at", { ascending: false })
      .limit(1);          // NO .single() — avoids throwing when empty
 
    if (error) {
      console.error("[Validation] GET error:", JSON.stringify(error));
      return res.status(500).json({ error: "Database error: " + error.message });
    }
    if (!data || data.length === 0) {
      return res.status(404).json({ error: "No validation found" });
    }
    res.json(data[0].report);
  } catch (e) {
    console.error("[Validation] GET exception:", e.message);
    res.status(500).json({ error: "Failed to fetch validation" });
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