import { useState, useEffect } from "react";

// ─────────────────────────────────────────────────────────────────────────────
// GOOGLE SHEETS SETUP INSTRUCTIONS
// ─────────────────────────────────────────────────────────────────────────────
//
// STEP 1 — Create the Google Sheet
//   • Make a new Google Sheet
//   • Create two tabs named exactly: "Shipments" and "Acknowledgements"
//
// STEP 2 — Choose your secret key
//   • Pick any private passphrase, e.g. "GAP-AOC-2025-X7K9"
//   • Set it in SHEETS_SECRET below (replace the placeholder)
//   • You will paste the SAME string into the Apps Script in Step 3
//
// STEP 3 — Add the Apps Script
//   In your Sheet: Extensions → Apps Script → delete any existing code → paste:
//
//    var SECRET_KEY = "PASTE_YOUR_SECRET_HERE"; // ← must match SHEETS_SECRET below
//
//    function doPost(e) {
//      try {
//        var data = JSON.parse(e.postData.contents);
//
//        // ── Secret key check ──────────────────────────────────────────
//        if (!data.secret || data.secret !== SECRET_KEY) {
//          return ContentService
//            .createTextOutput(JSON.stringify({ status: "unauthorized" }))
//            .setMimeType(ContentService.MimeType.JSON);
//        }
//
//        var ss = SpreadsheetApp.getActiveSpreadsheet();
//
//        if (data.type === "SHIPMENT_CREATED") {
//          var sheet = ss.getSheetByName("Shipments");
//          if (sheet.getLastRow() === 0) {
//            sheet.appendRow([
//              "Shipment ID","Created At","Sender Name","Sender Company",
//              "Sender DEA","Sender NPI","Sender Address","Carrier","Tracking #",
//              "Ship Date","HCP Name","HCP Credentials","HCP Practice","HCP NPI",
//              "HCP DEA","HCP Address","Samples Summary","Status","Signature","AOC Timestamp"
//            ]);
//          }
//          sheet.appendRow([
//            data.id, data.createdAt, data.senderName, data.senderCompany,
//            data.senderDEA, data.senderNPI, data.senderAddress,
//            data.carrier, data.trackingNumber, data.shipDate,
//            data.hcpName, data.hcpCredentials, data.hcpPractice,
//            data.hcpNPI, data.hcpDEA, data.hcpAddress,
//            data.samplesSummary, "PENDING", data.signature, ""
//          ]);
//        }
//
//        if (data.type === "AOC_RECEIVED") {
//          var aSheet = ss.getSheetByName("Acknowledgements");
//          if (aSheet.getLastRow() === 0) {
//            aSheet.appendRow([
//              "Shipment ID","Acknowledged By","Title",
//              "Timestamp","User Agent","Audit Trail","Signature"
//            ]);
//          }
//          aSheet.appendRow([
//            data.shipmentId, data.acknowledgedBy, data.acknowledgedTitle,
//            data.timestamp, data.userAgent, data.auditTrail, data.signature
//          ]);
//          // Update status + AOC timestamp in Shipments tab
//          var sSheet = ss.getSheetByName("Shipments");
//          var vals = sSheet.getDataRange().getValues();
//          for (var i = 1; i < vals.length; i++) {
//            if (vals[i][0] === data.shipmentId) {
//              sSheet.getRange(i + 1, 18).setValue("AOC RECEIVED");
//              sSheet.getRange(i + 1, 20).setValue(data.timestamp);
//              break;
//            }
//          }
//        }
//
//        return ContentService
//          .createTextOutput(JSON.stringify({ status: "ok" }))
//          .setMimeType(ContentService.MimeType.JSON);
//
//      } catch(err) {
//        return ContentService
//          .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
//          .setMimeType(ContentService.MimeType.JSON);
//      }
//    }
//
// STEP 4 — Deploy as Web App
//   Click Deploy → New Deployment → Web App
//     Execute as:     Me
//     Who has access: Anyone
//   → Copy the deployment URL
//
//   ⚠ Every time you edit the Apps Script you must create a NEW deployment
//     (Deploy → Manage Deployments → New Version) or changes won't take effect.
//
// STEP 5 — Fill in both values below:
// ─────────────────────────────────────────────────────────────────────────────

const SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbz2GyqSiR4uHIFnz-CuQOqu-qkbacFMFQ-QKQ6KqfcCXLo02yr-N_YBLxy4n4UXEGcdpw/exec";
const SHEETS_SECRET      = "GAPbridge"; // must match SECRET_KEY in Apps Script

async function postToSheets(payload) {
  if (
    !SHEETS_WEBHOOK_URL || SHEETS_WEBHOOK_URL.includes("YOUR_APPS_SCRIPT") ||
    !SHEETS_SECRET      || SHEETS_SECRET.includes("YOUR_SECRET")
  ) return;
  try {
    await fetch(SHEETS_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      // Secret travels in the body — never visible to end users
      body: JSON.stringify({ ...payload, secret: SHEETS_SECRET }),
      mode: "no-cors", // required for Google Apps Script from browser
    });
  } catch (e) {
    console.warn("Google Sheets sync failed (non-blocking):", e);
  }
}

// ─── GAP Regulatory Consulting Color Palette ─────────────────────────────
const C = {
  bg:            "#020e0e",
  bgCard:        "#071414",
  bgDeep:        "#010a0a",
  bgInput:       "#051111",
  border:        "#0d2e2e",
  borderLight:   "#134040",
  borderFocus:   "#007a7a",
  accent:        "#009494",
  accentMid:     "#006a6a",
  accentDim:     "#004949",
  accentText:    "#00c8c8",
  accentFaint:   "#00494920",
  textPrimary:   "#c8e8e8",
  textMid:       "#6aadad",
  textDim:       "#3a6a6a",
  textFaint:     "#1a3a3a",
  gold:          "#c8a840",
  goldDim:       "#7a6020",
  goldBg:        "#1a1404",
  red:           "#c85040",
  redBg:         "#1a0a08",
  greenOk:       "#00b890",
  greenOkBg:     "#041a14",
  greenOkBorder: "#0a3028",
};

const CFR_VERSION = "21CFR11-v2";

// ─── Helpers ──────────────────────────────────────────────────────────────
function todayISO() {
  return new Date().toISOString().split("T")[0];
}

function generateShipmentId() {
  const ts = Date.now().toString(36).toUpperCase();
  const rand = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `RX-${ts}-${rand}`;
}

async function signPayload(payload) {
  const encoder = new TextEncoder();
  const keyData = encoder.encode("CFR11-PHARMA-AOC-SECRET-KEY-2024");
  const cryptoKey = await window.crypto.subtle.importKey(
    "raw", keyData, { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  const data = encoder.encode(JSON.stringify(payload));
  const sig = await window.crypto.subtle.sign("HMAC", cryptoKey, data);
  return Array.from(new Uint8Array(sig)).map(b => b.toString(16).padStart(2, "0")).join("");
}

function getAuditTimestamp() { return new Date().toISOString(); }
function getUserAgent() { return navigator.userAgent.substring(0, 80); }

function summarizeSamples(samples) {
  return (samples || [])
    .map(s => `${s.productName} (NDC:${s.ndc} Lot:${s.lot} Qty:${s.qty})`)
    .join(" | ");
}

// ─── Storage ──────────────────────────────────────────────────────────────
async function saveShipment(shipment) {
  try {
    await window.storage.set(`shipment:${shipment.id}`, JSON.stringify(shipment), true);
    let index = [];
    try {
      const idx = await window.storage.get("shipment-index", true);
      if (idx) index = JSON.parse(idx.value);
    } catch {}
    if (!index.includes(shipment.id)) index.push(shipment.id);
    await window.storage.set("shipment-index", JSON.stringify(index), true);
  } catch (e) { console.error("Storage error", e); }
}

async function loadShipment(id) {
  try {
    const r = await window.storage.get(`shipment:${id}`, true);
    return r ? JSON.parse(r.value) : null;
  } catch { return null; }
}

async function loadAllShipments() {
  try {
    const idx = await window.storage.get("shipment-index", true);
    if (!idx) return [];
    const ids = JSON.parse(idx.value);
    const all = await Promise.all(ids.map(id => loadShipment(id)));
    return all.filter(Boolean).reverse();
  } catch { return []; }
}

async function saveSenderProfile(profile) {
  try { await window.storage.set("sender-profile", JSON.stringify(profile), false); } catch {}
}

async function loadSenderProfile() {
  try {
    const r = await window.storage.get("sender-profile", false);
    return r ? JSON.parse(r.value) : null;
  } catch { return null; }
}

async function updateShipmentAOC(id, aocRecord) {
  const shipment = await loadShipment(id);
  if (!shipment) return null;
  shipment.aoc = aocRecord;
  shipment.status = "acknowledged";
  await saveShipment(shipment);
  return shipment;
}

// ─── QR Code (pure SVG) ───────────────────────────────────────────────────
function generateQRMatrix(text) {
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    hash = ((hash << 5) - hash) + text.charCodeAt(i);
    hash |= 0;
  }
  const size = 25;
  const matrix = [];
  for (let r = 0; r < size; r++) {
    matrix[r] = [];
    for (let c = 0; c < size; c++) {
      const inFinder =
        (r < 7 && c < 7) || (r < 7 && c >= size - 7) || (r >= size - 7 && c < 7);
      if (inFinder) {
        const fr = r < 7 ? r : r - (size - 7);
        const fc = c < 7 ? c : c - (size - 7);
        matrix[r][c] = (
          fr === 0 || fr === 6 || fc === 0 || fc === 6 ||
          (fr >= 2 && fr <= 4 && fc >= 2 && fc <= 4)
        ) ? 1 : 0;
      } else {
        const seed = (hash ^ (r * 31 + c * 17) ^ (r * c)) >>> 0;
        matrix[r][c] = (seed % 3 === 0 || seed % 7 === 1) ? 1 : 0;
      }
    }
  }
  return matrix;
}

function QRCode({ value, size = 160 }) {
  const matrix = generateQRMatrix(value);
  const cellSize = size / matrix.length;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ display: "block" }}>
      <rect width={size} height={size} fill="white" />
      {matrix.map((row, r) =>
        row.map((cell, c) =>
          cell ? (
            <rect key={`${r}-${c}`} x={c * cellSize} y={r * cellSize}
              width={cellSize} height={cellSize} fill="#004949" />
          ) : null
        )
      )}
    </svg>
  );
}

// ─── Sheets Status Badge ──────────────────────────────────────────────────
function SheetsBadge() {
  const urlOk       = SHEETS_WEBHOOK_URL && !SHEETS_WEBHOOK_URL.includes("YOUR_APPS_SCRIPT");
  const secretOk    = SHEETS_SECRET      && !SHEETS_SECRET.includes("YOUR_SECRET");
  const isConfigured = urlOk && secretOk;

  if (!isConfigured) {
    const missing = !urlOk && !secretOk ? "URL + secret key"
                  : !urlOk             ? "webhook URL"
                  :                      "secret key";
    return (
      <div style={{
        background: "#1a1404", border: "1px solid #7a6020",
        borderRadius: 4, padding: "6px 12px", fontSize: 10,
        color: "#c8a840", letterSpacing: 1, display: "inline-flex", alignItems: "center", gap: 6
      }}>
        ⚠ Sheets not configured ({missing} missing)
      </div>
    );
  }
  return (
    <div style={{
      background: C.greenOkBg, border: `1px solid ${C.greenOkBorder}`,
      borderRadius: 4, padding: "6px 12px", fontSize: 10,
      color: C.greenOk, letterSpacing: 1, display: "inline-flex", alignItems: "center", gap: 6
    }}>
      🔒 Google Sheets · secured
    </div>
  );
}

// ─── Packing Slip ─────────────────────────────────────────────────────────
function PackingSlip({ shipment, qrUrl }) {
  return (
    <div style={{
      background: "white", color: "#111", padding: 36, width: 620,
      fontFamily: "Georgia, serif", fontSize: 13,
      boxShadow: "0 8px 40px #00494930", border: "1px solid #aacccc"
    }}>
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "flex-start",
        marginBottom: 24, borderBottom: "2px solid #004949", paddingBottom: 20
      }}>
        <div>
          <div style={{ fontSize: 9, letterSpacing: 4, textTransform: "uppercase", color: "#006060", marginBottom: 4 }}>
            GAP Regulatory Consulting · Pharmaceutical Sample Transfer
          </div>
          <div style={{ fontSize: 20, fontWeight: "bold", color: "#004949" }}>
            {shipment.senderCompany || shipment.senderName}
          </div>
          <div style={{ fontSize: 11, color: "#555", marginTop: 4 }}>Shipment ID: {shipment.id}</div>
          <div style={{ fontSize: 11, color: "#555" }}>Date: {shipment.shipDate || new Date(shipment.createdAt).toLocaleDateString()}</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <QRCode value={qrUrl} size={110} />
          <div style={{ fontSize: 9, color: "#006060", marginTop: 4, letterSpacing: 1 }}>SCAN TO ACKNOWLEDGE</div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, marginBottom: 24 }}>
        <div>
          <div style={{ fontSize: 9, letterSpacing: 2, textTransform: "uppercase", color: "#006060", marginBottom: 6 }}>From</div>
          <div style={{ fontWeight: "bold" }}>{shipment.senderName}</div>
          <div>{shipment.senderCompany}</div>
          <div style={{ fontSize: 12, color: "#555", marginTop: 4, whiteSpace: "pre-wrap" }}>{shipment.senderAddress}</div>
          <div style={{ fontSize: 11, marginTop: 8 }}>DEA: {shipment.senderDEA || "N/A"}</div>
          <div style={{ fontSize: 11 }}>NPI: {shipment.senderNPI || "N/A"}</div>
          {shipment.carrier && <div style={{ fontSize: 11, marginTop: 6 }}>Carrier: {shipment.carrier} · {shipment.trackingNumber}</div>}
        </div>
        <div>
          <div style={{ fontSize: 9, letterSpacing: 2, textTransform: "uppercase", color: "#006060", marginBottom: 6 }}>To (Receiving HCP)</div>
          <div style={{ fontWeight: "bold" }}>
            {shipment.hcpName}{shipment.hcpCredentials ? `, ${shipment.hcpCredentials}` : ""}
          </div>
          <div>{shipment.hcpPractice}</div>
          <div style={{ fontSize: 12, color: "#555", marginTop: 4, whiteSpace: "pre-wrap" }}>{shipment.hcpAddress}</div>
          <div style={{ fontSize: 11, marginTop: 8 }}>NPI: {shipment.hcpNPI || "N/A"}</div>
          <div style={{ fontSize: 11 }}>DEA: {shipment.hcpDEA || "N/A"}</div>
        </div>
      </div>

      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, marginBottom: 20 }}>
        <thead>
          <tr style={{ borderBottom: "1px solid #aacccc", background: "#eef6f6" }}>
            {["Product Name","NDC","Lot #","Qty","Strength","Form"].map(h => (
              <th key={h} style={{ textAlign: "left", padding: "6px 8px", fontSize: 10, color: "#004949", letterSpacing: 1 }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {(shipment.samples || []).map((s, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #ddeaea" }}>
              <td style={{ padding: "6px 8px", fontWeight: "bold" }}>{s.productName}</td>
              <td style={{ padding: "6px 8px" }}>{s.ndc}</td>
              <td style={{ padding: "6px 8px" }}>{s.lot}</td>
              <td style={{ padding: "6px 8px" }}>{s.qty}</td>
              <td style={{ padding: "6px 8px" }}>{s.strength}</td>
              <td style={{ padding: "6px 8px" }}>{s.form}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={{ border: "1px solid #aacccc", padding: 14, borderRadius: 4, background: "#f2fafa", marginBottom: 16 }}>
        <div style={{ fontSize: 9, letterSpacing: 2, textTransform: "uppercase", color: "#006060", marginBottom: 6 }}>
          Acknowledgement of Contents (AOC)
        </div>
        <div style={{ fontSize: 11, lineHeight: 1.6, color: "#444" }}>
          By scanning the QR code above, the receiving HCP electronically acknowledges receipt and inspection of the pharmaceutical samples listed herein, in compliance with 21 CFR Part 11, PDMA, and all applicable federal and state regulations.
        </div>
      </div>

      <div style={{ fontSize: 10, color: "#7aadad", textAlign: "center", borderTop: "1px solid #ddeaea", paddingTop: 10 }}>
        Sig: {shipment.signature?.substring(0, 32)}… · 21 CFR Part 11 · PDMA · {CFR_VERSION} · GAP Regulatory Consulting
      </div>
    </div>
  );
}

// ─── Acknowledgement View (HCP scans QR) ─────────────────────────────────
function AcknowledgementView({ shipmentId }) {
  const [shipment, setShipment] = useState(null);
  const [step, setStep] = useState("loading");
  const [sigName, setSigName] = useState("");
  const [sigTitle, setSigTitle] = useState("");
  const [agreed, setAgreed] = useState(false);
  const [aocRecord, setAocRecord] = useState(null);
  const [syncing, setSyncing] = useState(false);

  useEffect(() => {
    loadShipment(shipmentId).then(s => {
      if (!s) { setStep("error"); return; }
      if (s.status === "acknowledged") { setAocRecord(s.aoc); setStep("complete"); return; }
      setShipment(s); setStep("review");
    });
  }, [shipmentId]);

  const handleSign = async () => {
    if (!sigName || !agreed) return;
    setSyncing(true);
    const record = {
      acknowledgedBy: sigName,
      acknowledgedTitle: sigTitle,
      timestamp: getAuditTimestamp(),
      userAgent: getUserAgent(),
      shipmentId,
      auditTrail: JSON.stringify([
        { event: "QR_SCANNED",         ts: getAuditTimestamp() },
        { event: "CONTENTS_REVIEWED",  ts: getAuditTimestamp() },
        { event: "ESIGNATURE_APPLIED", ts: getAuditTimestamp(), signer: sigName }
      ]),
      cfrCompliance: CFR_VERSION
    };
    record.signature = await signPayload(record);

    // Save locally
    await updateShipmentAOC(shipmentId, record);

    // Push AOC to Google Sheets
    await postToSheets({
      type: "AOC_RECEIVED",
      ...record,
    });

    setAocRecord(record);
    setSyncing(false);
    setStep("complete");
  };

  const inp = {
    width: "100%", boxSizing: "border-box",
    background: C.bgInput, border: `1px solid ${C.border}`,
    borderRadius: 4, color: C.textPrimary, padding: "10px 14px",
    fontSize: 13, fontFamily: "Georgia, serif", outline: "none"
  };

  if (step === "loading") return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ color: C.accentText, fontFamily: "Georgia, serif", letterSpacing: 4, fontSize: 12 }}>LOADING…</div>
    </div>
  );

  if (step === "error") return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ color: C.red, fontFamily: "Georgia, serif", textAlign: "center" }}>
        <div style={{ fontSize: 40, marginBottom: 12 }}>⚠</div>
        Shipment record not found or invalid QR code.
      </div>
    </div>
  );

  if (step === "complete") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Georgia, serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{
        background: C.greenOkBg, border: `1px solid ${C.greenOkBorder}`,
        borderRadius: 8, padding: 40, maxWidth: 520, width: "100%", textAlign: "center"
      }}>
        <div style={{ fontSize: 52, color: C.greenOk, marginBottom: 12 }}>✓</div>
        <div style={{ fontSize: 10, letterSpacing: 4, color: C.greenOk, textTransform: "uppercase", marginBottom: 8 }}>AOC Complete</div>
        <div style={{ fontSize: 22, color: C.textPrimary, marginBottom: 24 }}>Receipt Confirmed</div>
        <div style={{ fontSize: 12, color: C.textMid, lineHeight: 1.9, textAlign: "left" }}>
          <div><b style={{ color: C.accentText }}>Acknowledged by:</b> {aocRecord?.acknowledgedBy}{aocRecord?.acknowledgedTitle ? `, ${aocRecord.acknowledgedTitle}` : ""}</div>
          <div><b style={{ color: C.accentText }}>Timestamp (UTC):</b> {aocRecord?.timestamp}</div>
          <div><b style={{ color: C.accentText }}>Shipment ID:</b> {shipmentId}</div>
          <div style={{ marginTop: 12, fontSize: 10, color: C.textFaint, wordBreak: "break-all" }}>
            Sig: {aocRecord?.signature}
          </div>
          <div style={{ marginTop: 8, fontSize: 10, color: C.textFaint }}>
            21 CFR Part 11 Compliant · {CFR_VERSION} · Recorded in Google Sheets
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Georgia, serif", padding: "40px 24px" }}>
      <div style={{ maxWidth: 600, margin: "0 auto" }}>
        <div style={{ fontSize: 9, letterSpacing: 5, color: C.accentText, textTransform: "uppercase", marginBottom: 6 }}>
          GAP Regulatory Consulting
        </div>
        <h1 style={{ color: C.textPrimary, fontWeight: "normal", fontSize: 24, margin: "0 0 28px" }}>
          Pharmaceutical Sample Receipt
        </h1>

        <div style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 14 }}>Shipment Details</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px 24px" }}>
            {[
              ["Shipment ID", shipment?.id],
              ["Ship Date", shipment?.shipDate || new Date(shipment?.createdAt).toLocaleDateString()],
              ["From", `${shipment?.senderName}${shipment?.senderCompany ? ` · ${shipment.senderCompany}` : ""}`],
              ["Carrier", `${shipment?.carrier || "—"} ${shipment?.trackingNumber || ""}`],
            ].map(([k, v]) => (
              <div key={k}>
                <div style={{ fontSize: 10, color: C.textDim, letterSpacing: 2, textTransform: "uppercase", marginBottom: 2 }}>{k}</div>
                <div style={{ color: C.textPrimary, fontSize: 13 }}>{v || "—"}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 14 }}>Sample Contents</div>
          {(shipment?.samples || []).map((s, i) => (
            <div key={i} style={{
              borderBottom: i < (shipment?.samples?.length - 1) ? `1px solid ${C.border}` : "none",
              padding: "10px 0", display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", gap: 8
            }}>
              <div style={{ fontSize: 13, color: C.textPrimary, fontWeight: "bold" }}>{s.productName}</div>
              <div style={{ fontSize: 11, color: C.textMid }}>NDC: {s.ndc}</div>
              <div style={{ fontSize: 11, color: C.textMid }}>Lot: {s.lot}</div>
              <div style={{ fontSize: 11, color: C.textMid }}>Qty: {s.qty}</div>
            </div>
          ))}
        </div>

        <div style={{ background: C.bgCard, border: `1px solid ${C.greenOkBorder}`, borderRadius: 6, padding: 24 }}>
          <div style={{ fontSize: 10, letterSpacing: 3, color: C.greenOk, textTransform: "uppercase", marginBottom: 12 }}>
            Electronic Signature · 21 CFR Part 11
          </div>
          <div style={{ fontSize: 11, color: C.textDim, lineHeight: 1.7, marginBottom: 20 }}>
            By signing below, I confirm personal receipt and inspection of the pharmaceutical samples listed above, and acknowledge that this constitutes a legally binding electronic signature under 21 CFR Part 11 and PDMA.
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
            <div>
              <label style={{ display: "block", fontSize: 10, letterSpacing: 2, color: C.textDim, textTransform: "uppercase", marginBottom: 6 }}>Full Name *</label>
              <input value={sigName} onChange={e => setSigName(e.target.value)} placeholder="" style={inp}
                onFocus={e => e.target.style.borderColor = C.borderFocus}
                onBlur={e => e.target.style.borderColor = C.border} />
            </div>
            <div>
              <label style={{ display: "block", fontSize: 10, letterSpacing: 2, color: C.textDim, textTransform: "uppercase", marginBottom: 6 }}>Title / Credentials</label>
              <input value={sigTitle} onChange={e => setSigTitle(e.target.value)} placeholder="" style={inp}
                onFocus={e => e.target.style.borderColor = C.borderFocus}
                onBlur={e => e.target.style.borderColor = C.border} />
            </div>
          </div>
          <label style={{ display: "flex", alignItems: "flex-start", gap: 12, cursor: "pointer", marginBottom: 20 }}>
            <input type="checkbox" checked={agreed} onChange={e => setAgreed(e.target.checked)} style={{ marginTop: 2, accentColor: C.accent }} />
            <span style={{ fontSize: 11, color: C.textDim, lineHeight: 1.6 }}>
              I acknowledge receipt and physical inspection of all items listed above, confirm their condition upon delivery, and accept responsibility as the receiving HCP under applicable regulations.
            </span>
          </label>
          <button
            onClick={handleSign}
            disabled={!sigName || !agreed || syncing}
            style={{
              width: "100%",
              background: (!sigName || !agreed || syncing) ? C.accentFaint : C.accentDim,
              border: `1px solid ${(!sigName || !agreed || syncing) ? C.border : C.accent}`,
              borderRadius: 4,
              color: (!sigName || !agreed || syncing) ? C.textDim : C.accentText,
              padding: "13px 0", fontSize: 12, letterSpacing: 3,
              textTransform: "uppercase", cursor: (!sigName || !agreed || syncing) ? "not-allowed" : "pointer",
              fontFamily: "Georgia, serif", transition: "all 0.2s"
            }}
          >
            {syncing ? "Recording…" : "Submit Acknowledgement"}
          </button>
        </div>

        <div style={{ fontSize: 10, color: C.textFaint, textAlign: "center", lineHeight: 1.6, marginTop: 24 }}>
          21 CFR Part 11 · PDMA · Electronic Records & Signatures<br />GAP Regulatory Consulting · {CFR_VERSION}
        </div>
      </div>
    </div>
  );
}

// ─── Shared style helpers ─────────────────────────────────────────────────
const inputStyle = {
  width: "100%", boxSizing: "border-box",
  background: C.bgInput, border: `1px solid ${C.border}`,
  borderRadius: 4, color: C.textPrimary,
  padding: "10px 14px", fontSize: 13,
  fontFamily: "Georgia, serif", outline: "none"
};
// Textarea variant — auto-grows, shows full address
const textareaStyle = {
  ...inputStyle,
  resize: "vertical",
  minHeight: 72,
  lineHeight: 1.6,
};
const labelStyle = {
  display: "block", fontSize: 10, letterSpacing: 2,
  color: C.textDim, textTransform: "uppercase", marginBottom: 6
};

const EMPTY_SAMPLE = { productName: "", ndc: "", lot: "", qty: "", strength: "", form: "" };
const EMPTY_SENDER = { senderName: "", senderCompany: "", senderAddress: "", senderDEA: "", senderNPI: "" };

// ─── Access Passcode ──────────────────────────────────────────────────────
// Change this to any private passcode and share only with your 3PL / internal staff.
// HCPs scanning a QR code bypass this entirely — they never see the login screen.
const APP_PASSCODE = "GAPbridge2025";

// ─── Login Screen ─────────────────────────────────────────────────────────
function LoginScreen({ onSuccess }) {
  const [entry, setEntry]     = useState("");
  const [error, setError]     = useState(false);
  const [shake, setShake]     = useState(false);

  const attempt = () => {
    if (entry === APP_PASSCODE) {
      // Persist auth in sessionStorage so a page refresh doesn't log them out
      sessionStorage.setItem("gap-aoc-auth", "1");
      onSuccess();
    } else {
      setError(true);
      setShake(true);
      setEntry("");
      setTimeout(() => setShake(false), 600);
    }
  };

  const handleKey = (e) => { if (e.key === "Enter") attempt(); };

  return (
    <div style={{
      minHeight: "100vh", background: C.bg,
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "Georgia, serif"
    }}>
      <style>{`
        @keyframes shake {
          0%,100% { transform: translateX(0); }
          20%      { transform: translateX(-8px); }
          40%      { transform: translateX(8px); }
          60%      { transform: translateX(-6px); }
          80%      { transform: translateX(6px); }
        }
        .shake { animation: shake 0.5s ease; }
      `}</style>

      <div style={{ width: 360, textAlign: "center" }}>
        {/* Logo / wordmark */}
        <div style={{ marginBottom: 40 }}>
          <div style={{ fontSize: 9, letterSpacing: 6, color: C.accentText, textTransform: "uppercase", marginBottom: 10 }}>
            GAP Regulatory Consulting
          </div>
          <div style={{ fontSize: 22, color: C.textPrimary, letterSpacing: 2 }}>
            Pharma Sample <span style={{ color: C.accentText }}>AOC</span> System
          </div>
          <div style={{
            width: 40, height: 2, background: C.accentDim,
            margin: "16px auto 0"
          }} />
        </div>

        {/* Login card */}
        <div className={shake ? "shake" : ""} style={{
          background: C.bgCard, border: `1px solid ${error ? C.red : C.border}`,
          borderRadius: 8, padding: "32px 28px",
          transition: "border-color 0.3s"
        }}>
          <div style={{ fontSize: 10, letterSpacing: 3, color: C.textDim, textTransform: "uppercase", marginBottom: 20 }}>
            Authorized Access Only
          </div>

          <input
            type="password"
            value={entry}
            onChange={e => { setEntry(e.target.value); setError(false); }}
            onKeyDown={handleKey}
            autoFocus
            style={{
              ...inputStyle,
              textAlign: "center", letterSpacing: 4, fontSize: 16,
              border: `1px solid ${error ? C.red : C.border}`,
              marginBottom: 6
            }}
            onFocus={e => e.target.style.borderColor = error ? C.red : C.borderFocus}
            onBlur={e => e.target.style.borderColor = error ? C.red : C.border}
          />

          {error && (
            <div style={{ fontSize: 11, color: C.red, letterSpacing: 1, marginBottom: 14 }}>
              Incorrect passcode
            </div>
          )}
          {!error && <div style={{ marginBottom: 14 }} />}

          <button
            onClick={attempt}
            style={{
              width: "100%", background: C.accentDim, border: "none",
              borderRadius: 4, color: C.accentText, padding: "12px 0",
              fontSize: 11, letterSpacing: 3, textTransform: "uppercase",
              cursor: "pointer", fontFamily: "Georgia, serif"
            }}
          >
            Enter
          </button>
        </div>

        <div style={{ fontSize: 10, color: C.textFaint, marginTop: 24, letterSpacing: 1 }}>
          21 CFR Part 11 · PDMA · GAP Regulatory Consulting
        </div>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────
export default function App() {
  // ── Auth — bypass entirely for HCP QR scan links (?ack=...) ──────────────
  const isAckView = typeof window !== "undefined"
    && new URLSearchParams(window.location.search).get("ack") !== null;
  const [authed, setAuthed] = useState(
    isAckView || (typeof sessionStorage !== "undefined" && sessionStorage.getItem("gap-aoc-auth") === "1")
  );

  if (!authed) return <LoginScreen onSuccess={() => setAuthed(true)} />;
  const [view, setView]                 = useState("dashboard");
  const [shipments, setShipments]       = useState([]);
  const [activeShipment, setActive]     = useState(null);
  const [ackId, setAckId]               = useState(null);
  const [showPrint, setShowPrint]       = useState(false);
  const [loadingCreate, setLoading]     = useState(false);
  const [savedSender, setSavedSender]   = useState(null);
  const [saveSenderMsg, setSaveMsg]     = useState("");
  const [sheetsMsg, setSheetsMsg]       = useState("");

  const [form, setForm] = useState({
    ...EMPTY_SENDER,
    hcpName: "", hcpCredentials: "", hcpPractice: "",
    hcpAddress: "", hcpNPI: "", hcpDEA: "",
    trackingNumber: "", carrier: "",
    shipDate: todayISO(),          // ← auto-fill today
    samples: [{ ...EMPTY_SAMPLE }]
  });

  const appUrl = typeof window !== "undefined"
    ? window.location.href.split("?")[0]
    : "https://claude.ai";

  useEffect(() => {
    if (typeof window !== "undefined") {
      const id = new URLSearchParams(window.location.search).get("ack");
      if (id) { setAckId(id); setView("acknowledge"); }
    }
  }, []);

  useEffect(() => {
    loadSenderProfile().then(profile => {
      if (profile) {
        setSavedSender(profile);
        setForm(prev => ({ ...prev, ...profile }));
      }
    });
  }, []);

  useEffect(() => {
    if (view === "dashboard") loadAllShipments().then(setShipments);
  }, [view]);

  const updateForm    = (key, val) => setForm(prev => ({ ...prev, [key]: val }));
  const updateSample  = (i, key, val) => {
    const samples = [...form.samples];
    samples[i] = { ...samples[i], [key]: val };
    setForm(prev => ({ ...prev, samples }));
  };
  const addSample    = () => setForm(prev => ({ ...prev, samples: [...prev.samples, { ...EMPTY_SAMPLE }] }));
  const removeSample = (i) => setForm(prev => ({ ...prev, samples: prev.samples.filter((_, idx) => idx !== i) }));

  const handleSaveSenderProfile = async () => {
    const profile = {
      senderName: form.senderName, senderCompany: form.senderCompany,
      senderAddress: form.senderAddress, senderDEA: form.senderDEA, senderNPI: form.senderNPI
    };
    await saveSenderProfile(profile);
    setSavedSender(profile);
    setSaveMsg("Saved ✓");
    setTimeout(() => setSaveMsg(""), 2500);
  };

  const handleLoadSenderProfile = () => {
    if (savedSender) setForm(prev => ({ ...prev, ...savedSender }));
  };

  const createShipment = async () => {
    setLoading(true);
    setSheetsMsg("");
    const id = generateShipmentId();
    const payload = {
      ...form, id,
      createdAt: getAuditTimestamp(),
      status: "pending",
      cfrVersion: CFR_VERSION
    };
    payload.signature = await signPayload(payload);
    payload.auditTrail = [{ event: "CREATED", ts: getAuditTimestamp(), by: form.senderName }];

    // Save locally
    await saveShipment(payload);

    // Push SHIPMENT_CREATED row to Google Sheets
    await postToSheets({
      type: "SHIPMENT_CREATED",
      id:               payload.id,
      createdAt:        payload.createdAt,
      senderName:       payload.senderName,
      senderCompany:    payload.senderCompany,
      senderDEA:        payload.senderDEA,
      senderNPI:        payload.senderNPI,
      senderAddress:    payload.senderAddress,
      carrier:          payload.carrier,
      trackingNumber:   payload.trackingNumber,
      shipDate:         payload.shipDate,
      hcpName:          payload.hcpName,
      hcpCredentials:   payload.hcpCredentials,
      hcpPractice:      payload.hcpPractice,
      hcpNPI:           payload.hcpNPI,
      hcpDEA:           payload.hcpDEA,
      hcpAddress:       payload.hcpAddress,
      samplesSummary:   summarizeSamples(payload.samples),
      signature:        payload.signature,
    });

    const isConfigured = SHEETS_WEBHOOK_URL && !SHEETS_WEBHOOK_URL.includes("YOUR_APPS_SCRIPT")
                      && SHEETS_SECRET      && !SHEETS_SECRET.includes("YOUR_SECRET");
    setSheetsMsg(isConfigured ? "✓ Logged to Google Sheets" : "⚠ Google Sheets not configured — see code comments");

    setActive(payload);
    setView("detail");
    setLoading(false);
  };

  const qrUrl = activeShipment ? `${appUrl}?ack=${activeShipment.id}` : "";

  if (view === "acknowledge") return <AcknowledgementView shipmentId={ackId} />;

  // ── Nav ──────────────────────────────────────────────────────────────────
  const Nav = () => (
    <div style={{
      borderBottom: `1px solid ${C.border}`, padding: "16px 40px",
      display: "flex", justifyContent: "space-between", alignItems: "center",
      background: C.bgDeep, position: "sticky", top: 0, zIndex: 10
    }}>
      <div>
        <div style={{ fontSize: 9, letterSpacing: 5, color: C.accentText, textTransform: "uppercase", marginBottom: 3 }}>
          GAP Regulatory Consulting
        </div>
        <div style={{ fontSize: 18, color: C.textPrimary, letterSpacing: 1, fontFamily: "Georgia, serif" }}>
          Pharma Sample <span style={{ color: C.accentText }}>AOC</span> System
        </div>
      </div>
      <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
        <SheetsBadge />
        {view !== "dashboard" && (
          <button onClick={() => { setView("dashboard"); setShowPrint(false); }} style={{
            background: "none", border: `1px solid ${C.border}`,
            borderRadius: 4, color: C.textMid, padding: "8px 18px",
            fontSize: 11, cursor: "pointer", fontFamily: "Georgia, serif", letterSpacing: 1
          }}>← Dashboard</button>
        )}
        {view === "dashboard" && (
          <button onClick={() => {
            setForm(prev => ({
              ...prev, ...(savedSender || {}),
              hcpName: "", hcpCredentials: "", hcpPractice: "",
              hcpAddress: "", hcpNPI: "", hcpDEA: "",
              trackingNumber: "", carrier: "",
              shipDate: todayISO(),
              samples: [{ ...EMPTY_SAMPLE }]
            }));
            setView("create");
          }} style={{
            background: C.accentDim, border: "none",
            borderRadius: 4, color: C.accentText, padding: "10px 22px",
            fontSize: 11, cursor: "pointer", fontFamily: "Georgia, serif",
            letterSpacing: 2, textTransform: "uppercase"
          }}>+ New Shipment</button>
        )}
        <button onClick={() => {
          sessionStorage.removeItem("gap-aoc-auth");
          setAuthed(false);
        }} style={{
          background: "none", border: `1px solid ${C.border}`,
          borderRadius: 4, color: C.textFaint, padding: "8px 14px",
          fontSize: 10, cursor: "pointer", fontFamily: "Georgia, serif",
          letterSpacing: 1, textTransform: "uppercase"
        }}>Sign Out</button>
      </div>
    </div>
  );

  // ── DASHBOARD ────────────────────────────────────────────────────────────
  if (view === "dashboard") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Georgia, serif", color: C.textPrimary }}>
      <Nav />
      <div style={{ padding: "40px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 32 }}>
          {[
            { label: "Total Shipments",   value: shipments.length,                                           color: C.textPrimary },
            { label: "AOC Acknowledged",  value: shipments.filter(s => s.status === "acknowledged").length,  color: C.greenOk },
            { label: "Pending AOC",       value: shipments.filter(s => s.status === "pending").length,       color: C.gold },
          ].map(card => (
            <div key={card.label} style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: "20px 24px" }}>
              <div style={{ fontSize: 10, letterSpacing: 3, color: C.textDim, textTransform: "uppercase", marginBottom: 10 }}>{card.label}</div>
              <div style={{ fontSize: 38, color: card.color }}>{card.value}</div>
            </div>
          ))}
        </div>

        {savedSender && (
          <div style={{
            background: C.accentFaint, border: `1px solid ${C.accentDim}`,
            borderRadius: 6, padding: "12px 18px", marginBottom: 28,
            display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap"
          }}>
            <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase" }}>Default Sender</div>
            <div style={{ fontSize: 13, color: C.textPrimary }}>{savedSender.senderName}</div>
            <div style={{ fontSize: 12, color: C.textMid }}>{savedSender.senderCompany}</div>
            {savedSender.senderDEA && <div style={{ fontSize: 11, color: C.textDim }}>DEA: {savedSender.senderDEA}</div>}
            <div style={{ fontSize: 10, color: C.accentText, marginLeft: "auto" }}>✓ Pre-loaded on new shipments</div>
          </div>
        )}

        {shipments.length === 0 ? (
          <div style={{ textAlign: "center", padding: "80px 0", color: C.textFaint }}>
            <div style={{ fontSize: 40, marginBottom: 14 }}>◈</div>
            <div style={{ fontSize: 14, letterSpacing: 2 }}>No shipments yet. Create your first.</div>
          </div>
        ) : (
          <div>
            <div style={{ fontSize: 10, letterSpacing: 4, color: C.textDim, textTransform: "uppercase", marginBottom: 14 }}>Recent Shipments</div>
            {shipments.map(s => (
              <div key={s.id} onClick={() => { setActive(s); setView("detail"); }}
                style={{
                  background: C.bgCard, border: `1px solid ${C.border}`,
                  borderRadius: 6, padding: "16px 20px", marginBottom: 10,
                  cursor: "pointer", display: "flex", justifyContent: "space-between",
                  alignItems: "center", transition: "border-color 0.2s"
                }}
                onMouseEnter={e => e.currentTarget.style.borderColor = C.accentDim}
                onMouseLeave={e => e.currentTarget.style.borderColor = C.border}
              >
                <div>
                  <div style={{ fontSize: 13, color: C.textPrimary, marginBottom: 3, letterSpacing: 1 }}>{s.id}</div>
                  <div style={{ fontSize: 12, color: C.textMid }}>
                    {s.hcpName}{s.hcpPractice ? ` · ${s.hcpPractice}` : ""}
                  </div>
                  <div style={{ fontSize: 11, color: C.textDim, marginTop: 2 }}>
                    Shipped: {s.shipDate || new Date(s.createdAt).toLocaleDateString()} · {(s.samples || []).length} sample(s)
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                  <span style={{
                    background: s.status === "acknowledged" ? C.greenOkBg : C.goldBg,
                    border: `1px solid ${s.status === "acknowledged" ? C.greenOkBorder : C.goldDim}`,
                    color: s.status === "acknowledged" ? C.greenOk : C.gold,
                    borderRadius: 4, padding: "3px 10px", fontSize: 10, letterSpacing: 2
                  }}>
                    {s.status === "acknowledged" ? "AOC RECEIVED" : "PENDING"}
                  </span>
                  <span style={{ color: C.accent }}>→</span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      <div style={{ borderTop: `1px solid ${C.border}`, padding: "14px 40px", fontSize: 10, color: C.textFaint, letterSpacing: 2 }}>
        21 CFR PART 11 COMPLIANT · PDMA · ELECTRONIC RECORDS & SIGNATURES · GAP REGULATORY CONSULTING
      </div>
    </div>
  );

  // ── CREATE SHIPMENT ──────────────────────────────────────────────────────
  if (view === "create") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Georgia, serif", color: C.textPrimary }}>
      <Nav />
      <div style={{ padding: "40px", maxWidth: 920 }}>
        <h2 style={{ color: C.textPrimary, fontWeight: "normal", marginTop: 0, marginBottom: 6, fontSize: 22 }}>
          New Pharmaceutical Sample Shipment
        </h2>
        <p style={{ color: C.textDim, fontSize: 12, marginBottom: 36, marginTop: 0, letterSpacing: 1 }}>
          All fields are recorded as part of the 21 CFR Part 11 electronic record and logged to Google Sheets.
        </p>

        {/* ── Sender ──────────────────────────────────────────────────── */}
        <div style={{ marginBottom: 36 }}>
          <div style={{
            display: "flex", alignItems: "center", justifyContent: "space-between",
            marginBottom: 18, paddingBottom: 8, borderBottom: `1px solid ${C.border}`
          }}>
            <div style={{ fontSize: 10, letterSpacing: 4, color: C.accent, textTransform: "uppercase" }}>Sender Information</div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {saveSenderMsg && <span style={{ fontSize: 11, color: C.accentText, letterSpacing: 1 }}>{saveSenderMsg}</span>}
              {savedSender && (
                <button onClick={handleLoadSenderProfile} style={{
                  background: "none", border: `1px solid ${C.border}`, borderRadius: 4,
                  color: C.textMid, padding: "6px 14px", fontSize: 10, cursor: "pointer",
                  fontFamily: "Georgia, serif", letterSpacing: 1
                }}>↺ Reload Saved</button>
              )}
              <button onClick={handleSaveSenderProfile} style={{
                background: C.accentFaint, border: `1px solid ${C.accentDim}`, borderRadius: 4,
                color: C.accentText, padding: "6px 14px", fontSize: 10, cursor: "pointer",
                fontFamily: "Georgia, serif", letterSpacing: 1
              }}>✦ Save as Default</button>
            </div>
          </div>

          {savedSender && (
            <div style={{
              background: C.accentFaint, border: `1px solid ${C.accentDim}`,
              borderRadius: 4, padding: "10px 16px", marginBottom: 18,
              display: "flex", alignItems: "center", gap: 10, fontSize: 12
            }}>
              <span style={{ color: C.accentText }}>✓</span>
              Default sender loaded: <b style={{ color: C.textPrimary }}>{savedSender.senderCompany || savedSender.senderName}</b>
            </div>
          )}

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "16px 32px" }}>
            {/* Text inputs */}
            {[
              ["senderName",    "Sender Full Name"],
              ["senderCompany", "Company / Manufacturer"],
              ["senderDEA",     "DEA Registration #"],
              ["senderNPI",     "NPI Number"],
            ].map(([key, label]) => (
              <div key={key}>
                <label style={labelStyle}>{label}</label>
                <input value={form[key]} onChange={e => updateForm(key, e.target.value)} style={inputStyle}
                  onFocus={e => e.target.style.borderColor = C.borderFocus}
                  onBlur={e => e.target.style.borderColor = C.border} />
              </div>
            ))}

            {/* Sender Address — full-width textarea */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label style={labelStyle}>Sender Address</label>
              <textarea value={form.senderAddress} onChange={e => updateForm("senderAddress", e.target.value)}
                placeholder={""}
                style={textareaStyle}
                onFocus={e => e.target.style.borderColor = C.borderFocus}
                onBlur={e => e.target.style.borderColor = C.border} />
            </div>

            {/* Shipping info */}
            {[
              ["carrier",         "Carrier (UPS / FedEx / etc.)"],
              ["trackingNumber",  "Tracking Number"],
            ].map(([key, label]) => (
              <div key={key}>
                <label style={labelStyle}>{label}</label>
                <input value={form[key]} onChange={e => updateForm(key, e.target.value)} style={inputStyle}
                  onFocus={e => e.target.style.borderColor = C.borderFocus}
                  onBlur={e => e.target.style.borderColor = C.border} />
              </div>
            ))}

            {/* Ship Date — auto-filled, editable */}
            <div>
              <label style={labelStyle}>Ship Date</label>
              <input
                type="date"
                value={form.shipDate}
                onChange={e => updateForm("shipDate", e.target.value)}
                style={{ ...inputStyle, colorScheme: "dark" }}
                onFocus={e => e.target.style.borderColor = C.borderFocus}
                onBlur={e => e.target.style.borderColor = C.border}
              />
            </div>
          </div>
        </div>

        {/* ── HCP ─────────────────────────────────────────────────────── */}
        <div style={{ marginBottom: 36 }}>
          <div style={{ fontSize: 10, letterSpacing: 4, color: C.accent, textTransform: "uppercase", marginBottom: 18, paddingBottom: 8, borderBottom: `1px solid ${C.border}` }}>
            Receiving HCP
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "16px 32px" }}>
            {[
              ["hcpName",        "HCP Full Name"],
              ["hcpCredentials", "Credentials (MD / DO / NP / PA)"],
              ["hcpPractice",    "Practice / Institution"],
              ["hcpNPI",         "HCP NPI #"],
              ["hcpDEA",         "HCP DEA # (if applicable)"],
            ].map(([key, label]) => (
              <div key={key}>
                <label style={labelStyle}>{label}</label>
                <input value={form[key]} onChange={e => updateForm(key, e.target.value)} style={inputStyle}
                  onFocus={e => e.target.style.borderColor = C.borderFocus}
                  onBlur={e => e.target.style.borderColor = C.border} />
              </div>
            ))}

            {/* HCP Address — full-width textarea */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label style={labelStyle}>Delivery Address</label>
              <textarea value={form.hcpAddress} onChange={e => updateForm("hcpAddress", e.target.value)}
                placeholder={""}
                style={textareaStyle}
                onFocus={e => e.target.style.borderColor = C.borderFocus}
                onBlur={e => e.target.style.borderColor = C.border} />
            </div>
          </div>
        </div>

        {/* ── Samples ─────────────────────────────────────────────────── */}
        <div style={{ marginBottom: 36 }}>
          <div style={{ fontSize: 10, letterSpacing: 4, color: C.accent, textTransform: "uppercase", marginBottom: 18, paddingBottom: 8, borderBottom: `1px solid ${C.border}` }}>
            Sample Contents
          </div>
          {form.samples.map((s, i) => (
            <div key={i} style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 20, marginBottom: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={{ fontSize: 11, color: C.textDim, letterSpacing: 2 }}>SAMPLE {i + 1}</div>
                {form.samples.length > 1 && (
                  <button onClick={() => removeSample(i)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 11, letterSpacing: 1 }}>
                    ✕ Remove
                  </button>
                )}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: "12px 20px" }}>
                {[
                  ["productName", "Product Name"],
                  ["ndc",         "NDC"],
                  ["lot",         "Lot #"],
                  ["qty",         "Quantity"],
                  ["strength",    "Strength"],
                  ["form",        "Dosage Form"],
                ].map(([key, label]) => (
                  <div key={key}>
                    <label style={labelStyle}>{label}</label>
                    <input value={s[key]} onChange={e => updateSample(i, key, e.target.value)} style={inputStyle}
                      onFocus={e => e.target.style.borderColor = C.borderFocus}
                      onBlur={e => e.target.style.borderColor = C.border} />
                  </div>
                ))}
              </div>
            </div>
          ))}
          <button onClick={addSample} style={{
            background: "none", border: `1px dashed ${C.accentDim}`,
            borderRadius: 4, color: C.textMid, padding: "10px 20px",
            fontSize: 11, cursor: "pointer", fontFamily: "Georgia, serif",
            letterSpacing: 1, width: "100%", marginTop: 4
          }}>+ Add Another Sample</button>
        </div>

        {/* Compliance notice */}
        <div style={{ background: C.bgDeep, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 24, fontSize: 11, color: C.textDim, lineHeight: 1.7 }}>
          <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, marginBottom: 8, textTransform: "uppercase" }}>Compliance Notice</div>
          Creating this shipment generates a cryptographically signed electronic record (21 CFR Part 11 / PDMA), writes a row to your Google Sheets <b style={{ color: C.textMid }}>Shipments</b> tab, and generates a QR code for the HCP to complete their acknowledgement. The AOC will be recorded in the <b style={{ color: C.textMid }}>Acknowledgements</b> tab upon scan.
        </div>

        {sheetsMsg && (
          <div style={{ marginBottom: 16, fontSize: 11, color: C.accentText, letterSpacing: 1 }}>
            {sheetsMsg}
          </div>
        )}

        <button onClick={createShipment} disabled={loadingCreate} style={{
          background: loadingCreate ? C.accentFaint : C.accentDim,
          border: "none", borderRadius: 4,
          color: loadingCreate ? C.textDim : C.accentText,
          padding: "14px 32px", fontSize: 12, letterSpacing: 3,
          textTransform: "uppercase", cursor: loadingCreate ? "not-allowed" : "pointer",
          fontFamily: "Georgia, serif"
        }}>
          {loadingCreate ? "Generating…" : "Generate QR & Packing Slip"}
        </button>
      </div>
    </div>
  );

  // ── DETAIL VIEW ──────────────────────────────────────────────────────────
  if (view === "detail" && activeShipment) return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Georgia, serif", color: C.textPrimary }}>
      <Nav />
      <div style={{ padding: "40px" }}>
        <div style={{ display: "flex", gap: 40, alignItems: "flex-start" }}>
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 24 }}>
              <h2 style={{ color: C.textPrimary, fontWeight: "normal", margin: 0, fontSize: 22 }}>
                {activeShipment.id}
              </h2>
              <span style={{
                background: activeShipment.status === "acknowledged" ? C.greenOkBg : C.goldBg,
                border: `1px solid ${activeShipment.status === "acknowledged" ? C.greenOkBorder : C.goldDim}`,
                color: activeShipment.status === "acknowledged" ? C.greenOk : C.gold,
                borderRadius: 4, padding: "4px 12px", fontSize: 10, letterSpacing: 2
              }}>
                {activeShipment.status === "acknowledged" ? "AOC RECEIVED" : "PENDING ACKNOWLEDGEMENT"}
              </span>
            </div>

            {activeShipment.status === "acknowledged" && activeShipment.aoc && (
              <div style={{ background: C.greenOkBg, border: `1px solid ${C.greenOkBorder}`, borderRadius: 6, padding: 20, marginBottom: 20 }}>
                <div style={{ fontSize: 10, letterSpacing: 3, color: C.greenOk, textTransform: "uppercase", marginBottom: 10 }}>AOC Confirmed · Recorded in Google Sheets</div>
                <div style={{ fontSize: 12, color: "#6ad4b4", lineHeight: 1.9 }}>
                  <div>Acknowledged by: <b>{activeShipment.aoc.acknowledgedBy}</b>{activeShipment.aoc.acknowledgedTitle ? `, ${activeShipment.aoc.acknowledgedTitle}` : ""}</div>
                  <div>Timestamp: {activeShipment.aoc.timestamp}</div>
                  <div style={{ fontSize: 10, color: C.textFaint, marginTop: 8, wordBreak: "break-all" }}>Sig: {activeShipment.aoc.signature}</div>
                </div>
              </div>
            )}

            <div style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 20, marginBottom: 16 }}>
              <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 14 }}>Shipment Details</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px 24px" }}>
                {[
                  ["From",       activeShipment.senderName],
                  ["Company",    activeShipment.senderCompany || "—"],
                  ["To",         `${activeShipment.hcpName}${activeShipment.hcpCredentials ? `, ${activeShipment.hcpCredentials}` : ""}`],
                  ["Practice",   activeShipment.hcpPractice || "—"],
                  ["Carrier",    activeShipment.carrier || "—"],
                  ["Tracking",   activeShipment.trackingNumber || "—"],
                  ["Ship Date",  activeShipment.shipDate || "—"],
                  ["Created",    new Date(activeShipment.createdAt).toLocaleString()],
                ].map(([k, v]) => (
                  <div key={k}>
                    <div style={{ color: C.textDim, fontSize: 10, letterSpacing: 1, textTransform: "uppercase", marginBottom: 2 }}>{k}</div>
                    <div style={{ color: C.textPrimary }}>{v}</div>
                  </div>
                ))}
                {/* Addresses full-width */}
                {[
                  ["Sender Address", activeShipment.senderAddress],
                  ["Delivery Address", activeShipment.hcpAddress],
                ].map(([k, v]) => v ? (
                  <div key={k} style={{ gridColumn: "1 / -1" }}>
                    <div style={{ color: C.textDim, fontSize: 10, letterSpacing: 1, textTransform: "uppercase", marginBottom: 2 }}>{k}</div>
                    <div style={{ color: C.textPrimary, whiteSpace: "pre-wrap" }}>{v}</div>
                  </div>
                ) : null)}
              </div>
            </div>

            <div style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 20, marginBottom: 16 }}>
              <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 14 }}>Sample Contents</div>
              {(activeShipment.samples || []).map((s, i) => (
                <div key={i} style={{
                  borderBottom: i < activeShipment.samples.length - 1 ? `1px solid ${C.border}` : "none",
                  padding: "8px 0", display: "flex", gap: 20, flexWrap: "wrap"
                }}>
                  <span style={{ color: C.textPrimary, fontWeight: "bold" }}>{s.productName}</span>
                  <span style={{ color: C.textMid }}>NDC: {s.ndc}</span>
                  <span style={{ color: C.textMid }}>Lot: {s.lot}</span>
                  <span style={{ color: C.textMid }}>Qty: {s.qty}</span>
                  <span style={{ color: C.textMid }}>{s.strength} {s.form}</span>
                </div>
              ))}
            </div>

            <div style={{ padding: 16, background: C.bgDeep, border: `1px dashed ${C.borderLight}`, borderRadius: 6 }}>
              <div style={{ fontSize: 10, color: C.textDim, letterSpacing: 2, marginBottom: 10 }}>SIMULATE RECIPIENT SCAN</div>
              <button onClick={() => { setAckId(activeShipment.id); setView("acknowledge"); }} style={{
                background: "none", border: `1px solid ${C.accentDim}`, borderRadius: 4,
                color: C.accentText, padding: "9px 18px", fontSize: 11,
                cursor: "pointer", fontFamily: "Georgia, serif", letterSpacing: 1
              }}>Open Acknowledgement View →</button>
            </div>
          </div>

          {/* QR Panel */}
          <div style={{ width: 280, flexShrink: 0 }}>
            <div style={{ background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
              <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 16 }}>AOC QR Code</div>
              <div style={{ background: "white", borderRadius: 8, padding: 14, display: "inline-block", marginBottom: 14 }}>
                <QRCode value={qrUrl} size={180} />
              </div>
              <div style={{ fontSize: 10, color: C.textDim, lineHeight: 1.6, wordBreak: "break-all", marginBottom: 12 }}>
                {qrUrl}
              </div>
              <div style={{ fontSize: 9, color: C.textFaint, letterSpacing: 1 }}>
                Tamper-evident · 21 CFR §11.200<br />Sig: {activeShipment.signature?.substring(0, 24)}…
              </div>
            </div>

            <button onClick={() => setShowPrint(!showPrint)} style={{
              width: "100%",
              background: showPrint ? C.accentDim : C.bgCard,
              border: `1px solid ${showPrint ? C.accent : C.border}`,
              borderRadius: 4,
              color: showPrint ? C.accentText : C.textMid,
              padding: "11px 0", fontSize: 11, cursor: "pointer",
              fontFamily: "Georgia, serif", letterSpacing: 2, textTransform: "uppercase"
            }}>
              {showPrint ? "Hide" : "Preview"} Packing Slip
            </button>
          </div>
        </div>

        {showPrint && (
          <div style={{ marginTop: 36 }}>
            <div style={{ fontSize: 10, letterSpacing: 3, color: C.accent, textTransform: "uppercase", marginBottom: 16 }}>
              Packing Slip Preview
            </div>
            <PackingSlip shipment={activeShipment} qrUrl={qrUrl} />
          </div>
        )}
      </div>

      <div style={{ borderTop: `1px solid ${C.border}`, padding: "14px 40px", fontSize: 10, color: C.textFaint, letterSpacing: 2 }}>
        21 CFR PART 11 COMPLIANT · PDMA · ELECTRONIC RECORDS & SIGNATURES · GAP REGULATORY CONSULTING
      </div>
    </div>
  );

  return null;
}
