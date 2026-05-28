import React, { useState, useRef, useCallback, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════════
   BANK2TALLY — Commercial Grade v2.0
   Full pipeline: Auth → Upload → Column Map → Ledger → Preview → Export
   Features: Multi-company, Tally XML, PDF OCR, Duplicate detection,
   Bank templates, BRS, Audit trail, Role-based access,
   Supabase Auth (admin + user with approval flow)
═══════════════════════════════════════════════════════════════════ */

// ── Supabase Client ───────────────────────────────────────────────
const SUPABASE_URL  = "https://jiuvpncqxvntrrmjpmuv.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppdXZwbmNxeHZudHJybWpwbXV2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkxNzkxMDMsImV4cCI6MjA5NDc1NTEwM30.ojAykG5pZLZEusjKSD5D_6N6mGiDssr3NWAXDm9rnak";

// Lightweight Supabase REST helper (no SDK dependency)
const sb = {
  _headers: () => ({
    "Content-Type": "application/json",
    "apikey": SUPABASE_ANON,
    "Authorization": `Bearer ${sb._token || SUPABASE_ANON}`,
  }),
  _token: null,

  // Auth: sign in
  async signIn(email, password) {
    let res, data;
    try {
      res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "apikey": SUPABASE_ANON,
          "x-client-info": "bank2tally/2.0",
        },
        body: JSON.stringify({ email, password }),
      });
      data = await res.json();
    } catch (e) {
      throw new Error("Network error — check your internet connection and try again.");
    }
    if (!res.ok) throw new Error(data.error_description || data.msg || data.message || "Invalid email or password.");
    sb._token = data.access_token;
    sb._refreshToken = data.refresh_token || null;
    try { localStorage.setItem("sb_session", JSON.stringify(data)); } catch {}
    return data;
  },

  // Auth: refresh expired token using refresh_token
  async refreshSession() {
    const stored = (() => { try { return JSON.parse(localStorage.getItem("sb_session")||"{}"); } catch { return {}; } })();
    const rt = stored.refresh_token;
    if (!rt) throw new Error("No refresh token available");
    const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON },
      body: JSON.stringify({ refresh_token: rt }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error_description || "Session expired — please sign in again.");
    sb._token = data.access_token;
    sb._refreshToken = data.refresh_token || rt;
    try { localStorage.setItem("sb_session", JSON.stringify(data)); } catch {}
    return data;
  },

  // Auth: sign up
  async signUp(email, password, meta) {
    let res, data;
    try {
      res = await fetch(`${SUPABASE_URL}/auth/v1/signup`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "apikey": SUPABASE_ANON,
          "x-client-info": "bank2tally/2.0",
        },
        body: JSON.stringify({ email, password, data: meta }),
      });
      data = await res.json();
    } catch (e) {
      throw new Error("Network error — check your internet connection and try again.");
    }
    if (!res.ok) throw new Error(data.error_description || data.msg || data.message || "Signup failed.");
    // Do NOT overwrite sb._token here — if an admin is adding a user, we must preserve
    // the admin's own session token. Only set token for self-registration flows.
    // Callers that need the new user's token should use data.access_token directly.
    return data;
  },

  // Auth: sign out
  async signOut() {
    try {
      await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
        method: "POST", headers: sb._headers(),
      });
    } catch {}
    sb._token = null;
  },

  // DB: auto-retry helper — refreshes token on 401 and retries once
  async _fetch(url, opts) {
    let res = await fetch(url, opts);
    if (res.status === 401) {
      try {
        await sb.refreshSession();
        opts.headers = { ...opts.headers, "Authorization": `Bearer ${sb._token}` };
        res = await fetch(url, opts);
      } catch { throw new Error("Session expired — please sign in again."); }
    }
    const data = res.status === 204 ? {} : await res.json();
    if (!res.ok) throw new Error(data.message || data.error_description || JSON.stringify(data));
    return data;
  },

  // DB: generic select
  async from(table, query = "") {
    return sb._fetch(`${SUPABASE_URL}/rest/v1/${table}${query ? "?" + query : ""}`, {
      headers: { ...sb._headers(), "Prefer": "return=representation" },
    });
  },

  // DB: update
  async update(table, match, payload) {
    const q = Object.entries(match).map(([k,v]) => `${k}=eq.${encodeURIComponent(v)}`).join("&");
    return sb._fetch(`${SUPABASE_URL}/rest/v1/${table}?${q}`, {
      method: "PATCH",
      headers: { ...sb._headers(), "Prefer": "return=representation" },
      body: JSON.stringify(payload),
    });
  },

  // DB: insert
  async insert(table, payload) {
    return sb._fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
      method: "POST",
      headers: { ...sb._headers(), "Prefer": "return=representation" },
      body: JSON.stringify(payload),
    });
  },
  // Auth: update password for the currently authenticated user (used after recovery link)
  async updatePassword(newPassword) {
    const res = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        "apikey": SUPABASE_ANON,
        "Authorization": `Bearer ${sb._token}`,
      },
      body: JSON.stringify({ password: newPassword }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error_description || data.message || "Failed to update password.");
    return data;
  },
};

// ── Design Tokens ────────────────────────────────────────────────
const T = {
  // Backgrounds — clean premium light
  bg:          "#f0f4f9",
  surface:     "#ffffff",
  card:        "#ffffff",
  cardHover:   "#f8fafd",
  // Borders — soft, visible
  border:      "#dde3ee",
  borderLight: "#e8edf6",
  borderMid:   "#c8d3e8",
  // Accent blue — rich, professional
  accent:      "#2563eb",
  accentDim:   "#dbeafe",
  accentGlow:  "rgba(37,99,235,0.15)",
  accentSoft:  "rgba(37,99,235,0.06)",
  // Semantic colours — vivid on light
  green:       "#059669",
  greenDim:    "#d1fae5",
  greenMid:    "#a7f3d0",
  amber:       "#d97706",
  amberDim:    "#fef3c7",
  red:         "#dc2626",
  redDim:      "#fee2e2",
  purple:      "#7c3aed",
  purpleDim:   "#ede9fe",
  gold:        "#b45309",
  goldDim:     "#fef3c7",
  teal:        "#0891b2",
  tealDim:     "#cffafe",
  // Text — strong contrast on white/light
  text:        "#0f172a",     // headings, primary content
  textHigh:    "#000000",     // maximum contrast labels
  textMid:     "#334155",     // secondary text — clearly readable
  textSub:     "#64748b",     // tertiary — subtle but visible
  textDim:     "#94a3b8",     // decorative only — NOT for readable text
  // Brand
  font:        "'DM Sans', 'Inter', 'Segoe UI', sans-serif",
  mono:        "'JetBrains Mono', 'Fira Code', 'Consolas', monospace",
};

// ── Motivational Quotes ──────────────────────────────────────────
const QUOTES = [
  { text: "Accounting is the language of business. Speak it fluently.", author: "Warren Buffett (adapted)" },
  { text: "The secret of getting ahead is getting started.", author: "Mark Twain" },
  { text: "Beware of little expenses. A small leak will sink a great ship.", author: "Benjamin Franklin" },
  { text: "An investment in knowledge pays the best interest.", author: "Benjamin Franklin" },
  { text: "It's not about having time. It's about making time.", author: "Unknown" },
  { text: "Do not save what is left after spending; spend what is left after saving.", author: "Warren Buffett" },
  { text: "Success is the sum of small efforts repeated day in and day out.", author: "Robert Collier" },
  { text: "Financial freedom is available to those who learn about it and work for it.", author: "Robert Kiyosaki" },
];

const todayQuote = QUOTES[new Date().getDate() % QUOTES.length];

// ── Constants ────────────────────────────────────────────────────
// No hardcoded companies — all fetched live from Tally gateway
const TALLY_COMPANIES_FALLBACK = []; // empty until gateway responds

/* ── Tally Gateway Layer ─────────────────────────────────────────
   All requests go to the Tally HTTP/XML gateway (default port 9000).
   Tally must be running with Gateway enabled:
     Tally Prime  → F12 > Advanced Config > Enable Tally Gateway Server
     Tally ERP 9  → F12 > Configure > Enable ODBC/HTTP Server
   CORS note: In production wrap in a local proxy (e.g. electron, or
   a tiny Express server on localhost:3001 that forwards to :9000).
─────────────────────────────────────────────────────────────────── */


// ── Bank2Tally Universal Connector ──────────────────────────────
// Uses Chrome Extension to bypass HTTPS→HTTP mixed content restriction
// Extension communicates with Tally directly on localhost

// Check if extension is installed
let _extensionReady = false;
function _markExtensionReady() {
  _extensionReady = true;
  window.__bank2tallyExtension = true; // sync both flags
}
window.addEventListener("message", (e) => {
  // Only accept messages from same origin or extension (null origin = extension)
  if (e.origin && e.origin !== window.location.origin && !e.origin.startsWith("chrome-extension://")) return;
  if (e.data?.type === "BANK2TALLY_EXTENSION_PRESENT") {
    _markExtensionReady();
  }
});
// Ping on load — extension may have already sent its signal before listener registered
setTimeout(() => {
  window.postMessage({ type: "CHECK_EXTENSION" }, "*");
}, 100);
setTimeout(() => {
  window.postMessage({ type: "CHECK_EXTENSION" }, "*");
}, 1500);

// Send message to extension and wait for response.
// Handles multiple response naming conventions since different extension builds
// may respond with: TYPE_RESPONSE, TALLY_RESPONSE, or BANK2TALLY_RESPONSE.
function sendToExtension(msg, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const requestId = Math.random().toString(36).slice(2);
    // Accept any of these response type names
    const acceptedTypes = new Set([
      msg.type + "_RESPONSE",      // e.g. TALLY_REQUEST_RESPONSE
      "TALLY_RESPONSE",            // generic response
      "TALLY_REQUEST_RESPONSE",    // explicit
      "BANK2TALLY_RESPONSE",       // alternate naming
      "TALLY_PING_RESPONSE",       // ping response
    ]);
    const timer = setTimeout(() => {
      window.removeEventListener("message", handler);
      reject(new Error("Extension timed out — make sure Tally is open and the extension is active"));
    }, timeoutMs);

    function handler(e) {
      if (!e.data?.type) return;
      // Match by type (with OR without requestId — support both strict and broadcast modes)
      const typeMatch = acceptedTypes.has(e.data.type);
      const idMatch = !e.data.requestId || e.data.requestId === requestId;
      if (typeMatch && idMatch) {
        clearTimeout(timer);
        window.removeEventListener("message", handler);
        resolve(e.data);
      }
    }
    window.addEventListener("message", handler);
    window.postMessage({ ...msg, requestId }, "*");
  });
}

// ── Direct HTTP fetch to Tally gateway ───────────────────────────
// Tally's HTTP gateway on port 9000 accepts plain POST with XML body.
// When running on HTTPS (vercel.app) direct fetch is blocked by mixed-content.
// We try it anyway — on localhost / http origins it works directly.
// If it fails, we fall back to the extension message bus.
async function tallyDirectFetch(host, port, xmlBody, timeoutMs = 10000) {
  const url = `http://${host}:${port}`;
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/xml; charset=utf-8" },
      body: xmlBody,
      signal: ctrl.signal,
    });
    clearTimeout(timer);
    if (!res.ok) throw new Error(`Tally HTTP ${res.status}`);
    return await res.text();
  } catch (e) {
    clearTimeout(timer);
    throw e;
  }
}

// ── Extension message bus (fallback when direct fetch blocked) ────
// CRITICAL FIX: The extension expects message type "TALLY_XML_REQUEST" — NOT "TALLY_REQUEST".
// Sending "TALLY_REQUEST" causes the extension to trigger its own internal TDL Form:Company
// dialog which has no PARTS defined → crashes Tally with "Error in TDL. No 'PARTS'!".
// The correct type is "TALLY_XML_REQUEST" which tells the extension to pass-through our XML.
async function tallyExtensionFetch(host, port, xmlBody, timeoutMs = 10000) {
  if (!_extensionReady) throw new Error("EXTENSION_NOT_READY");
  return new Promise((resolve, reject) => {
    const requestId = Math.random().toString(36).slice(2);
    const acceptedTypes = new Set([
      "TALLY_XML_REQUEST_RESPONSE",
      "TALLY_XML_RESPONSE",
      "TALLY_RESPONSE",
      "BANK2TALLY_RESPONSE",
      "TALLY_REQUEST_RESPONSE",   // legacy — some extension builds still use this
    ]);
    const timer = setTimeout(() => {
      window.removeEventListener("message", handler);
      reject(new Error("Extension timed out — make sure Tally is open and the extension is active"));
    }, timeoutMs);
    function handler(e) {
      if (!e.data?.type) return;
      const typeMatch = acceptedTypes.has(e.data.type);
      const idMatch = !e.data.requestId || e.data.requestId === requestId;
      if (typeMatch && idMatch) {
        clearTimeout(timer);
        window.removeEventListener("message", handler);
        if (e.data.success === false || e.data.error) {
          reject(new Error(e.data.error || "Extension returned error"));
        } else {
          resolve(e.data.data || e.data.body || "");
        }
      }
    }
    window.addEventListener("message", handler);
    // Send with corrected type — "TALLY_XML_REQUEST" not "TALLY_REQUEST"
    window.postMessage({ type: "TALLY_XML_REQUEST", requestId, host, port, body: xmlBody }, "*");
  });
}

// ── Master tallyPost — direct fetch first, extension fallback ─────
async function tallyPost(host, port, xmlBody, timeoutMs = 10000) {
  // Strategy 1: direct HTTP fetch (works on http:// origins and localhost)
  try {
    const result = await tallyDirectFetch(host, port, xmlBody, Math.min(timeoutMs, 8000));
    if (result) {
      _usedDirectFetch = true;
      return result;
    }
  } catch (directErr) {
    // Mixed-content block or network error — fall through to extension
    if (directErr.name === "TypeError" && directErr.message.includes("Failed to fetch")) {
      _usedDirectFetch = false;
    } else if (directErr.name === "AbortError") {
      throw new Error("Tally not responding — make sure Tally is open and Gateway is enabled");
    }
    // Other errors fall through to extension
  }

  // Strategy 2: Chrome extension message bus
  if (!_extensionReady) {
    throw new Error(
      "Cannot reach Tally.\n\n" +
      "Option A (Recommended): Run the app on http://localhost instead of the Vercel URL, so direct HTTP to Tally works.\n\n" +
      "Option B: Install the Bank2Tally Connector Chrome extension and make sure it is enabled."
    );
  }
  return await tallyExtensionFetch(host, port, xmlBody, timeoutMs);
}

let _usedDirectFetch = false; // track which method succeeded

// Parse Tally XML response — extract text content of all matching tags
function parseTallyTags(xml, tag) {
  const re = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "gi");
  const results = [];
  let m;
  while ((m = re.exec(xml)) !== null) results.push(m[1].trim());
  return results;
}

// Fetch all companies from Tally — tries every known request format
async function fetchTallyCompanies(host, port) {
  // These are the correct XML formats Tally Prime actually responds to:

  // Format A: Get currently OPEN company via SVCURRENTCOMPANY system variable
  const xmlA = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>Company Info</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  // Format B: List of Ledgers (lighter than Day Book)
  const xmlB = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Accounts</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  // Format C: Direct collection request — the most reliable for Tally Prime 3.x
  const xmlC = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Primary Groups</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  // Format D: TDL-style collection export (works on all Tally Gateway versions)
  const xmlD = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>COMPANY</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  // Format E: Fetch via GETCOLLECTION — works on Tally Prime 2.x+
  const xmlE = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY></SVCURRENTCOMPANY></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  const formats = [xmlA, xmlC, xmlD, xmlB, xmlE]; // xmlE (List of Companies) last — can crash older Tally
  let lastRaw = "";

  for (const xml of formats) {
    try {
      const raw = await tallyPost(host, port, xml, 8000);
      lastRaw = raw;
      if (raw.includes("LINEERROR")) continue; // try next format

      // Try all known company/name tags
      const candidates = [
        ...parseTallyTags(raw, "COMPANY"),
        ...parseTallyTags(raw, "BASICCOMPANYNAME"),
        ...parseTallyTags(raw, "COMPANYNAME"),
        ...parseTallyTags(raw, "NAME"),
        ...parseTallyTags(raw, "REMOTECMPINFO\.LIST"),
      ].map(n => n.trim()).filter(n => n && !n.includes("<") && n.length > 1 && n.length < 100);

      const unique = [...new Set(candidates)];
      if (unique.length > 0) {
        // console.info("Tally companies found:", unique); // debug only
        return unique.map((name, i) => ({ id: `tc${i}`, name, gstin: "", state: "", fy: "2024-25" }));
      }

      // If we got a valid XML response but no company tags,
      // the company name may be in COMPANYINFO or header — try extracting it
      const headerMatch = raw.match(/<SVCURRENTCOMPANY[^>]*>([^<]+)<\/SVCURRENTCOMPANY>/i);
      if (headerMatch?.[1]?.trim()) {
        const name = headerMatch[1].trim();
        // debug: company from header
        return [{ id: "tc0", name, gstin: "", state: "", fy: "2024-25" }];
      }
    } catch(e) {
      if (e.message === "EXTENSION_NOT_READY") throw e;
      lastRaw = e.message;
    }
  }

  // Last resort — if Tally responded at all, ask user to enter company manually
  // Tally company fetch exhausted — handled by error throw below
  throw new Error("Tally is connected but could not read company list. Please enter your company name manually below.");
}

// Fetch company details (GSTIN, state, FY) for a specific company
async function fetchTallyCompanyDetails(host, port, companyName) {
  const xml = `<ENVELOPE>
    <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
    <BODY><EXPORTDATA><REQUESTDESC>
      <REPORTNAME>Company Info</REPORTNAME>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>${companyName.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</SVCURRENTCOMPANY>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </REQUESTDESC></EXPORTDATA></BODY>
  </ENVELOPE>`;
  try {
    const raw = await tallyPost(host, port, xml, 4000);
    const gstin = (parseTallyTags(raw, "INCOMETAXNUMBER")[0] || parseTallyTags(raw, "GSTREGISTRATIONNUMBER")[0] || "").trim();
    const state = (parseTallyTags(raw, "STATENAME")[0] || "").trim();
    const fyStart = (parseTallyTags(raw, "BOOKSBEGINNINGFROM")[0] || "").trim();
    let fy = "2024-25";
    if (fyStart) {
      const d = new Date(fyStart);
      if (!isNaN(d)) {
        const y = d.getFullYear();
        const m = d.getMonth(); // 0=Jan
        fy = m >= 3 ? `${y}-${String(y+1).slice(2)}` : `${y-1}-${String(y).slice(2)}`;
      }
    }
    return { gstin, state, fy };
  } catch { return { gstin: "", state: "", fy: "2024-25" }; }
}

// Fetch all ledgers for a company from Tally
async function fetchTallyLedgers(host, port, companyName) {
  const xml = `<ENVELOPE>
    <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
    <BODY><EXPORTDATA><REQUESTDESC>
      <REPORTNAME>List of Accounts</REPORTNAME>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>${companyName.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</SVCURRENTCOMPANY>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </REQUESTDESC></EXPORTDATA></BODY>
  </ENVELOPE>`;
  const raw = await tallyPost(host, port, xml, 6000);
  const names = parseTallyTags(raw, "NAME").filter(n => n && !n.includes("<") && n.length < 80);
  // Also try LEDGER tag
  const led = parseTallyTags(raw, "LEDGER").filter(n => n && !n.includes("<") && n.length < 80);
  const all = [...new Set([...names, ...led])].filter(Boolean);
  return all.length ? all : null; // null = use built-in fallback
}

// Test connection — tries direct fetch first, then extension
// Uses the safest possible Tally XML request (License Info — zero data processing)
async function testTallyConnection(host, port) {
  // Safest ping XML — Tally just returns license info, no company data needed
  const pingXml = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>License Info</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;
  const fallbackXml = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Primary Groups</REPORTNAME><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

  // Try direct HTTP first (bypasses extension entirely — no TDL Form:Company bug)
  for (const xml of [pingXml, fallbackXml]) {
    try {
      const result = await tallyDirectFetch(host, port, xml, 6000);
      if (result && result.includes("<")) {
        _usedDirectFetch = true;
        return true; // Direct fetch works — best case
      }
    } catch (e) {
      if (e.name === "AbortError") {
        throw new Error(
          "Tally is not responding on " + host + ":" + port + ".\n\n" +
          "Please ensure:\n" +
          "1. Tally Prime is open\n" +
          "2. Gateway is enabled: F12 → Advanced Config → Enable Tally Gateway Server\n" +
          "3. Port " + port + " is not blocked by firewall"
        );
      }
      // TypeError = mixed-content block — try extension next
    }
  }

  // Extension fallback
  if (!_extensionReady) {
    throw new Error(
      "Cannot reach Tally directly (mixed-content blocked on HTTPS).\n\n" +
      "Fix options:\n" +
      "• Install the Bank2Tally Connector Chrome extension (see below)\n" +
      "• Or run the app at http://localhost:3000 instead of the Vercel URL"
    );
  }

  // Try via extension with correct message type
  for (const xml of [pingXml, fallbackXml]) {
    try {
      const result = await tallyExtensionFetch(host, port, xml, 6000);
      if (result) return true;
    } catch (e) {
      if (e.message === "EXTENSION_NOT_READY") throw e;
      // Try next format
    }
  }
  throw new Error(
    "Extension is installed but Tally is not responding.\n\n" +
    "In Tally Prime: F12 → Advanced Config → Enable Tally Gateway Server → set port to " + port
  );
}

// ── useTallyGateway hook ─────────────────────────────────────────
// Provides live companies + ledgers from the Tally gateway.
// Falls back to built-in statics if Tally is unreachable.
function useTallyGateway(host = "localhost", port = "9000") {
  const [status, setStatus] = useState("idle"); // idle | connecting | ok | error
  const [companies, setCompanies] = useState([]);
  const [ledgerMap, setLedgerMap] = useState({}); // companyName → string[]
  const [error, setError] = useState("");
  const [lastFetch, setLastFetch] = useState(null);

  const fetch_ = useCallback(async (h, p) => {
    setStatus("connecting"); setError("");
    try {
      const cos = await fetchTallyCompanies(h, p);
      _markExtensionReady(); // connection worked — mark as ready
      setCompanies(cos);
      setStatus("ok");
      setLastFetch(new Date());
      // Fetch ledgers + details in parallel — using Promise.allSettled to handle errors per company
      Promise.allSettled(cos.map(co => fetchTallyLedgers(h, p, co.name).then(ledgers => {
        if (ledgers) setLedgerMap(m => ({ ...m, [co.name]: ledgers }));
      })));
      Promise.allSettled(cos.map((co, i) => fetchTallyCompanyDetails(h, p, co.name).then(details => {
        setCompanies(cs => cs.map((c, j) => j === i ? { ...c, ...details } : c));
      })));
    } catch (e) {
      if (e.message === "EXTENSION_NOT_READY") {
        // Extension wasn't ready yet — stay idle, the useEffect listener will retry
        setStatus("idle"); setError("");
      } else {
        setStatus("error"); setError(e.message);
        setCompanies([]);
      }
    }
  }, []);

  // Connect once extension is ready — listen for its presence signal
  useEffect(() => {
    // If already ready (e.g. page reload with extension active), connect immediately
    if (_extensionReady) {
      fetch_(host, port);
      return;
    }
    // Otherwise wait for the extension presence message
    const handler = (e) => {
      if (e.data?.type === "BANK2TALLY_EXTENSION_PRESENT") {
        fetch_(host, port);
      }
    };
    window.addEventListener("message", handler);
    // Send pings to prompt the extension to identify itself
    window.postMessage({ type: "CHECK_EXTENSION" }, "*");
    const ping2 = setTimeout(() => window.postMessage({ type: "CHECK_EXTENSION" }, "*"), 1000);
    const ping3 = setTimeout(() => window.postMessage({ type: "CHECK_EXTENSION" }, "*"), 3000);
    // After 5s with no response, mark as error so UI shows "Offline"
    const giveUp = setTimeout(() => {
      setStatus(s => s === "connecting" || s === "idle" ? "error" : s);
      setError("Bank2Tally Connector extension not detected. Install it from Settings.");
    }, 5000);
    return () => {
      window.removeEventListener("message", handler);
      clearTimeout(ping2); clearTimeout(ping3); clearTimeout(giveUp);
    };
  }, []); // eslint-disable-line

  // Re-fetch when host/port change (call manually or via refetch)
  const refetch = useCallback((h, p) => fetch_(h || host, p || port), [fetch_, host, port]);

  return { status, companies, ledgerMap, error, lastFetch, refetch }; // error is exposed for UI use
}

const BANK_TEMPLATES = {
  // ── Public Sector ──────────────────────────────────────────────
  sbi:     { name:"SBI",           cols:{ date:"Txn Date",          narration:"Description",         debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Ref No./Cheque No." }},
  pnb:     { name:"PNB",           cols:{ date:"Date",              narration:"Particulars",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Ref. No." }},
  bob:     { name:"Bank of Baroda",cols:{ date:"Txn Date",          narration:"Description",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference No." }},
  boi:     { name:"Bank of India", cols:{ date:"Date",              narration:"Narration",            debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Chq No" }},
  canara:  { name:"Canara Bank",   cols:{ date:"Date",              narration:"Description",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference No" }},
  union:   { name:"Union Bank",    cols:{ date:"Date",              narration:"Narration",            debit:"Debit Amount",           credit:"Credit Amount",          balance:"Balance",              ref:"Reference Number" }},
  // ── Private Sector ─────────────────────────────────────────────
  hdfc:    { name:"HDFC Bank",     cols:{ date:"Date",              narration:"Narration",            debit:"Withdrawal Amt.",        credit:"Deposit Amt.",           balance:"Closing Balance",      ref:"Chq./Ref.No." }},
  icici:   { name:"ICICI Bank",    cols:{ date:"Date",               narration:"Description",          debit:"Withdrawal (Dr)",        credit:"Deposit (Cr)",           balance:"Available Balance",    ref:"Transaction ID" }},  // "Date" matches both "Date" and "Value Date" via prefix
  axis:    { name:"Axis Bank",     cols:{ date:"Tran Date",         narration:"PARTICULARS",          debit:"DR",                     credit:"CR",                     balance:"BAL",                  ref:"CHQNO" }},
  kotak:   { name:"Kotak Bank",    cols:{ date:"Transaction Date",  narration:"Description",          debit:"Debit Amount",           credit:"Credit Amount",          balance:"Balance",              ref:"Reference No" }},
  yes:     { name:"Yes Bank",      cols:{ date:"Date",              narration:"Transaction Details",  debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference Number" }},
  idfc:    { name:"IDFC First",    cols:{ date:"Date",              narration:"Transaction Remarks",  debit:"Debit Amount",           credit:"Credit Amount",          balance:"Balance",              ref:"Transaction ID" }},
  indus:   { name:"IndusInd Bank", cols:{ date:"Date",              narration:"Description",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference No." }},
  rbl:     { name:"RBL Bank",      cols:{ date:"Txn Date",          narration:"Description",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Ref No." }},
  centralbank: { name:"Central Bank", cols:{ date:"Value Date",       narration:"Details",              debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Chq.No." }},
  andhra:  { name:"Andhra/Union Bank",cols:{ date:"Tran Date",        narration:"Remarks",              debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Tran Id" }},
  pnb:     { name:"PNB",             cols:{ date:"Transaction Date",  narration:"Narration",            debit:"Withdrawal",             credit:"Deposit",                balance:"Balance",              ref:"Cheque Number" }},
  boi:     { name:"Bank of India",   cols:{ date:"Txn Date",          narration:"Description",          debit:"Withdrawal",             credit:"Deposits",               balance:"Balance",              ref:"Cheque No" }},
  bob:     { name:"Bank of Baroda",  cols:{ date:"Value Date",        narration:"Details",              debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Chq.No." }},
  federal: { name:"Federal Bank",  cols:{ date:"Transaction Date",  narration:"Particulars",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference No" }},
  iob:     { name:"IOB",           cols:{ date:"Date",              narration:"Particulars",          debit:"Debit",                  credit:"Credit",                 balance:"Balance",              ref:"Reference No" }},
  // ── Credit Cards / Other ───────────────────────────────────────
  amex:    { name:"Amex Card",     cols:{ date:"Date",              narration:"Description",          crdr:"Amount",                  crdrFlag:"CR/DR",                balance:"Balance",              ref:"Reference" }},
};

const TALLY_LEDGERS = [
  { group: "Suspense", items: ["Suspense Account"] },
  { group: "Income", items: ["Sales Account", "Commission Received", "Interest Received", "Rent Received", "Discount Received"] },
  { group: "Expenses", items: ["Purchase Account", "Bank Charges", "Salary & Wages", "Rent Expenses", "Office Expenses", "Travel Expenses", "Advertisement Expenses", "Printing & Stationery", "Telephone Expenses", "Electricity Charges", "Repair & Maintenance", "Legal & Professional Charges", "Audit Fees", "Miscellaneous Expenses"] },
  { group: "Tax", items: ["GST Payable", "IGST Payable", "CGST Payable", "SGST Payable", "TDS Payable", "TCS Payable", "Advance Tax"] },
  { group: "Bank & Cash", items: ["HDFC Bank", "ICICI Bank", "SBI Bank", "Axis Bank", "Kotak Bank", "Petty Cash", "Cash in Hand"] },
  { group: "Parties", items: ["Sundry Debtors", "Sundry Creditors", "Director Loan", "Shareholder Loan"] },
  { group: "Capital", items: ["Capital Account", "Drawings Account", "Retained Earnings", "Share Capital", "Securities Premium"] },
];
const ALL_LEDGERS = TALLY_LEDGERS.flatMap(g => g.items);

const VOUCHER_TYPES = ["Receipt", "Payment", "Contra", "Journal", "Sales", "Purchase"];

const SCREENS = { LOGIN: -1, UPLOAD: 0, COLUMN_MAP: 1, LEDGER: 2, PREVIEW: 3, HISTORY: 4, SETTINGS: 5, DASHBOARD: 6, USER_MGMT: 7 };

// User data is loaded live from Supabase — no hardcoded users

// ── Helpers ──────────────────────────────────────────────────────
const genId = () => Math.random().toString(36).slice(2, 9);
const fmt = n => n == null || n === "" ? "" : Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
// Parse any Indian/ISO date string into a JS Date (noon IST to avoid timezone shift)
const parseIndianDate = v => {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const p = s.split(/[\/\-\.]/);
  if (p.length !== 3) return new Date(s); // fallback
  const [a, b, c] = p.map(x => x.trim());
  if (a.length === 4) {
    // ISO: YYYY-MM-DD or YYYY/MM/DD
    return new Date(`${a}-${b.padStart(2,"0")}-${c.padStart(2,"0")}T12:00:00`);
  } else if (c.length === 4) {
    // Indian: DD/MM/YYYY or DD-MM-YYYY  (a=day, b=month, c=year)
    // IMPORTANT: never pass DD/MM/YYYY to new Date() directly — JS reads it as MM/DD
    return new Date(`${c}-${b.padStart(2,"0")}-${a.padStart(2,"0")}T12:00:00`);
  } else if (c.length === 2) {
    // Short year DD/MM/YY
    return new Date(`20${c}-${b.padStart(2,"0")}-${a.padStart(2,"0")}T12:00:00`);
  }
  return new Date(s);
};
// Format date as "04 May 2026" — manual format avoids browser locale differences
const MONTHS_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const fmtDate = v => {
  if (!v) return "";
  const d = parseIndianDate(v);
  if (isNaN(d)) return String(v);
  return `${String(d.getDate()).padStart(2,"0")} ${MONTHS_ABBR[d.getMonth()]} ${d.getFullYear()}`;
};
const fmtDateShort = v => {
  if (!v) return "";
  const d = parseIndianDate(v);
  if (isNaN(d)) return String(v);
  return `${String(d.getDate()).padStart(2,"0")} ${MONTHS_ABBR[d.getMonth()]} ${d.getFullYear()}`;
};

const aiLedger = desc => {
  if (!desc) return "Suspense Account";
  const d = desc.toLowerCase();
  if (/salary|sal\/|empl|wage|ctc|payroll/i.test(d)) return "Salary & Wages";
  if (/rent|lease|premise/i.test(d)) return "Rent Expenses";
  if (/bank charge|chgs|service fee|annual fee|atm|sms alert|locker/i.test(d)) return "Bank Charges";
  if (/igst/i.test(d)) return "IGST Payable";
  if (/cgst/i.test(d)) return "CGST Payable";
  if (/sgst/i.test(d)) return "SGST Payable";
  if (/gst|tax/i.test(d)) return "GST Payable";
  if (/tds/i.test(d)) return "TDS Payable";
  if (/tcs/i.test(d)) return "TCS Payable";
  if (/petrol|travel|tour|hotel|cab|ola|uber|flight|railway|irctc/i.test(d)) return "Travel Expenses";
  if (/office|stationary|canteen|pantry|supply|amazon|flipkart/i.test(d)) return "Office Expenses";
  if (/interest|int rcv|int cr/i.test(d)) return "Interest Received";
  if (/commission|comm rcv/i.test(d)) return "Commission Received";
  if (/capital|invest|fdr|fd open/i.test(d)) return "Capital Account";
  if (/drawing|withdraw|atm w\/d/i.test(d)) return "Drawings Account";
  if (/sales|sell|sold|invoice/i.test(d)) return "Sales Account";
  if (/purch|buy|vendor|supplier/i.test(d)) return "Purchase Account";
  if (/neft|rtgs|imps|upi|transfer to/i.test(d)) return "Sundry Creditors";
  if (/received|receipt|neft rcv|upi rcv/i.test(d)) return "Sundry Debtors";
  if (/electricity|bescom|msedcl|tpddl|power bill/i.test(d)) return "Electricity Charges";
  if (/telephone|airtel|jio|bsnl|broadband|internet/i.test(d)) return "Telephone Expenses";
  if (/audit|ca |chartered/i.test(d)) return "Audit Fees";
  if (/legal|advocate|court/i.test(d)) return "Legal & Professional Charges";
  if (/repair|maintenance|amc/i.test(d)) return "Repair & Maintenance";
  if (/advertisement|ads|google|meta|facebook|digital marketing/i.test(d)) return "Advertisement Expenses";
  if (/printing|stationery|xerox/i.test(d)) return "Printing & Stationery";
  if (/contra|transfer to self|own account/i.test(d)) return "HDFC Bank";
  return "Suspense Account";
};

// ── Custom Keyword Rule Engine ────────────────────────────────────
// Each rule: { id, keyword, matchType:"contains"|"startsWith"|"endsWith"|"regex",
//              ledger, voucherType:"Payment"|"Receipt"|"Contra"|"Journal"|"", enabled, priority }
const DEFAULT_RULES = [
  { id:"r01", keyword:"zomato|swiggy|dominos|dunzo|blinkit",      matchType:"regex",    ledger:"Office Expenses",              voucherType:"Payment",  enabled:true,  priority:1  },
  { id:"r02", keyword:"uber|ola|rapido|irctc|railway|flight",     matchType:"regex",    ledger:"Travel Expenses",              voucherType:"Payment",  enabled:true,  priority:2  },
  { id:"r03", keyword:"atm",                                       matchType:"contains", ledger:"Cash Account",                 voucherType:"Contra",   enabled:true,  priority:3  },
  { id:"r04", keyword:"sweep|auto sweep|fdr|term deposit",         matchType:"regex",    ledger:"Fixed Deposit Account",        voucherType:"Contra",   enabled:true,  priority:4  },
  { id:"r05", keyword:"salary|sal/|payroll|wages|empl",           matchType:"regex",    ledger:"Salary & Wages",               voucherType:"Payment",  enabled:true,  priority:5  },
  { id:"r06", keyword:"rent|lease|premise",                        matchType:"regex",    ledger:"Rent Expenses",                voucherType:"Payment",  enabled:true,  priority:6  },
  { id:"r07", keyword:"amazon|flipkart|meesho|myntra",             matchType:"regex",    ledger:"Office Expenses",              voucherType:"Payment",  enabled:true,  priority:7  },
  { id:"r08", keyword:"google|meta|facebook|instagram|ads",        matchType:"regex",    ledger:"Advertisement Expenses",       voucherType:"Payment",  enabled:true,  priority:8  },
  { id:"r09", keyword:"tds",                                       matchType:"contains", ledger:"TDS Payable",                  voucherType:"Journal",  enabled:true,  priority:9  },
  { id:"r10", keyword:"gst|igst|cgst|sgst",                       matchType:"regex",    ledger:"GST Payable",                  voucherType:"Journal",  enabled:true,  priority:10 },
  { id:"r11", keyword:"airtel|jio|bsnl|broadband|internet",       matchType:"regex",    ledger:"Telephone Expenses",           voucherType:"Payment",  enabled:true,  priority:11 },
  { id:"r12", keyword:"electricity|bescom|msedcl|tpddl|power",    matchType:"regex",    ledger:"Electricity Charges",          voucherType:"Payment",  enabled:true,  priority:12 },
  { id:"r13", keyword:"interest rcv|int cr|interest received",     matchType:"regex",    ledger:"Interest Received",            voucherType:"Receipt",  enabled:true,  priority:13 },
  { id:"r14", keyword:"bank charge|service charge|annual fee|sms alert", matchType:"regex", ledger:"Bank Charges",            voucherType:"Payment",  enabled:true,  priority:14 },
  { id:"r15", keyword:"neft|rtgs|imps transfer to",               matchType:"regex",    ledger:"Sundry Creditors",             voucherType:"Payment",  enabled:true,  priority:15 },
  { id:"r16", keyword:"sales|invoice|received from customer",      matchType:"regex",    ledger:"Sales Account",               voucherType:"Receipt",  enabled:true,  priority:16 },
  { id:"r17", keyword:"audit fee|ca fee|chartered accountant",     matchType:"regex",    ledger:"Audit Fees",                   voucherType:"Payment",  enabled:true,  priority:17 },
  { id:"r18", keyword:"petrol|fuel|hp pump|bp pump|iocl",         matchType:"regex",    ledger:"Travel Expenses",              voucherType:"Payment",  enabled:true,  priority:18 },
  { id:"r19", keyword:"drawing|personal withdrawal",               matchType:"regex",    ledger:"Drawings Account",             voucherType:"Journal",  enabled:true,  priority:19 },
  { id:"r20", keyword:"transfer to self|own account|hdfc to icici", matchType:"regex",  ledger:"HDFC Bank",                    voucherType:"Contra",   enabled:true,  priority:20 },
];

// Persist rules in localStorage across sessions
const loadRules = () => {
  try { const s = localStorage.getItem("b2t_rules"); return s ? JSON.parse(s) : DEFAULT_RULES; } catch { return DEFAULT_RULES; }
};
const saveRules = rules => { try { localStorage.setItem("b2t_rules", JSON.stringify(rules)); } catch {} };
const genRuleId = () => "r" + Math.random().toString(36).slice(2,8);

// Core matcher — tests a single rule against a narration string
const matchRule = (rule, narration) => {
  if (!rule.enabled) return false;
  const hay = (narration || "").toLowerCase();
  const needle = rule.keyword.toLowerCase();
  try {
    switch (rule.matchType) {
      case "regex":      return new RegExp(needle, "i").test(hay);
      case "startsWith": return hay.startsWith(needle);
      case "endsWith":   return hay.endsWith(needle);
      case "contains":
      default:           return hay.includes(needle);
    }
  } catch { return false; } // invalid regex — skip
};

// Apply all active rules (priority order) to a set of rows
// Returns new rows array with ledger/voucherType overwritten where a rule fires
const applyRulesToRows = (rows, rules) => {
  const sorted = [...rules].filter(r=>r.enabled).sort((a,b)=>a.priority-b.priority);
  return rows.map(row => {
    for (const rule of sorted) {
      if (matchRule(rule, row.narration)) {
        return {
          ...row,
          ledger:      rule.ledger      || row.ledger,
          voucherType: rule.voucherType || row.voucherType,
          ruleMatched: rule.keyword,
        };
      }
    }
    return row;
  });
};
const escXml = s => String(s||"")
  .replace(/&/g,"&amp;")
  .replace(/</g,"&lt;")
  .replace(/>/g,"&gt;")
  .replace(/"/g,"&quot;")
  .replace(/'/g,"&apos;");

// ── Voucher type logic — with credit/debit validation ────────────
const voucherType = (debit, credit, ledger) => {
  const lg = (ledger || "").toLowerCase();
  // Contra: inter-bank or bank-cash transfers
  if ((lg.includes("bank") || lg.includes("cash") || lg.includes("contra")) &&
      !lg.includes("charges") && !lg.includes("sundry") && !lg.includes("salary")) return "Contra";
  if (credit && !debit) return "Receipt";
  if (debit && !credit) return "Payment";
  return "Journal";
};

// ── Voucher validation: flag credit+Payment or debit+Receipt ─────
const validateVoucherType = (row) => {
  const vt = row.voucherType || voucherType(row.debit, row.credit, row.ledger);
  if (row.credit && !row.debit && vt === "Payment") return { valid:false, issue:"Credit entry mapped to Payment voucher", fix:"Receipt" };
  if (row.debit && !row.credit && vt === "Receipt") return { valid:false, issue:"Debit entry mapped to Receipt voucher", fix:"Payment" };
  return { valid:true };
};

// ── Running balance checker — flags rows where stated balance drifts ──
const checkRunningBalance = (rows) => {
  let running = null;
  return rows.map(r => {
    const stated = parseFloat(String(r.balance||"").replace(/,/g,""));
    const debit  = parseFloat(String(r.debit||"0").replace(/,/g,""));
    const credit = parseFloat(String(r.credit||"0").replace(/,/g,""));
    if (isNaN(stated)) return { ...r, balanceMismatch: false };
    if (running === null) { running = stated; return { ...r, balanceMismatch: false }; }
    const expected = parseFloat((running - debit + credit).toFixed(2));
    const mismatch = Math.abs(expected - stated) > 0.02; // 2 paise tolerance
    if (!mismatch) running = stated;
    return { ...r, balanceMismatch: mismatch, expectedBalance: expected };
  });
};

// ── Duplicate detection (hashing map) ────────────────────────────
const detectDuplicates = rows => {
  const seen = new Map();
  return rows.map(r => {
    const key = `${String(r.date).slice(0,10)}|${String(r.debit||"").replace(/,/g,"")}|${String(r.credit||"").replace(/,/g,"")}|${String(r.narration).slice(0,60).toLowerCase().replace(/\s+/g," ").trim()}`;
    if (seen.has(key)) return { ...r, isDuplicate: true, duplicateOf: seen.get(key) };
    seen.set(key, r.id);
    return { ...r, isDuplicate: false };
  });
};

// ── Compliance Score calculator ──────────────────────────────────
const calcComplianceScore = (rows) => {
  if (!rows.length) return 100;
  const total = rows.length;
  const suspense   = rows.filter(r => r.ledger === "Suspense Account").length;
  const dups       = rows.filter(r => r.isDuplicate).length;
  const mismatches = rows.filter(r => r.balanceMismatch).length;
  const badVoucher = rows.filter(r => !validateVoucherType(r).valid).length;
  const deductions = (suspense/total)*35 + (dups/total)*25 + (mismatches/total)*20 + (badVoucher/total)*20;
  return Math.max(0, Math.round(100 - deductions));
};

// ── Tally-Prime compliant XML generator (audit-fixed) ────────────
const toTallyXML = (rows, company, fy = "2024-25") => {
  const tallyDate = raw => {
    const s = String(raw || "");
    const months = {jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
    const m1 = s.match(/(\d{1,2})[\s\-\/]([a-z]{3})[\s\-\/](\d{4})/i);
    if (m1) return m1[3]+(months[m1[2].toLowerCase()]||"01")+String(m1[1]).padStart(2,"0");
    const m2 = s.match(/^(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})/);
    if (m2) return m2[1]+String(m2[2]).padStart(2,"0")+String(m2[3]).padStart(2,"0");
    const m3 = s.match(/^(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{4})/);
    if (m3) return m3[3]+String(m3[2]).padStart(2,"0")+String(m3[1]).padStart(2,"0");
    return s.replace(/[^0-9]/g,"").slice(0,8);
  };

  const stableGuid = r => {
    const raw = String(r.date||"")+"|"+String(r.debit||r.credit||"")+"|"+String(r.narration||"").slice(0,40);
    let h = 5381;
    for (let i=0;i<raw.length;i++){h=((h<<5)+h)+raw.charCodeAt(i);h=h&h;}
    return "B2T-"+Math.abs(h).toString(36).toUpperCase().padStart(8,"0");
  };

  const bankLedger = escXml(company.bankLedger || company.name || "Bank Account");

  const vouchers = rows.filter(r => !r.isDuplicate || r.forceImport).map((r, idx) => {
    const amt     = parseFloat(String(r.debit||r.credit||"0").replace(/,/g,""));
    const isDebit = !!r.debit && !r.credit;  // true = money left bank (Payment/Contra)
    const vtype   = r.voucherType || voucherType(r.debit, r.credit, r.ledger);
    const ledger  = escXml(r.ledger || "Suspense A/c");
    const dt      = tallyDate(r.date);
    const narr    = escXml((r.narration||"")+(r.ref?` [Ref: ${r.ref}]`:"")+" [Imported via Bank2Tally]");

    // Tally Prime double-entry convention:
    // For a PAYMENT (bank debit): bank entry amount = negative, expense entry = positive
    // For a RECEIPT (bank credit): bank entry amount = positive, income entry = negative
    const bankAmt   = isDebit ? -amt : amt;
    const ledgerAmt = isDebit ?  amt : -amt;
    const bankPos   = isDebit ? "No"  : "Yes";  // ISDEEMEDPOSITIVE for bank
    const ledgPos   = isDebit ? "Yes" : "No";   // ISDEEMEDPOSITIVE for ledger

    return `        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER REMOTEID="${stableGuid(r)}" VCHTYPE="${escXml(vtype)}" ACTION="Create" OBJVIEW="Accounting VoucherView">
            <DATE>${dt}</DATE>
            <EFFECTIVEDATE>${dt}</EFFECTIVEDATE>
            <GUID>${stableGuid(r)}-${idx}</GUID>
            <VOUCHERTYPENAME>${escXml(vtype)}</VOUCHERTYPENAME>
            <PARTYLEDGERNAME>${ledger}</PARTYLEDGERNAME>
            <NARRATION>${narr}</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${bankLedger}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>${bankPos}</ISDEEMEDPOSITIVE>
              <AMOUNT>${bankAmt.toFixed(2)}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${ledger}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>${ledgPos}</ISDEEMEDPOSITIVE>
              <AMOUNT>${ledgerAmt.toFixed(2)}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>`;
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>${escXml(company.name||"")}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
${vouchers}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
};

// ── PDF OCR Pipeline ─────────────────────────────────────────────
// Loads pdfjs + Tesseract from CDN, extracts text (or OCRs scanned pages),
// then heuristically parses the result into { headers, rows }.

async function loadScript(src) {
  if (document.querySelector(`script[src="${src}"]`)) {
    return new Promise(res => {
      const poll = setInterval(() => {
        if (document.querySelector(`script[src="${src}"]`)?.dataset?.loaded) { clearInterval(poll); res(); }
      }, 80);
    });
  }
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => { s.dataset.loaded = "1"; res(); };
    s.onerror = () => rej(new Error(`Failed to load ${src}`));
    document.head.appendChild(s);
  });
}

// ══════════════════════════════════════════════════════════════════
//  BANK STATEMENT PDF PARSER — MULTI-ENGINE ARCHITECTURE
//  Each known bank gets a precise X-position-aware parser.
//  A fingerprint detector routes to the right parser automatically.
//  Unknown banks fall back to the universal text-based engine.
// ══════════════════════════════════════════════════════════════════

// ── Shared helpers ────────────────────────────────────────────────
const RX_DATE   = /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$|^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$|\b\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{2,4}\b/i;
const RX_AMOUNT = /^-?[\d,]+\.?\d{0,2}$|^-?[\d,]+\.\d{2}\s*(CR|DR|cr|dr)?$/;
const RX_FOOTER = /^(total|grand total|closing|opening balance|page\s*\d|statement|account summary|sr\.?\s*no\.?$|date\s*description|transactions\s*for)/i;
const RX_HDR    = /date|narr|desc|debit|credit|withdraw|deposit|balance|particulars|amount|remarks|details|tran|txn|cheque|chq|ref|value\s*date|posted/i;

function isDateStr(s)   { return RX_DATE.test(String(s).trim()); }
function isAmountStr(s) { const c = String(s).trim().replace(/,/g,""); return RX_AMOUNT.test(c) && !isNaN(parseFloat(c)); }

// ── Load pdfjs once, share across all parsers ─────────────────────
async function loadPdfJs() {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
  const lib = window["pdfjs-dist/build/pdf"];
  lib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  return lib;
}

// Extract raw word items with x/y positions from all pages
// Y is converted to top-down (0 = page top) for consistent processing
// Uses disableCombineTextItems:false so pdfjs merges adjacent chars into words
async function extractWordItems(buf) {
  const pdfjsLib = await loadPdfJs();
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp   = page.getViewport({ scale: 1 });
    const con  = await page.getTextContent({
      normalizeWhitespace:     true,
      disableCombineTextItems: false,   // let pdfjs combine adjacent chars
    });

    // pdfjs still sometimes emits single chars; we merge them into words here.
    // Strategy: group items by their baseline Y (transform[5]), then within each
    // line merge items whose X gap is < (fontSize * 0.4) — i.e. no visible space.
    const raw = con.items
      .filter(it => it.str && it.str.trim())
      .map(it => ({
        x:    Math.round(it.transform[4] * 10) / 10,
        y:    Math.round((vp.height - it.transform[5]) * 10) / 10,
        baseY: Math.round(it.transform[5] * 10) / 10,  // raw baseline for merging
        w:    it.width || 0,
        fs:   Math.abs(it.transform[0]) || 10,         // font size from transform[0]
        str:  it.str,
      }));

    // 1. Group by baseline Y with 2pt tolerance
    const byLine = {};
    raw.forEach(it => {
      const key = Math.round(it.baseY / 2) * 2;
      if (!byLine[key]) byLine[key] = [];
      byLine[key].push(it);
    });

    // 2. Within each line, merge chars whose gap < 40% of font size into words
    const items = [];
    Object.values(byLine).forEach(line => {
      line.sort((a, b) => a.x - b.x);
      const words = [];
      line.forEach(it => {
        const prev = words[words.length - 1];
        const gap  = prev ? it.x - (prev.x + prev.w) : Infinity;
        const spaceThreshold = (it.fs || 10) * 0.35;
        if (prev && gap < spaceThreshold && gap > -2) {
          // merge into previous word
          prev.str += it.str;
          prev.w    = (it.x + it.w) - prev.x;
        } else {
          words.push({ x: it.x, y: it.y, w: it.w, str: it.str });
        }
      });
      items.push(...words);
    });

    pages.push({ pageNum: p, items, width: vp.width, height: vp.height });
  }
  return pages;
}

// ── Bank fingerprint detector ─────────────────────────────────────
function detectBank(pages) {
  // Build two samples:
  // 1. joined with spaces (works for merged words)
  // 2. joined without spaces (works when pdfjs emits individual chars: "I C I C I" → "ICICI")
  const rawItems = pages.flatMap(p => p.items).slice(0, 300).map(it => it.str);
  const sample   = rawItems.join(" ").toLowerCase();
  const compact  = rawItems.join("").toLowerCase();  // no spaces — catches "I C I C I B a n k"
  const test     = s => sample.includes(s) || compact.includes(s);

  if (/\bdbs\b|development\s*bank\s*of\s*singapore|dbsssgsg|posb\s*bank|fast\s*payment.*26-\d{3}|autosave\s*transfer|ebusiness\s*lite|856210064336|0856IT\d+|dbs\s*bank.*india|dbs.*ebusiness/i.test(sample)) return "dbs";
  if (test("icicibank") || test("icici bank") || test("icic0") || test("icicibank.com") || /detailed\s*statement|cr\/dr.*transaction.*amount/i.test(sample) || (/account\s*statement/i.test(sample) && /icic/i.test(compact))) return "icici";
  if (test("state bank of india") || test("statebankofindia") || /\bsbi\b|sbchq|sbin\d{7}/i.test(sample)) return "sbi";
  if (test("hdfc bank") || test("hdfcbank") || /withdrawal\s*amt\.|deposit\s*amt\./i.test(sample)) return "hdfc";
  if (test("axis bank") || test("axisbank") || /tran\s*date.*particulars/i.test(sample))       return "axis";
  if (test("kotak mahindra") || test("kotakbank") || /kotak/i.test(compact))                    return "kotak";
  if (/punjab\s*national\s*bank|\bpnb\b/i.test(sample))                                        return "pnb";
  if (/bank\s*of\s*baroda|\bbaroda\b/i.test(sample))                                           return "bob";
  if (/yes\s*bank/i.test(sample))                                                               return "yes";
  if (/idfc\s*(first|bank)/i.test(sample))                                                     return "idfc";
  if (test("canara bank") || test("canarabank"))                                                 return "canara";
  if (/union\s*bank/i.test(sample))                                                             return "union";
  if (/bank\s*of\s*india/i.test(sample))                                                       return "boi";
  if (/federal\s*bank/i.test(sample))                                                           return "federal";
  if (/indusind/i.test(sample))                                                                 return "indus";
  if (/rbl\s*bank/i.test(sample))                                                               return "rbl";
  if (/andhra\s*bank|andhra.*corp/i.test(sample))                                               return "andhra";
  if (/central\s*bank\s*of\s*india/i.test(sample))                                             return "centralbank";
  if (/bank\s*of\s*baroda|\bbaroda\b/i.test(sample))                                          return "bob";
  if (/bank\s*of\s*india/i.test(sample))                                                       return "boi";
  if (/post\s*date.*value\s*date|value\s*date.*post\s*date/i.test(sample))                   return "txnhistory";
  return "generic";
}

// ══════════════════════════════════════════════════════════════════
//  BANK-SPECIFIC PARSERS  (work directly on word-position data)
// ══════════════════════════════════════════════════════════════════

// Helper: group word items into physical lines by Y proximity
function groupIntoLines(items, pageTolerance = 3) {
  if (!items.length) return [];
  // Adaptive tolerance from median Y gap
  const ys = [...new Set(items.map(it => it.y))].sort((a,b)=>a-b);
  let tol = pageTolerance;
  if (ys.length > 3) {
    const gaps = [];
    for (let i = 1; i < Math.min(ys.length, 30); i++) { const g = ys[i]-ys[i-1]; if (g>0.5&&g<20) gaps.push(g); }
    if (gaps.length) { gaps.sort((a,b)=>a-b); tol = Math.max(1.5, gaps[Math.floor(gaps.length/2)] * 0.6); }
  }
  const lineMap = {};
  items.forEach(it => {
    const key = Math.round(it.y / tol) * tol;
    if (!lineMap[key]) lineMap[key] = [];
    lineMap[key].push(it);
  });
  return Object.keys(lineMap)
    .sort((a,b) => Number(a)-Number(b))
    .map(k => ({ y: Number(k), items: lineMap[k].sort((a,b)=>a.x-b.x) }));
}

// Helper: convert lines into tab-separated text using X-band snapping
function linesToText(lines, pageWidth) {
  // Detect X column bands from lines with 2+ items
  const allX = lines.filter(l=>l.items.length>=2).flatMap(l=>l.items.map(it=>it.x)).sort((a,b)=>a-b);
  const thresh = Math.max(10, pageWidth * 0.02);
  const bands = [];
  if (allX.length) { let bx=allX[0]; for(let i=1;i<allX.length;i++){if(allX[i]-allX[i-1]>thresh){bands.push(bx);bx=allX[i];}} bands.push(allX[allX.length-1]); }
  const snap = x => { for(let i=0;i<bands.length-1;i++){if(x<(bands[i]+bands[i+1])/2)return i;} return Math.max(0,bands.length-1); };

  return lines.map(line => {
    // Merge micro-fragments first
    const merged = [];
    line.items.forEach(it => {
      const last = merged[merged.length-1];
      if (last && (it.x-(last.x+last.w))<3) { last.str+=it.str; last.w=it.x+it.w-last.x; }
      else merged.push({...it});
    });
    if (bands.length <= 2) return merged.map(it=>it.str.trim()).join("  ");
    const slots = new Array(bands.length).fill("");
    merged.forEach(it => { const c=snap(it.x); slots[c]=slots[c]?slots[c]+" "+it.str.trim():it.str.trim(); });
    let s = slots.join("\t"); while(s.endsWith("\t"))s=s.slice(0,-1); return s;
  }).filter(l=>l.trim());
}

// ── ICICI Bank Statement ─────────────────────────────────────────
// FORMAT A — "Detailed Statement" (900pt wide, OpTransactionHistory):
//   Layout per transaction (3 Y levels, ~8pt apart):
//     Y+0:  valDate(108) | postedDate(184) | desc(417) | prevBalance(829)  ← balance belongs to PREV row
//     Y+7:  sno(15) | txnId(43) | cheque(369) | [desc cont] | crdr(704) | amount(780)
//     Y+14: [desc continuation lines at x=417]
//   The balance shown at Y+0 is the RUNNING balance BEFORE this transaction.
//   So we carry it forward: balance[n] = prevBalance line of row[n+1].
//
// FORMAT B — "Account Statement" A4 (595pt wide, 2024+ Current Account):
//   Layout per transaction (first line + continuation lines):
//     Y+0:  sno(50) | txnId(94) | date(161) | desc(296) | debit(363) | credit(430) | balance(497)
//     Y+12: date-year(161) | desc-cont(296)
//     Y+24: desc-cont(296)  ...
function parseICICIWords(pages) {
  if (!pages.length) return null;
  const pageWidth = pages[0].width || 595;

  // ═══════════════════════════════════════════════════════════════
  // FORMAT A — Wide "Detailed Statement" (900pt)
  // ═══════════════════════════════════════════════════════════════
  if (pageWidth > 700) {
    // Exact X bands from measured PDF
    const col = x =>
      x < 36   ? "sno"     :
      x < 108  ? "txnid"   :
      x < 183  ? "valdate" :
      x < 365  ? "posted"  :
      x < 415  ? "cheque"  :
      x < 695  ? "desc"    :
      x < 760  ? "crdr"    :
      x < 825  ? "amount"  : "balance";

    // Group all items into physical lines (2pt tolerance)
    const allItems = pages.flatMap(p => p.items);
    const physLines = groupIntoLines(allItems, 2);

    // Each transaction block spans ~3 Y-lines:
    //   line A (date/desc/prevBalance), line B (sno/txnid/crdr/amount), line C+ (desc cont)
    // We identify transaction starts by: line has valdate AND NO sno
    // Then the NEXT line with sno/txnid gives us the ID and amount.
    // Balance at line A = balance AFTER the previous transaction, i.e. opening for this one.
    // We instead use the balance from line A of the NEXT transaction as closing balance.

    const rows = []; // {valdate, txnid, desc[], crdr, amount, balance(running after)}
    let pending = null; // transaction being built
    let pendingBalance = ""; // balance line seen BEFORE current transaction

    const flushPending = (closingBal) => {
      if (!pending) return;
      if (!pending.valdate || !pending.amount) { pending = null; return; }
      const debit  = pending.crdr === "DR" ? pending.amount : "";
      const credit = pending.crdr === "CR" ? pending.amount : "";
      rows.push([
        pending.valdate,
        pending.txnid,
        pending.desc.filter(Boolean).join(" ").trim(),
        debit,
        credit,
        closingBal || pending.balance
      ]);
      pending = null;
    };

    physLines.forEach(line => {
      const txt = line.items.map(i => i.str).join(" ");
      // Skip header/footer
      if (/detailed statement|transactions list|no\.\s*transact|value date|txn posted|generated on|page \d|legends|bbps|bctt|^bil |^bpay/i.test(txt)) return;

      const slots = { sno:[], txnid:[], valdate:[], posted:[], cheque:[], desc:[], crdr:[], amount:[], balance:[] };
      line.items.forEach(it => { const c = col(it.x); slots[c].push(it.str.trim()); });

      const valdate = slots.valdate.find(isDateStr) || "";
      const txnid   = slots.txnid.join("").trim();
      const sno     = slots.sno.join("").trim();
      const desc    = slots.desc.filter(s => s && s !== "-").join(" ").trim();
      const crdr    = slots.crdr.join("").trim().toUpperCase();
      const amount  = slots.amount.find(isAmountStr) || "";
      const balance = slots.balance.find(isAmountStr) || "";

      // A line with balance but no date/txnid = running balance header line
      // This is the CLOSING balance of the previous transaction
      if (balance && !valdate && !txnid && !sno) {
        if (pending) pending.balance = balance; // assign as closing to pending
        pendingBalance = balance;
        return;
      }

      // Line with valdate = start of new transaction block
      if (valdate) {
        // The balance on THIS line is actually the closing balance of the PREVIOUS tx
        flushPending(balance || pendingBalance);
        pending = { valdate, txnid, desc: desc ? [desc] : [], crdr, amount, balance: "" };
        pendingBalance = "";
        return;
      }

      // Line with sno/txnid (the identification line, 7pt below date line)
      if (/^S\d{5,}/i.test(txnid) || /^\d{1,4}$/.test(sno)) {
        if (pending) {
          if (!pending.txnid && txnid) pending.txnid = txnid;
          if (!pending.crdr && crdr)   pending.crdr   = crdr;
          if (!pending.amount && amount) pending.amount = amount;
          if (desc) pending.desc.push(desc);
        }
        return;
      }

      // Continuation desc line
      if (pending && desc) pending.desc.push(desc);
      if (pending && !pending.crdr && crdr)     pending.crdr   = crdr;
      if (pending && !pending.amount && amount)  pending.amount = amount;
    });

    flushPending("");

    const HDR = ["Value Date","Transaction ID","Description","Withdrawal (Dr)","Deposit (Cr)","Available Balance"];
    return rows.length ? { headers: HDR, rows, _bankHint: "icici" } : null;
  }

  // ═══════════════════════════════════════════════════════════════
  // FORMAT B — New A4 "Account Statement" (595pt, 2024+)
  // Uses RAW TEXT approach — far more reliable than X/Y coordinate
  // parsing for this PDF which emits individual chars via pdfjs.
  //
  // Raw text per line (from pdfjs joined):
  //   "1 S1070358 04-May- UPI/109740 1050.00 30980.30"   ← row start
  //   "2026 316919/milk"                                  ← year + desc cont
  //   "payment/q8"                                        ← desc cont only
  // ═══════════════════════════════════════════════════════════════
  const MON_B = {jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",
                 jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};

  // Row-start: "1 S1070358 04-May- UPI/109740 1050.00 30980.30"
  //   group 1 = sno, 2 = txnid, 3 = "DD-Mon-", 4 = desc_start (optional),
  //   5 = amount, 6 = balance
  const ROW_RE = /^(\d{1,3})\s+(S\d{5,})\s+(\d{1,2}-[A-Za-z]{3}-)\s*(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$/;
  // Continuation with year: "2026 316919/milk"
  const YEAR_RE = /^(\d{4})\s+(.*)$/;
  // Skip patterns
  const SKIP_B = /generated\s+on|page\s+\d+\s+of\s+\d+|legends\s+used|^s\.no\s+transaction|^account\s+(name|number|type|currency)|^ifsc|^customer\s+id|^available\s+balance|^total\s+effective|^\*this\s+is|^\d+\.\s+[A-Z]{2,}\s*-/i;

  const normDateB = (part1, year) => {
    const m = part1.match(/^(\d{1,2})-([A-Za-z]{3})-?$/i);
    if (!m) return part1 + year;
    const mo = MON_B[m[2].toLowerCase()] || m[2];
    return `${m[1].padStart(2,"0")}/${mo}/${year}`;
  };

  const rowsB = [];
  let curB = null;

  const flushB = () => {
    if (!curB) return;
    if (!curB.year) { curB = null; return; } // no year = incomplete, discard
    const date = normDateB(curB.datePart, curB.year);
    const desc = curB.desc.join(" ").trim();
    if (!date || !curB.amt) { curB = null; return; }
    // Determine debit/credit by comparing prev balance
    const amt  = parseFloat(curB.amt.replace(/,/g,""));
    const bal  = parseFloat(curB.balance.replace(/,/g,""));
    const prev = parseFloat((curB.prevBalance||"0").replace(/,/g,""));
    let debit = "", credit = "";
    if (prev > 0) {
      if (Math.abs((prev - amt) - bal) < 0.02) debit  = curB.amt;
      else if (Math.abs((prev + amt) - bal) < 0.02) credit = curB.amt;
      else debit = curB.amt; // fallback — balance went down
    } else {
      debit = curB.amt; // first row, no prev — default debit
    }
    rowsB.push([date, curB.txnid, desc, debit, credit, curB.balance]);
    curB = null;
  };

  // Get raw text lines from all pages (same as what pdfjs produces)
  const allLines = pages.flatMap(({ items }) => {
    // Group items into lines by Y (6pt tolerance), join by space
    const lm = {};
    items.forEach(it => {
      const k = Math.round(it.y / 6) * 6;
      if (!lm[k]) lm[k] = [];
      lm[k].push(it);
    });
    return Object.keys(lm).sort((a,b)=>+a-+b)
      .map(k => lm[k].sort((a,b)=>a.x-b.x).map(i=>i.str).join(" ").trim())
      .filter(Boolean);
  });

  let prevBal = "";
  allLines.forEach(line => {
    if (SKIP_B.test(line)) { flushB(); return; }

    const mRow = ROW_RE.exec(line);
    if (mRow) {
      if (curB) {
        prevBal = curB.balance;
        flushB();
      }
      curB = {
        txnid:    mRow[2],
        datePart: mRow[3],
        year:     "",          // filled by next continuation line
        desc:     mRow[4] ? [mRow[4]] : [],
        amt:      mRow[5],
        balance:  mRow[6],
        prevBalance: prevBal,
      };
      return;
    }

    if (curB) {
      const mYear = YEAR_RE.exec(line);
      if (mYear && !curB.year) {
        curB.year = mYear[1];
        if (mYear[2].trim()) curB.desc.push(mYear[2].trim());
      } else if (line.trim()) {
        curB.desc.push(line.trim());
      }
    }
  });
  if (curB) { prevBal = curB.balance; flushB(); }
  const HDR_B = ["Date","Transaction ID","Description","Withdrawal (Dr)","Deposit (Cr)","Available Balance"];
  return rowsB.length ? { headers: HDR_B, rows: rowsB, _bankHint: "icici" } : null;
}

// ── SBI Account Statement ─────────────────────────────────────────
// Clean 7-col table preceded by a metadata block (Label : Value pairs).
// Header: Txn Date | Value Date | Description | Ref No./Cheque No. | Debit | Credit | Balance
// Description wraps across 2-3 lines, Ref No. column header also wraps.
function parseSBIWords(pages) {
  const allLines = pages.flatMap(({ items, width }) => {
    const physLines = groupIntoLines(items, 3);
    return linesToText(physLines, width);
  });

  const SEP = /\t|\s{2,}/;
  const isKV = line => {
    // KV lines: "Account Name    : Mr. Atul Kumar Verma"
    const colons = (line.match(/:/g)||[]).length;
    const parts  = line.split(":").map(s=>s.trim());
    return colons === 1 && /^[A-Za-z\s\.%\(\)\/]+$/.test(parts[0]) && parts[0].split(" ").length <= 6;
  };

  // Find the real header row (score by bank keyword hits, skip KV lines)
  let hdrIdx = -1, best = 0;
  allLines.forEach((line, i) => {
    if (RX_FOOTER.test(line) || isKV(line)) return;
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (cols.length < 3) return;
    const score = cols.filter(c=>RX_HDR.test(c)).length + (cols.length >= 5 ? 2 : 0);
    if (score > best && cols.filter(c=>RX_HDR.test(c)).length >= 3) { best = score; hdrIdx = i; }
  });
  if (hdrIdx === -1) return null;

  // Merge wrapped header line (SBI wraps "Ref No./Cheque" / "No." onto next line)
  let headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const nxt = (allLines[hdrIdx+1]||"").split(SEP).map(c=>c.trim()).filter(Boolean);
  if (nxt.length && nxt.length <= headers.length && !nxt.some(isDateStr) && !nxt.some(isAmountStr)
      && nxt.every(w => /^[A-Za-z\.\(\)\/\-]+$/.test(w))) {
    headers = headers.map((h,i) => nxt[i] ? h+" "+nxt[i] : h);
    hdrIdx++;
  }

  // Sliding-window row stitcher (description wraps across lines)
  const SNO    = /^\d{1,6}$/;
  const NEWROW = cols => isDateStr(cols[0]) || (cols.length >= 2 && isDateStr(cols[1]));
  const txnRows = []; let cur = null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line) || isKV(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if (isAmountStr(tok)) {
          let pl=false;
          for(let k=cur.length-1;k>=Math.max(0,cur.length-4);k--){if(!cur[k]){cur[k]=tok;pl=true;break;}}
          if(!pl)cur.push(tok);
        } else {
          const ni=cur.findIndex((c,i)=>i>0&&!isDateStr(c)&&!isAmountStr(c)&&!SNO.test(c));
          if(ni!==-1)cur[ni]=(cur[ni]+" "+tok).trim();else cur.push(tok);
        }
      });
    } else { if(cur)txnRows.push(cur); cur=[...cols]; }
  });
  if (cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});

  return rows.length ? { headers, rows, _bankHint:"sbi" } : null;
}

// ── HDFC Bank Statement ───────────────────────────────────────────
// 6-col: Date | Narration | Chq/Ref | Withdrawal Amt. | Deposit Amt. | Closing Balance
function parseHDFCWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = -1, best = 0;
  allLines.forEach((line,i) => {
    if (i > 40 || RX_FOOTER.test(line)) return;
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    const score = cols.filter(c=>RX_HDR.test(c)).length;
    if (score >= 3 && score > best) { best=score; hdrIdx=i; }
  });
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const NEWROW  = cols => isDateStr(cols[0]);
  const txnRows = []; let cur = null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if (isAmountStr(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if (cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"hdfc" } : null;
}

// ── Axis Bank Statement ───────────────────────────────────────────
// 6-col: Tran Date | PARTICULARS | CHQNO | DR | CR | BAL
function parseAxisWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,4), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<40 && /PARTICULARS|Tran\s*Date/i.test(line) && line.split(SEP).filter(c=>RX_HDR.test(c)).length >= 3);
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const NEWROW  = cols => isDateStr(cols[0]);
  const txnRows = []; let cur = null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if (isAmountStr(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if (cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"axis" } : null;
}

// ── DBS Bank — Singapore & India formats ─────────────────────────
// Both use: Trans. Date | Value Date | Transaction Details (multi-line) | Debits | Credits | Running Balance
//
// DBS Singapore specifics:
//   - Product: Current Account / Savings
//   - Refs: EBGPP60408436435..., IBG BEXP-..., RTF EBACT...
//   - "26-067" batch codes, "FAST PAYMENT", "AUTOSAVE TRANSFER FEE"
//   - Currency line: "SGD 755" at bottom of description block
//
// DBS India specifics:
//   - Product: EBUSINESS LITE / Current Account
//   - Account: 856210064336-INR format
//   - Refs: 0856IT0509568, IMPS-609112023254, EBHVT60403143451
//   - Transaction types: TRANSFER, IMPS, NEFT, RTGS, REMITTANCE
//   - Currency line: "INR 2000000" or "USD 25000 AT INR92.5425000"
//   - Continuation-only rows: date + value date + "INR XXXXX" + amount + balance (no TRANSFER word)
//   - Multiple pages with "Account Details / From / To" header repeated
//
// Strategy: use PDF positional data (x-coordinates) to classify each token
// as Date | Description | Debit | Credit | Balance by column position.
function parseDBSWords(pages) {
  // ── 1. Detect variant ───────────────────────────────────────────
  const allText = pages.flatMap(p => p.items.map(i => i.str || "")).join(" ");
  const isIndia = /ebusiness\s*lite|856210064336|0856IT\d+|INR\s*\d{4,}/i.test(allText);

  // ── 2. Noise filter (common to both variants) ───────────────────
  const isNoise = line => /deposit\s*insurance|supplementary\s*retirement|sdic|scheme\s*member|non-bank\s*depositor|aggregate\s*per\s*depositor|end\s*of\s*report|printed\s*(on|by)|page\s*\d+\s*of\s*\d+|transactions\s*performed\s*on\s*a\s*non.working|date\s*requested\s*is\s*a\s*non\s*business|select\s*the\s*next\s*business\s*day|non\s*working\s*day\s*will\s*be\s*posted|foreign\s*currency\s*deposits|dual\s*currency|structured\s*deposits|investment\s*products\s*are\s*not\s*insured|opening\s*balance\s*:|ledger\s*balance\s*:|available\s*balance\s*:|earmark|overdraft\s*limit|account\s*number\s*:|account\s*name\s*:|product\s*type\s*:|total\s*debit\s*count|total\s*credit\s*count|total\s*debit\s*amount|total\s*credit\s*amount|effective\s*available|from\s+to\s*$|account\s*details/i.test(line);

  const isTableHeader = line => /trans.*date.*value.*date|value\s*date.*transaction.*details|debits.*credits.*running/i.test(line);

  const DATE_RE = /^\d{2}-[A-Za-z]{3}-\d{4}$/;
  const isDate = s => DATE_RE.test(s.trim());

  // ── 3. Build all lines across all pages ─────────────────────────
  const allLines = pages.flatMap(({ items, width }) =>
    linesToText(groupIntoLines(items, 3), width)
  );

  // Find header row — skip everything before it (account info, dates, etc.)
  let dataStart = 0;
  for (let i = 0; i < Math.min(allLines.length, 40); i++) {
    if (isTableHeader(allLines[i])) { dataStart = i + 1; break; }
  }

  // ── 4. Group lines into transaction blocks ──────────────────────
  // A new transaction always starts with DD-Mon-YYYY in position 0.
  // The value date (same or different) follows immediately.
  // Subsequent lines until the next date are continuation of the same txn.
  const txns = [];
  let cur = null;

  const cleanLines = allLines.slice(dataStart).filter(l => !isNoise(l) && l.trim().length > 0);

  for (const line of cleanLines) {
    // Re-detect header on each page (DBS India repeats "Account Details / From / To" per page)
    if (isTableHeader(line)) { continue; }

    const parts = line.split(/\t|\s{2,}/).map(s => s.trim()).filter(Boolean);
    if (!parts.length) continue;

    const p0IsDate = isDate(parts[0]);
    const p1IsDate = parts.length > 1 && isDate(parts[1]);

    if (p0IsDate) {
      // Save previous transaction
      if (cur) txns.push(cur);

      cur = {
        transDate: parts[0],
        valueDate: p1IsDate ? parts[1] : parts[0],
        descParts: [],
        _amts: [],
      };

      // Process rest of first line
      const rest = parts.slice(p1IsDate ? 2 : 1);
      for (const tok of rest) {
        if (isAmountStr(tok))  cur._amts.push(tok);
        else if (!isDate(tok)) cur.descParts.push(tok);
      }
    } else if (cur) {
      // Continuation line
      for (const tok of parts) {
        if (isAmountStr(tok))  cur._amts.push(tok);
        else if (!isDate(tok)) cur.descParts.push(tok);
      }
    }
  }
  if (cur) txns.push(cur);
  if (!txns.length) return null;

  // ── 5. Resolve amounts into debit / credit / balance ───────────
  // DBS layout: amounts appear in THREE columns (Debits | Credits | Running Balance)
  // Within a transaction block all amounts appear in order left→right, line by line.
  // The LAST amount is always the Running Balance.
  // The SECOND-TO-LAST (if present) is the transaction amount.
  // Direction (debit vs credit) is determined by running balance delta.
  //
  // Special case for DBS India: some transactions have 3 amounts (debit + credit slot empty + balance)
  // or only 2 (txnAmt + balance). We use delta from previous balance as ground truth.

  const rows = [];
  let prevBal = null;

  for (const txn of txns) {
    const amts = txn._amts
      .map(a => parseFloat(String(a).replace(/,/g, "")))
      .filter(n => !isNaN(n) && n > 0);

    if (amts.length === 0) continue;

    const bal     = amts[amts.length - 1];
    const txnAmt  = amts.length >= 2 ? amts[amts.length - 2] : null;

    let debit = "", credit = "";

    if (txnAmt !== null && prevBal !== null) {
      const delta = parseFloat((bal - prevBal).toFixed(2));
      const tol   = Math.max(0.10, txnAmt * 0.0001); // 0.01% tolerance for FX rounding
      if (Math.abs(delta + txnAmt) <= tol)      debit  = txnAmt.toFixed(2); // balance fell
      else if (Math.abs(delta - txnAmt) <= tol) credit = txnAmt.toFixed(2); // balance rose
      else {
        // Delta doesn't match neatly — use sign of delta
        if (delta < 0) debit  = txnAmt.toFixed(2);
        else           credit = txnAmt.toFixed(2);
      }
    } else if (txnAmt !== null) {
      // No previous balance — use description keywords
      const desc = txn.descParts.join(" ").toUpperCase();
      const isExpense = /charge|fee|commission|debit|rtgs.*thrymr|remittance\s*charges|agent\s*charges|gst\s*on|autosave.*fee|transfer.*fee/i.test(desc);
      if (isExpense) debit  = txnAmt.toFixed(2);
      else           credit = txnAmt.toFixed(2);
    } else {
      // Only one amount — must be balance only (no transaction amount, skip row)
      prevBal = bal;
      continue;
    }

    prevBal = bal;

    // ── 6. Build clean narration ──────────────────────────────────
    const rawJoined = txn.descParts.join(" ").replace(/\s+/g, " ").trim();

    // Extract the primary ref code
    // DBS India: 0856IT0509568, IMPS-609112023254, EBHVT60403143451, BatchId:0039
    // DBS SG: EBGPP604084364350..., IBG BEXP-..., RTF EBACT...
    const refMatch = rawJoined.match(
      /\b(IMPS[-\s]\d{12,}|RTGS\s+\S+|NEFTIN\s+\S+|\d{4}IT\d{7,}|EBHVT\d{14,}|EBGPP\S{10,}|IBG\s+\S+|RTF\s+\S+|BatchId:\S+|[A-Z]{2,4}\d{10,})\b/i
    );
    const ref = refMatch ? refMatch[0].trim() : "";

    // Clean narration: remove ref codes, currency lines, "INR/USD XXXX AT INRXX" lines
    const narration = rawJoined
      .replace(/\b(IMPS[-\s]\d{12,}|RTGS\s+\S+|NEFTIN\s+\S+|\d{4}IT\d{7,}|EBHVT\d{14,}|EBGPP\S{10,}|IBG\s+\S+|RTF\s+EBACT\S*)\b/gi, "")
      .replace(/\b(USD|EUR|GBP|SGD|INR)\s+[\d,]+(\.\d+)?\s+(AT\s+INR[\d.]+)?/gi, "")
      .replace(/\b(BatchId:\S+|REF[-\s]?\S+|AC\s+ENDING\s+\w+|ICIC\w+|CITI\w+|UTIB\w+|BOFA\w+|CBIN\w+|KMB\w*)\b/gi, "")
      .replace(/\b26-\d{3}\b/g, "")
      .replace(/\s{2,}/g, " ")
      .replace(/^\s*TRANSFER\s*/i, "")
      .trim();

    if (!debit && !credit) continue;

    rows.push([
      txn.transDate,
      narration || rawJoined.slice(0, 80),
      ref,
      debit,
      credit,
      bal.toFixed(2),
    ]);
  }

  if (!rows.length) return null;

  return {
    headers: ["Date", "Narration", "Ref", "Debit", "Credit", "Balance"],
    rows,
    _bankHint: "dbs",
  };
}

// ── Generic / universal fallback parser (text-based) ─────────────
function parsePdfText(rawText) {
  const lines = rawText.split("\n").map(l=>l.trim()).filter(Boolean);
  const useTab = lines.filter(l=>l.includes("\t")).length > lines.length*0.3;
  const SEP = useTab ? /\t/ : /\s{2,}/;

  const isKVLine = line => {
    const colons=(line.match(/:/g)||[]).length;
    if(colons!==1)return false;
    const p=line.split(":").map(s=>s.trim());
    return /^[A-Za-z\s\.\(\)%\/]+$/.test(p[0])&&p[0].split(/\s+/).length<=5;
  };

  let hdrIdx=-1,best=0;
  for(let i=0;i<Math.min(lines.length,60);i++){
    if(RX_FOOTER.test(lines[i])||isKVLine(lines[i]))continue;
    const cols=lines[i].split(SEP).map(c=>c.trim()).filter(Boolean);
    if(cols.length<3)continue;
    const s=cols.filter(c=>RX_HDR.test(c)).length;
    const b=s>=4?3:s>=3?1:0;
    if((s+b)>best&&s>=2){best=s+b;hdrIdx=i;}
  }

  let synth=false;
  if(hdrIdx===-1){
    for(let i=0;i<Math.min(lines.length,60);i++){
      const c=lines[i].split(SEP).map(c=>c.trim()).filter(Boolean);
      if(c.length>=2&&isDateStr(c[0])){hdrIdx=Math.max(0,i-1);synth=true;break;}
    }
    if(hdrIdx===-1)for(let i=0;i<Math.min(lines.length,60);i++){
      const c=lines[i].split(SEP).map(c=>c.trim()).filter(Boolean);
      if(c.filter(isAmountStr).length>=2){hdrIdx=Math.max(0,i-1);synth=true;break;}
    }
    if(hdrIdx===-1)throw new Error("Could not find transaction data. Try downloading as Excel/CSV from your bank portal.");
  }

  let headers;
  if(synth){
    const fd=lines.slice(hdrIdx+1).find(l=>{const c=l.split(SEP).filter(Boolean);return c.length>=2&&(isDateStr(c[0])||isAmountStr(c[c.length-1]));});
    const cc=fd?fd.split(SEP).filter(Boolean).length:4;
    const gn=["Date","Description","Withdrawal","Deposit","Balance","Ref"];
    headers=Array.from({length:cc},(_,i)=>gn[i]||`Col${i+1}`);
  } else {
    const h=lines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
    const n=lines[hdrIdx+1]?lines[hdrIdx+1].split(SEP).map(c=>c.trim()).filter(Boolean):[];
    const ok=n.length>0&&n.length<=h.length&&!n.some(isDateStr)&&!n.some(isAmountStr)&&
      (n.filter(c=>RX_HDR.test(c)).length>=2||(n.length<=h.length&&n.every(w=>/^[A-Za-z\.\(\)\/]+$/.test(w))));
    if(ok){headers=h.map((x,i)=>n[i]?x+" "+n[i]:x);if(n.length>h.length)headers.push(...n.slice(h.length));hdrIdx++;}
    else headers=h;
  }

  const dataLines=lines.slice(hdrIdx+1).filter(l=>!RX_FOOTER.test(l)&&l.trim());
  const dateOnly=dataLines.filter(l=>isDateStr(l.split(SEP)[0])&&l.split(SEP).filter(Boolean).length<=2);
  if(!useTab&&dateOnly.length>dataLines.length*0.25){
    const txns=[];let cur=null;
    dataLines.forEach(line=>{
      if(RX_FOOTER.test(line))return;
      const p=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      if(isDateStr(p[0])){if(cur)txns.push(cur);cur={date:p[0],narr:p.slice(1).join(" "),amts:[]};}
      else if(cur){const a=p.filter(isAmountStr);const w=p.filter(x=>!isAmountStr(x));if(a.length)cur.amts.push(...a);if(w.length)cur.narr+=" "+w.join(" ");}
    });
    if(cur)txns.push(cur);
    const oR=txns.filter(t=>t.amts.length).map(t=>{
      const a=t.amts;const bal=a[a.length-1];const ta=a.length>=2?a[a.length-2]:a[0];
      const iD=/dr$/i.test(ta)||t.narr.toUpperCase().includes(" DR");const c=ta.replace(/[^0-9.]/g,"");
      return[t.date,t.narr.trim(),iD?c:"",iD?"":c,bal.replace(/[^0-9.]/g,"")];
    });
    if(oR.length)return{headers:["Date","Narration","Debit","Credit","Balance"],rows:oR};
  }

  const SNO=/^\d{1,6}$/;
  const NEWROW=c=>isDateStr(c[0])||SNO.test(c[0])||(c.length>=2&&isDateStr(c[1]));
  const tL=[];let cur2=null;
  lines.slice(hdrIdx+1).forEach(line=>{
    if(!line.trim()||RX_FOOTER.test(line)){if(cur2){tL.push(cur2);cur2=null;}return;}
    const c=line.split(SEP).map(x=>x.trim()).filter(Boolean);
    if(!c.length)return;
    if(NEWROW(c)){if(cur2)tL.push(cur2);cur2=[...c];}
    else if(cur2){
      c.forEach(tok=>{
        if(isAmountStr(tok)){let pl=false;for(let k=cur2.length-1;k>=Math.max(0,cur2.length-4);k--){if(!cur2[k]){cur2[k]=tok;pl=true;break;}}if(!pl)cur2.push(tok);}
        else if(/^(DR|CR)$/i.test(tok)){for(let k=cur2.length-1;k>=0;k--){if(isAmountStr(cur2[k])){cur2[k]+=" "+tok.toUpperCase();break;}}}
        else{let ni=-1;for(let k=1;k<cur2.length;k++){if(!isAmountStr(cur2[k])&&!isDateStr(cur2[k])&&!SNO.test(cur2[k])){ni=k;break;}}if(ni!==-1)cur2[ni]=(cur2[ni]+" "+tok).trim();else cur2.push(tok);}
      });
    } else{if(cur2)tL.push(cur2);cur2=[...c];}
  });
  if(cur2)tL.push(cur2);
  if(!tL.length)throw new Error("PDF parsed but no rows found. Try exporting as Excel or CSV.");

  const rows=tL.filter(r=>r.some(c=>c)).map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  if(headers.length<=3){
    const ac=headers.findIndex((_,i)=>rows.slice(0,10).filter(r=>isAmountStr(r[i])).length>3);
    if(ac!==-1){const ex=rows.map(r=>{const ra=(r[ac]||"").trim();const iD=/dr$/i.test(ra)||ra.startsWith("-");const a=ra.replace(/[^0-9.]/g,"");const re=r.filter((_,i)=>i!==ac);return[re[0]||"",re[1]||"",iD?a:"",iD?"":a,r[r.length-1]||""];});return{headers:["Date","Narration","Debit","Credit","Balance"],rows:ex};}
  }
  return{headers,rows};
}


// ── Canara Bank Statement ─────────────────────────────────────────
// Format: Txn Date | Value Date | Cheque No | Description | Branch Code | Debit | Credit | Balance
// Also handles older format: Date | Value Date | Ref | Description | Debit | Credit | Balance
function parseCanaraWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = -1, best = 0;
  allLines.forEach((line,i) => {
    if (i > 50 || RX_FOOTER.test(line)) return;
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    const score = cols.filter(c=>RX_HDR.test(c)).length;
    if (score >= 3 && score > best) { best=score; hdrIdx=i; }
  });
  if (hdrIdx === -1) return null;

  // Canara sometimes wraps "Branch Code" onto next line
  let headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const nxt = (allLines[hdrIdx+1]||"").split(SEP).map(c=>c.trim()).filter(Boolean);
  if (nxt.length && nxt.length <= headers.length && !nxt.some(isDateStr) && !nxt.some(isAmountStr)
      && nxt.every(w=>/^[A-Za-z\.\(\)\/\-]+$/.test(w))) {
    headers = headers.map((h,i) => nxt[i] ? h+" "+nxt[i] : h);
    hdrIdx++;
  }

  const NEWROW = cols => isDateStr(cols[0]) || (cols.length >= 2 && isDateStr(cols[1]));
  const txnRows = []; let cur = null;
  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if (isAmountStr(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if (cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"canara" } : null;
}

// ── Kotak Mahindra Bank Statement ────────────────────────────────
// Format: Sl.No | Date | Description | Chc/Ref number | Amount | Dr/Cr | Balance | Dr/Cr
// The Dr/Cr column tells direction; Amount is always positive
function parseKotakWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /description|narration/i.test(line) &&
    /dr.*cr|credit.*debit|amount/i.test(line));
  if (hdrIdx === -1) {
    hdrIdx = -1; let best = 0;
    allLines.forEach((line,i) => {
      if (i>60) return;
      const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score = cols.filter(c=>RX_HDR.test(c)).length;
      if (score >= 3 && score > best) { best=score; hdrIdx=i; }
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const drcrIdx = headers.findIndex(h => /^dr\/cr$|^cr\/dr$/i.test(h));
  const amtIdx  = headers.findIndex(h => /^amount$/i.test(h));

  const NEWROW = cols => {
    const first = cols[0];
    // Kotak has serial number first, then date
    return /^\d{1,4}$/.test(first) ? isDateStr(cols[1]) : isDateStr(first);
  };
  const txnRows = []; let cur = null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols = line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if (!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if (isAmountStr(tok)) cur.push(tok);
        else if (/^(DR|CR)$/i.test(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)&&!/^(DR|CR)$/i.test(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if (cur) txnRows.push(cur);

  // Normalise: if Dr/Cr + Amount columns exist, split into Debit/Credit
  let outHeaders = headers;
  let outRows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});

  if (drcrIdx !== -1 && amtIdx !== -1) {
    outHeaders = headers.filter((_,i)=>i!==drcrIdx && i!==amtIdx).concat(["Debit","Credit","Balance"]);
    const balIdx = headers.findIndex(h=>/balance/i.test(h));
    outRows = outRows.map(r => {
      const drCr = (r[drcrIdx]||"").toUpperCase();
      const amt  = r[amtIdx] || "";
      const bal  = balIdx !== -1 ? r[balIdx] : r[r.length-1];
      const rest = r.filter((_,i)=>i!==drcrIdx&&i!==amtIdx&&i!==balIdx);
      return [...rest, drCr==="DR"?amt:"", drCr==="CR"?amt:"", bal];
    });
  }

  return outRows.length ? { headers: outHeaders, rows: outRows, _bankHint:"kotak" } : null;
}

// ── Andhra Bank / Union Bank / UCO Bank Statement ─────────────────
// Format: Tran Id | Tran Date | Remarks | Amount (Rs.) | Balance (Rs.)
// Amount has "(Dr)" or "(Cr)" suffix embedded in same cell
function parseAndhraWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /tran.*date|transaction.*date/i.test(line) &&
    /remarks|description|narration/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>60)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  // Remove Rs. suffix from column headers for cleaner display
  const cleanHeaders = headers.map(h => h.replace(/\s*\(rs\.?\)/i,"").trim());

  const NEWROW = cols => isDateStr(cols[0]) || (cols.length>=2 && isDateStr(cols[1]));
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if(isAmountStr(tok.replace(/\s*\(dr\)|\s*\(cr\)/gi,""))) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c.replace(/\s*\(dr\)|\s*\(cr\)/gi,""))&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if (cur) txnRows.push(cur);

  // Andhra encodes Dr/Cr as "300.00 (Dr)" in the Amount column
  // Split into Debit / Credit columns
  const amtColIdx = cleanHeaders.findIndex(h=>/^amount$/i.test(h));
  const balColIdx = cleanHeaders.findIndex(h=>/balance/i.test(h));

  let outHeaders = cleanHeaders;
  let outRows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<cleanHeaders.length)r.push("");return r.slice(0,cleanHeaders.length);});

  if (amtColIdx !== -1) {
    const newHdrs = cleanHeaders.filter((_,i)=>i!==amtColIdx).slice(0,balColIdx!==-1?balColIdx:-1);
    outHeaders = [...(balColIdx!==-1?cleanHeaders.filter((_,i)=>i!==amtColIdx&&i!==balColIdx):cleanHeaders.filter((_,i)=>i!==amtColIdx)), "Debit","Credit","Balance"];
    outRows = outRows.map(r => {
      const raw  = r[amtColIdx] || "";
      const isDr = /dr/i.test(raw);
      const amt  = raw.replace(/[^0-9.]/g,"");
      const bal  = balColIdx !== -1 ? r[balColIdx] : "";
      const rest = r.filter((_,i)=>i!==amtColIdx&&i!==balColIdx);
      return [...rest, isDr?amt:"", isDr?"":amt, bal];
    });
  }

  return outRows.length ? { headers: outHeaders, rows: outRows, _bankHint:"andhra" } : null;
}

// ── Punjab National Bank (PNB) Statement ─────────────────────────
// Format: Transaction Date | Cheque Number | Withdrawal | Deposit | Balance | Narration
// Balance has "Cr." suffix
function parsePNBWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /transaction.*date|txn.*date/i.test(line) &&
    /withdrawal|deposit/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>60)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  // PNB wraps "Transaction Date" across two lines
  let headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const nxt = (allLines[hdrIdx+1]||"").split(SEP).map(c=>c.trim()).filter(Boolean);
  if (nxt.length && nxt.length <= headers.length && !nxt.some(isDateStr) && !nxt.some(isAmountStr)
      && nxt.every(w=>/^[A-Za-z\.\(\)\/\-]+$/.test(w))) {
    headers = headers.map((h,i) => nxt[i] ? h+" "+nxt[i] : h);
    hdrIdx++;
  }

  const NEWROW = cols => isDateStr(cols[0]);
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        // PNB balance has "Cr." suffix e.g. "570.53 Cr."
        const clean=tok.replace(/\s*(cr\.?|dr\.?)$/i,"");
        if(isAmountStr(clean)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c.replace(/\s*(cr\.?|dr\.?)$/i,""))&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"pnb" } : null;
}

// ── Central Bank of India Statement ──────────────────────────────
// Format: Value Date | Post Date | Details | Chq.No. | Debit | Credit | Balance
// Balance has "Cr" or "Dr" suffix. Multi-line descriptions grouped by blank Post Date
function parseCentralBankWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /value\s*date|post\s*date/i.test(line) &&
    /debit|credit/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>60)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  // Central Bank: new txn has date in col 0; continuation rows have "." or blank in col 0
  const NEWROW = cols => isDateStr(cols[0]);
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length || cols[0]===".") {
      // continuation line — merge description
      if(cur) {
        const descIdx=cur.findIndex((c,i)=>i>1&&!isAmountStr(c.replace(/\s*(cr|dr)$/i,""))&&!isDateStr(c));
        if(descIdx!==-1&&cols.length) cur[descIdx]+=" "+cols.join(" ");
      }
      return;
    }
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        const clean=tok.replace(/\s*(cr|dr)$/i,"");
        if(isAmountStr(clean)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>1&&!isAmountStr(c.replace(/\s*(cr|dr)$/i,""))&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  // Clean "Cr"/"Dr" suffix from balance
  const rows = txnRows.filter(r=>r.some(Boolean)).map(r => {
    const out = r.map(c => typeof c==="string" ? c.replace(/\s+(Cr|Dr)$/i,"").trim() : c);
    while(out.length<headers.length)out.push("");
    return out.slice(0,headers.length);
  });
  return rows.length ? { headers, rows, _bankHint:"centralbank" } : null;
}

// ── Bank of India Statement ───────────────────────────────────────
// Format: SI No | Txn Date | Description | Cheque No | Withdrawal | Deposits | Balance
// Older format with wide narration, SI number starts each row
function parseBOIWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /withdrawal|deposits|deposit/i.test(line) &&
    /txn.*date|date/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>60)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  // BOI rows start with serial number or date
  const NEWROW = cols => /^\d{1,6}$/.test(cols[0]) || isDateStr(cols[0]) || (isDateStr(cols[1]||""));
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if(isAmountStr(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)&&!/^\d{1,6}$/.test(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"boi" } : null;
}

// ── Bank of Baroda Statement ──────────────────────────────────────
// Format: Value Date | Post Date | Details | Chq.No. | Debit | Credit | Balance
// Balance has "Cr" suffix; continuation rows have blank Value Date
function parseBOBWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /value\s*date|post\s*date/i.test(line) &&
    /debit|credit/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>60)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const NEWROW = cols => isDateStr(cols[0]);
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        const clean=tok.replace(/\s*(cr|dr)$/i,"");
        if(isAmountStr(clean)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c.replace(/\s*(cr|dr)$/i,""))&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean)).map(r => {
    const out=r.map(c=>typeof c==="string"?c.replace(/\s+(Cr|Dr)$/i,"").trim():c);
    while(out.length<headers.length)out.push("");
    return out.slice(0,headers.length);
  });
  return rows.length ? { headers, rows, _bankHint:"bob" } : null;
}

// ── RBL Bank Credit Card Statement ───────────────────────────────
// Format: Date | Description | Amount (can be debit or credit)
// Opening/closing balance in header. Amounts with sign or label
function parseRBLWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<80 &&
    /date/i.test(line) &&
    /description|narration|particulars/i.test(line) &&
    /amount/i.test(line));
  if (hdrIdx === -1) {
    let best=0;
    allLines.forEach((line,i) => {
      if(i>80)return;
      const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
      const score=cols.filter(c=>RX_HDR.test(c)).length;
      if(score>=3&&score>best){best=score;hdrIdx=i;}
    });
  }
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const NEWROW = cols => isDateStr(cols[0]);
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)||/reward.*summary|account.*summary|payment.*due|minimum.*amount/i.test(line)) {
      if(cur){txnRows.push(cur);cur=null;} return;
    }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        if(isAmountStr(tok)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c)&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean))
    .map(r=>{while(r.length<headers.length)r.push("");return r.slice(0,headers.length);});
  return rows.length ? { headers, rows, _bankHint:"rbl" } : null;
}

// ── Transaction History style (SBI/IOB/BOB online portal exports) ─
// Format: Post Date | Value Date | Description | Debit | Credit | Balance
// These have "CR" suffix on Balance column
function parseTransactionHistoryWords(pages) {
  const allLines = pages.flatMap(({items,width}) => linesToText(groupIntoLines(items,3), width));
  const SEP = /\t|\s{2,}/;

  let hdrIdx = allLines.findIndex((line,i) => i<60 &&
    /post\s*date|value\s*date/i.test(line) &&
    /description|narration/i.test(line));
  if (hdrIdx === -1) return null;

  const headers = allLines[hdrIdx].split(SEP).map(c=>c.trim()).filter(Boolean);
  const NEWROW = cols => isDateStr(cols[0]);
  const txnRows=[]; let cur=null;

  allLines.slice(hdrIdx+1).forEach(line => {
    if (RX_FOOTER.test(line)) { if(cur){txnRows.push(cur);cur=null;} return; }
    const cols=line.split(SEP).map(c=>c.trim()).filter(Boolean);
    if(!cols.length) return;
    if (NEWROW(cols)) { if(cur)txnRows.push(cur); cur=[...cols]; }
    else if (cur) {
      cols.forEach(tok => {
        const clean=tok.replace(/\s*(cr|dr)$/i,"");
        if(isAmountStr(clean)) cur.push(tok);
        else { const ni=cur.findIndex((c,i)=>i>0&&!isAmountStr(c.replace(/\s*(cr|dr)$/i,""))&&!isDateStr(c)); if(ni!==-1)cur[ni]+=" "+tok; else cur.push(tok); }
      });
    }
  });
  if(cur) txnRows.push(cur);

  const rows = txnRows.filter(r=>r.some(Boolean)).map(r => {
    const out=r.map(c=>typeof c==="string"?c.replace(/\s+(CR|DR)$/i,"").trim():c);
    while(out.length<headers.length)out.push("");
    return out.slice(0,headers.length);
  });
  return rows.length ? { headers, rows, _bankHint:"txnhistory" } : null;
}

// ── Generic text extraction for the universal parser ──────────────
async function extractPdfText(buf) {
  const pages = await extractWordItems(buf);
  return pages.map(({items, width}) => linesToText(groupIntoLines(items,3), width).join("\n")).join("\n");
}

// ── OCR fallback for scanned PDFs ────────────────────────────────
async function ocrPdfText(buf, onProgress) {
  await loadPdfJs();
  await loadScript("https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js");
  const pdfjsLib = window["pdfjs-dist/build/pdf"];
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  const worker = await window.Tesseract.createWorker("eng");
  let fullText = "";
  for (let i=1;i<=pdf.numPages;i++) {
    onProgress&&onProgress(`OCR page ${i}/${pdf.numPages}…`);
    const page=await pdf.getPage(i);const vp=page.getViewport({scale:2.5});
    const canvas=document.createElement("canvas");canvas.width=vp.width;canvas.height=vp.height;
    await page.render({canvasContext:canvas.getContext("2d"),viewport:vp}).promise;
    const {data:{text}}=await worker.recognize(canvas.toDataURL("image/png"));
    fullText+=text+"\n";
  }
  await worker.terminate();
  return fullText.trim();
}

// ── File Parser entry point ───────────────────────────────────────
async function parseFile(file, onProgress) {
  const ext = file.name.split(".").pop().toLowerCase();

  // ── PDF path ──────────────────────────────────────────────────────
  if (ext === "pdf") {
    const buf = await file.arrayBuffer();
    const headerBytes = new TextDecoder("latin1").decode(new Uint8Array(buf.slice(0,2048)));
    if (/\/Encrypt/i.test(headerBytes)) {
      const e = new Error("This PDF is password-protected. Open it in Acrobat → File → Save as (removes password), then re-upload.");
      e.code = "ERR_002"; throw e;
    }
    onProgress && onProgress("Loading PDF engine…");

    // ── Step 1: Extract word items with X/Y positions ──────────────
    let pages = [];
    let isScanned = false;
    try {
      pages = await extractWordItems(buf);
      const totalWords = pages.reduce((n, p) => n + p.items.length, 0);
      isScanned = totalWords < 30; // very few words = scanned image PDF
    } catch { isScanned = true; }

    // ── Step 2: Scanned PDF → OCR ──────────────────────────────────
    if (isScanned) {
      onProgress && onProgress("Scanned PDF detected — starting OCR (may take 30–60 s)…");
      const ocrText = await ocrPdfText(buf, onProgress);
      onProgress && onProgress("Parsing OCR output…");
      return parsePdfText(ocrText);
    }

    // ── Step 3: Detect bank and route to the right parser ──────────
    const bank = detectBank(pages);
    onProgress && onProgress(`Detected: ${bank.toUpperCase()} — parsing…`);

    const PARSERS = {
      dbs:         parseDBSWords,
      icici:       parseICICIWords,
      sbi:         parseSBIWords,
      hdfc:        parseHDFCWords,
      axis:        parseAxisWords,
      canara:      parseCanaraWords,
      kotak:       parseKotakWords,
      andhra:      parseAndhraWords,
      pnb:         parsePNBWords,
      centralbank: parseCentralBankWords,
      boi:         parseBOIWords,
      bob:         parseBOBWords,
      rbl:         parseRBLWords,
      txnhistory:  parseTransactionHistoryWords,
      union:       parseAndhraWords,
      federal:     parseCanaraWords,
      indus:       parseCanaraWords,
      yes:         parseCanaraWords,
    };

    // Try bank-specific parser first
    if (PARSERS[bank]) {
      try {
        const result = PARSERS[bank](pages);
        if (result && result.rows.length > 0) return result;
        // If it returned 0 rows, fall through to generic
        onProgress && onProgress("Bank parser found no rows — trying generic…");
      } catch (e) {
        onProgress && onProgress("Bank parser error — trying generic…");
        console.warn(`${bank} parser failed:`, e);
      }
    }

    // ── Step 4: Generic fallback (works for Kotak, PNB, Yes, IDFC, etc.) ──
    onProgress && onProgress("Parsing transaction table…");
    const text = pages.map(({items, width}) => linesToText(groupIntoLines(items, 3), width).join("\n")).join("\n");
    return parsePdfText(text);
  }

  // ── Excel / CSV path ─────────────────────────────────────────────
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true, raw:false });

  // Find the most data-rich sheet (some banks put statement on sheet 2)
  let bestSheet = wb.SheetNames[0];
  let bestCount = 0;
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    const count = (range.e.r - range.s.r) * (range.e.c - range.s.c);
    if (count > bestCount) { bestCount = count; bestSheet = name; }
  });

  const raw = XLSX.utils.sheet_to_json(wb.Sheets[bestSheet], { header:1, defval:"", raw:false });

  // Find the real header row — skip bank metadata rows at top
  // Strategy: score each row by how many cells match bank header keywords.
  // KV rows like ["Date", "8 May 2026"] or ["Account Number", "0001234"] must be skipped.
  const isKVRow = (row) => {
    // A KV row has exactly 2 non-empty cells where col[0] is a short label word(s) with no digits
    const cells = row.map(c=>String(c).trim()).filter(Boolean);
    if (cells.length !== 2) return false;
    return /^[A-Za-z\s\.\(\)%\/]+$/.test(cells[0]) && cells[0].split(/\s+/).length <= 5;
  };

  let headerRowIdx = 0;
  let bestHdrScore = 0;
  for (let i = 0; i < Math.min(raw.length, 30); i++) {
    const row = raw[i].map(c=>String(c).trim());
    const nonEmpty = row.filter(Boolean);
    if (nonEmpty.length < 3) continue;    // real headers have 3+ columns
    if (isKVRow(nonEmpty)) continue;      // skip "Label | Value" pairs
    const score = nonEmpty.filter(c => RX_HDR.test(c)).length;
    const bonus = score >= 4 ? 3 : score >= 3 ? 1 : 0;
    if ((score + bonus) > bestHdrScore && score >= 2) { bestHdrScore = score + bonus; headerRowIdx = i; }
  }
  // Fallback: first row with 3+ non-empty cells that has a date somewhere nearby
  if (bestHdrScore === 0) {
    for (let i = 0; i < Math.min(raw.length, 20); i++) {
      const row = raw[i].map(c=>String(c).trim()).filter(Boolean);
      if (row.length >= 3) { headerRowIdx = i; break; }
    }
  }

  const rawHeaders = raw[headerRowIdx].map(h => String(h).trim());
  // Remove empty trailing headers and deduplicate blank ones
  const headers = rawHeaders.map((h,i) => h || `Col${i+1}`);
  // Trim to last non-empty header
  let lastNonEmpty = headers.length-1;
  while (lastNonEmpty > 0 && !raw[headerRowIdx][lastNonEmpty]) lastNonEmpty--;
  const finalHeaders = headers.slice(0, lastNonEmpty+1);

  const rows = raw.slice(headerRowIdx+1)
    .map(r => finalHeaders.map((_,i) => String(r[i]??"")))
    .filter(r => r.some(c => c.trim()));

  if (!finalHeaders.length || !rows.length) throw new Error("No data found in this file. Check the sheet and try again.");
  return { headers: finalHeaders, rows };
}

// ── Styled Primitives ────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=DM+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${T.bg}; color: ${T.text}; font-family: ${T.font}; }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: #e8edf6; }
  ::-webkit-scrollbar-thumb { background: #c8d3e8; border-radius: 99px; }
  select, input, textarea { background: ${T.surface}; color: ${T.text}; border: 1px solid ${T.border}; border-radius: 8px; font-family: ${T.font}; font-size: 13px; }
  select:focus, input:focus { outline: none; border-color: ${T.accent}; box-shadow: 0 0 0 3px ${T.accentGlow}; }
  @keyframes fadeIn { from { opacity:0; transform:translateY(8px) } to { opacity:1; transform:translateY(0) } }
  @keyframes shimmer { from { background-position: -200% 0 } to { background-position: 200% 0 } }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.5} }
  @keyframes slideIn { from{transform:translateX(100%);opacity:0} to{transform:translateX(0);opacity:1} }
  .fade-in { animation: fadeIn 0.3s ease forwards; }
  .row-hover:hover { background: #f0f4f9 !important; }
  .btn-hover:hover { filter: brightness(0.95); transform: translateY(-1px); }
  .card-hover:hover { border-color: ${T.accent} !important; box-shadow: 0 0 0 3px ${T.accentGlow}; }
`;

function Pill({ children, color = "gray", size = "sm", dot = false }) {
  const map = {
    gray:   [T.textDim,  T.surface],
    blue:   [T.accent,   T.accentDim],
    green:  [T.green,    T.greenDim],
    amber:  [T.amber,    T.amberDim],
    red:    [T.red,      T.redDim],
    purple: [T.purple,   T.purpleDim],
  };
  const [col, bg] = map[color] || map.gray;
  return (
    <span style={{ display:"inline-flex", alignItems:"center", gap:4, background:bg, color:col, fontSize:size==="xs"?10:11, fontWeight:600, padding:size==="xs"?"2px 7px":"3px 9px", borderRadius:99, letterSpacing:"0.02em", border:`1px solid ${col}22` }}>
      {dot && <span style={{width:5,height:5,borderRadius:"50%",background:col,flexShrink:0}}/>}
      {children}
    </span>
  );
}

function Btn({ children, onClick, variant="primary", disabled, icon, size="md", fullWidth, style:extra={} }) {
  const sz = size==="sm" ? { padding:"5px 12px", fontSize:12 } : size==="lg" ? { padding:"11px 22px", fontSize:15 } : { padding:"8px 16px", fontSize:13 };
  const vars = {
    primary: { background:`linear-gradient(135deg, ${T.accent}, #3b6fd4)`, color:"#fff", border:"none", boxShadow:`0 0 20px ${T.accentGlow}` },
    secondary: { background:T.card, color:T.text, border:`1px solid ${T.border}` },
    outline: { background:"transparent", color:T.accent, border:`1px solid ${T.accent}` },
    ghost: { background:"transparent", color:T.textMid, border:"none" },
    danger: { background:T.redDim, color:T.red, border:`1px solid ${T.red}44` },
    success: { background:T.greenDim, color:T.green, border:`1px solid ${T.green}44` },
    amber: { background:T.amberDim, color:T.amber, border:`1px solid ${T.amber}44` },
  };
  return (
    <button onClick={onClick} disabled={disabled} className="btn-hover"
      style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", gap:6, borderRadius:9, fontFamily:T.font, fontWeight:500, cursor:disabled?"not-allowed":"pointer", opacity:disabled?0.45:1, transition:"all 0.18s", width:fullWidth?"100%":undefined, ...sz, ...vars[variant], ...extra }}>
      {icon && <span style={{fontSize:size==="sm"?13:15}}>{icon}</span>}
      {children}
    </button>
  );
}

function Card({ children, style:extra={}, className="" }) {
  return <div className={className} style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:20, boxShadow:"0 1px 4px rgba(0,0,0,0.06), 0 4px 16px rgba(0,0,0,0.04)", ...extra }}>{children}</div>;
}

function Input({ value, onChange, placeholder, prefix, suffix, style:extra={} }) {
  return (
    <div style={{ position:"relative", display:"flex", alignItems:"center" }}>
      {prefix && <span style={{ position:"absolute", left:10, color:T.textDim, fontSize:13 }}>{prefix}</span>}
      <input value={value} onChange={onChange} placeholder={placeholder}
        style={{ width:"100%", padding:"8px 12px", paddingLeft:prefix?32:12, paddingRight:suffix?32:12, ...extra }} />
      {suffix && <span style={{ position:"absolute", right:10, color:T.textDim, fontSize:13 }}>{suffix}</span>}
    </div>
  );
}

function Modal({ open, onClose, title, children, width=540 }) {
  if (!open) return null;
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,0.45)", zIndex:9000, display:"flex", alignItems:"center", justifyContent:"center", backdropFilter:"blur(4px)" }} onClick={onClose}>
      <div className="fade-in" style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:18, width, maxWidth:"95vw", maxHeight:"85vh", overflow:"auto", padding:28 }} onClick={e=>e.stopPropagation()}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20 }}>
          <span style={{ fontWeight:700, fontSize:16, color:T.text }}>{title}</span>
          <button onClick={onClose} style={{ background:"none", border:"none", color:T.textMid, cursor:"pointer", fontSize:20, lineHeight:1 }}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function Toast({ toasts }) {
  return (
    <div style={{ position:"fixed", top:20, right:20, zIndex:9999, display:"flex", flexDirection:"column", gap:8 }}>
      {toasts.map(t => (
        <div key={t.id} className="fade-in"
          style={{ background:T.card, border:`1px solid ${t.type==="error"?T.red:t.type==="warn"?T.amber:T.green}55`, borderRadius:11, padding:"12px 18px", fontSize:13, color:t.type==="error"?T.red:t.type==="warn"?T.amber:T.green, boxShadow:`0 8px 32px rgba(0,0,0,0.4)`, maxWidth:380, display:"flex", alignItems:"center", gap:8 }}>
          <span>{t.type==="error"?"✕":t.type==="warn"?"⚠":"✓"}</span>
          <span style={{ color:T.text }}>{t.msg}</span>
        </div>
      ))}
    </div>
  );
}

function Steps({ steps, current }) {
  return (
    <div style={{ display:"flex", alignItems:"center", gap:0, marginBottom:24 }}>
      {steps.map((s, i) => (
        <div key={i} style={{ display:"flex", alignItems:"center", flex: i < steps.length-1 ? 1 : "none" }}>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <div style={{ width:30, height:30, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700,
              background: i < current ? T.green : i === current ? T.accent : T.border,
              color: i <= current ? "#fff" : T.textDim,
              boxShadow: i === current ? `0 0 16px ${T.accentGlow}` : "none",
              flexShrink:0, transition:"all 0.3s" }}>
              {i < current ? "✓" : i+1}
            </div>
            <span style={{ fontSize:12, fontWeight:i===current?600:400, color:i===current?T.text:T.textDim, whiteSpace:"nowrap" }}>{s}</span>
          </div>
          {i < steps.length-1 && (
            <div style={{ flex:1, height:1, background:`linear-gradient(90deg, ${i < current ? T.green : T.border}, ${i+1 <= current ? T.green : T.border})`, margin:"0 10px", transition:"background 0.5s" }} />
          )}
        </div>
      ))}
    </div>
  );
}

function StatCard({ icon, label, value, sub, color=T.accent }) {
  return (
    <Card style={{ padding:"18px 20px", position:"relative", overflow:"hidden", borderLeft:`3px solid ${color}` }}>
      <div style={{ fontSize:24, fontWeight:800, color, letterSpacing:"-0.5px", marginBottom:4 }}>{value}</div>
      <div style={{ fontSize:12, fontWeight:600, color:T.text, marginBottom:2 }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:T.textSub, marginTop:2 }}>{sub}</div>}
    </Card>
  );
}

// ── ERR Card ─────────────────────────────────────────────────────
function ErrCard({ code, message, onDismiss }) {
  return (
    <div style={{ background:T.redDim, border:`1px solid ${T.red}55`, borderRadius:12, padding:"14px 18px", marginBottom:16 }} className="fade-in">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
        <div style={{ flex:1 }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:8 }}>
            <Pill color="red">{code}</Pill>
            <span style={{ fontWeight:600, color:T.red, fontSize:13 }}>{code === "ERR_002" ? "Password-Protected PDF" : code === "ERR_PDF_AGENT" ? "PDF Agent Required" : "Import Error"}</span>
          </div>
          <p style={{ color:T.textMid, fontSize:13, lineHeight:1.6 }}>{message}</p>
          {code === "ERR_002" && (
            <div style={{ marginTop:12, background:"rgba(0,0,0,0.2)", borderRadius:8, padding:"10px 14px" }}>
              <p style={{ fontWeight:600, fontSize:12, color:T.red, marginBottom:6 }}>How to fix ERR_002:</p>
              <ol style={{ paddingLeft:16, fontSize:12, color:T.textMid, lineHeight:2 }}>
                <li>Log in to your bank's net banking portal</li>
                <li>Download the statement as <strong style={{color:T.text}}>Excel (.xlsx)</strong> or <strong style={{color:T.text}}>CSV</strong> format</li>
                <li>Alternatively: open the PDF → Print → Save as new PDF (removes password)</li>
                <li>Upload the unlocked file here</li>
              </ol>
            </div>
          )}
          {code === "ERR_PDF_AGENT" && (
            <div style={{ marginTop:12, display:"flex", gap:8 }}>
              <Btn size="sm" variant="outline" icon="⬇">Download Desktop Agent</Btn>
              <Btn size="sm" variant="ghost">Use Excel/CSV instead</Btn>
            </div>
          )}
        </div>
        {onDismiss && <button onClick={onDismiss} style={{ background:"none", border:"none", cursor:"pointer", color:T.textDim, fontSize:18, marginLeft:12, flexShrink:0 }}>×</button>}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Login / Register / Pending Approval
// ══════════════════════════════════════════════════════════════════

// ── Extension Status Component ───────────────────────────────────
function ExtensionStatus() {
  const [status, setStatus] = useState("checking"); // checking | installed | missing

  useEffect(() => {
    // If either flag already set (e.g. from earlier ping), resolve immediately
    if (_extensionReady || window.__bank2tallyExtension) {
      setStatus("installed");
      return;
    }

    const handler = (e) => {
      if (e.data?.type === "BANK2TALLY_EXTENSION_PRESENT") {
        _markExtensionReady();
        setStatus("installed");
        clearTimeout(timer);
      }
    };
    window.addEventListener("message", handler);
    window.postMessage({ type: "CHECK_EXTENSION" }, "*");

    // After 2s with no signal, mark as missing
    const timer = setTimeout(() => {
      setStatus(_extensionReady || window.__bank2tallyExtension ? "installed" : "missing");
    }, 2000);

    return () => { window.removeEventListener("message", handler); clearTimeout(timer); };
  }, []);

  if (status === "checking") return (
    <div style={{ marginTop:14, background:T.surface, borderRadius:8, padding:"10px 14px", fontSize:12, color:T.textDim }}>
      ⏳ Checking for Bank2Tally Connector extension...
    </div>
  );

  if (status === "installed") return (
    <div style={{ marginTop:14, background:T.greenDim, border:`1px solid ${T.green}44`, borderRadius:8, padding:"10px 14px" }}>
      <p style={{ fontSize:12, fontWeight:600, color:T.green }}>✓ Bank2Tally Connector Installed</p>
      <p style={{ fontSize:11, color:T.textMid, marginTop:4 }}>Extension is active. Tally connection is ready.</p>
    </div>
  );

  return (
    <div style={{ marginTop:14, background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"12px 14px" }}>
      <p style={{ fontSize:12, fontWeight:700, color:T.red, marginBottom:6 }}>⚠ Extension Not Installed</p>
      <p style={{ fontSize:11, color:T.textMid, lineHeight:1.7, marginBottom:10 }}>
        Install the <strong style={{color:T.text}}>Bank2Tally Connector</strong> Chrome extension to connect to Tally. One-time setup, no software needed.
      </p>
      <div style={{ background:T.surface, borderRadius:7, padding:"8px 12px", fontSize:11, color:T.textMid, lineHeight:1.8, marginBottom:10 }}>
        <strong style={{color:T.text}}>How to install:</strong><br/>
        1. Download the extension ZIP below<br/>
        2. Open Chrome → <code style={{color:T.accent}}>chrome://extensions</code><br/>
        3. Enable <strong style={{color:T.text}}>Developer Mode</strong> (top right)<br/>
        4. Click <strong style={{color:T.text}}>Load unpacked</strong> → select extracted folder<br/>
        5. Refresh this page ✓
      </div>
      <a href="/tally-extension.zip" download
        style={{ display:"inline-flex", alignItems:"center", gap:6, padding:"8px 14px", background:T.accent, color:"#fff", borderRadius:7, fontSize:12, fontWeight:600, textDecoration:"none" }}>
        ⬇ Download Extension
      </a>
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [tab, setTab] = useState("login"); // login | register | forgot
  // Login
  const [email, setEmail] = useState("");
  const [pass, setPass]   = useState("");
  // Register
  const [regName,    setRegName]    = useState("");
  const [regEmail,   setRegEmail]   = useState("");
  const [regPass,    setRegPass]    = useState("");
  const [regPass2,   setRegPass2]   = useState("");
  const [regCompany, setRegCompany] = useState("");
  // Forgot password
  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotSent,  setForgotSent]  = useState(false);
  // State
  const [loading, setLoading] = useState(false);
  const [err,     setErr]     = useState("");
  const [success, setSuccess] = useState("");
  const [pendingUser, setPendingUser] = useState(null); // waiting for approval

  // ── Login ────────────────────────────────────────────────────────
  const [loginAttempts, setLoginAttempts] = useState(0);
  const [loginLockedUntil, setLoginLockedUntil] = useState(null);

  const handleLogin = async () => {
    // Client-side rate limiting — 5 attempts then 60s lockout
    if (loginLockedUntil && Date.now() < loginLockedUntil) {
      const secs = Math.ceil((loginLockedUntil - Date.now()) / 1000);
      setErr(`Too many failed attempts. Try again in ${secs} seconds.`);
      return;
    }
    setErr(""); setSuccess(""); setLoading(true);
    try {
      const session = await sb.signIn(email, pass);
      const userEmail = session.user.email;
      const userId    = session.user.id;
      // Fetch profile (no email column — email comes from session)
      let profiles = [];
      try { profiles = await sb.from("profiles", `id=eq.${userId}&select=*`); } catch {}
      const profile = profiles[0];
      if (!profile) {
        // No profile row yet — auto-allow if Supabase metadata says admin, else block
        const meta = session.user.user_metadata || {};
        const role = meta.role || "user";
        if (role === "admin") {
          onLogin({ id: userId, name: meta.name || userEmail.split("@")[0], role: "admin", status: "approved", avatar: upper2(meta.name || userEmail), email: userEmail, mobile: profile?.mobile||"", mobile_verified: profile?.mobile_verified||false, sessionToken: session.access_token });
          return;
        }
        throw new Error("Your account is pending admin approval.");
      }
      if (profile.status === "pending")  { setPendingUser({ ...profile, email: userEmail }); setLoading(false); return; }
      if (profile.status === "on_hold")  throw new Error("Your account is on hold. Contact admin.");
      if (profile.status === "rejected") throw new Error("Your access request was rejected. Contact admin.");
      // Backfill email into profiles if missing (fixes old accounts)
      if (!profile.email && userEmail) {
        try { await sb.update("profiles", { id: userId }, { email: userEmail }); } catch {}
      }
      // Normalise role to lowercase DB format
      const normRole = (profile.role || "user").toLowerCase().trim();
      onLogin({ ...profile, role: normRole, email: userEmail, sessionToken: session.access_token });
    } catch (e) {
      const attempts = loginAttempts + 1;
      setLoginAttempts(attempts);
      if (attempts >= 5) {
        setLoginLockedUntil(Date.now() + 60000);
        setLoginAttempts(0);
        setErr('Too many failed attempts. Please wait 60 seconds before trying again.');
      } else {
        setErr(e.message + (attempts >= 3 ? ` (${5-attempts} attempts remaining)` : ''));
      }
    }
    setLoading(false);
  };

  const upper2 = str => (str || "").slice(0,2).toUpperCase();

  // ── Email verification state ─────────────────────────────────────
  const [verifyStep,  setVerifyStep]  = useState(false); // show "check email" screen
  const [pendingReg,  setPendingReg]  = useState(null);  // holds reg data

  // ── Password recovery (from email link) ────────────────────────
  const [recoveryToken, setRecoveryToken] = useState(null); // access_token from URL hash
  const [newPass,       setNewPass]       = useState("");
  const [newPass2,      setNewPass2]      = useState("");
  const [resetDone,     setResetDone]     = useState(false);

  useEffect(() => {
    // Supabase puts #access_token=...&type=recovery in the URL when the user
    // clicks the password-reset email link. Detect it here and switch to the
    // set-new-password form instead of showing the normal login screen.
    const hash = window.location.hash;
    if (!hash) return;
    const params = new URLSearchParams(hash.replace(/^#/, ""));
    if (params.get("type") === "recovery" && params.get("access_token")) {
      const token = params.get("access_token");
      sb._token = token; // authenticate the sb client with this one-time token
      setRecoveryToken(token);
      // Clean the URL so a refresh doesn't re-trigger recovery mode
      window.history.replaceState(null, "", window.location.pathname + window.location.search);
    }
  }, []);

  const handleSetNewPassword = async () => {
    setErr(""); setSuccess("");
    if (newPass.length < 8)      return setErr("Password must be at least 8 characters.");
    if (!/[0-9!@#$%^&*()]/.test(newPass)) return setErr("Password must contain at least one number or special character.");
    if (newPass !== newPass2)    return setErr("Passwords do not match.");
    setLoading(true);
    try {
      await sb.updatePassword(newPass);
      setResetDone(true);
      setNewPass(""); setNewPass2("");
    } catch (e) { setErr(e.message); }
    setLoading(false);
  };

  // ── Register ─────────────────────────────────────────────────────
  const handleRegister = async () => {
    setErr(""); setSuccess("");
    if (!regName.trim())         return setErr("Full name is required.");
    if (!regEmail.includes("@")) return setErr("Enter a valid email address.");
    if (regPass.length < 8)      return setErr("Password must be at least 8 characters.");
    if (!/[0-9!@#$%^&*()_+\-=\[\]{}]/.test(regPass)) return setErr("Password must contain at least one number or special character.");
    if (regPass !== regPass2)    return setErr("Passwords do not match.");
    setLoading(true);
    try {
      const session = await sb.signUp(regEmail, regPass, {
        name: regName.trim(),
        role: "user",
        company: regCompany.trim(),
      });
      // Insert profile row with pending status
      const userId = session.user?.id;
      if (userId) {
        try {
          await sb.insert("profiles", {
            id:      userId,
            name:    regName.trim(),
            email:   regEmail.trim().toLowerCase(),   // store email so admin can see it
            role:    "user",                          // always "user" — DB constraint safe
            company: regCompany.trim(),
            status:  "pending",
            avatar:  regName.trim().slice(0,2).toUpperCase(),
          });
          // Insert approval request so admin can see it
          await sb.insert("approval_requests", {
            user_id: userId,
            status: "pending",
            requested_at: new Date().toISOString(),
          });
        } catch (profileErr) {
          // Profile insert may fail if email not verified yet (Supabase policy)
          // The Supabase trigger / webhook should handle this on email confirmation
        }
      }
      setPendingReg({ name: regName.trim(), email: regEmail, company: regCompany.trim() });
      setVerifyStep(true);
    } catch (e) { setErr(e.message); }
    setLoading(false);
  };

  // ── Resend verification email ─────────────────────────────────────
  const handleResendVerification = async () => {
    setErr(""); setSuccess("");
    try {
      const res = await fetch(`${SUPABASE_URL}/auth/v1/resend`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON },
        body: JSON.stringify({ type: "signup", email: pendingReg.email }),
      });
      if (res.ok) setSuccess("Verification email resent!");
      else throw new Error("Failed to resend. Try again later.");
    } catch (e) { setErr(e.message); }
  };

  // ── Forgot Password ─────────────────────────────────────────────
  const handleForgotPassword = async () => {
    setErr(""); setSuccess("");
    if (!forgotEmail.includes("@")) return setErr("Enter a valid email address.");
    setLoading(true);
    try {
      const res = await fetch(`${SUPABASE_URL}/auth/v1/recover`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON },
        body: JSON.stringify({ email: forgotEmail }),
      });
      if (!res.ok) {
        const d = await res.json();
        throw new Error(d.error_description || d.message || "Failed to send reset email");
      }
      setForgotSent(true);
    } catch (e) { setErr(e.message); }
    setLoading(false);
  };

  // ── Password recovery screen (after clicking email reset link) ──
  if (recoveryToken) {
    return (
      <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:T.font }}>
        <div style={{ position:"absolute", inset:0, backgroundImage:`radial-gradient(ellipse at 30% 60%, ${T.accentDim}44 0%, transparent 55%)`, pointerEvents:"none" }} />
        <div className="fade-in" style={{ width:440, position:"relative" }}>
          <div style={{ textAlign:"center", marginBottom:24 }}>
            <div style={{ width:76, height:76, borderRadius:22, background:"linear-gradient(145deg, #1d4ed8, #2563eb, #7c3aed)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:36, margin:"0 auto 14px", boxShadow:"0 0 40px rgba(37,99,235,0.20)" }}>🔑</div>
            <h1 style={{ fontSize:24, fontWeight:900, letterSpacing:"-0.5px", marginBottom:4, background:"linear-gradient(135deg, #1d4ed8 40%, #7c3aed)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>Set New Password</h1>
            <p style={{ color:T.textMid, fontSize:13 }}>Choose a strong new password for your account.</p>
          </div>
          <Card style={{ padding:28 }}>
            {resetDone ? (
              <div style={{ textAlign:"center", padding:"10px 0" }}>
                <div style={{ fontSize:48, marginBottom:12 }}>✅</div>
                <p style={{ fontSize:15, fontWeight:700, color:T.green, marginBottom:8 }}>Password updated!</p>
                <p style={{ fontSize:13, color:T.textDim, marginBottom:24 }}>Your password has been changed successfully.</p>
                <Btn variant="primary" fullWidth onClick={() => { setRecoveryToken(null); setResetDone(false); setTab("login"); }}>
                  → Sign In Now
                </Btn>
              </div>
            ) : (
              <>
                {err && <div style={{ background:T.redDim, border:`1px solid ${T.red}44`, borderRadius:8, padding:"10px 14px", fontSize:12, color:T.red, marginBottom:14 }}>{err}</div>}
                <div style={{ marginBottom:16 }}>
                  <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>New Password</label>
                  <Input value={newPass} onChange={e=>setNewPass(e.target.value)} placeholder="Min 8 chars + number/symbol" prefix="" type="password" />
                </div>
                <div style={{ marginBottom:22 }}>
                  <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Confirm New Password</label>
                  <Input value={newPass2} onChange={e=>setNewPass2(e.target.value)} placeholder="Repeat new password" prefix="" type="password" />
                </div>
                <Btn onClick={handleSetNewPassword} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"✔"}>
                  {loading ? "Saving…" : "Set New Password"}
                </Btn>
              </>
            )}
          </Card>
        </div>
      </div>
    );
  }

  // ── Pending screen ───────────────────────────────────────────────
  if (pendingUser) {
    return (
      <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:T.font }}>
        <div style={{ position:"absolute", inset:0, backgroundImage:`radial-gradient(ellipse at 30% 60%, ${T.amberDim}44 0%, transparent 55%)`, pointerEvents:"none" }} />
        <div className="fade-in" style={{ width:460, position:"relative", textAlign:"center" }}>
          <div style={{ width:72, height:72, borderRadius:"50%", background:T.amberDim, border:`2px solid ${T.amber}55`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:32, margin:"0 auto 20px" }}>⏳</div>
          <h2 style={{ fontSize:22, fontWeight:700, color:T.text, marginBottom:8 }}>Awaiting Approval</h2>
          <p style={{ color:T.textDim, fontSize:14, lineHeight:1.7, marginBottom:24 }}>
            Hi <strong style={{color:T.text}}>{pendingUser.name}</strong>, your account has been created.<br/>
            An <span style={{color:T.amber}}>admin</span> needs to approve your access before you can log in.
          </p>
          <Card style={{ padding:20, marginBottom:16, textAlign:"left" }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:12 }}>
              <div style={{ width:36, height:36, borderRadius:"50%", background:T.accentDim, color:T.accent, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:700 }}>
                {pendingUser.name?.slice(0,2).toUpperCase()}
              </div>
              <div>
                <div style={{ fontSize:13, fontWeight:600, color:T.text }}>{pendingUser.name}</div>
                <div style={{ fontSize:11, color:T.textDim }}>{pendingUser.email}</div>
              </div>
              <Pill color="amber" size="xs" style={{marginLeft:"auto"}}>Pending</Pill>
            </div>
            <p style={{ fontSize:12, color:T.textDim }}>You'll be notified once approved. Please contact your admin if this takes too long.</p>
          </Card>
          <Btn variant="secondary" onClick={()=>{setPendingUser(null); setTab("login"); setEmail(pendingUser.email); setPass("");}} fullWidth icon="←">
            Back to Login
          </Btn>
        </div>
      </div>
    );
  }

  const BG = `radial-gradient(ellipse at 20% 50%, ${T.accentDim}55 0%, transparent 60%), radial-gradient(ellipse at 80% 20%, ${T.purpleDim}44 0%, transparent 50%)`;

  return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", fontFamily:T.font }}>
      <div style={{ position:"absolute", inset:0, backgroundImage:BG, pointerEvents:"none" }} />
      <div className="fade-in" style={{ width:440, position:"relative" }}>
        {/* Logo */}
        {/* ── Brand Header ── */}
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ position:"relative", display:"inline-block", marginBottom:14 }}>
            <div style={{ width:76, height:76, borderRadius:22, background:"linear-gradient(145deg, #1d4ed8, #2563eb, #7c3aed)", display:"flex", alignItems:"center", justifyContent:"center", boxShadow:"0 0 0 1px rgba(37,99,235,0.2), 0 0 40px rgba(37,99,235,0.15), 0 8px 24px rgba(0,0,0,0.12)" }}>
                <svg width="38" height="38" viewBox="0 0 38 38" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect x="4" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                  <rect x="10" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                  <rect x="16" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                  <rect x="22" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                  <rect x="28" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                  <rect x="2" y="31" width="33" height="3" rx="1.5" fill="white"/>
                  <polygon points="19,4 2,15 36,15" fill="white" fillOpacity="0.95"/>
                </svg>
              </div>
          </div>
          <h1 style={{ fontSize:28, fontWeight:900, letterSpacing:"-1px", marginBottom:3, background:"linear-gradient(135deg, #1d4ed8 40%, #7c3aed)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>Bank2Tally</h1>
          <p style={{ color:T.textMid, fontSize:11, letterSpacing:"0.12em", textTransform:"uppercase", fontWeight:600, marginBottom:14 }}>Professional Bank Statement Importer</p>
          <div style={{ background:`linear-gradient(135deg, rgba(61,127,255,0.08), rgba(180,124,255,0.08))`, border:`1px solid ${T.border}`, borderRadius:12, padding:"10px 14px", textAlign:"left" }}>
            <p style={{ fontSize:12, color:T.textMid, lineHeight:1.65, fontStyle:"italic", marginBottom:4 }}>"{todayQuote.text}"</p>
            <p style={{ fontSize:10, color:T.textDim, fontWeight:600, letterSpacing:"0.05em" }}>— {todayQuote.author}</p>
          </div>
        </div>

        {/* Tab switcher - only show when not on forgot */}
        {tab !== "forgot" && (
          <div style={{ display:"flex", background:T.surface, borderRadius:11, padding:4, marginBottom:20, border:`1px solid ${T.border}` }}>
            {[["login","Sign In"],["register","Request Access"]].map(([t,label]) => (
              <button key={t} onClick={()=>{setTab(t);setErr("");setSuccess("");}}
                style={{ flex:1, padding:"8px 0", borderRadius:8, border:"none", cursor:"pointer", fontFamily:T.font, fontSize:13, fontWeight:tab===t?600:400, transition:"all 0.2s",
                  background:tab===t?T.accent:"transparent", color:tab===t?"#fff":T.textMid,
                  boxShadow:tab===t?`0 0 16px ${T.accentGlow}`:"none" }}>
                {label}
              </button>
            ))}
          </div>
        )}

        <Card style={{ padding:28 }}>
          {err && (
            <div style={{ background:T.redDim, border:`1px solid ${T.red}44`, borderRadius:8, padding:"10px 14px", fontSize:12, color:T.red, marginBottom:16 }}>
              ✕ {err}
            </div>
          )}
          {success && (
            <div style={{ background:T.greenDim, border:`1px solid ${T.green}44`, borderRadius:8, padding:"10px 14px", fontSize:12, color:T.green, marginBottom:16 }}>
              ✓ {success}
            </div>
          )}

          {tab === "forgot" ? (
            <>
              <div style={{ textAlign:"center", marginBottom:18 }}>
                <div style={{ fontSize:32, marginBottom:8 }}>🔑</div>
                <h3 style={{ fontSize:16, fontWeight:700, color:T.text, marginBottom:4 }}>Reset your password</h3>
                <p style={{ fontSize:12, color:T.textDim }}>Enter your registered email and we'll send you a reset link.</p>
              </div>
              {forgotSent ? (
                <div style={{ textAlign:"center", padding:"20px 0" }}>
                  <div style={{ fontSize:40, marginBottom:12 }}>📬</div>
                  <p style={{ fontSize:13, color:T.green, fontWeight:600, marginBottom:8 }}>Reset link sent!</p>
                  <p style={{ fontSize:12, color:T.textDim, marginBottom:20 }}>Check your inbox at <strong style={{color:T.text}}>{forgotEmail}</strong> and follow the link to reset your password.</p>
                  <Btn variant="secondary" onClick={()=>{setTab("login");setForgotSent(false);setForgotEmail("");setErr("");}} fullWidth>
                    ← Back to Sign In
                  </Btn>
                </div>
              ) : (
                <>
                  <div style={{ marginBottom:18 }}>
                    <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Registered email address</label>
                    <Input value={forgotEmail} onChange={e=>setForgotEmail(e.target.value)} placeholder="you@company.in" prefix="✉" />
                  </div>
                  <Btn onClick={handleForgotPassword} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"→"}>
                    {loading ? "Sending…" : "Send Reset Link"}
                  </Btn>
                  <div style={{ marginTop:14, textAlign:"center" }}>
                    <button onClick={()=>{setTab("login");setErr("");setForgotEmail("");}}
                      style={{ background:"none", border:"none", cursor:"pointer", color:T.textDim, fontSize:12, fontFamily:T.font }}>
                      ← Back to Sign In
                    </button>
                  </div>
                </>
              )}
            </>
          ) : verifyStep ? (
            <>
              <div style={{ textAlign:"center", padding:"10px 0 20px" }}>
                <div style={{ width:52, height:52, borderRadius:"50%", background:T.accentDim, display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 12px", border:`1px solid ${T.accent}33` }}><div style={{ fontSize:10, fontWeight:700, color:T.accent }}>EMAIL</div></div>
                <h3 style={{ fontSize:16, fontWeight:700, color:T.text, marginBottom:8 }}>Check your email</h3>
                <p style={{ fontSize:13, color:T.textMid, lineHeight:1.7, marginBottom:4 }}>
                  We sent a verification link to:
                </p>
                <p style={{ fontSize:14, fontWeight:700, color:T.accent, marginBottom:16 }}>
                  {pendingReg?.email}
                </p>
                <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"14px 16px", textAlign:"left", marginBottom:20 }}>
                  <p style={{ fontSize:12, color:T.textMid, lineHeight:1.8 }}>
                    1. Open your email inbox<br/>
                    2. Click the <strong style={{color:T.text}}>verification link</strong> from Bank2Tally<br/>
                    3. Come back here and <strong style={{color:T.text}}>Sign In</strong>
                  </p>
                </div>
                <Btn onClick={handleResendVerification} variant="secondary" fullWidth icon="→">
                  Resend verification email
                </Btn>
                <div style={{ marginTop:14 }}>
                  <button onClick={()=>{setVerifyStep(false);setTab("login");setErr("");setSuccess("");}}
                    style={{ background:"none", border:"none", cursor:"pointer", color:T.textDim, fontSize:12, fontFamily:T.font }}>
                    ← Back to Sign In
                  </button>
                </div>
              </div>
            </>
          ) : tab === "login" ? (
            <>
              <div style={{ marginBottom:14 }}>
                <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Email address</label>
                <Input value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@company.in" prefix="✉" />
              </div>
              <div style={{ marginBottom:22 }}>
                <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Password</label>
                <Input value={pass} onChange={e=>setPass(e.target.value)} placeholder="••••••••" prefix="" />
              </div>
              <Btn onClick={handleLogin} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"→"}>
                {loading ? "Signing in…" : "Sign In"}
              </Btn>
              <div style={{ marginTop:14, textAlign:"center" }}>
                <button onClick={()=>{setTab("forgot");setErr("");}}
                  style={{ background:"none", border:"none", cursor:"pointer", color:T.accent, fontSize:12, fontFamily:T.font, textDecoration:"underline" }}>
                  Forgot password?
                </button>
              </div>
            </>
          ) : (
            <>
              <div style={{ background:T.accentDim+"55", border:`1px solid ${T.accent}33`, borderRadius:9, padding:"10px 14px", fontSize:12, color:T.textMid, marginBottom:18 }}>
                ℹ After registering, your account will be <strong style={{color:T.accent}}>reviewed by an admin</strong> before you can log in.
              </div>
              {[
                { label:"Full Name *", val:regName, set:setRegName, ph:"Rajesh Kumar", pre:"👤" },
                { label:"Work Email *", val:regEmail, set:setRegEmail, ph:"rajesh@company.in", pre:"✉" },
                { label:"Company", val:regCompany, set:setRegCompany, ph:"Acme Corp Pvt Ltd", pre:"🏢" },
                { label:"Password * (min 8 chars)", val:regPass, set:setRegPass, ph:"••••••••", pre:"🔒" },
                { label:"Confirm Password *", val:regPass2, set:setRegPass2, ph:"••••••••", pre:"🔒" },
              ].map(f => (
                <div key={f.label} style={{ marginBottom:12 }}>
                  <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>{f.label}</label>
                  <Input value={f.val} onChange={e=>f.set(e.target.value)} placeholder={f.ph} prefix={f.pre} />
                </div>
              ))}
              <div style={{ marginTop:6 }}>
                <Btn onClick={handleRegister} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"→"}>
                  {loading ? "Submitting…" : "Request Access"}
                </Btn>
              </div>
            </>
          )}
        </Card>

        {/* ── Verma Consultancy Branding Footer ── */}
        <div style={{ textAlign:"center", marginTop:22, paddingTop:16, borderTop:`1px solid ${T.border}` }}>
          <div style={{ marginBottom:6 }}>
            <span style={{ fontSize:11, fontWeight:600, color:T.textSub, letterSpacing:"0.04em", textTransform:"uppercase" }}>Produced by</span>
            <span style={{ fontSize:12, fontWeight:800, color:T.text, letterSpacing:"0.02em", marginLeft:6 }}>Verma Consultancy Services</span>
          </div>
          <p style={{ fontSize:11, color:T.textDim, marginBottom:8 }}>For purchase, support &amp; enquiries</p>
          <div style={{ display:"flex", justifyContent:"center", gap:12, flexWrap:"wrap" }}>
            <a href="tel:+918707401846"
              style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 14px", background:T.accentSoft, border:`1px solid ${T.accent}44`, borderRadius:20, fontSize:11, fontWeight:600, color:T.accent, textDecoration:"none" }}>
              +91 8707401846
            </a>
            <a href="mailto:svtiger543939@gmail.com"
              style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 14px", background:T.accentSoft, border:`1px solid ${T.accent}44`, borderRadius:20, fontSize:11, fontWeight:600, color:T.accent, textDecoration:"none" }}>
              svtiger543939@gmail.com
            </a>
          </div>
          <p style={{ fontSize:10, color:T.textDim, marginTop:10, letterSpacing:"0.04em" }}>v2.0 Commercial · Tally ERP 9 &amp; Prime Compatible</p>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// ADMIN PANEL: Approval Requests
// ══════════════════════════════════════════════════════════════════
function AdminApprovalPanel({ user, onClose }) {
  const [requests, setRequests] = useState([]);
  const [loading, setLoading]   = useState(true);
  const [actioning, setActioning] = useState(null);
  const [toast_, setToast_] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      // Fetch approval requests only (no join — avoids RLS recursion)
      const rows = await sb.from(
        "approval_requests",
        "status=eq.pending&select=*&order=requested_at.asc"
      );
      // Enrich each with profile data separately
      const enriched = await Promise.all(rows.map(async (req) => {
        try {
          const pr = await sb.from("profiles", `id=eq.${req.user_id}&select=name,company,avatar`);
          return { ...req, profiles: pr?.[0] || {} };
        } catch { return { ...req, profiles: {} }; }
      }));
      setRequests(enriched);
    } catch (e) { setToast_("Error: " + e.message); }
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const action = async (req, decision) => {
    setActioning(req.id);
    try {
      // Update approval_request
      await sb.update("approval_requests", { id: req.id }, {
        status: decision, admin_id: user.id, actioned_at: new Date().toISOString(),
      });
      // Update profile status
      await sb.update("profiles", { id: req.user_id }, {
        status: decision,
        approved_by: user.id,
        approved_at: new Date().toISOString(),
      });
      setToast_(`${decision === "approved" ? "✓ Approved" : "✕ Rejected"}: ${req.profiles?.name}`);
      setRequests(r => r.filter(x => x.id !== req.id));
    } catch (e) { setToast_("Error: " + e.message); }
    setActioning(null);
  };

  return (
    <div>
      {toast_ && (
        <div style={{ background:T.greenDim, border:`1px solid ${T.green}44`, borderRadius:8, padding:"9px 14px", fontSize:12, color:T.green, marginBottom:14 }}>
          {toast_}
        </div>
      )}
      {loading ? (
        <div style={{ textAlign:"center", padding:"40px 0", color:T.textDim }}>
          <span style={{ animation:"pulse 1.5s infinite", fontSize:24 }}>⏳</span>
          <p style={{ marginTop:10, fontSize:13 }}>Loading requests…</p>
        </div>
      ) : requests.length === 0 ? (
        <div style={{ textAlign:"center", padding:"40px 0", color:T.textDim }}>
          <div style={{ fontSize:36, marginBottom:10 }}>✅</div>
          <p style={{ fontSize:14 }}>No pending approval requests</p>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {requests.map(req => (
            <div key={req.id} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, padding:"16px 18px", display:"flex", justifyContent:"space-between", alignItems:"center", gap:12 }}>
              <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                <div style={{ width:40, height:40, borderRadius:"50%", background:T.accentDim, color:T.accent, display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, fontWeight:700, flexShrink:0 }}>
                  {req.profiles?.name?.slice(0,2).toUpperCase() || "?"}
                </div>
                <div>
                  <div style={{ fontSize:13, fontWeight:600, color:T.text }}>{req.profiles?.name}</div>
                  <div style={{ fontSize:11, color:T.textDim }}>{req.profiles?.email || "—"}</div>
                  {req.profiles?.company && <div style={{ fontSize:11, color:T.textMid, marginTop:2 }}>🏢 {req.profiles.company}</div>}
                  <div style={{ fontSize:10, color:T.textDim, marginTop:2 }}>
                    Requested {new Date(req.requested_at).toLocaleString("en-IN")}
                  </div>
                </div>
              </div>
              <div style={{ display:"flex", gap:8, flexShrink:0 }}>
                <Btn size="sm" variant="success" icon="✓"
                  disabled={actioning === req.id}
                  onClick={() => action(req, "approved")}>
                  {actioning === req.id ? "…" : "Approve"}
                </Btn>
                <Btn size="sm" variant="danger" icon="✕"
                  disabled={actioning === req.id}
                  onClick={() => action(req, "rejected")}>
                  Reject
                </Btn>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}



// ══════════════════════════════════════════════════════════════════
// SCREEN: Dashboard
// ══════════════════════════════════════════════════════════════════
function DashboardScreen({ history, setScreen, user, tally }) {
  const totalRows = history.reduce((s,h) => s+h.rows, 0);
  const totalSuspense = history.reduce((s,h) => s+(h.suspense||0), 0);
  const totalDups = history.reduce((s,h) => s+(h.duplicates||0), 0);
  const thisMonth = history.filter(h => {
    const d = new Date(h.rawDate || h.date);
    const n = new Date();
    return d.getMonth() === n.getMonth() && d.getFullYear() === n.getFullYear();
  });

  return (
    <div className="fade-in">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:24 }}>
        <div>
          <div>
            <h2 style={{ fontSize:22, fontWeight:900, letterSpacing:"-0.6px", color:T.text }}>Dashboard</h2>
            <p style={{ color:T.textMid, fontSize:13, marginTop:3 }}>Welcome back, <strong style={{color:T.text}}>{user?.name?.split(" ")[0]}</strong> 👋</p>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          {tally?.status === "ok" && <Pill color="green" dot>Tally Live · {tally.companies.length} co.</Pill>}
          {tally?.status === "connecting" && <Pill color="amber" dot>Connecting…</Pill>}
          {tally?.status === "idle" && <Pill color="amber" dot>Waiting for extension…</Pill>}
          {tally?.status === "error" && (
            <span style={{display:"flex",alignItems:"center",gap:6}}>
              <Pill color="red" dot>Tally Offline</Pill>
              <button onClick={()=>tally.refetch()} title="Retry connection"
                style={{fontSize:13,background:"none",border:"none",cursor:"pointer",color:T.textDim,padding:"2px 4px"}}>↺</button>
            </span>
          )}
          <Btn onClick={() => setScreen(SCREENS.UPLOAD)} icon="+" size="lg">New Import</Btn>
        </div>
      </div>

      {/* Daily Motivation */}
      <div style={{ background:"rgba(37,99,235,0.04)", border:"1px solid rgba(37,99,235,0.14)", borderRadius:12, padding:"12px 18px", marginBottom:18, display:"flex", alignItems:"center", gap:14 }}>
        <div style={{ width:3, alignSelf:"stretch", borderRadius:2, background:T.accent, flexShrink:0 }} />
        <div>
          <p style={{ fontSize:12, color:T.textMid, lineHeight:1.6, fontStyle:"italic", marginBottom:2 }}>"{todayQuote.text}"</p>
          <p style={{ fontSize:10, color:T.textSub, fontWeight:600 }}>— {todayQuote.author}</p>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:24 }}>
        <StatCard icon="📑" label="Total Transactions" value={totalRows.toLocaleString()} sub={`${history.length} imports`} color={T.accent} />
        <StatCard icon="📅" label="This Month" value={thisMonth.length} sub="imports" color={T.purple} />
        <StatCard icon="⚠️" label="Suspense Pending" value={totalSuspense} sub="need review" color={T.amber} />
        <StatCard icon="🔁" label="Duplicates Found" value={totalDups} sub="auto-detected" color={T.red} />
      </div>

      {/* Recent activity */}
      <Card>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
          <span style={{ fontWeight:600, fontSize:14 }}>Recent Imports</span>
          <Btn size="sm" variant="ghost" onClick={()=>setScreen(SCREENS.HISTORY)}>View all →</Btn>
        </div>
        {history.length === 0 ? (
          <div style={{ textAlign:"center", padding:"40px 0", color:T.textDim }}>
            <div style={{ width:48, height:48, borderRadius:12, background:T.accentDim, margin:"0 auto 12px", display:"flex", alignItems:"center", justifyContent:"center" }}>
              <div style={{ width:18, height:22, borderRadius:3, border:`2px solid ${T.accent}`, opacity:0.5 }} />
            </div>
            <p style={{ fontSize:14, color:T.textSub, fontWeight:500 }}>No imports yet — start your first import</p>
          </div>
        ) : history.slice(0,5).map(h => (
          <div key={h.id} className="row-hover" style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"11px 14px", borderRadius:9, marginBottom:4, cursor:"pointer", transition:"background 0.15s" }}>
            <div style={{ display:"flex", alignItems:"center", gap:12 }}>
              <div style={{ width:36, height:36, borderRadius:9, background:T.accentDim, display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, color:T.accent, flexShrink:0 }}>
                {h.filename?.split(".").pop().toUpperCase().slice(0,3)||"XLS"}
              </div>
              <div>
                <div style={{ fontWeight:500, fontSize:13, color:T.text }}>{h.filename}</div>
                <div style={{ fontSize:11, color:T.textDim }}>{h.date} · {h.company}</div>
              </div>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <Pill color="blue" size="xs">{h.rows} rows</Pill>
              {h.suspense > 0 && <Pill color="amber" size="xs">{h.suspense} suspense</Pill>}
              <Pill color="green" size="xs" dot>Imported</Pill>
            </div>
          </div>
        ))}
      </Card>
    </div>
  );
}

// Small helper for offline/manual company entry
function ManualCompanyEntry({ onAdd }) {
  const [val, setVal] = useState("");
  return (
    <div style={{ display:"flex", gap:8, flex:1 }}>
      <Input value={val} onChange={e=>setVal(e.target.value)}
        placeholder="e.g. Acme Corp Pvt Ltd" prefix="🏢"
        style={{ flex:1 }} />
      <Btn size="sm" variant="secondary" icon="+" onClick={()=>{ if(val.trim()){ onAdd(val.trim()); setVal(""); } }}>Add</Btn>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Upload
// ══════════════════════════════════════════════════════════════════
function UploadScreen({ onParsed, selectedCompanies, setSelectedCompanies, tally }) {
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [pdfStatus, setPdfStatus] = useState("");
  const [error, setError] = useState(null);
  const [selectedTemplate, setSelectedTemplate] = useState("");
  const fileRef = useRef();

  const handleFile = useCallback(async (file) => {
    // 50MB limit — large PDFs crash the browser tab
    if (file.size > 50 * 1024 * 1024) {
      setError({ code: "ERR_001", message: `File too large (${(file.size/1024/1024).toFixed(1)} MB). Maximum is 50 MB. For large bank statements, split the PDF into smaller date ranges or export as Excel/CSV.` });
      return;
    }
    setError(null); setLoading(true); setPdfStatus("");
    try {
      const result = await parseFile(file, (msg) => setPdfStatus(msg));
      onParsed(result, file.name, selectedTemplate);
    } catch(e) {
      setError({ code: e.code || "ERR_001", message: e.message });
    } finally { setLoading(false); setPdfStatus(""); }
  }, [onParsed, selectedTemplate]);

  const onDrop = e => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if(f) handleFile(f); };

  // Tally connection banner
  const StatusBadge = () => {
    if (tally.status === "connecting") return <Pill color="amber" dot>Connecting to Tally…</Pill>;
    if (tally.status === "ok") return <Pill color="green" dot>Tally Live · {tally.companies.length} companies</Pill>;
    if (tally.status === "error") return <Pill color="red" dot>Tally offline — using defaults</Pill>;
    return null;
  };

  const displayCompanies = tally.companies.length > 0 ? tally.companies : TALLY_COMPANIES_FALLBACK;

  return (
    <div className="fade-in">
      <h2 style={{ fontSize:20, fontWeight:700, marginBottom:4, color:T.text }}>Import Bank Statement</h2>
      <p style={{ color:T.textDim, fontSize:13, marginBottom:20 }}>Upload your statement — CSV, Excel (.xlsx/.xls) or PDF (text + OCR supported)</p>

      {/* Company multi-select */}
      <Card style={{ marginBottom:20, padding:18 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
          <p style={{ fontWeight:600, fontSize:13, color:T.text }}>Target Tally Companies <span style={{color:T.red}}>*</span></p>
          <StatusBadge />
        </div>

        {tally.status === "error" && (
          <div style={{ background:T.amberDim, border:`1px solid ${T.amber}44`, borderRadius:8, padding:"9px 14px", fontSize:12, color:T.amber, marginBottom:12 }}>
            ⚠ {tally.error?.includes("manually") 
              ? "Tally is connected but company list could not be read automatically. Type your company name below."
              : `Could not reach Tally at localhost:9000. Go to Settings → Test Connection to reconnect, or type a company name below.`}
          </div>
        )}

        {tally.status === "connecting" && (
          <div style={{ display:"flex", gap:10, alignItems:"center", padding:"12px 0", color:T.textDim, fontSize:13 }}>
            <span style={{ animation:"pulse 1.5s infinite", fontSize:18 }}>⏳</span> Fetching companies from Tally gateway…
          </div>
        )}

        {(() => {
          // Merge Tally companies + manually-added ad-hoc companies
          const adHoc = Object.values(window.__adHocCompanies || {});
          const allCos = [...displayCompanies, ...adHoc.filter(a => !displayCompanies.find(d => d.id === a.id))];
          if (!allCos.length) return null;
          return (
            <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
              {allCos.map(c => {
                const sel = selectedCompanies.includes(c.id);
                const isAdHoc = !!(window.__adHocCompanies||{})[c.id];
                return (
                  <div key={c.id} style={{ display:"flex", alignItems:"center", gap:0, borderRadius:9, overflow:"hidden",
                    border: sel ? `2px solid ${isAdHoc ? T.green : T.accent}` : `1px solid ${T.border}`,
                    background: sel ? (isAdHoc ? T.greenDim : T.accentDim) : T.surface,
                    boxShadow: sel ? `0 0 0 3px ${isAdHoc ? T.green : T.accent}18` : "none",
                    transition:"all 0.15s" }}>
                    <button onClick={() => setSelectedCompanies(p => sel ? p.filter(x=>x!==c.id) : [...p,c.id])}
                      style={{ padding:"7px 12px", fontSize:12, fontWeight:sel?600:500, cursor:"pointer", border:"none", background:"transparent",
                        color: sel ? (isAdHoc ? T.green : T.accent) : T.textMid, fontFamily:T.font }}>
                      {sel && <span style={{marginRight:4}}>✓</span>}{c.name}
                      {isAdHoc && <span style={{ marginLeft:5, fontSize:9, opacity:0.65, fontWeight:400 }}>manual</span>}
                      {c.state && !isAdHoc && <span style={{ marginLeft:5, fontSize:10, opacity:0.6 }}>· {c.state}</span>}
                    </button>
                    {isAdHoc && (
                      <button title="Remove company" onClick={() => {
                        delete (window.__adHocCompanies||{})[c.id];
                        setSelectedCompanies(p => p.filter(x=>x!==c.id));
                      }} style={{ padding:"7px 9px 7px 4px", fontSize:13, border:"none", background:"transparent", cursor:"pointer",
                        color: T.red, fontFamily:T.font, lineHeight:1, fontWeight:700 }}>×</button>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}

        {/* Manual entry if Tally offline */}
        {tally.status === "error" && (
          <div style={{ marginTop:12 }}>
            <p style={{ fontSize:11, color:T.textDim, marginBottom:6 }}>Or type a company name manually:</p>
            <div style={{ display:"flex", gap:8 }}>
              <ManualCompanyEntry onAdd={(name) => {
                const id = `mc_${genId()}`;
                // We can't mutate tally.companies directly, so keep in selectedCompanies as a name-keyed entry
                setSelectedCompanies(p => [...p, id]);
                // Store ad-hoc company in session (pass through as-is)
                window.__adHocCompanies = window.__adHocCompanies || {};
                window.__adHocCompanies[id] = { id, name, gstin:"", state:"", fy:"2024-25" };
              }} />
            </div>
          </div>
        )}

        {selectedCompanies.length === 0 && <p style={{ marginTop:8, fontSize:11, color:T.red }}>⚠ Select at least one company</p>}
      </Card>

      {/* Bank template */}
      <Card style={{ marginBottom:20, padding:18 }}>
        <p style={{ fontWeight:600, fontSize:13, marginBottom:10, color:T.text }}>Bank Template <span style={{ fontWeight:400, color:T.textDim }}>(optional — auto-maps columns)</span></p>
        <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
          {Object.entries(BANK_TEMPLATES).map(([key, t]) => (
            <button key={key} onClick={() => setSelectedTemplate(selectedTemplate===key ? "" : key)}
              style={{ padding:"6px 14px", borderRadius:8, fontSize:12, cursor:"pointer", fontFamily:T.font, transition:"all 0.15s",
                border: selectedTemplate===key ? `2px solid ${T.green}` : `1px solid ${T.border}`,
                background: selectedTemplate===key ? T.greenDim : T.surface,
                color: selectedTemplate===key ? T.green : T.textMid }}>
              {selectedTemplate===key && "✓ "}{t.name}
            </button>
          ))}
        </div>
      </Card>

      {error && <ErrCard code={error.code} message={error.message} onDismiss={() => setError(null)} />}

      {/* Drop zone */}
      <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={onDrop}
        onClick={()=>fileRef.current?.click()}
        style={{ border:`2px dashed ${dragging?T.accent:T.border}`, borderRadius:16, padding:"52px 24px", textAlign:"center", cursor:"pointer",
          background: dragging ? T.accentDim+"33" : T.surface,
          boxShadow: dragging ? `0 0 30px ${T.accentGlow}` : "none",
          transition:"all 0.25s" }}>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.pdf" style={{display:"none"}} onChange={e=>e.target.files[0]&&handleFile(e.target.files[0])} />
        {loading ? (
          <div>
            <div style={{ fontSize:32, marginBottom:12, animation:"pulse 1.5s infinite" }}>⏳</div>
            <p style={{ color:T.textMid, fontSize:14 }}>{pdfStatus || "Parsing file…"}</p>
          </div>
        ) : (
          <>
            <div style={{ marginBottom:14, display:"flex", justifyContent:"center" }}>
                <svg width="48" height="48" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect x="6" y="28" width="36" height="4" rx="2" fill={T.border}/>
                  <rect x="6" y="36" width="36" height="4" rx="2" fill={T.border}/>
                  <path d="M24 6 L24 26 M24 6 L16 14 M24 6 L32 14" stroke={T.accent} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
            <p style={{ fontWeight:700, fontSize:16, color:T.text, marginBottom:6 }}>Drop bank statement here</p>
            <p style={{ color:T.textDim, fontSize:13 }}>or <span style={{color:T.accent, textDecoration:"underline"}}>browse files</span></p>
            <div style={{ display:"flex", gap:8, justifyContent:"center", marginTop:16, flexWrap:"wrap" }}>
              {["CSV", ".xlsx / .xls", "PDF (text + OCR)"].map(l => (
                <span key={l} style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:6, padding:"3px 10px", fontSize:11, color:T.textDim }}>{l}</span>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Column Mapper
// ══════════════════════════════════════════════════════════════════
const REQUIRED_COLS = ["date","narration","debit","credit"];
const COL_LABELS = {
  date:     "Date",
  narration:"Narration / Description",
  debit:    "Debit Amount",
  credit:   "Credit Amount",
  crdr:     "Combined Cr/Dr Amount (auto-split)",
  crdrFlag: "Cr/Dr Indicator column (optional)",
  balance:  "Balance (optional)",
  ref:      "Reference / Cheque No. (optional)",
};

function ColumnMapScreen({ headers, templateKey, onMapped, onBack }) {
  const [mapping, setMapping] = useState({});
  const [draggingCol, setDraggingCol] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  const autoMap = useCallback(() => {
    const m = {};

    // ── If a template is selected, try exact+normalised match first ──
    if (templateKey && BANK_TEMPLATES[templateKey]) {
      const tpl = BANK_TEMPLATES[templateKey].cols;
      const norm = s => s.toLowerCase().replace(/[\s_\-\.\(\)\/,]/g,"");
      Object.entries(tpl).forEach(([field, colName]) => {
        if (m[field]) return;
        const cn = norm(colName);
        // 1. Exact normalised match
        const exact = headers.find(h => norm(h) === cn);
        if (exact) { m[field] = exact; return; }
        // 2. Starts-with (handles trailing spaces/suffixes in PDF-extracted headers)
        const sw = headers.find(h => norm(h).startsWith(cn.slice(0,12)) || cn.startsWith(norm(h).slice(0,12)));
        if (sw) { m[field] = sw; return; }
        // 3. Contains key word
        const kw = colName.toLowerCase().split(/\s+/).find(w=>w.length>4);
        if (kw) { const has = headers.find(h=>h.toLowerCase().includes(kw)); if(has) m[field]=has; }
      });
    }

    // ── Universal fuzzy pass — fills any unmapped fields ─────────────
    // This works WITHOUT any template by pattern-matching column names
    const norm2 = s => s.toLowerCase().replace(/[\s_\-\.\(\)\/,]/g,"");
    headers.forEach(h => {
      const hn = norm2(h);

      // DATE — must contain 'date' and not be a narration field
      if (!m.date && /date/.test(hn) && !/narr|desc|particular|remark/.test(hn)) m.date = h;

      // NARRATION — description / particulars / narration / remarks
      if (!m.narration && /narr|description|particulars|detail|remark|transactionremark/.test(hn)) m.narration = h;

      // DEBIT — withdrawal / debit (not combined)
      const isCombined = /^(amount|txnamount|crdr|drcr|cr\/dr|dr\/cr)$/.test(hn) || (hn.includes("cr")&&hn.includes("dr")&&!hn.includes("credit")&&!hn.includes("debit"));
      if (!m.debit && !isCombined && (/debit|withdraw|paid/.test(hn) || /\(dr\)$/.test(hn))) m.debit = h;

      // CREDIT — deposit / credit (not combined)
      if (!m.credit && !isCombined && (/credit|deposit|received/.test(hn) || /\(cr\)$/.test(hn))) m.credit = h;

      // BALANCE
      if (!m.balance && /balance|bal(?!ance)?$/.test(hn)) m.balance = h;

      // REF
      if (!m.ref && /ref|chq|cheque|utr|neft|imps|instrument|sno|serial/.test(hn)) m.ref = h;
    });

    // ── Combined Cr/Dr detection (for banks with single amount column) ──
    if (!m.debit || !m.credit) {
      const combined = headers.find(h => {
        const hn = norm2(h);
        return /^(amount|txnamount|netamount)$/.test(hn)
          || (/amount/.test(hn) && !/(debit|credit|withdraw|deposit|balance)/.test(hn))
          || (hn.includes("cr")&&hn.includes("dr"));
      });
      if (combined && !m.crdr) {
        m.crdr = combined;
        // Look for a Cr/Dr flag column
        const flag = headers.find(h => {
          const hn = norm2(h);
          return /type|flag|indicator|crdr|drorflag/.test(hn) && h !== combined;
        });
        if (flag) m.crdrFlag = flag;
      }
    }

    setMapping(m);
  }, [headers, templateKey]);

  useEffect(() => { autoMap(); }, [autoMap]);

  const assignedSet = new Set(Object.values(mapping).filter(Boolean));
  // Satisfied if: date+narration + (debit+credit) OR (crdr combined column)
  const hasCrDr = !!(mapping.crdr);
  const allRequired = mapping.date && mapping.narration && (
    (mapping.debit && mapping.credit) || hasCrDr
  );

  const dropOn = field => {
    if (!draggingCol) return;
    const prev = mapping[field];
    setMapping(m => ({ ...m, [field]: draggingCol }));
    setDraggingCol(null); setDragOver(null);
  };

  const unassign = field => {
    setMapping(m => { const n={...m}; delete n[field]; return n; });
  };

  return (
    <div className="fade-in">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
        <div>
          <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:4 }}>Map Columns</h2>
          <p style={{ color:T.textDim, fontSize:13 }}>Drag source columns onto the target fields. Required fields marked <span style={{color:T.red}}>*</span></p>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          {templateKey && <Pill color="green">Template: {BANK_TEMPLATES[templateKey]?.name}</Pill>}
          <Btn size="sm" variant="secondary" onClick={autoMap} icon="✨">Auto-detect</Btn>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 230px", gap:20, alignItems:"start" }}>
        {/* Target fields */}
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {Object.entries(COL_LABELS).map(([field, label]) => {
            const req = REQUIRED_COLS.includes(field);
            const mapped = mapping[field];
            const isOver = dragOver === field;
            return (
              <div key={field}
                onDragOver={e=>{e.preventDefault();setDragOver(field);}}
                onDragLeave={()=>setDragOver(null)}
                onDrop={()=>dropOn(field)}
                style={{ display:"flex", alignItems:"center", gap:12, padding:"12px 16px", borderRadius:11,
                  border: isOver ? `2px solid ${T.accent}` : mapped ? `1px solid ${T.green}44` : `1px dashed ${T.border}`,
                  background: isOver ? T.accentDim+"44" : mapped ? T.greenDim+"33" : T.surface,
                  transition:"all 0.15s", boxShadow: isOver ? `0 0 16px ${T.accentGlow}` : "none" }}>
                <div style={{ minWidth:170 }}>
                  <span style={{ fontSize:13, fontWeight:500, color:T.text }}>{label}</span>
                  {req && <span style={{ color:T.red, marginLeft:3 }}>*</span>}
                </div>
                {mapped ? (
                  <div style={{ display:"flex", alignItems:"center", gap:8, flex:1 }}>
                    <span style={{ background:T.greenDim, color:T.green, padding:"3px 12px", borderRadius:7, fontSize:12, fontWeight:500, fontFamily:T.mono }}>{mapped}</span>
                    <button onClick={()=>unassign(field)} style={{ background:"none", border:"none", cursor:"pointer", color:T.textDim, fontSize:16, lineHeight:1 }}>×</button>
                  </div>
                ) : (
                  <span style={{ fontSize:12, color:T.textDim, fontStyle:"italic" }}>← drop column here</span>
                )}
              </div>
            );
          })}
        </div>

        {/* Source columns */}
        <Card style={{ padding:14, position:"sticky", top:0 }}>
          <p style={{ fontSize:12, fontWeight:600, color:T.textMid, marginBottom:10 }}>FILE COLUMNS ({headers.length})</p>
          <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
            {headers.map(h => {
              const used = assignedSet.has(h);
              return (
                <div key={h} draggable={!used} onDragStart={()=>{ if(!used) setDraggingCol(h); }}
                  style={{ padding:"8px 11px", borderRadius:8, fontSize:12, fontWeight:500, fontFamily:T.mono,
                    background: used ? T.greenDim : T.surface,
                    color: used ? T.green : T.text,
                    border: `1px solid ${used ? T.green+"44" : T.border}`,
                    cursor: used ? "default" : "grab",
                    textDecoration: used ? "none" : "none",
                    opacity: used ? 0.7 : 1,
                    transition:"all 0.15s", userSelect:"none" }}>
                  {used ? "✓ " : ""}{h}
                </div>
              );
            })}
          </div>
        </Card>
      </div>

      {!allRequired && (
        <div style={{ background:T.amberDim, border:`1px solid ${T.amber}44`, borderRadius:9, padding:"10px 16px", marginTop:14, fontSize:12, color:T.amber }}>
          ⚠ Required: {!mapping.date && "Date, "}{!mapping.narration && "Narration, "}
          {!mapping.crdr && !mapping.debit && "Debit Amount, "}
          {!mapping.crdr && !mapping.credit && "Credit Amount"}
          {!mapping.crdr && !mapping.debit && !mapping.credit && " — or map a single Combined Cr/Dr Amount column"}
        </div>
      )}

      <div style={{ display:"flex", justifyContent:"space-between", marginTop:24 }}>
        <Btn variant="secondary" onClick={onBack} icon="←">Back</Btn>
        <Btn onClick={()=>onMapped(mapping)} disabled={!allRequired} icon="→">Assign Ledgers</Btn>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Ledger Assignment
// ══════════════════════════════════════════════════════════════════
function LedgerScreen({ rows, setRows, onNext, onBack, auditLog, setAuditLog, user, tally }) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const [forceImport, setForceImport] = useState({});
  const [bulkLedger, setBulkLedger] = useState("");
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [showAudit, setShowAudit] = useState(false);
  const [showRules, setShowRules] = useState(false);
  const [rules, setRules] = useState(loadRules);
  const [editRule, setEditRule] = useState(null); // rule being edited in modal
  const [newRule, setNewRule] = useState({ keyword:"", matchType:"contains", ledger:"", voucherType:"Payment", enabled:true });
  const [ruleTestNarr, setRuleTestNarr] = useState("");

  // Persist rules whenever they change
  useEffect(() => { saveRules(rules); }, [rules]);

  const reApplyRules = () => {
    setRows(prev => {
      const reset = prev.map(r => ({ ...r, ledger: r.aiLedger || r.ledger, ruleMatched: undefined }));
      return applyRulesToRows(reset, rules);
    });
    toast(`Rules re-applied to ${rows.length} rows`, "success");
  };

  const addRule = () => {
    if (!newRule.keyword || !newRule.ledger) return;
    const r = { ...newRule, id: genRuleId(), priority: rules.length + 1 };
    const updated = [...rules, r];
    setRules(updated);
    setNewRule({ keyword:"", matchType:"contains", ledger:"", voucherType:"Payment", enabled:true });
    toast(`Rule added — click Re-Apply to update rows`, "success");
  };

  const deleteRule = id => setRules(prev => prev.filter(r=>r.id!==id).map((r,i)=>({...r,priority:i+1})));
  const toggleRule = id => setRules(prev => prev.map(r=>r.id===id?{...r,enabled:!r.enabled}:r));
  const moveRule = (id, dir) => setRules(prev => {
    const idx = prev.findIndex(r=>r.id===id);
    if ((dir===-1&&idx===0)||(dir===1&&idx===prev.length-1)) return prev;
    const arr = [...prev];
    [arr[idx],arr[idx+dir]] = [arr[idx+dir],arr[idx]];
    return arr.map((r,i)=>({...r,priority:i+1}));
  });

  const ruleMatchCount = (rule) => rows.filter(r => matchRule(rule, r.narration)).length;

  const dupRows = rows.filter(r=>r.isDuplicate);
  const suspenseRows = rows.filter(r=>r.ledger==="Suspense Account"&&!r.isDuplicate);

  // Merge live Tally ledgers (if available) with built-in fallback
  const activeLedgerGroups = useMemo(() => {
    const allLive = Object.values(tally?.ledgerMap || {}).flat();
    const liveSet = new Set(allLive);
    if (liveSet.size === 0) return TALLY_LEDGERS; // pure fallback
    // Keep existing groups but append live ledgers that aren't already present
    const groups = TALLY_LEDGERS.map(g => ({
      group: g.group,
      items: g.items.filter(i => liveSet.has(i) || liveSet.size === 0)
    })).filter(g => g.items.length > 0);
    // Add "From Tally" group with any live ledgers not in our static list
    const allStatic = new Set(TALLY_LEDGERS.flatMap(g => g.items));
    const extra = allLive.filter(l => !allStatic.has(l));
    if (extra.length) groups.push({ group: "From Tally", items: [...new Set(extra)].sort() });
    return groups.length ? groups : TALLY_LEDGERS;
  }, [tally?.ledgerMap]);

  const filtered = useMemo(() => rows.filter(r => {
    if (filter==="suspense" && r.ledger!=="Suspense Account") return false;
    if (filter==="duplicate" && !r.isDuplicate) return false;
    if (filter==="ready" && (r.ledger==="Suspense Account"||r.isDuplicate)) return false;
    if (search) {
      const s = search.toLowerCase();
      return r.narration?.toLowerCase().includes(s) || String(r.date).includes(s) || r.ledger?.toLowerCase().includes(s);
    }
    return true;
  }), [rows, filter, search]);

  const setLedger = (id, val, reason="manual") => {
    const row = rows.find(r=>r.id===id);
    setRows(rs => rs.map(r => r.id===id ? {...r, ledger:val} : r));
    setAuditLog(l => [...l, { id:genId(), ts:new Date().toISOString(), user:user?.name, rowId:id, from:row?.ledger, to:val, narration:row?.narration?.slice(0,30), reason }]);
  };

  const applyBulk = () => {
    if (!bulkLedger) return;
    const targets = selectedRows.size > 0
      ? [...selectedRows]
      : rows.filter(r=>r.ledger==="Suspense Account").map(r=>r.id);
    targets.forEach(id => setLedger(id, bulkLedger, "bulk"));
    setSelectedRows(new Set());
    setBulkLedger("");
  };

  const toggleSelect = id => setSelectedRows(s => { const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const selectAll = () => setSelectedRows(new Set(filtered.map(r=>r.id)));
  const clearSelect = () => setSelectedRows(new Set());

  const readyCount = rows.filter(r=>!r.isDuplicate&&r.ledger!=="Suspense Account").length;
  const skipCount = rows.filter(r=>r.isDuplicate&&!forceImport[r.id]).length;

  return (
    <div className="fade-in">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:18 }}>
        <div>
          <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:4 }}>Assign Ledgers</h2>
          <p style={{ color:T.textDim, fontSize:13 }}>Review AI suggestions · correct assignments · handle duplicates</p>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <Pill color="green" dot>{readyCount} ready</Pill>
          {suspenseRows.length>0 && <Pill color="amber" dot>{suspenseRows.length} suspense</Pill>}
          {dupRows.length>0 && <Pill color="red" dot>{dupRows.length} duplicates</Pill>}
          <Btn size="sm" variant="ghost" onClick={()=>setShowAudit(true)}>Audit log ({auditLog.length})</Btn>
          <Btn size="sm" variant="secondary" onClick={()=>setShowRules(true)}>Mapping Rules ({rules.filter(r=>r.enabled).length} active)</Btn>
        </div>
      </div>

      {/* Duplicate warning */}
      {dupRows.length > 0 && (
        <div style={{ background:T.redDim, border:`1px solid ${T.red}44`, borderRadius:11, padding:"12px 16px", marginBottom:14 }} className="fade-in">
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <span style={{ color:T.red, fontWeight:600, fontSize:13 }}>⚠ {dupRows.length} potential duplicate transaction{dupRows.length>1?"s":""} auto-detected</span>
            <div style={{ display:"flex", gap:8 }}>
              <Btn size="sm" variant="secondary" onClick={()=>setFilter(filter==="duplicate"?"all":"duplicate")} icon="🔍">
                {filter==="duplicate"?"Show All":"Inspect Duplicates"}
              </Btn>
              <Btn size="sm" variant="danger" onClick={()=>{const fi={};dupRows.forEach(r=>fi[r.id]=true);setForceImport(fi);}} icon="⚡">Force Import All</Btn>
            </div>
          </div>
        </div>
      )}

      {/* Toolbar */}
      <div style={{ display:"flex", gap:8, marginBottom:12, flexWrap:"wrap", alignItems:"center" }}>
        <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search narration, ledger…" prefix="" style={{ width:220 }} />
        <div style={{ display:"flex", gap:4 }}>
          {["all","suspense","duplicate","ready"].map(f => (
            <button key={f} onClick={()=>setFilter(f)}
              style={{ padding:"6px 12px", borderRadius:7, fontSize:12, fontFamily:T.font, cursor:"pointer", transition:"all 0.15s",
                background:filter===f?T.accent:T.surface, color:filter===f?"#fff":T.textMid, border:`1px solid ${filter===f?T.accent:T.border}` }}>
              {f.charAt(0).toUpperCase()+f.slice(1)}
            </button>
          ))}
        </div>
        <div style={{ flex:1 }} />
        <select value={bulkLedger} onChange={e=>setBulkLedger(e.target.value)} style={{ padding:"7px 10px", minWidth:220, borderRadius:8 }}>
          <option value="">Bulk assign {selectedRows.size>0?`${selectedRows.size} selected`:"suspense rows"} to…</option>
          {activeLedgerGroups.map(g => <optgroup key={g.group} label={g.group}>{g.items.map(l=><option key={l}>{l}</option>)}</optgroup>)}
        </select>
        <Btn size="sm" variant="amber" onClick={applyBulk} disabled={!bulkLedger} icon="⚡">Apply</Btn>
        {selectedRows.size===0 ? <Btn size="sm" variant="ghost" onClick={selectAll}>Select all</Btn> : <Btn size="sm" variant="ghost" onClick={clearSelect}>Clear ({selectedRows.size})</Btn>}
      </div>

      {/* Table */}
      <div style={{ overflowX:"auto", borderRadius:12, border:`1px solid ${T.border}`, maxHeight:420, overflowY:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead style={{ position:"sticky", top:0, zIndex:2 }}>
            <tr style={{ background:T.surface }}>
              <th style={{ padding:"10px 8px", width:36, borderBottom:`1px solid ${T.border}` }}>
                <input
                  type="checkbox"
                  title={selectedRows.size===filtered.length&&filtered.length>0?"Deselect all":"Select all visible"}
                  checked={filtered.length>0 && selectedRows.size===filtered.length}
                  ref={el=>{ if(el) el.indeterminate = selectedRows.size>0 && selectedRows.size<filtered.length; }}
                  onChange={()=> selectedRows.size===filtered.length ? clearSelect() : selectAll()}
                  style={{ accentColor:T.accent, cursor:"pointer" }}
                />
              </th>
              {["Date","Narration","Ref","Debit ₹","Credit ₹","Voucher","AI Ledger","Assign Ledger","Status"].map(h=>(
                <th key={h} style={{ padding:"10px 12px", textAlign:"left", fontSize:11, fontWeight:600, color:T.textDim, borderBottom:`1px solid ${T.border}`, whiteSpace:"nowrap", letterSpacing:"0.04em" }}>{h.toUpperCase()}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map(r => {
              const isDup = r.isDuplicate && !forceImport[r.id];
              const isSus = r.ledger === "Suspense Account";
              const vtype = r.voucherType || voucherType(r.debit, r.credit, r.ledger);
              return (
                <tr key={r.id} className="row-hover"
                  style={{ background: isDup ? T.redDim+"66" : isSus ? T.amberDim+"44" : "transparent", borderBottom:`1px solid ${T.border}22`, transition:"background 0.15s" }}>
                  <td style={{ padding:"8px 8px", textAlign:"center" }}>
                    <input type="checkbox" checked={selectedRows.has(r.id)} onChange={()=>toggleSelect(r.id)}
                      style={{ accentColor:T.accent, cursor:"pointer" }} />
                  </td>
                  <td style={{ padding:"8px 12px", color:T.textMid, whiteSpace:"nowrap", fontFamily:T.mono, fontSize:11 }}>{fmtDate(r.date)}</td>
                  <td style={{ padding:"8px 12px", maxWidth:220, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:T.text }} title={r.narration}>{r.narration}</td>
                  <td style={{ padding:"8px 12px", color:T.textDim, fontFamily:T.mono, fontSize:11 }}>{r.ref||"—"}</td>
                  <td style={{ padding:"8px 12px", color:T.red, fontWeight:500, textAlign:"right", fontFamily:T.mono }}>{r.debit ? fmt(r.debit) : ""}</td>
                  <td style={{ padding:"8px 12px", color:T.green, fontWeight:500, textAlign:"right", fontFamily:T.mono }}>{r.credit ? fmt(r.credit) : ""}</td>
                  <td style={{ padding:"8px 12px" }}>
                    <Pill color={vtype==="Receipt"?"green":vtype==="Payment"?"red":vtype==="Contra"?"blue":"purple"} size="xs">{vtype}</Pill>
                  </td>
                  <td style={{ padding:"8px 12px" }}>
                    <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                      <Pill color={r.aiLedger==="Suspense Account"?"amber":"blue"} size="xs">{(r.aiLedger||"").slice(0,18)}</Pill>
                      {r.ruleMatched && (
                        <span title={`Matched rule: "${r.ruleMatched}"`} style={{ fontSize:9, fontWeight:700, background:T.greenDim, color:T.green, border:`1px solid ${T.green}44`, borderRadius:4, padding:"1px 4px", whiteSpace:"nowrap" }}>
                          Rule
                        </span>
                      )}
                    </div>
                  </td>
                  <td style={{ padding:"8px 12px", minWidth:200 }}>
                    {isDup ? (
                      <Btn size="sm" variant="danger" onClick={()=>setForceImport(f=>({...f,[r.id]:true}))} icon="⚡">Force import</Btn>
                    ) : (
                      <select value={r.ledger} onChange={e=>setLedger(r.id,e.target.value)} style={{ padding:"5px 8px", width:"100%", fontSize:11, borderRadius:7 }}>
                        {activeLedgerGroups.map(g=><optgroup key={g.group} label={g.group}>{g.items.map(l=><option key={l}>{l}</option>)}</optgroup>)}
                      </select>
                    )}
                  </td>
                  <td style={{ padding:"8px 12px" }}>
                    {isDup ? <Pill color="red" size="xs">Skipped</Pill>
                      : isSus ? <Pill color="amber" size="xs">Review</Pill>
                      : <Pill color="green" size="xs" dot>Ready</Pill>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <p style={{ fontSize:11, color:T.textDim, marginTop:8 }}>Showing {filtered.length} of {rows.length} rows · {skipCount} duplicates will be skipped</p>

      {/* ── Rules Manager Modal ── */}
      <Modal open={showRules} onClose={()=>setShowRules(false)} title="Keyword Mapping Rules" width={780}>
        <div style={{ fontSize:12, color:T.textSub, marginBottom:14, lineHeight:1.6 }}>
          Rules run in priority order on every narration during ingestion. Higher-priority rules win. Click <strong>Re-Apply All Rules</strong> to update already-imported rows instantly.
        </div>

        {/* Re-apply + Reset buttons */}
        <div style={{ display:"flex", gap:8, marginBottom:16 }}>
          <Btn variant="primary" onClick={()=>{ reApplyRules(); setShowRules(false); }}>Re-Apply All Rules to Rows</Btn>
          <Btn variant="secondary" onClick={()=>{ setRules(DEFAULT_RULES); saveRules(DEFAULT_RULES); toast("Rules reset to defaults","success"); }}>Reset to Defaults</Btn>
        </div>

        {/* Live test input */}
        <div style={{ background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:9, padding:"10px 14px", marginBottom:16, display:"flex", gap:10, alignItems:"center" }}>
          <div style={{ fontSize:12, fontWeight:600, color:T.accent, whiteSpace:"nowrap" }}>Test narration:</div>
          <input value={ruleTestNarr} onChange={e=>setRuleTestNarr(e.target.value)}
            placeholder="e.g. Zomato order 8923…" style={{ flex:1, padding:"6px 10px", borderRadius:7 }} />
          {ruleTestNarr && (() => {
            const hit = [...rules].filter(r=>r.enabled).sort((a,b)=>a.priority-b.priority).find(r=>matchRule(r,ruleTestNarr));
            return hit
              ? <div style={{ fontSize:11, fontWeight:700, color:T.green, whiteSpace:"nowrap" }}>→ {hit.ledger} · {hit.voucherType} <span style={{ fontWeight:400, color:T.textSub }}>(rule: "{hit.keyword}")</span></div>
              : <div style={{ fontSize:11, color:T.amber, fontWeight:600, whiteSpace:"nowrap" }}>No rule matched → Suspense Account</div>;
          })()}
        </div>

        {/* Add new rule form */}
        <Card style={{ marginBottom:14, padding:"12px 14px", border:`1px dashed ${T.border}` }}>
          <div style={{ fontSize:12, fontWeight:700, color:T.text, marginBottom:10 }}>+ Add New Rule</div>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"flex-end" }}>
            <div style={{ flex:2, minWidth:160 }}>
              <div style={{ fontSize:10, color:T.textSub, marginBottom:3, fontWeight:600 }}>KEYWORD / REGEX</div>
              <input value={newRule.keyword} onChange={e=>setNewRule(p=>({...p,keyword:e.target.value}))}
                placeholder="e.g. zomato|swiggy or ATM" style={{ width:"100%", padding:"6px 10px", borderRadius:7 }} />
            </div>
            <div style={{ minWidth:130 }}>
              <div style={{ fontSize:10, color:T.textSub, marginBottom:3, fontWeight:600 }}>MATCH TYPE</div>
              <select value={newRule.matchType} onChange={e=>setNewRule(p=>({...p,matchType:e.target.value}))} style={{ width:"100%", padding:"6px 8px", borderRadius:7 }}>
                <option value="contains">Contains</option>
                <option value="startsWith">Starts with</option>
                <option value="endsWith">Ends with</option>
                <option value="regex">Regex (|, *, ?)</option>
              </select>
            </div>
            <div style={{ flex:2, minWidth:160 }}>
              <div style={{ fontSize:10, color:T.textSub, marginBottom:3, fontWeight:600 }}>MAP TO LEDGER</div>
              <select value={newRule.ledger} onChange={e=>setNewRule(p=>({...p,ledger:e.target.value}))} style={{ width:"100%", padding:"6px 8px", borderRadius:7 }}>
                <option value="">Select ledger…</option>
                {TALLY_LEDGERS.map(g=><optgroup key={g.group} label={g.group}>{g.items.map(l=><option key={l}>{l}</option>)}</optgroup>)}
              </select>
            </div>
            <div style={{ minWidth:120 }}>
              <div style={{ fontSize:10, color:T.textSub, marginBottom:3, fontWeight:600 }}>VOUCHER TYPE</div>
              <select value={newRule.voucherType} onChange={e=>setNewRule(p=>({...p,voucherType:e.target.value}))} style={{ width:"100%", padding:"6px 8px", borderRadius:7 }}>
                {["Payment","Receipt","Contra","Journal",""].map(v=><option key={v} value={v}>{v||"(auto)"}</option>)}
              </select>
            </div>
            <Btn variant="primary" onClick={addRule} disabled={!newRule.keyword||!newRule.ledger}>Add Rule</Btn>
          </div>
        </Card>

        {/* Rules list */}
        <div style={{ maxHeight:320, overflowY:"auto" }}>
          {rules.map((rule, idx) => {
            const hits = ruleMatchCount(rule);
            return (
              <div key={rule.id} style={{ display:"flex", alignItems:"center", gap:8, padding:"9px 12px", borderRadius:9, marginBottom:4,
                background: rule.enabled ? (hits>0?T.greenDim+"66":T.surface) : T.border+"33",
                border:`1px solid ${rule.enabled?(hits>0?T.green+"44":T.border):T.border+"66"}`,
                opacity: rule.enabled?1:0.55, transition:"all 0.15s" }}>

                {/* Priority drag handles */}
                <div style={{ display:"flex", flexDirection:"column", gap:2 }}>
                  <button onClick={()=>moveRule(rule.id,-1)} disabled={idx===0}
                    style={{ border:"none", background:"transparent", cursor:idx===0?"not-allowed":"pointer", color:T.textSub, fontSize:11, padding:"0 2px", lineHeight:1 }}>▲</button>
                  <button onClick={()=>moveRule(rule.id,1)} disabled={idx===rules.length-1}
                    style={{ border:"none", background:"transparent", cursor:idx===rules.length-1?"not-allowed":"pointer", color:T.textSub, fontSize:11, padding:"0 2px", lineHeight:1 }}>▼</button>
                </div>

                {/* Priority number */}
                <div style={{ width:22, height:22, borderRadius:6, background:T.accentDim, color:T.accent, fontSize:10, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{idx+1}</div>

                {/* Keyword */}
                <div style={{ flex:2, minWidth:0 }}>
                  <div style={{ fontSize:12, fontWeight:600, color:T.text, fontFamily:T.mono, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{rule.keyword}</div>
                  <div style={{ fontSize:10, color:T.textSub }}>{rule.matchType}</div>
                </div>

                {/* Arrow */}
                <div style={{ color:T.textDim, fontSize:14, fontWeight:700 }}>→</div>

                {/* Ledger + voucher */}
                <div style={{ flex:2, minWidth:0 }}>
                  <div style={{ fontSize:12, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{rule.ledger}</div>
                  <Pill size="xs" color={rule.voucherType==="Payment"?"red":rule.voucherType==="Receipt"?"green":rule.voucherType==="Contra"?"blue":"purple"}>{rule.voucherType||"auto"}</Pill>
                </div>

                {/* Hit count badge */}
                <div style={{ minWidth:52, textAlign:"right" }}>
                  {hits > 0
                    ? <span style={{ fontSize:11, fontWeight:700, background:T.greenDim, color:T.green, borderRadius:99, padding:"2px 8px" }}>{hits} rows</span>
                    : <span style={{ fontSize:10, color:T.textDim }}>0 rows</span>}
                </div>

                {/* Toggle enable */}
                <button onClick={()=>toggleRule(rule.id)} title={rule.enabled?"Disable rule":"Enable rule"}
                  style={{ padding:"3px 10px", borderRadius:6, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:"none",
                    background: rule.enabled?T.greenDim:T.border+"44", color:rule.enabled?T.green:T.textSub }}>
                  {rule.enabled?"On":"Off"}
                </button>

                {/* Delete */}
                <button onClick={()=>deleteRule(rule.id)} title="Delete rule"
                  style={{ padding:"3px 8px", borderRadius:6, fontSize:11, fontWeight:700, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.red}33`, background:T.redDim, color:T.red }}>
                  ×
                </button>
              </div>
            );
          })}
        </div>
        <div style={{ fontSize:11, color:T.textSub, marginTop:10 }}>
          {rules.filter(r=>r.enabled).length} active rules · {rules.filter(r=>r.enabled).reduce((s,r)=>s+ruleMatchCount(r),0)} total row matches
        </div>
      </Modal>

      {/* Audit Modal */}
      <Modal open={showAudit} onClose={()=>setShowAudit(false)} title="Audit Trail" width={640}>
        {auditLog.length === 0 ? (
          <p style={{ color:T.textDim, fontSize:13, textAlign:"center", padding:20 }}>No changes yet. All ledger changes will be logged here.</p>
        ) : (
          <div style={{ maxHeight:400, overflowY:"auto" }}>
            {[...auditLog].reverse().map(l=>(
              <div key={l.id} style={{ padding:"10px 0", borderBottom:`1px solid ${T.border}`, display:"flex", gap:12, alignItems:"flex-start" }}>
                <div style={{ width:32, height:32, borderRadius:"50%", background:T.accentDim, color:T.accent, display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, flexShrink:0 }}>
                  {l.user?.split(" ").map(w=>w[0]).join("")}
                </div>
                <div style={{ flex:1 }}>
                  <p style={{ fontSize:12, color:T.text, marginBottom:2 }}>
                    <strong>{l.user}</strong> changed <span style={{fontFamily:T.mono,color:T.textMid,fontSize:11}}>&ldquo;{l.narration}…&rdquo;</span>
                  </p>
                  <p style={{ fontSize:11, color:T.textDim }}>
                    <span style={{color:T.red}}>{l.from}</span> → <span style={{color:T.green}}>{l.to}</span>
                    <span style={{marginLeft:10}}>{l.reason}</span>
                  </p>
                  <p style={{ fontSize:10, color:T.textDim, marginTop:2 }}>{new Date(l.ts).toLocaleString("en-IN")}</p>
                </div>
              </div>
            ))}
          </div>
        )}
      </Modal>

      <div style={{ display:"flex", justifyContent:"space-between", marginTop:20 }}>
        <Btn variant="secondary" onClick={onBack} icon="←">Back</Btn>
        <Btn onClick={()=>{
          // Sync forceImport state into row objects before proceeding
          setRows(rs => rs.map(r => ({...r, forceImport: !!forceImport[r.id]})));
          onNext();
        }} icon="→">Preview ({rows.filter(r=>!r.isDuplicate||forceImport[r.id]).length} rows)</Btn>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Preview & Export
// ══════════════════════════════════════════════════════════════════
function PreviewScreen({ rows, setRows, filename, selectedCompanies, onBack, onImport, auditLog, tally }) {
  const [exporting, setExporting] = useState(null);
  const [showXml, setShowXml] = useState(false);
  const [selectedCompanyForXml, setSelectedCompanyForXml] = useState(selectedCompanies[0]);

  const validRows = rows.filter(r => !r.isDuplicate || r.forceImport);
  const totalDebit = validRows.reduce((s,r) => s+(parseFloat(r.debit)||0), 0);
  const totalCredit = validRows.reduce((s,r) => s+(parseFloat(r.credit)||0), 0);
  const net = totalCredit - totalDebit;
  const liveCompanies = (tally?.companies || []).filter(c=>selectedCompanies.includes(c.id));
  const companies = liveCompanies.length ? liveCompanies
    : selectedCompanies.map(id => window.__adHocCompanies?.[id] || { id, name: id });

  const exportExcel = () => {
    setExporting("excel");
    const data = validRows.map(r => ({
      "Date": fmtDate(r.date), "Narration": r.narration, "Reference": r.ref||"",
      "Debit (₹)": r.debit||"", "Credit (₹)": r.credit||"",
      "Voucher Type": r.voucherType||voucherType(r.debit,r.credit,r.ledger),
      "Ledger Account": r.ledger, "AI Suggestion": r.aiLedger,
      "Status": r.isDuplicate?"Duplicate":r.ledger==="Suspense Account"?"Suspense":"Ready"
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    // Column widths
    ws["!cols"] = [12,40,16,14,14,14,24,24,10].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, "Tally Vouchers");
    // Summary sheet
    const summary = [
      ["Bank2Tally Export Summary"],[""],
      ["File", filename],["Generated", new Date().toLocaleString("en-IN")],
      ["Total Rows", validRows.length],["Total Debit", totalDebit],["Total Credit", totalCredit],["Net", net],
      ["Suspense Rows", validRows.filter(r=>r.ledger==="Suspense Account").length],
      ["Companies", companies.map(c=>c.name).join(", ")],[""],
      ["Ledger Breakdown"], ["Ledger","Count","Total Amount"],
    ];
    const lgMap = {};
    validRows.forEach(r=>{ lgMap[r.ledger]=(lgMap[r.ledger]||{count:0,amt:0}); lgMap[r.ledger].count++; lgMap[r.ledger].amt+=parseFloat(r.debit||r.credit||0); });
    Object.entries(lgMap).sort((a,b)=>b[1].count-a[1].count).forEach(([l,v])=>summary.push([l,v.count,v.amt]));
    const ws2 = XLSX.utils.aoa_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, ws2, "Summary");
    XLSX.writeFile(wb, `Bank2Tally_${filename.replace(/\.[^.]+$/,"")}_${Date.now()}.xlsx`);
    setTimeout(()=>setExporting(null),1000);
  };

  const exportCSV = () => {
    setExporting("csv");
    const hdr = ["Date","Narration","Reference","Debit","Credit","Voucher Type","Ledger","AI Suggestion"];
    const csvRows = [hdr.join(","), ...validRows.map(r=>[
      fmtDate(r.date), `"${(r.narration||"").replace(/"/g,'""')}"`, r.ref||"",
      r.debit||"", r.credit||"", r.voucherType||voucherType(r.debit,r.credit,r.ledger),
      `"${r.ledger}"`, `"${r.aiLedger||""}"`
    ].join(","))];
    const blob = new Blob([csvRows.join("\n")], {type:"text/csv"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`Bank2Tally_${Date.now()}.csv`; a.click();
    setTimeout(()=>setExporting(null),1000);
  };

  const exportXML = (company) => {
    setExporting("xml");
    const xml = toTallyXML(validRows, company);
    const blob = new Blob([xml], {type:"text/xml"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`TallyImport_${company.name.replace(/\s+/g,"_")}_${Date.now()}.xml`; a.click();
    setTimeout(()=>setExporting(null),1000);
  };

  const xmlPreview = showXml ? toTallyXML(validRows.slice(0,2), companies.find(c=>c.id===selectedCompanyForXml)||companies[0]) : "";

  // Ledger breakdown
  const lgBreakdown = useMemo(() => {
    const m = {};
    validRows.forEach(r=>{ m[r.ledger]=(m[r.ledger]||0)+1; });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,8);
  }, [validRows]);

  return (
    <div className="fade-in">
      <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:4 }}>Preview & Export</h2>
      <p style={{ color:T.textDim, fontSize:13, marginBottom:20 }}>Final review before pushing to Tally or exporting</p>

      {/* ── Audit Compliance Bento Grid ── */}
      {(() => {
        const score = calcComplianceScore(validRows);
        const mismatches = validRows.filter(r=>r.balanceMismatch).length;
        const badVoucher = validRows.filter(r=>!validateVoucherType(r).valid).length;
        const suspense   = validRows.filter(r=>r.ledger==="Suspense Account").length;
        const dups       = rows.filter(r=>r.isDuplicate).length;
        const scoreColor = score>=90?T.green:score>=70?T.amber:T.red;
        return (
          <div style={{ marginBottom:20 }}>
            {/* Top stat row */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:10, marginBottom:10 }}>
              <StatCard label="Transactions" value={validRows.length} color={T.accent} />
              <StatCard label="Total Debit" value={"₹"+fmt(totalDebit)} color={T.red} />
              <StatCard label="Total Credit" value={"₹"+fmt(totalCredit)} color={T.green} />
              <StatCard label="Net Balance" value={(net>=0?"+ ":"- ")+"₹"+fmt(Math.abs(net))} color={net>=0?T.green:T.red} />
              <StatCard label="Suspense" value={suspense} color={suspense>0?T.amber:T.green} />
            </div>
            {/* Compliance bento */}
            <Card style={{ padding:"16px 20px" }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:14 }}>
                <div>
                  <div style={{ fontSize:13, fontWeight:700, color:T.text }}>Audit Compliance Score</div>
                  <div style={{ fontSize:11, color:T.textSub, marginTop:1 }}>Based on unmapped, duplicate, balance & voucher checks</div>
                </div>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ fontSize:32, fontWeight:900, color:scoreColor, fontFamily:T.mono }}>{score}%</div>
                  <div style={{ width:60, height:60, position:"relative" }}>
                    <svg viewBox="0 0 36 36" style={{ transform:"rotate(-90deg)" }}>
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke={T.border} strokeWidth="3"/>
                      <circle cx="18" cy="18" r="15.9" fill="none" stroke={scoreColor} strokeWidth="3"
                        strokeDasharray={`${score} ${100-score}`} strokeLinecap="round"/>
                    </svg>
                  </div>
                </div>
              </div>
              {/* Score deduction breakdown */}
              <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8 }}>
                {[
                  { label:"Balance Mismatches", val:mismatches, color:mismatches>0?T.red:T.green,
                    action: mismatches>0 ? ()=>{ setRows(prev=>prev.map(r=>r.balanceMismatch?{...r,balance:String(r.expectedBalance||r.balance)}:r)); toast("Balance mismatches auto-repaired","success"); } : null,
                    actionLabel:"Auto-Repair" },
                  { label:"Voucher Type Errors", val:badVoucher, color:badVoucher>0?T.red:T.green,
                    action: badVoucher>0 ? ()=>{ setRows(prev=>prev.map(r=>{ const v=validateVoucherType(r); return !v.valid?{...r,voucherType:v.fix}:r; })); toast("Voucher types auto-corrected","success"); } : null,
                    actionLabel:"Quick-Fix All" },
                  { label:"Duplicates Skipped", val:dups, color:dups>0?T.amber:T.green, action:null },
                  { label:"Suspense Rows", val:suspense, color:suspense>0?T.amber:T.green, action:null },
                ].map(item=>(
                  <div key={item.label} style={{ background:item.val>0?`${item.color}10`:T.greenDim, border:`1px solid ${item.val>0?item.color+"33":T.green+"33"}`, borderRadius:9, padding:"10px 12px" }}>
                    <div style={{ fontSize:20, fontWeight:800, color:item.color, marginBottom:2 }}>{item.val}</div>
                    <div style={{ fontSize:11, color:T.textMid, fontWeight:500, marginBottom:item.action?6:0 }}>{item.label}</div>
                    {item.action && (
                      <button onClick={item.action}
                        style={{ fontSize:10, fontWeight:700, color:item.color, background:`${item.color}15`, border:`1px solid ${item.color}44`, borderRadius:5, padding:"2px 8px", cursor:"pointer", fontFamily:T.font }}>
                        {item.actionLabel}
                      </button>
                    )}
                  </div>
                ))}
              </div>
              {/* Mismatch detail rows */}
              {mismatches > 0 && (
                <div style={{ marginTop:12, background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"8px 12px" }}>
                  <div style={{ fontSize:11, fontWeight:700, color:T.red, marginBottom:6 }}>Balance Mismatch Details</div>
                  {validRows.filter(r=>r.balanceMismatch).slice(0,3).map((r,i)=>(
                    <div key={i} style={{ fontSize:11, color:T.textMid, marginBottom:3 }}>
                      Row {i+1} · {r.date} · {r.narration?.slice(0,30)} · Stated: ₹{r.balance} · Expected: ₹{r.expectedBalance?.toFixed(2)}
                    </div>
                  ))}
                  {mismatches>3 && <div style={{ fontSize:10, color:T.textSub }}>…and {mismatches-3} more. Click Auto-Repair to fix all.</div>}
                </div>
              )}
              {/* Voucher error detail */}
              {badVoucher > 0 && (
                <div style={{ marginTop:8, background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"8px 12px" }}>
                  <div style={{ fontSize:11, fontWeight:700, color:T.red, marginBottom:4 }}>Voucher Type Issues</div>
                  {validRows.filter(r=>!validateVoucherType(r).valid).slice(0,3).map((r,i)=>{
                    const v = validateVoucherType(r);
                    return <div key={i} style={{ fontSize:11, color:T.textMid, marginBottom:2 }}>{r.date} · {r.narration?.slice(0,25)} · {v.issue} → should be "{v.fix}"</div>;
                  })}
                </div>
              )}
            </Card>
          </div>
        );
      })()}

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:20 }}>
        {/* Ledger breakdown */}
        <Card>
          <p style={{ fontWeight:600, fontSize:13, marginBottom:12, color:T.text }}>Top Ledger Accounts</p>
          {lgBreakdown.map(([l,n])=>(
            <div key={l} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
              <span style={{ fontSize:12, color:T.textMid, maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{l}</span>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <div style={{ width:80, height:4, borderRadius:99, background:T.border, overflow:"hidden" }}>
                  <div style={{ width:`${Math.min(100,(n/validRows.length)*100)}%`, height:"100%", background:l==="Suspense Account"?T.amber:T.accent, borderRadius:99 }} />
                </div>
                <span style={{ fontSize:11, color:T.textDim, fontFamily:T.mono, minWidth:24, textAlign:"right" }}>{n}</span>
              </div>
            </div>
          ))}
        </Card>

        {/* Company targets */}
        <Card>
          <p style={{ fontWeight:600, fontSize:13, marginBottom:12, color:T.text }}>Export Targets</p>
          {companies.map(c=>(
            <div key={c.id} style={{ padding:"10px 12px", background:T.surface, borderRadius:9, marginBottom:8, border:`1px solid ${T.border}` }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div>
                  <div style={{ fontWeight:500, fontSize:13, color:T.text }}>{c.name}</div>
                  <div style={{ fontSize:11, color:T.textDim, fontFamily:T.mono }}>{c.gstin} · {c.state} · FY {c.fy}</div>
                </div>
                <Btn size="sm" variant="outline" onClick={()=>exportXML(c)} disabled={exporting==="xml"} icon="⬇">XML</Btn>
              </div>
            </div>
          ))}
          <div style={{ marginTop:8 }}>
            <Btn size="sm" variant="ghost" onClick={()=>setShowXml(!showXml)} icon="👁">
              {showXml ? "Hide" : "Preview"} Tally XML
            </Btn>
          </div>
        </Card>
      </div>

      {/* XML preview */}
      {showXml && (
        <Card style={{ marginBottom:16, padding:14 }} className="fade-in">
          <p style={{ fontSize:11, color:T.textDim, marginBottom:8 }}>Tally XML preview (first 2 vouchers)</p>
          <pre style={{ fontFamily:T.mono, fontSize:10, color:T.textMid, overflow:"auto", maxHeight:200, lineHeight:1.6 }}>{xmlPreview}</pre>
        </Card>
      )}

      {/* Export buttons */}
      <div style={{ display:"flex", gap:10, marginBottom:20, flexWrap:"wrap" }}>
        <Btn onClick={exportExcel} disabled={exporting==="excel"} icon="📊" variant="success">
          {exporting==="excel" ? "Exporting…" : "Export Excel (with summary)"}
        </Btn>
        <Btn onClick={exportCSV} disabled={exporting==="csv"} icon="📄" variant="secondary">
          {exporting==="csv" ? "Exporting…" : "Export CSV"}
        </Btn>
        <Btn onClick={()=>window.print()} icon="🖨" variant="secondary">Print / PDF</Btn>
      </div>

      {/* Preview table */}
      <Card style={{ padding:0, overflow:"hidden" }}>
        <div style={{ padding:"14px 18px", borderBottom:`1px solid ${T.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span style={{ fontWeight:600, fontSize:13 }}>Transaction Preview</span>
          <span style={{ fontSize:11, color:T.textDim }}>{validRows.length} rows</span>
        </div>
        <div style={{ overflowX:"auto", maxHeight:320, overflowY:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead style={{ position:"sticky", top:0, background:T.surface, zIndex:1 }}>
              <tr>
                {["Date","Narration","Debit ₹","Credit ₹","Type","Ledger Account"].map(h=>(
                  <th key={h} style={{ padding:"10px 14px", textAlign:"left", fontSize:11, fontWeight:600, color:T.textDim, borderBottom:`1px solid ${T.border}`, whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {validRows.map(r=>(
                <tr key={r.id} className="row-hover" style={{ borderBottom:`1px solid ${T.border}22` }}>
                  <td style={{ padding:"8px 14px", color:T.textMid, fontFamily:T.mono, fontSize:11, whiteSpace:"nowrap" }}>{fmtDate(r.date)}</td>
                  <td style={{ padding:"8px 14px", color:T.text, maxWidth:260, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{r.narration}</td>
                  <td style={{ padding:"8px 14px", textAlign:"right", color:T.red, fontFamily:T.mono }}>{r.debit?fmt(r.debit):""}</td>
                  <td style={{ padding:"8px 14px", textAlign:"right", color:T.green, fontFamily:T.mono }}>{r.credit?fmt(r.credit):""}</td>
                  <td style={{ padding:"8px 14px" }}>
                    {(() => {
                      const vt = r.voucherType || voucherType(r.debit, r.credit, r.ledger);
                      const validation = validateVoucherType({...r, voucherType: vt});
                      return (
                        <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                          <Pill size="xs" color={!validation.valid?"red":vt==="Receipt"?"green":vt==="Payment"?"red":"blue"}>{vt}</Pill>
                          {!validation.valid && (
                            <button title={validation.issue} onClick={()=>setRows(prev=>prev.map(x=>x.id===r.id?{...x,voucherType:validation.fix}:x))}
                              style={{ fontSize:9, fontWeight:700, background:T.greenDim, color:T.green, border:`1px solid ${T.green}44`, borderRadius:4, padding:"1px 5px", cursor:"pointer", fontFamily:T.font }}>
                              Fix→{validation.fix}
                            </button>
                          )}
                        </div>
                      );
                    })()}
                  </td>
                  <td style={{ padding:"8px 14px" }}>
                    <Pill size="xs" color={r.ledger==="Suspense Account"?"amber":"blue"}>{r.ledger}</Pill>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <div style={{ display:"flex", justifyContent:"space-between", marginTop:20, alignItems:"center", gap:12, flexWrap:"wrap" }}>
        <Btn variant="secondary" onClick={onBack} icon="←">Back</Btn>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
          {/* Download XML — Manual import in Tally */}
          <div>
            {companies.length > 1 ? (
              <div style={{ display:"flex", gap:6 }}>
                {companies.map(c=>(
                  <Btn key={c.id} onClick={()=>exportXML(c)} disabled={exporting==="xml"} icon="📥" variant="outline">
                    {exporting==="xml"?"Exporting…":`Download XML – ${c.name.split(" ")[0]}`}
                  </Btn>
                ))}
              </div>
            ) : (
              <Btn onClick={()=>exportXML(companies[0]||{id:"default",name:"Company"})} disabled={exporting==="xml"} icon="📥" variant="outline">
                {exporting==="xml"?"Generating…":"Download XML (Manual Import)"}
              </Btn>
            )}
          </div>
          {/* Push to Tally — Automatic */}
          <Btn onClick={onImport} icon="🚀" size="lg" style={{ background:`linear-gradient(135deg,${T.accent},${T.purple})`, color:"#fff", border:"none", boxShadow:`0 4px 16px ${T.accentGlow}`, fontWeight:700 }}>
            Push to Tally ({companies.length} {companies.length===1?"company":"companies"})
          </Btn>
        </div>
      </div>

      {/* Export method explainer */}
      <div style={{ marginTop:14, display:"flex", gap:10, flexWrap:"wrap" }}>
        <div style={{ flex:1, minWidth:240, background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:10, padding:"10px 14px", display:"flex", gap:10, alignItems:"flex-start" }}>
          <span style={{ fontSize:18 }}>📥</span>
          <div>
            <div style={{ fontSize:12, fontWeight:700, color:T.accent, marginBottom:2 }}>Download XML — Manual Import</div>
            <div style={{ fontSize:11, color:T.textMid, lineHeight:1.6 }}>Downloads a Tally-compatible XML file. Open Tally → Gateway of Tally → Import Data → Vouchers → select this file.</div>
          </div>
        </div>
        <div style={{ flex:1, minWidth:240, background:T.greenDim, border:`1px solid ${T.green}44`, borderRadius:10, padding:"10px 14px", display:"flex", gap:10, alignItems:"flex-start" }}>
          <span style={{ fontSize:18 }}>🚀</span>
          <div>
            <div style={{ fontSize:12, fontWeight:700, color:T.green, marginBottom:2 }}>Push to Tally — Automatic</div>
            <div style={{ fontSize:11, color:T.textMid, lineHeight:1.6 }}>Sends vouchers directly to Tally via the Bank2Tally Connector extension. Requires Tally running with Gateway enabled on port {tally?.port||9000}.</div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: History
// ══════════════════════════════════════════════════════════════════
function HistoryScreen({ history, onReimport, onDeleteEntry, onClearAll, onBack }) {
  const [search,      setSearch]      = useState("");
  const [expandedId,  setExpandedId]  = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);

  const filtered = history.filter(h =>
    !search || h.filename?.toLowerCase().includes(search.toLowerCase()) || h.company?.toLowerCase().includes(search.toLowerCase())
  );

  const exportFromHistory = (h) => {
    if (!h.rows_data?.length) return;
    const data = h.rows_data.map(r=>({
      Date:fmtDate(r.date), Narration:r.narration, Debit:r.debit||"", Credit:r.credit||"", Ledger:r.ledger, Reference:r.ref||""
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Export");
    XLSX.writeFile(wb, `ReExport_${h.filename.replace(/\.[^.]+$/,"")}_${Date.now()}.xlsx`);
  };

  // Days remaining helper
  const daysLeft = (h) => {
    if (!h.savedAt) return null;
    const saved = new Date(h.savedAt).getTime();
    const expiry = saved + 7 * 24 * 60 * 60 * 1000;
    const left = Math.ceil((expiry - Date.now()) / (24 * 60 * 60 * 1000));
    return Math.max(0, left);
  };

  return (
    <div className="fade-in">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div>
          <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:4 }}>Import History</h2>
          <p style={{ color:T.textDim, fontSize:13 }}>{history.length} import{history.length!==1?"s":""} · auto-deleted after 7 days</p>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          {history.length > 0 && (
            confirmClear ? (
              <div style={{ display:"flex", gap:6, alignItems:"center" }}>
                <span style={{ fontSize:12, color:T.textDim }}>Sure?</span>
                <Btn size="sm" variant="danger" onClick={()=>{ onClearAll(); setConfirmClear(false); }} icon="🗑">Yes, clear all</Btn>
                <Btn size="sm" variant="secondary" onClick={()=>setConfirmClear(false)}>Cancel</Btn>
              </div>
            ) : (
              <Btn size="sm" variant="secondary" onClick={()=>setConfirmClear(true)} icon="🗑">Clear All</Btn>
            )
          )}
          <Btn variant="secondary" onClick={onBack} icon="←">Back</Btn>
        </div>
      </div>

      {/* 7-day info banner */}
      <div style={{ background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:8, padding:"9px 14px", marginBottom:14, display:"flex", alignItems:"center", gap:8 }}>
        <span style={{ fontSize:14 }}>ℹ️</span>
        <p style={{ fontSize:11, color:T.textMid }}>History is stored on this device only and <strong style={{color:T.text}}>auto-deleted after 7 days</strong> to save space. Download exports before they expire.</p>
      </div>

      <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by file name or company…" prefix="" style={{ marginBottom:16, width:"100%" }} />

      {filtered.length === 0 ? (
        <div style={{ textAlign:"center", padding:"60px 0", color:T.textDim }}>
          <div style={{ fontSize:40, marginBottom:12 }}>📭</div>
          <p style={{ fontSize:14 }}>{history.length===0 ? "No imports yet" : "No results matching your search"}</p>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {filtered.map(h => (
            <Card key={h.id} className="card-hover" style={{ padding:0, cursor:"pointer", transition:"border-color 0.2s" }}>
              <div style={{ padding:"16px 20px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div style={{ display:"flex", alignItems:"center", gap:14, flex:1, cursor:"pointer" }}
                  onClick={()=>setExpandedId(expandedId===h.id?null:h.id)}>
                  <div style={{ width:42, height:42, borderRadius:11, background:T.accentDim, display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, flexShrink:0 }}>
                    {h.filename?.endsWith(".pdf")?"📑":h.filename?.endsWith(".csv")?"📄":"📊"}
                  </div>
                  <div>
                    <div style={{ fontWeight:600, fontSize:14, color:T.text, marginBottom:3 }}>{h.filename}</div>
                    <div style={{ fontSize:11, color:T.textDim }}>{h.date} · {h.company}</div>
                    <div style={{ display:"flex", gap:6, marginTop:6, flexWrap:"wrap" }}>
                      <Pill color="blue" size="xs">{h.rows} rows</Pill>
                      {(() => { const d = daysLeft(h); return d !== null ? (
                        <Pill color={d<=1?"red":d<=3?"amber":"green"} size="xs">
                          {d===0?"expires today":`${d}d left`}
                        </Pill>
                      ) : null; })()}
                      {h.suspense>0 && <Pill color="amber" size="xs">{h.suspense} suspense</Pill>}
                      {h.duplicates>0 && <Pill color="red" size="xs">{h.duplicates} duplicates</Pill>}
                      <Pill color="green" size="xs" dot>Imported</Pill>
                    </div>
                  </div>
                </div>
                <div style={{ display:"flex", gap:8, flexShrink:0 }}>
                  <Btn size="sm" variant="secondary" onClick={e=>{e.stopPropagation();onReimport(h);}} icon="🔄">Re-import</Btn>
                  <Btn size="sm" variant="secondary" onClick={e=>{e.stopPropagation();exportFromHistory(h);}} icon="📊">Export</Btn>
                  <Btn size="sm" variant="danger" onClick={e=>{e.stopPropagation();onDeleteEntry(h.id);}} icon="🗑">Delete</Btn>
                  <span style={{ color:T.textDim, fontSize:18, padding:"4px 6px", cursor:"pointer" }}>{expandedId===h.id?"▲":"▼"}</span>
                </div>
              </div>
              {expandedId===h.id && h.rows_data?.length>0 && (
                <div className="fade-in" style={{ borderTop:`1px solid ${T.border}`, padding:"14px 20px", overflowX:"auto" }}>
                  <p style={{ fontSize:11, color:T.textDim, marginBottom:10 }}>First 5 rows preview</p>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                    <thead>
                      <tr>{["Date","Narration","Debit","Credit","Ledger"].map(c=>(
                        <th key={c} style={{ padding:"6px 10px", textAlign:"left", color:T.textDim, fontWeight:600, borderBottom:`1px solid ${T.border}` }}>{c}</th>
                      ))}</tr>
                    </thead>
                    <tbody>
                      {h.rows_data.slice(0,5).map(r=>(
                        <tr key={r.id} style={{ borderBottom:`1px solid ${T.border}22` }}>
                          <td style={{ padding:"6px 10px", color:T.textMid, fontFamily:T.mono }}>{fmtDate(r.date)}</td>
                          <td style={{ padding:"6px 10px", color:T.text, maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{r.narration}</td>
                          <td style={{ padding:"6px 10px", color:T.red, fontFamily:T.mono, textAlign:"right" }}>{r.debit?fmt(r.debit):""}</td>
                          <td style={{ padding:"6px 10px", color:T.green, fontFamily:T.mono, textAlign:"right" }}>{r.credit?fmt(r.credit):""}</td>
                          <td style={{ padding:"6px 10px" }}><Pill size="xs" color={r.ledger==="Suspense Account"?"amber":"blue"}>{r.ledger}</Pill></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SCREEN: Settings
// ══════════════════════════════════════════════════════════════════
function SettingsScreen({ user, onLogout, onUserUpdate, tally, tallyHost, setTallyHost, tallyPort, setTallyPort, defaultLedger, setDefaultLedger, autoDetectLedger, setAutoDetectLedger }) {
  const [saved, setSaved] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState(null);
  const [testMsg, setTestMsg] = useState("");

  // ── Profile edit state ────────────────────────────────────────
  const [profForm, setProfForm]     = useState({ name: user?.name||"", company: user?.company||"", mobile: user?.mobile||"" });
  const [profSaving, setProfSaving] = useState(false);
  const [profErr, setProfErr]       = useState("");
  const [profOk, setProfOk]         = useState("");

  // ── Mobile OTP state ─────────────────────────────────────────
  const [otpSent, setOtpSent]         = useState(false);
  const [otpValue, setOtpValue]       = useState("");
  const [otpSending, setOtpSending]   = useState(false);
  const [otpVerifying, setOtpVerifying] = useState(false);
  const [otpErr, setOtpErr]           = useState("");
  const [mobileVerified, setMobileVerified] = useState(user?.mobile_verified||false);

  // Sync form when user changes
  React.useEffect(() => {
    setProfForm({ name: user?.name||"", company: user?.company||"", mobile: user?.mobile||"" });
    setMobileVerified(user?.mobile_verified||false);
  }, [user?.id]);

  const save = () => { setSaved(true); setTimeout(()=>setSaved(false),2000); };

  // ── Save profile ──────────────────────────────────────────────
  const saveProfile = async () => {
    if (!profForm.name.trim()) return setProfErr("Name is required.");
    setProfSaving(true); setProfErr(""); setProfOk("");
    try {
      const payload = {
        name:    profForm.name.trim(),
        company: profForm.company.trim(),
        mobile:  profForm.mobile.trim(),
        avatar:  profForm.name.trim().slice(0,2).toUpperCase(),
        updated_at: new Date().toISOString(),
      };
      await sb.update("profiles", { id: user.id }, payload);
      if (onUserUpdate) onUserUpdate({ ...user, ...payload });
      setProfOk("Profile saved successfully!");
      setTimeout(()=>setProfOk(""),3000);
    } catch(e) { setProfErr("Save failed: "+e.message); }
    setProfSaving(false);
  };

  // ── Send OTP via Fast2SMS Supabase Edge Function ──────────────
  const sendOtp = async () => {
    const mob = profForm.mobile.trim();
    if (!/^[6-9]\d{9}$/.test(mob)) return setOtpErr("Enter a valid 10-digit Indian mobile number.");
    setOtpSending(true); setOtpErr(""); setOtpValue("");
    try {
      // Generate 6-digit OTP, store hash + expiry in DB (never expose in frontend)
      const code = String(Math.floor(100000 + Math.random() * 900000));
      const expiry = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 min
      // Save to DB first (so verifyOtp can check server-side)
      await sb.update("profiles", { id: user.id }, {
        mobile: mob,
        otp_code: code,
        otp_expiry: expiry,
        mobile_verified: false,
      });
      // Call Supabase Edge Function which calls Fast2SMS (API key stays on server)
      const fnRes = await fetch(`${SUPABASE_URL}/functions/v1/send-otp`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${sb._token || SUPABASE_ANON}`,
        },
        body: JSON.stringify({ mobile: mob, otp: code }),
      });
      const fnData = await fnRes.json();
      if (!fnRes.ok || fnData.error) {
        throw new Error(fnData.error || "SMS gateway error. Please try again.");
      }
      setOtpSent(true);
      setOtpErr("");
    } catch(e) { setOtpErr("Failed to send OTP: " + e.message); }
    setOtpSending(false);
  };

  // ── Verify OTP (server-side check against DB) ─────────────────
  const verifyOtp = async () => {
    if (otpValue.length !== 6) return setOtpErr("Enter the 6-digit OTP sent to your mobile.");
    setOtpVerifying(true); setOtpErr("");
    try {
      const rows = await sb.from("profiles", `id=eq.${user.id}&select=otp_code,otp_expiry`);
      const prof = rows?.[0];
      if (!prof) throw new Error("Profile not found.");
      if (!prof.otp_code) throw new Error("No OTP requested. Please send again.");
      if (new Date(prof.otp_expiry) < new Date()) throw new Error("OTP expired. Please resend.");
      if (prof.otp_code !== otpValue) throw new Error("Incorrect OTP. Please try again.");
      // OTP correct — mark verified, clear from DB
      await sb.update("profiles", { id: user.id }, {
        mobile: profForm.mobile.trim(),
        mobile_verified: true,
        otp_code: null,
        otp_expiry: null,
      });
      setMobileVerified(true);
      setOtpSent(false); setOtpValue("");
      if (onUserUpdate) onUserUpdate({ ...user, mobile: profForm.mobile.trim(), mobile_verified: true });
      setProfOk("Mobile number verified successfully!");
      setTimeout(()=>setProfOk(""), 3000);
    } catch(e) { setOtpErr(e.message); }
    setOtpVerifying(false);
  };

  const runTest = async () => {
    setTesting(true); setTestResult(null); setTestMsg("");
    try {
      await testTallyConnection(tallyHost, tallyPort);
      // Mark extension/connection as ready globally so all other checks pass
      _markExtensionReady();
      setTestResult("ok"); setTestMsg("Connected! Fetching companies from Tally…");
      // Small delay to ensure _extensionReady propagates before gateway fetch
      setTimeout(() => {
        tally.refetch(tallyHost, tallyPort);
      }, 200);
      // Watch tally status reactively — update message when companies load
      let attempts = 0;
      const watchInterval = setInterval(() => {
        attempts++;
        if (tally.status === "ok") {
          setTestMsg(`Connected · ${tally.companies.length} ${tally.companies.length === 1 ? "company" : "companies"} loaded`);
          clearInterval(watchInterval);
        } else if (tally.status === "error" || attempts > 10) {
          setTestMsg(tally.status === "error" ? `Error: ${tally.error}` : "Connected · 0 companies (open a company in Tally)");
          clearInterval(watchInterval);
        }
      }, 500);
    } catch (e) {
      setTestResult("error"); setTestMsg(e.message);
    } finally { setTesting(false); }
  };

  return (
    <div className="fade-in">
      <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:20 }}>Settings</h2>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
        {/* Profile Card */}
        <Card style={{ gridColumn:"1 / -1" }}>
          {/* Profile header */}
          <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:22, padding:"16px 18px", borderRadius:12, background:`linear-gradient(135deg, ${T.accentDim}88, ${T.surface})`, border:`1px solid ${T.accent}22` }}>
            <div style={{ width:64, height:64, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22, fontWeight:800, color:"#fff", flexShrink:0, boxShadow:`0 4px 20px ${T.accent}55` }}>
              {(profForm.name||user?.name||"?").slice(0,2).toUpperCase()}
            </div>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontWeight:700, fontSize:17, color:T.text, marginBottom:3 }}>{user?.name || "—"}</div>
              {/* Email — high contrast so it's always readable */}
              <div style={{ fontSize:13, color:"#94a3b8", fontFamily:T.mono, marginBottom:5, display:"flex", alignItems:"center", gap:6 }}>
                <span style={{ fontSize:11, opacity:0.7 }}>✉</span>
                <span style={{ color: user?.email ? "#cbd5e1" : "#64748b", fontStyle: user?.email ? "normal" : "italic" }}>
                  {user?.email || "No email on file"}
                </span>
              </div>
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                <Pill color="blue" size="xs">{fromDbRole(user?.role)}</Pill>
                <Pill color={user?.status==="approved"?"green":"amber"} size="xs">{user?.status||"pending"}</Pill>
                {mobileVerified && <Pill color="green" size="xs">📱 Verified</Pill>}
                {user?.company && <Pill color="gray" size="xs">🏢 {user.company}</Pill>}
              </div>
            </div>
            <Btn variant="danger" onClick={onLogout} icon="🚪" size="sm">Sign Out</Btn>
          </div>

          {profOk && <div style={{ background:T.greenDim, border:`1px solid ${T.green}44`, borderRadius:8, padding:"9px 14px", fontSize:12, color:T.green, marginBottom:12 }}>✓ {profOk}</div>}
          {profErr && <div style={{ background:T.redDim, border:`1px solid ${T.red}44`, borderRadius:8, padding:"9px 14px", fontSize:12, color:T.red, marginBottom:12 }}>✕ {profErr}</div>}

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
            <div>
              <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Full Name *</label>
              <Input value={profForm.name} onChange={e=>setProfForm(f=>({...f,name:e.target.value}))} placeholder="Your full name" prefix="👤" />
            </div>
            <div>
              <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Company</label>
              <Input value={profForm.company} onChange={e=>setProfForm(f=>({...f,company:e.target.value}))} placeholder="Your company name" prefix="🏢" />
            </div>
          </div>

          {/* Email — read-only, from auth */}
          <div style={{ marginBottom:14 }}>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>
              Email Address
              <span style={{ marginLeft:6, color:"#64748b", fontSize:11 }}>(managed by Supabase Auth · contact admin to change)</span>
            </label>
            <div style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 14px", borderRadius:8, border:`1px solid ${T.accent}33`, background:`${T.accentDim}55`, fontSize:13 }}>
              <span style={{ fontSize:15 }}>✉</span>
              <span style={{ flex:1, color: user?.email ? "#e2e8f0" : "#64748b", fontStyle: user?.email ? "normal" : "italic", fontFamily:T.mono, fontSize:12 }}>
                {user?.email || "No email on file — ask admin to update your profile"}
              </span>
              {user?.email && <Pill color="green" size="xs">✓ Verified</Pill>}
            </div>
          </div>

          {/* Mobile with OTP verification */}
          <div style={{ marginBottom:16 }}>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>
              Mobile Number
              {mobileVerified && <span style={{ marginLeft:8, color:T.green, fontSize:11 }}>✓ Verified</span>}
              {!mobileVerified && profForm.mobile && <span style={{ marginLeft:8, color:T.amber||"#f59e0b", fontSize:11 }}>⚠ Not verified</span>}
            </label>
            <div style={{ display:"flex", gap:8 }}>
              <div style={{ flex:1 }}>
                <Input
                  value={profForm.mobile}
                  onChange={e=>{ setProfForm(f=>({...f,mobile:e.target.value.replace(/\D/g,"")})); setMobileVerified(false); setOtpSent(false); setOtpErr(""); }}
                  placeholder="10-digit mobile number"
                  prefix="📱"
                  maxLength={10}
                />
              </div>
              <Btn
                size="sm"
                variant={mobileVerified?"success":"secondary"}
                onClick={mobileVerified ? undefined : sendOtp}
                disabled={otpSending || mobileVerified || !profForm.mobile}
                icon={otpSending?"⏳":mobileVerified?"✓":"📤"}
              >
                {otpSending ? "Sending…" : mobileVerified ? "Verified" : otpSent ? "Resend OTP" : "Send OTP"}
              </Btn>
            </div>

            {/* OTP input */}
            {otpSent && !mobileVerified && (
              <div style={{ marginTop:10, padding:"14px 16px", background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:10 }}>
                <p style={{ fontSize:12, color:T.textMid, marginBottom:10 }}>Enter the 6-digit OTP sent to +91 {profForm.mobile}</p>
                {otpErr && <p style={{ fontSize:11, color:T.red, marginBottom:8 }}>✕ {otpErr}</p>}
                <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap" }}>
                  {/* 6 individual OTP boxes */}
                  {[0,1,2,3,4,5].map(i => (
                    <input
                      key={i}
                      id={`otp-${i}`}
                      maxLength={1}
                      value={otpValue[i]||""}
                      onChange={e => {
                        const v = e.target.value.replace(/\D/,"");
                        const arr = otpValue.split("");
                        arr[i] = v;
                        const next = arr.join("").slice(0,6);
                        setOtpValue(next);
                        setOtpErr("");
                        if (v && i < 5) document.getElementById(`otp-${i+1}`)?.focus();
                      }}
                      onKeyDown={e => {
                        if (e.key==="Backspace" && !otpValue[i] && i>0) document.getElementById(`otp-${i-1}`)?.focus();
                      }}
                      style={{ width:38, height:44, textAlign:"center", fontSize:18, fontWeight:700, borderRadius:9, border:`2px solid ${otpValue[i]?T.accent:T.border}`, background:T.surface, color:T.text, outline:"none", fontFamily:T.font }}
                    />
                  ))}
                  <Btn
                    variant="primary"
                    onClick={verifyOtp}
                    disabled={otpVerifying || otpValue.length < 6}
                    icon={otpVerifying?"⏳":"✓"}
                  >
                    {otpVerifying ? "Verifying…" : "Verify"}
                  </Btn>
                </div>
                <p style={{ fontSize:11, color:T.textDim, marginTop:8 }}>OTP valid for 10 minutes · Sent via SMS to your mobile</p>
              </div>
            )}
          </div>

          <Btn variant="primary" onClick={saveProfile} disabled={profSaving} icon={profSaving?"⏳":"💾"}>
            {profSaving ? "Saving…" : "Save Profile"}
          </Btn>
        </Card>

        {/* Tally connection */}
        <Card>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
            <p style={{ fontWeight:600, fontSize:14, color:T.text }}>Tally Gateway</p>
            {tally.status === "ok" && <Pill color="green" dot>Live</Pill>}
            {tally.status === "connecting" && <Pill color="amber" dot>Connecting…</Pill>}
            {tally.status === "error" && <Pill color="red" dot>Offline</Pill>}
          </div>
          <div style={{ marginBottom:12 }}>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Host</label>
            <Input value={tallyHost} onChange={e=>setTallyHost(e.target.value)} prefix="🖥" />
          </div>
          <div style={{ marginBottom:16 }}>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Port</label>
            <Input value={tallyPort} onChange={e=>setTallyPort(e.target.value)} prefix="🔌" />
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <Btn size="sm" variant={testResult==="ok"?"success":testResult==="error"?"danger":"secondary"}
              icon={testing?"⏳":testResult==="ok"?"✓":testResult==="error"?"✕":"🔗"}
              fullWidth onClick={runTest} disabled={testing}>
              {testing ? "Testing…" : testResult==="ok" ? "Connected!" : testResult==="error" ? "Failed" : "Test Connection"}
            </Btn>
          </div>
          {testMsg && (
            <div style={{ marginTop:10, borderRadius:10, padding:"12px 14px",
              background: testResult==="ok" ? T.greenDim : T.redDim,
              border: `1px solid ${testResult==="ok" ? T.green+"44" : T.red+"44"}` }}>
              {testResult === "ok" ? (
                <p style={{ fontSize:12, color:T.green, fontWeight:600 }}>✓ {testMsg}</p>
              ) : (
                <>
                  <p style={{ fontSize:12, color:T.red, fontWeight:700, marginBottom:8 }}>Connection Failed</p>
                  {/* Detect the TDL Form:Company error specifically */}
                  {(testMsg.includes("Form:Company") || testMsg.includes("No 'PARTS'") || testMsg.includes("TDL")) ? (
                    <div>
                      <p style={{ fontSize:11, color:T.text, fontWeight:600, marginBottom:6 }}>
                        Root cause: The Chrome extension is sending a broken TDL request to Tally that triggers "Form:Company — No PARTS" error.
                      </p>
                      <p style={{ fontSize:11, color:T.textMid, marginBottom:8, lineHeight:1.6 }}>
                        This happens when the extension intercepts the connection test and wraps it in its own TDL form definition that Tally cannot parse.
                      </p>
                      <div style={{ fontSize:11, color:T.textMid, lineHeight:1.8 }}>
                        <strong style={{color:T.text}}>Fix options (try in order):</strong><br/>
                        <strong style={{color:T.accent}}>1.</strong> Update the Bank2Tally Connector extension to the latest version<br/>
                        <strong style={{color:T.accent}}>2.</strong> Disable the extension → use <strong>Download XML</strong> instead of Push to Tally<br/>
                        <strong style={{color:T.accent}}>3.</strong> In Tally: press <strong>Escape</strong> to close any open dialog, then retry<br/>
                        <strong style={{color:T.accent}}>4.</strong> In Tally Prime: F12 → Advanced Config → restart Gateway Server<br/>
                        <strong style={{color:T.accent}}>5.</strong> Use <strong>Download XML</strong> and import manually: Tally → Gateway of Tally → Import Data → Vouchers
                      </div>
                    </div>
                  ) : testMsg.includes("mixed-content") || testMsg.includes("Failed to fetch") ? (
                    <div>
                      <p style={{ fontSize:11, color:T.text, fontWeight:600, marginBottom:6 }}>
                        Mixed-content block: Browser prevents HTTPS → HTTP (localhost) connections.
                      </p>
                      <div style={{ fontSize:11, color:T.textMid, lineHeight:1.8 }}>
                        <strong style={{color:T.accent}}>Fix:</strong> Install the Bank2Tally Connector Chrome extension below — it acts as a local bridge to bypass this restriction.
                      </div>
                    </div>
                  ) : (
                    <div>
                      <p style={{ fontSize:11, color:T.textMid, lineHeight:1.6, marginBottom:8 }}>{testMsg}</p>
                      <div style={{ fontSize:11, color:T.textMid, lineHeight:1.8 }}>
                        <strong style={{color:T.text}}>Checklist:</strong><br/>
                        <strong style={{color:T.accent}}>1.</strong> Tally Prime is open and a company is loaded<br/>
                        <strong style={{color:T.accent}}>2.</strong> F12 → Advanced Config → <strong>Enable Tally Gateway Server</strong> is ON<br/>
                        <strong style={{color:T.accent}}>3.</strong> Port is set to <strong>{tallyPort}</strong> (same in Tally and above)<br/>
                        <strong style={{color:T.accent}}>4.</strong> No firewall blocking localhost:{tallyPort}<br/>
                        <strong style={{color:T.accent}}>5.</strong> Try <strong>Download XML</strong> as a workaround — import manually in Tally
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          )}
          {tally.status === "ok" && tally.companies.length > 0 && (
            <div style={{ marginTop:12, borderTop:`1px solid ${T.border}`, paddingTop:12 }}>
              <p style={{ fontSize:11, color:T.textDim, marginBottom:6 }}>Loaded companies ({tally.companies.length}):</p>
              <div style={{ display:"flex", flexDirection:"column", gap:4, maxHeight:140, overflowY:"auto" }}>
                {tally.companies.map(c => (
                  <div key={c.id} style={{ fontSize:11, color:T.text, display:"flex", justifyContent:"space-between" }}>
                    <span>{c.name}</span>
                    <span style={{ color:T.textDim }}>{c.state || ""} {c.fy ? `· FY ${c.fy}` : ""}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          <p style={{ fontSize:11, color:T.textDim, marginTop:10 }}>
            Tally Prime → F12 &gt; Advanced Config → Enable Tally Gateway on port {tallyPort}
          </p>
          {tally.status === "ok" && (
            <p style={{ fontSize:10, marginTop:4, color:T.textSub }}>
              Connection method: <strong style={{color:T.green}}>{_usedDirectFetch ? "Direct HTTP (fastest)" : "Chrome Extension bridge"}</strong>
            </p>
          )}
          {/* Extension status */}
          <ExtensionStatus />
        </Card>

        {/* Defaults */}
        <Card>
          <p style={{ fontWeight:600, fontSize:14, marginBottom:16, color:T.text }}>Import Defaults</p>
          <div style={{ marginBottom:14 }}>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>Default Ledger (unrecognised)</label>
            <select value={defaultLedger} onChange={e=>setDefaultLedger(e.target.value)} style={{ width:"100%", padding:"8px 10px" }}>
              {ALL_LEDGERS.map(l=><option key={l}>{l}</option>)}
            </select>
          </div>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"10px 0" }}>
            <span style={{ fontSize:13, color:T.text }}>AI auto-detect ledgers</span>
            <div onClick={()=>setAutoDetectLedger(!autoDetectLedger)}
              style={{ width:44, height:24, borderRadius:99, background:autoDetectLedger?T.accent:T.border, cursor:"pointer", position:"relative", transition:"background 0.2s" }}>
              <div style={{ position:"absolute", top:3, left:autoDetectLedger?22:3, width:18, height:18, borderRadius:"50%", background:"#fff", transition:"left 0.2s" }} />
            </div>
          </div>
        </Card>

        {/* About */}
        <Card>
          <p style={{ fontWeight:600, fontSize:14, marginBottom:16, color:T.text }}>About</p>
          <div style={{ display:"flex", flexDirection:"column", gap:10, fontSize:12, color:T.textDim }}>
            {[["Version","2.0.0 (Commercial)"],["License","Professional — Unlimited companies"],["Support","support@bank2tally.in"],["Tally Compat.","Tally Prime 3.x, ERP 9 (6.6+)"],["GST","Compliant with CGST/SGST/IGST"],["Produced by","Verma Consultancy Services"]].map(([k,v])=>(
              <div key={k} style={{ display:"flex", justifyContent:"space-between" }}>
                <span>{k}</span>
                <span style={{ color:k==="Produced by"?T.gold:T.text, fontWeight:k==="Produced by"?700:500 }}>{v}</span>
              </div>
            ))}
            <div style={{ marginTop:12, paddingTop:12, borderTop:`1px solid ${T.border}` }}>
              <p style={{ color:T.textDim, marginBottom:8, fontSize:11 }}>For purchase &amp; support:</p>
              <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
                <a href="tel:+918707401846" style={{ color:T.accent, textDecoration:"none", fontWeight:600, fontSize:12 }}>📞 8707401846</a>
                <a href="mailto:svtiger543939@gmail.com" style={{ color:T.accent, textDecoration:"none", fontWeight:600, fontSize:12 }}>✉ svtiger543939@gmail.com</a>
              </div>
            </div>
          </div>
        </Card>
      </div>
      <div style={{ marginTop:16, display:"flex", gap:8 }}>
        <Btn onClick={save} icon={saved?"✓":"💾"} variant={saved?"success":"primary"}>{saved?"Saved!":"Save Settings"}</Btn>
      </div>
    </div>
  );
}

// ── Role helpers (module-level so all screens can use them) ────────
const DB_ROLES    = ["user","accountant","ca","admin"];
const ROLE_LABELS = { user:"User", accountant:"Accountant", ca:"CA", admin:"Admin" };
const toDbRole   = r => (r||"user").toLowerCase().trim();
const fromDbRole = r => ROLE_LABELS[String(r||"").toLowerCase()] || r || "user";
const roleColor  = r => { const n=String(r||"").toLowerCase(); return n==="admin"?"red":n==="ca"?"purple":n==="accountant"?"blue":"gray"; };

// ══════════════════════════════════════════════════════════════════
// SCREEN: User Management (Admin only)
// ══════════════════════════════════════════════════════════════════
function UserManagementScreen({ adminUser }) {
  const [users, setUsers]           = useState([]);
  const [loading, setLoading]       = useState(true);
  const [search, setSearch]         = useState("");
  const [filterRole, setFilterRole] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [actioning, setActioning]   = useState(null);
  const [toast_, setToast_]         = useState({ msg:"", type:"success" });
  const [confirmDel, setConfirmDel] = useState(null);
  const [viewUser, setViewUser]     = useState(null);
  const [editUser, setEditUser]     = useState(null);   // ← NEW: edit modal
  const [editForm, setEditForm]     = useState({});     // ← NEW: edit form state
  const [editErr, setEditErr]       = useState("");
  const [editLoading, setEditLoading] = useState(false);
  const [resetModal, setResetModal] = useState(null);
  const [newPass, setNewPass]       = useState("");
  const [newPassErr, setNewPassErr] = useState("");
  const [addModal, setAddModal]     = useState(false);
  const [addForm, setAddForm]       = useState({ name:"", email:"", role:"user", company:"", password:"" });
  const [addErr, setAddErr]         = useState("");
  const [addLoading, setAddLoading] = useState(false);
  const [tab, setTab] = useState("users");

  const ROLES = DB_ROLES; // alias for JSX below

  const statusColor = s => s==="approved"?"green":s==="pending"?"amber":s==="on_hold"?"purple":s==="rejected"?"red":"gray";

  const notify = (msg, type="success") => {
    setToast_({ msg, type });
    setTimeout(() => setToast_(t => t.msg===msg?{msg:"",type:"success"}:t), 3500);
  };

  // ── Load all profiles + backfill missing emails ────────────────
  const loadUsers = useCallback(async () => {
    setLoading(true);
    try {
      // Try with email column first; fall back gracefully if column doesn't exist yet
      let profiles;
      try {
        profiles = await sb.from("profiles", "select=id,name,email,role,status,company,avatar,mobile,mobile_verified,created_at,approved_by,approved_at&order=created_at.asc&limit=500");
      } catch(emailErr) {
        if (/email.*not exist|column.*email/i.test(emailErr.message)) {
          // email column not yet added — fetch without it and notify admin
          profiles = await sb.from("profiles", "select=id,name,role,status,company,avatar,mobile,mobile_verified,created_at,approved_by,approved_at&order=created_at.asc&limit=500");
          notify("⚠ profiles.email column missing. Run in Supabase SQL Editor: ALTER TABLE profiles ADD COLUMN IF NOT EXISTS email text; — then refresh.", "error");
        } else throw emailErr;
      }

      // Fetch emails from auth admin endpoint (requires admin JWT, best-effort)
      let authEmailMap = {};
      try {
        const res = await fetch(`${SUPABASE_URL}/auth/v1/admin/users?per_page=1000`, {
          headers: {
            "apikey": SUPABASE_ANON,
            "Authorization": `Bearer ${sb._token||SUPABASE_ANON}`,
            "Content-Type": "application/json",
          }
        });
        if (res.ok) {
          const d = await res.json();
          (d.users||d||[]).forEach(u => { if(u.id&&u.email) authEmailMap[u.id]=u.email; });
        }
      } catch {}

      // Also get current signed-in user's own email as a guaranteed source
      try {
        const me = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
          headers: { "apikey": SUPABASE_ANON, "Authorization": `Bearer ${sb._token}` }
        });
        if (me.ok) { const d=await me.json(); if(d.id&&d.email) authEmailMap[d.id]=d.email; }
      } catch {}

      const enriched = profiles.map(p => ({
        ...p,
        email:   p.email || authEmailMap[p.id] || "",
        role:    p.role || "user",
        avatar:  p.avatar || (p.name||"?").slice(0,2).toUpperCase(),
        company: (p.company && isNaN(String(p.company).trim())) ? p.company : "",
        _emailMissing: !p.email && !authEmailMap[p.id],
      }));

      // Auto-backfill: write resolved emails back to profiles table so future loads work
      const needsBackfill = enriched.filter(p => !profiles.find(x=>x.id===p.id)?.email && p.email);
      for (const p of needsBackfill) {
        try { await sb.update("profiles", { id: p.id }, { email: p.email }); } catch {}
      }

      setUsers(enriched);
    } catch (e) {
      const msg = e.message || "";
      if (/infinite recursion/i.test(msg)) {
        notify("RLS policy error — run the SQL fix in Settings to resolve.", "error");
      } else if (/406|not acceptable/i.test(msg)) {
        notify("Profiles table RLS is blocking access. Run the SQL fix from the audit report.", "error");
      } else {
        notify("Error loading users: "+msg, "error");
      }
    }
    setLoading(false);
  }, []);

  useEffect(() => { loadUsers(); }, [loadUsers]);

  // ── Helper: patch a user locally + in viewUser/editUser if open ─
  const patchUser = (id, changes) => {
    setUsers(us => us.map(x => x.id===id ? {...x,...changes} : x));
    setViewUser(v => v && v.id===id ? {...v,...changes} : v);
    setEditUser(v => v && v.id===id ? {...v,...changes} : v);
  };

  // ── Hold / Unhold ──────────────────────────────────────────────
  const toggleHold = async (u) => {
    const newStatus = u.status==="on_hold" ? "approved" : "on_hold";
    setActioning(u.id);
    try {
      await sb.update("profiles", { id: u.id }, { status: newStatus, updated_at: new Date().toISOString() });
      patchUser(u.id, { status: newStatus });
      notify(`${u.name} ${newStatus==="on_hold"?"put on hold":"reactivated"}`);
    } catch (e) { notify("Hold error: "+e.message, "error"); }
    setActioning(null);
  };

  // ── Change Role ────────────────────────────────────────────────
  const changeRole = async (u, role) => {
    const dbRole = toDbRole(role);
    if (dbRole === toDbRole(u.role)) return;
    setActioning(u.id);
    try {
      await sb.update("profiles", { id: u.id }, { role: dbRole, updated_at: new Date().toISOString() });
      patchUser(u.id, { role: dbRole });
      notify(`${u.name}'s role changed to ${fromDbRole(dbRole)}`);
    } catch (e) {
      if (e.message?.includes("role_check") || e.message?.includes("check constraint")) {
        notify(
          `DB constraint error — run this SQL in Supabase:\n` +
          `ALTER TABLE profiles DROP CONSTRAINT profiles_role_check;\n` +
          `ALTER TABLE profiles ADD CONSTRAINT profiles_role_check CHECK (role IN ('admin','user','accountant','ca'));`,
          "error"
        );
      } else {
        notify("Role change error: "+e.message, "error");
      }
    }
    setActioning(null);
  };

  // ── Edit / Save Profile ────────────────────────────────────────
  const openEditUser = (u) => {
    setEditUser(u);
    setEditForm({ name: u.name||"", email: u.email||"", company: u.company||"", role: toDbRole(u.role||"user"), status: u.status||"approved" });
    setEditErr("");
  };

  const saveEditUser = async () => {
    if (!editForm.name.trim()) return setEditErr("Name is required.");
    setEditLoading(true); setEditErr("");
    try {
      if (!editForm.email.trim() || !editForm.email.includes("@"))
        return setEditErr("A valid email address is required.");
      const payload = {
        name:       editForm.name.trim(),
        email:      editForm.email.trim().toLowerCase(),
        company:    editForm.company.trim(),
        role:       toDbRole(editForm.role),   // normalise to DB lowercase
        status:     editForm.status,
        avatar:     editForm.name.trim().slice(0,2).toUpperCase(),
        updated_at: new Date().toISOString(),
      };
      await sb.update("profiles", { id: editUser.id }, payload);
      patchUser(editUser.id, payload);
      notify(`${editForm.name} profile updated`);
      setEditUser(null);
    } catch (e) { setEditErr("Save failed: "+e.message); }
    setEditLoading(false);
  };

  // ── Delete ─────────────────────────────────────────────────────
  const deleteUser = async (u) => {
    setActioning(u.id);
    try {
      await sb.update("profiles", { id: u.id }, { status: "deleted", updated_at: new Date().toISOString() });
      // Hard delete profile row
      const q = `id=eq.${encodeURIComponent(u.id)}`;
      const res = await fetch(`${SUPABASE_URL}/rest/v1/profiles?${q}`, {
        method: "DELETE",
        headers: { ...sb._headers(), "Prefer": "return=minimal" },
      });
      if (!res.ok && res.status !== 204) throw new Error(`Delete failed (${res.status})`);
      setUsers(us => us.filter(x => x.id!==u.id));
      notify(`${u.name} deleted`);
    } catch (e) { notify("Delete error: "+e.message, "error"); }
    setActioning(null);
    setConfirmDel(null);
  };

  // ── Reset Password ─────────────────────────────────────────────
  const sendPasswordReset = async (u) => {
    setNewPassErr("");
    try {
      const res = await fetch(`${SUPABASE_URL}/auth/v1/recover`, {
        method: "POST",
        headers: { "Content-Type":"application/json", "apikey": SUPABASE_ANON },
        body: JSON.stringify({ email: u.email }),
      });
      if (!res.ok) { const d=await res.json(); throw new Error(d.error_description||d.message||"Failed"); }
      notify(`Password reset email sent to ${u.email}`);
      setResetModal(null);
    } catch (e) { setNewPassErr(e.message); }
  };

  // ── Add User ───────────────────────────────────────────────────
  const handleAddUser = async () => {
    setAddErr("");
    if (!addForm.name.trim())         return setAddErr("Name is required.");
    if (!addForm.email.includes("@")) return setAddErr("Valid email required.");
    if (addForm.password.length < 8)  return setAddErr("Password must be at least 8 characters.");
    // Check if email already exists in profiles
    setAddLoading(true);
    try {
      const existing = await sb.from("profiles", `email=eq.${encodeURIComponent(addForm.email.trim().toLowerCase())}&select=id`);
      if (existing && existing.length > 0) { setAddErr("A user with this email already exists."); setAddLoading(false); return; }
    } catch {}
    try {
      const dbRole = toDbRole(addForm.role);
      const session = await sb.signUp(addForm.email, addForm.password, {
        name: addForm.name.trim(), role: dbRole, company: addForm.company.trim(),
      });
      const userId = session.user?.id;
      if (!userId) throw new Error("User created but no ID returned — check Supabase email confirmation settings.");
      await sb.insert("profiles", {
        id:      userId,
        name:    addForm.name.trim(),
        email:   addForm.email.trim().toLowerCase(),
        role:    dbRole,                               // always lowercase to satisfy constraint
        company: addForm.company.trim(),
        status:  "approved",
        avatar:  addForm.name.trim().slice(0,2).toUpperCase(),
      });
      notify(`User ${addForm.name} created successfully`);
      setAddModal(false);
      setAddForm({ name:"", email:"", role:"user", company:"", password:"" });
      loadUsers();
    } catch (e) { setAddErr(e.message); }
    setAddLoading(false);
  };

  // ── Filters ────────────────────────────────────────────────────
  const filtered = users.filter(u => {
    const q = search.toLowerCase();
    const matchSearch = !search || u.name?.toLowerCase().includes(q) || u.email?.toLowerCase().includes(q) || u.company?.toLowerCase().includes(q);
    const matchRole   = filterRole==="all"   || u.role===filterRole;
    const matchStatus = filterStatus==="all" || u.status===filterStatus;
    return matchSearch && matchRole && matchStatus;
  });

  return (
    <div className="fade-in">
      {/* Toast */}
      {toast_.msg && (
        <div style={{ position:"fixed", top:20, right:20, zIndex:9999, background:T.card, border:`1px solid ${toast_.type==="error"?T.red:T.green}55`, borderRadius:11, padding:"12px 18px", fontSize:13, color:toast_.type==="error"?T.red:T.green, boxShadow:"0 8px 32px rgba(0,0,0,0.4)", display:"flex", alignItems:"center", gap:8 }} className="fade-in">
          <span>{toast_.type==="error"?"✕":"✓"}</span>
          <span style={{ color:T.text }}>{toast_.msg}</span>
        </div>
      )}

      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18 }}>
        <div>
          <h2 style={{ fontSize:22, fontWeight:900, letterSpacing:"-0.5px", color:T.text }}>User Management</h2>
          <p style={{ color:T.textMid, fontSize:13, marginTop:3 }}>Admin-only panel · full user control</p>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          {/* Backfill Emails — fixes users created before email was stored in profiles */}
          {users.some(u => u._emailMissing) && (
            <Btn size="sm" variant="ghost" icon="✉" onClick={async () => {
              notify("Fetching emails from auth…", "success");
              await loadUsers();
              notify("Email backfill complete — refresh to verify", "success");
            }}>Backfill Emails ({users.filter(u=>u._emailMissing).length})</Btn>
          )}
          {tab==="users" && <Btn icon="+" onClick={()=>setAddModal(true)}>Add User</Btn>}
        </div>
      </div>

      {/* ── SQL Migration Notice (shown once if constraint error is likely) ── */}
      {users.some(u => u.role && !["admin","user","accountant","ca"].includes(String(u.role).toLowerCase())) && (
        <div style={{ background:T.amberDim, border:`1px solid ${T.amber}44`, borderRadius:10, padding:"12px 16px", marginBottom:16, fontSize:12, color:T.amber, lineHeight:1.7 }}>
          ⚠ <strong>Role constraint mismatch detected.</strong> Some users have roles not allowed by your DB constraint.<br/>
          Run this SQL once in <strong>Supabase → SQL Editor</strong> to fix:
          <pre style={{ background:"rgba(0,0,0,0.3)", borderRadius:6, padding:"8px 12px", marginTop:8, fontSize:11, color:T.text, overflowX:"auto", userSelect:"all" }}>
{`ALTER TABLE profiles DROP CONSTRAINT IF EXISTS profiles_role_check;
ALTER TABLE profiles ADD CONSTRAINT profiles_role_check
  CHECK (role IN ('admin','user','accountant','ca'));`}
          </pre>
        </div>
      )}

      {/* Tabs */}
      <div style={{ display:"flex", borderBottom:`2px solid ${T.border}`, marginBottom:20, gap:0 }}>
        {[["users","Users"],["approvals","Approvals"]].map(([t,label]) => (
          <button key={t} onClick={()=>setTab(t)}
            style={{ padding:"10px 22px", border:"none", borderBottom: tab===t ? `2px solid ${T.accent}` : "2px solid transparent", marginBottom:"-2px", cursor:"pointer", fontFamily:T.font, fontSize:13, fontWeight:tab===t?700:500, transition:"all 0.15s", background:"transparent", color:tab===t?T.accent:T.textSub, letterSpacing:"-0.1px" }}>
            {label}
          </button>
        ))}
      </div>

      {tab==="approvals" && <AdminApprovalPanel user={adminUser} onClose={()=>{}} />}
      {tab==="users" && (
      <div>
      {/* Stats row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:20 }}>
        {[
          ["Total Users",  users.length,                                        T.accent],
          ["Active",       users.filter(u=>u.status==="approved").length,        T.green],
          ["Pending",      users.filter(u=>u.status==="pending").length,         T.amber],
          ["On Hold",      users.filter(u=>u.status==="on_hold").length,         T.purple],
        ].map(([label,val,color]) => (
          <Card key={label} style={{ padding:"16px 18px", borderLeft:`3px solid ${color}` }}>
            <div style={{ fontSize:24, fontWeight:800, color, letterSpacing:"-0.5px", marginBottom:3 }}>{val}</div>
            <div style={{ fontSize:12, fontWeight:600, color:T.text }}>{label}</div>
          </Card>
        ))}
      </div>

      {/* Filters */}
      <Card style={{ marginBottom:16, padding:"14px 18px" }}>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          <div style={{ flex:1, minWidth:200 }}>
            <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, email, company…" prefix="" />
          </div>
          <select value={filterRole} onChange={e=>setFilterRole(e.target.value)} style={{ padding:"8px 12px", borderRadius:8, fontSize:12, border:`1px solid ${T.border}`, background:T.surface, color:T.text, minWidth:120 }}>
            <option value="all">All Roles</option>
            {ROLES.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
          <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{ padding:"8px 12px", borderRadius:8, fontSize:12, border:`1px solid ${T.border}`, background:T.surface, color:T.text, minWidth:130 }}>
            <option value="all">All Statuses</option>
            {["approved","pending","on_hold","rejected"].map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <Btn size="sm" variant="ghost" icon="↺" onClick={loadUsers}>Refresh</Btn>
        </div>
      </Card>

      {/* Table */}
      <Card style={{ padding:0, overflow:"hidden" }}>
        {loading ? (
          <div style={{ textAlign:"center", padding:"60px 0", color:T.textDim }}>
            <span style={{ fontSize:13, color:T.textSub }}>Loading…</span>
            <p style={{ marginTop:12, fontSize:13 }}>Loading users…</p>
          </div>
        ) : filtered.length===0 ? (
          <div style={{ textAlign:"center", padding:"60px 0", color:T.textDim }}>
            <div style={{ fontSize:36, marginBottom:10 }}>""</div>
            {users.length===0 ? (
              <>
                <p style={{ fontSize:14, fontWeight:600, color:T.text, marginBottom:6 }}>No users loaded</p>
                <p style={{ fontSize:12, color:T.textDim, marginBottom:16 }}>
                  This is usually caused by an RLS (Row Level Security) policy error in Supabase.
                </p>
                <div style={{ display:"flex", gap:8, justifyContent:"center" }}>
                  <Btn size="sm" variant="secondary" icon="↺" onClick={loadUsers}>Retry</Btn>
                  <Btn size="sm" variant="primary" icon="→" onClick={() => {
                    const sql = `-- Fix RLS infinite recursion\nDROP POLICY IF EXISTS "Admins can read all profiles" ON profiles;\nDROP POLICY IF EXISTS "Admins can update profiles" ON profiles;\nCREATE POLICY "Admins can read all profiles" ON profiles FOR SELECT USING ((auth.jwt() -> 'user_metadata' ->> 'role') = 'admin' OR auth.uid() = id);\nCREATE POLICY "Admins can update profiles" ON profiles FOR UPDATE USING ((auth.jwt() -> 'user_metadata' ->> 'role') = 'admin');`;
                    navigator.clipboard?.writeText(sql);
                    alert("SQL copied to clipboard!\n\nPaste and run it in:\nSupabase Dashboard → SQL Editor");
                  }}>Copy RLS Fix SQL</Btn>
                </div>
              </>
            ) : (
              <p style={{ fontSize:14 }}>No users match your filters</p>
            )}
          </div>
        ) : (
          <>
            <div style={{ display:"grid", gridTemplateColumns:"2fr 1.5fr 1fr 1fr 1.2fr 2fr", gap:0, padding:"10px 20px", borderBottom:`1px solid ${T.border}`, background:T.surface }}>
              {["User","Company","Role","Status","Mobile","Actions"].map(h => (
                <span key={h} style={{ fontSize:11, fontWeight:600, color:T.textDim, textTransform:"uppercase", letterSpacing:"0.06em" }}>{h}</span>
              ))}
            </div>
            {filtered.map((u, idx) => (
              <div key={u.id} className="row-hover" style={{ display:"grid", gridTemplateColumns:"2fr 1.5fr 1fr 1fr 1.2fr 2fr", gap:0, padding:"13px 20px", borderBottom:idx<filtered.length-1?`1px solid ${T.border}`:"none", alignItems:"center", transition:"background 0.15s" }}>
                {/* User */}
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ width:36, height:36, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, color:"#fff", flexShrink:0 }}>
                    {(u.avatar||(u.name||"?").slice(0,2)).toUpperCase()}
                  </div>
                  <div style={{ minWidth:0 }}>
                    <div style={{ fontSize:13, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.name||"—"}</div>
                    <div style={{ fontSize:11, color:T.textDim, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.email||"—"}</div>
                  </div>
                </div>
                {/* Company */}
                <div style={{ fontSize:12, color:T.textMid, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.company||<span style={{color:T.textDim}}>—</span>}</div>
                {/* Role */}
                <div><Pill color={roleColor(u.role)} size="xs">{fromDbRole(u.role)}</Pill></div>
                {/* Status */}
                <div><Pill color={statusColor(u.status)} size="xs" dot>{u.status||"unknown"}</Pill></div>
                {/* Mobile */}
                <div style={{ fontSize:11 }}>
                  {u.mobile ? (
                    <span style={{ color: u.mobile_verified ? T.green : T.amber||"#f59e0b" }}>
                      {u.mobile_verified ? "✓ " : "⚠ "}{u.mobile}
                    </span>
                  ) : <span style={{ color:T.textDim }}>—</span>}
                </div>
                {/* Actions */}
                <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                  {/* View */}
                  <button title="View Profile" disabled={actioning===u.id}
                    onClick={() => setViewUser(u)}
                    style={{ padding:"4px 10px", borderRadius:6, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.border}`, background:T.surface, color:T.textMid }}>
                    View
                  </button>
                  {/* Edit */}
                  <button title="Edit Profile" disabled={actioning===u.id || u.id===adminUser.id}
                    onClick={() => openEditUser(u)}
                    style={{ padding:"4px 10px", borderRadius:6, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:(u.id===adminUser.id)?"not-allowed":"pointer", border:`1px solid ${T.accent}44`, background:T.accentDim, color:T.accent, opacity:u.id===adminUser.id?0.45:1 }}>
                    Edit
                  </button>
                  {/* Hold / Unhold */}
                  <button title={u.status==="on_hold"?"Unhold":"Put on Hold"} disabled={actioning===u.id||u.id===adminUser.id}
                    onClick={() => toggleHold(u)}
                    style={{ padding:"4px 10px", borderRadius:6, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:(actioning===u.id||u.id===adminUser.id)?"not-allowed":"pointer", border:`1px solid ${u.status==="on_hold"?T.green+"66":T.purple+"66"}`, background:u.status==="on_hold"?T.greenDim:T.purpleDim, color:u.status==="on_hold"?T.green:T.purple, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1 }}>
                    {actioning===u.id ? "…" : u.status==="on_hold" ? "Unhold" : "Hold"}
                  </button>
                  {/* Reset Password */}
                  <button title="Reset Password" disabled={actioning===u.id}
                    onClick={() => { setResetModal(u); setNewPassErr(""); setNewPass(""); }}
                    style={{ padding:"4px 10px", borderRadius:6, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.amber}44`, background:T.amberDim, color:T.amber }}>
                    Pwd Reset
                  </button>
                  {/* Change Role inline */}
                  <select title="Change Role" disabled={actioning===u.id||u.id===adminUser.id}
                    value={toDbRole(u.role||"user")}
                    onChange={e => changeRole(u, e.target.value)}
                    style={{ padding:"4px 6px", borderRadius:7, fontSize:11, fontFamily:T.font, border:`1px solid ${T.border}`, background:T.surface, color:T.textMid, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1, cursor:"pointer" }}>
                    {ROLES.map(r => <option key={r} value={r}>{ROLE_LABELS[r]||r}</option>)}
                  </select>
                  {/* Delete */}
                  <button title="Delete User" disabled={actioning===u.id||u.id===adminUser.id}
                    onClick={() => setConfirmDel(u)}
                    style={{ padding:"5px 9px", borderRadius:7, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:(actioning===u.id||u.id===adminUser.id)?"not-allowed":"pointer", border:`1px solid ${T.red}44`, background:T.redDim, color:T.red, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1 }}>
                    🗑
                  </button>
                </div>
              </div>
            ))}
          </>
        )}
      </Card>
      </div>)}

      {/* ── Delete confirm ──────────────────────────────────────── */}
      <Modal open={!!confirmDel} onClose={() => setConfirmDel(null)} title="⚠ Confirm Delete" width={420}>
        {confirmDel && (
          <div>
            <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:10, padding:"14px 16px", marginBottom:18 }}>
              <p style={{ fontSize:13, color:T.textMid, lineHeight:1.7 }}>
                You are about to <strong style={{color:T.red}}>permanently delete</strong> the account for:<br/>
                <strong style={{color:T.text}}>{confirmDel.name}</strong> ({confirmDel.email||confirmDel.id})<br/>
                This cannot be undone.
              </p>
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <Btn variant="secondary" fullWidth onClick={() => setConfirmDel(null)}>Cancel</Btn>
              <Btn variant="danger" fullWidth icon="🗑" onClick={() => deleteUser(confirmDel)} disabled={actioning===confirmDel.id}>
                {actioning===confirmDel.id ? "Deleting…" : "Delete Permanently"}
              </Btn>
            </div>
          </div>
        )}
      </Modal>

      {/* ── Reset Password ──────────────────────────────────────── */}
      <Modal open={!!resetModal} onClose={() => setResetModal(null)} title="🔑 Reset Password" width={440}>
        {resetModal && (() => {
          const effectiveEmail = resetModal.email || manualEmail.trim();
          return (
            <div>
              <div style={{ background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:10, padding:"12px 16px", marginBottom:14 }}>
                <p style={{ fontSize:13, color:T.textMid, lineHeight:1.7 }}>
                  Send a password-reset link to <strong style={{color:T.text}}>{resetModal.name}</strong>
                </p>
                {resetModal.email
                  ? <p style={{ fontSize:12, color:T.green, marginTop:4 }}>✓ {resetModal.email}</p>
                  : <p style={{ fontSize:12, color:"#f59e0b", marginTop:4 }}>⚠ No email stored in profile</p>
                }
              </div>

              {/* If email missing — let admin enter it manually */}
              {!resetModal.email && (
                <div style={{ marginBottom:14 }}>
                  <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:6 }}>
                    Enter email address for this user
                  </label>
                  <Input
                    value={manualEmail}
                    onChange={e => setManualEmail(e.target.value)}
                    placeholder="user@example.com"
                    prefix="✉"
                    type="email"
                  />
                  <p style={{ fontSize:11, color:T.textDim, marginTop:5 }}>
                    This will also be saved to the user's profile for future use.
                  </p>
                </div>
              )}

              {newPassErr && <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"9px 13px", fontSize:12, color:T.red, marginBottom:12 }}>✕ {newPassErr}</div>}

              <div style={{ display:"flex", gap:10 }}>
                <Btn variant="secondary" fullWidth onClick={() => setResetModal(null)}>Cancel</Btn>
                <Btn
                  variant="primary" fullWidth icon="📧"
                  disabled={!effectiveEmail || !effectiveEmail.includes("@")}
                  onClick={async () => {
                    // If admin typed a manual email, save it to profiles first
                    if (!resetModal.email && manualEmail.trim()) {
                      try {
                        await sb.update("profiles", { id: resetModal.id }, { email: manualEmail.trim().toLowerCase() });
                        patchUser(resetModal.id, { email: manualEmail.trim().toLowerCase(), _emailMissing: false });
                        setResetModal(m => ({ ...m, email: manualEmail.trim().toLowerCase() }));
                      } catch(e) { setNewPassErr("Could not save email: "+e.message); return; }
                    }
                    sendPasswordReset({ ...resetModal, email: effectiveEmail });
                  }}
                >
                  Send Reset Email
                </Btn>
              </div>
            </div>
          );
        })()}
      </Modal>

      {/* ── View User Profile ────────────────────────────────────── */}
      <Modal open={!!viewUser} onClose={() => setViewUser(null)} title="👤 User Profile" width={480}>
        {viewUser && (
          <div>
            <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:20, padding:"16px", background:T.surface, borderRadius:12 }}>
              <div style={{ width:60, height:60, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, fontWeight:700, color:"#fff", flexShrink:0 }}>
                {(viewUser.avatar||(viewUser.name||"?").slice(0,2)).toUpperCase()}
              </div>
              <div>
                <div style={{ fontSize:17, fontWeight:700, color:T.text, marginBottom:4 }}>{viewUser.name}</div>
                <div style={{ fontSize:12, color:T.textDim, marginBottom:6 }}>{viewUser.email||<span style={{fontStyle:"italic"}}>no email on file</span>}</div>
                <div style={{ display:"flex", gap:6 }}>
                  <Pill color={roleColor(viewUser.role)} size="xs">{fromDbRole(viewUser.role)}</Pill>
                  <Pill color={statusColor(viewUser.status)} size="xs" dot>{viewUser.status||"unknown"}</Pill>
                </div>
              </div>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:8, fontSize:13, marginBottom:16 }}>
              {[
                ["🏢","Company",    viewUser.company||"—"],
                ["✉","Email",      viewUser.email||"—"],
                ["📱","Mobile",     viewUser.mobile ? `${viewUser.mobile}${viewUser.mobile_verified?" ✓ Verified":" ⚠ Unverified"}` : "—"],
                ["🪪","User ID",    viewUser.id],
                ["📅","Joined",     viewUser.created_at ? new Date(viewUser.created_at).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}) : "—"],
                ["✅","Approved By",users.find(x=>x.id===viewUser.approved_by)?.name || viewUser.approved_by || "—"],
                ["🕒","Approved At",viewUser.approved_at ? new Date(viewUser.approved_at).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}) : "—"],
              ].map(([icon,label,val]) => (
                <div key={label} style={{ display:"flex", justifyContent:"space-between", padding:"8px 12px", background:T.surface, borderRadius:8 }}>
                  <span style={{ color:T.textDim }}>{icon} {label}</span>
                  <span style={{ color:T.text, fontWeight:500, wordBreak:"break-all", maxWidth:250, textAlign:"right" }}>{val}</span>
                </div>
              ))}
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <Btn variant="secondary" fullWidth onClick={() => setViewUser(null)}>Close</Btn>
              <Btn variant="outline" fullWidth icon="✏" onClick={() => { setViewUser(null); openEditUser(viewUser); }}>Edit</Btn>
              <Btn variant="outline" fullWidth icon="🔑" onClick={() => { setViewUser(null); setResetModal(viewUser); setNewPassErr(""); }}>Reset Pwd</Btn>
              {viewUser.id !== adminUser.id && (
                viewUser.status==="on_hold"
                  ? <Btn variant="success" fullWidth icon="▶" onClick={() => { toggleHold(viewUser); setViewUser(null); }}>Reactivate</Btn>
                  : <Btn variant="amber" fullWidth icon="⏸" onClick={() => { toggleHold(viewUser); setViewUser(null); }}>Hold</Btn>
              )}
            </div>
          </div>
        )}
      </Modal>

      {/* ── Edit User Profile ────────────────────────────────────── */}
      <Modal open={!!editUser} onClose={() => setEditUser(null)} title="✏ Edit User Profile" width={480}>
        {editUser && (
          <div style={{ display:"flex", flexDirection:"column", gap:13 }}>
            {editErr && <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"9px 13px", fontSize:12, color:T.red }}>✕ {editErr}</div>}
            <div>
              <label style={{ fontSize:12, color:T.textDim, display:"block", marginBottom:5 }}>Full Name *</label>
              <Input value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))} placeholder="e.g. Priya Sharma" prefix="👤" />
            </div>
            <div>
              <label style={{ fontSize:12, color:T.textDim, display:"block", marginBottom:5 }}>
                Email Address *
                {editUser && !editUser.email && <span style={{color:"#f59e0b",marginLeft:6,fontSize:11}}>⚠ Missing — enter to enable password reset</span>}
              </label>
              <Input
                value={editForm.email}
                onChange={e=>setEditForm(f=>({...f,email:e.target.value}))}
                placeholder="user@example.com"
                prefix="✉"
                type="email"
              />
            </div>
            <div>
              <label style={{ fontSize:12, color:T.textDim, display:"block", marginBottom:5 }}>Company</label>
              <Input value={editForm.company} onChange={e=>setEditForm(f=>({...f,company:e.target.value}))} placeholder="e.g. Acme Corp Pvt Ltd" prefix="🏢" />
            </div>
            <div>
              <label style={{ fontSize:12, color:T.textDim, display:"block", marginBottom:5 }}>Role</label>
              <select value={editForm.role} onChange={e=>setEditForm(f=>({...f,role:e.target.value}))}
                style={{ width:"100%", padding:"9px 12px", borderRadius:8, fontSize:13, border:`1px solid ${T.border}`, background:T.surface, color:T.text }}>
                {ROLES.map(r => <option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize:12, color:T.textDim, display:"block", marginBottom:5 }}>Status</label>
              <select value={editForm.status} onChange={e=>setEditForm(f=>({...f,status:e.target.value}))}
                style={{ width:"100%", padding:"9px 12px", borderRadius:8, fontSize:13, border:`1px solid ${T.border}`, background:T.surface, color:T.text }}>
                {["approved","pending","on_hold","rejected"].map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div style={{ display:"flex", gap:10, marginTop:4 }}>
              <Btn variant="secondary" fullWidth onClick={() => setEditUser(null)}>Cancel</Btn>
              <Btn variant="primary" fullWidth icon="💾" onClick={saveEditUser} disabled={editLoading}>
                {editLoading ? "Saving…" : "Save Changes"}
              </Btn>
            </div>
          </div>
        )}
      </Modal>

      {/* ── Add User ────────────────────────────────────────────── */}
      <Modal open={addModal} onClose={() => { setAddModal(false); setAddErr(""); }} title="➕ Add New User" width={480}>
        <div style={{ display:"flex", flexDirection:"column", gap:13 }}>
          {addErr && (
            <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"9px 13px", fontSize:12, color:T.red }}>✕ {addErr}</div>
          )}
          <div>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Full Name *</label>
            <Input value={addForm.name} onChange={e=>setAddForm(f=>({...f,name:e.target.value}))} placeholder="e.g. Priya Sharma" prefix="👤" />
          </div>
          <div>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Email Address *</label>
            <Input value={addForm.email} onChange={e=>setAddForm(f=>({...f,email:e.target.value}))} placeholder="user@company.in" prefix="✉" />
          </div>
          <div>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Temporary Password * (min 8 chars)</label>
            <Input value={addForm.password} onChange={e=>setAddForm(f=>({...f,password:e.target.value}))} placeholder="••••••••" prefix="🔑" />
          </div>
          <div>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Role</label>
            <select value={addForm.role} onChange={e=>setAddForm(f=>({...f,role:e.target.value}))}
              style={{ width:"100%", padding:"9px 12px", borderRadius:8, fontSize:13, border:`1px solid ${T.border}`, background:T.surface, color:T.text }}>
              {ROLES.map(r=><option key={r} value={r}>{ROLE_LABELS[r]||r}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize:12, color:T.textMid, display:"block", marginBottom:5 }}>Company</label>
            <Input value={addForm.company} onChange={e=>setAddForm(f=>({...f,company:e.target.value}))} placeholder="e.g. Acme Corp Pvt Ltd" prefix="🏢" />
          </div>
          <div style={{ display:"flex", gap:10, marginTop:4 }}>
            <Btn variant="secondary" fullWidth onClick={() => { setAddModal(false); setAddErr(""); }}>Cancel</Btn>
            <Btn variant="primary" fullWidth icon="+" disabled={addLoading} onClick={handleAddUser}>
              {addLoading ? "Creating…" : "Create User"}
            </Btn>
          </div>
        </div>
      </Modal>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// ROOT APP
// ══════════════════════════════════════════════════════════════════
const INITIAL_HISTORY = []; // No demo data — each user sees only their own imports

// ── Error Boundary — catches uncaught render errors ──────────────
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { console.error("Bank2Tally render error:", error, info); }
  render() {
    if (this.state.error) {
      return (
        <div style={{ minHeight:"100vh", background:"#f0f4f9", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"sans-serif", color:"#0f172a" }}>
          <div style={{ maxWidth:480, textAlign:"center", padding:32 }}>
            <div style={{ fontSize:48, marginBottom:16 }}>⚠</div>
            <h2 style={{ fontSize:20, fontWeight:700, marginBottom:12, color:"#dc2626" }}>Something went wrong</h2>
            <p style={{ color:"#64748b", lineHeight:1.7, marginBottom:20 }}>{this.state.error?.message || "An unexpected error occurred."}</p>
            <button onClick={() => { this.setState({ error:null }); window.location.reload(); }}
              style={{ padding:"10px 24px", background:"#2563eb", color:"#fff", border:"none", borderRadius:8, fontSize:14, fontWeight:600, cursor:"pointer" }}>
              Reload App
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}


function AppInner() {
  const [user, setUser] = useState(null);
  const [screen, setScreen] = useState(SCREENS.LOGIN);
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [filename, setFilename] = useState("");
  const [templateKey, setTemplateKey] = useState("");
  const [mapping, setMapping] = useState({});
  const [rows, setRows] = useState([]);
  const [selectedCompanies, setSelectedCompanies] = useState([]);
  const [history, setHistory] = useState(INITIAL_HISTORY); // always starts empty
  const [auditLog, setAuditLog] = useState([]);
  const [toasts, setToasts] = useState([]);
  const [tallyHost, setTallyHost] = useState("localhost");
  const [tallyPort, setTallyPort] = useState("9000");
  const [defaultLedger, setDefaultLedger] = useState("Suspense Account");
  const [autoDetectLedger, setAutoDetectLedger] = useState(true);

  // Live Tally gateway
  const tally = useTallyGateway(tallyHost, tallyPort);

  const toast = (msg, type="success") => {
    const id = genId();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),4000);
  };

  const [pendingCount, setPendingCount] = useState(0);

  // Load user-specific import history — auto-purge entries older than 7 days
  const loadHistory = (userId) => {
    try {
      const key = `import_history_${userId}`;
      const stored = localStorage.getItem(key);
      if (!stored) { setHistory([]); return; }
      const parsed = JSON.parse(stored);
      const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000; // 7 days ago
      const fresh  = parsed.filter(h => {
        const ts = h.savedAt ? new Date(h.savedAt).getTime() : new Date(h.rawDate || 0).getTime();
        return ts > cutoff;
      });
      // If some entries were purged, save the cleaned list
      if (fresh.length !== parsed.length) {
        try { localStorage.setItem(key, JSON.stringify(fresh)); } catch (e) {
        // Storage full or private mode — remove oldest entry and retry
        try {
          const keys = Object.keys(localStorage).filter(k=>k.startsWith("b2t_hist_"));
          if (keys.length) { localStorage.removeItem(keys[0]); localStorage.setItem(key, JSON.stringify(fresh)); }
        } catch {}
      }
      }
      setHistory(fresh);
    } catch { setHistory([]); }
  };

  // Save import history to localStorage scoped by user
  const saveHistory = (userId, hist) => {
    try {
      localStorage.setItem(`import_history_${userId}`, JSON.stringify(hist));
    } catch {}
  };

  // Delete a single history entry
  const deleteHistoryEntry = (entryId) => {
    setHistory(h => {
      const updated = h.filter(x => x.id !== entryId);
      if (user?.id) saveHistory(user.id, updated);
      return updated;
    });
  };

  // Clear all history for current user
  const clearAllHistory = () => {
    setHistory([]);
    if (user?.id) {
      localStorage.removeItem(`import_history_${user.id}`);
    }
  };

  const onLogin = (u) => {
    setUser(u);
    setScreen(SCREENS.DASHBOARD);
    loadHistory(u.id); // load ONLY this user's history
    if (u.role === "admin") {
      sb.from("approval_requests", "status=eq.pending&select=id")
        .then(rows => setPendingCount(rows.length))
        .catch(() => {});
    }
  };

  const onLogout = async () => {
    await sb.signOut().catch(() => {});
    localStorage.removeItem("sb_session");
    setUser(null);
    setScreen(SCREENS.LOGIN);
    setPendingCount(0);
    // Clear ALL user-specific state so next user starts clean
    setHistory([]);
    setRows([]); setHeaders([]); setRawRows([]); setFilename("");
    setMapping({}); setTemplateKey(""); setAuditLog([]);
    setSelectedCompanies([]);
  };

  // Restore session on mount — runs once, directly sets user state
  useEffect(() => {
    const stored = localStorage.getItem("sb_session");
    if (!stored) return;
    try {
      const session = JSON.parse(stored);
      if (!session?.access_token || !session?.user?.id) {
        localStorage.removeItem("sb_session"); return;
      }
      // Check token expiry
      let payload;
      try { payload = JSON.parse(atob(session.access_token.split(".")[1])); } catch { payload = {}; }
      if (payload.exp && payload.exp * 1000 < Date.now()) {
        localStorage.removeItem("sb_session"); return;
      }
      sb._token = session.access_token;

      // Helper: restore from session metadata alone (fallback)
      const restoreFromSession = async (profile) => {
        const meta = session.user.user_metadata || {};
        const name = profile?.name || meta.name || session.user.email.split("@")[0];
        const role = (profile?.role || meta.role || "user").toLowerCase().trim();
        const status = profile?.status || (role === "admin" ? "approved" : "pending");
        const avatar = profile?.avatar || name.slice(0,2).toUpperCase();
        if (status !== "approved" && status !== "on_hold") { localStorage.removeItem("sb_session"); return; }
        if (status === "on_hold") { localStorage.removeItem("sb_session"); return; }
        // Backfill email into profile if missing
        if (profile && !profile.email && session.user.email) {
          try { await sb.update("profiles", { id: session.user.id }, { email: session.user.email }); } catch {}
        }
        setUser({ id: session.user.id, name, role, status, avatar, company: profile?.company || "", email: session.user.email, mobile: profile?.mobile || "", mobile_verified: profile?.mobile_verified || false, sessionToken: session.access_token });
        setScreen(SCREENS.DASHBOARD);
        if (role === "admin") {
          sb.from("approval_requests", "status=eq.pending&select=id")
            .then(r => setPendingCount(r.length)).catch(() => {});
        }
      };

      sb.from("profiles", "id=eq." + session.user.id + "&select=*")
        .then(profiles => {
          const profile = profiles?.[0];
          restoreFromSession(profile || null);
        })
        .catch(() => {
          // RLS or network error — still try to restore from session metadata
          restoreFromSession(null);
        });
    } catch (e) { localStorage.removeItem("sb_session"); }
  }, []); // eslint-disable-line

  const onParsed = (result, fname, tmplKey) => {
    if (selectedCompanies.length === 0) { toast("Select at least one Tally company first","warn"); return; }

    // ── Auto-detect bank template from parsed headers ─────────────────
    // Compare each template's columns against the actual headers.
    // Score = number of template column names that fuzzy-match a real header.
    // Highest score wins.
    let resolvedKey = tmplKey || "";
    // ── Step 1: Trust the parser's own bank hint (most reliable) ──────
    if (!resolvedKey && result._bankHint && BANK_TEMPLATES[result._bankHint]) {
      resolvedKey = result._bankHint;
    }
    // ── Step 2: Score templates against actual headers (fallback) ─────
    if (!resolvedKey) {
      const norm = s => s.toLowerCase().replace(/[\s_\-\.\(\)\/,]/g,"");
      const hNorms = result.headers.map(h => norm(h));
      let bestScore = 0;
      Object.entries(BANK_TEMPLATES).forEach(([key, tpl]) => {
        const score = Object.values(tpl.cols).filter(colName => {
          const cn = norm(colName);
          return hNorms.some(hn => hn === cn || hn.startsWith(cn.slice(0,10)) || cn.startsWith(hn.slice(0,10)));
        }).length;
        if (score > bestScore) { bestScore = score; resolvedKey = key; }
      });
      if (bestScore < 2) resolvedKey = ""; // not confident enough — let user map manually
    }

    setHeaders(result.headers);
    setRawRows(result.rows);
    setFilename(fname);
    setTemplateKey(resolvedKey);
    setScreen(SCREENS.COLUMN_MAP);
    const bankName = resolvedKey ? BANK_TEMPLATES[resolvedKey]?.name : "";
    toast(
      `Parsed ${result.rows.length} rows from ${fname}${bankName ? ` · ${bankName} detected` : ""}`,
      "success"
    );
  };

  const onMapped = (m) => {
    setMapping(m);
    const built = rawRows.map(r => {
      const get = field => {
        if (!m[field]) return "";
        const idx = headers.indexOf(m[field]);
        return idx >= 0 ? (r[idx] ?? "") : "";
      };

      // ── Amount resolution — handles every format banks emit ──────
      let debit = "", credit = "";

      // Helper: parse any amount string to a positive float
      const parseAmt = raw => {
        const s = String(raw||"").replace(/,/g,"").trim();
        const n = parseFloat(s.replace(/[^0-9.\-]/g,""));
        return isNaN(n) ? 0 : Math.abs(n);
      };
      // Detect DR/CR suffix (e.g. "5000.00 DR", "12345.67CR")
      const getDrCrSuffix = raw => {
        const s = String(raw||"").trim();
        if (/dr$/i.test(s)) return "DR";
        if (/cr$/i.test(s)) return "CR";
        return null;
      };

      if (m.crdr) {
        // Combined amount column
        const raw = String(get("crdr")||"").replace(/,/g,"").trim();
        const suffix = getDrCrSuffix(raw);
        const flagRaw = m.crdrFlag ? String(get("crdrFlag")||"").trim().toUpperCase() : "";
        const amt = parseAmt(raw);
        // Determine direction: explicit flag > suffix > sign
        let isDr;
        if (flagRaw)         isDr = /^(dr|debit|d)/.test(flagRaw);
        else if (suffix)     isDr = suffix === "DR";
        else if (raw.startsWith("-")) isDr = true;
        else                 isDr = false; // positive = credit by convention
        if (isDr) debit = amt || ""; else credit = amt || "";
      } else {
        // Separate debit / credit columns
        const rawDr = String(get("debit")||"").replace(/,/g,"").trim();
        const rawCr = String(get("credit")||"").replace(/,/g,"").trim();
        // Some banks put amount in debit col with DR suffix, credit col with CR suffix
        const drAmt = parseAmt(rawDr);
        const crAmt = parseAmt(rawCr);
        // Treat DR-suffixed value in credit column as debit (and vice versa)
        const drSuffix = getDrCrSuffix(rawDr);
        const crSuffix = getDrCrSuffix(rawCr);
        if (drSuffix === "CR" && drAmt) { credit = drAmt; }
        else if (drAmt) debit = drAmt;
        if (crSuffix === "DR" && crAmt) { debit = crAmt; }
        else if (crAmt) credit = crAmt;
        // If both resolved to same direction, keep the larger one in the right bucket
        if (debit && credit && debit === credit) { credit = ""; }
      }

      const narr = String(get("narration")||"").trim();
      const dateVal = String(get("date")||"").trim();
      const ai = autoDetectLedger ? aiLedger(narr) : defaultLedger;
      return {
        id: genId(),
        date: dateVal,
        narration: narr,
        debit: debit || "",
        credit: credit || "",
        balance: String(get("balance")||"").replace(/,/g,"").trim(),
        ref: String(get("ref")||"").trim(),
        aiLedger: ai, ledger: ai, isDuplicate: false,
      };
    }).filter(r => r.date || r.debit || r.credit);
    const withDups = detectDuplicates(built);
    const rules = loadRules();
    setRows(checkRunningBalance(applyRulesToRows(withDups, rules)));
    setScreen(SCREENS.LEDGER);
    const dups = withDups.filter(r=>r.isDuplicate).length;
    const ruleHits = applyRulesToRows(withDups, rules).filter(r=>r.ruleMatched).length;
    const mismatches = built.filter(r=>r.balanceMismatch).length;
    toast(
      `${built.length} transactions · ${ruleHits} auto-mapped by rules${dups>0?` · ${dups} duplicates`:""}${mismatches>0?` · ${mismatches} balance issues`:""}`,
      dups>0||mismatches>0?"warn":"success"
    );
  };

  const onImport = async () => {
    const targetCos = tally.companies.filter(c=>selectedCompanies.includes(c.id));
    const companyNames = targetCos.map(c=>c.name).join(", ") || selectedCompanies.join(", ");
    const validRows = rows.filter(r=>!r.isDuplicate||r.forceImport);

    // History entry (always saved regardless of push result)
    const entry = {
      id:genId(), filename, date:new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}),
      rawDate:new Date().toISOString(), rows:validRows.length, company:companyNames, status:"Imported",
      suspense:validRows.filter(r=>r.ledger==="Suspense Account").length,
      duplicates:rows.filter(r=>r.isDuplicate).length,
      rows_data: rows.slice(0,500),
      rows_truncated: rows.length > 500,
    };
    setHistory(h => {
      const updated = [{ ...entry, savedAt: new Date().toISOString() }, ...h];
      saveHistory(user.id, updated);
      return updated;
    });

    // If no companies selected or Tally offline — just save history and go to dashboard
    if (!targetCos.length || tally.status !== "ok") {
      toast(`${validRows.length} vouchers saved. ${tally.status!=="ok"?"Use Download XML to import manually into Tally.":""}`, "success");
      setScreen(SCREENS.DASHBOARD);
      setHeaders([]); setRawRows([]); setRows([]); setFilename(""); setAuditLog([]);
      return;
    }

    // Push to each selected company
    let successCount = 0; let failMsg = "";
    for (const co of targetCos) {
      const xml = toTallyXML(validRows, co);
      try {
        const result = await tallyPost(tally.host || "localhost", tally.port || "9000", xml, 30000);
        // Tally returns LINEERROR or IMPORTERROR on failure
        if (result && (result.includes("LINEERROR") || result.includes("IMPORTERROR"))) {
          const errMatch = result.match(/<LINEERROR>([\s\S]*?)<\/LINEERROR>/i) ||
                           result.match(/<IMPORTERROR>([\s\S]*?)<\/IMPORTERROR>/i);
          failMsg = errMatch?.[1]?.trim() || "Tally import error";
          toast(`Tally rejected import for ${co.name}: ${failMsg}`, "warn");
        } else {
          successCount++;
        }
      } catch(e) {
        const isTDLError = e.message.includes("Form:Company") || e.message.includes("No 'PARTS'") || e.message.includes("TDL");
        if (isTDLError) {
          toast(
            "Tally Extension Error: The extension sent a broken TDL request. " +
            "Use 'Download XML' instead and import manually in Tally → Gateway of Tally → Import Data → Vouchers.",
            "warn"
          );
        } else {
          toast(`Push to ${co.name} failed: ${e.message}`, "warn");
        }
        break;
      }
    }

    if (successCount > 0) {
      toast(`✓ ${validRows.length} vouchers pushed to Tally for ${targetCos.slice(0,successCount).map(c=>c.name).join(", ")}`, "success");
    }

    setScreen(SCREENS.DASHBOARD);
    setHeaders([]); setRawRows([]); setRows([]); setFilename(""); setAuditLog([]);
  };

  const onReimport = (h) => {
    if (h.rows_data?.length) {
      setRows(h.rows_data); setFilename(h.filename);
      setScreen(SCREENS.LEDGER);
      toast(`Re-importing ${h.filename}`,"success");
    } else {
      toast("Row data unavailable — please upload the original file again","warn");
      setScreen(SCREENS.UPLOAD);
    }
  };

  if (!user) return (
    <>
      <style>{css}</style>
      <LoginScreen onLogin={onLogin} />
    </>
  );

  const isAdmin = user?.role === "admin";
  const NAV = [
    { id:SCREENS.DASHBOARD, label:"Dashboard",  abbr:"DB" },
    { id:SCREENS.UPLOAD,    label:"New Import",  abbr:"NI" },
    { id:SCREENS.HISTORY,   label:"History",     abbr:"HX" },
    { id:SCREENS.SETTINGS,  label:"Settings",    abbr:"ST" },
    ...(user?.role === "admin" ? [
      { id:SCREENS.USER_MGMT, label:"Users",     abbr:"US", badge: pendingCount > 0 ? pendingCount : null },
    ] : []),
  ];

  // Refresh pending count when admin approval modal closes


  const IMPORT_STEPS = ["Upload","Map Columns","Assign Ledgers","Preview & Export"];
  const isImportScreen = [SCREENS.UPLOAD, SCREENS.COLUMN_MAP, SCREENS.LEDGER, SCREENS.PREVIEW].includes(screen);

  return (
    <>
      <style>{css}</style>
      <Toast toasts={toasts} />

      <div style={{ display:"flex", minHeight:"100vh", background:T.bg, fontFamily:T.font }}>
        {/* Sidebar */}
        <div style={{ width:220, background:T.surface, borderRight:`1px solid ${T.border}`, padding:"20px 0", display:"flex", flexDirection:"column", flexShrink:0, position:"fixed", top:0, bottom:0, left:0, zIndex:100 }}>
          <div style={{ padding:"18px 16px 16px", borderBottom:`1px solid ${T.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:11, marginBottom:12 }}>
              <div style={{ width:40, height:40, borderRadius:12, background:"linear-gradient(145deg,#1d4ed8,#2563eb,#7c3aed)", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, boxShadow:"0 4px 20px rgba(37,99,235,0.30)" }}>
                  <svg width="22" height="22" viewBox="0 0 38 38" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <rect x="4" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                    <rect x="10" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                    <rect x="16" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                    <rect x="22" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                    <rect x="28" y="17" width="4" height="13" rx="1" fill="white" fillOpacity="0.95"/>
                    <rect x="2" y="31" width="33" height="3" rx="1.5" fill="white"/>
                    <polygon points="19,4 2,15 36,15" fill="white" fillOpacity="0.95"/>
                  </svg>
                </div>
              <div>
                <div style={{ fontWeight:900, fontSize:16, letterSpacing:"-0.5px", background:"linear-gradient(135deg,#1d4ed8 30%,#7c3aed)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent", lineHeight:1.2 }}>Bank2Tally</div>
                <div style={{ fontSize:9.5, color:"#64748b", letterSpacing:"0.1em", textTransform:"uppercase", fontWeight:700, marginTop:1 }}>By Verma Consultancy</div>
              </div>
            </div>
            <div style={{ background:"rgba(37,99,235,0.06)", borderLeft:`2px solid rgba(37,99,235,0.4)`, borderRadius:"0 6px 6px 0", padding:"7px 10px" }}>
              <p style={{ fontSize:10.5, color:"#64748b", lineHeight:1.55, fontStyle:"italic", margin:0 }}>"{todayQuote.text.slice(0,72)}{todayQuote.text.length>72?"…":""}"</p>
            </div>
          </div>
          <nav style={{ flex:1, padding:"14px 10px" }}>
            {NAV.map(n=>(
              <button key={n.id} onClick={()=>setScreen(n.id)}
                style={{ width:"100%", display:"flex", alignItems:"center", gap:0, padding:"0", borderRadius:8, fontSize:13, fontFamily:T.font, cursor:"pointer", border:"none", marginBottom:3, transition:"all 0.15s", background:"transparent", overflow:"hidden" }}>
                <div style={{
                  width:3, alignSelf:"stretch", borderRadius:2, flexShrink:0, marginRight:10,
                  background: screen===n.id ? T.accent : "transparent",
                  transition:"background 0.15s"
                }} />
                <div style={{
                  flex:1, display:"flex", alignItems:"center", justifyContent:"space-between",
                  padding:"9px 10px", borderRadius:7,
                  background: screen===n.id ? T.accentDim : n.badge ? T.amberDim : "transparent",
                  color: screen===n.id ? T.accent : n.badge ? T.amber : T.textSub,
                  fontWeight: screen===n.id ? 700 : 500,
                }}>
                  <span>{n.label}</span>
                  {n.badge && (
                    <span style={{ background:T.amber, color:"#fff", borderRadius:99, fontSize:10, fontWeight:700, padding:"1px 7px", minWidth:18, textAlign:"center" }}>
                      {n.badge}
                    </span>
                  )}
                </div>
              </button>
            ))}
          </nav>
          <div style={{ padding:"14px 14px", borderTop:`1px solid ${T.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10, padding:"8px 10px", borderRadius:10, background:T.surface }}>
              <div style={{ width:34, height:34, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:800, color:"#fff", flexShrink:0, boxShadow:`0 2px 10px ${T.accentGlow}` }}>{user.avatar}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:12, fontWeight:700, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{user.name.split(" ")[0]}</div>
                <div style={{ fontSize:10, color:T.textSub, fontWeight:600 }}>{fromDbRole(user.role)}</div>
                {user.email && <div style={{ fontSize:9.5, color:T.textDim, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:130 }}>{user.email}</div>}
              </div>
            </div>
            <button onClick={onLogout}
              style={{ width:"100%", display:"flex", alignItems:"center", justifyContent:"center", gap:8, padding:"8px 12px", borderRadius:8, fontSize:12, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.red}44`, background:T.redDim, color:T.red, transition:"all 0.15s" }}>
              Sign Out
            </button>
          </div>
        </div>

        {/* Main */}
        <div style={{ flex:1, marginLeft:220, padding:"28px 32px", maxWidth:1100, minWidth:0 }}>
          {isImportScreen && <Steps steps={IMPORT_STEPS} current={screen} />}

          {screen === SCREENS.DASHBOARD && <DashboardScreen history={history} setScreen={setScreen} user={user} tally={tally} />}
          {screen === SCREENS.UPLOAD && <UploadScreen onParsed={onParsed} selectedCompanies={selectedCompanies} setSelectedCompanies={setSelectedCompanies} tally={tally} />}
          {screen === SCREENS.COLUMN_MAP && <ColumnMapScreen headers={headers} templateKey={templateKey} onMapped={onMapped} onBack={()=>setScreen(SCREENS.UPLOAD)} />}
          {screen === SCREENS.LEDGER && <LedgerScreen rows={rows} setRows={setRows} onNext={()=>setScreen(SCREENS.PREVIEW)} onBack={()=>setScreen(SCREENS.COLUMN_MAP)} auditLog={auditLog} setAuditLog={setAuditLog} user={user} tally={tally} />}
          {screen === SCREENS.PREVIEW && <PreviewScreen rows={rows} setRows={setRows} filename={filename} selectedCompanies={selectedCompanies} onBack={()=>setScreen(SCREENS.LEDGER)} onImport={onImport} auditLog={auditLog} tally={tally} />}
          {screen === SCREENS.HISTORY && <HistoryScreen history={history} onReimport={onReimport} onDeleteEntry={deleteHistoryEntry} onClearAll={clearAllHistory} onBack={()=>setScreen(SCREENS.DASHBOARD)} />}
          {screen === SCREENS.SETTINGS && <SettingsScreen user={user} onLogout={onLogout} onUserUpdate={u=>setUser(u)} tally={tally} tallyHost={tallyHost} setTallyHost={setTallyHost} tallyPort={tallyPort} setTallyPort={setTallyPort} defaultLedger={defaultLedger} setDefaultLedger={setDefaultLedger} autoDetectLedger={autoDetectLedger} setAutoDetectLedger={setAutoDetectLedger} />}
          {screen === SCREENS.USER_MGMT && isAdmin && <UserManagementScreen adminUser={user} />}
        </div>
      </div>
    </>
  );
}

// Wrap with ErrorBoundary so any uncaught render error shows a recovery screen
export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}
