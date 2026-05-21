import { useState, useRef, useCallback, useEffect, useMemo } from "react";
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
    sb._token = data.access_token;
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

  // DB: generic select
  async from(table, query = "") {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}${query ? "?" + query : ""}`, {
      headers: { ...sb._headers(), "Prefer": "return=representation" },
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.message || JSON.stringify(data));
    return data;
  },

  // DB: update
  async update(table, match, payload) {
    const q = Object.entries(match).map(([k,v]) => `${k}=eq.${encodeURIComponent(v)}`).join("&");
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${q}`, {
      method: "PATCH",
      headers: { ...sb._headers(), "Prefer": "return=representation" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.message || JSON.stringify(data));
    return data;
  },

  // DB: insert
  async insert(table, payload) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
      method: "POST",
      headers: { ...sb._headers(), "Prefer": "return=representation" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.message || JSON.stringify(data));
    return data;
  },
};

// ── Design Tokens ────────────────────────────────────────────────
const T = {
  bg:          "#080b12",
  surface:     "#0e1220",
  card:        "#131825",
  border:      "#1e2640",
  borderLight: "#263050",
  accent:      "#3d7fff",
  accentDim:   "#0d1f4a",
  accentGlow:  "rgba(61,127,255,0.18)",
  accentSoft:  "rgba(61,127,255,0.08)",
  green:       "#10d98c",
  greenDim:    "#052e1e",
  amber:       "#ffb547",
  amberDim:    "#3d2200",
  red:         "#ff4f6a",
  redDim:      "#3d0a14",
  purple:      "#b47cff",
  purpleDim:   "#1e0a40",
  gold:        "#ffd166",
  goldDim:     "#3d2800",
  text:        "#eef2ff",
  textMid:     "#8896b3",
  textDim:     "#3d4f6e",
  font:        "'DM Sans', 'Segoe UI', sans-serif",
  mono:        "'JetBrains Mono', 'Fira Code', monospace",
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

// Build Tally XML request envelope
function tallyRequest(body) {
  return `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME>${body}</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;
}

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

// Send message to extension and wait for response
function sendToExtension(msg, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const requestId = Math.random().toString(36).slice(2);
    const responseType = msg.type + "_RESPONSE";
    const timer = setTimeout(() => {
      window.removeEventListener("message", handler);
      reject(new Error("Extension response timed out"));
    }, timeoutMs);

    function handler(e) {
      if (e.data?.type === responseType && e.data?.requestId === requestId) {
        clearTimeout(timer);
        window.removeEventListener("message", handler);
        resolve(e.data);
      }
    }
    window.addEventListener("message", handler);
    window.postMessage({ ...msg, requestId }, "*");
  });
}

// Generic POST to Tally via the Chrome extension (required — app runs on HTTPS,
// so direct HTTP to localhost is blocked by mixed-content policy)
async function tallyPost(host, port, xmlBody, timeoutMs = 10000) {
  if (!_extensionReady) {
    throw new Error("EXTENSION_NOT_READY");
  }
  const res = await sendToExtension({
    type: "TALLY_REQUEST",
    host, port, body: xmlBody,
  }, timeoutMs);
  if (!res.success) throw new Error(res.error || "Tally request failed");
  return res.data;
}

// Parse Tally XML response — extract text content of all matching tags
function parseTallyTags(xml, tag) {
  const re = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "gi");
  const results = [];
  let m;
  while ((m = re.exec(xml)) !== null) results.push(m[1].trim());
  return results;
}

// Fetch all companies from Tally
async function fetchTallyCompanies(host, port) {
  const xml = `<ENVELOPE>
    <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
    <BODY><EXPORTDATA><REQUESTDESC>
      <REPORTNAME>List of Companies</REPORTNAME>
      <STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>
    </REQUESTDESC></EXPORTDATA></BODY>
  </ENVELOPE>`;
  const raw = await tallyPost(host, port, xml);
  const names = parseTallyTags(raw, "COMPANY");
  // Also try BASICCOMPANYNAME tag (Tally Prime)
  const prime = parseTallyTags(raw, "BASICCOMPANYNAME");
  const all = [...new Set([...names, ...prime])].filter(Boolean);
  if (!all.length) throw new Error("Tally responded but returned no companies. Make sure at least one company is loaded in Tally.");
  return all.map((name, i) => ({ id: `tc${i}`, name, gstin: "", state: "", fy: "2024-25" }));
}

// Fetch company details (GSTIN, state, FY) for a specific company
async function fetchTallyCompanyDetails(host, port, companyName) {
  const xml = `<ENVELOPE>
    <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
    <BODY><EXPORTDATA><REQUESTDESC>
      <REPORTNAME>Company Info</REPORTNAME>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>${companyName}</SVCURRENTCOMPANY>
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
        <SVCURRENTCOMPANY>${companyName}</SVCURRENTCOMPANY>
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

// Test connection via extension
async function testTallyConnection(host, port) {
  if (!_extensionReady) {
    throw new Error("Bank2Tally Connector extension not detected. Please install it from the instructions below, then refresh this page.");
  }
  const res = await sendToExtension({ type: "TALLY_PING", host, port }, 8000);
  if (res.success) return true;
  throw new Error(res.error || "Extension is installed but cannot reach Tally. Make sure Tally is open on this computer.");
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
      // Fetch ledgers for each company (fire and forget per company)
      cos.forEach(async (co) => {
        try {
          const ledgers = await fetchTallyLedgers(h, p, co.name);
          if (ledgers) setLedgerMap(m => ({ ...m, [co.name]: ledgers }));
        } catch { /* ignore per-company ledger errors */ }
      });
      // Fetch company details (GSTIN, state, FY) in background
      cos.forEach(async (co, i) => {
        try {
          const details = await fetchTallyCompanyDetails(h, p, co.name);
          setCompanies(cs => cs.map((c, j) => j === i ? { ...c, ...details } : c));
        } catch { /* ignore */ }
      });
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

  return { status, companies, ledgerMap, error, lastFetch, refetch };
}

const BANK_TEMPLATES = {
  hdfc:  { name: "HDFC Bank", cols: { date: "Date", narration: "Narration", debit: "Withdrawal Amt.", credit: "Deposit Amt.", balance: "Closing Balance", ref: "Chq./Ref.No." }},
  sbi:   { name: "State Bank of India", cols: { date: "Txn Date", narration: "Description", debit: "Debit", credit: "Credit", balance: "Balance", ref: "Ref No./Cheque No." }},
  icici: { name: "ICICI Bank", cols: { date: "Transaction Date", narration: "Transaction Remarks", debit: "Withdrawal Amount (INR )", credit: "Deposit Amount (INR )", balance: "Balance (INR )", ref: "S No." }},
  axis:  { name: "Axis Bank", cols: { date: "Tran Date", narration: "PARTICULARS", debit: "DR", credit: "CR", balance: "BAL", ref: "CHQNO" }},
  kotak: { name: "Kotak Mahindra Bank", cols: { date: "Transaction Date", narration: "Description", debit: "Debit Amount", credit: "Credit Amount", balance: "Balance", ref: "Reference No" }},
  pnb:   { name: "Punjab National Bank", cols: { date: "Date", narration: "Particulars", debit: "Debit", credit: "Credit", balance: "Balance", ref: "Ref. No." }},
  yes:   { name: "Yes Bank", cols: { date: "Date", narration: "Transaction Details", debit: "Debit", credit: "Credit", balance: "Balance", ref: "Reference Number" }},
  idfc:  { name: "IDFC First Bank", cols: { date: "Date", narration: "Transaction Remarks", debit: "Debit Amount", credit: "Credit Amount", balance: "Balance", ref: "Transaction ID" }},
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

const USERS = [
  { id: "u1", name: "Rajesh Kumar", email: "admin@acmecorp.in", role: "Admin", avatar: "RK", company: "Acme Corp Pvt Ltd" },
  { id: "u2", name: "Priya Sharma", email: "ca@acmecorp.in", role: "CA", avatar: "PS", company: "Acme Corp Pvt Ltd" },
  { id: "u3", name: "Amit Verma", email: "accountant@acmecorp.in", role: "Accountant", avatar: "AV", company: "Acme Corp Pvt Ltd" },
];

// ── Helpers ──────────────────────────────────────────────────────
const genId = () => Math.random().toString(36).slice(2, 9);
const fmt = n => n == null || n === "" ? "" : Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = v => {
  if (!v) return "";
  let d = v instanceof Date ? v : new Date(String(v));
  if (isNaN(d)) {
    const p = String(v).split(/[\/\-\.]/);
    if (p.length === 3) {
      const [a, b, c] = p;
      const y = c?.length === 4 ? c : a?.length === 4 ? a : "20" + c;
      d = new Date(`${y}-${String(b).padStart(2,"0")}-${String(a?.length===4?b:a).padStart(2,"0")}`);
    }
  }
  return isNaN(d) ? String(v) : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
};
const fmtDateShort = v => {
  if (!v) return "";
  let d = v instanceof Date ? v : new Date(String(v));
  return isNaN(d) ? String(v) : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
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

const voucherType = (debit, credit, ledger) => {
  if (ledger?.toLowerCase().includes("bank") || ledger?.toLowerCase().includes("cash")) return "Contra";
  if (credit && !debit) return "Receipt";
  if (debit && !credit) return "Payment";
  return "Journal";
};

const detectDuplicates = rows => {
  const seen = new Map();
  return rows.map(r => {
    const key = `${String(r.date).slice(0,10)}|${r.debit}|${r.credit}|${String(r.narration).slice(0,25).toLowerCase()}`;
    if (seen.has(key)) return { ...r, isDuplicate: true, duplicateOf: seen.get(key) };
    seen.set(key, r.id);
    return { ...r, isDuplicate: false };
  });
};

// Tally XML generator
const toTallyXML = (rows, company, fy = "2024-25") => {
  const [fyStart, fyEnd] = fy.split("-");
  const vouchers = rows.filter(r => !r.isDuplicate || r.forceImport).map(r => {
    const amt = parseFloat(r.debit || r.credit || 0);
    const isDebit = !!r.debit;
    const vtype = r.voucherType || voucherType(r.debit, r.credit, r.ledger);
    return `
  <VOUCHER REMOTEID="${r.id}" VCHTYPE="${vtype}" ACTION="Create">
    <DATE>${String(r.date).replace(/[^0-9]/g, "").slice(0,8)}</DATE>
    <NARRATION>${(r.narration || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</NARRATION>
    <VOUCHERTYPENAME>${vtype}</VOUCHERTYPENAME>
    <PARTYLEDGERNAME>${r.ledger}</PARTYLEDGERNAME>
    <ALLLEDGERENTRIES.LIST>
      <LEDGERNAME>${isDebit ? r.ledger : company.bankLedger || "HDFC Bank"}</LEDGERNAME>
      <ISDEEMEDPOSITIVE>${isDebit ? "No" : "Yes"}</ISDEEMEDPOSITIVE>
      <AMOUNT>${isDebit ? amt : -amt}</AMOUNT>
    </ALLLEDGERENTRIES.LIST>
    <ALLLEDGERENTRIES.LIST>
      <LEDGERNAME>${isDebit ? company.bankLedger || "HDFC Bank" : r.ledger}</LEDGERNAME>
      <ISDEEMEDPOSITIVE>${isDebit ? "Yes" : "No"}</ISDEEMEDPOSITIVE>
      <AMOUNT>${isDebit ? -amt : amt}</AMOUNT>
    </ALLLEDGERENTRIES.LIST>
  </VOUCHER>`;
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
          <SVCURRENTCOMPANY>${company.name}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>${vouchers}
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

// Best-effort extraction of text from a PDF page using pdfjs
async function extractPdfText(buf) {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
  const pdfjsLib = window["pdfjs-dist/build/pdf"];
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  let fullText = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    // Preserve columnar structure: group items by approximate Y, sort by X
    const byY = {};
    content.items.forEach(item => {
      const y = Math.round(item.transform[5] / 6) * 6;
      if (!byY[y]) byY[y] = [];
      byY[y].push({ x: item.transform[4], str: item.str });
    });
    const lines = Object.keys(byY)
      .sort((a, b) => Number(b) - Number(a))
      .map(y => byY[y].sort((a, b) => a.x - b.x).map(it => it.str.trim()).filter(Boolean).join("\t"));
    fullText += lines.join("\n") + "\n";
  }
  return fullText.trim();
}

// OCR fallback: rasterise pages and run Tesseract
async function ocrPdfText(buf, onProgress) {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
  await loadScript("https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js");
  const pdfjsLib = window["pdfjs-dist/build/pdf"];
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
  const worker = await window.Tesseract.createWorker("eng");
  let fullText = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    onProgress && onProgress(`OCR page ${i}/${pdf.numPages}…`);
    const page = await pdf.getPage(i);
    const vp = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = vp.width; canvas.height = vp.height;
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    const dataUrl = canvas.toDataURL("image/png");
    const { data: { text } } = await worker.recognize(dataUrl);
    fullText += text + "\n";
  }
  await worker.terminate();
  return fullText.trim();
}

// Parse flat tabular text (TSV/space-aligned) into {headers, rows}
function parsePdfText(text) {
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);

  // Heuristic: find a header line containing bank-statement keywords
  const HDR_RE = /date|narr|desc|debit|credit|withdraw|deposit|balance|particulars/i;
  let headerIdx = -1;
  for (let i = 0; i < Math.min(lines.length, 30); i++) {
    if (HDR_RE.test(lines[i]) && lines[i].split(/\t|  {2,}/).length >= 3) { headerIdx = i; break; }
  }

  // If no header found, try to infer columns from first data-heavy line
  const delim = (line) => {
    const tabCount = (line.match(/\t/g) || []).length;
    const spaceCount = (line.match(/  {2,}/g) || []).length;
    return tabCount >= 2 ? "\t" : spaceCount >= 2 ? /  {2,}/ : null;
  };

  if (headerIdx === -1) {
    // Try to auto-detect by finding consecutive numeric-heavy lines
    for (let i = 0; i < Math.min(lines.length, 40); i++) {
      const cols = lines[i].split(/\t|  {2,}/).filter(Boolean);
      if (cols.length >= 4 && /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(lines[i])) {
        headerIdx = Math.max(0, i - 1); break;
      }
    }
  }

  if (headerIdx === -1) throw new Error("Could not detect column structure in PDF. Please try CSV or Excel export from your bank's portal.");

  const sep = delim(lines[headerIdx]) || "\t";
  const headers = lines[headerIdx].split(sep).map(h => h.trim()).filter(Boolean);
  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const cols = lines[i].split(sep).map(c => c.trim());
    if (cols.length < 2) continue;
    // Skip totals/footer lines
    if (/^(total|closing|opening|grand|page)/i.test(lines[i])) continue;
    // Pad or trim to header length
    while (cols.length < headers.length) cols.push("");
    rows.push(cols.slice(0, headers.length));
  }
  if (!rows.length) throw new Error("PDF parsed but no data rows found. Try exporting as Excel/CSV.");
  return { headers, rows };
}

// ── File Parser ──────────────────────────────────────────────────
async function parseFile(file, onProgress) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "pdf") {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf.slice(0, 2048));
    const header = new TextDecoder("latin1").decode(bytes);
    if (/\/Encrypt/i.test(header)) {
      const e = new Error("This PDF is password-protected. Bank2Tally cannot read encrypted PDFs.");
      e.code = "ERR_002"; throw e;
    }
    onProgress && onProgress("Loading PDF engine…");
    let text = "";
    try {
      text = await extractPdfText(buf);
    } catch (pdfErr) {
      // pdfjs failed — possibly a scanned PDF; fall through to OCR
      text = "";
    }
    // If text extraction yielded too little content (scanned PDF), run OCR
    const wordCount = text.replace(/\s+/g, " ").split(" ").filter(Boolean).length;
    if (wordCount < 20) {
      onProgress && onProgress("Scanned PDF detected — starting OCR…");
      text = await ocrPdfText(buf, onProgress);
    }
    onProgress && onProgress("Parsing table structure…");
    return parsePdfText(text);
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetNames = wb.SheetNames;
  const ws = wb.Sheets[sheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  // Find header row (skip metadata rows)
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const nonEmpty = raw[i].filter(c => String(c).trim()).length;
    if (nonEmpty >= 3) { headerRowIdx = i; break; }
  }
  const headers = raw[headerRowIdx].map(h => String(h).trim()).filter(Boolean);
  const rows = raw.slice(headerRowIdx + 1).filter(r => r.some(c => c !== "" && c !== null && c !== undefined));
  if (!headers.length) throw new Error("Could not detect column headers in this file.");
  return { headers, rows, sheetNames };
}

// ── Styled Primitives ────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=DM+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${T.bg}; color: ${T.text}; font-family: ${T.font}; }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: ${T.surface}; }
  ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 99px; }
  select, input, textarea { background: ${T.card}; color: ${T.text}; border: 1px solid ${T.border}; border-radius: 8px; font-family: ${T.font}; font-size: 13px; }
  select:focus, input:focus { outline: none; border-color: ${T.accent}; box-shadow: 0 0 0 3px ${T.accentGlow}; }
  @keyframes fadeIn { from { opacity:0; transform:translateY(8px) } to { opacity:1; transform:translateY(0) } }
  @keyframes shimmer { from { background-position: -200% 0 } to { background-position: 200% 0 } }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.5} }
  @keyframes slideIn { from{transform:translateX(100%);opacity:0} to{transform:translateX(0);opacity:1} }
  .fade-in { animation: fadeIn 0.3s ease forwards; }
  .row-hover:hover { background: ${T.borderLight} !important; }
  .btn-hover:hover { filter: brightness(1.15); transform: translateY(-1px); }
  .card-hover:hover { border-color: ${T.accent} !important; }
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
  return <div className={className} style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:20, ...extra }}>{children}</div>;
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
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.75)", zIndex:9000, display:"flex", alignItems:"center", justifyContent:"center", backdropFilter:"blur(4px)" }} onClick={onClose}>
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
    <Card style={{ padding:"16px 20px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:-10, right:-10, fontSize:48, opacity:0.06 }}>{icon}</div>
      <div style={{ fontSize:22, marginBottom:6 }}>{icon}</div>
      <div style={{ fontSize:22, fontWeight:700, color, letterSpacing:"-0.5px" }}>{value}</div>
      <div style={{ fontSize:12, color:T.textDim, marginTop:2 }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:T.textMid, marginTop:4 }}>{sub}</div>}
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
  const handleLogin = async () => {
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
          onLogin({ id: userId, name: meta.name || userEmail.split("@")[0], role: "admin", status: "approved", avatar: upper2(meta.name || userEmail), email: userEmail, sessionToken: session.access_token });
          return;
        }
        throw new Error("Your account is pending admin approval.");
      }
      if (profile.status === "pending")  { setPendingUser({ ...profile, email: userEmail }); setLoading(false); return; }
      if (profile.status === "rejected") throw new Error("Your access request was rejected. Contact admin.");
      // Merge email from session since profiles table has no email column
      onLogin({ ...profile, email: userEmail, sessionToken: session.access_token });
    } catch (e) { setErr(e.message); }
    setLoading(false);
  };

  const upper2 = str => (str || "").slice(0,2).toUpperCase();

  // ── Email verification state ─────────────────────────────────────
  const [verifyStep,  setVerifyStep]  = useState(false); // show "check email" screen
  const [pendingReg,  setPendingReg]  = useState(null);  // holds reg data

  // ── Register ─────────────────────────────────────────────────────
  const handleRegister = async () => {
    setErr(""); setSuccess("");
    if (!regName.trim())         return setErr("Full name is required.");
    if (!regEmail.includes("@")) return setErr("Enter a valid email address.");
    if (regPass.length < 8)      return setErr("Password must be at least 8 characters.");
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
            id: userId,
            name: regName.trim(),
            role: "user",
            company: regCompany.trim(),
            status: "pending",
            avatar: regName.trim().slice(0,2).toUpperCase(),
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
            <div style={{ width:76, height:76, borderRadius:22, background:"linear-gradient(145deg, #1a4fd6, #3d7fff, #7c3aed)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:36, boxShadow:"0 0 0 1px rgba(61,127,255,0.3), 0 0 60px rgba(61,127,255,0.25), 0 8px 32px rgba(0,0,0,0.5)" }}>🏦</div>
          </div>
          <h1 style={{ fontSize:28, fontWeight:900, letterSpacing:"-1px", marginBottom:3, background:"linear-gradient(135deg, #eef2ff 40%, #3d7fff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>Bank2Tally</h1>
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
                  <Btn onClick={handleForgotPassword} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"📨"}>
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
                <div style={{ fontSize:52, marginBottom:12 }}>📬</div>
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
                <Btn onClick={handleResendVerification} variant="secondary" fullWidth icon="📨">
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
                <Input value={pass} onChange={e=>setPass(e.target.value)} placeholder="••••••••" prefix="🔒" />
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
                <Btn onClick={handleRegister} disabled={loading} fullWidth size="lg" icon={loading?"⏳":"📝"}>
                  {loading ? "Submitting…" : "Request Access"}
                </Btn>
              </div>
            </>
          )}
        </Card>

        {/* ── Verma Consultancy Branding Footer ── */}
        <div style={{ textAlign:"center", marginTop:22, paddingTop:16, borderTop:`1px solid ${T.border}` }}>
          <div style={{ display:"inline-flex", alignItems:"center", gap:6, marginBottom:8 }}>
            <span style={{ color:T.gold, fontSize:14 }}>✦</span>
            <span style={{ fontSize:12, fontWeight:700, color:T.textMid, letterSpacing:"0.03em" }}>Produced by</span>
            <span style={{ fontSize:12, fontWeight:800, color:T.text, letterSpacing:"0.02em" }}>Verma Consultancy Services</span>
            <span style={{ color:T.gold, fontSize:14 }}>✦</span>
          </div>
          <p style={{ fontSize:11, color:T.textDim, marginBottom:8 }}>For purchase, support &amp; enquiries</p>
          <div style={{ display:"flex", justifyContent:"center", gap:16, flexWrap:"wrap" }}>
            <a href="tel:+918707401846"
              style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", background:T.accentSoft, border:`1px solid ${T.accent}44`, borderRadius:20, fontSize:11, fontWeight:600, color:T.accent, textDecoration:"none" }}>
              📞 8707401846
            </a>
            <a href="mailto:svtiger543939@gmail.com"
              style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 12px", background:T.accentSoft, border:`1px solid ${T.accent}44`, borderRadius:20, fontSize:11, fontWeight:600, color:T.accent, textDecoration:"none" }}>
              ✉ svtiger543939@gmail.com
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
            <h2 style={{ fontSize:22, fontWeight:900, letterSpacing:"-0.6px", background:"linear-gradient(135deg, #eef2ff 50%, #3d7fff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>Dashboard</h2>
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
      <div style={{ background:"linear-gradient(135deg, rgba(61,127,255,0.07), rgba(180,124,255,0.07))", border:"1px solid rgba(61,127,255,0.18)", borderRadius:14, padding:"12px 18px", marginBottom:18, display:"flex", alignItems:"center", gap:14 }}>
        <div style={{ fontSize:26, flexShrink:0 }}>💡</div>
        <div>
          <p style={{ fontSize:12, color:T.textMid, lineHeight:1.6, fontStyle:"italic", marginBottom:2 }}>"{todayQuote.text}"</p>
          <p style={{ fontSize:10, color:T.textDim, fontWeight:600 }}>— {todayQuote.author}</p>
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
            <div style={{ fontSize:36, marginBottom:10 }}>📂</div>
            <p style={{ fontSize:14 }}>No imports yet. Start your first import →</p>
          </div>
        ) : history.slice(0,5).map(h => (
          <div key={h.id} className="row-hover" style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"11px 14px", borderRadius:9, marginBottom:4, cursor:"pointer", transition:"background 0.15s" }}>
            <div style={{ display:"flex", alignItems:"center", gap:12 }}>
              <div style={{ width:36, height:36, borderRadius:9, background:T.accentDim, display:"flex", alignItems:"center", justifyContent:"center", fontSize:16 }}>📄</div>
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
            ⚠ Could not reach Tally at {tally.status === "error" ? "localhost:9000" : ""}. Go to <strong>Settings → Test Connection</strong> to reconnect, or type a company name below.
          </div>
        )}

        {tally.status === "connecting" && (
          <div style={{ display:"flex", gap:10, alignItems:"center", padding:"12px 0", color:T.textDim, fontSize:13 }}>
            <span style={{ animation:"pulse 1.5s infinite", fontSize:18 }}>⏳</span> Fetching companies from Tally gateway…
          </div>
        )}

        {displayCompanies.length > 0 && (
          <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
            {displayCompanies.map(c => {
              const sel = selectedCompanies.includes(c.id);
              return (
                <button key={c.id} onClick={() => setSelectedCompanies(p => sel ? p.filter(x=>x!==c.id) : [...p,c.id])}
                  style={{ padding:"8px 14px", borderRadius:9, fontSize:12, fontWeight:500, cursor:"pointer", transition:"all 0.15s", fontFamily:T.font,
                    border: sel ? `2px solid ${T.accent}` : `1px solid ${T.border}`,
                    background: sel ? T.accentDim : T.surface,
                    color: sel ? T.accent : T.textMid,
                    boxShadow: sel ? `0 0 12px ${T.accentGlow}` : "none" }}>
                  {sel && "✓ "}{c.name}
                  {c.state && <span style={{ marginLeft:6, fontSize:10, opacity:0.6 }}>· {c.state}</span>}
                </button>
              );
            })}
          </div>
        )}

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
            <div style={{ fontSize:40, marginBottom:14 }}>🏦</div>
            <p style={{ fontWeight:700, fontSize:16, color:T.text, marginBottom:6 }}>Drop bank statement here</p>
            <p style={{ color:T.textDim, fontSize:13 }}>or <span style={{color:T.accent, textDecoration:"underline"}}>browse files</span></p>
            <div style={{ display:"flex", gap:8, justifyContent:"center", marginTop:16, flexWrap:"wrap" }}>
              {["📄 CSV", "📊 .xlsx / .xls", "📑 PDF (text + OCR)"].map(l => (
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
const COL_LABELS = { date:"Date", narration:"Narration / Description", debit:"Debit Amount", credit:"Credit Amount", balance:"Balance (optional)", ref:"Reference / Cheque No. (optional)" };

function ColumnMapScreen({ headers, templateKey, onMapped, onBack }) {
  const [mapping, setMapping] = useState({});
  const [draggingCol, setDraggingCol] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  const autoMap = useCallback(() => {
    if (templateKey && BANK_TEMPLATES[templateKey]) {
      const tpl = BANK_TEMPLATES[templateKey].cols;
      const m = {};
      Object.entries(tpl).forEach(([field, colName]) => {
        const found = headers.find(h => h.toLowerCase().trim() === colName.toLowerCase().trim()) || headers.find(h => h.toLowerCase().includes(colName.toLowerCase().slice(0,6)));
        if (found) m[field] = found;
      });
      setMapping(m); return;
    }
    const m = {};
    headers.forEach(h => {
      const hl = h.toLowerCase().replace(/[\s_\-\.]/g,"");
      if (!m.date && /date|dt/.test(hl)) m.date = h;
      else if (!m.narration && /narr|desc|particular|detail|remark|note|trxn|txn|transaction/.test(hl)) m.narration = h;
      else if (!m.debit && /debit|dr|withdraw|paid|debitamt/.test(hl)) m.debit = h;
      else if (!m.credit && /credit|cr|deposit|received|creditamt/.test(hl)) m.credit = h;
      else if (!m.balance && /balance|bal/.test(hl)) m.balance = h;
      else if (!m.ref && /ref|chq|cheque|utr|neft|imps|refno/.test(hl)) m.ref = h;
    });
    setMapping(m);
  }, [headers, templateKey]);

  useEffect(() => { autoMap(); }, [autoMap]);

  const assignedSet = new Set(Object.values(mapping).filter(Boolean));
  const allRequired = REQUIRED_COLS.every(k => mapping[k]);

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
          ⚠ Required: {REQUIRED_COLS.filter(k=>!mapping[k]).map(k=>COL_LABELS[k]).join(", ")}
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
  const [filter, setFilter] = useState("all"); // all | suspense | duplicate | ready
  const [forceImport, setForceImport] = useState({});
  const [bulkLedger, setBulkLedger] = useState("");
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [showAudit, setShowAudit] = useState(false);

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
          <Btn size="sm" variant="ghost" onClick={()=>setShowAudit(true)} icon="📋">Audit log ({auditLog.length})</Btn>
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
        <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search narration, ledger…" prefix="🔍" style={{ width:220 }} />
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
              <th style={{ padding:"10px 8px", width:32 }}></th>
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
                  <td style={{ padding:"8px 12px", color:T.textMid, whiteSpace:"nowrap", fontFamily:T.mono, fontSize:11 }}>{fmtDateShort(r.date)}</td>
                  <td style={{ padding:"8px 12px", maxWidth:220, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:T.text }} title={r.narration}>{r.narration}</td>
                  <td style={{ padding:"8px 12px", color:T.textDim, fontFamily:T.mono, fontSize:11 }}>{r.ref||"—"}</td>
                  <td style={{ padding:"8px 12px", color:T.red, fontWeight:500, textAlign:"right", fontFamily:T.mono }}>{r.debit ? fmt(r.debit) : ""}</td>
                  <td style={{ padding:"8px 12px", color:T.green, fontWeight:500, textAlign:"right", fontFamily:T.mono }}>{r.credit ? fmt(r.credit) : ""}</td>
                  <td style={{ padding:"8px 12px" }}>
                    <Pill color={vtype==="Receipt"?"green":vtype==="Payment"?"red":vtype==="Contra"?"blue":"purple"} size="xs">{vtype}</Pill>
                  </td>
                  <td style={{ padding:"8px 12px" }}>
                    <Pill color={r.aiLedger==="Suspense Account"?"amber":"blue"} size="xs">{(r.aiLedger||"").slice(0,18)}</Pill>
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
function PreviewScreen({ rows, filename, selectedCompanies, onBack, onImport, auditLog, tally }) {
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

      {/* Summary */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:12, marginBottom:20 }}>
        <StatCard icon="📑" label="Transactions" value={validRows.length} color={T.accent} />
        <StatCard icon="📤" label="Total Debit" value={"₹"+fmt(totalDebit)} color={T.red} />
        <StatCard icon="📥" label="Total Credit" value={"₹"+fmt(totalCredit)} color={T.green} />
        <StatCard icon="⚖️" label="Net Balance" value={(net>=0?"↑":"↓")+"₹"+fmt(Math.abs(net))} color={net>=0?T.green:T.red} />
        <StatCard icon="⚠️" label="Suspense" value={validRows.filter(r=>r.ledger==="Suspense Account").length} color={T.amber} />
      </div>

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
                    <Pill size="xs" color={voucherType(r.debit,r.credit,r.ledger)==="Receipt"?"green":voucherType(r.debit,r.credit,r.ledger)==="Payment"?"red":"blue"}>
                      {voucherType(r.debit,r.credit,r.ledger)}
                    </Pill>
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

      <div style={{ display:"flex", justifyContent:"space-between", marginTop:20 }}>
        <Btn variant="secondary" onClick={onBack} icon="←">Back</Btn>
        <Btn onClick={onImport} icon="🚀" size="lg">Push to Tally ({companies.length} {companies.length===1?"company":"companies"})</Btn>
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

      <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by file name or company…" prefix="🔍" style={{ marginBottom:16, width:"100%" }} />

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
function SettingsScreen({ user, onLogout, tally, tallyHost, setTallyHost, tallyPort, setTallyPort, defaultLedger, setDefaultLedger, autoDetectLedger, setAutoDetectLedger }) {
  const [saved, setSaved] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState(null); // null | "ok" | "error"
  const [testMsg, setTestMsg] = useState("");

  const save = () => { setSaved(true); setTimeout(()=>setSaved(false),2000); };

  const runTest = async () => {
    setTesting(true); setTestResult(null); setTestMsg("");
    try {
      await testTallyConnection(tallyHost, tallyPort);
      // Mark extension/connection as ready globally so all other checks pass
      _markExtensionReady();
      setTestResult("ok"); setTestMsg(`Connected! Fetching companies…`);
      tally.refetch(tallyHost, tallyPort);
      setTimeout(() => setTestMsg(`Connected · ${tally.companies.length} companies loaded`), 2500);
    } catch (e) {
      setTestResult("error"); setTestMsg(e.message);
    } finally { setTesting(false); }
  };

  return (
    <div className="fade-in">
      <h2 style={{ fontSize:20, fontWeight:700, color:T.text, marginBottom:20 }}>Settings</h2>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
        {/* User */}
        <Card>
          <p style={{ fontWeight:600, fontSize:14, marginBottom:16, color:T.text }}>Account</p>
          <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:20 }}>
            <div style={{ width:52, height:52, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, fontWeight:700, color:"#fff" }}>{user?.avatar}</div>
            <div>
              <div style={{ fontWeight:600, fontSize:15, color:T.text }}>{user?.name}</div>
              <div style={{ fontSize:12, color:T.textDim }}>{user?.email}</div>
              <Pill color="blue" size="xs">{user?.role}</Pill>
            </div>
          </div>
          <Btn variant="danger" onClick={onLogout} fullWidth icon="→">Sign Out</Btn>
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
            <p style={{ fontSize:11, marginTop:8, color:testResult==="ok"?T.green:T.red }}>{testMsg}</p>
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
          <p style={{ fontSize:11, color:T.textDim, marginTop:10 }}>Tally Prime → F12 &gt; Advanced Config → Enable Tally Gateway on port {tallyPort}</p>
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

// ══════════════════════════════════════════════════════════════════
// SCREEN: User Management (Admin only)
// ══════════════════════════════════════════════════════════════════
function UserManagementScreen({ adminUser }) {
  const [users, setUsers]           = useState([]);
  const [loading, setLoading]       = useState(true);
  const [search, setSearch]         = useState("");
  const [filterRole, setFilterRole] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [actioning, setActioning]   = useState(null); // userId being acted on
  const [toast_, setToast_]         = useState({ msg:"", type:"success" });
  const [confirmDel, setConfirmDel] = useState(null); // user to confirm-delete
  const [viewUser, setViewUser]     = useState(null); // user detail modal
  const [resetModal, setResetModal] = useState(null); // user for pwd reset
  const [newPass, setNewPass]       = useState("");
  const [newPassErr, setNewPassErr] = useState("");
  const [addModal, setAddModal]     = useState(false);
  const [addForm, setAddForm]       = useState({ name:"", email:"", role:"user", company:"", password:"" });
  const [addErr, setAddErr]         = useState("");
  const [addLoading, setAddLoading] = useState(false);
  const [tab, setTab] = useState("users"); // users | approvals

  const notify = (msg, type="success") => {
    setToast_({ msg, type });
    setTimeout(() => setToast_(t => t.msg === msg ? { msg:"", type:"success" } : t), 3500);
  };

  // ── Load all profiles ──────────────────────────────────────────
  const loadUsers = useCallback(async () => {
    setLoading(true);
    try {
      const profiles = await sb.from("profiles", "select=*&order=created_at.asc");

      // Try to fetch auth user list for emails (requires admin JWT — best effort)
      let authEmailMap = {};
      try {
        const res = await fetch(`${SUPABASE_URL}/auth/v1/admin/users?per_page=1000`, {
          headers: { "apikey": SUPABASE_ANON, "Authorization": `Bearer ${sb._token || SUPABASE_ANON}` }
        });
        if (res.ok) {
          const d = await res.json();
          const list = d.users || d || [];
          list.forEach(u => { if (u.id && u.email) authEmailMap[u.id] = u.email; });
        }
      } catch {}

      setUsers(profiles.map(p => ({
        ...p,
        email: p.email || authEmailMap[p.id] || "",
        avatar: p.avatar || (p.name || "?").slice(0,2).toUpperCase(),
        // Sanitize company: if it's a pure number it's likely a DB artifact
        company: (p.company && isNaN(String(p.company).trim())) ? p.company : (p.company ? "" : ""),
      })));
    } catch (e) {
      notify("Error loading users: " + e.message, "error");
    }
    setLoading(false);
  }, []);

  useEffect(() => { loadUsers(); }, [loadUsers]);

  // ── Hold / Unhold ──────────────────────────────────────────────
  const toggleHold = async (u) => {
    const newStatus = u.status === "on_hold" ? "approved" : "on_hold";
    setActioning(u.id);
    try {
      await sb.update("profiles", { id: u.id }, { status: newStatus });
      setUsers(us => us.map(x => x.id === u.id ? { ...x, status: newStatus } : x));
      notify(`${u.name} ${newStatus === "on_hold" ? "put on hold" : "reactivated"}`);
    } catch (e) { notify("Error: " + e.message, "error"); }
    setActioning(null);
  };

  // ── Change Role ────────────────────────────────────────────────
  const changeRole = async (u, role) => {
    setActioning(u.id);
    try {
      await sb.update("profiles", { id: u.id }, { role });
      setUsers(us => us.map(x => x.id === u.id ? { ...x, role } : x));
      notify(`${u.name}'s role changed to ${role}`);
    } catch (e) { notify("Error: " + e.message, "error"); }
    setActioning(null);
  };

  // ── Delete ─────────────────────────────────────────────────────
  const deleteUser = async (u) => {
    setActioning(u.id);
    try {
      // Delete profile row (Supabase auth user must be deleted via service-role key in production)
      const q = `id=eq.${u.id}`;
      await fetch(`${SUPABASE_URL}/rest/v1/profiles?${q}`, {
        method: "DELETE",
        headers: { ...sb._headers(), "Prefer": "return=representation" },
      });
      setUsers(us => us.filter(x => x.id !== u.id));
      notify(`${u.name} deleted`);
    } catch (e) { notify("Error: " + e.message, "error"); }
    setActioning(null);
    setConfirmDel(null);
  };

  // ── Admin Reset Password ───────────────────────────────────────
  const sendPasswordReset = async (u) => {
    setNewPassErr("");
    try {
      // Trigger Supabase password-reset email (works without service key)
      const res = await fetch(`${SUPABASE_URL}/auth/v1/recover`, {
        method: "POST",
        headers: { "Content-Type":"application/json", "apikey": SUPABASE_ANON },
        body: JSON.stringify({ email: u.email }),
      });
      if (!res.ok) {
        const d = await res.json();
        throw new Error(d.error_description || d.message || "Failed");
      }
      notify(`Password reset email sent to ${u.email}`);
      setResetModal(null);
    } catch (e) { setNewPassErr(e.message); }
  };

  // ── Add User ───────────────────────────────────────────────────
  const handleAddUser = async () => {
    setAddErr("");
    if (!addForm.name.trim())          return setAddErr("Name is required.");
    if (!addForm.email.includes("@"))  return setAddErr("Valid email required.");
    if (addForm.password.length < 8)   return setAddErr("Password must be at least 8 characters.");
    setAddLoading(true);
    try {
      const session = await sb.signUp(addForm.email, addForm.password, {
        name: addForm.name.trim(),
        role: addForm.role,
        company: addForm.company.trim(),
      });
      // Insert profile with approved status (admin-created accounts skip approval)
      await sb.insert("profiles", {
        id: session.user?.id,
        name: addForm.name.trim(),
        role: addForm.role,
        company: addForm.company.trim(),
        status: "approved",
        avatar: addForm.name.trim().slice(0,2).toUpperCase(),
      });
      notify(`User ${addForm.name} created successfully`);
      setAddModal(false);
      setAddForm({ name:"", email:"", role:"user", company:"", password:"" });
      loadUsers();
    } catch (e) { setAddErr(e.message); }
    setAddLoading(false);
  };

  // ── Login visibility (view as user) ───────────────────────────
  const loginAsUser = (u) => {
    // Opens a special read-only view — in production this would use a service-role impersonation token.
    // Here we show the user's profile details in a modal as "view-as".
    setViewUser(u);
  };

  // ── Filters ────────────────────────────────────────────────────
  const filtered = users.filter(u => {
    const matchSearch = !search ||
      u.name?.toLowerCase().includes(search.toLowerCase()) ||
      u.email?.toLowerCase().includes(search.toLowerCase()) ||
      u.company?.toLowerCase().includes(search.toLowerCase());
    const matchRole   = filterRole   === "all" || u.role   === filterRole;
    const matchStatus = filterStatus === "all" || u.status === filterStatus;
    return matchSearch && matchRole && matchStatus;
  });

  const statusColor = s => s === "approved" ? "green" : s === "pending" ? "amber" : s === "on_hold" ? "purple" : s === "rejected" ? "red" : "gray";
  const roleColor   = r => r === "admin" ? "red" : r === "CA" ? "purple" : r === "Accountant" ? "blue" : "gray";

  const ROLES = ["user","Accountant","CA","admin"];

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
          <h2 style={{ fontSize:22, fontWeight:900, letterSpacing:"-0.5px", background:"linear-gradient(135deg,#eef2ff 50%,#3d7fff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>User Management</h2>
          <p style={{ color:T.textMid, fontSize:13, marginTop:3 }}>Admin-only panel · full user control</p>
        </div>
        {tab === "users" && <Btn icon="+" onClick={() => setAddModal(true)}>Add User</Btn>}
      </div>

      {/* Tabs */}
      <div style={{ display:"flex", background:T.surface, borderRadius:11, padding:4, marginBottom:20, border:`1px solid ${T.border}`, width:"fit-content", gap:2 }}>
        {[["users","👥 Users"],["approvals","⏳ Approvals"]].map(([t,label]) => (
          <button key={t} onClick={()=>setTab(t)}
            style={{ padding:"7px 18px", borderRadius:8, border:"none", cursor:"pointer", fontFamily:T.font, fontSize:13, fontWeight:tab===t?600:400, transition:"all 0.2s",
              background:tab===t?T.accent:"transparent", color:tab===t?"#fff":T.textMid,
              boxShadow:tab===t?`0 0 16px ${T.accentGlow}`:"none" }}>
            {label}
          </button>
        ))}
      </div>

      {tab === "approvals" && <AdminApprovalPanel user={adminUser} onClose={()=>{}} />}
      {tab === "users" && (
      <div>
      {/* Stats row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:20 }}>
        {[
          ["👥","Total Users",    users.length,                                           T.accent],
          ["✅","Active",         users.filter(u=>u.status==="approved").length,           T.green],
          ["⏳","Pending",        users.filter(u=>u.status==="pending").length,            T.amber],
          ["🔒","On Hold",        users.filter(u=>u.status==="on_hold").length,            T.purple],
        ].map(([icon,label,val,color]) => (
          <Card key={label} style={{ padding:"14px 18px" }}>
            <div style={{ fontSize:20, marginBottom:4 }}>{icon}</div>
            <div style={{ fontSize:22, fontWeight:700, color }}>{val}</div>
            <div style={{ fontSize:11, color:T.textDim }}>{label}</div>
          </Card>
        ))}
      </div>

      {/* Filters */}
      <Card style={{ marginBottom:16, padding:"14px 18px" }}>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          <div style={{ flex:1, minWidth:200 }}>
            <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, email, company…" prefix="🔍" />
          </div>
          <select value={filterRole} onChange={e=>setFilterRole(e.target.value)} style={{ padding:"8px 12px", borderRadius:8, fontSize:12, minWidth:120 }}>
            <option value="all">All Roles</option>
            {ROLES.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
          <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{ padding:"8px 12px", borderRadius:8, fontSize:12, minWidth:130 }}>
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
            <span style={{ fontSize:28, animation:"pulse 1.5s infinite" }}>⏳</span>
            <p style={{ marginTop:12, fontSize:13 }}>Loading users…</p>
          </div>
        ) : filtered.length === 0 ? (
          <div style={{ textAlign:"center", padding:"60px 0", color:T.textDim }}>
            <div style={{ fontSize:36, marginBottom:10 }}>🔍</div>
            <p style={{ fontSize:14 }}>No users match your filters</p>
          </div>
        ) : (
          <>
            {/* Table header */}
            <div style={{ display:"grid", gridTemplateColumns:"2fr 1.5fr 1fr 1fr 1.8fr", gap:0, padding:"10px 20px", borderBottom:`1px solid ${T.border}`, background:T.surface }}>
              {["User","Company","Role","Status","Actions"].map(h => (
                <span key={h} style={{ fontSize:11, fontWeight:600, color:T.textDim, textTransform:"uppercase", letterSpacing:"0.06em" }}>{h}</span>
              ))}
            </div>
            {filtered.map((u, idx) => (
              <div key={u.id} className="row-hover" style={{ display:"grid", gridTemplateColumns:"2fr 1.5fr 1fr 1fr 1.8fr", gap:0, padding:"13px 20px", borderBottom: idx < filtered.length-1 ? `1px solid ${T.border}` : "none", alignItems:"center", transition:"background 0.15s" }}>
                {/* User */}
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ width:36, height:36, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, color:"#fff", flexShrink:0 }}>
                    {(u.avatar || (u.name||"?").slice(0,2)).toUpperCase()}
                  </div>
                  <div style={{ minWidth:0 }}>
                    <div style={{ fontSize:13, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.name || "—"}</div>
                    <div style={{ fontSize:11, color:T.textDim, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.email || "—"}</div>
                  </div>
                </div>
                {/* Company */}
                <div style={{ fontSize:12, color:T.textMid, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{u.company || <span style={{color:T.textDim}}>—</span>}</div>
                {/* Role */}
                <div>
                  <Pill color={roleColor(u.role)} size="xs">{u.role || "user"}</Pill>
                </div>
                {/* Status */}
                <div>
                  <Pill color={statusColor(u.status)} size="xs" dot>{u.status || "unknown"}</Pill>
                </div>
                {/* Actions */}
                <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                  {/* View */}
                  <button title="View Profile" disabled={actioning===u.id}
                    onClick={async () => {
                      // Resolve approved_by UUID → name using already-loaded users list
                      const approvedByName = u.approved_by
                        ? (users.find(x => x.id === u.approved_by)?.name || u.approved_by)
                        : "—";
                      // Fetch email from Supabase auth (admin users endpoint)
                      let email = u.email || "";
                      if (!email) {
                        try {
                          const res = await fetch(`${SUPABASE_URL}/auth/v1/admin/users/${u.id}`, {
                            headers: { "apikey": SUPABASE_ANON, "Authorization": `Bearer ${sb._token || SUPABASE_ANON}` }
                          });
                          if (res.ok) { const d = await res.json(); email = d.email || ""; }
                        } catch {}
                      }
                      setViewUser({ ...u, email, approvedByName });
                    }}
                    style={{ padding:"5px 9px", borderRadius:7, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.border}`, background:T.surface, color:T.textMid, transition:"all 0.15s" }}>
                    👁
                  </button>
                  {/* Hold / Unhold */}
                  <button title={u.status==="on_hold"?"Unhold":"Put on Hold"} disabled={actioning===u.id || u.id===adminUser.id}
                    onClick={() => toggleHold(u)}
                    style={{ padding:"5px 9px", borderRadius:7, fontSize:11, fontWeight:600, fontFamily:T.font, cursor: (actioning===u.id||u.id===adminUser.id)?"not-allowed":"pointer", border:`1px solid ${u.status==="on_hold"?T.green+"66":T.purple+"66"}`, background:u.status==="on_hold"?T.greenDim:T.purpleDim, color:u.status==="on_hold"?T.green:T.purple, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1, transition:"all 0.15s" }}>
                    {u.status==="on_hold" ? "▶" : "⏸"}
                  </button>
                  {/* Reset Password */}
                  <button title="Reset Password" disabled={actioning===u.id}
                    onClick={() => { setResetModal(u); setNewPassErr(""); }}
                    style={{ padding:"5px 9px", borderRadius:7, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:actioning===u.id?"not-allowed":"pointer", border:`1px solid ${T.accent}44`, background:T.accentDim, color:T.accent, opacity:actioning===u.id?0.45:1, transition:"all 0.15s" }}>
                    🔑
                  </button>
                  {/* Change Role */}
                  <select title="Change Role" disabled={actioning===u.id || u.id===adminUser.id}
                    value={u.role || "user"}
                    onChange={e => changeRole(u, e.target.value)}
                    style={{ padding:"4px 7px", borderRadius:7, fontSize:11, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.border}`, background:T.surface, color:T.textMid, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1 }}>
                    {ROLES.map(r => <option key={r} value={r}>{r}</option>)}
                  </select>
                  {/* Delete */}
                  <button title="Delete User" disabled={actioning===u.id || u.id===adminUser.id}
                    onClick={() => setConfirmDel(u)}
                    style={{ padding:"5px 9px", borderRadius:7, fontSize:11, fontWeight:600, fontFamily:T.font, cursor:(actioning===u.id||u.id===adminUser.id)?"not-allowed":"pointer", border:`1px solid ${T.red}44`, background:T.redDim, color:T.red, opacity:(actioning===u.id||u.id===adminUser.id)?0.45:1, transition:"all 0.15s" }}>
                    🗑
                  </button>
                </div>
              </div>
            ))}
          </>
        )}
      </Card>
      </div>)}

      {/* ── Delete confirm modal ────────────────────────────────── */}
      <Modal open={!!confirmDel} onClose={() => setConfirmDel(null)} title="⚠ Confirm Delete" width={420}>
        {confirmDel && (
          <div>
            <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:10, padding:"14px 16px", marginBottom:18 }}>
              <p style={{ fontSize:13, color:T.textMid, lineHeight:1.7 }}>
                You are about to <strong style={{color:T.red}}>permanently delete</strong> the account for:<br/>
                <strong style={{color:T.text}}>{confirmDel.name}</strong> ({confirmDel.email})<br/>
                This action cannot be undone.
              </p>
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <Btn variant="secondary" fullWidth onClick={() => setConfirmDel(null)}>Cancel</Btn>
              <Btn variant="danger" fullWidth icon="🗑" onClick={() => deleteUser(confirmDel)}>Delete Permanently</Btn>
            </div>
          </div>
        )}
      </Modal>

      {/* ── Reset Password modal ────────────────────────────────── */}
      <Modal open={!!resetModal} onClose={() => setResetModal(null)} title="🔑 Reset Password" width={420}>
        {resetModal && (
          <div>
            <div style={{ background:T.accentDim, border:`1px solid ${T.accent}33`, borderRadius:10, padding:"12px 16px", marginBottom:16 }}>
              <p style={{ fontSize:12, color:T.textMid, lineHeight:1.7 }}>
                Send a password-reset email to <strong style={{color:T.text}}>{resetModal.name}</strong> ({resetModal.email}).<br/>
                They will receive a secure link to set a new password.
              </p>
            </div>
            {newPassErr && (
              <div style={{ background:T.redDim, border:`1px solid ${T.red}33`, borderRadius:8, padding:"9px 13px", fontSize:12, color:T.red, marginBottom:12 }}>✕ {newPassErr}</div>
            )}
            <div style={{ display:"flex", gap:10 }}>
              <Btn variant="secondary" fullWidth onClick={() => setResetModal(null)}>Cancel</Btn>
              <Btn variant="primary" fullWidth icon="📧" onClick={() => sendPasswordReset(resetModal)}>Send Reset Email</Btn>
            </div>
          </div>
        )}
      </Modal>

      {/* ── View User modal ─────────────────────────────────────── */}
      <Modal open={!!viewUser} onClose={() => setViewUser(null)} title="👤 User Profile" width={480}>
        {viewUser && (
          <div>
            <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:20, padding:"16px", background:T.surface, borderRadius:12 }}>
              <div style={{ width:60, height:60, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, fontWeight:700, color:"#fff", flexShrink:0 }}>
                {(viewUser.avatar||(viewUser.name||"?").slice(0,2)).toUpperCase()}
              </div>
              <div>
                <div style={{ fontSize:17, fontWeight:700, color:T.text, marginBottom:4 }}>{viewUser.name}</div>
                <div style={{ fontSize:12, color:T.textDim, marginBottom:6 }}>{viewUser.email || <span style={{color:T.textDim,fontStyle:"italic"}}>loading email…</span>}</div>
                <div style={{ display:"flex", gap:6 }}>
                  <Pill color={roleColor(viewUser.role)} size="xs">{viewUser.role||"user"}</Pill>
                  <Pill color={statusColor(viewUser.status)} size="xs" dot>{viewUser.status}</Pill>
                </div>
              </div>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:8, fontSize:13 }}>
              {[
                ["🏢","Company",   (viewUser.company && isNaN(viewUser.company)) ? viewUser.company : (viewUser.company ? `ID: ${viewUser.company}` : "—")],
                ["✉","Email",     viewUser.email || "—"],
                ["🪪","User ID",   viewUser.id],
                ["📅","Joined",    viewUser.created_at ? new Date(viewUser.created_at).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}) : "—"],
                ["✅","Approved By", viewUser.approvedByName || viewUser.approved_by || "—"],
                ["🕒","Approved At", viewUser.approved_at ? new Date(viewUser.approved_at).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}) : "—"],
              ].map(([icon, label, val]) => (
                <div key={label} style={{ display:"flex", justifyContent:"space-between", padding:"8px 12px", background:T.surface, borderRadius:8 }}>
                  <span style={{ color:T.textDim }}>{icon} {label}</span>
                  <span style={{ color:T.text, fontWeight:500, wordBreak:"break-all", maxWidth:220, textAlign:"right" }}>{val}</span>
                </div>
              ))}
            </div>
            <div style={{ display:"flex", gap:8, marginTop:16 }}>
              <Btn variant="secondary" fullWidth onClick={() => setViewUser(null)}>Close</Btn>
              <Btn variant="outline" fullWidth icon="🔑" onClick={() => { setViewUser(null); setResetModal(viewUser); setNewPassErr(""); }}>Reset Password</Btn>
              {viewUser.status === "on_hold"
                ? <Btn variant="success" fullWidth icon="▶" onClick={() => { toggleHold(viewUser); setViewUser(null); }}>Reactivate</Btn>
                : <Btn variant="amber" fullWidth icon="⏸" onClick={() => { toggleHold(viewUser); setViewUser(null); }}>Put on Hold</Btn>
              }
            </div>
          </div>
        )}
      </Modal>

      {/* ── Add User modal ──────────────────────────────────────── */}
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
            <select value={addForm.role} onChange={e=>setAddForm(f=>({...f,role:e.target.value}))} style={{ width:"100%", padding:"8px 10px" }}>
              {["user","Accountant","CA","admin"].map(r=><option key={r} value={r}>{r}</option>)}
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

export default function App() {
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
        localStorage.setItem(key, JSON.stringify(fresh));
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
    setHistory([]); // clear history so next user can't see previous user's data
    setRows([]); setHeaders([]); setRawRows([]); setFilename("");
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
      const restoreFromSession = (profile) => {
        const meta = session.user.user_metadata || {};
        const name = profile?.name || meta.name || session.user.email.split("@")[0];
        const role = profile?.role || meta.role || "user";
        const status = profile?.status || (role === "admin" ? "approved" : "pending");
        const avatar = profile?.avatar || name.slice(0,2).toUpperCase();
        if (status !== "approved") { localStorage.removeItem("sb_session"); return; }
        setUser({ id: session.user.id, name, role, status, avatar, company: profile?.company || "", email: session.user.email, sessionToken: session.access_token });
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
    setHeaders(result.headers); setRawRows(result.rows); setFilename(fname); setTemplateKey(tmplKey||"");
    setScreen(SCREENS.COLUMN_MAP);
    toast(`Parsed ${result.rows.length} rows from ${fname}`,"success");
  };

  const onMapped = (m) => {
    setMapping(m);
    const built = rawRows.map(r => {
      const get = field => m[field] ? r[headers.indexOf(m[field])] : "";
      const narr = String(get("narration")||"").trim();
      const ai = autoDetectLedger ? aiLedger(narr) : defaultLedger;
      return { id:genId(), date:get("date"), narration:narr, debit:get("debit"), credit:get("credit"), balance:get("balance"), ref:get("ref"), aiLedger:ai, ledger:ai, isDuplicate:false };
    }).filter(r => r.date||r.debit||r.credit);
    setRows(detectDuplicates(built));
    setScreen(SCREENS.LEDGER);
    const dups = built.filter(r=>r.isDuplicate).length;
    toast(`${built.length} transactions processed${dups>0?` · ${dups} duplicates detected`:""}`, dups>0?"warn":"success");
  };

  const onImport = () => {
    const companies = tally.companies.filter(c=>selectedCompanies.includes(c.id)).map(c=>c.name).join(", ") || selectedCompanies.join(", ");
    const validRows = rows.filter(r=>!r.isDuplicate||r.forceImport);
    const entry = {
      id:genId(), filename, date:new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}),
      rawDate:new Date().toISOString(), rows:validRows.length, company:companies, status:"Imported",
      suspense:validRows.filter(r=>r.ledger==="Suspense Account").length,
      duplicates:rows.filter(r=>r.isDuplicate).length, rows_data:rows,
    };
    setHistory(h => {
      const updated = [{ ...entry, savedAt: new Date().toISOString() }, ...h];
      saveHistory(user.id, updated); // save per-user
      return updated;
    });
    toast(`✓ ${validRows.length} vouchers pushed to Tally for ${companies}`,"success");
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
    { id:SCREENS.DASHBOARD, label:"Dashboard",  icon:"📊" },
    { id:SCREENS.UPLOAD,    label:"New Import",  icon:"⬆" },
    { id:SCREENS.HISTORY,   label:"History",     icon:"📋" },
    { id:SCREENS.SETTINGS,  label:"Settings",    icon:"⚙" },
    ...(user?.role === "admin" ? [
      { id:SCREENS.USER_MGMT, label:"Users",     icon:"👥", badge: pendingCount > 0 ? pendingCount : null },
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
          <div style={{ padding:"0 20px 20px", borderBottom:`1px solid ${T.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
              <div style={{ width:38, height:38, borderRadius:11, background:"linear-gradient(145deg,#1a4fd6,#3d7fff,#7c3aed)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, boxShadow:"0 0 24px rgba(61,127,255,0.3)" }}>🏦</div>
              <div>
                <div style={{ fontWeight:900, fontSize:15, letterSpacing:"-0.5px", background:"linear-gradient(135deg,#eef2ff,#3d7fff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>Bank2Tally</div>
                <div style={{ fontSize:9, color:T.textDim, letterSpacing:"0.08em", textTransform:"uppercase", fontWeight:600 }}>by Verma Consultancy</div>
              </div>
            </div>
            <div style={{ background:"rgba(61,127,255,0.06)", borderLeft:`2px solid ${T.accent}44`, borderRadius:"0 6px 6px 0", padding:"6px 10px" }}>
              <p style={{ fontSize:10, color:T.textMid, lineHeight:1.5, fontStyle:"italic" }}>"{todayQuote.text.slice(0,72)}{todayQuote.text.length>72?"…":""}"</p>
            </div>
          </div>
          <nav style={{ flex:1, padding:"14px 10px" }}>
            {NAV.map(n=>(
              <button key={n.id} onClick={()=>setScreen(n.id)}
                style={{ width:"100%", display:"flex", alignItems:"center", gap:10, padding:"9px 12px", borderRadius:9, fontSize:13, fontWeight:screen===n.id?600:400, fontFamily:T.font, cursor:"pointer", border:n.badge?`1px solid ${T.amber}66`:"none", marginBottom:3, transition:"all 0.15s",
                  background: screen===n.id ? T.accentDim : n.badge ? T.amberDim : "transparent",
                  color: screen===n.id ? T.accent : n.badge ? T.amber : T.textMid,
                  boxShadow: screen===n.id ? `0 0 12px ${T.accentGlow}` : "none" }}>
                <span style={{fontSize:16}}>{n.icon}</span>
                <span style={{flex:1,textAlign:"left"}}>{n.label}</span>
                {n.badge && (
                  <span style={{ background:T.amber, color:"#000", borderRadius:99, fontSize:10, fontWeight:700, padding:"1px 7px", minWidth:18, textAlign:"center" }}>
                    {n.badge}
                  </span>
                )}
              </button>
            ))}
          </nav>
          <div style={{ padding:"14px 14px", borderTop:`1px solid ${T.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
              <div style={{ width:32, height:32, borderRadius:"50%", background:`linear-gradient(135deg,${T.accent},${T.purple})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, color:"#fff" }}>{user.avatar}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:12, fontWeight:600, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{user.name.split(" ")[0]}</div>
                <div style={{ fontSize:10, color:T.textDim }}>{user.role}</div>
              </div>
            </div>
            <button onClick={onLogout}
              style={{ width:"100%", display:"flex", alignItems:"center", justifyContent:"center", gap:8, padding:"8px 12px", borderRadius:8, fontSize:12, fontWeight:600, fontFamily:T.font, cursor:"pointer", border:`1px solid ${T.red}44`, background:T.redDim, color:T.red, transition:"all 0.15s" }}>
              <span>⏻</span> Sign Out
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
          {screen === SCREENS.PREVIEW && <PreviewScreen rows={rows} filename={filename} selectedCompanies={selectedCompanies} onBack={()=>setScreen(SCREENS.LEDGER)} onImport={onImport} auditLog={auditLog} tally={tally} />}
          {screen === SCREENS.HISTORY && <HistoryScreen history={history} onReimport={onReimport} onDeleteEntry={deleteHistoryEntry} onClearAll={clearAllHistory} onBack={()=>setScreen(SCREENS.DASHBOARD)} />}
          {screen === SCREENS.SETTINGS && <SettingsScreen user={user} onLogout={onLogout} tally={tally} tallyHost={tallyHost} setTallyHost={setTallyHost} tallyPort={tallyPort} setTallyPort={setTallyPort} defaultLedger={defaultLedger} setDefaultLedger={setDefaultLedger} autoDetectLedger={autoDetectLedger} setAutoDetectLedger={setAutoDetectLedger} />}
          {screen === SCREENS.USER_MGMT && isAdmin && <UserManagementScreen adminUser={user} />}
        </div>
      </div>
    </>
  );
}
