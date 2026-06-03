import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── helpers ────────────────────────────────────────────────────────────────

function normalizePhone(raw) {
  if (raw == null) return "";
  return String(raw).replace(/\D/g, "").replace(/^0+/, "").slice(-12);
}

function extractPhoneFromTicketName(name) {
  if (!name) return "";
  const digits = String(name).replace(/\D/g, "");
  return digits.slice(-12);
}

// Normalize ANY input (JS Date, "DD/MM/YYYY ...", "YYYY-MM-DD ...") to "YYYY-MM-DD".
// IMPORTANT: for Date objects we read LOCAL components, not toISOString(),
// because toISOString() converts to UTC and can shift the day for evening tickets.
function toDateStr(val) {
  if (val === null || val === undefined || val === "") return "";

  if (val instanceof Date && !isNaN(val)) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  const s = String(val).trim();

  // DD/MM/YYYY or D/M/YYYY (the spooler format), with optional trailing time
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) {
    const [, d, m, y] = dmy;
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }

  // Already ISO-ish: YYYY-MM-DD...
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];

  return s.slice(0, 10);
}

// Business rule: feedback must go out on the SAME calendar day as the ticket.
// Next-day (or later) counts as a late submission. To allow a wider window,
// bump FEEDBACK_DAY_TOLERANCE to 1 (next day OK), 2, etc.
const FEEDBACK_DAY_TOLERANCE = 0;

function datesWithin(d1, d2, tol = FEEDBACK_DAY_TOLERANCE) {
  if (!d1 || !d2) return false;
  const a = new Date(`${d1}T00:00:00`);
  const b = new Date(`${d2}T00:00:00`);
  if (isNaN(a) || isNaN(b)) return false;
  const diffDays = Math.abs(a - b) / 86400000;
  return diffDays <= tol;
}

function parseSpooler(rows) {
  // returns Map: phone -> Set of canonical date strings (YYYY-MM-DD)
  const map = new Map();
  rows.forEach((r) => {
    const dest = r["DEST"] ?? r["dest"] ?? r["Dest"] ?? "";
    const createdAt =
      r["CREATED AT"] ?? r["created at"] ?? r["Created At"] ?? r["createdAt"] ?? "";
    const phone = normalizePhone(dest);
    const date = toDateStr(createdAt);
    if (!phone || !date) return;
    if (!map.has(phone)) map.set(phone, new Set());
    map.get(phone).add(date);
  });
  return map;
}

function processTickets(ticketRows, spoolerMap) {
  const seen = new Map();

  const results = ticketRows.map((row, idx) => {
    const ticketName = row["Ticket name"] ?? row["ticket name"] ?? row["TICKET NAME"] ?? "";
    const customerCli = row["Customer CLI"] ?? row["customer cli"] ?? row["CUSTOMER CLI"] ?? "";
    const createdAt = row["Created At"] ?? row["created at"] ?? row["CREATED AT"] ?? "";

    const phoneFromCli = normalizePhone(customerCli);
    const phoneFromName = extractPhoneFromTicketName(ticketName);
    const phone = phoneFromCli || phoneFromName;
    const date = toDateStr(createdAt);

    const key = `${phone}__${date}`;
    const isDuplicate = seen.has(key);
    if (!isDuplicate) seen.set(key, idx);

    const existsInSpooler = spoolerMap.has(phone);
    // A feedback message counts only if it was sent within FEEDBACK_DAY_TOLERANCE
    // days of the ticket. With tolerance 0 this means the exact same calendar day;
    // anything later falls through to "Date mismatch (late feedback)".
    const sameDateInSpooler =
      existsInSpooler &&
      [...spoolerMap.get(phone)].some((sd) => datesWithin(sd, date));

    let status = "Valid";
    let reason = "";

    if (isDuplicate) {
      status = "Invalid";
      reason = "Duplicate ticket";
    } else if (!existsInSpooler) {
      status = "Invalid";
      reason = "Phone not in spooler";
    } else if (!sameDateInSpooler) {
      status = "Invalid";
      reason = "Date mismatch (late feedback)";
    }

    return { ...row, Status: status, Reason: reason, _phone: phone, _date: date };
  });

  // Duplicate rule: the FIRST ticket for a given phone+date is kept as-is
  // (valid if it otherwise passes); every later copy is flagged "Duplicate ticket".
  // This is handled inline above via the `seen` map, so no second pass is needed —
  // an earlier version re-marked every copy invalid, which wrongly voided the
  // original too.
  return results;
}

// Parse a CSV/TSV text blob into row objects WITHOUT date coercion.
// This is the key fix: the spooler is a tab-separated text file whose dates
// are DD/MM/YYYY. If we hand it to SheetJS, it guesses US MM/DD and locks in
// the wrong date (e.g. 03/06/2026 -> March 6 instead of June 3). Keeping the
// raw string lets toDateStr() apply the correct DD/MM parsing.
function parseDelimitedText(text) {
  const lines = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter((l) => l.length);
  if (!lines.length) return [];
  const delim = lines[0].includes("\t") ? "\t" : ",";
  const headers = lines[0].split(delim).map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const cells = line.split(delim);
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = (cells[i] ?? "").trim();
    });
    return obj;
  });
}

async function readFile(file) {
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  // Real .xlsx starts with "PK" (zip); legacy .xls starts with the OLE magic D0 CF.
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  const isOle = bytes[0] === 0xd0 && bytes[1] === 0xcf;

  if (!isZip && !isOle) {
    // Plain text masquerading as .xls/.csv — parse manually, preserve raw dates.
    const text = new TextDecoder("utf-8").decode(buf);
    return parseDelimitedText(text);
  }

  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function downloadFile(results, format) {
  // Long numeric IDs (Ticket number, Customer CLI) get rendered as scientific
  // notation (e.g. 1.78042E+12) when written as numeric cells, destroying the ID.
  // Force ID-like columns to strings so the full digits survive the export.
  const ID_COLUMNS = ["Ticket number", "Customer CLI"];
  const clean = results.map(({ _phone, _date, ...rest }) => {
    const row = { ...rest };
    ID_COLUMNS.forEach((col) => {
      if (row[col] !== undefined && row[col] !== null && row[col] !== "") {
        row[col] = String(row[col]);
      }
    });
    return row;
  });
  if (format === "csv") {
    const ws = XLSX.utils.json_to_sheet(clean);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "validated_tickets.csv";
    a.click();
  } else {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(clean);
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    XLSX.writeFile(wb, "validated_tickets.xlsx");
  }
}

// ─── components ──────────────────────────────────────────────────────────────

function DropZone({ label, icon, file, onFile, accept }) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef();

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDrag(false);
      const f = e.dataTransfer.files[0];
      if (f) onFile(f);
    },
    [onFile]
  );

  return (
    <div
      className={`dropzone ${drag ? "drag-over" : ""} ${file ? "has-file" : ""}`}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])}
      />
      <div className="drop-icon">{icon}</div>
      <div className="drop-label">{label}</div>
      {file ? (
        <div className="drop-filename">📎 {file.name}</div>
      ) : (
        <div className="drop-hint">Drop file here or click to browse<br /><span>.xlsx · .xls · .csv</span></div>
      )}
    </div>
  );
}

function Badge({ status }) {
  return (
    <span className={`badge badge-${status.toLowerCase()}`}>{status}</span>
  );
}

export default function App() {
  const [spoolerFile, setSpoolerFile] = useState(null);
  const [ticketsFile, setTicketsFile] = useState(null);
  const [results, setResults] = useState(null);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");

  const handleProcess = async () => {
    if (!spoolerFile || !ticketsFile) return;
    setLoading(true);
    setError(null);
    try {
      const [spoolerRows, ticketRows] = await Promise.all([
        readFile(spoolerFile),
        readFile(ticketsFile),
      ]);
      const spoolerMap = parseSpooler(spoolerRows);
      const processed = processTickets(ticketRows, spoolerMap);
      const valid = processed.filter((r) => r.Status === "Valid").length;
      const invalid = processed.length - valid;
      setResults(processed);
      setStats({ total: processed.length, valid, invalid });
    } catch (e) {
      setError("Error processing files: " + e.message);
    }
    setLoading(false);
  };

  const filtered = results
    ? results.filter((r) => {
        const matchFilter = filter === "All" || r.Status === filter;
        const matchSearch =
          !search ||
          Object.values(r).some((v) =>
            String(v).toLowerCase().includes(search.toLowerCase())
          );
        return matchFilter && matchSearch;
      })
    : [];

  return (
    <div className="app">
      <header>
        <div className="header-inner">
          <div className="logo">
            <span className="logo-mark">✦</span>
            <span className="logo-text">TicketLens</span>
          </div>
          <p className="tagline">WhatsApp Feedback · Spooler Validator</p>
        </div>
      </header>

      <main>
        {/* Upload Section */}
        <section className="upload-section">
          <div className="section-label">01 — Upload Files</div>
          <div className="dropzone-grid">
            <DropZone
              label="Spooler File"
              icon="📤"
              file={spoolerFile}
              onFile={setSpoolerFile}
              accept=".xlsx,.xls,.csv"
            />
            <div className="arrow-divider">→</div>
            <DropZone
              label="Tickets File"
              icon="🎫"
              file={ticketsFile}
              onFile={setTicketsFile}
              accept=".xlsx,.xls,.csv"
            />
          </div>

          <button
            className={`process-btn ${loading ? "loading" : ""}`}
            disabled={!spoolerFile || !ticketsFile || loading}
            onClick={handleProcess}
          >
            {loading ? (
              <><span className="spinner" /> Processing…</>
            ) : (
              <><span>⚡</span> Validate Tickets</>
            )}
          </button>

          {error && <div className="error-msg">{error}</div>}
        </section>

        {/* Results Section */}
        {results && (
          <section className="results-section">
            <div className="section-label">02 — Results</div>

            {/* Stats */}
            <div className="stats-row">
              <div className="stat-card stat-total">
                <div className="stat-num">{stats.total}</div>
                <div className="stat-lbl">Total Tickets</div>
              </div>
              <div className="stat-card stat-valid">
                <div className="stat-num">{stats.valid}</div>
                <div className="stat-lbl">Valid</div>
              </div>
              <div className="stat-card stat-invalid">
                <div className="stat-num">{stats.invalid}</div>
                <div className="stat-lbl">Invalid</div>
              </div>
              <div className="stat-card stat-rate">
                <div className="stat-num">
                  {stats.total ? Math.round((stats.valid / stats.total) * 100) : 0}%
                </div>
                <div className="stat-lbl">Valid Rate</div>
              </div>
            </div>

            {/* Filters & Search */}
            <div className="controls-row">
              <div className="filter-tabs">
                {["All", "Valid", "Invalid"].map((f) => (
                  <button
                    key={f}
                    className={`filter-tab ${filter === f ? "active" : ""}`}
                    onClick={() => setFilter(f)}
                  >
                    {f}
                    <span className="tab-count">
                      {f === "All"
                        ? stats.total
                        : f === "Valid"
                        ? stats.valid
                        : stats.invalid}
                    </span>
                  </button>
                ))}
              </div>
              <input
                className="search-input"
                placeholder="Search tickets…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>

            {/* Table */}
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Ticket Number</th>
                    <th>Phone</th>
                    <th>Date</th>
                    <th>Branch</th>
                    <th>Feedback</th>
                    <th>Status</th>
                    <th>Reason</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.slice(0, 200).map((row, i) => (
                    <tr key={i} className={row.Status === "Invalid" ? "row-invalid" : "row-valid"}>
                      <td className="td-num">{i + 1}</td>
                      <td className="td-ticket">{row["Ticket number"] ?? "—"}</td>
                      <td className="td-phone">{row._phone || "—"}</td>
                      <td className="td-date">{row._date || "—"}</td>
                      <td>{row["Branch Name"] ?? "—"}</td>
                      <td>{row["Feedback Head"] ?? "—"}</td>
                      <td><Badge status={row.Status} /></td>
                      <td className="td-reason">{row.Reason || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtered.length > 200 && (
                <div className="table-overflow-note">
                  Showing 200 of {filtered.length} rows. Download to see all.
                </div>
              )}
            </div>

            {/* Download */}
            <div className="download-row">
              <div className="download-label">03 — Export</div>
              <div className="download-btns">
                <button className="dl-btn dl-xlsx" onClick={() => downloadFile(results, "xlsx")}>
                  ⬇ Download XLSX
                </button>
                <button className="dl-btn dl-csv" onClick={() => downloadFile(results, "csv")}>
                  ⬇ Download CSV
                </button>
              </div>
            </div>
          </section>
        )}
      </main>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');

        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        :root {
          --bg: #0a0c0f;
          --surface: #111418;
          --surface2: #181c22;
          --border: #252a33;
          --accent: #00e5a0;
          --accent2: #ff5e5e;
          --accent3: #ffd166;
          --text: #e8edf5;
          --muted: #6b7685;
          --valid: #00e5a0;
          --invalid: #ff5e5e;
          --font-head: 'Syne', sans-serif;
          --font-mono: 'DM Mono', monospace;
        }

        body { background: var(--bg); color: var(--text); font-family: var(--font-head); }

        .app { min-height: 100vh; }

        /* Header */
        header {
          border-bottom: 1px solid var(--border);
          padding: 20px 32px;
          background: linear-gradient(135deg, #0a0c0f 0%, #0f1419 100%);
        }
        .header-inner { display: flex; align-items: center; gap: 20px; }
        .logo { display: flex; align-items: center; gap: 10px; }
        .logo-mark { font-size: 22px; color: var(--accent); }
        .logo-text { font-size: 22px; font-weight: 800; letter-spacing: -0.5px; }
        .tagline { font-size: 12px; color: var(--muted); font-family: var(--font-mono); }

        /* Main */
        main { max-width: 1200px; margin: 0 auto; padding: 40px 32px; }

        /* Section label */
        .section-label {
          font-family: var(--font-mono);
          font-size: 11px;
          color: var(--accent);
          letter-spacing: 2px;
          text-transform: uppercase;
          margin-bottom: 20px;
        }

        /* Upload */
        .upload-section { margin-bottom: 56px; }
        .dropzone-grid {
          display: flex;
          align-items: center;
          gap: 16px;
          margin-bottom: 24px;
        }
        .dropzone {
          flex: 1;
          border: 1.5px dashed var(--border);
          border-radius: 12px;
          padding: 32px 24px;
          text-align: center;
          cursor: pointer;
          transition: all 0.2s ease;
          background: var(--surface);
        }
        .dropzone:hover, .dropzone.drag-over {
          border-color: var(--accent);
          background: rgba(0, 229, 160, 0.04);
        }
        .dropzone.has-file {
          border-color: var(--accent);
          border-style: solid;
        }
        .drop-icon { font-size: 32px; margin-bottom: 10px; }
        .drop-label { font-size: 14px; font-weight: 700; margin-bottom: 8px; }
        .drop-hint { font-size: 12px; color: var(--muted); font-family: var(--font-mono); line-height: 1.6; }
        .drop-hint span { color: var(--accent); }
        .drop-filename {
          font-size: 12px;
          font-family: var(--font-mono);
          color: var(--accent);
          margin-top: 8px;
          word-break: break-all;
        }
        .arrow-divider { font-size: 20px; color: var(--muted); flex-shrink: 0; }

        .process-btn {
          display: flex;
          align-items: center;
          gap: 10px;
          background: var(--accent);
          color: #000;
          border: none;
          border-radius: 8px;
          padding: 14px 32px;
          font-size: 15px;
          font-weight: 700;
          font-family: var(--font-head);
          cursor: pointer;
          transition: all 0.2s;
          letter-spacing: 0.3px;
        }
        .process-btn:hover:not(:disabled) { background: #00ffb3; transform: translateY(-1px); }
        .process-btn:disabled { opacity: 0.4; cursor: not-allowed; }
        .process-btn.loading { background: var(--surface2); color: var(--accent); border: 1px solid var(--accent); }
        .spinner {
          width: 16px; height: 16px;
          border: 2px solid rgba(0,229,160,0.3);
          border-top-color: var(--accent);
          border-radius: 50%;
          animation: spin 0.7s linear infinite;
          display: inline-block;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .error-msg {
          margin-top: 16px;
          padding: 12px 16px;
          background: rgba(255,94,94,0.1);
          border: 1px solid rgba(255,94,94,0.3);
          border-radius: 8px;
          color: var(--invalid);
          font-size: 13px;
          font-family: var(--font-mono);
        }

        /* Stats */
        .stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 16px;
          margin-bottom: 28px;
        }
        .stat-card {
          background: var(--surface);
          border-radius: 12px;
          padding: 20px;
          border: 1px solid var(--border);
          text-align: center;
        }
        .stat-num { font-size: 32px; font-weight: 800; }
        .stat-lbl { font-size: 11px; font-family: var(--font-mono); color: var(--muted); margin-top: 4px; }
        .stat-total .stat-num { color: var(--text); }
        .stat-valid .stat-num { color: var(--valid); }
        .stat-invalid .stat-num { color: var(--invalid); }
        .stat-rate .stat-num { color: var(--accent3); }

        /* Controls */
        .controls-row {
          display: flex;
          align-items: center;
          gap: 16px;
          margin-bottom: 16px;
          flex-wrap: wrap;
        }
        .filter-tabs { display: flex; gap: 6px; }
        .filter-tab {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 8px 16px;
          border-radius: 8px;
          border: 1px solid var(--border);
          background: var(--surface);
          color: var(--muted);
          font-size: 13px;
          font-weight: 600;
          font-family: var(--font-head);
          cursor: pointer;
          transition: all 0.2s;
        }
        .filter-tab:hover { border-color: var(--accent); color: var(--text); }
        .filter-tab.active { background: rgba(0,229,160,0.1); border-color: var(--accent); color: var(--accent); }
        .tab-count {
          background: var(--surface2);
          border-radius: 4px;
          padding: 1px 6px;
          font-size: 11px;
          font-family: var(--font-mono);
        }
        .search-input {
          flex: 1;
          min-width: 200px;
          background: var(--surface);
          border: 1px solid var(--border);
          border-radius: 8px;
          padding: 8px 14px;
          color: var(--text);
          font-size: 13px;
          font-family: var(--font-mono);
          outline: none;
          transition: border-color 0.2s;
        }
        .search-input:focus { border-color: var(--accent); }
        .search-input::placeholder { color: var(--muted); }

        /* Table */
        .table-wrap {
          overflow-x: auto;
          border-radius: 12px;
          border: 1px solid var(--border);
          margin-bottom: 32px;
        }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        thead { background: var(--surface2); }
        th {
          padding: 12px 14px;
          text-align: left;
          font-family: var(--font-mono);
          font-size: 11px;
          color: var(--muted);
          text-transform: uppercase;
          letter-spacing: 1px;
          border-bottom: 1px solid var(--border);
          font-weight: 500;
        }
        td {
          padding: 11px 14px;
          border-bottom: 1px solid rgba(37,42,51,0.6);
          font-family: var(--font-mono);
          font-size: 12px;
          vertical-align: middle;
        }
        tr:last-child td { border-bottom: none; }
        tbody tr { background: var(--surface); transition: background 0.1s; }
        tbody tr:hover { background: var(--surface2); }
        .row-invalid td { opacity: 0.75; }

        .td-num { color: var(--muted); width: 44px; }
        .td-ticket { font-weight: 500; color: var(--text); }
        .td-phone { color: var(--accent); }
        .td-date { color: var(--accent3); }
        .td-reason { color: var(--invalid); font-size: 11px; }

        .badge {
          display: inline-flex;
          align-items: center;
          padding: 3px 10px;
          border-radius: 100px;
          font-size: 11px;
          font-weight: 700;
          font-family: var(--font-mono);
          letter-spacing: 0.5px;
        }
        .badge-valid { background: rgba(0,229,160,0.12); color: var(--valid); }
        .badge-invalid { background: rgba(255,94,94,0.12); color: var(--invalid); }

        .table-overflow-note {
          padding: 12px 16px;
          font-size: 12px;
          font-family: var(--font-mono);
          color: var(--muted);
          text-align: center;
          border-top: 1px solid var(--border);
          background: var(--surface);
        }

        /* Download */
        .download-row {
          display: flex;
          align-items: center;
          gap: 24px;
          flex-wrap: wrap;
        }
        .download-btns { display: flex; gap: 12px; }
        .dl-btn {
          padding: 12px 24px;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 700;
          font-family: var(--font-head);
          cursor: pointer;
          border: 1.5px solid;
          transition: all 0.2s;
        }
        .dl-xlsx {
          background: rgba(0,229,160,0.08);
          border-color: var(--accent);
          color: var(--accent);
        }
        .dl-xlsx:hover { background: rgba(0,229,160,0.18); }
        .dl-csv {
          background: rgba(255,209,102,0.08);
          border-color: var(--accent3);
          color: var(--accent3);
        }
        .dl-csv:hover { background: rgba(255,209,102,0.18); }

        @media (max-width: 700px) {
          main { padding: 24px 16px; }
          .dropzone-grid { flex-direction: column; }
          .arrow-divider { transform: rotate(90deg); }
          .stats-row { grid-template-columns: 1fr 1fr; }
        }
      `}</style>
    </div>
  );
}
