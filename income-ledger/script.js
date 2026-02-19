(() => {
  "use strict";

  // 실수로 script가 중복 include되어도 2번 초기화되지 않게 가드
  if (window.__incomeLedgerInitialized) return;
  window.__incomeLedgerInitialized = true;

  /* =========================================================
     DOM
     ========================================================= */
  const $id = (id) => document.getElementById(id);

  const el = {
    // Upload
    dropZone: $id("dropZone"),
    fileInput: $id("fileInput"),
    btnPick: $id("btnPick"),
    btnReset: $id("btnReset"),
    fileMeta: $id("fileMeta"),
    btnRun: $id("btnRun"),

    // Progress
    pLoad: $id("pLoad"),
    loadLabel: $id("loadLabel"),
    loadPct: $id("loadPct"),
    pAnalyze: $id("pAnalyze"),
    analyzeLabel: $id("analyzeLabel"),
    analyzePct: $id("analyzePct"),
    libStatus: $id("libStatus"),

    // Mapping
    mappingDetails: $id("mappingDetails"),
    mappingStatus: $id("mappingStatus"),
    inpHeaderRow: $id("inpHeaderRow"),
    selSeq: $id("selSeq"),
    selDatetime: $id("selDatetime"),
    selWithdraw: $id("selWithdraw"),
    selDeposit: $id("selDeposit"),
    selBalance: $id("selBalance"),
    selContent: $id("selContent"),
    selNote: $id("selNote"),
    selBranch: $id("selBranch"),
    btnApplyMapping: $id("btnApplyMapping"),

    // Log
    logCard: $id("logCard"),
    logSummary: $id("logSummary"),
    logList: $id("logList"),

    // Sections
    secDashboard: $id("secDashboard"),
    secDetails: $id("secDetails"),
    secDownload: $id("secDownload"),

    // Dashboard
    periodLabel: $id("periodLabel"),
    dashboardTotals: $id("dashboardTotals"),
    monthlyTable: $id("monthlyTable"),
    depositDaysTable: $id("depositDaysTable"),

    // Details
    selOrder: $id("selOrder"),
    selMode: $id("selMode"),
    dateList: $id("dateList"),
    detailPanel: $id("detailPanel"),

    // Download
    btnDownloadXlsx: $id("btnDownloadXlsx"),
    btnDownloadCsvAll: $id("btnDownloadCsvAll"),
    btnDownloadCsvDeposit: $id("btnDownloadCsvDeposit"),

    // Top
    btnToTop: $id("btnToTop"),
  };

  /* =========================================================
     State
     ========================================================= */
  const DEFAULT_START = {
    seq: 1,        // B
    datetime: 2,   // C
    withdraw: 3,   // D (E fallback)
    deposit: 5,    // F
    balance: 6,    // G (H fallback)
    content: 8,    // I (J fallback)
    note: 10,      // K (L/M fallback)
    branch: 13,    // N
  };

  const state = {
    file: null,

    // workbook cache
    wb: null,
    sheetName: "",
    ws: null,
    rows: [],
    colCount: 14,

    // mapping
    headerRowIdx: 11, // 0-based (12행)
    mapping: null,

    // meta
    meta: { owner: "", account: "", period: "" },

    // results
    txns: [],
    dailyList: [],
    monthlyList: [],
    totals: null,

    // original header cosmetics for export
    original: {
      fileName: "",
      sheetName: "",
      ws: null,
      headerRowIdx: null,
      headerAoa: [],
      allMerges: [],
      headerMerges: [],
      cols: null,
      colCount: 14,
    },

    // logs
    logs: [], // {level:'info'|'warn'|'error', text:string}
  };

  /* =========================================================
     UI helpers
     ========================================================= */
  function setBtnEnabled(btn, enabled) {
    if (!btn) return;
    btn.disabled = !enabled;
  }

  function setHidden(elm, hidden) {
    if (!elm) return;
    elm.hidden = !!hidden;
  }

  function clampPct(n) {
    const v = Number(n);
    if (!Number.isFinite(v)) return 0;
    return Math.max(0, Math.min(100, Math.round(v)));
  }

  function setLoadProgress(pct, label = "") {
    const p = clampPct(pct);
    if (el.pLoad) el.pLoad.value = p;
    if (el.loadPct) el.loadPct.textContent = `${p}%`;
    if (el.loadLabel) el.loadLabel.textContent = label;
  }

  function setAnalyzeProgress(pct, label = "") {
    const p = clampPct(pct);
    if (el.pAnalyze) el.pAnalyze.value = p;
    if (el.analyzePct) el.analyzePct.textContent = `${p}%`;
    if (el.analyzeLabel) el.analyzeLabel.textContent = label;
  }

  function setLibStatus(text, { error = false } = {}) {
    if (!el.libStatus) return;
    el.libStatus.textContent = text;
    el.libStatus.classList.toggle("err", !!error);
  }

  function resetResultsUI() {
    setHidden(el.secDashboard, true);
    setHidden(el.secDetails, true);
    setHidden(el.secDownload, true);

    if (el.dashboardTotals) el.dashboardTotals.innerHTML = "";
    if (el.monthlyTable) el.monthlyTable.innerHTML = "";
    if (el.depositDaysTable) el.depositDaysTable.innerHTML = "";
    if (el.dateList) el.dateList.innerHTML = "";
    if (el.detailPanel) el.detailPanel.innerHTML = "";

    setBtnEnabled(el.btnDownloadXlsx, false);
    setBtnEnabled(el.btnDownloadCsvAll, false);
    setBtnEnabled(el.btnDownloadCsvDeposit, false);
  }

  /* =========================================================
     Logging
     ========================================================= */
  function clearLogs() {
    state.logs = [];
    if (el.logSummary) el.logSummary.textContent = "";
    if (el.logList) el.logList.innerHTML = "";
    if (el.logCard) el.logCard.style.display = "none";
  }

  function addLog(level, text) {
    state.logs.push({ level, text: String(text) });
  }

  function flushLogs(summaryLines = []) {
    if (!el.logCard || !el.logSummary || !el.logList) return;

    el.logSummary.textContent = summaryLines.filter(Boolean).join("\n");
    el.logList.innerHTML = "";

    const max = 30;
    const shown = state.logs.slice(0, max);

    for (const item of shown) {
      const li = document.createElement("li");
      li.textContent = item.text;
      if (item.level === "error") li.style.color = "#b42318";
      if (item.level === "warn") li.style.color = "#b45309";
      el.logList.appendChild(li);
    }

    if (state.logs.length > max) {
      const li = document.createElement("li");
      li.textContent = `… 외 ${state.logs.length - max}건 더 있음`;
      li.className = "muted";
      el.logList.appendChild(li);
    }

    el.logCard.style.display = "block";
  }

  /* =========================================================
     XLSX loader (local → CDN fallback)
     ========================================================= */
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.async = true;
      s.onload = () => resolve(src);
      s.onerror = () => reject(new Error("failed: " + src));
      document.head.appendChild(s);
    });
  }

  async function ensureXLSX() {
    if (window.XLSX) {
      setLoadProgress(100, "이미 로드됨");
      setLibStatus("XLSX 라이브러리 로드 완료");
      return true;
    }

    setLoadProgress(10, "라이브러리 로드 중...");
    setLibStatus("로딩 중…");

    const sources = [
      "/static/xlsx.full.min.js",
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
    ];

    for (let i = 0; i < sources.length; i++) {
      const src = sources[i];
      try {
        setLoadProgress(20 + i * 40, `로드 시도: ${src}`);
        await loadScript(src);
        if (window.XLSX) {
          setLoadProgress(100, "로드 완료");
          setLibStatus(`xlsx.full.min.js 로드 완료 (${i === 0 ? "로컬" : "CDN"})`);
          return true;
        }
      } catch (_) {
        // try next
      }
    }

    setLoadProgress(0, "로드 실패");
    setLibStatus("오류: XLSX 라이브러리 로드 실패. /static/xlsx.full.min.js 경로를 확인하세요.", { error: true });
    return false;
  }

  /* =========================================================
     Parsing utilities
     ========================================================= */
  function bytesToHuman(bytes) {
    const b = Number(bytes);
    if (!Number.isFinite(b)) return "-";
    const units = ["B", "KB", "MB", "GB"];
    let v = b, i = 0;
    while (v >= 1024 && i < units.length - 1) { v /= 1024; i++; }
    return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
  }

  function decodeExcelEscapes(s) {
    const str = String(s ?? "");
    return str.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => {
      const code = parseInt(hex, 16);
      if (!Number.isFinite(code)) return "";
      return String.fromCharCode(code);
    });
  }

  function normalizeText(v) {
    if (v === null || v === undefined) return "";
    let s = String(v);
    s = decodeExcelEscapes(s);
    s = s.replace(/\u00A0/g, " ");
    s = s.replace(/\s+/g, " ").trim();
    return s;
  }

  function parseAmount(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number" && Number.isFinite(v)) return Math.round(v);
    const s0 = normalizeText(v);
    if (!s0) return 0;
    const s = s0.replace(/,/g, "").replace(/원/g, "").trim();
    if (!s) return 0;
    const n = Number(s);
    return Number.isFinite(n) ? Math.round(n) : 0;
  }

  function pad2(n) { return String(n).padStart(2, "0"); }

  function toDateKey(dateObj) {
    const y = dateObj.getFullYear();
    const m = pad2(dateObj.getMonth() + 1);
    const d = pad2(dateObj.getDate());
    return `${y}-${m}-${d}`;
  }

  function toMonthKey(dateObj) {
    const y = dateObj.getFullYear();
    const m = pad2(dateObj.getMonth() + 1);
    return `${y}-${m}`;
  }

  function formatDateTime(d) {
    const y = d.getFullYear();
    const m = pad2(d.getMonth() + 1);
    const day = pad2(d.getDate());
    const hh = pad2(d.getHours());
    const mm = pad2(d.getMinutes());
    const ss = pad2(d.getSeconds());
    return `${y}/${m}/${day} ${hh}:${mm}:${ss}`;
  }

  function parseDateTime(v) {
    if (!v && v !== 0) return { date: null, dateKey: "", monthKey: "", dateTimeStr: "" };

    if (v instanceof Date && !isNaN(v.getTime())) {
      return { date: v, dateKey: toDateKey(v), monthKey: toMonthKey(v), dateTimeStr: formatDateTime(v) };
    }

    // Excel serial number
    if (typeof v === "number" && Number.isFinite(v) && window.XLSX?.SSF?.parse_date_code) {
      const dc = XLSX.SSF.parse_date_code(v);
      if (dc && dc.y && dc.m && dc.d) {
        const date = new Date(dc.y, dc.m - 1, dc.d, dc.H || 0, dc.M || 0, dc.S || 0);
        return { date, dateKey: toDateKey(date), monthKey: toMonthKey(date), dateTimeStr: formatDateTime(date) };
      }
    }

    const s = normalizeText(v);
    const m = s.match(
      /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/
    );
    if (!m) return { date: null, dateKey: "", monthKey: "", dateTimeStr: s };

    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const hh = Number(m[4] ?? 0);
    const mm = Number(m[5] ?? 0);
    const ss = Number(m[6] ?? 0);

    const date = new Date(y, mo - 1, d, hh, mm, ss);
    if (isNaN(date.getTime())) return { date: null, dateKey: "", monthKey: "", dateTimeStr: s };

    return { date, dateKey: toDateKey(date), monthKey: toMonthKey(date), dateTimeStr: formatDateTime(date) };
  }

  function fmtMoney(n, { blankZero = true } = {}) {
    const v = Number(n);
    if (!Number.isFinite(v)) return "";
    if (blankZero && v === 0) return "";
    return v.toLocaleString("ko-KR");
  }

  function toIntLike(v) {
    if (typeof v === "number" && Number.isFinite(v)) return Math.trunc(v);
    const s = normalizeText(v);
    if (!s) return null;
    return /^\d+$/.test(s) ? parseInt(s, 10) : null;
  }

  function firstNonEmpty(row, idxList) {
    for (const i of idxList) {
      const v = row?.[i];
      if (v !== null && v !== undefined && String(v).trim() !== "") return v;
    }
    return "";
  }

  function inferColCountFromRows(rows) {
    let max = 0;
    const limit = Math.min(rows.length, 80);
    for (let i = 0; i < limit; i++) {
      max = Math.max(max, (rows[i] || []).length);
    }
    return max || 14;
  }

  /* =========================================================
     Column label helpers
     ========================================================= */
  function indexToColLetter(idx) {
    let n = idx + 1;
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /* =========================================================
     Header detect / mapping
     ========================================================= */
  function detectHeaderMap(rows) {
    const maxScan = Math.min(rows.length, 40);
    const key = (v) => normalizeText(v).replace(/\s+/g, "");

    const want = {
      seq: /구분/,
      datetime: /거래일자/,
      withdraw: /출금금액|출금/,
      deposit: /입금금액|입금/,
      balance: /거래후잔액|잔액/,
      content: /거래내용/,
      note: /거래기록사항|거래기록/,
      branch: /거래점/,
    };

    for (let r = 0; r < maxScan; r++) {
      const row = rows[r] || [];
      const keys = row.map(key);

      const idxSeq = keys.findIndex((x) => want.seq.test(x));
      const idxDt = keys.findIndex((x) => want.datetime.test(x));
      const idxW = keys.findIndex((x) => want.withdraw.test(x));
      const idxD = keys.findIndex((x) => want.deposit.test(x));
      const idxB = keys.findIndex((x) => want.balance.test(x));
      const idxC = keys.findIndex((x) => want.content.test(x));
      const idxN = keys.findIndex((x) => want.note.test(x));
      const idxBr = keys.findIndex((x) => want.branch.test(x));

      const ok = idxSeq >= 0 && idxDt >= 0 && idxW >= 0 && idxD >= 0 && idxB >= 0 && idxC >= 0 && idxN >= 0;
      if (!ok) continue;

      return {
        headerRowIdx: r,
        start: {
          seq: idxSeq,
          datetime: idxDt,
          withdraw: idxW,
          deposit: idxD,
          balance: idxB,
          content: idxC,
          note: idxN,
          branch: idxBr >= 0 ? idxBr : null,
        },
      };
    }

    return null;
  }

  function extractMeta(rows) {
    const meta = { owner: "", account: "", period: "" };

    const findRightValue = (row, idx) => {
      for (let j = idx + 1; j < row.length; j++) {
        const v = normalizeText(row[j]);
        if (v) return v;
      }
      return "";
    };

    for (let r = 0; r < Math.min(rows.length, 25); r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const v = normalizeText(row[c]);
        if (!v) continue;

        if (v === "계좌번호") meta.account = findRightValue(row, c);
        if (v === "예금주명") meta.owner = findRightValue(row, c);
        if (v === "조회기간") meta.period = findRightValue(row, c);
      }
    }
    return meta;
  }

  function buildMappingFromStart(start, colCount) {
    const seq = start.seq ?? DEFAULT_START.seq;
    const datetime = start.datetime ?? DEFAULT_START.datetime;
    const withdrawStart = start.withdraw ?? DEFAULT_START.withdraw;
    const deposit = start.deposit ?? DEFAULT_START.deposit;
    const balanceStart = start.balance ?? DEFAULT_START.balance;
    const contentStart = start.content ?? DEFAULT_START.content;
    const noteStart = start.note ?? DEFAULT_START.note;

    const branch =
      (start.branch === null || start.branch === undefined)
        ? (DEFAULT_START.branch < colCount ? DEFAULT_START.branch : null)
        : start.branch;

    return {
      seq,
      datetime,
      withdraw: [withdrawStart, withdrawStart + 1],
      deposit,
      balance: [balanceStart, balanceStart + 1],
      content: [contentStart, contentStart + 1],
      note: [noteStart, noteStart + 1, noteStart + 2],
      branch,
    };
  }

  function applyMappingToUI(start, headerRowIdx) {
    if (el.inpHeaderRow) el.inpHeaderRow.value = String(headerRowIdx + 1);

    setSelect(el.selSeq, start.seq);
    setSelect(el.selDatetime, start.datetime);
    setSelect(el.selWithdraw, start.withdraw);
    setSelect(el.selDeposit, start.deposit);
    setSelect(el.selBalance, start.balance);
    setSelect(el.selContent, start.content);
    setSelect(el.selNote, start.note);
    setSelect(el.selBranch, start.branch);

    function setSelect(sel, idx) {
      if (!sel) return;
      if (idx === null || idx === undefined) { sel.value = ""; return; }
      sel.value = String(idx);
    }
  }

  function populateMappingSelectOptions({ colCount, headerRow }) {
    const selects = [
      el.selSeq, el.selDatetime, el.selWithdraw, el.selDeposit, el.selBalance, el.selContent, el.selNote,
    ].filter(Boolean);

    const branchSelect = el.selBranch;

    for (const sel of selects) fillSelect(sel, { includeNone: false });
    if (branchSelect) fillSelect(branchSelect, { includeNone: true });

    function fillSelect(sel, { includeNone }) {
      const prev = sel.value;
      sel.innerHTML = "";

      if (includeNone) {
        const opt = document.createElement("option");
        opt.value = "";
        opt.textContent = "(없음)";
        sel.appendChild(opt);
      }

      for (let i = 0; i < colCount; i++) {
        const letter = indexToColLetter(i);
        const headerText = normalizeText(headerRow?.[i]);
        const opt = document.createElement("option");
        opt.value = String(i);
        opt.textContent = headerText ? `${letter} : ${headerText}` : `${letter}`;
        sel.appendChild(opt);
      }

      if (prev) sel.value = prev;
    }
  }

  function readMappingFromUI() {
    const headerRow1 = Number(el.inpHeaderRow?.value);
    const headerRowIdx = Number.isFinite(headerRow1) ? Math.max(0, Math.round(headerRow1) - 1) : 11;

    const mustInt = (sel, name) => {
      if (!sel) throw new Error(`${name} 선택 박스를 찾을 수 없습니다.`);
      const n = parseInt(sel.value, 10);
      if (!Number.isInteger(n) || n < 0) throw new Error(`${name} 매핑이 올바르지 않습니다.`);
      return n;
    };

    const maybeInt = (sel) => {
      if (!sel) return null;
      if (sel.value === "") return null;
      const n = parseInt(sel.value, 10);
      return Number.isInteger(n) && n >= 0 ? n : null;
    };

    const start = {
      seq: mustInt(el.selSeq, "구분"),
      datetime: mustInt(el.selDatetime, "거래일자"),
      withdraw: mustInt(el.selWithdraw, "출금금액"),
      deposit: mustInt(el.selDeposit, "입금금액"),
      balance: mustInt(el.selBalance, "거래후잔액"),
      content: mustInt(el.selContent, "거래내용"),
      note: mustInt(el.selNote, "거래기록사항"),
      branch: maybeInt(el.selBranch),
    };

    if (state.rows.length && (headerRowIdx < 0 || headerRowIdx >= state.rows.length)) {
      throw new Error(`헤더 행이 범위를 벗어났습니다. (1 ~ ${state.rows.length})`);
    }

    return { headerRowIdx, start };
  }

  /* =========================================================
     Engine
     ========================================================= */
  function parseTransactions(rows, headerRowIdx, map, ctx) {
    const txns = [];
    let skipped = 0;
    let errors = 0;

    const startRow = headerRowIdx + 1;

    for (let r = startRow; r < rows.length; r++) {
      const row = rows[r] || [];

      const seq = toIntLike(row[map.seq]);
      if (seq === null) { skipped++; continue; }

      const dt = parseDateTime(row[map.datetime]);
      if (!dt.dateKey) { errors++; skipped++; continue; }

      const withdrawRaw = firstNonEmpty(row, map.withdraw);
      const depositRaw = firstNonEmpty(row, [map.deposit]);
      const balanceRaw = firstNonEmpty(row, map.balance);
      const contentRaw = firstNonEmpty(row, map.content);
      const noteRaw = firstNonEmpty(row, map.note);
      const branchRaw = map.branch !== null && map.branch !== undefined ? row[map.branch] : "";

      const withdraw = parseAmount(withdrawRaw);
      const deposit = parseAmount(depositRaw);
      const balance = parseAmount(balanceRaw);

      txns.push({
        seq,
        date: dt.date,
        dateKey: dt.dateKey,
        monthKey: dt.monthKey,
        dateTimeStr: dt.dateTimeStr,

        withdraw,
        deposit,
        balance,

        content: normalizeText(contentRaw),
        note: normalizeText(noteRaw),
        branch: normalizeText(branchRaw),

        source: { fileName: ctx?.fileName || "", sheetName: ctx?.sheetName || "", rowIndex: r + 1 },
      });
    }

    return { txns, skippedRows: skipped, parseErrors: errors };
  }

  function aggregate(txns) {
    const dayMap = new Map();
    const monthMap = new Map();

    let totalWithdraw = 0;
    let totalDeposit = 0;

    for (const t of txns) {
      totalWithdraw += t.withdraw;
      totalDeposit += t.deposit;

      if (!dayMap.has(t.dateKey)) {
        dayMap.set(t.dateKey, {
          dateKey: t.dateKey,
          monthKey: t.monthKey,
          txns: [],
          withdrawSum: 0,
          depositSum: 0,
          withdrawCount: 0,
          depositCount: 0,
        });
      }
      const d = dayMap.get(t.dateKey);
      d.txns.push(t);
      d.withdrawSum += t.withdraw;
      d.depositSum += t.deposit;
      if (t.withdraw > 0) d.withdrawCount += 1;
      if (t.deposit > 0) d.depositCount += 1;

      if (!monthMap.has(t.monthKey)) {
        monthMap.set(t.monthKey, { monthKey: t.monthKey, withdrawSum: 0, depositSum: 0, txnCount: 0 });
      }
      const m = monthMap.get(t.monthKey);
      m.withdrawSum += t.withdraw;
      m.depositSum += t.deposit;
      m.txnCount += 1;
    }

    const dailyList = [...dayMap.values()].map((d) => ({
      ...d,
      txnCount: d.txns.length,
      hasDeposit: d.depositSum > 0,
    }));
    dailyList.sort((a, b) => b.dateKey.localeCompare(a.dateKey));

    const monthlyList = [...monthMap.values()].sort((a, b) => a.monthKey.localeCompare(b.monthKey));

    const totals = {
      totalWithdraw,
      totalDeposit,
      txnCount: txns.length,
      depositDayCount: dailyList.filter((d) => d.hasDeposit).length,
    };

    return { dailyList, monthlyList, totals };
  }

  /* =========================================================
     Render
     ========================================================= */
  function derivePeriodFromTxns(txns) {
    if (!txns || !txns.length) return "";
    const keys = txns.map((t) => t.dateKey).filter(Boolean).sort();
    if (!keys.length) return "";
    return `${keys[0]} ~ ${keys[keys.length - 1]}`;
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function renderDashboard() {
    if (el.periodLabel) {
      const p = state.meta.period || derivePeriodFromTxns(state.txns);
      const owner = state.meta.owner ? ` · 예금주: ${state.meta.owner}` : "";
      const acc = state.meta.account ? ` · 계좌: ${state.meta.account}` : "";
      el.periodLabel.textContent = `조회기간: ${p || "-"}${owner}${acc}`;
    }

    if (el.dashboardTotals && state.totals) {
      const t = state.totals;
      el.dashboardTotals.innerHTML = "";

      const items = [
        { cls: "badge badge-deposit", label: "총 입금", value: fmtMoney(t.totalDeposit, { blankZero: false }) + "원" },
        { cls: "badge badge-withdraw", label: "총 출금", value: fmtMoney(t.totalWithdraw, { blankZero: false }) + "원" },
        { cls: "badge badge-muted", label: "거래 건수", value: `${t.txnCount}건` },
        { cls: "badge badge-muted", label: "입금 발생 일수", value: `${t.depositDayCount}일` },
      ];

      for (const it of items) {
        const span = document.createElement("span");
        span.className = it.cls;
        span.textContent = `${it.label}: ${it.value}`;
        el.dashboardTotals.appendChild(span);
      }
    }

    renderMonthlyTable();
    renderDepositDaysTable();
  }

  function renderMonthlyTable() {
    if (!el.monthlyTable) return;
    const rows = state.monthlyList || [];

    el.monthlyTable.innerHTML = `
      <thead>
        <tr>
          <th>월</th>
          <th class="numeric">입금 합계</th>
          <th class="numeric">출금 합계</th>
          <th class="numeric">건수</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(m => `
          <tr>
            <td>${escapeHtml(m.monthKey)}</td>
            <td class="numeric">${fmtMoney(m.depositSum, { blankZero: false })}</td>
            <td class="numeric">${fmtMoney(m.withdrawSum, { blankZero: false })}</td>
            <td class="numeric">${m.txnCount}</td>
          </tr>
        `).join("")}
      </tbody>
    `;
  }

  function renderDepositDaysTable() {
    if (!el.depositDaysTable) return;

    const depositDays = (state.dailyList || [])
      .filter((d) => d.hasDeposit)
      .slice()
      .sort((a, b) => b.dateKey.localeCompare(a.dateKey));

    el.depositDaysTable.innerHTML = `
      <thead>
        <tr>
          <th>날짜</th>
          <th class="numeric">입금 합계</th>
          <th class="numeric">출금 합계</th>
          <th class="numeric">입금 건수</th>
          <th class="numeric">전체 건수</th>
        </tr>
      </thead>
      <tbody>
        ${depositDays.map(d => `
          <tr>
            <td>${escapeHtml(d.dateKey)}</td>
            <td class="numeric">${fmtMoney(d.depositSum, { blankZero: false })}</td>
            <td class="numeric">${fmtMoney(d.withdrawSum, { blankZero: false })}</td>
            <td class="numeric">${d.depositCount}</td>
            <td class="numeric">${d.txnCount}</td>
          </tr>
        `).join("")}
      </tbody>
    `;
  }

  function renderDetails() {
    const order = el.selOrder?.value === "asc" ? "asc" : "desc";
    const mode = el.selMode?.value === "all" ? "all" : "deposit";

    const list = (state.dailyList || []).slice().sort((a, b) => {
      return order === "asc" ? a.dateKey.localeCompare(b.dateKey) : b.dateKey.localeCompare(a.dateKey);
    });

    // panel 1: 날짜 칩
    if (el.dateList) {
      el.dateList.innerHTML = "";
      for (const d of list) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = `date-chip badge ${d.hasDeposit ? "badge-deposit" : "badge-muted"}`;
        btn.textContent = d.dateKey;
        btn.addEventListener("click", () => {
          const target = document.getElementById(`day-${d.dateKey}`);
          if (target) {
            target.scrollIntoView({ behavior: "smooth", block: "start" });
            target.open = true;
          }
        });
        el.dateList.appendChild(btn);
      }
    }

    // panel 2: 일자별 상세
    if (el.detailPanel) {
      el.detailPanel.innerHTML = "";
      for (const d of list) {
        el.detailPanel.appendChild(buildDayBlock(d, { mode }));
      }
    }
  }

  function buildDayBlock(day, { mode }) {
    const details = document.createElement("details");
    details.className = "day-block";
    details.id = `day-${day.dateKey}`;
    details.open = !!day.hasDeposit; // 기본 open 규칙

    const summary = document.createElement("summary");
    summary.className = "day-summary";

    const left = document.createElement("div");
    left.className = "day-title";
    left.textContent = day.dateKey;

    const right = document.createElement("div");
    right.className = "day-badges";

    const b1 = document.createElement("span");
    b1.className = "badge badge-deposit";
    b1.textContent = `입금 ${fmtMoney(day.depositSum, { blankZero: false })}`;

    const b2 = document.createElement("span");
    b2.className = "badge badge-withdraw";
    b2.textContent = `출금 ${fmtMoney(day.withdrawSum, { blankZero: false })}`;

    const b3 = document.createElement("span");
    b3.className = "badge badge-muted";
    b3.textContent = `${day.txnCount}건`;

    right.appendChild(b1);
    right.appendChild(b2);
    right.appendChild(b3);

    summary.appendChild(left);
    summary.appendChild(right);
    details.appendChild(summary);

    const inner = document.createElement("div");
    inner.className = "day-inner";

    let txns = (day.txns || []).slice();
    txns.sort((a, b) => b.dateTimeStr.localeCompare(a.dateTimeStr));
    if (mode === "deposit") txns = txns.filter((t) => t.deposit > 0);

    const wrap = document.createElement("div");
    wrap.className = "table-wrap";

    const table = document.createElement("table");
    table.className = "sheetlike sticky-table";

    table.innerHTML = `
      <thead>
        <tr>
          <th class="col-seq sticky-1">구분</th>
          <th class="col-datetime sticky-2">거래일자</th>
          <th class="numeric">출금금액</th>
          <th class="numeric">입금금액</th>
          <th class="numeric">거래후잔액</th>
          <th>거래내용</th>
          <th>거래기록사항</th>
          <th>거래점</th>
        </tr>
      </thead>
      <tbody></tbody>
      <tfoot>
        <tr>
          <td class="sticky-1" colspan="2"><strong>일합계</strong></td>
          <td class="numeric"><strong>${fmtMoney(day.withdrawSum, { blankZero: false })}</strong></td>
          <td class="numeric"><strong>${fmtMoney(day.depositSum, { blankZero: false })}</strong></td>
          <td colspan="4"></td>
        </tr>
      </tfoot>
    `;

    const tbody = table.querySelector("tbody");
    if (tbody) {
      if (!txns.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td colspan="8" class="muted">${
          mode === "deposit" ? "입금 내역이 없습니다. (현재: 입금만 표시)" : "표시할 내역이 없습니다."
        }</td>`;
        tbody.appendChild(tr);
      } else {
        const frag = document.createDocumentFragment();
        for (const t of txns) {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td class="sticky-1 col-seq">${escapeHtml(t.seq)}</td>
            <td class="sticky-2 col-datetime">${escapeHtml(t.dateTimeStr)}</td>
            <td class="numeric">${fmtMoney(t.withdraw)}</td>
            <td class="numeric">${fmtMoney(t.deposit)}</td>
            <td class="numeric">${fmtMoney(t.balance, { blankZero: false })}</td>
            <td>${escapeHtml(t.content)}</td>
            <td>${escapeHtml(t.note)}</td>
            <td>${escapeHtml(t.branch)}</td>
          `;
          frag.appendChild(tr);
        }
        tbody.appendChild(frag);
      }
    }

    wrap.appendChild(table);
    inner.appendChild(wrap);
    details.appendChild(inner);

    return details;
  }

  /* =========================================================
     Export (xlsx / csv)
     ========================================================= */
  function sanitizeFilename(name) {
    return String(name ?? "").replace(/[\\/:*?"<>|]/g, "_");
  }

  function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1200);
  }

  function buildExportWorkbook() {
    if (!window.XLSX) throw new Error("XLSX 라이브러리가 없습니다.");

    const wb = XLSX.utils.book_new();

    // sheet1: 원본
    if (state.original.ws) {
      XLSX.utils.book_append_sheet(wb, state.original.ws, "통장거래내역 (원본)");
    } else {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["원본 시트를 만들 수 없습니다."]]), "통장거래내역 (원본)");
    }

    const headerAoa = state.original.headerAoa?.length ? state.original.headerAoa : buildFallbackHeaderAoa(state.meta);
    const colCount = state.original.colCount || 14;

    // sheet2: 입출금 + 일합계
    const aoaAll = [];
    for (const row of headerAoa) aoaAll.push(padRow(row, colCount));

    const txnsSorted = state.txns.slice().sort((a, b) => {
      if (a.dateKey !== b.dateKey) return b.dateKey.localeCompare(a.dateKey);
      if (a.dateTimeStr !== b.dateTimeStr) return b.dateTimeStr.localeCompare(a.dateTimeStr);
      return (a.seq ?? 0) - (b.seq ?? 0);
    });

    const byDay = new Map();
    for (const t of txnsSorted) {
      if (!byDay.has(t.dateKey)) byDay.set(t.dateKey, []);
      byDay.get(t.dateKey).push(t);
    }

    const dayKeysAll = [...byDay.keys()].sort((a, b) => b.localeCompare(a));
    for (const dayKey of dayKeysAll) {
      const list = byDay.get(dayKey) || [];
      let dayW = 0;
      let dayD = 0;

      for (const t of list) {
        dayW += t.withdraw;
        dayD += t.deposit;
        aoaAll.push(txnToAoaRow(t, { colCount, depositZeroWhenBlank: true }));
      }

      aoaAll.push(makeSubtotalRow(dayKey, dayW, dayD, { colCount }));
    }

    const wsAll = XLSX.utils.aoa_to_sheet(aoaAll);
    applyHeaderCosmetics(wsAll);
    XLSX.utils.book_append_sheet(wb, wsAll, "통장거래내역 (입출금)");

    // sheet3: 입금만 + 일합계
    const aoaDep = [];
    for (const row of headerAoa) aoaDep.push(padRow(row, colCount));

    const dailyDesc = state.dailyList.slice().sort((a, b) => b.dateKey.localeCompare(a.dateKey));
    for (const day of dailyDesc) {
      const deposits = (day.txns || []).filter((t) => t.deposit > 0).slice();
      if (!deposits.length) continue;

      deposits.sort((a, b) => b.dateTimeStr.localeCompare(a.dateTimeStr));

      let dayW = 0;
      let dayD = 0;
      for (const t of deposits) {
        dayW += t.withdraw;
        dayD += t.deposit;
        aoaDep.push(txnToAoaRow(t, { colCount, depositZeroWhenBlank: false }));
      }

      aoaDep.push(makeSubtotalRow(day.dateKey, dayW, dayD, { colCount, depositOnly: true }));
    }

    const wsDep = XLSX.utils.aoa_to_sheet(aoaDep);
    applyHeaderCosmetics(wsDep);
    XLSX.utils.book_append_sheet(wb, wsDep, "통장거래내역 (입금)");

    return wb;

    function padRow(row, n) {
      return Array.from({ length: n }, (_, i) => row?.[i] ?? "");
    }

    function buildFallbackHeaderAoa(meta) {
      return [
        ["", "입출금거래내역"],
        [""],
        ["", "예금주명", "", "", meta.owner || ""],
        ["", "계좌번호", "", "", meta.account || ""],
        ["", "조회기간", "", "", meta.period || ""],
        [""],
        ["", "구분", "거래일자", "출금금액", "", "입금금액", "거래후잔액", "", "거래내용", "", "거래기록사항", "", "", "거래점"],
      ];
    }

    function txnToAoaRow(t, { colCount, depositZeroWhenBlank }) {
      const row = Array.from({ length: colCount }, () => "");
      row[0] = "";
      row[1] = t.seq ?? "";
      row[2] = t.dateTimeStr || "";
      row[3] = t.withdraw ? t.withdraw : "";
      row[4] = "";
      row[5] = t.deposit ? t.deposit : (depositZeroWhenBlank ? 0 : "");
      row[6] = t.balance ? t.balance : "";
      row[7] = "";
      row[8] = t.content || "";
      row[9] = "";
      row[10] = t.note || "";
      row[11] = "";
      row[12] = "";
      row[13] = t.branch || "";
      return row;
    }

    function makeSubtotalRow(dayKey, withdrawSum, depositSum, { colCount, depositOnly = false }) {
      const row = Array.from({ length: colCount }, () => "");
      row[2] = `${dayKey} 합계`;
      row[3] = depositOnly ? "" : withdrawSum;
      row[5] = depositSum;
      return row;
    }

    function applyHeaderCosmetics(ws) {
      if (state.original.headerMerges?.length) {
        ws["!merges"] = state.original.headerMerges.map((m) => ({ s: m.s, e: m.e }));
      }
      if (state.original.cols) {
        ws["!cols"] = state.original.cols;
      }
    }
  }

  function downloadXlsx() {
    if (!state.txns.length) {
      addLog("warn", "다운로드할 데이터가 없습니다.");
      flushLogs(["다운로드 실패: 데이터 없음"]);
      return;
    }

    try {
      const wb = buildExportWorkbook();
      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const filename = sanitizeFilename(`통장거래내역_집계_${new Date().toISOString().slice(0, 10)}.xlsx`);
      downloadBlob(filename, blob);
    } catch (e) {
      addLog("error", `엑셀 다운로드 실패: ${e.message}`);
      flushLogs(["엑셀 다운로드 실패", e.message]);
    }
  }

  function toCsvRow(values) {
    return values.map((v) => {
      const s = String(v ?? "");
      if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    }).join(",");
  }

  function downloadCsv(kind) {
    if (!state.txns.length) {
      addLog("warn", "다운로드할 데이터가 없습니다.");
      flushLogs(["다운로드 실패: 데이터 없음"]);
      return;
    }

    const rows = [];
    rows.push(toCsvRow(["구분","거래일자","출금금액","입금금액","거래후잔액","거래내용","거래기록사항","거래점"]));

    const data = kind === "deposit" ? state.txns.filter((t) => t.deposit > 0) : state.txns;
    const sorted = data.slice().sort((a, b) => {
      if (a.dateKey !== b.dateKey) return b.dateKey.localeCompare(a.dateKey);
      return b.dateTimeStr.localeCompare(a.dateTimeStr);
    });

    for (const t of sorted) {
      rows.push(toCsvRow([t.seq, t.dateTimeStr, t.withdraw, t.deposit, t.balance, t.content, t.note, t.branch]));
    }

    const csv = "\ufeff" + rows.join("\r\n"); // BOM
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const filename = sanitizeFilename(
      `통장거래내역_${kind === "deposit" ? "입금" : "입출금"}_${new Date().toISOString().slice(0, 10)}.csv`
    );
    downloadBlob(filename, blob);
  }

  /* =========================================================
     File handling
     ========================================================= */
  function setFileMeta(file) {
    if (!el.fileMeta) return;
    if (!file) { el.fileMeta.textContent = ""; return; }
    const mod = new Date(file.lastModified).toLocaleString();
    el.fileMeta.innerHTML = `<b>${escapeHtml(file.name)}</b> · ${bytesToHuman(file.size)} · 수정 ${escapeHtml(mod)}`;
  }

  async function prepareWorkbook(file) {
    if (!file) return;

    clearLogs();
    resetResultsUI();

    setAnalyzeProgress(0, "파일 읽는 중...");

    const ok = await ensureXLSX();
    if (!ok) return;

    try {
      setAnalyzeProgress(10, "엑셀 파싱 중...");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

      state.file = file;
      state.wb = wb;
      state.sheetName = sheetName;
      state.ws = ws;
      state.rows = rows;
      state.colCount = Math.max(14, inferColCountFromRows(rows));

      state.meta = extractMeta(rows);

      // cosmetics for export
      state.original.fileName = file.name;
      state.original.sheetName = sheetName;
      state.original.ws = ws;
      state.original.allMerges = ws["!merges"] || [];
      state.original.cols = ws["!cols"] || null;
      state.original.colCount = state.colCount;

      const detected = detectHeaderMap(rows);
      let headerRowIdx = detected?.headerRowIdx ?? 11;
      let start = detected?.start ?? { ...DEFAULT_START, branch: DEFAULT_START.branch };

      const headerRow = rows[headerRowIdx] || [];
      populateMappingSelectOptions({ colCount: state.colCount, headerRow });
      applyMappingToUI(start, headerRowIdx);

      state.headerRowIdx = headerRowIdx;
      state.mapping = buildMappingFromStart(start, state.colCount);

      state.original.headerRowIdx = headerRowIdx;
      state.original.headerAoa = rows.slice(0, headerRowIdx + 1);
      state.original.headerMerges = state.original.allMerges.filter((m) => (m.e?.r ?? 9999) <= headerRowIdx);

      if (el.mappingStatus) {
        if (detected) {
          el.mappingStatus.textContent = `자동 감지 성공: 헤더 ${headerRowIdx + 1}행 · 기본 매핑이 자동 설정되었습니다.`;
        } else {
          el.mappingStatus.textContent = `자동 감지 실패: 기본값(헤더 12행 + B/C/D/F…)을 적용했습니다. 필요하면 수동 매핑을 조정하세요.`;
          if (el.mappingDetails) el.mappingDetails.open = true;
        }
      }

      setAnalyzeProgress(100, "파일 준비 완료");

      setBtnEnabled(el.btnReset, true);
      setBtnEnabled(el.btnRun, true);
      setBtnEnabled(el.btnApplyMapping, true);

      addLog("info", `파일 준비 완료: ${file.name} (${bytesToHuman(file.size)})`);
      addLog("info", `시트: ${sheetName} · 총 ${rows.length.toLocaleString("ko-KR")}행`);
      flushLogs([
        `파일: ${file.name}`,
        `시트: ${sheetName}`,
        `메타: 예금주=${state.meta.owner || "-"}, 계좌=${state.meta.account || "-"}, 조회기간=${state.meta.period || "-"}`,
        `헤더행: ${headerRowIdx + 1}행 (현재)`,
      ]);
    } catch (e) {
      setAnalyzeProgress(0, "파일 준비 실패");
      addLog("error", `파일 파싱 실패: ${e.message}`);
      flushLogs(["파일 파싱 실패", e.message]);
      throw e;
    }
  }

  function resetAll() {
    clearLogs();

    state.file = null;
    state.wb = null;
    state.sheetName = "";
    state.ws = null;
    state.rows = [];
    state.colCount = 14;
    state.headerRowIdx = 11;
    state.mapping = null;
    state.meta = { owner: "", account: "", period: "" };
    state.txns = [];
    state.dailyList = [];
    state.monthlyList = [];
    state.totals = null;

    state.original = {
      fileName: "",
      sheetName: "",
      ws: null,
      headerRowIdx: null,
      headerAoa: [],
      allMerges: [],
      headerMerges: [],
      cols: null,
      colCount: 14,
    };

    if (el.fileInput) el.fileInput.value = "";
    setFileMeta(null);

    setBtnEnabled(el.btnReset, false);
    setBtnEnabled(el.btnRun, false);
    setBtnEnabled(el.btnApplyMapping, false);

    resetResultsUI();

    setAnalyzeProgress(0, "파일 분석 대기");
  }

  /* =========================================================
     Analysis
     ========================================================= */
  async function runAnalysis() {
    clearLogs();
    resetResultsUI();

    if (!state.file || !state.rows.length) {
      addLog("warn", "파일이 없습니다. 먼저 업로드하세요.");
      flushLogs(["파일이 없습니다."]);
      return;
    }

    const ok = await ensureXLSX();
    if (!ok) return;

    let headerRowIdx;
    let start;
    try {
      ({ headerRowIdx, start } = readMappingFromUI());
    } catch (e) {
      addLog("error", e.message);
      flushLogs(["매핑 오류", e.message]);
      if (el.mappingDetails) el.mappingDetails.open = true;
      return;
    }

    state.headerRowIdx = headerRowIdx;
    state.mapping = buildMappingFromStart(start, state.colCount);

    // export header block sync
    state.original.headerRowIdx = headerRowIdx;
    state.original.headerAoa = state.rows.slice(0, headerRowIdx + 1);
    state.original.headerMerges = (state.original.allMerges || []).filter((m) => (m.e?.r ?? 9999) <= headerRowIdx);

    setAnalyzeProgress(10, "거래내역 파싱 중...");

    const parsed = parseTransactions(state.rows, headerRowIdx, state.mapping, {
      fileName: state.file.name,
      sheetName: state.sheetName,
    });

    state.txns = parsed.txns;

    setAnalyzeProgress(60, "집계 중...");

    const agg = aggregate(state.txns);
    state.dailyList = agg.dailyList;
    state.monthlyList = agg.monthlyList;
    state.totals = agg.totals;

    setAnalyzeProgress(85, "화면 렌더링 중...");

    renderDashboard();
    renderDetails();

    setHidden(el.secDashboard, false);
    setHidden(el.secDetails, false);
    setHidden(el.secDownload, false);

    setBtnEnabled(el.btnDownloadXlsx, true);
    setBtnEnabled(el.btnDownloadCsvAll, true);
    setBtnEnabled(el.btnDownloadCsvDeposit, true);

    setAnalyzeProgress(100, "분석 완료 ✅");

    const period = state.meta.period || derivePeriodFromTxns(state.txns) || "-";
    addLog("info", `거래 파싱: ${parsed.txns.length.toLocaleString("ko-KR")}건`);
    addLog("info", `스킵: ${parsed.skippedRows.toLocaleString("ko-KR")}행`);
    if (parsed.parseErrors) addLog("warn", `파싱 오류(거래일자 등): ${parsed.parseErrors}건`);

    flushLogs([
      `파일: ${state.file.name}`,
      `조회기간: ${period}`,
      `거래 파싱: ${parsed.txns.length.toLocaleString("ko-KR")}건 · 스킵 ${parsed.skippedRows.toLocaleString("ko-KR")}행 · 오류 ${parsed.parseErrors.toLocaleString("ko-KR")}건`,
      `총 입금: ${fmtMoney(state.totals.totalDeposit, { blankZero: false })}원 · 총 출금: ${fmtMoney(state.totals.totalWithdraw, { blankZero: false })}원`,
    ]);
  }

  /* =========================================================
     Events
     ========================================================= */
  function bindEvents() {
    // file pick
    el.btnPick?.addEventListener("click", () => el.fileInput?.click());

    // dropzone
    if (el.dropZone) {
      el.dropZone.addEventListener("click", () => el.fileInput?.click());
      el.dropZone.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          el.fileInput?.click();
        }
      });

      el.dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        // ✅ CSS와 맞춤: .dragover
        el.dropZone.classList.add("dragover");
      });

      el.dropZone.addEventListener("dragleave", () => {
        el.dropZone.classList.remove("dragover");
      });

      el.dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        el.dropZone.classList.remove("dragover");
        const files = e.dataTransfer?.files;
        if (files && files.length) {
          const file = files[0];
          setFileMeta(file);
          prepareWorkbook(file).catch(() => {});
        }
      });
    }

    // input change
    el.fileInput?.addEventListener("change", (e) => {
      const files = e.target.files;
      if (!files || !files.length) return;
      const file = files[0];
      setFileMeta(file);
      prepareWorkbook(file).catch(() => {});
    });

    // reset
    el.btnReset?.addEventListener("click", resetAll);

    // apply mapping
    el.btnApplyMapping?.addEventListener("click", () => {
      if (!state.rows.length) return;
      try {
        const { headerRowIdx, start } = readMappingFromUI();
        state.headerRowIdx = headerRowIdx;
        state.mapping = buildMappingFromStart(start, state.colCount);

        // 헤더행 바뀌었을 수 있으니 옵션 텍스트 갱신
        const headerRow = state.rows[headerRowIdx] || [];
        populateMappingSelectOptions({ colCount: state.colCount, headerRow });
        applyMappingToUI(start, headerRowIdx);

        state.original.headerRowIdx = headerRowIdx;
        state.original.headerAoa = state.rows.slice(0, headerRowIdx + 1);
        state.original.headerMerges = (state.original.allMerges || []).filter((m) => (m.e?.r ?? 9999) <= headerRowIdx);

        if (el.mappingStatus) el.mappingStatus.textContent = `매핑이 적용되었습니다. (헤더 ${headerRowIdx + 1}행)`;

        addLog("info", "수동 매핑 적용 완료");
        flushLogs([
          `매핑 적용 완료`,
          `헤더행: ${headerRowIdx + 1}행`,
          `구분=${indexToColLetter(start.seq)} · 거래일자=${indexToColLetter(start.datetime)} · 출금=${indexToColLetter(start.withdraw)} · 입금=${indexToColLetter(start.deposit)}`,
        ]);
      } catch (e) {
        addLog("error", e.message);
        flushLogs(["매핑 적용 실패", e.message]);
      }
    });

    // header row change → 옵션 라벨 갱신
    el.inpHeaderRow?.addEventListener("change", () => {
      if (!state.rows.length) return;
      const headerRow1 = Number(el.inpHeaderRow.value);
      const headerRowIdx = Number.isFinite(headerRow1) ? Math.max(0, Math.round(headerRow1) - 1) : 11;
      const headerRow = state.rows[headerRowIdx] || [];
      populateMappingSelectOptions({ colCount: state.colCount, headerRow });
    });

    // run
    el.btnRun?.addEventListener("click", () => {
      runAnalysis().catch((e) => {
        addLog("error", e.message);
        flushLogs(["분석 실패", e.message]);
      });
    });

    // detail order/mode
    el.selOrder?.addEventListener("change", () => {
      if (!state.txns.length) return;
      renderDetails();
    });

    el.selMode?.addEventListener("change", () => {
      if (!state.txns.length) return;
      renderDetails();
    });

    // downloads
    el.btnDownloadXlsx?.addEventListener("click", downloadXlsx);
    el.btnDownloadCsvAll?.addEventListener("click", () => downloadCsv("all"));
    el.btnDownloadCsvDeposit?.addEventListener("click", () => downloadCsv("deposit"));

    // to top
    if (el.btnToTop) {
      window.addEventListener("scroll", () => {
        el.btnToTop.style.display = window.scrollY > 500 ? "block" : "none";
      });
      el.btnToTop.addEventListener("click", () => {
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
    }
  }

  /* =========================================================
     Init
     ========================================================= */
  function init() {
    setFileMeta(null);

    setBtnEnabled(el.btnReset, false);
    setBtnEnabled(el.btnRun, false);
    setBtnEnabled(el.btnApplyMapping, false);

    setLoadProgress(0, "대기");
    setAnalyzeProgress(0, "파일 분석 대기");

    resetResultsUI();
    clearLogs();

    bindEvents();
    ensureXLSX(); // 미리 로드 시도
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
