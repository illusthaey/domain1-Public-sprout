(() => {
  "use strict";

  /* =========================================================
     DOM
     ========================================================= */
  const $ = (sel) => document.querySelector(sel);

  const el = {
    // Upload
    dropzone: $("#dropzone"),
    fileInput: $("#fileInput"),
    btnPickFiles: $("#btnPickFiles"),
    btnAnalyze: $("#btnAnalyze"),
    btnClear: $("#btnClear"),
    fileMeta: $("#fileMeta"),
    fileList: $("#fileList"),

    // Progress
    libStatusText: $("#libStatusText"),
    libProgress: $("#libProgress"),
    parseStatusText: $("#parseStatusText"),
    parseProgress: $("#parseProgress"),

    // Options
    optAutoDetect: $("#optAutoDetect"),
    optDedupe: $("#optDedupe"),
    optDepositOnly: $("#optDepositOnly"),
    optShowAllTxns: $("#optShowAllTxns"),
    sortOrderSelect: $("#sortOrderSelect"),

    // Manual mapping
    mappingDetails: $("#mappingDetails"),
    mapSeq: $("#mapSeq"),
    mapDate: $("#mapDate"),
    mapDebit: $("#mapDebit"),
    mapCredit: $("#mapCredit"),
    mapBalance: $("#mapBalance"),
    mapDesc: $("#mapDesc"),
    mapMemo: $("#mapMemo"),
    mapBranch: $("#mapBranch"),
    btnReparse: $("#btnReparse"),

    // Log
    logDetails: $("#logDetails"),
    logOutput: $("#logOutput"),

    // Sections
    dashboardSection: $("#dashboardSection"),
    dailySection: $("#dailySection"),
    downloadSection: $("#downloadSection"),

    // Dashboard
    periodText: $("#periodText"),
    ownerText: $("#ownerText"),
    accountText: $("#accountText"),
    totalCredit: $("#totalCredit"),
    totalDebit: $("#totalDebit"),
    txnCount: $("#txnCount"),
    depositDayCount: $("#depositDayCount"),
    monthlyTbody: $("#monthlyTbody"),
    depositDayChips: $("#depositDayChips"),

    // Daily
    dateListUl: $("#dateListUl"),
    dayGroups: $("#dayGroups"),
    tplDayGroup: $("#tplDayGroup"),

    // Download
    btnDownloadXlsx: $("#btnDownloadXlsx"),

    // ScrollTop
    btnScrollTop: $("#btnScrollTop"),
  };

  /* =========================================================
     State (UI와 엔진이 공유하는 단 하나의 상태)
     ========================================================= */
  const state = {
    files: [],

    // 원본: 첫 파일의 첫 시트를 sheet1에 넣기 위해 보관
    original: {
      fileName: "",
      sheetName: "",
      ws: null,       // SheetJS worksheet object
      headerAoa: [],  // header block(메타~헤더행) AOA
      headerMerges: [],
      cols: null,
      headerRowIdx: null, // 0-based
      colCount: 14,
    },

    // 파싱 결과
    meta: {
      owner: "",
      account: "",
      period: "",
    },

    txns: [],          // 모든 거래(정규화된 객체 배열)
    dayMap: new Map(), // dateKey -> {dateKey, monthKey, txns:[], debitSum, creditSum, ...}
    dailyList: [],     // dayMap을 정렬해 만든 배열
    monthlyList: [],   // 월별 합계 배열

    // 통계/에러
    stats: {
      parsedRows: 0,
      skippedRows: 0,
      parseErrors: 0,
      deduped: 0,
      detectFailed: 0,
    },

    // 마지막 렌더 옵션
    ui: {
      sortOrder: "desc",      // desc | asc
      showAllTxns: false,     // 상세에서 입출금 전체 보기
    },
  };

  /* =========================================================
     Logging & Progress
     ========================================================= */
  function log(line) {
    if (!el.logOutput) return;
    const ts = new Date().toLocaleTimeString();
    el.logOutput.textContent += `[${ts}] ${line}\n`;
    el.logOutput.scrollTop = el.logOutput.scrollHeight;
  }

  function clearLog() {
    if (!el.logOutput) return;
    el.logOutput.textContent = "";
  }

  function setLibProgress(pct, text) {
    if (el.libProgress) el.libProgress.value = clampPct(pct);
    if (el.libStatusText) el.libStatusText.textContent = text || "";
  }

  function setParseProgress(pct, text) {
    if (el.parseProgress) el.parseProgress.value = clampPct(pct);
    if (el.parseStatusText) el.parseStatusText.textContent = text || "";
  }

  function clampPct(n) {
    const v = Number(n);
    if (!Number.isFinite(v)) return 0;
    return Math.max(0, Math.min(100, Math.round(v)));
  }

  /* =========================================================
     Utils (텍스트/숫자/날짜 파서)
     ========================================================= */
  function bytesToHuman(bytes) {
    const b = Number(bytes);
    if (!Number.isFinite(b)) return "-";
    const units = ["B", "KB", "MB", "GB"];
    let v = b, i = 0;
    while (v >= 1024 && i < units.length - 1) { v /= 1024; i++; }
    return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
  }

  // Excel에서 _x00A0_ 같은 이스케이프가 들어오는 케이스 대응
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
    // NBSP(0xA0)도 공백으로 통일
    s = s.replace(/\u00A0/g, " ");
    // 줄바꿈/연속 공백 정리
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

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

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

  // 거래일자: 'YYYY/MM/DD HH:MM:SS' (샘플 기준)
  function parseDateTime(v) {
    if (!v && v !== 0) return { date: null, dateKey: "", monthKey: "", dateTimeStr: "" };

    // Date 객체로 들어오는 케이스
    if (v instanceof Date && !isNaN(v.getTime())) {
      const dateKey = toDateKey(v);
      const monthKey = toMonthKey(v);
      return { date: v, dateKey, monthKey, dateTimeStr: formatDateTime(v) };
    }

    // Excel serial number 케이스(가끔 있음)
    if (typeof v === "number" && Number.isFinite(v) && window.XLSX?.SSF?.parse_date_code) {
      const dc = XLSX.SSF.parse_date_code(v);
      if (dc && dc.y && dc.m && dc.d) {
        const date = new Date(dc.y, dc.m - 1, dc.d, dc.H || 0, dc.M || 0, dc.S || 0);
        const dateKey = toDateKey(date);
        const monthKey = toMonthKey(date);
        return { date, dateKey, monthKey, dateTimeStr: formatDateTime(date) };
      }
    }

    const s = normalizeText(v);

    // YYYY/MM/DD HH:MM:SS 또는 YYYY-MM-DD HH:MM:SS 대응
    const m = s.match(
      /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/
    );
    if (!m) {
      return { date: null, dateKey: "", monthKey: "", dateTimeStr: s };
    }

    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const hh = Number(m[4] ?? 0);
    const mm = Number(m[5] ?? 0);
    const ss = Number(m[6] ?? 0);

    const date = new Date(y, mo - 1, d, hh, mm, ss);
    if (isNaN(date.getTime())) {
      return { date: null, dateKey: "", monthKey: "", dateTimeStr: s };
    }

    return {
      date,
      dateKey: toDateKey(date),
      monthKey: toMonthKey(date),
      dateTimeStr: formatDateTime(date),
    };
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

  function fmtMoney(n, { blankZero = true } = {}) {
    const v = Number(n);
    if (!Number.isFinite(v)) return "";
    if (blankZero && v === 0) return "";
    return v.toLocaleString("ko-KR");
  }

  function firstNonEmpty(row, idxList) {
    for (const i of idxList) {
      const v = row[i];
      if (v !== null && v !== undefined && String(v).trim() !== "") return v;
    }
    return "";
  }

  /* =========================================================
     XLSX loader (로컬 → CDN 폴백)
     ========================================================= */
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = () => resolve(true);
      s.onerror = () => reject(new Error("script load fail: " + src));
      document.head.appendChild(s);
    });
  }

  async function ensureXLSX() {
    if (window.XLSX) return true;

    setLibProgress(10, "라이브러리(XLSX) 불러오는 중...");
    const sources = [
      "/static/xlsx.full.min.js",
      "static/xlsx.full.min.js",
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
    ];

    for (let i = 0; i < sources.length; i++) {
      const src = sources[i];
      try {
        setLibProgress(20 + i * 20, `라이브러리(XLSX) 로드 시도: ${src}`);
        log(`XLSX 로드 시도: ${src}`);
        await loadScript(src);
        if (window.XLSX) {
          setLibProgress(100, "라이브러리 로드 완료 ✅");
          log("XLSX 로드 완료 ✅");
          return true;
        }
      } catch {
        log(`XLSX 로드 실패: ${src}`);
      }
    }

    setLibProgress(0, "라이브러리 로드 실패 ❌ (/static 또는 CDN 확인)");
    return false;
  }

  /* =========================================================
     Engine (파싱/집계/다운로드)  ✅ UI와 분리
     ========================================================= */
  const Engine = {
    // 수동 매핑 입력 파싱 (예: "D,E" -> [3,4])
    parseMappingFromUI() {
      const seq = parseColSingle(el.mapSeq?.value || "B");
      const date = parseColSingle(el.mapDate?.value || "C");
      const debit = parseColList(el.mapDebit?.value || "D,E");
      const credit = parseColSingle(el.mapCredit?.value || "F");
      const balance = parseColList(el.mapBalance?.value || "G,H");
      const desc = parseColList(el.mapDesc?.value || "I,J");
      const memo = parseColList(el.mapMemo?.value || "K,L,M");
      const branch = parseColSingle(el.mapBranch?.value || "N", { allowEmpty: true });

      if (seq === null || date === null || credit === null) {
        throw new Error("수동 매핑 입력이 올바르지 않습니다 (필수: 구분/거래일자/입금).");
      }
      if (!debit.length || !balance.length || !desc.length || !memo.length) {
        throw new Error("수동 매핑 입력이 올바르지 않습니다 (후보 컬럼 목록이 비었습니다).");
      }

      return { seq, date, debit, credit, balance, desc, memo, branch };
    },

    // 자동 헤더 감지: 12행 고정이더라도 “검증” 겸 안전장치
    detectHeaderMap(rows) {
      const maxScan = Math.min(rows.length, 40);

      const key = (v) => normalizeText(v).replace(/\s+/g, "");

      const want = {
        seq: /구분/,
        date: /거래일자/,
        debit: /출금금액/,
        credit: /입금금액/,
        balance: /거래후잔액/,
        desc: /거래내용/,
        memo: /거래기록사항/,
        branch: /거래점/,
      };

      for (let r = 0; r < maxScan; r++) {
        const row = rows[r] || [];
        const keys = row.map(key);

        const idxSeq = keys.findIndex((x) => want.seq.test(x));
        const idxDate = keys.findIndex((x) => want.date.test(x));
        const idxDebit = keys.findIndex((x) => want.debit.test(x));
        const idxCredit = keys.findIndex((x) => want.credit.test(x));
        const idxBalance = keys.findIndex((x) => want.balance.test(x));
        const idxDesc = keys.findIndex((x) => want.desc.test(x));
        const idxMemo = keys.findIndex((x) => want.memo.test(x));
        const idxBranch = keys.findIndex((x) => want.branch.test(x));

        // 핵심 헤더가 한 줄에 모였는지 확인
        const ok =
          idxSeq >= 0 && idxDate >= 0 && idxDebit >= 0 && idxCredit >= 0 &&
          idxBalance >= 0 && idxDesc >= 0 && idxMemo >= 0;

        if (!ok) continue;

        // 병합 열 후보(농협 샘플 기준)
        const debit = makeCandidates(idxDebit, 2);
        const balance = makeCandidates(idxBalance, 2);
        const desc = makeCandidates(idxDesc, 2);
        const memo = makeCandidates(idxMemo, 3);
        const branch = idxBranch >= 0 ? idxBranch : null;

        return {
          headerRowIdx: r,
          map: {
            seq: idxSeq,
            date: idxDate,
            debit,
            credit: idxCredit,
            balance,
            desc,
            memo,
            branch,
          },
        };
      }

      return null;

      function makeCandidates(startIdx, n) {
        const out = [];
        for (let i = 0; i < n; i++) out.push(startIdx + i);
        return out;
      }
    },

    extractMeta(rows) {
      // “계좌번호/예금주명/조회기간” 라벨을 찾고 오른쪽 값을 뽑는다.
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
    },

    parseTransactions(rows, headerRowIdx, map, ctx) {
      // ctx: {fileName, sheetName}
      const txns = [];
      let skipped = 0;
      let errors = 0;

      const start = headerRowIdx + 1;
      for (let r = start; r < rows.length; r++) {
        const row = rows[r] || [];

        const seqRaw = row[map.seq];
        const seq = toIntLike(seqRaw);
        if (seq === null) {
          // 거래행이 아닌 메타/푸터는 통째로 스킵
          skipped++;
          continue;
        }

        const dateRaw = row[map.date];
        const dt = parseDateTime(dateRaw);

        if (!dt.dateKey) {
          errors++;
          skipped++;
          continue;
        }

        const debitRaw = firstNonEmpty(row, map.debit);
        const creditRaw = row[map.credit];
        const balRaw = firstNonEmpty(row, map.balance);
        const descRaw = firstNonEmpty(row, map.desc);
        const memoRaw = firstNonEmpty(row, map.memo);
        const branchRaw = map.branch !== null && map.branch !== undefined ? row[map.branch] : "";

        const debit = parseAmount(debitRaw);
        const credit = parseAmount(creditRaw);
        const balance = parseAmount(balRaw);

        const txn = {
          seq,
          date: dt.date,
          dateKey: dt.dateKey,
          monthKey: dt.monthKey,
          dateTimeStr: dt.dateTimeStr,

          debit,
          credit,
          balance,

          desc: normalizeText(descRaw),
          memo: normalizeText(memoRaw),
          branch: normalizeText(branchRaw),

          source: {
            fileName: ctx.fileName,
            sheetName: ctx.sheetName,
            rowIndex: r + 1, // 1-based for humans
          },
        };

        txns.push(txn);
      }

      return { txns, skippedRows: skipped, parseErrors: errors };
    },

    dedupe(txns) {
      // 기간 겹치는 파일을 여러 개 올렸을 때 중복 거래 제거
      // key: dateTime + debit + credit + balance + desc + memo
      const seen = new Set();
      const out = [];
      let removed = 0;

      for (const t of txns) {
        const key =
          `${t.dateTimeStr}|${t.debit}|${t.credit}|${t.balance}|${t.desc}|${t.memo}|${t.branch}`;

        if (seen.has(key)) {
          removed++;
          continue;
        }
        seen.add(key);
        out.push(t);
      }

      return { txns: out, removed };
    },

    aggregate(txns) {
      const dayMap = new Map();
      const monthMap = new Map();

      let totalDebit = 0;
      let totalCredit = 0;

      for (const t of txns) {
        totalDebit += t.debit;
        totalCredit += t.credit;

        // day
        if (!dayMap.has(t.dateKey)) {
          dayMap.set(t.dateKey, {
            dateKey: t.dateKey,
            monthKey: t.monthKey,
            txns: [],
            debitSum: 0,
            creditSum: 0,
            txnCount: 0,
          });
        }
        const d = dayMap.get(t.dateKey);
        d.txns.push(t);
        d.debitSum += t.debit;
        d.creditSum += t.credit;
        d.txnCount += 1;

        // month
        if (!monthMap.has(t.monthKey)) {
          monthMap.set(t.monthKey, {
            monthKey: t.monthKey,
            debitSum: 0,
            creditSum: 0,
            txnCount: 0,
          });
        }
        const m = monthMap.get(t.monthKey);
        m.debitSum += t.debit;
        m.creditSum += t.credit;
        m.txnCount += 1;
      }

      const dailyList = [...dayMap.values()].map((d) => ({
        ...d,
        hasDeposit: d.creditSum > 0,
        depositCount: d.txns.reduce((acc, x) => acc + (x.credit > 0 ? 1 : 0), 0),
      }));

      // 기본 정렬은 내림차순(최신 날짜 먼저)
      dailyList.sort((a, b) => b.dateKey.localeCompare(a.dateKey));

      const monthlyList = [...monthMap.values()].sort((a, b) =>
        a.monthKey.localeCompare(b.monthKey)
      );

      return {
        dayMap,
        dailyList,
        monthlyList,
        totals: {
          totalDebit,
          totalCredit,
          txnCount: txns.length,
          depositDayCount: dailyList.filter((d) => d.hasDeposit).length,
        },
      };
    },

    buildExportWorkbook({ meta, txns, dailyList, originalWsInfo }) {
      if (!window.XLSX) throw new Error("XLSX 라이브러리가 없습니다.");

      const wb = XLSX.utils.book_new();

      // 1) 원본 (첫 파일의 첫 시트 그대로)
      if (originalWsInfo?.ws) {
        XLSX.utils.book_append_sheet(wb, originalWsInfo.ws, "통장거래내역 (원본)");
      } else {
        const wsEmpty = XLSX.utils.aoa_to_sheet([["원본 시트를 만들 수 없습니다."]]);
        XLSX.utils.book_append_sheet(wb, wsEmpty, "통장거래내역 (원본)");
      }

      // 공통: 헤더 블럭(메타~헤더행) AOA 복사
      const headerAoa = originalWsInfo?.headerAoa?.length
        ? originalWsInfo.headerAoa
        : buildFallbackHeaderAoa(meta);

      const colCount = originalWsInfo?.colCount || 14;

      // 2) 입출금 전체 + 일자별 합계 행
      const aoaAll = [];
      for (const row of headerAoa) aoaAll.push(padRow(row, colCount));

      // 날짜 내림차순으로 그룹 유지(원본 정렬을 최대한 유지하려면 txns를 dateTime 내림차순 정렬)
      const txnsSorted = txns.slice().sort((a, b) => {
        // 1) 날짜내림차순
        if (a.dateKey !== b.dateKey) return b.dateKey.localeCompare(a.dateKey);
        // 2) 시간내림차순(문자열로도 충분히 정렬됨: YYYY/MM/DD HH:MM:SS)
        if (a.dateTimeStr !== b.dateTimeStr) return b.dateTimeStr.localeCompare(a.dateTimeStr);
        // 3) tie-break: seq 오름차순
        return (a.seq ?? 0) - (b.seq ?? 0);
      });

      // dayKey -> rows
      const byDay = new Map();
      for (const t of txnsSorted) {
        if (!byDay.has(t.dateKey)) byDay.set(t.dateKey, []);
        byDay.get(t.dateKey).push(t);
      }

      const dayKeysAll = [...byDay.keys()].sort((a, b) => b.localeCompare(a)); // desc
      for (const dayKey of dayKeysAll) {
        const list = byDay.get(dayKey) || [];
        let dayDebit = 0;
        let dayCredit = 0;

        for (const t of list) {
          dayDebit += t.debit;
          dayCredit += t.credit;
          aoaAll.push(txnToAoaRow(t, { colCount, creditZeroWhenBlank: true }));
        }

        // 합계행(가독성 위해 거래일자 칸에 라벨)
        aoaAll.push(makeSubtotalRow(dayKey, dayDebit, dayCredit, { colCount }));
      }

      const wsAll = XLSX.utils.aoa_to_sheet(aoaAll);
      applyHeaderCosmetics(wsAll, originalWsInfo);
      XLSX.utils.book_append_sheet(wb, wsAll, "통장거래내역 (입출금)");

      // 3) 입금만 + 일자별 합계
      const aoaDep = [];
      for (const row of headerAoa) aoaDep.push(padRow(row, colCount));

      for (const day of dailyList.slice().sort((a, b) => b.dateKey.localeCompare(a.dateKey))) {
        const deposits = (day.txns || []).filter((t) => t.credit > 0);

        let dayDebit = 0;
        let dayCredit = 0;

        // 입금 있는 날만 출력(요구사항)
        if (deposits.length === 0) continue;

        // 시간 내림차순
        deposits.sort((a, b) => b.dateTimeStr.localeCompare(a.dateTimeStr));

        for (const t of deposits) {
          dayDebit += t.debit;
          dayCredit += t.credit;
          aoaDep.push(txnToAoaRow(t, { colCount, creditZeroWhenBlank: false }));
        }

        aoaDep.push(makeSubtotalRow(day.dateKey, dayDebit, dayCredit, { colCount, depositOnly: true }));
      }

      const wsDep = XLSX.utils.aoa_to_sheet(aoaDep);
      applyHeaderCosmetics(wsDep, originalWsInfo);
      XLSX.utils.book_append_sheet(wb, wsDep, "통장거래내역 (입금)");

      return wb;

      /* ----- helpers for export ----- */

      function padRow(row, n) {
        const out = Array.from({ length: n }, (_, i) => (row?.[i] ?? ""));
        return out;
      }

      function buildFallbackHeaderAoa(meta) {
        // 원본 헤더를 못 가져왔을 때 최소한의 헤더만 만든다.
        // (가능하면 원본 headerAoa를 쓰는 게 더 예쁨)
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

      function txnToAoaRow(t, { colCount, creditZeroWhenBlank }) {
        // 농협 샘플 구조(14열): A빈칸, B구분, C거래일자, D출금, E(병합), F입금, G잔액, H(병합), I내용, J(병합), K기록, L/M(병합), N거래점
        const row = Array.from({ length: colCount }, () => "");
        row[0] = ""; // A
        row[1] = t.seq ?? "";
        row[2] = t.dateTimeStr || "";
        row[3] = t.debit ? t.debit : "";
        row[4] = ""; // merged
        // 입금칸: 빈값을 0으로 넣을지 여부(샘플은 debit행 credit=0)
        row[5] = t.credit ? t.credit : (creditZeroWhenBlank ? 0 : "");
        row[6] = t.balance ? t.balance : "";
        row[7] = "";
        row[8] = t.desc || "";
        row[9] = "";
        row[10] = t.memo || "";
        row[11] = "";
        row[12] = "";
        row[13] = t.branch || "";
        return row;
      }

      function makeSubtotalRow(dayKey, debitSum, creditSum, { colCount, depositOnly = false }) {
        const row = Array.from({ length: colCount }, () => "");
        // B(구분) 비움, C에 라벨
        row[2] = `${dayKey} 합계`;
        row[3] = depositOnly ? "" : debitSum;  // 입금 시트는 출금합계 굳이 안 보여줘도 됨
        row[5] = creditSum;
        return row;
      }

      function applyHeaderCosmetics(ws, originalWsInfo) {
        // 원본의 헤더부 병합/열너비를 “가능한 선”에서 복사
        if (originalWsInfo?.headerMerges?.length) {
          ws["!merges"] = originalWsInfo.headerMerges.map((m) => ({ s: m.s, e: m.e }));
        }
        if (originalWsInfo?.cols) {
          ws["!cols"] = originalWsInfo.cols;
        }
      }
    },
  };

  function parseColSingle(col, { allowEmpty = false } = {}) {
    const s = normalizeText(col).toUpperCase();
    if (!s && allowEmpty) return null;
    if (!s) return null;
    if (!/^[A-Z]{1,3}$/.test(s)) return null;
    return colToIndex(s);
  }

  function parseColList(spec) {
    const s = normalizeText(spec).toUpperCase();
    if (!s) return [];
    return s.split(",").map(x => x.trim()).filter(Boolean)
      .map(x => colToIndex(x))
      .filter(x => Number.isInteger(x) && x >= 0);
  }

  function colToIndex(col) {
    let n = 0;
    for (let i = 0; i < col.length; i++) {
      const code = col.charCodeAt(i);
      if (code < 65 || code > 90) return null;
      n = n * 26 + (code - 64);
    }
    return n - 1;
  }

  function toIntLike(v) {
    if (typeof v === "number" && Number.isFinite(v)) return Math.trunc(v);
    const s = normalizeText(v);
    if (!s) return null;
    return /^\d+$/.test(s) ? parseInt(s, 10) : null;
  }

  /* =========================================================
     UI Rendering  ✅ Engine과 분리
     ========================================================= */
  const UI = {
    hideResults() {
      if (el.dashboardSection) el.dashboardSection.hidden = true;
      if (el.dailySection) el.dailySection.hidden = true;
      if (el.downloadSection) el.downloadSection.hidden = true;
    },

    showResults() {
      if (el.dashboardSection) el.dashboardSection.hidden = false;
      if (el.dailySection) el.dailySection.hidden = false;
      if (el.downloadSection) el.downloadSection.hidden = false;
    },

    renderFileList() {
      if (!el.fileMeta || !el.fileList) return;

      if (!state.files.length) {
        el.fileMeta.textContent = "아직 파일이 없습니다.";
        el.fileList.innerHTML = "";
        return;
      }

      const totalBytes = state.files.reduce((a, f) => a + (f.size || 0), 0);
      el.fileMeta.innerHTML = `파일 <b>${state.files.length}</b>개 · 총 ${bytesToHuman(totalBytes)}`;

      el.fileList.innerHTML = state.files.map((f) => {
        const mod = new Date(f.lastModified).toLocaleString();
        return `<div>• <b>${escapeHtml(f.name)}</b><br/><span class="muted">${bytesToHuman(f.size)} · 수정 ${escapeHtml(mod)}</span></div>`;
      }).join("<div style='height:8px'></div>");
    },

    renderDashboard() {
      const { owner, account, period } = state.meta;
      const totals = state._totals;

      if (el.ownerText) el.ownerText.textContent = owner || "-";
      if (el.accountText) el.accountText.textContent = account || "-";
      if (el.periodText) el.periodText.textContent = period || derivePeriodText();

      if (el.totalCredit) el.totalCredit.textContent = fmtMoney(totals.totalCredit, { blankZero: false });
      if (el.totalDebit) el.totalDebit.textContent = fmtMoney(totals.totalDebit, { blankZero: false });
      if (el.txnCount) el.txnCount.textContent = String(totals.txnCount || 0);
      if (el.depositDayCount) el.depositDayCount.textContent = String(totals.depositDayCount || 0);

      UI.renderMonthlyTable();
      UI.renderDepositDayChips();
    },

    renderMonthlyTable() {
      if (!el.monthlyTbody) return;
      el.monthlyTbody.innerHTML = "";

      const frag = document.createDocumentFragment();
      for (const m of state.monthlyList) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${escapeHtml(m.monthKey)}</td>
          <td class="numeric">${fmtMoney(m.creditSum, { blankZero: false })}</td>
          <td class="numeric">${fmtMoney(m.debitSum, { blankZero: false })}</td>
          <td class="numeric">${String(m.txnCount || 0)}</td>
        `;
        frag.appendChild(tr);
      }
      el.monthlyTbody.appendChild(frag);
    },

    renderDepositDayChips() {
      if (!el.depositDayChips) return;
      el.depositDayChips.innerHTML = "";

      const depositDays = state.dailyList
        .filter((d) => d.hasDeposit)
        .slice()
        .sort((a, b) => b.dateKey.localeCompare(a.dateKey)); // desc

      const frag = document.createDocumentFragment();
      for (const d of depositDays) {
        const span = document.createElement("span");
        span.className = "chip is-deposit";
        span.textContent = d.dateKey;
        span.title = "클릭하면 해당 날짜로 이동";
        span.addEventListener("click", () => {
          const target = document.getElementById(`day-${d.dateKey}`);
          if (target) {
            target.scrollIntoView({ behavior: "smooth", block: "start" });
            target.open = true;
          } else {
            // 아직 렌더 안됐을 때 대비
            document.getElementById("dailySection")?.scrollIntoView({ behavior: "smooth" });
          }
        });
        frag.appendChild(span);
      }
      el.depositDayChips.appendChild(frag);
    },

    renderDaily() {
      if (!el.dateListUl || !el.dayGroups || !el.tplDayGroup) return;

      // 정렬
      const order = state.ui.sortOrder;
      const list = state.dailyList.slice().sort((a, b) => {
        return order === "asc"
          ? a.dateKey.localeCompare(b.dateKey)
          : b.dateKey.localeCompare(a.dateKey);
      });

      // 패널1: 날짜 목록
      el.dateListUl.innerHTML = "";
      const ulFrag = document.createDocumentFragment();

      for (const d of list) {
        const li = document.createElement("li");

        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "chip" + (d.hasDeposit ? " is-deposit" : "");
        btn.textContent = d.dateKey;
        btn.addEventListener("click", () => {
          const target = document.getElementById(`day-${d.dateKey}`);
          if (target) {
            target.scrollIntoView({ behavior: "smooth", block: "start" });
            // 클릭했으면 펼쳐주는 게 UX가 좋음
            target.open = true;
          }
        });

        li.appendChild(btn);
        ulFrag.appendChild(li);
      }

      el.dateListUl.appendChild(ulFrag);

      // 패널2: 날짜별 details 그룹
      el.dayGroups.innerHTML = "";
      const groupFrag = document.createDocumentFragment();

      for (const d of list) {
        const details = UI.renderDayGroup(d);
        groupFrag.appendChild(details);
      }

      el.dayGroups.appendChild(groupFrag);
    },

    renderDayGroup(dayAgg) {
      const tpl = el.tplDayGroup;
      const node = tpl.content.firstElementChild.cloneNode(true);

      node.dataset.date = dayAgg.dateKey;
      node.id = `day-${dayAgg.dateKey}`;

      // 기본 open 규칙:
      // - 입금 있는 날: open
      // - 출금만 있는 날: closed
      node.open = !!dayAgg.hasDeposit;

      const label = node.querySelector(".day-label");
      const meta = node.querySelector(".day-meta");
      if (label) label.textContent = dayAgg.dateKey;
      if (meta) meta.textContent =
        `입금 ${fmtMoney(dayAgg.creditSum, { blankZero:false })} / 출금 ${fmtMoney(dayAgg.debitSum, { blankZero:false })} / ${dayAgg.txnCount}건`;

      // 표 채우기
      const tbody = node.querySelector("tbody");
      const tfoot = node.querySelector("tfoot");
      const showAll = !!state.ui.showAllTxns;
      const depositOnly = !!(el.optDepositOnly?.checked) && !showAll;

      // 날짜에 속한 거래들
      let txns = (dayAgg.txns || []).slice();

      // 시간 정렬(기본: 내림차순)
      txns.sort((a, b) => b.dateTimeStr.localeCompare(a.dateTimeStr));

      // 입금만 필터(기본)
      if (depositOnly) {
        txns = txns.filter((t) => t.credit > 0);
      }

      // tbody 구성
      if (tbody) {
        tbody.innerHTML = "";

        if (!txns.length) {
          const tr = document.createElement("tr");
          tr.innerHTML = `<td colspan="7" class="muted">표시할 내역이 없습니다.</td>`;
          tbody.appendChild(tr);
        } else {
          const frag = document.createDocumentFragment();
          for (const t of txns) {
            const tr = document.createElement("tr");
            tr.innerHTML = `
              <td>${escapeHtml(t.dateTimeStr || "")}</td>
              <td class="numeric">${fmtMoney(t.debit)}</td>
              <td class="numeric">${fmtMoney(t.credit)}</td>
              <td class="numeric">${fmtMoney(t.balance, { blankZero:false })}</td>
              <td>${escapeHtml(t.desc)}</td>
              <td>${escapeHtml(t.memo)}</td>
              <td>${escapeHtml(t.branch)}</td>
            `;
            frag.appendChild(tr);
          }
          tbody.appendChild(frag);
        }
      }

      // tfoot 합계(표시된 필터 기준 합계)
      if (tfoot) {
        const debitSum = txns.reduce((a, x) => a + (x.debit || 0), 0);
        const creditSum = txns.reduce((a, x) => a + (x.credit || 0), 0);

        const tds = tfoot.querySelectorAll("td");
        // [0]=합계, [1]=출금, [2]=입금
        if (tds[1]) tds[1].textContent = fmtMoney(debitSum, { blankZero:false });
        if (tds[2]) tds[2].textContent = fmtMoney(creditSum, { blankZero:false });
      }

      return node;
    },
  };

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function derivePeriodText() {
    // meta.period가 없으면 txns의 min/max로 추정
    if (!state.txns.length) return "-";
    const dates = state.txns.map(t => t.dateKey).filter(Boolean).sort();
    if (!dates.length) return "-";
    const min = dates[0];
    const max = dates[dates.length - 1];
    return `${min} ~ ${max}`;
  }

  /* =========================================================
     App pipeline (업로드 → 분석 → 렌더 → 다운로드)
     ========================================================= */
  async function analyze() {
    clearLog();

    UI.hideResults();
    setParseProgress(0, "파일 분석 대기");

    if (!state.files.length) {
      log("파일이 없습니다. 먼저 업로드하세요.");
      return;
    }

    const ok = await ensureXLSX();
    if (!ok) {
      log("XLSX 로드 실패 → 분석 중단");
      return;
    }

    // 상태 초기화
    state.txns = [];
    state.dayMap = new Map();
    state.dailyList = [];
    state.monthlyList = [];
    state.stats = { parsedRows: 0, skippedRows: 0, parseErrors: 0, deduped: 0, detectFailed: 0 };
    state.meta = { owner: "", account: "", period: "" };
    state.original = {
      fileName: "",
      sheetName: "",
      ws: null,
      headerAoa: [],
      headerMerges: [],
      cols: null,
      headerRowIdx: null,
      colCount: 14,
    };

    setParseProgress(3, `파일 ${state.files.length}개 읽는 중...`);
    log(`분석 시작: 파일 ${state.files.length}개`);

    const optAuto = !!el.optAutoDetect?.checked;
    const optDedupe = !!el.optDedupe?.checked;

    // 1) 모든 파일 읽어서 rows 확보 + (첫 파일) 원본 ws/헤더블럭 저장
    const fileSheets = []; // [{file, fileName, wb, sheetName, ws, rows}]
    let totalRowCount = 0;

    for (let i = 0; i < state.files.length; i++) {
      const f = state.files[i];
      log(`파일 읽기: ${f.name}`);

      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

      totalRowCount += rows.length;
      fileSheets.push({ file: f, fileName: f.name, wb, sheetName, ws, rows });

      // 첫 파일을 원본으로 보관(요구서 sheet1)
      if (i === 0) {
        state.original.fileName = f.name;
        state.original.sheetName = sheetName;
        state.original.ws = ws;

        // 헤더 감지(첫 파일 기준으로 meta/header block 추출)
        const detect = Engine.detectHeaderMap(rows);
        const headerRowIdx = detect?.headerRowIdx ?? 11; // fallback: 12행(0-based 11)
        state.original.headerRowIdx = headerRowIdx;

        // 헤더 블럭: 0 ~ headerRowIdx(포함)
        state.original.headerAoa = rows.slice(0, headerRowIdx + 1);

        // merges/cols는 원본 ws에서 복사(헤더부만)
        const merges = ws["!merges"] || [];
        state.original.headerMerges = merges.filter(m => (m.e?.r ?? 9999) <= headerRowIdx);
        state.original.cols = ws["!cols"] || null;
        state.original.colCount = Math.max(14, inferColCountFromRows(rows));
      }
    }

    setParseProgress(10, `파싱 시작 (총 ${totalRowCount}행 스캔)...`);

    // 2) 파일별 파싱
    const allTxns = [];
    let processedRows = 0;

    // 수동 매핑은 미리 읽어둔다(자동 실패 시 fallback)
    let manualMap = null;
    try {
      manualMap = Engine.parseMappingFromUI();
    } catch (e) {
      // 수동 매핑이 이상해도 자동 감지가 성공하면 괜찮으므로 즉시 중단하지 않음
      log(`수동 매핑 경고: ${e.message}`);
    }

    for (const fs of fileSheets) {
      const { rows, fileName, sheetName } = fs;

      // meta는 첫 파일 기준으로 채우되, 비어 있으면 뒤 파일에서 보완
      const meta = Engine.extractMeta(rows);
      if (!state.meta.owner && meta.owner) state.meta.owner = meta.owner;
      if (!state.meta.account && meta.account) state.meta.account = meta.account;
      if (!state.meta.period && meta.period) state.meta.period = meta.period;

      // 헤더/맵 결정
      let headerRowIdx = 11;
      let map = null;

      if (optAuto) {
        const detected = Engine.detectHeaderMap(rows);
        if (detected) {
          headerRowIdx = detected.headerRowIdx;
          map = detected.map;
          log(`[${fileName}] 헤더 자동 감지 OK: headerRow=${headerRowIdx + 1}행`);
        } else {
          state.stats.detectFailed++;
          log(`[${fileName}] 헤더 자동 감지 실패 → 수동 매핑 사용 시도`);
        }
      }

      if (!map) {
        if (!manualMap) {
          throw new Error("자동 감지 실패 + 수동 매핑도 유효하지 않아서 파싱할 수 없습니다.");
        }
        // 수동 매핑은 headerRowIdx를 “농협 기본(12행)”으로 가정(요구서에 서식 일정)
        headerRowIdx = 11;
        map = manualMap;
      }

      const parsed = Engine.parseTransactions(rows, headerRowIdx, map, { fileName, sheetName });
      allTxns.push(...parsed.txns);

      state.stats.skippedRows += parsed.skippedRows;
      state.stats.parseErrors += parsed.parseErrors;
      state.stats.parsedRows += parsed.txns.length;

      processedRows += rows.length;
      const pct = 10 + (processedRows / Math.max(1, totalRowCount)) * 75;
      setParseProgress(pct, `파일 분석 중... (${Math.round(pct)}%)`);
    }

    // 3) 중복 제거
    let finalTxns = allTxns;
    if (optDedupe) {
      const d = Engine.dedupe(allTxns);
      finalTxns = d.txns;
      state.stats.deduped = d.removed;
      log(`중복 제거: ${d.removed}건 제거됨`);
    }

    // 4) 집계
    const agg = Engine.aggregate(finalTxns);
    state.txns = finalTxns;
    state.dayMap = agg.dayMap;
    state.dailyList = agg.dailyList;
    state.monthlyList = agg.monthlyList;
    state._totals = agg.totals;

    // dailyList 각 항목에 txns 붙여주기(렌더 편의)
    // (Engine.aggregate에서 txns를 이미 갖고 있지만, 여기서 안전하게 동기화)
    for (const d of state.dailyList) {
      const mapItem = state.dayMap.get(d.dateKey);
      d.txns = mapItem?.txns || [];
    }

    // 5) 렌더
    UI.showResults();
    UI.renderDashboard();
    UI.renderDaily();

    setParseProgress(100, "분석 완료 ✅");
    log(`분석 완료 ✅`);
    log(`- 거래행 파싱: ${state.stats.parsedRows}건`);
    log(`- 스킵된 행: ${state.stats.skippedRows}행`);
    log(`- 파싱 에러(거래일자 등): ${state.stats.parseErrors}건`);
    log(`- 감지 실패: ${state.stats.detectFailed}회`);
    log(`- 중복 제거: ${state.stats.deduped}건`);
  }

  function inferColCountFromRows(rows) {
    let max = 0;
    for (let i = 0; i < Math.min(rows.length, 80); i++) {
      max = Math.max(max, (rows[i] || []).length);
    }
    return max || 14;
  }

  /* =========================================================
     Download
     ========================================================= */
  function downloadXlsx() {
    if (!state.txns.length) {
      log("다운로드할 데이터가 없습니다.");
      return;
    }
    try {
      const wb = Engine.buildExportWorkbook({
        meta: state.meta,
        txns: state.txns,
        dailyList: state.dailyList,
        originalWsInfo: state.original,
      });

      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      const safe = sanitizeFilename(`통장거래내역_집계_${new Date().toISOString().slice(0, 10)}.xlsx`);
      a.download = safe;
      document.body.appendChild(a);
      a.click();
      a.remove();

      setTimeout(() => URL.revokeObjectURL(url), 1200);
      log(`엑셀 다운로드: ${safe}`);
    } catch (e) {
      console.error(e);
      log(`다운로드 실패: ${e.message}`);
    }
  }

  function sanitizeFilename(name) {
    return String(name ?? "").replace(/[\\/:*?"<>|]/g, "_");
  }

  /* =========================================================
     File handling (추가 업로드 지원)
     ========================================================= */
  function addFiles(files) {
    const list = Array.from(files || []).filter(Boolean);
    if (!list.length) return;

    // 기존 + 신규 합치고(같은 파일은 중복 방지)
    const byKey = new Map();
    for (const f of state.files) byKey.set(fileKey(f), f);
    for (const f of list) byKey.set(fileKey(f), f);

    state.files = [...byKey.values()];
    UI.renderFileList();
    log(`파일 추가됨: 현재 ${state.files.length}개`);
  }

  function fileKey(f) {
    return `${f.name}||${f.size}||${f.lastModified}`;
  }

  function resetAll() {
    state.files = [];
    state.txns = [];
    state.dayMap = new Map();
    state.dailyList = [];
    state.monthlyList = [];
    state.meta = { owner: "", account: "", period: "" };
    state.stats = { parsedRows: 0, skippedRows: 0, parseErrors: 0, deduped: 0, detectFailed: 0 };
    state.original = {
      fileName: "",
      sheetName: "",
      ws: null,
      headerAoa: [],
      headerMerges: [],
      cols: null,
      headerRowIdx: null,
      colCount: 14,
    };

    if (el.fileInput) el.fileInput.value = "";
    UI.renderFileList();
    UI.hideResults();

    setParseProgress(0, "파일 분석 대기");
    clearLog();
    log("초기화 완료");
  }

  /* =========================================================
     Events
     ========================================================= */
  function bindEvents() {
    // file input
    el.fileInput?.addEventListener("change", (e) => {
      const files = e.target.files;
      addFiles(files);
      // input은 같은 파일을 다시 선택할 수도 있게 value를 비워도 되는데,
      // 여기서는 “추가 업로드” UX를 위해 비우지 않고 유지.
    });

    // dropzone
    if (el.dropzone) {
      el.dropzone.addEventListener("click", () => el.fileInput?.click());

      el.dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        el.dropzone.classList.add("is-dragover");
      });

      el.dropzone.addEventListener("dragleave", () => {
        el.dropzone.classList.remove("is-dragover");
      });

      el.dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        el.dropzone.classList.remove("is-dragover");
        addFiles(e.dataTransfer.files);
      });
    }

    // analyze / clear
    el.btnAnalyze?.addEventListener("click", () => {
      analyze().catch((e) => {
        console.error(e);
        log(`분석 실패: ${e.message}`);
      });
    });

    el.btnReparse?.addEventListener("click", () => {
      analyze().catch((e) => {
        console.error(e);
        log(`재분석 실패: ${e.message}`);
      });
    });

    el.btnClear?.addEventListener("click", resetAll);

    // sort / showAll
    el.sortOrderSelect?.addEventListener("change", () => {
      state.ui.sortOrder = el.sortOrderSelect.value === "asc" ? "asc" : "desc";
      UI.renderDaily();
    });

    el.optShowAllTxns?.addEventListener("change", () => {
      state.ui.showAllTxns = !!el.optShowAllTxns.checked;
      UI.renderDaily();
    });

    // deposit-only toggle(상세 필터)
    el.optDepositOnly?.addEventListener("change", () => {
      UI.renderDaily();
    });

    // download
    el.btnDownloadXlsx?.addEventListener("click", downloadXlsx);

    // scroll top button
    window.addEventListener("scroll", () => {
      const show = window.scrollY > 500;
      if (el.btnScrollTop) {
        el.btnScrollTop.classList.toggle("is-visible", show);
      }
    });

    el.btnScrollTop?.addEventListener("click", () => {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

  /* =========================================================
     Init
     ========================================================= */
  async function init() {
    UI.hideResults();
    UI.renderFileList();

    setLibProgress(0, "라이브러리(XLSX) 로드 대기");
    setParseProgress(0, "파일 분석 대기");

    bindEvents();

    // 페이지 로드 시점에 미리 XLSX를 당겨서 “초반 %”를 보여줄 수도 있음
    const ok = await ensureXLSX();
    if (!ok) {
      // 오프라인/폐쇄망이면 여기서 실패할 수 있음 → 로그로 안내
      log("⚠️ XLSX 로드 실패: 폐쇄망이면 /static/xlsx.full.min.js를 준비하세요.");
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
