// script.js
(() => {
  'use strict';

  /* =========================================================
     DOM helpers
     ========================================================= */
  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

  const nextFrame = () =>
    new Promise((resolve) => (window.requestAnimationFrame ? requestAnimationFrame(resolve) : setTimeout(resolve, 16)));

  /* =========================================================
     Engine (파싱/집계/다운로드)
     - UI와 독립적으로 동작하도록 순수 함수 중심
     ========================================================= */
  const Engine = (() => {
    // --- 상수/스키마 ---
    const FIELD_LABELS = Object.freeze({
      headerRow: '헤더 행',
      seq: '구분',
      datetime: '거래일자',
      withdraw: '출금금액',
      deposit: '입금금액',
      balance: '거래후잔액',
      content: '거래내용',
      note: '거래기록사항',
      branch: '거래점',
    });

    const REQUIRED_FIELDS = Object.freeze(['datetime', 'withdraw', 'deposit', 'balance', 'content', 'note']);

    // 농협 원본의 사실상 고정 매핑(0-based)
    const DEFAULT_MAPPING = Object.freeze({
      seq: 1, // B
      datetime: 2, // C
      withdraw: 3, // D (+E merge)
      deposit: 5, // F
      balance: 6, // G (+H merge)
      content: 8, // I (+J merge)
      note: 10, // K (+L+M merge)
      branch: 13, // N
    });

    const DEFAULT_HEADER_ROW_INDEX = 11; // 12행(1-based) → 0-based 11

    // === 문자열 디코딩/정규화 ===
    function decodeExcelEscapes(value) {
      if (value == null) return value;
      if (typeof value !== 'string') return value;
      // Excel export에서 종종 보이는 _x00A0_ 같은 escape 처리
      const out = value.replace(/_x([0-9A-Fa-f]{4})_/g, (m, hex) => {
        const code = Number.parseInt(hex, 16);
        if (Number.isNaN(code)) return m;
        return String.fromCharCode(code);
      });
      // NBSP → 일반 공백
      return out.replace(/\u00A0/g, ' ');
    }

    function normalizeHeader(value) {
      const s = decodeExcelEscapes(value);
      if (s == null) return '';
      return String(s)
        .replace(/\s+/g, '') // 공백/줄바꿈 제거
        .replace(/[()［］\[\]{}<>]/g, '')
        .replace(/[:：]/g, '')
        .trim();
    }

    function isEmptyCell(v) {
      return v == null || (typeof v === 'string' && v.trim() === '');
    }

    function padRow(row, len) {
      const out = new Array(len).fill(null);
      if (!Array.isArray(row)) return out;
      for (let i = 0; i < len; i++) out[i] = i < row.length ? row[i] : null;
      return out;
    }

    // === Sheet → Matrix (AOA) ===
    function sheetToMatrix(ws) {
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      return Array.isArray(matrix) ? matrix : [];
    }

    function getMaxCols(ws) {
      const ref = ws && ws['!ref'];
      if (!ref) return 0;
      const range = XLSX.utils.decode_range(ref);
      return range.e.c + 1;
    }

    function fillMerges(matrix, merges) {
      if (!Array.isArray(merges)) return;
      for (const m of merges) {
        const sr = m.s.r,
          sc = m.s.c,
          er = m.e.r,
          ec = m.e.c;
        if (!matrix[sr]) continue;
        const v = matrix[sr][sc];
        for (let r = sr; r <= er; r++) {
          if (!matrix[r]) continue;
          for (let c = sc; c <= ec; c++) {
            if (isEmptyCell(matrix[r][c])) matrix[r][c] = v;
          }
        }
      }
    }

    // === 헤더/매핑 자동 감지 ===
    function detectHeaderRow(matrix, maxScanRows = 60) {
      const candidates = [
        '거래일자',
        '출금금액',
        '입금금액',
        '거래후잔액',
        '거래내용',
        '거래기록사항',
      ].map((s) => normalizeHeader(s));

      let bestRow = null;
      let bestScore = -1;

      const limit = Math.min(matrix.length, maxScanRows);
      for (let r = 0; r < limit; r++) {
        const row = matrix[r] || [];
        const norms = row.map(normalizeHeader);
        let score = 0;
        for (const key of candidates) {
          if (norms.some((cell) => cell.includes(key))) score++;
        }
        if (score > bestScore) {
          bestScore = score;
          bestRow = r;
        }
      }
      // 최소 3개 이상 헤더가 맞아야 "헤더 행"으로 인정
      if (bestScore >= 3) return { rowIndex: bestRow, score: bestScore };
      return { rowIndex: null, score: bestScore };
    }

    function findColByKeywords(headerRow, keywords) {
      if (!Array.isArray(headerRow)) return null;
      for (let c = 0; c < headerRow.length; c++) {
        const cell = normalizeHeader(headerRow[c]);
        if (!cell) continue;
        for (const kw of keywords) {
          if (cell.includes(kw)) return c;
        }
      }
      return null;
    }

    function autoDetectMapping(headerRow, maxCols) {
      const kw = {
        seq: ['구분'],
        datetime: ['거래일자'],
        withdraw: ['출금금액', '출금'],
        deposit: ['입금금액', '입금'],
        balance: ['거래후잔액', '잔액'],
        content: ['거래내용', '적요', '내용'],
        note: ['거래기록사항', '기록사항', '거래기록'],
        branch: ['거래점', '거래점정보', '거래점정보'],
      };

      const mapping = {};
      for (const [field, keywords] of Object.entries(kw)) {
        const col = findColByKeywords(headerRow, keywords.map((s) => normalizeHeader(s)));
        if (col != null) mapping[field] = col;
      }

      // 누락된 필드는 NH 기본 매핑으로 보정(단, 범위 안이면)
      for (const [field, defCol] of Object.entries(DEFAULT_MAPPING)) {
        if (mapping[field] == null && typeof defCol === 'number' && defCol < maxCols) {
          mapping[field] = defCol;
        }
      }

      const missing = REQUIRED_FIELDS.filter((f) => mapping[f] == null);
      return { mapping, missing };
    }

    // === 메타 정보(조회기간 등) 추출 ===
    function findValueRightOfLabel(matrix, labelNorm, maxRows) {
      const limit = Math.min(matrix.length, maxRows);
      for (let r = 0; r < limit; r++) {
        const row = matrix[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellNorm = normalizeHeader(row[c]);
          if (cellNorm === labelNorm) {
            for (let cc = c + 1; cc < row.length; cc++) {
              if (!isEmptyCell(row[cc])) return decodeExcelEscapes(row[cc]);
            }
          }
        }
      }
      return null;
    }

    function parseMeta(matrix, headerRowIndex) {
      const maxRows = Math.min(headerRowIndex != null ? headerRowIndex : 12, 30);
      const queryPeriod = findValueRightOfLabel(matrix, normalizeHeader('조회기간'), maxRows);
      const account = findValueRightOfLabel(matrix, normalizeHeader('계좌번호'), maxRows);
      const owner = findValueRightOfLabel(matrix, normalizeHeader('예금주명'), maxRows);
      const currentBalance = findValueRightOfLabel(matrix, normalizeHeader('현재통화잔액'), maxRows);
      return { queryPeriod, account, owner, currentBalance };
    }

    // === 값 파서(날짜/금액) ===
    function parseDateTime(value) {
      if (value == null) return null;

      if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

      // 숫자면 Excel date serial 가능성
      if (typeof value === 'number' && typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
        try {
          const o = XLSX.SSF.parse_date_code(value);
          if (o && o.y && o.m && o.d) return new Date(o.y, o.m - 1, o.d, o.H || 0, o.M || 0, o.S || 0);
        } catch (_) {
          // ignore
        }
      }

      const raw = decodeExcelEscapes(value);
      const s = String(raw).trim();
      if (!s) return null;

      // YYYY/MM/DD HH:MM:SS  or YYYY-MM-DD HH:MM:SS  or YYYY.MM.DD ...
      const m = s.match(
        /^(\d{4})[./-](\d{1,2})[./-](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/
      );
      if (!m) return null;

      const y = Number(m[1]);
      const mo = Number(m[2]);
      const d = Number(m[3]);
      const hh = Number(m[4] || 0);
      const mm = Number(m[5] || 0);
      const ss = Number(m[6] || 0);

      const dt = new Date(y, mo - 1, d, hh, mm, ss);
      if (Number.isNaN(dt.getTime())) return null;
      return dt;
    }

    function dateKeyFromDate(dt) {
      const y = dt.getFullYear();
      const m = String(dt.getMonth() + 1).padStart(2, '0');
      const d = String(dt.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    function monthKeyFromDate(dt) {
      const y = dt.getFullYear();
      const m = String(dt.getMonth() + 1).padStart(2, '0');
      return `${y}-${m}`;
    }

    function parseAmount(value) {
      if (value == null) return 0;
      if (typeof value === 'number' && Number.isFinite(value)) return Math.round(value);

      let s = decodeExcelEscapes(value);
      s = String(s);
      s = s.replace(/\u00A0/g, ' ');
      // 괄호 음수(회계표기) 지원
      let negative = false;
      const pm = s.match(/^\((.*)\)$/);
      if (pm) {
        negative = true;
        s = pm[1];
      }
      s = s.replace(/원/g, '').replace(/￦/g, '');
      s = s.replace(/[,\s]/g, '');
      s = s.replace(/[^0-9-]/g, '');

      if (!s || s === '-' || s === '--') return 0;
      const n = Number.parseInt(s, 10);
      if (!Number.isFinite(n)) return 0;
      return negative ? -Math.abs(n) : n;
    }

    function formatDateTime(dt) {
      const y = dt.getFullYear();
      const mo = String(dt.getMonth() + 1).padStart(2, '0');
      const d = String(dt.getDate()).padStart(2, '0');
      const hh = String(dt.getHours()).padStart(2, '0');
      const mm = String(dt.getMinutes()).padStart(2, '0');
      const ss = String(dt.getSeconds()).padStart(2, '0');
      return `${y}/${mo}/${d} ${hh}:${mm}:${ss}`;
    }

    function toIntOrNull(v) {
      if (v == null) return null;
      if (typeof v === 'number' && Number.isFinite(v)) return Math.trunc(v);
      const s = String(v).trim();
      if (!s) return null;
      const n = Number.parseInt(s.replace(/[^0-9-]/g, ''), 10);
      return Number.isFinite(n) ? n : null;
    }

    async function parseTransactions(matrix, headerRowIndex, mapping, onProgress) {
      const start = headerRowIndex + 1;
      const total = Math.max(0, matrix.length - start);

      const txs = [];
      const skipped = [];
      const unclassified = [];
      const anomalies = [];

      let lastTxRowIndex = null;

      for (let r = start; r < matrix.length; r++) {
        const row = matrix[r] || [];

        const dt = parseDateTime(row[mapping.datetime]);
        if (!dt) {
          // 의미 있는 값이 있는데 거래일자가 없으면 "스킵"으로 기록
          const meaningful = row.some((v) => !isEmptyCell(v));
          if (meaningful) {
            skipped.push({
              rowNumber: r + 1,
              preview: row
                .slice(0, 14)
                .map((v) => {
                  const dv = decodeExcelEscapes(v);
                  return dv == null ? '' : String(dv).replace(/\s+/g, ' ').trim();
                })
                .filter(Boolean)
                .slice(0, 4)
                .join(' | '),
            });
          }
          continue;
        }

        lastTxRowIndex = r;

        const seq = mapping.seq != null ? toIntOrNull(row[mapping.seq]) : null;
        const withdraw = parseAmount(row[mapping.withdraw]);
        const deposit = parseAmount(row[mapping.deposit]);
        const balance = parseAmount(row[mapping.balance]);
        const content = decodeExcelEscapes(row[mapping.content]) ?? '';
        const note = decodeExcelEscapes(row[mapping.note]) ?? '';
        const branch = mapping.branch != null ? decodeExcelEscapes(row[mapping.branch]) ?? '' : '';

        let type = 'none';
        if (withdraw > 0 && deposit > 0) {
          type = 'both';
          anomalies.push({ rowNumber: r + 1, reason: '출금/입금이 동시에 존재', withdraw, deposit });
        } else if (deposit > 0) {
          type = 'deposit';
        } else if (withdraw > 0) {
          type = 'withdraw';
        } else {
          type = 'none';
          unclassified.push({ rowNumber: r + 1, datetime: formatDateTime(dt) });
        }

        txs.push({
          rowNumber: r + 1,
          seq,
          datetime: dt,
          dateKey: dateKeyFromDate(dt),
          monthKey: monthKeyFromDate(dt),
          type,
          withdraw,
          deposit,
          balance,
          content: String(content).trim(),
          note: String(note).trim(),
          branch: String(branch).trim(),
        });

        // 진행률(대용량 파일에서 UI 멈춤 방지)
        if (onProgress && (r - start) % 250 === 0) {
          onProgress(r - start, total);
          await nextFrame();
        }
      }

      if (onProgress) onProgress(total, total);

      const footerStartIndex = lastTxRowIndex != null ? lastTxRowIndex + 1 : start;
      return { txs, logs: { skipped, unclassified, anomalies }, lastTxRowIndex, footerStartIndex };
    }

    function aggregate(transactions) {
      const totals = {
        withdraw: 0,
        deposit: 0,
        count: transactions.length,
      };

      let minDt = null;
      let maxDt = null;

      const daily = new Map(); // dateKey -> dayData
      const monthly = new Map(); // monthKey -> monthData

      for (const tx of transactions) {
        totals.withdraw += tx.withdraw;
        totals.deposit += tx.deposit;

        if (!minDt || tx.datetime < minDt) minDt = tx.datetime;
        if (!maxDt || tx.datetime > maxDt) maxDt = tx.datetime;

        // Daily
        if (!daily.has(tx.dateKey)) {
          daily.set(tx.dateKey, {
            dateKey: tx.dateKey,
            txs: [],
            withdraw: 0,
            deposit: 0,
            withdrawCount: 0,
            depositCount: 0,
          });
        }
        const day = daily.get(tx.dateKey);
        day.txs.push(tx);
        day.withdraw += tx.withdraw;
        day.deposit += tx.deposit;
        if (tx.withdraw > 0) day.withdrawCount++;
        if (tx.deposit > 0) day.depositCount++;

        // Monthly
        if (!monthly.has(tx.monthKey)) {
          monthly.set(tx.monthKey, { monthKey: tx.monthKey, withdraw: 0, deposit: 0 });
        }
        const mon = monthly.get(tx.monthKey);
        mon.withdraw += tx.withdraw;
        mon.deposit += tx.deposit;
      }

      // 정렬(기본: 내림차순)
      const dailyListDesc = Array.from(daily.values()).sort((a, b) => b.dateKey.localeCompare(a.dateKey));
      const monthlyListDesc = Array.from(monthly.values()).sort((a, b) => b.monthKey.localeCompare(a.monthKey));

      const depositDays = dailyListDesc.filter((d) => d.deposit > 0).map((d) => d.dateKey);

      return { totals, dailyListDesc, monthlyListDesc, depositDays, minDt, maxDt };
    }

    // === Export(엑셀/CSV) ===
    function cloneWorksheet(ws) {
      // 간단 복제(원본 ws를 건드리지 않기 위함)
      const out = {};
      for (const k of Object.keys(ws)) {
        const v = ws[k];
        if (k.startsWith('!')) out[k] = JSON.parse(JSON.stringify(v));
        else if (v && typeof v === 'object') out[k] = { ...v };
        else out[k] = v;
      }
      return out;
    }

    function buildAoAWithPadding(rows, maxCols) {
      return (rows || []).map((r) => padRow(r, maxCols));
    }

    function decodeAoAStrings(aoa) {
      return aoa.map((row) =>
        row.map((cell) => {
          if (typeof cell === 'string') return decodeExcelEscapes(cell);
          return cell;
        })
      );
    }

    function blankRow(maxCols) {
      return new Array(maxCols).fill(null);
    }

    function makeTxRow(tx, maxCols) {
      // 14열(NH 원본 규격) 형태로 맞춰서 뽑음
      // A(0): 빈칸
      // B(1): 구분
      // C(2): 거래일자
      // D(3): 출금금액, E(4): merge
      // F(5): 입금금액
      // G(6): 거래후잔액, H(7): merge
      // I(8): 거래내용, J(9): merge
      // K(10): 거래기록사항, L(11), M(12): merge
      // N(13): 거래점
      const row = blankRow(maxCols);
      row[1] = tx.seq != null ? tx.seq : null;
      row[2] = formatDateTime(tx.datetime);
      row[3] = tx.withdraw > 0 ? tx.withdraw : null;
      row[5] = tx.deposit > 0 ? tx.deposit : null;
      row[6] = tx.balance;
      row[8] = tx.content || null;
      row[10] = tx.note || null;
      row[13] = tx.branch || null;
      return row;
    }

    function makeSubtotalRow(day, maxCols) {
      const row = blankRow(maxCols);
      row[3] = day.withdraw || 0;
      row[5] = day.deposit || 0;
      return row;
    }

    function mergesForDataRow(r) {
      // r: 0-based row index in worksheet
      return [
        { s: { r, c: 3 }, e: { r, c: 4 } }, // D:E
        { s: { r, c: 6 }, e: { r, c: 7 } }, // G:H
        { s: { r, c: 8 }, e: { r, c: 9 } }, // I:J
        { s: { r, c: 10 }, e: { r, c: 12 } }, // K:M
      ];
    }

    function buildSheetAoA(analysis, mode /* 'all' | 'deposit' */) {
      const { maxCols, topRows, footerRows, originalMerges } = analysis;

      // top/footer는 "서식"이 아니라 "값"만 보존하면 충분하므로 AOA로 복제해서 씀
      const top = decodeAoAStrings(buildAoAWithPadding(topRows, maxCols));
      const footer = decodeAoAStrings(buildAoAWithPadding(footerRows, maxCols));

      // 그룹화(Export는 항상 날짜 내림차순)
      const dayMap = new Map();
      for (const d of analysis.agg.dailyListDesc) {
        dayMap.set(d.dateKey, {
          dateKey: d.dateKey,
          withdraw: d.withdraw,
          deposit: d.deposit,
          txs: [],
        });
      }
      for (const tx of analysis.transactions) {
        const day = dayMap.get(tx.dateKey);
        if (!day) continue;
        if (mode === 'deposit' && tx.deposit <= 0) continue;
        day.txs.push(tx);
      }

      const aoa = [...top];
      const merges = [];

      // 원본 상단(헤더 포함) merge는 그대로 복사(행 위치가 같을 때만 의미 있음)
      if (Array.isArray(originalMerges)) {
        for (const m of originalMerges) {
          // headerRowIndex(0-based) 이하만 복사(그 아래는 행이 달라지므로 깨짐)
          if (m.e.r <= analysis.headerRowIndex) merges.push(JSON.parse(JSON.stringify(m)));
        }
      }

      for (const day of analysis.agg.dailyListDesc) {
        const g = dayMap.get(day.dateKey);
        if (!g) continue;

        // mode=deposit인 경우: 해당 일자에 입금 행이 하나도 없으면 스킵(샘플 "입금" 시트 동작에 가깝게)
        if (mode === 'deposit' && g.txs.length === 0) continue;

        // 거래 행 (시간 내림차순 유지)
        g.txs.sort((a, b) => b.datetime - a.datetime || (a.seq ?? 0) - (b.seq ?? 0));
        for (const tx of g.txs) {
          const rowIndex = aoa.length; // push 전 인덱스
          aoa.push(makeTxRow(tx, maxCols));
          merges.push(...mergesForDataRow(rowIndex));
        }

        // 일자 합계 행 + 공백행
        aoa.push(makeSubtotalRow(g, maxCols));
        aoa.push(blankRow(maxCols));
      }

      aoa.push(...footer);

      // footer(농협 확인 문구 등): B열 단독 메시지면 B~마지막열까지 병합(원본 느낌)
      const footerStartRow = aoa.length - footer.length;
      for (let i = 0; i < footer.length; i++) {
        const r = footerStartRow + i;
        const row = footer[i] || [];
        const hasMsg = typeof row[1] === 'string' && row[1].trim() !== '';
        if (!hasMsg) continue;
        const restEmpty = row.slice(2).every((v) => isEmptyCell(v));
        if (restEmpty) merges.push({ s: { r, c: 1 }, e: { r, c: maxCols - 1 } });
      }

      return { aoa, merges };
    }

    function applyNumberFormat(ws, maxCols) {
      if (!ws || !ws['!ref']) return;

      const range = XLSX.utils.decode_range(ws['!ref']);
      const fmt = '#,##0';

      const targetCols = [3, 5, 6]; // D, F, G
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (const c of targetCols) {
          if (c > range.e.c) continue;
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws[addr];
          if (!cell) continue;
          if (cell.t === 'n') cell.z = fmt;
        }
      }
    }

    function buildExportWorkbook(originalWb, originalSheetName, analysis) {
      const out = XLSX.utils.book_new();

      // Sheet 1: 원본
      const orig = cloneWorksheet(originalWb.Sheets[originalSheetName]);
      XLSX.utils.book_append_sheet(out, orig, '통장거래내역 (원본)');

      // Sheet 2: 입출금(전체)
      const s2 = buildSheetAoA(analysis, 'all');
      const ws2 = XLSX.utils.aoa_to_sheet(s2.aoa);
      ws2['!merges'] = s2.merges;
      applyNumberFormat(ws2, analysis.maxCols);
      XLSX.utils.book_append_sheet(out, ws2, '통장거래내역 (입출금)');

      // Sheet 3: 입금(필터)
      const s3 = buildSheetAoA(analysis, 'deposit');
      const ws3 = XLSX.utils.aoa_to_sheet(s3.aoa);
      ws3['!merges'] = s3.merges;
      applyNumberFormat(ws3, analysis.maxCols);
      XLSX.utils.book_append_sheet(out, ws3, '통장거래내역 (입금)');

      return out;
    }

    function workbookToBlob(wb) {
      const array = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
      return new Blob([array], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    }

    function escapeCsvCell(value) {
      if (value == null) return '';
      const s = String(value);
      if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    }

    function makeCsv(transactions, mode /* 'all' | 'deposit' */) {
      const header = [
        '거래일자',
        '일자',
        '구분',
        '출금금액',
        '입금금액',
        '거래후잔액',
        '거래내용',
        '거래기록사항',
        '거래점',
        '원본행번호',
      ];

      const rows = [header];

      for (const tx of transactions) {
        if (mode === 'deposit' && tx.deposit <= 0) continue;
        rows.push([
          formatDateTime(tx.datetime),
          tx.dateKey,
          tx.type,
          tx.withdraw || 0,
          tx.deposit || 0,
          tx.balance,
          tx.content,
          tx.note,
          tx.branch,
          tx.rowNumber,
        ]);
      }

      return rows.map((r) => r.map(escapeCsvCell).join(',')).join('\n');
    }

    // === 공개 API ===
    function prepareWorksheet(ws) {
      const maxCols = getMaxCols(ws);
      const raw = sheetToMatrix(ws).map((r) => padRow(r, maxCols));
      const filled = raw.map((r) => r.slice());
      fillMerges(filled, ws['!merges'] || []);

      const headerDetect = detectHeaderRow(filled);
      const headerRowIndex = headerDetect.rowIndex != null ? headerDetect.rowIndex : DEFAULT_HEADER_ROW_INDEX;

      const headerRow = filled[headerRowIndex] || [];
      const auto = autoDetectMapping(headerRow, maxCols);
      const meta = parseMeta(filled, headerRowIndex);

      return {
        maxCols,
        rawMatrix: raw,
        matrix: filled,
        headerRowIndex,
        headerDetect,
        autoMapping: auto,
        meta,
        originalMerges: ws['!merges'] || [],
      };
    }

    async function analyzeWorkbook(workbook, sheetName, userConfig, onProgress) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) throw new Error('워크시트를 찾을 수 없습니다.');

      const prep = prepareWorksheet(ws);

      const headerRowIndex =
        userConfig && typeof userConfig.headerRowIndex === 'number' ? userConfig.headerRowIndex : prep.headerRowIndex;

      // 매핑: (1) 사용자 지정 → (2) 자동 감지 → (3) 기본값
      const mapping = {
        ...DEFAULT_MAPPING,
        ...(prep.autoMapping?.mapping || {}),
        ...(userConfig?.mapping || {}),
      };

      const missing = REQUIRED_FIELDS.filter((f) => mapping[f] == null);
      if (missing.length) {
        const missingNames = missing.map((f) => FIELD_LABELS[f] || f).join(', ');
        throw new Error(`필수 컬럼 매핑 누락: ${missingNames}`);
      }

      const parsed = await parseTransactions(prep.matrix, headerRowIndex, mapping, onProgress);
      const agg = aggregate(parsed.txs);

      const topRows = prep.rawMatrix.slice(0, headerRowIndex + 1);
      const footerRows = prep.rawMatrix.slice(parsed.footerStartIndex);

      return {
        // 입력
        workbook,
        sheetName,
        maxCols: prep.maxCols,
        headerRowIndex,
        headerDetect: prep.headerDetect,
        mapping,
        meta: prep.meta,
        originalMerges: prep.originalMerges,

        // 결과
        transactions: parsed.txs,
        logs: parsed.logs,
        agg,
        topRows,
        footerRows,
      };
    }

    return {
      FIELD_LABELS,
      REQUIRED_FIELDS,
      DEFAULT_MAPPING,
      DEFAULT_HEADER_ROW_INDEX,

      decodeExcelEscapes,
      normalizeHeader,
      prepareWorksheet,
      analyzeWorkbook,

      buildExportWorkbook,
      workbookToBlob,
      makeCsv,
      formatDateTime,
    };
  })();

  /* =========================================================
     UI (렌더링/이벤트)
     ========================================================= */
  const UI = (() => {
    // --- Elements ---
    const el = {
      dropZone: null,
      fileInput: null,
      btnPick: null,
      btnReset: null,
      fileMeta: null,

      pLoad: null,
      loadPct: null,
      loadLabel: null,

      pAnalyze: null,
      analyzePct: null,
      analyzeLabel: null,

      libStatus: null,

      btnRun: null,

      mappingDetails: null,
      mappingStatus: null,
      inpHeaderRow: null,
      selCols: {}, // field -> select
      btnApplyMapping: null,

      logCard: null,
      logSummary: null,
      logList: null,

      secDashboard: null,
      secDetails: null,
      secDownload: null,

      periodLabel: null,
      dashboardTotals: null,
      monthlyTable: null,
      depositDaysTable: null,

      selOrder: null,
      selMode: null,
      dateList: null,
      detailPanel: null,

      btnDownloadXlsx: null,
      btnDownloadCsvAll: null,
      btnDownloadCsvDeposit: null,

      btnToTop: null,
    };

    const state = {
      file: null,
      workbook: null,
      sheetName: null,
      prepared: null, // Engine.prepareWorksheet 결과
      mapping: null, // 수동 적용된 mapping (부분/전체)
      headerRowIndex: null,
      analysis: null,
      view: {
        order: 'desc',
        mode: 'deposit', // 'deposit' | 'all'
      },
    };

    // --- UI util ---
    function setText(node, text) {
      if (!node) return;
      node.textContent = text;
    }

    function show(node, yes) {
      if (!node) return;
      node.hidden = !yes;
    }

    function setDisabled(node, yes) {
      if (!node) return;
      node.disabled = !!yes;
    }

    function formatNumber(n) {
      if (n == null) return '';
      const num = Number(n);
      if (!Number.isFinite(num)) return '';
      return num.toLocaleString('ko-KR');
    }

    function humanFileSize(bytes) {
      if (!Number.isFinite(bytes)) return '';
      const units = ['B', 'KB', 'MB', 'GB'];
      let v = bytes;
      let i = 0;
      while (v >= 1024 && i < units.length - 1) {
        v /= 1024;
        i++;
      }
      return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
    }

    function downloadBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }

    function safeFileBaseName(name) {
      const base = (name || '집계결과').replace(/\.[^.]+$/, '');
      return base.replace(/[\\/:*?"<>|]+/g, '_');
    }

    function resetAll() {
      state.file = null;
      state.workbook = null;
      state.sheetName = null;
      state.prepared = null;
      state.mapping = null;
      state.headerRowIndex = null;
      state.analysis = null;

      el.fileInput.value = '';
      setText(el.fileMeta, '');
      el.dropZone.classList.remove('dragover');

      el.pLoad.value = 0;
      setText(el.loadPct, '0%');
      setText(el.loadLabel, '');

      el.pAnalyze.value = 0;
      setText(el.analyzePct, '0%');
      setText(el.analyzeLabel, '');

      setDisabled(el.btnRun, true);
      setDisabled(el.btnReset, true);

      setDisabled(el.btnDownloadXlsx, true);
      setDisabled(el.btnDownloadCsvAll, true);
      setDisabled(el.btnDownloadCsvDeposit, true);

      show(el.secDashboard, false);
      show(el.secDetails, false);
      show(el.secDownload, false);

      show(el.logCard, false);
      setText(el.logSummary, '');
      el.logList.innerHTML = '';

      setText(el.mappingStatus, '파일을 업로드하면 자동 감지 결과가 표시됩니다.');
      if (el.mappingDetails) el.mappingDetails.open = false;
    }

    // --- Mapping UI ---
    function getColumnOptions(maxCols) {
      const cols = [];
      for (let c = 0; c < maxCols; c++) {
        cols.push({ value: String(c), label: String.fromCharCode('A'.charCodeAt(0) + c) });
      }
      return cols;
    }

    function fillSelectOptions(selectEl, options, selectedValue) {
      selectEl.innerHTML = '';
      const optEmpty = document.createElement('option');
      optEmpty.value = '';
      optEmpty.textContent = '(선택 안 함)';
      selectEl.appendChild(optEmpty);

      for (const o of options) {
        const opt = document.createElement('option');
        opt.value = o.value;
        opt.textContent = o.label;
        selectEl.appendChild(opt);
      }
      selectEl.value = selectedValue != null ? String(selectedValue) : '';
    }

    function renderMappingUI(prep) {
      const maxCols = prep.maxCols || 14;
      const options = getColumnOptions(maxCols);

      // 헤더 행(1-based 표시)
      const headerRow1 = (prep.headerRowIndex ?? Engine.DEFAULT_HEADER_ROW_INDEX) + 1;
      el.inpHeaderRow.value = String(headerRow1);

      const mapping = prep.autoMapping?.mapping || { ...Engine.DEFAULT_MAPPING };

      // select들 채우기
      for (const field of ['seq', 'datetime', 'withdraw', 'deposit', 'balance', 'content', 'note', 'branch']) {
        const sel = el.selCols[field];
        if (!sel) continue;
        fillSelectOptions(sel, options, mapping[field]);
      }

      const detectOk = prep.headerDetect?.rowIndex != null && (prep.autoMapping?.missing || []).length === 0;

      const parts = [];
      parts.push(
        `헤더 감지: ${
          prep.headerDetect?.rowIndex != null ? `${prep.headerDetect.rowIndex + 1}행` : '실패(기본 12행 가정)'
        }`
      );
      parts.push(`컬럼 매핑: ${detectOk ? '자동 감지 성공' : '자동 감지 불완전(수동 확인 권장)'}`);

      if (prep.autoMapping?.missing?.length) {
        parts.push(`누락: ${prep.autoMapping.missing.map((f) => Engine.FIELD_LABELS[f] || f).join(', ')}`);
      }

      setText(el.mappingStatus, parts.join(' / '));

      // 자동 감지 실패 느낌이면 details를 열어줌(사용자 편의)
      if (!detectOk) el.mappingDetails.open = true;
    }

    function readMappingFromUI() {
      // 헤더 행
      let headerRowIndex = null;
      const v = Number.parseInt(el.inpHeaderRow.value, 10);
      if (Number.isFinite(v) && v >= 1) headerRowIndex = v - 1;

      const mapping = {};
      for (const [field, sel] of Object.entries(el.selCols)) {
        const s = sel.value;
        if (s === '') continue;
        const n = Number.parseInt(s, 10);
        if (Number.isFinite(n)) mapping[field] = n;
      }

      return { headerRowIndex, mapping };
    }

    // --- Logs ---
    function renderLogs(analysis) {
      const logs = analysis.logs || {};
      const skipped = logs.skipped || [];
      const unclassified = logs.unclassified || [];
      const anomalies = logs.anomalies || [];

      const meta = analysis.meta || {};

      const lines = [];
      lines.push(`• 거래행 파싱: ${analysis.transactions.length.toLocaleString('ko-KR')}건`);
      lines.push(`• 스킵된 행(거래일자 없음): ${skipped.length.toLocaleString('ko-KR')}건`);
      lines.push(`• 미분류(출금/입금 모두 0): ${unclassified.length.toLocaleString('ko-KR')}건`);
      lines.push(`• 이상치(출금/입금 동시): ${anomalies.length.toLocaleString('ko-KR')}건`);

      if (meta.queryPeriod) lines.push(`• 조회기간(원본): ${meta.queryPeriod}`);
      if (analysis.headerDetect?.rowIndex == null) lines.push(`• 헤더 자동 감지 실패 → 기본 12행 가정`);

      setText(el.logSummary, lines.join('\n'));
      el.logSummary.style.whiteSpace = 'pre-line';

      el.logList.innerHTML = '';
      const showItems = [];

      if (skipped.length) {
        showItems.push(...skipped.slice(0, 8).map((x) => `스킵 ${x.rowNumber}행: ${x.preview}`));
      }
      if (unclassified.length) {
        showItems.push(...unclassified.slice(0, 5).map((x) => `미분류 ${x.rowNumber}행: ${x.datetime}`));
      }
      if (anomalies.length) {
        showItems.push(
          ...anomalies.slice(0, 5).map((x) => `이상치 ${x.rowNumber}행: ${x.reason} (출금 ${x.withdraw}, 입금 ${x.deposit})`)
        );
      }

      for (const s of showItems) {
        const li = document.createElement('li');
        li.textContent = s;
        el.logList.appendChild(li);
      }

      show(el.logCard, true);
    }

    // --- Dashboard ---
    function renderDashboard(analysis) {
      const { agg, meta } = analysis;

      const period = meta?.queryPeriod
        ? meta.queryPeriod
        : agg.minDt && agg.maxDt
          ? `${Engine.formatDateTime(agg.minDt)} ~ ${Engine.formatDateTime(agg.maxDt)}`
          : '-';

      setText(el.periodLabel, `조회기간: ${period}`);

      // totals
      const totWithdraw = agg.totals.withdraw;
      const totDeposit = agg.totals.deposit;
      const net = totDeposit - totWithdraw;

      el.dashboardTotals.innerHTML = `
        <div class="card">
          <div class="muted">총 입금액</div>
          <div style="font-size:1.5rem; font-weight:800;">${formatNumber(totDeposit)}</div>
        </div>
        <div class="card">
          <div class="muted">총 출금액</div>
          <div style="font-size:1.5rem; font-weight:800;">${formatNumber(totWithdraw)}</div>
        </div>
        <div class="card">
          <div class="muted">순증감(입금-출금)</div>
          <div style="font-size:1.5rem; font-weight:800;">${formatNumber(net)}</div>
        </div>
      `;

      // monthly table
      el.monthlyTable.innerHTML = '';
      const thead = document.createElement('thead');
      thead.innerHTML = `
        <tr>
          <th>월</th>
          <th class="numeric">입금 합계</th>
          <th class="numeric">출금 합계</th>
          <th class="numeric">순증감</th>
        </tr>`;
      el.monthlyTable.appendChild(thead);

      const tbody = document.createElement('tbody');
      for (const m of agg.monthlyListDesc) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${m.monthKey}</td>
          <td class="numeric">${formatNumber(m.deposit)}</td>
          <td class="numeric">${formatNumber(m.withdraw)}</td>
          <td class="numeric">${formatNumber(m.deposit - m.withdraw)}</td>`;
        tbody.appendChild(tr);
      }
      el.monthlyTable.appendChild(tbody);

      // deposit days table (입금 있는 날짜만)
      el.depositDaysTable.innerHTML = '';
      const thead2 = document.createElement('thead');
      thead2.innerHTML = `
        <tr>
          <th>일자</th>
          <th class="numeric">입금 합계</th>
          <th class="numeric">입금 건수</th>
        </tr>`;
      el.depositDaysTable.appendChild(thead2);

      const tbody2 = document.createElement('tbody');
      for (const d of agg.dailyListDesc.filter((x) => x.deposit > 0)) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${d.dateKey}</td>
          <td class="numeric">${formatNumber(d.deposit)}</td>
          <td class="numeric">${formatNumber(d.depositCount)}</td>`;
        tbody2.appendChild(tr);
      }
      el.depositDaysTable.appendChild(tbody2);
    }

    // --- Details (패널1/패널2) ---
    function renderDateList(analysis, order) {
      const list = order === 'asc'
        ? [...analysis.agg.dailyListDesc].sort((a, b) => a.dateKey.localeCompare(b.dateKey))
        : analysis.agg.dailyListDesc;

      el.dateList.innerHTML = '';
      for (const d of list) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `btn btn-lightgrey date-chip ${d.deposit > 0 ? 'pastel-blue' : ''}`;
        btn.textContent = d.dateKey;
        btn.dataset.day = d.dateKey;
        btn.title = d.deposit > 0 ? `입금 ${formatNumber(d.deposit)} / 출금 ${formatNumber(d.withdraw)}` : `출금 ${formatNumber(d.withdraw)}`;

        btn.addEventListener('click', () => {
          const target = $(`#day-${d.dateKey.replaceAll('-', '')}`);
          if (target) {
            // 기본적으로 해당 일자 펼치기
            if (target.tagName.toLowerCase() === 'details') target.open = true;
            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        });

        el.dateList.appendChild(btn);
      }
    }

    function buildTxTable(txs) {
      const wrap = document.createElement('div');
      wrap.className = 'table-wrap';

      const table = document.createElement('table');
      table.className = 'sheetlike sticky-table';
      table.innerHTML = `
        <thead>
          <tr>
            <th class="sticky-1 col-seq">구분</th>
            <th class="sticky-2 col-datetime">거래일자</th>
            <th class="numeric">출금금액</th>
            <th class="numeric">입금금액</th>
            <th class="numeric">거래후잔액</th>
            <th>거래내용</th>
            <th>거래기록사항</th>
            <th>거래점</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;

      const tbody = $('tbody', table);

      for (const tx of txs) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td class="sticky-1 col-seq">${tx.seq ?? ''}</td>
          <td class="sticky-2 col-datetime">${Engine.formatDateTime(tx.datetime)}</td>
          <td class="numeric">${tx.withdraw > 0 ? formatNumber(tx.withdraw) : ''}</td>
          <td class="numeric">${tx.deposit > 0 ? formatNumber(tx.deposit) : ''}</td>
          <td class="numeric">${formatNumber(tx.balance)}</td>
          <td>${tx.content ?? ''}</td>
          <td>${tx.note ?? ''}</td>
          <td>${tx.branch ?? ''}</td>
        `;
        tbody.appendChild(tr);
      }

      wrap.appendChild(table);
      return wrap;
    }

    function renderDetailPanel(analysis, order, mode) {
      const list = order === 'asc'
        ? [...analysis.agg.dailyListDesc].sort((a, b) => a.dateKey.localeCompare(b.dateKey))
        : analysis.agg.dailyListDesc;

      // 날짜별 tx Map
      const txMap = new Map();
      for (const tx of analysis.transactions) {
        if (!txMap.has(tx.dateKey)) txMap.set(tx.dateKey, []);
        txMap.get(tx.dateKey).push(tx);
      }
      // 각 날짜 내부는 시간 내림차순(원본과 유사)
      for (const arr of txMap.values()) {
        arr.sort((a, b) => b.datetime - a.datetime || (a.seq ?? 0) - (b.seq ?? 0));
      }

      el.detailPanel.innerHTML = '';

      for (const day of list) {
        const dayTxsAll = txMap.get(day.dateKey) || [];
        const dayTxs = mode === 'deposit' ? dayTxsAll.filter((t) => t.deposit > 0) : dayTxsAll;

        const details = document.createElement('details');
        details.className = 'day-block';
        details.id = `day-${day.dateKey.replaceAll('-', '')}`;
        // 기본 펼침 규칙: 입금 1건 이상이면 펼침, 아니면 접기
        details.open = day.deposit > 0;

        const summary = document.createElement('summary');
        summary.className = 'day-summary';
        summary.innerHTML = `
          <span class="day-title">${day.dateKey}</span>
          <span class="day-badges">
            <span class="badge ${day.deposit > 0 ? 'badge-deposit' : 'badge-muted'}">입금 ${formatNumber(day.deposit)}</span>
            <span class="badge ${day.withdraw > 0 ? 'badge-withdraw' : 'badge-muted'}">출금 ${formatNumber(day.withdraw)}</span>
            <span class="badge badge-muted">거래 ${formatNumber(day.withdrawCount + day.depositCount)}건</span>
          </span>
        `;
        details.appendChild(summary);

        const inner = document.createElement('div');
        inner.className = 'day-inner';

        if (mode === 'deposit' && dayTxs.length === 0) {
          const p = document.createElement('p');
          p.className = 'muted';
          p.textContent = '입금 내역이 없습니다. (출금만 있는 일자)';
          inner.appendChild(p);
        } else {
          inner.appendChild(buildTxTable(dayTxs));
        }

        // 일자 합계 라인
        const sumLine = document.createElement('div');
        sumLine.className = 'day-sumline';
        sumLine.innerHTML = `
          <span class="muted">일자 합계</span>
          <span class="sumvals">
            <span>입금 <strong>${formatNumber(day.deposit)}</strong></span>
            <span>출금 <strong>${formatNumber(day.withdraw)}</strong></span>
          </span>
        `;
        inner.appendChild(sumLine);

        details.appendChild(inner);
        el.detailPanel.appendChild(details);
      }
    }

    function renderDetails(analysis) {
      renderDateList(analysis, state.view.order);
      renderDetailPanel(analysis, state.view.order, state.view.mode);
    }

    // --- Event handlers ---
    async function handleFile(file) {
      resetAll();
      state.file = file;

      setDisabled(el.btnReset, false);
      setText(el.fileMeta, `${file.name} (${humanFileSize(file.size)})`);

      // FileReader: ArrayBuffer 로드
      el.pLoad.value = 0;
      setText(el.loadPct, '0%');
      setText(el.loadLabel, '읽는 중…');

      const buf = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('파일을 읽을 수 없습니다.'));
        reader.onprogress = (e) => {
          if (e.lengthComputable) {
            const pct = Math.round((e.loaded / e.total) * 100);
            el.pLoad.value = pct;
            setText(el.loadPct, `${pct}%`);
          }
        };
        reader.onload = () => resolve(reader.result);
        reader.readAsArrayBuffer(file);
      });

      el.pLoad.value = 100;
      setText(el.loadPct, '100%');
      setText(el.loadLabel, '로드 완료');

      // XLSX 파싱
      setText(el.analyzeLabel, '대기 중');
      el.pAnalyze.value = 0;
      setText(el.analyzePct, '0%');

      try {
        const wb = XLSX.read(buf, { type: 'array', cellDates: true, cellNF: true, cellText: false });
        state.workbook = wb;
        state.sheetName = wb.SheetNames[0];
        if (!state.sheetName) throw new Error('시트가 없습니다.');

        // 매핑 자동 감지(가벼운 준비 단계)
        const prep = Engine.prepareWorksheet(wb.Sheets[state.sheetName]);
        state.prepared = prep;
        state.headerRowIndex = prep.headerRowIndex;

        renderMappingUI(prep);

        setDisabled(el.btnRun, false);
        setText(el.analyzeLabel, `분석 준비 완료 (시트: ${state.sheetName})`);
      } catch (err) {
        setDisabled(el.btnRun, true);
        setText(el.analyzeLabel, `오류: ${err.message}`);
      }
    }

    async function runAnalysis() {
      if (!state.workbook || !state.sheetName) return;

      setDisabled(el.btnRun, true);
      setDisabled(el.btnDownloadXlsx, true);
      setDisabled(el.btnDownloadCsvAll, true);
      setDisabled(el.btnDownloadCsvDeposit, true);

      show(el.secDashboard, false);
      show(el.secDetails, false);
      show(el.secDownload, false);

      show(el.logCard, false);
      el.logList.innerHTML = '';

      el.pAnalyze.value = 0;
      setText(el.analyzePct, '0%');
      setText(el.analyzeLabel, '파일 분석 중…');

      // 사용자 매핑 적용 상태 반영
      const userCfg = {};
      if (state.headerRowIndex != null) userCfg.headerRowIndex = state.headerRowIndex;
      if (state.mapping) userCfg.mapping = state.mapping;

      try {
        const analysis = await Engine.analyzeWorkbook(
          state.workbook,
          state.sheetName,
          userCfg,
          (done, total) => {
            const pct = total > 0 ? Math.round((done / total) * 100) : 100;
            el.pAnalyze.value = pct;
            setText(el.analyzePct, `${pct}%`);
          }
        );

        state.analysis = analysis;

        setText(el.analyzeLabel, '분석 완료 ✅');
        el.pAnalyze.value = 100;
        setText(el.analyzePct, '100%');

        renderLogs(analysis);
        renderDashboard(analysis);
        renderDetails(analysis);

        show(el.secDashboard, true);
        show(el.secDetails, true);
        show(el.secDownload, true);

        setDisabled(el.btnDownloadXlsx, false);
        setDisabled(el.btnDownloadCsvAll, false);
        setDisabled(el.btnDownloadCsvDeposit, false);
      } catch (err) {
        setText(el.analyzeLabel, `오류: ${err.message}`);
        setDisabled(el.btnRun, false);
      } finally {
        setDisabled(el.btnRun, false);
      }
    }

    function applyManualMapping() {
      if (!state.prepared) return;

      const { headerRowIndex, mapping } = readMappingFromUI();
      if (headerRowIndex != null) state.headerRowIndex = headerRowIndex;
      state.mapping = mapping;

      // 간단 검증(필수 매핑 누락)
      const merged = { ...Engine.DEFAULT_MAPPING, ...(state.prepared.autoMapping?.mapping || {}), ...(state.mapping || {}) };
      const missing = Engine.REQUIRED_FIELDS.filter((f) => merged[f] == null);
      if (missing.length) {
        setText(
          el.mappingStatus,
          `수동 매핑 적용됨 / 경고: 필수 컬럼 누락 → ${missing.map((f) => Engine.FIELD_LABELS[f] || f).join(', ')}`
        );
        el.mappingStatus.style.color = '#b42318';
      } else {
        setText(el.mappingStatus, '수동 매핑 적용됨 / 분석 버튼을 눌러 집계하세요.');
        el.mappingStatus.style.color = '';
      }
    }

    function bindDropZone() {
      const dz = el.dropZone;

      dz.addEventListener('click', () => el.fileInput.click());
      dz.addEventListener('dragover', (e) => {
        e.preventDefault();
        dz.classList.add('dragover');
      });
      dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
      dz.addEventListener('drop', (e) => {
        e.preventDefault();
        dz.classList.remove('dragover');
        const file = e.dataTransfer?.files?.[0];
        if (file) handleFile(file);
      });
    }

    function bindToTop() {
      const btn = el.btnToTop;
      const onScroll = () => {
        btn.style.display = window.scrollY > 400 ? 'inline-block' : 'none';
      };
      window.addEventListener('scroll', onScroll);
      onScroll();

      btn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
    }

    function bindControls() {
      el.btnPick.addEventListener('click', () => el.fileInput.click());
      el.fileInput.addEventListener('change', (e) => {
        const file = e.target.files?.[0];
        if (file) handleFile(file);
      });

      el.btnReset.addEventListener('click', resetAll);
      el.btnRun.addEventListener('click', runAnalysis);

      el.btnApplyMapping.addEventListener('click', applyManualMapping);

      el.selOrder.addEventListener('change', () => {
        state.view.order = el.selOrder.value;
        if (state.analysis) renderDetails(state.analysis);
      });
      el.selMode.addEventListener('change', () => {
        state.view.mode = el.selMode.value;
        if (state.analysis) renderDetails(state.analysis);
      });

      el.btnDownloadXlsx.addEventListener('click', () => {
        if (!state.analysis) return;
        const wb = Engine.buildExportWorkbook(state.workbook, state.sheetName, state.analysis);
        const blob = Engine.workbookToBlob(wb);
        const name = `${safeFileBaseName(state.file?.name || '통장거래내역')}_집계.xlsx`;
        downloadBlob(blob, name);
      });

      el.btnDownloadCsvAll.addEventListener('click', () => {
        if (!state.analysis) return;
        const csv = Engine.makeCsv(state.analysis.transactions, 'all');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const name = `${safeFileBaseName(state.file?.name || '통장거래내역')}_입출금.csv`;
        downloadBlob(blob, name);
      });

      el.btnDownloadCsvDeposit.addEventListener('click', () => {
        if (!state.analysis) return;
        const csv = Engine.makeCsv(state.analysis.transactions, 'deposit');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const name = `${safeFileBaseName(state.file?.name || '통장거래내역')}_입금.csv`;
        downloadBlob(blob, name);
      });
    }

    function cacheEls() {
      el.dropZone = $('#dropZone');
      el.fileInput = $('#fileInput');
      el.btnPick = $('#btnPick');
      el.btnReset = $('#btnReset');
      el.fileMeta = $('#fileMeta');

      el.pLoad = $('#pLoad');
      el.loadPct = $('#loadPct');
      el.loadLabel = $('#loadLabel');

      el.pAnalyze = $('#pAnalyze');
      el.analyzePct = $('#analyzePct');
      el.analyzeLabel = $('#analyzeLabel');

      el.libStatus = $('#libStatus');

      el.btnRun = $('#btnRun');

      el.mappingDetails = $('#mappingDetails');
      el.mappingStatus = $('#mappingStatus');
      el.inpHeaderRow = $('#inpHeaderRow');
      el.btnApplyMapping = $('#btnApplyMapping');

      el.selCols.seq = $('#selSeq');
      el.selCols.datetime = $('#selDatetime');
      el.selCols.withdraw = $('#selWithdraw');
      el.selCols.deposit = $('#selDeposit');
      el.selCols.balance = $('#selBalance');
      el.selCols.content = $('#selContent');
      el.selCols.note = $('#selNote');
      el.selCols.branch = $('#selBranch');

      el.logCard = $('#logCard');
      el.logSummary = $('#logSummary');
      el.logList = $('#logList');

      el.secDashboard = $('#secDashboard');
      el.secDetails = $('#secDetails');
      el.secDownload = $('#secDownload');

      el.periodLabel = $('#periodLabel');
      el.dashboardTotals = $('#dashboardTotals');
      el.monthlyTable = $('#monthlyTable');
      el.depositDaysTable = $('#depositDaysTable');

      el.selOrder = $('#selOrder');
      el.selMode = $('#selMode');
      el.dateList = $('#dateList');
      el.detailPanel = $('#detailPanel');

      el.btnDownloadXlsx = $('#btnDownloadXlsx');
      el.btnDownloadCsvAll = $('#btnDownloadCsvAll');
      el.btnDownloadCsvDeposit = $('#btnDownloadCsvDeposit');

      el.btnToTop = $('#btnToTop');
    }

    function init() {
      cacheEls();
      bindDropZone();
      bindControls();
      bindToTop();
      resetAll();
    }

    return { init };
  })();

  // Bootstrap
  window.addEventListener('DOMContentLoaded', UI.init);
})();
