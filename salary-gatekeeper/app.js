/* 출입문개폐전담원 월별 인건비 지출서식 (클라이언트 전용)
 * - Excel 업/다운로드: SheetJS(XLSX)
 * - PDF: html2canvas + jsPDF
 * - 입력값 localStorage 저장
 */
(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);

  const STORAGE_KEY = "doorGatekeeperWage_v1";

  const weekdayKorean = ["일", "월", "화", "수", "목", "금", "토"];
  const weekdayOrder = [1, 2, 3, 4, 5, 6, 0]; // schedule 표는 월~일 순으로 보여주기

  const defaultSchedule = () => ({
    // key: 0..6 (Sun..Sat)
    0: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    1: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    2: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    3: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    4: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    5: { s1: "", e1: "", s2: "", e2: "", memo: "" },
    6: { s1: "", e1: "", s2: "", e2: "", memo: "" },
  });

  const state = {
    payMonth: "", // YYYY-MM
    payDate: "",  // YYYY-MM-DD
    schoolName: "",
    workerName: "",
    jobTitle: "출입문개폐전담원",
    birthDate: "",
    hireDate: "",
    hourlyRate: 11500,

    // schedule by weekday
    schedule: defaultSchedule(),

    // attendance: { "YYYY-MM-DD": { on: boolean, manualHours: number|null } }
    attendance: {},

    // deductions & rates
    truncate10won: true,
    deductHealthOn: false,
    healthRate: 0.03595,
    healthAmt: 0,
    deductLtCareOn: false,
    ltCareRate: 0.1314,
    ltCareAmt: 0,
    deductEmpOn: false,
    empRate: 0.009,
    empAmt: 0,

    incomeTax: 0,
    localTax: 0,
    pension: 0,
    otherDeduct: 0,

    // manual override flags
    manual: {
      healthAmt: false,
      ltCareAmt: false,
      empAmt: false,
    },
    ui: {
      showManualDailyHours: false
    }
  };

  // ---------- utils ----------
  const pad2 = (n) => String(n).padStart(2, "0");
  const toISODate = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

  const parseNumber = (v, fallback = 0) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  };

  const floor10 = (n) => Math.floor(n / 10) * 10;

  const fmtWon = (n) => {
    const x = Math.round(parseNumber(n, 0));
    return x.toLocaleString("ko-KR");
  };

  const timeToMinutes = (t) => {
    if (!t || typeof t !== "string") return null;
    const m = /^(\d{1,2}):(\d{2})$/.exec(t.trim());
    if (!m) return null;
    const hh = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return hh * 60 + mm;
  };

  const diffMinutes = (start, end) => {
    const s = timeToMinutes(start);
    const e = timeToMinutes(end);
    if (s == null || e == null) return 0;
    let d = e - s;
    if (d < 0) d += 24 * 60; // 자정 넘어감
    return d;
  };

  const minutesToHours = (mins) => Math.round((mins / 60) * 100) / 100; // 소수 2자리

  const isLibReady = () => {
    const okXlsx = typeof window.XLSX !== "undefined";
    const okCanvas = typeof window.html2canvas !== "undefined";
    const okJsPdf = window.jspdf && window.jspdf.jsPDF;
    return { okXlsx, okCanvas, okJsPdf };
  };

  // ---------- schedule rendering ----------
  function renderSchedule() {
    const tbody = $("scheduleTbody");
    tbody.innerHTML = "";

    weekdayOrder.forEach((dow, rIdx) => {
      const row = document.createElement("tr");

      const labelTd = document.createElement("td");
      labelTd.textContent = weekdayKorean[dow];
      labelTd.style.fontWeight = "800";
      row.appendChild(labelTd);

      const makeTimeInput = (key, cIdx) => {
        const td = document.createElement("td");
        const inp = document.createElement("input");
        inp.type = "time";
        inp.value = state.schedule[dow][key] || "";
        inp.dataset.grid = "schedule";
        inp.dataset.r = String(rIdx);
        inp.dataset.c = String(cIdx);
        inp.addEventListener("input", () => {
          state.schedule[dow][key] = inp.value;
          const hBox = row.querySelector('[data-role="dayHours"]');
          if (hBox) hBox.value = String(calcScheduleHours(dow));
          computeAndRender();
          scheduleSaveDebounced();
        });
        td.appendChild(inp);
        return td;
      };

      row.appendChild(makeTimeInput("s1", 1));
      row.appendChild(makeTimeInput("e1", 2));
      row.appendChild(makeTimeInput("s2", 3));
      row.appendChild(makeTimeInput("e2", 4));

      const hoursTd = document.createElement("td");
      const hoursInp = document.createElement("input");
      hoursInp.type = "number";
      hoursInp.step = "0.1";
      hoursInp.className = "numeric";
      hoursInp.value = String(calcScheduleHours(dow));
      hoursInp.readOnly = true;
      hoursInp.dataset.role = "dayHours";
      hoursInp.tabIndex = -1;
      hoursTd.appendChild(hoursInp);
      row.appendChild(hoursTd);

      const memoTd = document.createElement("td");
      const memoInp = document.createElement("input");
      memoInp.type = "text";
      memoInp.placeholder = "예) 오전/오후, 행사 등";
      memoInp.value = state.schedule[dow].memo || "";
      memoInp.dataset.grid = "schedule";
      memoInp.dataset.r = String(rIdx);
      memoInp.dataset.c = "6";
      memoInp.addEventListener("input", () => {
        state.schedule[dow].memo = memoInp.value;
        scheduleSaveDebounced();
      });
      memoTd.appendChild(memoInp);
      row.appendChild(memoTd);

      tbody.appendChild(row);
    });

    rebuildGridMap("schedule");
  }

  function calcScheduleHours(dow) {
    const s = state.schedule[dow];
    const mins = diffMinutes(s.s1, s.e1) + diffMinutes(s.s2, s.e2);
    return minutesToHours(mins);
  }

  function calcWeeklyContractHours() {
    // 1주에 실제 근로하는 시간(모든 요일 합)
    let total = 0;
    for (let dow = 0; dow < 7; dow++) total += calcScheduleHours(dow);
    return Math.round(total * 100) / 100;
  }

  // ---------- calendar rendering ----------
  function getMonthParts(ym) {
    const m = /^(\d{4})-(\d{2})$/.exec(ym);
    if (!m) return null;
    return { y: parseInt(m[1], 10), m: parseInt(m[2], 10) };
  }

  function makeMonthWeeks(ym) {
    const p = getMonthParts(ym);
    if (!p) return [];
    const { y, m } = p;
    const first = new Date(y, m - 1, 1);
    const last = new Date(y, m, 0);
    const firstDow = first.getDay();
    const daysInMonth = last.getDate();

    const weeks = Array.from({ length: 6 }, () => Array(7).fill(null));
    let day = 1;
    let w = 0;
    let d = firstDow;
    while (day <= daysInMonth) {
      weeks[w][d] = new Date(y, m - 1, day);
      day += 1;
      d += 1;
      if (d === 7) {
        d = 0;
        w += 1;
      }
    }
    return weeks;
  }

  function ensureAttendanceKey(dateISO) {
    if (!state.attendance[dateISO]) {
      state.attendance[dateISO] = { on: false, manualHours: null };
    }
    return state.attendance[dateISO];
  }

  function renderCalendar() {
    const tbody = $("calendarTbody");
    tbody.innerHTML = "";

    const ym = state.payMonth || defaultPayMonth();
    const weeks = makeMonthWeeks(ym);
    const showManual = !!state.ui.showManualDailyHours;

    weeks.forEach((week, rIdx) => {
      const tr = document.createElement("tr");

      week.forEach((dateObj, cIdx) => {
        const td = document.createElement("td");
        if (!dateObj) {
          td.innerHTML = `<div class="cal-day"><span class="cal-num" style="color:#bbb;">-</span><span class="cal-sub"></span></div>`;
          td.style.background = "#fafafa";
          tr.appendChild(td);
          return;
        }

        const dateISO = toISODate(dateObj);
        const dow = dateObj.getDay();
        const a = ensureAttendanceKey(dateISO);

        const top = document.createElement("div");
        top.className = "cal-day";

        const num = document.createElement("div");
        num.className = "cal-num";
        num.textContent = String(dateObj.getDate());

        const sub = document.createElement("div");
        sub.className = "cal-sub";
        sub.textContent = weekdayKorean[dow];

        top.appendChild(num);
        top.appendChild(sub);

        const controls = document.createElement("div");
        controls.className = "cal-controls";

        // 출근 토글 버튼(셀 클릭으로도 토글)
        const toggleBtn = document.createElement("button");
        toggleBtn.type = "button";
        toggleBtn.className = "att-toggle" + (a.on ? " on" : "");
        toggleBtn.textContent = a.on ? "출근" : "미출근";
        toggleBtn.dataset.grid = "calendar";
        toggleBtn.dataset.r = String(rIdx);
        toggleBtn.dataset.c = String(cIdx);
        toggleBtn.dataset.date = dateISO;

        const applyToggleUI = () => {
          toggleBtn.classList.toggle("on", !!a.on);
          toggleBtn.textContent = a.on ? "출근" : "미출근";
          td.classList.toggle("day-on", !!a.on);
        };

        const toggleAttendance = () => {
          a.on = !a.on;
          ensureAttendanceKey(dateISO).on = a.on;
          applyToggleUI();
          computeAndRender();
          scheduleSaveDebounced();
        };

        toggleBtn.addEventListener("click", (ev) => {
          ev.preventDefault();
          toggleAttendance();
        });

        // 셀 아무 곳이나 클릭하면 토글(숫자/요일/여백)
        td.addEventListener("click", (ev) => {
          const target = ev.target;
          // input(수동시간)이나 버튼 클릭은 각각의 기본 동작을 존중
          if (target && target.closest && target.closest("input, button, select, textarea, a, label")) return;
          toggleAttendance();
        });

        controls.appendChild(toggleBtn);

        // 최초 UI 반영
        applyToggleUI();

        const hoursInp = document.createElement("input");
        hoursInp.type = "number";
        hoursInp.step = "0.1";
        hoursInp.placeholder = "시간";
        hoursInp.className = "numeric";
        hoursInp.style.display = showManual ? "inline-block" : "none";
        hoursInp.value = (a.manualHours == null || Number.isNaN(a.manualHours)) ? "" : String(a.manualHours);
        hoursInp.dataset.grid = "calendarHours";
        hoursInp.dataset.r = String(rIdx);
        hoursInp.dataset.c = String(cIdx);
        hoursInp.dataset.date = dateISO;

        hoursInp.addEventListener("input", () => {
          const val = hoursInp.value.trim();
          if (!val) {
            ensureAttendanceKey(dateISO).manualHours = null;
          } else {
            ensureAttendanceKey(dateISO).manualHours = parseNumber(val, null);
          }
          computeAndRender();
          scheduleSaveDebounced();
        });

        controls.appendChild(hoursInp);

        td.appendChild(top);
        td.appendChild(controls);

        const bottom = document.createElement("div");
        bottom.className = "muted";
        bottom.style.marginTop = "8px";
        bottom.style.fontSize = "0.92rem";
        bottom.dataset.role = "hoursHint";
        td.appendChild(bottom);

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    rebuildGridMap("calendar");
    rebuildGridMap("calendarHours");
  }

  // ---------- compute & render ----------
  function defaultPayMonth() {
    const now = new Date();
    return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;
  }

  function defaultPayDateForMonth(ym) {
    const p = getMonthParts(ym);
    if (!p) return "";
    const last = new Date(p.y, p.m, 0);
    return toISODate(last);
  }

  function sumWorkHoursForMonth(ym) {
    const weeks = makeMonthWeeks(ym);
    let days = 0;
    let hours = 0;

    weeks.flat().forEach((d) => {
      if (!d) return;
      const iso = toISODate(d);
      const a = state.attendance[iso];
      if (!a || !a.on) return;
      days += 1;

      const dow = d.getDay();
      const schedH = calcScheduleHours(dow);
      const used = (a.manualHours != null && Number.isFinite(a.manualHours)) ? a.manualHours : schedH;
      hours += parseNumber(used, 0);
    });

    // 소수점 오차 보정
    hours = Math.round(hours * 100) / 100;
    return { days, hours };
  }

  function autoInsuranceAmounts(gross) {
    // gross: 원(정수/실수 가능)
    const trunc = !!state.truncate10won;

    const calcHealth = () => {
      const raw = gross * state.healthRate;
      return trunc ? floor10(raw) : Math.round(raw);
    };

    const calcLt = (healthAmt) => {
      const raw = healthAmt * state.ltCareRate;
      return trunc ? floor10(raw) : Math.round(raw);
    };

    const calcEmp = () => {
      const raw = gross * state.empRate;
      return trunc ? floor10(raw) : Math.round(raw);
    };

    const health = calcHealth();
    const lt = calcLt(health);
    const emp = calcEmp();
    return { health, lt, emp };
  }

  function computeAndRender() {
    // schedule 재계산은 renderSchedule에서 readonly로 표시되지만, KPI에 반영
    const weeklyHours = calcWeeklyContractHours();
    $("contractWeeklyHoursPill").textContent = `주간 계약시간: ${weeklyHours}시간`;
    $("contractWeeklyHoursPill").className = "pill" + (weeklyHours > 0 ? " ok" : "");

    const ym = state.payMonth || defaultPayMonth();
    const { days, hours } = sumWorkHoursForMonth(ym);

    $("workDaysText").textContent = String(days);
    $("workHoursText").textContent = String(hours);

    // Gross
    const hourly = parseNumber(state.hourlyRate, 0);
    const gross = Math.round(hourly * hours); // 원 단위
    $("grossPayText").textContent = fmtWon(gross);

    // Auto insurance calc (only if toggle ON and not manual)
    const auto = autoInsuranceAmounts(gross);

    if (state.deductHealthOn && !state.manual.healthAmt) {
      state.healthAmt = auto.health;
      $("healthAmt").value = String(state.healthAmt);
    }
    if (state.deductLtCareOn && !state.manual.ltCareAmt) {
      state.ltCareAmt = auto.lt;
      $("ltCareAmt").value = String(state.ltCareAmt);
    }
    if (state.deductEmpOn && !state.manual.empAmt) {
      state.empAmt = auto.emp;
      $("empAmt").value = String(state.empAmt);
    }

    // Deductions total
    const healthAmt = state.deductHealthOn ? parseNumber(state.healthAmt, 0) : 0;
    const ltAmt = state.deductLtCareOn ? parseNumber(state.ltCareAmt, 0) : 0;
    const empAmt = state.deductEmpOn ? parseNumber(state.empAmt, 0) : 0;

    const incomeTax = parseNumber(state.incomeTax, 0);
    const localTax = parseNumber(state.localTax, 0);
    const pension = parseNumber(state.pension, 0);
    const otherDeduct = parseNumber(state.otherDeduct, 0);

    const deductTotal = Math.max(0, Math.round(healthAmt + ltAmt + empAmt + incomeTax + localTax + pension + otherDeduct));
    $("deductTotalText").textContent = fmtWon(deductTotal);

    const net = gross - deductTotal;
    $("netPayText").textContent = fmtWon(net);

    // Update per-day hour hint on calendar
    updateCalendarHints(ym);

    // Payslip preview
    renderPayslipPreview({
      ym,
      gross,
      deductTotal,
      net,
      hours,
      days,
      insurance: { healthAmt, ltAmt, empAmt },
      taxes: { incomeTax, localTax, pension, otherDeduct }
    });
  }

  function updateCalendarHints(ym) {
    const showManual = !!state.ui.showManualDailyHours;
    // DOM 전체를 순회하며 date를 가진 요소를 찾는 방식(간단/안전)
    document.querySelectorAll('[data-grid="calendar"][data-date]').forEach((cb) => {
      const dateISO = cb.dataset.date;
      const d = new Date(dateISO + "T00:00:00");
      if (Number.isNaN(d.getTime())) return;
      const dow = d.getDay();
      const schedH = calcScheduleHours(dow);
      const a = state.attendance[dateISO] || { on: false, manualHours: null };
      const used = (a.manualHours != null && Number.isFinite(a.manualHours)) ? a.manualHours : schedH;

      const cell = cb.closest("td");
      const hint = cell && cell.querySelector('[data-role="hoursHint"]');
      if (!hint) return;

      if (!a.on) {
        hint.textContent = showManual ? `근로시간: ${schedH}h (미출근)` : `근로시간: ${schedH}h`;
        hint.style.color = "#666";
      } else {
        hint.textContent = (a.manualHours != null && Number.isFinite(a.manualHours))
          ? `근로시간: ${used}h (수동)`
          : `근로시간: ${used}h`;
        hint.style.color = "#047857";
      }
    });
  }

  function renderPayslipPreview(calc) {
    const { ym, gross, deductTotal, net, hours, insurance, taxes } = calc;

    const title = ymToTitle(ym) + " 임금명세서";
    const payDate = state.payDate || defaultPayDateForMonth(ym);
    const school = state.schoolName || "";
    const worker = state.workerName || "";
    const jobTitle = state.jobTitle || "출입문개폐전담원";
    const birth = state.birthDate || "";
    const hire = state.hireDate || "";
    const hourly = parseNumber(state.hourlyRate, 0);

    const rowsPay = [
      { cat: "매월 지급", item: "기본급", formula: `시급 ${fmtWon(hourly)}원 × ${hours}시간 =`, amount: gross },
      { cat: "", item: "근속수당", formula: "", amount: "" },
      { cat: "", item: "정액급식비", formula: "", amount: "" },
      { cat: "", item: "위험근무수당", formula: "", amount: "" },
      { cat: "", item: "면허가산수당", formula: "", amount: "" },
      { cat: "", item: "특수업무수당", formula: "", amount: "" },
      { cat: "", item: "급식운영수당", formula: "", amount: "" },
      { cat: "", item: "가족수당", formula: "", amount: "" },
      { cat: "부정기 지급", item: "명절휴가비", formula: "", amount: "" },
      { cat: "부정기 지급", item: "명절휴가비", formula: "", amount: "" },
    ];

    const rowsDeduct = [
      { item: "소득세", amount: taxes.incomeTax || 0 },
      { item: "주민세", amount: taxes.localTax || 0 },
      { item: "건강보험", amount: insurance.healthAmt || 0 },
      { item: "장기요양보험", amount: insurance.ltAmt || 0 },
      { item: "국민연금", amount: taxes.pension || 0 },
      { item: "고용보험", amount: insurance.empAmt || 0 },
      { item: "기타", amount: taxes.otherDeduct || 0 },
    ];

    const el = $("payslipPreview");

    // HTML table (7 columns)
    const html = `
      <table class="payslip">
        <colgroup>
          <col style="width: 8%">
          <col style="width: 14%">
          <col style="width: 30%">
          <col style="width: 12%">
          <col style="width: 12%">
          <col style="width: 12%">
          <col style="width: 12%">
        </colgroup>

        <tr>
          <td class="title" colspan="7">${escapeHtml(title)}</td>
        </tr>

        <tr>
          <td class="h center" colspan="1">소속</td>
          <td colspan="2">${escapeHtml(school)}</td>
          <td class="h center" colspan="1">지급일</td>
          <td colspan="3">${escapeHtml(payDate)}</td>
        </tr>

        <tr>
          <td class="h center">성명</td>
          <td colspan="2">${escapeHtml(worker)}</td>
          <td class="h center">직종</td>
          <td colspan="3">${escapeHtml(jobTitle)}</td>
        </tr>

        <tr>
          <td class="h center">생년월일</td>
          <td colspan="2">${escapeHtml(birth)}</td>
          <td class="h center">최초임용일</td>
          <td colspan="3">${escapeHtml(hire)}</td>
        </tr>

        <tr>
          <td class="subhead center" colspan="4">급여내역</td>
          <td class="subhead center" colspan="3">공제내역</td>
        </tr>

        <tr>
          <td class="h center">구분</td>
          <td class="h center">임금항목</td>
          <td class="h center">산출식</td>
          <td class="h center">금액</td>
          <td class="h center" colspan="2">공제구분</td>
          <td class="h center">금액</td>
        </tr>

        ${rowsPay.map((p, idx) => `
          <tr>
            <td class="center">${idx === 0 ? "매월\n지급" : (idx === 8 ? "부정기\n지급" : "")}</td>
            <td>${escapeHtml(p.item)}</td>
            <td class="mono">${escapeHtml(p.formula || "")}</td>
            <td class="right">${p.amount === "" ? "" : fmtWon(p.amount)}</td>
            <td colspan="2">${escapeHtml(rowsDeduct[idx]?.item || "")}</td>
            <td class="right">${rowsDeduct[idx] ? fmtWon(rowsDeduct[idx].amount || 0) : ""}</td>
          </tr>
        `).join("")}

        <tr>
          <td class="h" colspan="3">급여총액 계 (A)</td>
          <td class="right"><strong>${fmtWon(gross)}</strong></td>
          <td class="h" colspan="2">공제액 계 (B)</td>
          <td class="right"><strong>${fmtWon(deductTotal)}</strong></td>
        </tr>

        <tr>
          <td class="h" colspan="6">실수령액 (A-B)</td>
          <td class="right"><strong>${fmtWon(net)}</strong></td>
        </tr>
      </table>

      <div class="hint muter" style="margin-top: 10px;">
        ·급여총액=시급×유급근로시간 · 사회보험료는 “공제” 선택 시 요율 자동 산출(수기 가능)
      </div>
    `;
    el.innerHTML = html;
  }

  function ymToTitle(ym) {
    const p = getMonthParts(ym);
    if (!p) return ym;
    return `${p.y}년 ${p.m}월`;
  }

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (ch) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[ch]));
  }

  // ---------- localStorage ----------
  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);

      // shallow merge
      Object.assign(state, parsed);

      // schedule/attendance 구조 보정
      if (!state.schedule) state.schedule = defaultSchedule();
      for (let i = 0; i < 7; i++) {
        if (!state.schedule[i]) state.schedule[i] = { s1: "", e1: "", s2: "", e2: "", memo: "" };
      }
      if (!state.attendance) state.attendance = {};
      if (!state.manual) state.manual = { healthAmt: false, ltCareAmt: false, empAmt: false };
      if (!state.ui) state.ui = { showManualDailyHours: false };

      $("saveStateText").textContent = "불러옴";
    } catch (e) {
      console.warn("Failed to load state:", e);
    }
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      $("saveStateText").textContent = new Date().toLocaleString("ko-KR");
    } catch (e) {
      console.warn("Failed to save state:", e);
      $("saveStateText").textContent = "저장 실패";
    }
  }

  let saveT = null;
  function scheduleSaveDebounced() {
    if (saveT) clearTimeout(saveT);
    saveT = setTimeout(() => saveState(), 400);
  }

  // ---------- wire inputs ----------
  function bindInputs() {
    $("payMonth").addEventListener("change", () => {
      state.payMonth = $("payMonth").value;
      if (!state.payDate) state.payDate = defaultPayDateForMonth(state.payMonth);
      $("payDate").value = state.payDate;
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("payDate").addEventListener("change", () => {
      state.payDate = $("payDate").value;
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("schoolName").addEventListener("input", () => { state.schoolName = $("schoolName").value; computeAndRender(); scheduleSaveDebounced(); });
    $("workerName").addEventListener("input", () => { state.workerName = $("workerName").value; computeAndRender(); scheduleSaveDebounced(); });
    $("jobTitle").addEventListener("input", () => { state.jobTitle = $("jobTitle").value; computeAndRender(); scheduleSaveDebounced(); });
    $("birthDate").addEventListener("change", () => { state.birthDate = $("birthDate").value; computeAndRender(); scheduleSaveDebounced(); });
    $("hireDate").addEventListener("change", () => { state.hireDate = $("hireDate").value; computeAndRender(); scheduleSaveDebounced(); });

    $("hourlyRate").addEventListener("input", () => {
      state.hourlyRate = parseNumber($("hourlyRate").value, 0);
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("preset2025Btn").addEventListener("click", () => {
      state.hourlyRate = 11200;
      $("hourlyRate").value = "11200";
      computeAndRender();
      scheduleSaveDebounced();
    });
    $("preset2026Btn").addEventListener("click", () => {
      state.hourlyRate = 11500;
      $("hourlyRate").value = "11500";
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("loadExampleScheduleBtn").addEventListener("click", () => {
      loadExampleSchedule();
      renderSchedule();
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("showManualDailyHours").addEventListener("change", () => {
      state.ui.showManualDailyHours = $("showManualDailyHours").checked;
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("fillByScheduleBtn").addEventListener("click", () => {
      fillCalendarBySchedule();
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("clearCalendarBtn").addEventListener("click", () => {
      clearCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    $("saveNowBtn").addEventListener("click", () => saveState());

    $("resetAllBtn").addEventListener("click", () => {
      if (!confirm("전체 입력값을 초기화할까요? (브라우저 저장값 포함)")) return;
      localStorage.removeItem(STORAGE_KEY);
      location.reload();
    });

    // deductions toggles
    $("truncate10won").addEventListener("change", () => { state.truncate10won = $("truncate10won").checked; computeAndRender(); scheduleSaveDebounced(); });

    $("deductHealthOn").addEventListener("change", () => { state.deductHealthOn = $("deductHealthOn").checked; computeAndRender(); scheduleSaveDebounced(); });
    $("healthRate").addEventListener("input", () => { state.healthRate = parseNumber($("healthRate").value, 0); state.manual.healthAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    $("healthAmt").addEventListener("input", () => { state.healthAmt = parseNumber($("healthAmt").value, 0); state.manual.healthAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    $("healthAutoBtn").addEventListener("click", () => { state.manual.healthAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    $("deductLtCareOn").addEventListener("change", () => { state.deductLtCareOn = $("deductLtCareOn").checked; computeAndRender(); scheduleSaveDebounced(); });
    $("ltCareRate").addEventListener("input", () => { state.ltCareRate = parseNumber($("ltCareRate").value, 0); state.manual.ltCareAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    $("ltCareAmt").addEventListener("input", () => { state.ltCareAmt = parseNumber($("ltCareAmt").value, 0); state.manual.ltCareAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    $("ltCareAutoBtn").addEventListener("click", () => { state.manual.ltCareAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    $("deductEmpOn").addEventListener("change", () => { state.deductEmpOn = $("deductEmpOn").checked; computeAndRender(); scheduleSaveDebounced(); });
    $("empRate").addEventListener("input", () => { state.empRate = parseNumber($("empRate").value, 0); state.manual.empAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    $("empAmt").addEventListener("input", () => { state.empAmt = parseNumber($("empAmt").value, 0); state.manual.empAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    $("empAutoBtn").addEventListener("click", () => { state.manual.empAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    $("incomeTax").addEventListener("input", () => { state.incomeTax = parseNumber($("incomeTax").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    $("localTax").addEventListener("input", () => { state.localTax = parseNumber($("localTax").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    $("pension").addEventListener("input", () => { state.pension = parseNumber($("pension").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    $("otherDeduct").addEventListener("input", () => { state.otherDeduct = parseNumber($("otherDeduct").value, 0); computeAndRender(); scheduleSaveDebounced(); });

    // file actions
    $("downloadTemplateBtn").addEventListener("click", () => downloadTemplateXlsx());
    $("uploadXlsx").addEventListener("change", (e) => uploadTemplateXlsx(e.target.files?.[0]));
    $("exportXlsxBtn").addEventListener("click", () => exportResultXlsx());
    $("exportPdfBtn").addEventListener("click", () => exportPayslipPdf());
  }

  function syncUIFromState() {
    $("payMonth").value = state.payMonth || defaultPayMonth();
    $("payDate").value = state.payDate || defaultPayDateForMonth($("payMonth").value);

    $("schoolName").value = state.schoolName || "";
    $("workerName").value = state.workerName || "";
    $("jobTitle").value = state.jobTitle || "출입문개폐전담원";
    $("birthDate").value = state.birthDate || "";
    $("hireDate").value = state.hireDate || "";

    $("hourlyRate").value = String(parseNumber(state.hourlyRate, 0));

    $("showManualDailyHours").checked = !!state.ui.showManualDailyHours;

    $("truncate10won").checked = !!state.truncate10won;

    $("deductHealthOn").checked = !!state.deductHealthOn;
    $("healthRate").value = String(parseNumber(state.healthRate, 0));
    $("healthAmt").value = String(parseNumber(state.healthAmt, 0));

    $("deductLtCareOn").checked = !!state.deductLtCareOn;
    $("ltCareRate").value = String(parseNumber(state.ltCareRate, 0));
    $("ltCareAmt").value = String(parseNumber(state.ltCareAmt, 0));

    $("deductEmpOn").checked = !!state.deductEmpOn;
    $("empRate").value = String(parseNumber(state.empRate, 0));
    $("empAmt").value = String(parseNumber(state.empAmt, 0));

    $("incomeTax").value = String(parseNumber(state.incomeTax, 0));
    $("localTax").value = String(parseNumber(state.localTax, 0));
    $("pension").value = String(parseNumber(state.pension, 0));
    $("otherDeduct").value = String(parseNumber(state.otherDeduct, 0));
  }

  function loadExampleSchedule() {
    // 월~목: 07:10~08:40, 16:40~18:10 (총 3시간)
    // 금: 07:40~08:40, 16:40~17:40 (총 2시간)
    const s = defaultSchedule();
    [1,2,3,4].forEach((dow) => {
      s[dow] = { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10", memo: "월~목" };
    });
    s[5] = { s1: "07:40", e1: "08:40", s2: "16:40", e2: "17:40", memo: "금" };
    s[0] = { s1: "", e1: "", s2: "", e2: "", memo: "" }; // 일
    s[6] = { s1: "", e1: "", s2: "", e2: "", memo: "" }; // 토
    state.schedule = s;
  }

  function fillCalendarBySchedule() {
    const ym = state.payMonth || defaultPayMonth();
    const weeks = makeMonthWeeks(ym);

    weeks.flat().forEach((d) => {
      if (!d) return;
      const iso = toISODate(d);
      const dow = d.getDay();
      const h = calcScheduleHours(dow);
      const a = ensureAttendanceKey(iso);
      a.on = h > 0;
      // 수동시간은 유지(원하면 직접 지우기)
    });

    // UI 반영: 체크만 갱신(전체 rerender로 간단 처리)
    renderCalendar();
  }

  function clearCalendar() {
    const ym = state.payMonth || defaultPayMonth();
    const weeks = makeMonthWeeks(ym);
    weeks.flat().forEach((d) => {
      if (!d) return;
      const iso = toISODate(d);
      const a = ensureAttendanceKey(iso);
      a.on = false;
      a.manualHours = null;
    });
    renderCalendar();
  }

  // ---------- keyboard navigation (excel-like) ----------
  const gridMaps = new Map(); // gridName -> {cells: Map("r,c"->el), maxR, maxC}

  function rebuildGridMap(gridName) {
    const els = Array.from(document.querySelectorAll(`[data-grid="${gridName}"]`))
      .filter((el) => !el.disabled && el.offsetParent !== null);
    const cells = new Map();
    let maxR = 0, maxC = 0;

    els.forEach((el) => {
      const r = parseInt(el.dataset.r || "0", 10);
      const c = parseInt(el.dataset.c || "0", 10);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
      cells.set(`${r},${c}`, el);
    });

    gridMaps.set(gridName, { cells, maxR, maxC });
  }

  function findNextInGrid(gridName, r, c, dr, dc) {
    const map = gridMaps.get(gridName);
    if (!map) return null;

    let nr = r + dr;
    let nc = c + dc;

    // 스캔하면서 존재하는 셀 찾기
    for (let i = 0; i < 100; i++) {
      const key = `${nr},${nc}`;
      const el = map.cells.get(key);
      if (el) return el;
      nr += dr;
      nc += dc;
    }
    return null;
  }

  function setupGridNavigation() {
    document.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLElement)) return;
      const gridName = t.dataset && t.dataset.grid;
      if (!gridName) return;

      // 입력 중 단축키는 제외
      if (e.altKey || e.metaKey || e.ctrlKey) return;

      const r = parseInt(t.dataset.r || "0", 10);
      const c = parseInt(t.dataset.c || "0", 10);

      // Space: checkbox / 출근 토글버튼
      if (e.key === " ") {
        // checkbox
        if ((t instanceof HTMLInputElement) && t.type === "checkbox") {
          e.preventDefault();
          t.checked = !t.checked;
          t.dispatchEvent(new Event("change", { bubbles: true }));
          return;
        }
        // calendar toggle button
        if ((t instanceof HTMLButtonElement) && t.classList.contains("att-toggle")) {
          e.preventDefault();
          t.click();
          return;
        }
      }

      const nav = (dr, dc) => {
        const nxt = findNextInGrid(gridName, r, c, dr, dc);
        if (nxt) {
          e.preventDefault();
          nxt.focus();
        }
      };

      switch (e.key) {
        case "ArrowUp": return nav(-1, 0);
        case "ArrowDown": return nav(1, 0);
        case "ArrowLeft": return nav(0, -1);
        case "ArrowRight": return nav(0, 1);
        case "Enter": return nav(1, 0);
        default: return;
      }
    });

    // focus style
    document.addEventListener("focusin", (e) => {
      const el = e.target;
      if (!(el instanceof HTMLElement)) return;
      if (!el.dataset || !el.dataset.grid) return;
      el.classList.add("cell-focus");
    });

    document.addEventListener("focusout", (e) => {
      const el = e.target;
      if (!(el instanceof HTMLElement)) return;
      el.classList.remove("cell-focus");
    });
  }

  // ---------- Excel I/O ----------
  function downloadTemplateXlsx() {
    const libs = isLibReady();
    if (!libs.okXlsx) {
      alert("XLSX 라이브러리를 불러오지 못했습니다. (인터넷 연결을 확인)");
      return;
    }

    const ym = state.payMonth || defaultPayMonth();
    const p = getMonthParts(ym);
    const wb = XLSX.utils.book_new();

    // 1) 입력 sheet (key-value)
    const kv = [
      ["key", "value"],
      ["payMonth", ym],
      ["payDate", state.payDate || defaultPayDateForMonth(ym)],
      ["schoolName", state.schoolName || ""],
      ["workerName", state.workerName || ""],
      ["jobTitle", state.jobTitle || "출입문개폐전담원"],
      ["birthDate", state.birthDate || ""],
      ["hireDate", state.hireDate || ""],
      ["hourlyRate", state.hourlyRate],
      ["truncate10won", state.truncate10won ? "TRUE" : "FALSE"],
      ["deductHealthOn", state.deductHealthOn ? "TRUE" : "FALSE"],
      ["healthRate", state.healthRate],
      ["healthAmt", state.healthAmt],
      ["deductLtCareOn", state.deductLtCareOn ? "TRUE" : "FALSE"],
      ["ltCareRate", state.ltCareRate],
      ["ltCareAmt", state.ltCareAmt],
      ["deductEmpOn", state.deductEmpOn ? "TRUE" : "FALSE"],
      ["empRate", state.empRate],
      ["empAmt", state.empAmt],
      ["incomeTax", state.incomeTax],
      ["localTax", state.localTax],
      ["pension", state.pension],
      ["otherDeduct", state.otherDeduct],
    ];
    const wsInput = XLSX.utils.aoa_to_sheet(kv);
    wsInput["!cols"] = [{ wch: 22 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, wsInput, "입력");

    // 2) 근무시간 sheet
    const sh = [
      ["요일", "시간대1(시작)", "시간대1(종료)", "시간대2(시작)", "시간대2(종료)", "일 근로시간", "비고"],
    ];
    weekdayOrder.forEach((dow) => {
      const s = state.schedule[dow];
      sh.push([
        weekdayKorean[dow],
        s.s1 || "", s.e1 || "", s.s2 || "", s.e2 || "",
        calcScheduleHours(dow),
        s.memo || "",
      ]);
    });
    const wsSched = XLSX.utils.aoa_to_sheet(sh);
    wsSched["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsSched, "근무시간");

    // 3) 출근내역(선택) sheet - 날짜별 O 및 수동시간
    const weeks = makeMonthWeeks(ym);
    const att = [["date", "weekday", "on(O)", "manualHours(선택)"]];
    weeks.flat().forEach((d) => {
      if (!d) return;
      const iso = toISODate(d);
      const a = state.attendance[iso] || { on: false, manualHours: null };
      att.push([iso, weekdayKorean[d.getDay()], a.on ? "O" : "", a.manualHours == null ? "" : a.manualHours]);
    });
    const wsAtt = XLSX.utils.aoa_to_sheet(att);
    wsAtt["!cols"] = [{ wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsAtt, "출근내역(선택)");

    const fname = `출입문개폐전담원_업로드서식_${p ? (p.y + pad2(p.m)) : ym}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  async function uploadTemplateXlsx(file) {
    const libs = isLibReady();
    if (!libs.okXlsx) {
      alert("XLSX 라이브러리를 불러오지 못했습니다. (인터넷 연결을 확인)");
      return;
    }
    if (!file) return;

    $("uploadStatus").textContent = "읽는 중...";
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      if (wb.Sheets["입력"]) {
        applyKvSheet(wb.Sheets["입력"]);
      }
      if (wb.Sheets["근무시간"]) {
        applyScheduleSheet(wb.Sheets["근무시간"]);
      }
      if (wb.Sheets["출근내역(선택)"]) {
        applyAttendanceSheet(wb.Sheets["출근내역(선택)"]);
      }

      // 다른 양식(기존 월별 파일) 추정 파싱
      if (!wb.Sheets["입력"]) {
        tryParseLegacyWorkbook(wb);
      }

      // UI 반영
      syncUIFromState();
      renderSchedule();
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();

      $("uploadStatus").textContent = `업로드 완료: ${file.name}`;
    } catch (err) {
      console.error(err);
      $("uploadStatus").textContent = "업로드 실패: 파일 형식/내용을 확인하세요.";
    }
  }

  function applyKvSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) || [];
    const kv = new Map();
    rows.slice(1).forEach((r) => {
      if (!r || r.length < 2) return;
      const k = String(r[0] ?? "").trim();
      const v = r[1];
      if (!k) return;
      kv.set(k, v);
    });

    const b = (x) => String(x).toUpperCase() === "TRUE";

    if (kv.has("payMonth")) state.payMonth = String(kv.get("payMonth") || "");
    if (kv.has("payDate")) state.payDate = String(kv.get("payDate") || "");
    if (kv.has("schoolName")) state.schoolName = String(kv.get("schoolName") || "");
    if (kv.has("workerName")) state.workerName = String(kv.get("workerName") || "");
    if (kv.has("jobTitle")) state.jobTitle = String(kv.get("jobTitle") || "");
    if (kv.has("birthDate")) state.birthDate = String(kv.get("birthDate") || "");
    if (kv.has("hireDate")) state.hireDate = String(kv.get("hireDate") || "");
    if (kv.has("hourlyRate")) state.hourlyRate = parseNumber(kv.get("hourlyRate"), 0);

    if (kv.has("truncate10won")) state.truncate10won = b(kv.get("truncate10won"));
    if (kv.has("deductHealthOn")) state.deductHealthOn = b(kv.get("deductHealthOn"));
    if (kv.has("healthRate")) state.healthRate = parseNumber(kv.get("healthRate"), state.healthRate);
    if (kv.has("healthAmt")) state.healthAmt = parseNumber(kv.get("healthAmt"), 0);
    if (kv.has("deductLtCareOn")) state.deductLtCareOn = b(kv.get("deductLtCareOn"));
    if (kv.has("ltCareRate")) state.ltCareRate = parseNumber(kv.get("ltCareRate"), state.ltCareRate);
    if (kv.has("ltCareAmt")) state.ltCareAmt = parseNumber(kv.get("ltCareAmt"), 0);
    if (kv.has("deductEmpOn")) state.deductEmpOn = b(kv.get("deductEmpOn"));
    if (kv.has("empRate")) state.empRate = parseNumber(kv.get("empRate"), state.empRate);
    if (kv.has("empAmt")) state.empAmt = parseNumber(kv.get("empAmt"), 0);

    if (kv.has("incomeTax")) state.incomeTax = parseNumber(kv.get("incomeTax"), 0);
    if (kv.has("localTax")) state.localTax = parseNumber(kv.get("localTax"), 0);
    if (kv.has("pension")) state.pension = parseNumber(kv.get("pension"), 0);
    if (kv.has("otherDeduct")) state.otherDeduct = parseNumber(kv.get("otherDeduct"), 0);

    // 업로드는 "수기 덮어쓰기" 취급
    state.manual.healthAmt = true;
    state.manual.ltCareAmt = true;
    state.manual.empAmt = true;
  }

  function applyScheduleSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) || [];
    // header: 요일, s1,e1,s2,e2,hours,memo
    const mapDow = { "일":0, "월":1, "화":2, "수":3, "목":4, "금":5, "토":6 };

    rows.slice(1).forEach((r) => {
      if (!r || r.length < 1) return;
      const day = String(r[0] ?? "").trim();
      if (!mapDow.hasOwnProperty(day)) return;
      const dow = mapDow[day];

      state.schedule[dow] = {
        s1: String(r[1] ?? "").trim(),
        e1: String(r[2] ?? "").trim(),
        s2: String(r[3] ?? "").trim(),
        e2: String(r[4] ?? "").trim(),
        memo: String(r[6] ?? "").trim(),
      };
    });
  }

  function applyAttendanceSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) || [];
    rows.slice(1).forEach((r) => {
      if (!r || r.length < 3) return;
      const dateISO = String(r[0] ?? "").trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) return;
      const on = String(r[2] ?? "").trim().toUpperCase() === "O";
      const mh = r[3] === "" || r[3] == null ? null : parseNumber(r[3], null);
      state.attendance[dateISO] = { on, manualHours: mh };
    });
  }

  function tryParseLegacyWorkbook(wb) {
    // 기존 월별 양식(예: '2025.3월 임금명세서' + '2025.3월 출근내역')을 최소한으로 파싱
    const names = wb.SheetNames || [];
    const paySheetName = names.find((n) => n.includes("임금명세서"));
    if (!paySheetName) return;

    const payWs = wb.Sheets[paySheetName];
    const get = (addr) => payWs[addr] ? payWs[addr].v : "";

    // 기본 정보
    state.schoolName = String(get("D3") || "");
    const payDate = get("G3");
    state.payDate = excelDateToISO(payDate) || "";
    state.workerName = String(get("D4") || "");
    state.jobTitle = String(get("G4") || "출입문개폐전담원");
    state.birthDate = excelDateToISO(get("D5")) || "";
    state.hireDate = excelDateToISO(get("G5")) || "";

    // 시급 추출: D9 문자열에서 '시급 11,200원' 형태
    const d9 = String(get("D9") || "");
    const m = /시급\s*([0-9,]+)원/.exec(d9.replace(/\s+/g, " "));
    if (m) state.hourlyRate = parseNumber(m[1].replace(/,/g, ""), state.hourlyRate);

    // 지급월 추정: 시트명 '2025.3월 ...'
    const ymMatch = /(\d{4})\.(\d{1,2})월/.exec(paySheetName);
    if (ymMatch) {
      state.payMonth = `${ymMatch[1]}-${pad2(parseInt(ymMatch[2], 10))}`;
    }

    // 출근내역 파싱
    const attSheetName = names.find((n) => n.includes("출근내역") && (ymMatch ? n.includes(`${ymMatch[1]}.${ymMatch[2]}월`) : true));
    if (!attSheetName) return;
    const attWs = wb.Sheets[attSheetName];

    // 달력 영역: 날짜는 B5:H? / 출근표시는 B6:H? (우리는 6주*2행 구조)
    // 기존 양식은 다를 수 있으므로 "날짜 셀 아래 한 줄에 O가 있는" 패턴을 찾아 추출한다.
    const cells = Object.keys(attWs).filter((k) => !k.startsWith("!"));

    // 날짜 셀로 보이는 값(Excel date number)들을 찾고, 바로 아래행 같은 열의 값이 "O"면 출근으로 판단
    const dateCandidates = cells.filter((addr) => {
      const v = attWs[addr]?.v;
      return typeof v === "number" && v > 40000 && v < 60000; // 2009~2064 정도 범위
    });

    dateCandidates.forEach((addr) => {
      const { c, r } = XLSX.utils.decode_cell(addr); // 0-index
      const below = XLSX.utils.encode_cell({ c, r: r + 1 });
      const mark = attWs[below]?.v;
      const dateISO = excelDateToISO(attWs[addr]?.v);
      if (!dateISO) return;
      ensureAttendanceKey(dateISO).on = String(mark || "").trim().toUpperCase() === "O";
    });
  }

  function excelDateToISO(v) {
    if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    if (typeof v === "number") {
      // Excel serial date -> JS date (SheetJS: 1900-based)
      const d = XLSX.SSF.parse_date_code(v);
      if (!d) return "";
      const js = new Date(d.y, d.m - 1, d.d);
      return toISODate(js);
    }
    // Date object in some cases
    if (v instanceof Date && !Number.isNaN(v.getTime())) return toISODate(v);
    return "";
  }

  function exportResultXlsx() {
    const libs = isLibReady();
    if (!libs.okXlsx) {
      alert("XLSX 라이브러리를 불러오지 못했습니다. (인터넷 연결을 확인)");
      return;
    }

    const ym = state.payMonth || defaultPayMonth();
    const p = getMonthParts(ym);
    const { days, hours } = sumWorkHoursForMonth(ym);
    const hourly = parseNumber(state.hourlyRate, 0);
    const gross = Math.round(hourly * hours);

    const healthAmt = state.deductHealthOn ? parseNumber(state.healthAmt, 0) : 0;
    const ltAmt = state.deductLtCareOn ? parseNumber(state.ltCareAmt, 0) : 0;
    const empAmt = state.deductEmpOn ? parseNumber(state.empAmt, 0) : 0;
    const incomeTax = parseNumber(state.incomeTax, 0);
    const localTax = parseNumber(state.localTax, 0);
    const pension = parseNumber(state.pension, 0);
    const otherDeduct = parseNumber(state.otherDeduct, 0);
    const deductTotal = Math.round(healthAmt + ltAmt + empAmt + incomeTax + localTax + pension + otherDeduct);
    const net = gross - deductTotal;

    const wb = XLSX.utils.book_new();

    const ws1 = buildSheet1Attendance(ym, { days, hours, gross });
    XLSX.utils.book_append_sheet(wb, ws1, "출근내역");

    const ws2 = buildSheet2Calc(ym, { days, hours, hourly, gross, healthAmt, ltAmt, empAmt, incomeTax, localTax, pension, otherDeduct, deductTotal, net });
    XLSX.utils.book_append_sheet(wb, ws2, "인건비 산정내역");

    const ws3 = buildSheet3Payslip(ym, { hours, hourly, gross, healthAmt, ltAmt, empAmt, incomeTax, localTax, pension, otherDeduct, deductTotal, net });
    XLSX.utils.book_append_sheet(wb, ws3, "임금명세서");

    const fname = `출입문개폐전담원_인건비_${p ? (p.y + pad2(p.m)) : ym}.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  function buildSheet1Attendance(ym, totals) {
    const title = `${ymToTitle(ym)} 출입문개폐요원 출근 내역`;
    const weeklyHours = calcWeeklyContractHours();
    const payDate = state.payDate || defaultPayDateForMonth(ym);

    // AoA: A~R(18 cols), rows 1~16
    const cols = 18;
    const rows = 16;
    const aoa = Array.from({ length: rows }, () => Array(cols).fill(""));

    // Title row (row2 col C..G)
    aoa[1][2] = title;

    // Weekday header row (row4, B..H)
    ["일", "월", "화", "수", "목", "금", "토"].forEach((w, i) => {
      aoa[3][1 + i] = w;
    });
    aoa[3][14] = `근로계약서 상 근로시간 (주 ${weeklyHours}시간)`;

    // summary block
    aoa[4][9] = "당월 근무일수";
    aoa[4][10] = totals.days;

    aoa[5][9] = "당월 유급 근로시간";
    aoa[5][10] = totals.hours;
    aoa[6][9] = "당월 보수 지급액";
    aoa[6][10] = totals.gross;
    aoa[7][9] = "지급일";
    aoa[7][10] = payDate;

    // schedule description (N5/O5 style)
    aoa[4][13] = "요일별 시간대(요약)";
    const schedLines = [];
    [1,2,3,4,5].forEach((dow) => {
      const h = calcScheduleHours(dow);
      if (h <= 0) return;
      const s = state.schedule[dow];
      const range = [s.s1 && s.e1 ? `${s.s1}~${s.e1}` : "", s.s2 && s.e2 ? `${s.s2}~${s.e2}` : ""].filter(Boolean).join(", ");
      schedLines.push(`${weekdayKorean[dow]} ${h}h (${range || "시간대 미기재"})`);
    });
    aoa[4][14] = schedLines.join(" / ");

    // Fill calendar weeks starting row5 (index 4), with date rows 5,7,9,11,13,15 and mark rows 6,8,10,12,14,16
    const weeks = makeMonthWeeks(ym);
    for (let w = 0; w < 6; w++) {
      const dateRow = 4 + w * 2; // 0-based row index
      const markRow = dateRow + 1;
      if (markRow >= rows) break;
      for (let dow = 0; dow < 7; dow++) {
        const col = 1 + dow; // B..H
        const d = weeks[w][dow];
        if (!d) continue;

        const iso = toISODate(d);
        aoa[dateRow][col] = iso; // string date
        const a = state.attendance[iso] || { on: false };
        aoa[markRow][col] = a.on ? "O" : "";
      }
    }

    // Put hourly rate as note (row10 in sample)
    aoa[9][10] = `시급: ${fmtWon(state.hourlyRate)}원`;

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // merges
    ws["!merges"] = [
      // C2:G2
      { s: { r: 1, c: 2 }, e: { r: 1, c: 6 } },
      // O4:R4
      { s: { r: 3, c: 14 }, e: { r: 3, c: 17 } },
      // O5:R5 (schedule summary)
      { s: { r: 4, c: 14 }, e: { r: 4, c: 17 } },
    ];

    ws["!cols"] = [
      { wch: 2 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
      { wch: 2 }, { wch: 18 }, { wch: 16 }, { wch: 2 }, { wch: 2 }, { wch: 18 }, { wch: 55 }, { wch: 2 }, { wch: 2 }, { wch: 2 }
    ];

    // format date strings as text (simple). If you prefer real excel date, convert serial.
    return ws;
  }

  function buildSheet2Calc(ym, calc) {
    const { days, hours, hourly, gross, healthAmt, ltAmt, empAmt, incomeTax, localTax, pension, otherDeduct, deductTotal, net } = calc;
    const title = `${ymToTitle(ym)} 인건비 산정내역`;

    const aoa = [
      [title],
      [],
      ["학교명", state.schoolName || "", "", "지급월", ym],
      ["근로자", state.workerName || "", "", "지급일", state.payDate || defaultPayDateForMonth(ym)],
      ["직종", state.jobTitle || "출입문개폐전담원", "", "시급(원)", hourly],
      [],
      ["항목", "값"],
      ["당월 근무일수(일)", days],
      ["당월 유급 근로시간(시간)", hours],
      ["당월 보수 지급액(원)", gross],
      [],
      ["공제(개인부담)", "금액(원)"],
      ["건강보험", healthAmt],
      ["장기요양보험", ltAmt],
      ["고용보험", empAmt],
      ["소득세", incomeTax],
      ["주민세", localTax],
      ["국민연금", pension],
      ["기타", otherDeduct],
      ["공제액 계(B)", deductTotal],
      [],
      ["실수령액(A-B)", net],
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 26 }, { wch: 22 }, { wch: 6 }, { wch: 14 }, { wch: 16 }];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // title merge
    ];
    return ws;
  }

  function buildSheet3Payslip(ym, calc) {
    const title = `${ymToTitle(ym)} 임금명세서`;
    const payDate = state.payDate || defaultPayDateForMonth(ym);

    const {
      hours, hourly, gross, healthAmt, ltAmt, empAmt,
      incomeTax, localTax, pension, otherDeduct, deductTotal, net
    } = calc;

    // 21 rows, 9 cols(A..I) but we mainly use B..I.
    const rows = 21;
    const cols = 9;
    const aoa = Array.from({ length: rows }, () => Array(cols).fill(""));

    aoa[0][1] = title; // B1

    // Row3 (index2)
    aoa[2][1] = "소속";
    aoa[2][3] = state.schoolName || "";
    aoa[2][4] = "지급일";
    aoa[2][6] = payDate;

    // Row4
    aoa[3][1] = "성명";
    aoa[3][3] = state.workerName || "";
    aoa[3][4] = "직종";
    aoa[3][6] = state.jobTitle || "출입문개폐전담원";

    // Row5
    aoa[4][1] = "생년월일";
    aoa[4][3] = state.birthDate || "";
    aoa[4][4] = "최초임용일";
    aoa[4][6] = state.hireDate || "";

    // Row7 headers
    aoa[6][1] = "급여내역";
    aoa[6][6] = "공제내역";

    // Row8 columns
    aoa[7][1] = "임금항목";
    aoa[7][3] = "산출식";
    aoa[7][5] = "금액";
    aoa[7][6] = "구분";
    aoa[7][8] = "금액";

    // Row9
    aoa[8][1] = "매월\n지급";
    aoa[8][2] = "기본급";
    aoa[8][3] = `시급 ${fmtWon(hourly)}원*${hours}시간=`;
    aoa[8][5] = gross;
    aoa[8][6] = "소득세";
    aoa[8][8] = incomeTax;

    // Row10~16 pay items blank, deductions
    const payItems = [
      ["근속수당", "주민세", localTax],
      ["정액급식비", "건강보험", healthAmt],
      ["위험근무수당", "장기요양보험", ltAmt],
      ["면허가산수당", "국민연금", pension],
      ["특수업무수당", "고용보험", empAmt],
      ["급식운영수당", "", ""],
      ["가족수당", "기타", otherDeduct],
    ];
    for (let i = 0; i < payItems.length; i++) {
      const r = 9 + i;
      aoa[r][2] = payItems[i][0];
      aoa[r][6] = payItems[i][1];
      aoa[r][8] = payItems[i][2];
    }

    // Row18 (index17) - 부정기
    aoa[17][1] = "부정기\n지급";
    aoa[17][2] = "명절휴가비";

    // Row20 totals (index19)
    aoa[19][1] = "급여총액 계 (A)";
    aoa[19][5] = gross;
    aoa[19][6] = "공제액 계 (B)";
    aoa[19][8] = deductTotal;

    // Row21 net (index20)
    aoa[20][1] = "실수령액 (A-B)";
    aoa[20][8] = net;

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 2 },  // A
      { wch: 10 }, // B
      { wch: 14 }, // C
      { wch: 28 }, // D
      { wch: 2 },  // E
      { wch: 14 }, // F
      { wch: 16 }, // G
      { wch: 2 },  // H
      { wch: 14 }, // I
    ];

    ws["!merges"] = [
      // B1:I1
      { s: { r: 0, c: 1 }, e: { r: 0, c: 8 } },
      // B3:C3
      { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } },
      // E3:F3
      { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } },
      // G3:I3
      { s: { r: 2, c: 6 }, e: { r: 2, c: 8 } },
      // B4:C4
      { s: { r: 3, c: 1 }, e: { r: 3, c: 2 } },
      // E4:F4
      { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
      // G4:I4
      { s: { r: 3, c: 6 }, e: { r: 3, c: 8 } },
      // B5:C5
      { s: { r: 4, c: 1 }, e: { r: 4, c: 2 } },
      // E5:F5
      { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } },
      // G5:I5
      { s: { r: 4, c: 6 }, e: { r: 4, c: 8 } },
      // B7:F7 (급여내역)
      { s: { r: 6, c: 1 }, e: { r: 6, c: 5 } },
      // G7:I7 (공제내역)
      { s: { r: 6, c: 6 }, e: { r: 6, c: 8 } },
      // D8:E8 (산출식)
      { s: { r: 7, c: 3 }, e: { r: 7, c: 4 } },
      // G8:H8 (구분)
      { s: { r: 7, c: 6 }, e: { r: 7, c: 7 } },
      // D9:E9
      { s: { r: 8, c: 3 }, e: { r: 8, c: 4 } },
      // G9:H9
      { s: { r: 8, c: 6 }, e: { r: 8, c: 7 } },
      // B9:B17
      { s: { r: 8, c: 1 }, e: { r: 16, c: 1 } },
      // B18:B19
      { s: { r: 17, c: 1 }, e: { r: 18, c: 1 } },
      // B20:E20
      { s: { r: 19, c: 1 }, e: { r: 19, c: 4 } },
      // G20:H20
      { s: { r: 19, c: 6 }, e: { r: 19, c: 7 } },
      // B21:H21
      { s: { r: 20, c: 1 }, e: { r: 20, c: 7 } },
    ];

    return ws;
  }

  async function exportPayslipPdf() {
    const libs = isLibReady();
    if (!libs.okCanvas || !libs.okJsPdf) {
      alert("PDF 라이브러리를 불러오지 못했습니다. (인터넷 연결을 확인)");
      return;
    }

    // payslipPreview를 캡처
    const node = $("payslipPreview");
    const scale = 2;

    const canvas = await window.html2canvas(node, {
      scale,
      backgroundColor: "#ffffff",
      useCORS: true
    });

    const imgData = canvas.toDataURL("image/png");

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "p", unit: "mm", format: "a4" });

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    // 이미지 비율 유지해서 A4에 맞추기
    const imgProps = pdf.getImageProperties(imgData);
    const imgW = pageWidth - 20; // 좌우 10mm 마진
    const imgH = (imgProps.height * imgW) / imgProps.width;

    let y = 10;
    if (imgH <= pageHeight - 20) {
      pdf.addImage(imgData, "PNG", 10, y, imgW, imgH);
    } else {
      // 여러 페이지로 분할 (세로 긴 경우)
      let remaining = imgH;
      let posY = 10;
      let srcY = 0;
      const ratio = imgProps.width / imgW; // px per mm (approx)
      const pageUsableH = pageHeight - 20;

      while (remaining > 0) {
        const sliceH = Math.min(pageUsableH, remaining);
        // 캔버스에서 slice를 잘라 새 이미지로
        const sliceCanvas = document.createElement("canvas");
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = Math.floor(sliceH * ratio);
        const ctx = sliceCanvas.getContext("2d");
        ctx.drawImage(canvas, 0, Math.floor(srcY * ratio), canvas.width, sliceCanvas.height, 0, 0, canvas.width, sliceCanvas.height);
        const sliceData = sliceCanvas.toDataURL("image/png");

        pdf.addImage(sliceData, "PNG", 10, posY, imgW, sliceH);

        remaining -= sliceH;
        srcY += sliceH;
        if (remaining > 0) pdf.addPage();
      }
    }

    const ym = state.payMonth || defaultPayMonth();
    const p = getMonthParts(ym);
    const fname = `임금명세서_${p ? (p.y + pad2(p.m)) : ym}_${(state.workerName || "근로자")}.pdf`;
    pdf.save(fname);
  }

  // ---------- init ----------
  function computeAndRenderAll() {
    renderSchedule();
    renderCalendar();
    computeAndRender();
  }

  function initLibraryWarning() {
    const libs = isLibReady();
    const warn = [];
    if (!libs.okXlsx) warn.push("XLSX(엑셀) 라이브러리 로드 실패");
    if (!libs.okCanvas) warn.push("html2canvas 로드 실패");
    if (!libs.okJsPdf) warn.push("jsPDF 로드 실패");

    $("libWarn").textContent = warn.length ? `⚠ ${warn.join(" / ")} · 인터넷 연결 확인` : "";
  }

  function init() {
    loadState();

    // 기본값 채우기(저장값이 없거나 일부 누락 시)
    if (!state.payMonth) state.payMonth = defaultPayMonth();
    if (!state.payDate) state.payDate = defaultPayDateForMonth(state.payMonth);

    if (!state.schedule || Object.keys(state.schedule).length < 7) state.schedule = defaultSchedule();

    syncUIFromState();
    renderSchedule();
    renderCalendar();

    bindInputs();
    setupGridNavigation();
    initLibraryWarning();
    computeAndRender();

    // grids rebuilt after initial render
    rebuildGridMap("schedule");
    rebuildGridMap("calendar");
    rebuildGridMap("calendarHours");
  }

  // Actually call init
  window.addEventListener("DOMContentLoaded", init);

})();
