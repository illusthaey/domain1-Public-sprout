/* 출입문개폐전담원 월별 인건비 지출서식 (클라이언트 전용)
 * - Excel 업/다운로드: SheetJS(XLSX)
 * - PDF: html2canvas + jsPDF
 * - 입력값 localStorage 저장
 *
 * [수정 포인트]
 * - index.html에서 삭제/이동된 버튼(id 없는 요소) 때문에 앱이 죽지 않도록 null-guard 적용
 * - 건강보험/장기요양 체크박스가 index에서 id가 중복(deductLtCareOn)이라도 동작하도록
 *   "카드(card) 내부에서 체크박스 찾기" 방식으로 바인딩
 */
(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);

  // 안전한 DOM 유틸 (요소 없으면 조용히 무시)
  const onEl = (el, evt, fn) => { if (el) el.addEventListener(evt, fn); };
  const onId = (id, evt, fn) => onEl($(id), evt, fn);
  const setText = (id, text) => { const el = $(id); if (el) el.textContent = text; };
  const setVal = (id, value) => { const el = $(id); if (el) el.value = value; };

  // index.html에서 건강/장기요양 체크박스 id가 중복일 수 있어,
  // "해당 카드 내부의 checkbox"를 anchor input을 기준으로 찾는다.
  function getCheckboxInCardByAnchorId(anchorId) {
    const anchor = $(anchorId);
    if (!anchor) return null;
    const card = anchor.closest(".card");
    if (!card) return null;
    return card.querySelector('input[type="checkbox"]');
  }

  // 건강보험 토글: healthRate가 있는 카드의 checkbox를 우선 사용
  function getHealthToggleEl() {
    const byCard = getCheckboxInCardByAnchorId("healthRate");
    if (byCard) return byCard;

    // 혹시 HTML을 나중에 정상 id로 고치면 여기로도 잡힘
    const byId = $("deductHealthOn");
    if (byId) return byId;

    // 현재 index처럼 id가 잘못되어 첫 번째 deductLtCareOn이 건강보험일 가능성이 높음
    const all = document.querySelectorAll("#deductLtCareOn");
    return all && all.length ? all[0] : null;
  }

  // 장기요양 토글: ltCareRate가 있는 카드의 checkbox를 우선 사용
  function getLtCareToggleEl() {
    const byCard = getCheckboxInCardByAnchorId("ltCareRate");
    if (byCard) return byCard;

    // 정상 id가 있으면 사용
    const byId = $("deductLtCareOn");
    // (주의) getElementById는 중복 id에서 첫 번째만 반환 → fallback으로만 사용
    if (byId) {
      // 그래도 중복이라면 두 번째를 더 우선
      const all = document.querySelectorAll("#deductLtCareOn");
      if (all && all.length >= 2) return all[1];
      return byId;
    }

    return null;
  }

  // 고용보험 토글: id가 정상인 경우가 많지만, 카드 기준도 지원
  function getEmpToggleEl() {
    const byCard = getCheckboxInCardByAnchorId("empRate");
    if (byCard) return byCard;
    return $("deductEmpOn");
  }

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
    if (!tbody) return;

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
    if (!tbody) return;

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

    hours = Math.round(hours * 100) / 100;
    return { days, hours };
  }

  function autoInsuranceAmounts(gross) {
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
    const weeklyHours = calcWeeklyContractHours();
    if ($("contractWeeklyHoursPill")) {
      $("contractWeeklyHoursPill").textContent = `주간 계약시간: ${weeklyHours}시간`;
      $("contractWeeklyHoursPill").className = "pill" + (weeklyHours > 0 ? " ok" : "");
    }

    const ym = state.payMonth || defaultPayMonth();
    const { days, hours } = sumWorkHoursForMonth(ym);

    setText("workDaysText", String(days));
    setText("workHoursText", String(hours));

    // Gross
    const hourly = parseNumber(state.hourlyRate, 0);
    const gross = Math.round(hourly * hours); // 원 단위
    setText("grossPayText", fmtWon(gross));

    // Auto insurance calc (only if toggle ON and not manual)
    const auto = autoInsuranceAmounts(gross);

    if (state.deductHealthOn && !state.manual.healthAmt) {
      state.healthAmt = auto.health;
      setVal("healthAmt", String(state.healthAmt));
    }
    if (state.deductLtCareOn && !state.manual.ltCareAmt) {
      state.ltCareAmt = auto.lt;
      setVal("ltCareAmt", String(state.ltCareAmt));
    }
    if (state.deductEmpOn && !state.manual.empAmt) {
      state.empAmt = auto.emp;
      setVal("empAmt", String(state.empAmt));
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
    setText("deductTotalText", fmtWon(deductTotal));

    const net = gross - deductTotal;
    setText("netPayText", fmtWon(net));

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
    const el = $("payslipPreview");
    if (!el) return;

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

      Object.assign(state, parsed);

      if (!state.schedule) state.schedule = defaultSchedule();
      for (let i = 0; i < 7; i++) {
        if (!state.schedule[i]) state.schedule[i] = { s1: "", e1: "", s2: "", e2: "", memo: "" };
      }
      if (!state.attendance) state.attendance = {};
      if (!state.manual) state.manual = { healthAmt: false, ltCareAmt: false, empAmt: false };
      if (!state.ui) state.ui = { showManualDailyHours: false };

      // index에 saveStateText가 없을 수 있으니 안전 처리
      setText("saveStateText", "불러옴");
    } catch (e) {
      console.warn("Failed to load state:", e);
    }
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      setText("saveStateText", new Date().toLocaleString("ko-KR"));
    } catch (e) {
      console.warn("Failed to save state:", e);
      setText("saveStateText", "저장 실패");
    }
  }

  let saveT = null;
  function scheduleSaveDebounced() {
    if (saveT) clearTimeout(saveT);
    saveT = setTimeout(() => saveState(), 400);
  }

  // ---------- wire inputs ----------
  function bindInputs() {
    onId("payMonth", "change", () => {
      state.payMonth = $("payMonth").value;
      if (!state.payDate) state.payDate = defaultPayDateForMonth(state.payMonth);
      setVal("payDate", state.payDate);
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("payDate", "change", () => {
      state.payDate = $("payDate").value;
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("schoolName", "input", () => { state.schoolName = $("schoolName").value; computeAndRender(); scheduleSaveDebounced(); });
    onId("workerName", "input", () => { state.workerName = $("workerName").value; computeAndRender(); scheduleSaveDebounced(); });
    onId("jobTitle", "input", () => { state.jobTitle = $("jobTitle").value; computeAndRender(); scheduleSaveDebounced(); });
    onId("birthDate", "change", () => { state.birthDate = $("birthDate").value; computeAndRender(); scheduleSaveDebounced(); });
    onId("hireDate", "change", () => { state.hireDate = $("hireDate").value; computeAndRender(); scheduleSaveDebounced(); });

    onId("hourlyRate", "input", () => {
      state.hourlyRate = parseNumber($("hourlyRate").value, 0);
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("preset2025Btn", "click", () => {
      state.hourlyRate = 11200;
      setVal("hourlyRate", "11200");
      computeAndRender();
      scheduleSaveDebounced();
    });
    onId("preset2026Btn", "click", () => {
      state.hourlyRate = 11500;
      setVal("hourlyRate", "11500");
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("loadExampleScheduleBtn", "click", () => {
      loadExampleSchedule();
      renderSchedule();
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("showManualDailyHours", "change", () => {
      state.ui.showManualDailyHours = $("showManualDailyHours").checked;
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("fillByScheduleBtn", "click", () => {
      fillCalendarBySchedule();
      computeAndRender();
      scheduleSaveDebounced();
    });

    onId("clearCalendarBtn", "click", () => {
      clearCalendar();
      computeAndRender();
      scheduleSaveDebounced();
    });

    // index에서 saveNowBtn이 없을 수 있음 → 있으면만 바인딩
    onId("saveNowBtn", "click", () => saveState());

    onId("resetAllBtn", "click", () => {
      if (!confirm("전체 입력값을 초기화할까요? (브라우저 저장값 포함)")) return;
      localStorage.removeItem(STORAGE_KEY);
      location.reload();
    });

    // truncate10won 토글이 index에서 제거된 상태일 수 있음 → 있으면만 바인딩
    onId("truncate10won", "change", () => {
      state.truncate10won = $("truncate10won").checked;
      computeAndRender();
      scheduleSaveDebounced();
    });

    // ===== 사회보험 토글(중요: 카드 기반으로 찾음) =====
    const healthToggle = getHealthToggleEl();
    onEl(healthToggle, "change", () => {
      state.deductHealthOn = !!healthToggle.checked;
      computeAndRender();
      scheduleSaveDebounced();
    });

    const ltToggle = getLtCareToggleEl();
    onEl(ltToggle, "change", () => {
      state.deductLtCareOn = !!ltToggle.checked;
      computeAndRender();
      scheduleSaveDebounced();
    });

    const empToggle = getEmpToggleEl();
    onEl(empToggle, "change", () => {
      state.deductEmpOn = !!empToggle.checked;
      computeAndRender();
      scheduleSaveDebounced();
    });

    // rates & amounts
    onId("healthRate", "input", () => { state.healthRate = parseNumber($("healthRate").value, 0); state.manual.healthAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    onId("healthAmt", "input", () => { state.healthAmt = parseNumber($("healthAmt").value, 0); state.manual.healthAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    onId("healthAutoBtn", "click", () => { state.manual.healthAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    onId("ltCareRate", "input", () => { state.ltCareRate = parseNumber($("ltCareRate").value, 0); state.manual.ltCareAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    onId("ltCareAmt", "input", () => { state.ltCareAmt = parseNumber($("ltCareAmt").value, 0); state.manual.ltCareAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    onId("ltCareAutoBtn", "click", () => { state.manual.ltCareAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    onId("empRate", "input", () => { state.empRate = parseNumber($("empRate").value, 0); state.manual.empAmt = false; computeAndRender(); scheduleSaveDebounced(); });
    onId("empAmt", "input", () => { state.empAmt = parseNumber($("empAmt").value, 0); state.manual.empAmt = true; computeAndRender(); scheduleSaveDebounced(); });
    onId("empAutoBtn", "click", () => { state.manual.empAmt = false; computeAndRender(); scheduleSaveDebounced(); });

    onId("incomeTax", "input", () => { state.incomeTax = parseNumber($("incomeTax").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    onId("localTax", "input", () => { state.localTax = parseNumber($("localTax").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    onId("pension", "input", () => { state.pension = parseNumber($("pension").value, 0); computeAndRender(); scheduleSaveDebounced(); });
    onId("otherDeduct", "input", () => { state.otherDeduct = parseNumber($("otherDeduct").value, 0); computeAndRender(); scheduleSaveDebounced(); });

    // file actions
    onId("downloadTemplateBtn", "click", () => downloadTemplateXlsx());
    onId("uploadXlsx", "change", (e) => uploadTemplateXlsx(e.target.files?.[0]));

    // index에서 export 버튼이 삭제됐을 수 있음 → 있으면만 바인딩
    onId("exportXlsxBtn", "click", () => exportResultXlsx());
    onId("exportPdfBtn", "click", () => exportPayslipPdf());
  }

  function syncUIFromState() {
    const ym = state.payMonth || defaultPayMonth();
    setVal("payMonth", ym);
    setVal("payDate", state.payDate || defaultPayDateForMonth(ym));

    setVal("schoolName", state.schoolName || "");
    setVal("workerName", state.workerName || "");
    setVal("jobTitle", state.jobTitle || "출입문개폐전담원");
    setVal("birthDate", state.birthDate || "");
    setVal("hireDate", state.hireDate || "");

    setVal("hourlyRate", String(parseNumber(state.hourlyRate, 0)));

    const manualDaily = $("showManualDailyHours");
    if (manualDaily) manualDaily.checked = !!state.ui.showManualDailyHours;

    // truncate10won 토글은 없을 수 있음
    const trunc = $("truncate10won");
    if (trunc) trunc.checked = !!state.truncate10won;

    // 보험 토글(카드 기준)
    const healthToggle = getHealthToggleEl();
    if (healthToggle) healthToggle.checked = !!state.deductHealthOn;

    const ltToggle = getLtCareToggleEl();
    if (ltToggle) ltToggle.checked = !!state.deductLtCareOn;

    const empToggle = getEmpToggleEl();
    if (empToggle) empToggle.checked = !!state.deductEmpOn;

    setVal("healthRate", String(parseNumber(state.healthRate, 0)));
    setVal("healthAmt", String(parseNumber(state.healthAmt, 0)));

    setVal("ltCareRate", String(parseNumber(state.ltCareRate, 0)));
    setVal("ltCareAmt", String(parseNumber(state.ltCareAmt, 0)));

    setVal("empRate", String(parseNumber(state.empRate, 0)));
    setVal("empAmt", String(parseNumber(state.empAmt, 0)));

    setVal("incomeTax", String(parseNumber(state.incomeTax, 0)));
    setVal("localTax", String(parseNumber(state.localTax, 0)));
    setVal("pension", String(parseNumber(state.pension, 0)));
    setVal("otherDeduct", String(parseNumber(state.otherDeduct, 0)));
  }

  function loadExampleSchedule() {
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
      // 수동시간은 유지
    });

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

      if (e.altKey || e.metaKey || e.ctrlKey) return;

      const r = parseInt(t.dataset.r || "0", 10);
      const c = parseInt(t.dataset.c || "0", 10);

      if (e.key === " ") {
        if ((t instanceof HTMLInputElement) && t.type === "checkbox") {
          e.preventDefault();
          t.checked = !t.checked;
          t.dispatchEvent(new Event("change", { bubbles: true }));
          return;
        }
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

    // 3) 출근내역(선택) sheet
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

    setText("uploadStatus", "읽는 중...");
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

      syncUIFromState();
      renderSchedule();
      renderCalendar();
      computeAndRender();
      scheduleSaveDebounced();

      setText("uploadStatus", `업로드 완료: ${file.name}`);
    } catch (err) {
      console.error(err);
      setText("uploadStatus", "업로드 실패: 파일 형식/내용을 확인하세요.");
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
    const names = wb.SheetNames || [];
    const paySheetName = names.find((n) => n.includes("임금명세서"));
    if (!paySheetName) return;

    const payWs = wb.Sheets[paySheetName];
    const get = (addr) => payWs[addr] ? payWs[addr].v : "";

    state.schoolName = String(get("D3") || "");
    const payDate = get("G3");
    state.payDate = excelDateToISO(payDate) || "";
    state.workerName = String(get("D4") || "");
    state.jobTitle = String(get("G4") || "출입문개폐전담원");
    state.birthDate = excelDateToISO(get("D5")) || "";
    state.hireDate = excelDateToISO(get("G5")) || "";

    const d9 = String(get("D9") || "");
    const m = /시급\s*([0-9,]+)원/.exec(d9.replace(/\s+/g, " "));
    if (m) state.hourlyRate = parseNumber(m[1].replace(/,/g, ""), state.hourlyRate);

    const ymMatch = /(\d{4})\.(\d{1,2})월/.exec(paySheetName);
    if (ymMatch) {
      state.payMonth = `${ymMatch[1]}-${pad2(parseInt(ymMatch[2], 10))}`;
    }

    const attSheetName = names.find((n) => n.includes("출근내역") && (ymMatch ? n.includes(`${ymMatch[1]}.${ymMatch[2]}월`) : true));
    if (!attSheetName) return;
    const attWs = wb.Sheets[attSheetName];

    const cells = Object.keys(attWs).filter((k) => !k.startsWith("!"));

    const dateCandidates = cells.filter((addr) => {
      const v = attWs[addr]?.v;
      return typeof v === "number" && v > 40000 && v < 60000;
    });

    dateCandidates.forEach((addr) => {
      const { c, r } = XLSX.utils.decode_cell(addr);
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
      const d = XLSX.SSF.parse_date_code(v);
      if (!d) return "";
      const js = new Date(d.y, d.m - 1, d.d);
      return toISODate(js);
    }
    if (v instanceof Date && !Number.isNaN(v.getTime())) return toISODate(v);
    return "";
  }

  // ===== 결과 export 함수들은 유지(버튼이 있으면만 실행됨) =====
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

    const cols = 18;
    const rows = 16;
    const aoa = Array.from({ length: rows }, () => Array(cols).fill(""));

    aoa[1][2] = title;

    ["일", "월", "화", "수", "목", "금", "토"].forEach((w, i) => {
      aoa[3][1 + i] = w;
    });
    aoa[3][14] = `근로계약서 상 근로시간 (주 ${weeklyHours}시간)`;

    aoa[4][9] = "당월 근무일수";
    aoa[4][10] = totals.days;

    aoa[5][9] = "당월 유급 근로시간";
    aoa[5][10] = totals.hours;
    aoa[6][9] = "당월 보수 지급액";
    aoa[6][10] = totals.gross;
    aoa[7][9] = "지급일";
    aoa[7][10] = payDate;

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

    const weeks = makeMonthWeeks(ym);
    for (let w = 0; w < 6; w++) {
      const dateRow = 4 + w * 2;
      const markRow = dateRow + 1;
      if (markRow >= rows) break;
      for (let dow = 0; dow < 7; dow++) {
        const col = 1 + dow;
        const d = weeks[w][dow];
        if (!d) continue;

        const iso = toISODate(d);
        aoa[dateRow][col] = iso;
        const a = state.attendance[iso] || { on: false };
        aoa[markRow][col] = a.on ? "O" : "";
      }
    }

    aoa[9][10] = `시급: ${fmtWon(state.hourlyRate)}원`;

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws["!merges"] = [
      { s: { r: 1, c: 2 }, e: { r: 1, c: 6 } },
      { s: { r: 3, c: 14 }, e: { r: 3, c: 17 } },
      { s: { r: 4, c: 14 }, e: { r: 4, c: 17 } },
    ];

    ws["!cols"] = [
      { wch: 2 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
      { wch: 2 }, { wch: 18 }, { wch: 16 }, { wch: 2 }, { wch: 2 }, { wch: 18 }, { wch: 55 }, { wch: 2 }, { wch: 2 }, { wch: 2 }
    ];

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
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
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

    const rows = 21;
    const cols = 9;
    const aoa = Array.from({ length: rows }, () => Array(cols).fill(""));

    aoa[0][1] = title;

    aoa[2][1] = "소속";
    aoa[2][3] = state.schoolName || "";
    aoa[2][4] = "지급일";
    aoa[2][6] = payDate;

    aoa[3][1] = "성명";
    aoa[3][3] = state.workerName || "";
    aoa[3][4] = "직종";
    aoa[3][6] = state.jobTitle || "출입문개폐전담원";

    aoa[4][1] = "생년월일";
    aoa[4][3] = state.birthDate || "";
    aoa[4][4] = "최초임용일";
    aoa[4][6] = state.hireDate || "";

    aoa[6][1] = "급여내역";
    aoa[6][6] = "공제내역";

    aoa[7][1] = "임금항목";
    aoa[7][3] = "산출식";
    aoa[7][5] = "금액";
    aoa[7][6] = "구분";
    aoa[7][8] = "금액";

    aoa[8][1] = "매월\n지급";
    aoa[8][2] = "기본급";
    aoa[8][3] = `시급 ${fmtWon(hourly)}원*${hours}시간=`;
    aoa[8][5] = gross;
    aoa[8][6] = "소득세";
    aoa[8][8] = incomeTax;

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

    aoa[17][1] = "부정기\n지급";
    aoa[17][2] = "명절휴가비";

    aoa[19][1] = "급여총액 계 (A)";
    aoa[19][5] = gross;
    aoa[19][6] = "공제액 계 (B)";
    aoa[19][8] = deductTotal;

    aoa[20][1] = "실수령액 (A-B)";
    aoa[20][8] = net;

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 2 },
      { wch: 10 },
      { wch: 14 },
      { wch: 28 },
      { wch: 2 },
      { wch: 14 },
      { wch: 16 },
      { wch: 2 },
      { wch: 14 },
    ];

    ws["!merges"] = [
      { s: { r: 0, c: 1 }, e: { r: 0, c: 8 } },
      { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } },
      { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } },
      { s: { r: 2, c: 6 }, e: { r: 2, c: 8 } },
      { s: { r: 3, c: 1 }, e: { r: 3, c: 2 } },
      { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
      { s: { r: 3, c: 6 }, e: { r: 3, c: 8 } },
      { s: { r: 4, c: 1 }, e: { r: 4, c: 2 } },
      { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } },
      { s: { r: 4, c: 6 }, e: { r: 4, c: 8 } },
      { s: { r: 6, c: 1 }, e: { r: 6, c: 5 } },
      { s: { r: 6, c: 6 }, e: { r: 6, c: 8 } },
      { s: { r: 7, c: 3 }, e: { r: 7, c: 4 } },
      { s: { r: 7, c: 6 }, e: { r: 7, c: 7 } },
      { s: { r: 8, c: 3 }, e: { r: 8, c: 4 } },
      { s: { r: 8, c: 6 }, e: { r: 8, c: 7 } },
      { s: { r: 8, c: 1 }, e: { r: 16, c: 1 } },
      { s: { r: 17, c: 1 }, e: { r: 18, c: 1 } },
      { s: { r: 19, c: 1 }, e: { r: 19, c: 4 } },
      { s: { r: 19, c: 6 }, e: { r: 19, c: 7 } },
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

    const node = $("payslipPreview");
    if (!node) {
      alert("임금명세서 미리보기 영역(payslipPreview)을 찾지 못했습니다.");
      return;
    }

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

    const imgProps = pdf.getImageProperties(imgData);
    const imgW = pageWidth - 20;
    const imgH = (imgProps.height * imgW) / imgProps.width;

    if (imgH <= pageHeight - 20) {
      pdf.addImage(imgData, "PNG", 10, 10, imgW, imgH);
    } else {
      let remaining = imgH;
      let srcY = 0;
      const ratio = imgProps.width / imgW;
      const pageUsableH = pageHeight - 20;

      while (remaining > 0) {
        const sliceH = Math.min(pageUsableH, remaining);

        const sliceCanvas = document.createElement("canvas");
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = Math.floor(sliceH * ratio);
        const ctx = sliceCanvas.getContext("2d");
        ctx.drawImage(
          canvas,
          0, Math.floor(srcY * ratio),
          canvas.width, sliceCanvas.height,
          0, 0,
          canvas.width, sliceCanvas.height
        );

        const sliceData = sliceCanvas.toDataURL("image/png");
        pdf.addImage(sliceData, "PNG", 10, 10, imgW, sliceH);

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
  function initLibraryWarning() {
    const libs = isLibReady();
    const warn = [];
    if (!libs.okXlsx) warn.push("XLSX(엑셀) 라이브러리 로드 실패");
    if (!libs.okCanvas) warn.push("html2canvas 로드 실패");
    if (!libs.okJsPdf) warn.push("jsPDF 로드 실패");

    setText("libWarn", warn.length ? `⚠ ${warn.join(" / ")} · 인터넷 연결 확인` : "");
  }

  function init() {
    loadState();

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

    rebuildGridMap("schedule");
    rebuildGridMap("calendar");
    rebuildGridMap("calendarHours");
  }

  window.addEventListener("DOMContentLoaded", init);

})();
