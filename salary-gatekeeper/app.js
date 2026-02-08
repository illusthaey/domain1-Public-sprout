(() => {
  "use strict";

  const STORAGE_KEY = "door-access-pay:v1";

  const DOW_LABELS = ["일", "월", "화", "수", "목", "금", "토"];

  function pad2(n) { return String(n).padStart(2, "0"); }

  function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  }
  function currentMonthStr() {
    const d = new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
  }

  function toInt(v, fallback = 0) {
    const n = Number(v);
    return Number.isFinite(n) ? Math.trunc(n) : fallback;
  }
  function toFloat(v, fallback = 0) {
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  }

  // 원단위 절삭: 1111 -> 1110
  function trunc10Won(n) {
    const x = Math.floor(Number(n) / 10) * 10;
    return Number.isFinite(x) ? x : 0;
  }

  function formatWon(n) {
    const x = Math.round(Number(n) || 0);
    return x.toLocaleString("ko-KR");
  }

  function timeToMinutes(t) {
    if (!t || typeof t !== "string") return null;
    const m = t.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    const hh = Number(m[1]);
    const mm = Number(m[2]);
    if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return hh * 60 + mm;
  }

  function intervalMinutes(s, e) {
    const ms = timeToMinutes(s);
    const me = timeToMinutes(e);
    if (ms == null || me == null) return 0;
    let diff = me - ms;
    if (diff < 0) diff += 24 * 60; // 혹시 자정 넘어가는 경우까지 방어
    return diff;
  }

  function calcDayHours(day) {
    const mins = intervalMinutes(day.s1, day.e1) + intervalMinutes(day.s2, day.e2);
    const hours = mins / 60;
    // 소수점 2자리 정도로 표기 안정화
    return Math.round(hours * 100) / 100;
  }

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function defaultState() {
    const base = {
      meta: {
        schoolName: "",
        docNo: "",
        payMonth: currentMonthStr(),
        payDate: todayStr(),
        workerName: "",
        jobTitle: "출입문개폐전담원",
        birthDate: "",
        firstHireDate: "",
      },
      rate: {
        hourly: 11500,
      },
      options: {
        monThuSame: true,
      },
      schedule: {
        // 문막초 예시를 기본으로 넣어둠(월~목 3시간, 금 2시간)
        0: { s1: "", e1: "", s2: "", e2: "" }, // 일
        1: { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10" }, // 월
        2: { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10" }, // 화
        3: { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10" }, // 수
        4: { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10" }, // 목
        5: { s1: "07:40", e1: "08:40", s2: "16:40", e2: "17:40" }, // 금
        6: { s1: "", e1: "", s2: "", e2: "" }, // 토
      },
      attendance: {
        // "YYYY-MM-DD": "O" | "휴가" | ...
      },
      payItems: {
        allowSeniority: 0,
        allowMeal: 0,
        allowDanger: 0,
        allowLicense: 0,
        allowSpecial: 0,
        allowMealOps: 0,
        allowFamily: 0,
        allowHolidayBonus: 0,
      },
      deductions: {
        incomeTax: 0,
        localTax: 0,
        pension: 0,
        etcDed: 0,
        health: {
          enabled: false,
          auto: true,
          // 기본값은 관행적인 값으로 넣되, 매년 바뀔 수 있어 UI에서 수정 가능하게 함
          rate: 7.09,     // 건강보험료율(총)
          ltcRate: 12.95, // 장기요양보험료율(건강보험료에 부과)
          amount: 0,      // 개인부담 합산(건강(개인)+요양(개인))
        },
        employment: {
          enabled: false,
          auto: true,
          rate: 0.9,   // 고용보험 개인부담(관행)
          amount: 0,
        },
      },
    };
    return base;
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const s = JSON.parse(raw);
      if (!s || typeof s !== "object") return null;
      return s;
    } catch {
      return null;
    }
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch {
      // storage 막힌 환경이면 조용히 무시
    }
  }

  let state = loadState() || defaultState();
  let renderedMonth = null;

  const $ = (id) => document.getElementById(id);

  function setInputValue(id, v) {
    const el = $(id);
    if (!el) return;
    el.value = v ?? "";
  }

  function setText(id, v) {
    const el = $(id);
    if (!el) return;
    el.textContent = v ?? "";
  }

  function monthParts(yyyyMM) {
    const m = String(yyyyMM || "").match(/^(\d{4})-(\d{2})$/);
    if (!m) return null;
    return { y: Number(m[1]), m: Number(m[2]) };
  }

  function dateStr(y, m, d) {
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }

  function getMonthInfo(yyyyMM) {
    const mp = monthParts(yyyyMM);
    if (!mp) return null;
    const first = new Date(mp.y, mp.m - 1, 1);
    const last = new Date(mp.y, mp.m, 0);
    return {
      y: mp.y,
      m: mp.m,
      firstDow: first.getDay(),
      days: last.getDate(),
    };
  }

  function buildScheduleTable() {
    const tbody = $("scheduleTbody");
    tbody.innerHTML = "";

    for (let dow = 0; dow < 7; dow++) {
      const tr = document.createElement("tr");

      const th = document.createElement("th");
      th.textContent = DOW_LABELS[dow];
      tr.appendChild(th);

      const makeTimeCell = (key, colIdx) => {
        const td = document.createElement("td");
        const inp = document.createElement("input");
        inp.type = "time";
        inp.value = state.schedule[dow]?.[key] || "";
        inp.id = `dow${dow}_${key}`;
        inp.dataset.grid = "sched";
        inp.dataset.r = String(dow);
        inp.dataset.c = String(colIdx);

        inp.addEventListener("input", () => {
          const val = inp.value || "";
          state.schedule[dow][key] = val;

          // 월~목 동일 모드: 월(1) 변경 시 화(2)~목(4) 자동 복사
          if (state.options.monThuSame && dow === 1) {
            for (let d = 2; d <= 4; d++) {
              state.schedule[d][key] = val;
              const twin = $(`dow${d}_${key}`);
              if (twin) twin.value = val;
            }
          }

          recomputeAndRender();
        });

        td.appendChild(inp);
        return td;
      };

      tr.appendChild(makeTimeCell("s1", 0));
      tr.appendChild(makeTimeCell("e1", 1));
      tr.appendChild(makeTimeCell("s2", 2));
      tr.appendChild(makeTimeCell("e2", 3));

      const tdHours = document.createElement("td");
      tdHours.className = "right mono";
      tdHours.innerHTML = `<span id="dow${dow}_hours">0</span> 시간`;
      tr.appendChild(tdHours);

      tbody.appendChild(tr);
    }
  }

  function buildCalendar() {
    const info = getMonthInfo(state.meta.payMonth);
    if (!info) return;

    const tbody = $("calendarTbody");
    tbody.innerHTML = "";

    let day = 1;
    for (let week = 0; week < 6; week++) {
      const tr = document.createElement("tr");

      for (let dow = 0; dow < 7; dow++) {
        const td = document.createElement("td");
        const cell = document.createElement("div");
        cell.className = "cal-cell";

        const isBefore = week === 0 && dow < info.firstDow;
        const isAfter = day > info.days;

        if (isBefore || isAfter) {
          cell.classList.add("cal-empty");
          cell.innerHTML = `<div class="cal-date">&nbsp;</div><div class="muted" style="font-size:0.9rem;">&nbsp;</div>`;
          td.appendChild(cell);
          tr.appendChild(td);
          continue;
        }

        const dStr = dateStr(info.y, info.m, day);
        const dateDiv = document.createElement("div");
        dateDiv.className = "cal-date";
        dateDiv.textContent = String(day);

        const inp = document.createElement("input");
        inp.type = "text";
        inp.className = "cal-mark";
        inp.setAttribute("list", "markList");
        inp.value = state.attendance[dStr] || "";
        inp.dataset.date = dStr;
        inp.dataset.grid = "cal";
        inp.dataset.r = String(week);
        inp.dataset.c = String(dow);

        inp.addEventListener("input", () => {
          const v = (inp.value || "").trim();
          if (!v) {
            delete state.attendance[dStr];
          } else {
            state.attendance[dStr] = v;
          }
          recomputeAndRender(/*noRebuild*/ true);
        });

        cell.appendChild(dateDiv);
        cell.appendChild(inp);

        td.appendChild(cell);
        tr.appendChild(td);

        day++;
      }

      tbody.appendChild(tr);
    }
  }

  // PDF 1페이지 달력(읽기전용)도 같이 생성
  function buildPdfCalendarMirror() {
    const info = getMonthInfo(state.meta.payMonth);
    if (!info) return;

    const tbody = $("pdfCalendarTbody");
    tbody.innerHTML = "";

    let day = 1;
    for (let week = 0; week < 6; week++) {
      const tr = document.createElement("tr");

      for (let dow = 0; dow < 7; dow++) {
        const td = document.createElement("td");

        const isBefore = week === 0 && dow < info.firstDow;
        const isAfter = day > info.days;

        if (isBefore || isAfter) {
          td.textContent = "";
          tr.appendChild(td);
          continue;
        }

        const dStr = dateStr(info.y, info.m, day);
        const mark = (state.attendance[dStr] || "").trim();
        // PDF용 셀에는 "일자\n표기" 스타일로 넣음
        td.innerHTML = `<div style="font-weight:700; font-size:0.92rem;">${day}</div>
                        <div style="margin-top:6px; font-size:0.92rem;">${escapeHtml(mark)}</div>`;
        tr.appendChild(td);

        day++;
      }

      tbody.appendChild(tr);
    }
  }

  function escapeHtml(s) {
    return String(s || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function computeScheduleHours() {
    const hoursByDow = {};
    let weekly = 0;
    for (let dow = 0; dow < 7; dow++) {
      const h = calcDayHours(state.schedule[dow]);
      hoursByDow[dow] = h;
      weekly += h;
    }
    weekly = Math.round(weekly * 100) / 100;
    return { hoursByDow, weekly };
  }

  function computeWorkStats() {
    const info = getMonthInfo(state.meta.payMonth);
    if (!info) {
      return {
        cntByDow: Array(7).fill(0),
        totalHours: 0,
        cntMonThu: 0,
        cntFri: 0,
      };
    }

    const { hoursByDow } = computeScheduleHours();
    const cntByDow = Array(7).fill(0);
    let totalHours = 0;

    for (let d = 1; d <= info.days; d++) {
      const ds = dateStr(info.y, info.m, d);
      const mark = (state.attendance[ds] || "").trim();
      const dow = new Date(info.y, info.m - 1, d).getDay();

      if (mark.toUpperCase() === "O") {
        cntByDow[dow]++;
        totalHours += (hoursByDow[dow] || 0);
      }
    }

    totalHours = Math.round(totalHours * 100) / 100;

    const cntMonThu = cntByDow[1] + cntByDow[2] + cntByDow[3] + cntByDow[4];
    const cntFri = cntByDow[5];

    return { cntByDow, totalHours, cntMonThu, cntFri };
  }

  function computePays() {
    const hourly = toInt(state.rate.hourly, 0);
    const { totalHours } = computeWorkStats();

    // 소수점 오차 방지용
    const basePayRaw = hourly * totalHours;
    const basePay = Math.round(basePayRaw);

    const items = clone(state.payItems);
    const allowancesSum =
      toInt(items.allowSeniority) +
      toInt(items.allowMeal) +
      toInt(items.allowDanger) +
      toInt(items.allowLicense) +
      toInt(items.allowSpecial) +
      toInt(items.allowMealOps) +
      toInt(items.allowFamily) +
      toInt(items.allowHolidayBonus);

    const sumA = basePay + allowancesSum;

    // 공제
    const d = clone(state.deductions);

    // 건강(요양포함)
    let healthTotal = 0;
    let healthEmployee = 0;
    let ltcEmployee = 0;

    if (d.health.enabled) {
      if (d.health.auto) {
        const rate = toFloat(d.health.rate, 0) / 100;      // 총
        const ltcRate = toFloat(d.health.ltcRate, 0) / 100;

        // 개인부담: 총 보험료의 50%로 가정
        healthEmployee = trunc10Won(sumA * rate * 0.5);
        ltcEmployee = trunc10Won(healthEmployee * ltcRate);
        healthTotal = trunc10Won(healthEmployee + ltcEmployee);

        d.health.amount = healthTotal;
      } else {
        healthTotal = trunc10Won(toInt(d.health.amount, 0));
        d.health.amount = healthTotal;
      }
    } else {
      d.health.amount = 0;
    }

    // 고용보험(개인)
    let empTotal = 0;
    if (d.employment.enabled) {
      if (d.employment.auto) {
        const r = toFloat(d.employment.rate, 0) / 100;
        empTotal = trunc10Won(sumA * r);
        d.employment.amount = empTotal;
      } else {
        empTotal = trunc10Won(toInt(d.employment.amount, 0));
        d.employment.amount = empTotal;
      }
    } else {
      d.employment.amount = 0;
    }

    // 기타 공제
    d.incomeTax = trunc10Won(toInt(d.incomeTax, 0));
    d.localTax = trunc10Won(toInt(d.localTax, 0));
    d.pension = trunc10Won(toInt(d.pension, 0));
    d.etcDed = trunc10Won(toInt(d.etcDed, 0));

    const sumB =
      d.incomeTax +
      d.localTax +
      d.pension +
      healthTotal +
      empTotal +
      d.etcDed;

    const net = sumA - sumB;

    return {
      hourly,
      totalHours,
      basePay,
      allowancesSum,
      sumA,
      sumB,
      net,
      ded: d,
      healthEmployee,
      ltcEmployee,
    };
  }

  function renderComputed() {
    const { hoursByDow, weekly } = computeScheduleHours();
    for (let dow = 0; dow < 7; dow++) {
      setText(`dow${dow}_hours`, String(hoursByDow[dow] || 0));
    }
    setText("weeklyHours", String(weekly));

    const stats = computeWorkStats();
    setText("cntMonThu", String(stats.cntMonThu));
    setText("cntFri", String(stats.cntFri));

    const calc = computePays();

    setText("totalHours", String(calc.totalHours));
    setText("viewHourly", formatWon(calc.hourly));
    setText("basePayView", formatWon(calc.basePay));

    setText("sumA", formatWon(calc.sumA));
    setText("sumB", formatWon(calc.sumB));
    setText("netPay", formatWon(calc.net));

    // 임금명세서 미리보기
    const info = getMonthInfo(state.meta.payMonth);
    const title = info ? `${info.y}년 ${info.m}월 임금명세서` : "임금명세서";
    setText("stubTitle", title);

    setText("stubSchool", state.meta.schoolName || "");
    setText("stubPayDate", state.meta.payDate || "");
    setText("stubName", state.meta.workerName || "");
    setText("stubJob", state.meta.jobTitle || "");
    setText("stubBirth", state.meta.birthDate || "");
    setText("stubFirstHire", state.meta.firstHireDate || "");

    // 급여내역 tbody
    const payTbody = $("payItemsTbody");
    payTbody.innerHTML = "";

    // 기본급(매월)
    payTbody.appendChild(makeRow3("매월", "기본급", calc.basePay));

    // 기타 수당 (매월로 처리)
    payTbody.appendChild(makeRow3("매월", "근속수당", toInt(state.payItems.allowSeniority)));
    payTbody.appendChild(makeRow3("매월", "정액급식비", toInt(state.payItems.allowMeal)));
    payTbody.appendChild(makeRow3("매월", "위험근무수당", toInt(state.payItems.allowDanger)));
    payTbody.appendChild(makeRow3("매월", "면허가산수당", toInt(state.payItems.allowLicense)));
    payTbody.appendChild(makeRow3("매월", "특수업무수당", toInt(state.payItems.allowSpecial)));
    payTbody.appendChild(makeRow3("매월", "급식운영수당", toInt(state.payItems.allowMealOps)));
    payTbody.appendChild(makeRow3("매월", "가족수당", toInt(state.payItems.allowFamily)));

    // 부정기(명절휴가비)
    payTbody.appendChild(makeRow3("부정기", "명절휴가비", toInt(state.payItems.allowHolidayBonus)));

    setText("stubSumA", formatWon(calc.sumA));

    // 공제내역 tbody
    const dedTbody = $("dedTbody");
    dedTbody.innerHTML = "";

    dedTbody.appendChild(makeRow3("세금", "소득세", calc.ded.incomeTax));
    dedTbody.appendChild(makeRow3("세금", "주민세", calc.ded.localTax));

    // 건강(요양포함) / 고용보험
    dedTbody.appendChild(makeRow3("사회보험", "건강보험(요양포함)", calc.ded.health.amount));
    dedTbody.appendChild(makeRow3("사회보험", "고용보험", calc.ded.employment.amount));

    // 국민연금/기타
    dedTbody.appendChild(makeRow3("사회보험", "국민연금", calc.ded.pension));
    dedTbody.appendChild(makeRow3("기타", "기타공제", calc.ded.etcDed));

    setText("stubSumB", formatWon(calc.sumB));
    setText("stubNet", formatWon(calc.net));

    // PDF 1페이지(출근내역+요약)
    const attTitle = info ? `${info.y}년 ${info.m}월 출입문개폐요원 출근 내역` : "출근내역";
    setText("attTitle", attTitle);
    setText("attSchool", state.meta.schoolName || "");
    setText("attName", state.meta.workerName || "");
    setText("attHourly", `${formatWon(calc.hourly)} 원`);
    setText("attWeeklyHours", String(weekly));
    setText("attCntMonThu", String(stats.cntMonThu));
    setText("attCntFri", String(stats.cntFri));
    setText("attTotalHours", String(calc.totalHours));
    setText("attBasePay", `${formatWon(calc.basePay)} 원`);

    setText("attScheduleText", buildScheduleText());

    buildPdfCalendarMirror();

    function makeRow3(kind, label, amount) {
      const tr = document.createElement("tr");
      const td1 = document.createElement("td");
      td1.textContent = kind;
      const td2 = document.createElement("td");
      td2.textContent = label;
      const td3 = document.createElement("td");
      td3.className = "num";
      td3.textContent = formatWon(amount || 0);
      tr.append(td1, td2, td3);
      return tr;
    }
  }

  function buildScheduleText() {
    // 보기 좋게 월~목/금 형태로 요약
    const mon = state.schedule[1];
    const fri = state.schedule[5];

    const monTxt = timeText(mon);
    const friTxt = timeText(fri);

    const monH = calcDayHours(mon);
    const tueH = calcDayHours(state.schedule[2]);
    const wedH = calcDayHours(state.schedule[3]);
    const thuH = calcDayHours(state.schedule[4]);
    const friH = calcDayHours(fri);

    const monThuSame = (monH === tueH && monH === wedH && monH === thuH);

    const a = [];
    if (monThuSame && monH > 0) {
      a.push(`·월~목: 1일 ${monH}시간 (${monTxt})`);
    } else {
      // 요일별 표시
      for (let d = 1; d <= 4; d++) {
        const h = calcDayHours(state.schedule[d]);
        if (h > 0) a.push(`·${DOW_LABELS[d]}: 1일 ${h}시간 (${timeText(state.schedule[d])})`);
      }
    }
    if (friH > 0) {
      a.push(`·금: 1일 ${friH}시간 (${friTxt})`);
    }
    return a.join("\n");
  }

  function timeText(day) {
    const p1 = (day.s1 && day.e1) ? `${day.s1}~${day.e1}` : "";
    const p2 = (day.s2 && day.e2) ? `${day.s2}~${day.e2}` : "";
    if (p1 && p2) return `${p1}, ${p2}`;
    if (p1) return p1;
    if (p2) return p2;
    return "시간대 없음";
  }

  function recomputeAndRender(noRebuildCalendar = false) {
    // 스케줄/공제 자동계산을 UI에 반영(자동일 때)
    const calc = computePays();

    // 건강/고용 auto이면 amount input에 반영
    if ($("healthEnabled").checked && $("healthAuto").value === "Y") {
      $("healthAmount").value = String(calc.ded.health.amount || 0);
    }
    if ($("empEnabled").checked && $("empAuto").value === "Y") {
      $("empAmount").value = String(calc.ded.employment.amount || 0);
    }

    // 달력 재생성 필요시(지급월 변경 등)
    if (!noRebuildCalendar && renderedMonth !== state.meta.payMonth) {
      buildCalendar();
      renderedMonth = state.meta.payMonth;
    }

    renderComputed();
    saveState();
  }

  function bindMeta() {
    $("schoolName").addEventListener("input", (e) => {
      state.meta.schoolName = e.target.value || "";
      recomputeAndRender(true);
    });
    $("docNo").addEventListener("input", (e) => {
      state.meta.docNo = e.target.value || "";
      recomputeAndRender(true);
    });
    $("payMonth").addEventListener("input", (e) => {
      const v = e.target.value || "";
      state.meta.payMonth = v;
      // 달력 표기는 “월 변경 시” 유지할지 애매하지만: 기본은 유지하지 않고 새 달력로 전환
      // 단, 같은 yyyy-mm로 보정되면 유지
      recomputeAndRender(false);
    });
    $("payDate").addEventListener("input", (e) => {
      state.meta.payDate = e.target.value || "";
      recomputeAndRender(true);
    });

    $("workerName").addEventListener("input", (e) => {
      state.meta.workerName = e.target.value || "";
      recomputeAndRender(true);
    });
    $("jobTitle").addEventListener("input", (e) => {
      state.meta.jobTitle = e.target.value || "";
      recomputeAndRender(true);
    });
    $("birthDate").addEventListener("input", (e) => {
      state.meta.birthDate = e.target.value || "";
      recomputeAndRender(true);
    });
    $("firstHireDate").addEventListener("input", (e) => {
      state.meta.firstHireDate = e.target.value || "";
      recomputeAndRender(true);
    });
  }

  function bindRate() {
    $("hourlyRate").addEventListener("input", (e) => {
      state.rate.hourly = toInt(e.target.value, 0);
      recomputeAndRender(true);
    });

    $("btnSet2025").addEventListener("click", () => {
      state.rate.hourly = 11200;
      $("hourlyRate").value = "11200";
      recomputeAndRender(true);
    });

    $("btnSet2026").addEventListener("click", () => {
      state.rate.hourly = 11500;
      $("hourlyRate").value = "11500";
      recomputeAndRender(true);
    });
  }

  function bindScheduleOptions() {
    $("chkMonThuSame").addEventListener("change", (e) => {
      state.options.monThuSame = !!e.target.checked;
      saveState();
    });

    $("btnCopyMonToTueThu").addEventListener("click", () => {
      for (const key of ["s1", "e1", "s2", "e2"]) {
        const val = state.schedule[1][key];
        for (let d = 2; d <= 4; d++) {
          state.schedule[d][key] = val;
          const el = $(`dow${d}_${key}`);
          if (el) el.value = val || "";
        }
      }
      recomputeAndRender(true);
    });

    $("btnPresetMoonmak").addEventListener("click", () => {
      applyMoonmakPreset();
      recomputeAndRender(false);
    });
  }

  function applyMoonmakPreset() {
    // 월~목
    for (let d = 1; d <= 4; d++) {
      state.schedule[d] = { s1: "07:10", e1: "08:40", s2: "16:40", e2: "18:10" };
    }
    // 금
    state.schedule[5] = { s1: "07:40", e1: "08:40", s2: "16:40", e2: "17:40" };
    // 일/토
    state.schedule[0] = { s1: "", e1: "", s2: "", e2: "" };
    state.schedule[6] = { s1: "", e1: "", s2: "", e2: "" };

    // UI 반영
    for (let dow = 0; dow < 7; dow++) {
      for (const key of ["s1", "e1", "s2", "e2"]) {
        const el = $(`dow${dow}_${key}`);
        if (el) el.value = state.schedule[dow][key] || "";
      }
    }
  }

  function bindAllowances() {
    const map = [
      ["allowSeniority", "allowSeniority"],
      ["allowMeal", "allowMeal"],
      ["allowDanger", "allowDanger"],
      ["allowLicense", "allowLicense"],
      ["allowSpecial", "allowSpecial"],
      ["allowMealOps", "allowMealOps"],
      ["allowFamily", "allowFamily"],
      ["allowHolidayBonus", "allowHolidayBonus"],
    ];

    for (const [id, key] of map) {
      $(id).addEventListener("input", (e) => {
        state.payItems[key] = trunc10Won(toInt(e.target.value, 0));
        // input에도 절삭된 값 반영
        e.target.value = String(state.payItems[key]);
        recomputeAndRender(true);
      });
    }
  }

  function bindDeductions() {
    // 기타 공제
    const nums = [
      ["incomeTax", "incomeTax"],
      ["localTax", "localTax"],
      ["pension", "pension"],
      ["etcDed", "etcDed"],
    ];
    for (const [id, key] of nums) {
      $(id).addEventListener("input", (e) => {
        state.deductions[key] = trunc10Won(toInt(e.target.value, 0));
        e.target.value = String(state.deductions[key]);
        recomputeAndRender(true);
      });
    }

    // 건강(요양포함)
    $("healthEnabled").addEventListener("change", (e) => {
      state.deductions.health.enabled = !!e.target.checked;
      // 미공제면 0으로 고정
      if (!state.deductions.health.enabled) {
        state.deductions.health.amount = 0;
        $("healthAmount").value = "0";
      }
      syncHealthInputs();
      recomputeAndRender(true);
    });

    $("healthRate").addEventListener("input", (e) => {
      state.deductions.health.rate = toFloat(e.target.value, 0);
      recomputeAndRender(true);
    });
    $("ltcRate").addEventListener("input", (e) => {
      state.deductions.health.ltcRate = toFloat(e.target.value, 0);
      recomputeAndRender(true);
    });
    $("healthAuto").addEventListener("change", (e) => {
      state.deductions.health.auto = (e.target.value === "Y");
      syncHealthInputs();
      recomputeAndRender(true);
    });
    $("healthAmount").addEventListener("input", (e) => {
      // 수기일 때만 반영
      if (!state.deductions.health.auto) {
        state.deductions.health.amount = trunc10Won(toInt(e.target.value, 0));
        e.target.value = String(state.deductions.health.amount);
        recomputeAndRender(true);
      }
    });

    // 고용보험
    $("empEnabled").addEventListener("change", (e) => {
      state.deductions.employment.enabled = !!e.target.checked;
      if (!state.deductions.employment.enabled) {
        state.deductions.employment.amount = 0;
        $("empAmount").value = "0";
      }
      syncEmpInputs();
      recomputeAndRender(true);
    });

    $("empRate").addEventListener("input", (e) => {
      state.deductions.employment.rate = toFloat(e.target.value, 0);
      recomputeAndRender(true);
    });
    $("empAuto").addEventListener("change", (e) => {
      state.deductions.employment.auto = (e.target.value === "Y");
      syncEmpInputs();
      recomputeAndRender(true);
    });
    $("empAmount").addEventListener("input", (e) => {
      if (!state.deductions.employment.auto) {
        state.deductions.employment.amount = trunc10Won(toInt(e.target.value, 0));
        e.target.value = String(state.deductions.employment.amount);
        recomputeAndRender(true);
      }
    });

    function syncHealthInputs() {
      const enabled = state.deductions.health.enabled;
      $("healthRate").disabled = !enabled;
      $("ltcRate").disabled = !enabled;
      $("healthAuto").disabled = !enabled;
      $("healthAmount").disabled = !enabled || state.deductions.health.auto;
    }
    function syncEmpInputs() {
      const enabled = state.deductions.employment.enabled;
      $("empRate").disabled = !enabled;
      $("empAuto").disabled = !enabled;
      $("empAmount").disabled = !enabled || state.deductions.employment.auto;
    }

    // 초기 동기화
    syncHealthInputs();
    syncEmpInputs();
  }

  function bindCalendarButtons() {
    $("btnPrevMonth").addEventListener("click", () => {
      const mp = monthParts(state.meta.payMonth);
      if (!mp) return;
      const d = new Date(mp.y, mp.m - 2, 1);
      const next = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
      state.meta.payMonth = next;
      $("payMonth").value = next;
      recomputeAndRender(false);
    });

    $("btnNextMonth").addEventListener("click", () => {
      const mp = monthParts(state.meta.payMonth);
      if (!mp) return;
      const d = new Date(mp.y, mp.m, 1);
      const next = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
      state.meta.payMonth = next;
      $("payMonth").value = next;
      recomputeAndRender(false);
    });

    $("btnAutoFillBlank").addEventListener("click", () => {
      autoFillMarks({ overwrite: false });
      recomputeAndRender(true);
      // UI input 값 반영
      syncCalendarInputsFromState();
    });

    $("btnAutoFillOverwrite").addEventListener("click", () => {
      autoFillMarks({ overwrite: true });
      recomputeAndRender(true);
      syncCalendarInputsFromState();
    });

    $("btnClearMarks").addEventListener("click", () => {
      clearMonthMarks();
      recomputeAndRender(true);
      syncCalendarInputsFromState();
    });
  }

  function syncCalendarInputsFromState() {
    const inputs = document.querySelectorAll('input.cal-mark[data-date]');
    for (const inp of inputs) {
      const ds = inp.dataset.date;
      inp.value = state.attendance[ds] || "";
    }
  }

  function autoFillMarks({ overwrite }) {
    const info = getMonthInfo(state.meta.payMonth);
    if (!info) return;
    const { hoursByDow } = computeScheduleHours();

    for (let d = 1; d <= info.days; d++) {
      const ds = dateStr(info.y, info.m, d);
      const dow = new Date(info.y, info.m - 1, d).getDay();
      const h = hoursByDow[dow] || 0;

      if (h <= 0) continue;

      if (!overwrite) {
        const cur = (state.attendance[ds] || "").trim();
        if (cur) continue;
      }
      state.attendance[ds] = "O";
    }
  }

  function clearMonthMarks() {
    const info = getMonthInfo(state.meta.payMonth);
    if (!info) return;
    for (let d = 1; d <= info.days; d++) {
      const ds = dateStr(info.y, info.m, d);
      delete state.attendance[ds];
    }
  }

  function bindExcelAndPdf() {
    $("btnDownloadTemplate").addEventListener("click", () => {
      downloadInputTemplate();
    });

    $("excelUpload").addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importFromExcel(file);
        applyStateToUI();
        recomputeAndRender(false);
      } catch (err) {
        alert("엑셀 업로드 처리 중 오류가 발생했습니다.\n서식이 다르면 읽지 못할 수 있습니다.\n\n" + String(err));
      } finally {
        e.target.value = "";
      }
    });

    $("btnExportXlsx").addEventListener("click", () => {
      exportToXlsx();
    });

    $("btnExportPdf").addEventListener("click", async () => {
      await exportToPdf();
    });

    $("btnReset").addEventListener("click", () => {
      if (!confirm("정말 초기화할까요?\n(브라우저 로컬저장 데이터도 삭제됩니다)")) return;
      localStorage.removeItem(STORAGE_KEY);
      state = defaultState();
      applyStateToUI();
      recomputeAndRender(false);
    });
  }

  function downloadInputTemplate() {
    const mp = getMonthInfo(state.meta.payMonth) || getMonthInfo(currentMonthStr());
    const payMonth = mp ? `${mp.y}-${pad2(mp.m)}` : currentMonthStr();

    const aoa = [];
    aoa.push(["출입문개폐전담원 인건비 입력서식"]);
    aoa.push([]);
    aoa.push(["학교명", state.meta.schoolName || ""]);
    aoa.push(["지급월(YYYY-MM)", payMonth]);
    aoa.push(["지급일자(YYYY-MM-DD)", state.meta.payDate || todayStr()]);
    aoa.push(["성명", state.meta.workerName || ""]);
    aoa.push(["직종", state.meta.jobTitle || "출입문개폐전담원"]);
    aoa.push(["생년월일(YYYY-MM-DD)", state.meta.birthDate || ""]);
    aoa.push(["최초임용일(YYYY-MM-DD)", state.meta.firstHireDate || ""]);
    aoa.push(["시급(원)", String(state.rate.hourly || 11500)]);
    aoa.push(["문서번호/비고", state.meta.docNo || ""]);
    aoa.push([]);
    aoa.push(["건강보험 공제여부(Y/N)", state.deductions.health.enabled ? "Y" : "N"]);
    aoa.push(["건강보험료율(%)", String(state.deductions.health.rate ?? 0)]);
    aoa.push(["요양보험료율(%)", String(state.deductions.health.ltcRate ?? 0)]);
    aoa.push(["고용보험 공제여부(Y/N)", state.deductions.employment.enabled ? "Y" : "N"]);
    aoa.push(["고용보험 개인요율(%)", String(state.deductions.employment.rate ?? 0)]);
    aoa.push(["소득세(원)", String(state.deductions.incomeTax ?? 0)]);
    aoa.push(["주민세(원)", String(state.deductions.localTax ?? 0)]);
    aoa.push(["국민연금(원)", String(state.deductions.pension ?? 0)]);
    aoa.push(["기타공제(원)", String(state.deductions.etcDed ?? 0)]);
    aoa.push([]);
    aoa.push(["요일별 근무시간대"]);
    aoa.push(["요일", "1차 시작", "1차 종료", "2차 시작", "2차 종료"]);
    for (let dow = 0; dow < 7; dow++) {
      const s = state.schedule[dow];
      aoa.push([DOW_LABELS[dow], s.s1 || "", s.e1 || "", s.s2 || "", s.e2 || ""]);
    }
    aoa.push([]);
    aoa.push(["출근내역"]);
    aoa.push(["날짜(YYYY-MM-DD)", "표기(O/휴가/연휴/공휴일 등)"]);

    if (mp) {
      for (let d = 1; d <= mp.days; d++) {
        const ds = dateStr(mp.y, mp.m, d);
        aoa.push([ds, state.attendance[ds] || ""]);
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
    ws["!cols"] = [{ wch: 26 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "입력");

    const fname = `${payMonth}_출입문개폐전담원_입력서식.xlsx`;
    XLSX.writeFile(wb, fname);
  }

  async function importFromExcel(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    // 1) 우선순위: "입력" 시트
    const inputSheetName = wb.SheetNames.find(n => n === "입력" || n.toUpperCase() === "INPUT");
    if (inputSheetName) {
      const ws = wb.Sheets[inputSheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      parseInputTemplateAoa(aoa);
      return;
    }

    // 2) 차선: 첫 시트에서 키-값 추정
    const ws0 = wb.Sheets[wb.SheetNames[0]];
    const aoa0 = XLSX.utils.sheet_to_json(ws0, { header: 1, defval: "" });
    const ok = parseKeyValueLoose(aoa0);

    if (!ok) {
      throw new Error("지원되는 입력서식(입력/INPUT) 또는 키-값 패턴을 찾지 못했습니다.");
    }
  }

  function parseInputTemplateAoa(aoa) {
    const getKV = (key) => {
      for (const row of aoa) {
        if (!row || row.length < 2) continue;
        if (String(row[0]).trim() === key) return String(row[1]).trim();
      }
      return "";
    };

    state.meta.schoolName = getKV("학교명");
    state.meta.payMonth = getKV("지급월(YYYY-MM)") || state.meta.payMonth;
    state.meta.payDate = getKV("지급일자(YYYY-MM-DD)") || state.meta.payDate;
    state.meta.workerName = getKV("성명");
    state.meta.jobTitle = getKV("직종") || state.meta.jobTitle;
    state.meta.birthDate = getKV("생년월일(YYYY-MM-DD)");
    state.meta.firstHireDate = getKV("최초임용일(YYYY-MM-DD)");
    state.meta.docNo = getKV("문서번호/비고");

    state.rate.hourly = toInt(getKV("시급(원)"), state.rate.hourly);

    // 공제
    state.deductions.health.enabled = (getKV("건강보험 공제여부(Y/N)").toUpperCase() === "Y");
    state.deductions.health.rate = toFloat(getKV("건강보험료율(%)"), state.deductions.health.rate);
    state.deductions.health.ltcRate = toFloat(getKV("요양보험료율(%)"), state.deductions.health.ltcRate);

    state.deductions.employment.enabled = (getKV("고용보험 공제여부(Y/N)").toUpperCase() === "Y");
    state.deductions.employment.rate = toFloat(getKV("고용보험 개인요율(%)"), state.deductions.employment.rate);

    state.deductions.incomeTax = trunc10Won(toInt(getKV("소득세(원)"), 0));
    state.deductions.localTax = trunc10Won(toInt(getKV("주민세(원)"), 0));
    state.deductions.pension = trunc10Won(toInt(getKV("국민연금(원)"), 0));
    state.deductions.etcDed = trunc10Won(toInt(getKV("기타공제(원)"), 0));

    // 스케줄 표 찾기
    const idxHeader = aoa.findIndex(r => String(r?.[0] || "").trim() === "요일" && String(r?.[1] || "").includes("시작"));
    if (idxHeader >= 0) {
      for (let i = idxHeader + 1; i < idxHeader + 8; i++) {
        const r = aoa[i];
        const dowLabel = String(r?.[0] || "").trim();
        const dow = DOW_LABELS.indexOf(dowLabel);
        if (dow < 0) continue;
        state.schedule[dow] = {
          s1: String(r?.[1] || "").trim(),
          e1: String(r?.[2] || "").trim(),
          s2: String(r?.[3] || "").trim(),
          e2: String(r?.[4] || "").trim(),
        };
      }
    }

    // 출근내역 표 찾기
    const idxAtt = aoa.findIndex(r => String(r?.[0] || "").trim().startsWith("날짜"));
    if (idxAtt >= 0) {
      // payMonth에 맞춰 attendance 초기화 후 채움
      state.attendance = {};
      for (let i = idxAtt + 1; i < aoa.length; i++) {
        const r = aoa[i];
        const ds = String(r?.[0] || "").trim();
        const mark = String(r?.[1] || "").trim();
        if (!ds) continue;
        if (!mark) continue;
        // YYYY-MM-DD 형태면 그대로
        state.attendance[ds] = mark;
      }
    }
  }

  function parseKeyValueLoose(aoa) {
    // 아주 느슨한 추정: "학교명", "지급월" 같은 라벨이 어딘가 있으면 읽는다.
    // 완벽한 호환은 불가하지만, 최소한 템플릿을 못 쓸 때 안전장치로.
    const map = {};
    for (const row of aoa.slice(0, 60)) {
      const k = String(row?.[0] || "").trim();
      const v = String(row?.[1] || "").trim();
      if (!k) continue;
      if (["학교명", "지급월", "지급월(YYYY-MM)", "성명", "직종"].includes(k)) {
        map[k] = v;
      }
    }
    const any = Object.keys(map).length > 0;
    if (!any) return false;

    if (map["학교명"]) state.meta.schoolName = map["학교명"];
    if (map["성명"]) state.meta.workerName = map["성명"];
    if (map["직종"]) state.meta.jobTitle = map["직종"];
    if (map["지급월(YYYY-MM)"]) state.meta.payMonth = map["지급월(YYYY-MM)"];
    if (map["지급월"]) state.meta.payMonth = map["지급월"];

    return true;
  }

  function exportToXlsx() {
    const info = getMonthInfo(state.meta.payMonth);
    const mpLabel = info ? `${info.y}.${info.m}월` : state.meta.payMonth;

    const wb = XLSX.utils.book_new();
    const ws1 = buildSheet1AttendanceCalendar();
    const ws2 = buildSheet2Calc();
    const ws3 = buildSheet3PayStub();

    XLSX.utils.book_append_sheet(wb, ws1, "Sheet1_출근내역");
    XLSX.utils.book_append_sheet(wb, ws2, "Sheet2_산정내역");
    XLSX.utils.book_append_sheet(wb, ws3, "Sheet3_임금명세서");

    const fileName = `${mpLabel}_출입문개폐전담원_출근내역_산정_임금명세서.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  function buildSheet1AttendanceCalendar() {
    const info = getMonthInfo(state.meta.payMonth);
    const calc = computePays();
    const stats = computeWorkStats();

    const title = info ? `${info.y}년 ${info.m}월 출입문개폐요원 출근 내역` : "출근 내역";

    const aoa = [];
    aoa.push([]);
    aoa.push(["", "", title, "", "", "", ""]);
    aoa.push([]);

    // 헤더
    aoa.push(["", "일", "월", "화", "수", "목", "금", "토", "", "", "월~목", "금"]);
    // 첫 주부터 6주(각 주 2행: 날짜행/표기행)
    if (!info) {
      aoa.push([]);
    } else {
      let day = 1;
      for (let w = 0; w < 6; w++) {
        const dateRow = Array(12).fill("");
        const markRow = Array(12).fill("");
        dateRow[0] = "";
        markRow[0] = "";

        for (let dow = 0; dow < 7; dow++) {
          const isBefore = w === 0 && dow < info.firstDow;
          const isAfter = day > info.days;
          const col = 1 + dow;

          if (isBefore || isAfter) {
            dateRow[col] = "";
            markRow[col] = "";
            continue;
          }

          dateRow[col] = `${info.m}/${day}`;
          const ds = dateStr(info.y, info.m, day);
          markRow[col] = state.attendance[ds] || "";
          day++;
        }

        // 우측 요약(첫 주 날짜행에만 일부 표시해도 되지만, 단순하게 고정 위치에 표시)
        if (w === 0) {
          dateRow[9] = "당월 근무일수";
          dateRow[10] = stats.cntMonThu;
          dateRow[11] = stats.cntFri;

          markRow[9] = "1일 근로시간";
          // 월~목 평균/대표: 월 근로시간
          const monH = calcDayHours(state.schedule[1]);
          const friH = calcDayHours(state.schedule[5]);
          markRow[10] = monH;
          markRow[11] = friH;
        }

        aoa.push(dateRow);
        aoa.push(markRow);
      }

      // 아래쪽 요약
      aoa.push([]);
      aoa.push(["", "", "", "", "", "", "", "", "", "당월 유급 근로시간", calc.totalHours, "(시간)"]);
      aoa.push(["", "", "", "", "", "", "", "", "", "당월 보수 지급액", calc.basePay, "(원)"]);
      aoa.push(["", "", "", "", "", "", "", "", "", "시급", calc.hourly, "(원)"]);
      aoa.push([]);
      aoa.push(["", "", "", "", "", "", "", "", "", "근로계약서 상 근로시간(주)", computeScheduleHours().weekly, "(시간)"]);
      aoa.push(["", "", "", "", "", "", "", "", "", "비고", state.meta.docNo || ""]);
      aoa.push([]);
      const scheduleTxt = buildScheduleText().split("\n");
      for (const line of scheduleTxt) {
        aoa.push(["", "", "", "", "", "", "", "", "", line]);
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 3 },  // A
      { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, // B~H
      { wch: 2 },  // I
      { wch: 18 }, { wch: 12 }, { wch: 10 }, // J~L
    ];
    ws["!merges"] = [
      { s: { r: 1, c: 2 }, e: { r: 1, c: 6 } }, // 제목 병합(대충)
    ];
    return ws;
  }

  function buildSheet2Calc() {
    const info = getMonthInfo(state.meta.payMonth);
    const calc = computePays();
    const { hoursByDow, weekly } = computeScheduleHours();

    const aoa = [];
    aoa.push(["인건비 산정 내역"]);
    aoa.push([]);
    aoa.push(["학교명", state.meta.schoolName || ""]);
    aoa.push(["지급월", state.meta.payMonth || ""]);
    aoa.push(["지급일자", state.meta.payDate || ""]);
    aoa.push(["성명", state.meta.workerName || ""]);
    aoa.push(["직종", state.meta.jobTitle || ""]);
    aoa.push([]);
    aoa.push(["요일", "근무일수(O)", "1일 근로시간", "당월 유급 근로시간"]);

    const stats = computeWorkStats();
    let sumDays = 0;
    let sumHours = 0;

    for (let dow = 0; dow < 7; dow++) {
      const days = stats.cntByDow[dow] || 0;
      const dh = hoursByDow[dow] || 0;
      const mh = Math.round(days * dh * 100) / 100;
      aoa.push([DOW_LABELS[dow], days, dh, mh]);
      sumDays += days;
      sumHours += mh;
    }
    sumHours = Math.round(sumHours * 100) / 100;

    aoa.push(["합계", sumDays, "", sumHours]);
    aoa.push([]);
    aoa.push(["주(요일패턴) 근로시간", weekly, "시간"]);
    aoa.push(["당월 유급 근로시간", calc.totalHours, "시간"]);
    aoa.push(["시급", calc.hourly, "원"]);
    aoa.push(["기본급(시급*유급근로시간)", calc.basePay, "원"]);
    aoa.push(["수당 합계", calc.allowancesSum, "원"]);
    aoa.push(["급여총액 계(A)", calc.sumA, "원"]);
    aoa.push(["공제액 계(B)", calc.sumB, "원"]);
    aoa.push(["실수령액(A-B)", calc.net, "원"]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    ws["!cols"] = [{ wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 20 }];
    return ws;
  }

  function buildSheet3PayStub() {
    const info = getMonthInfo(state.meta.payMonth);
    const calc = computePays();

    const title = info ? `${info.y}년 ${info.m}월 임금명세서` : "임금명세서";

    const aoa = [];
    aoa.push([title]);
    aoa.push([]);
    aoa.push(["소속", state.meta.schoolName || "", "지급일", state.meta.payDate || ""]);
    aoa.push(["성명", state.meta.workerName || "", "직종", state.meta.jobTitle || ""]);
    aoa.push(["생년월일", state.meta.birthDate || "", "최초임용일", state.meta.firstHireDate || ""]);
    aoa.push([]);
    aoa.push(["급여내역"]);
    aoa.push(["구분", "임금항목", "산출식(참고)", "금액(원)"]);
    aoa.push(["매월", "기본급", `시급 ${formatWon(calc.hourly)}원 * ${calc.totalHours}시간`, calc.basePay]);
    aoa.push(["매월", "근속수당", "", toInt(state.payItems.allowSeniority)]);
    aoa.push(["매월", "정액급식비", "", toInt(state.payItems.allowMeal)]);
    aoa.push(["매월", "위험근무수당", "", toInt(state.payItems.allowDanger)]);
    aoa.push(["매월", "면허가산수당", "", toInt(state.payItems.allowLicense)]);
    aoa.push(["매월", "특수업무수당", "", toInt(state.payItems.allowSpecial)]);
    aoa.push(["매월", "급식운영수당", "", toInt(state.payItems.allowMealOps)]);
    aoa.push(["매월", "가족수당", "", toInt(state.payItems.allowFamily)]);
    aoa.push(["부정기", "명절휴가비", "", toInt(state.payItems.allowHolidayBonus)]);
    aoa.push([]);
    aoa.push(["급여총액 계(A)", calc.sumA]);
    aoa.push([]);
    aoa.push(["공제내역"]);
    aoa.push(["구분", "항목", "", "금액(원)"]);
    aoa.push(["세금", "소득세", "", calc.ded.incomeTax]);
    aoa.push(["세금", "주민세", "", calc.ded.localTax]);
    aoa.push(["사회보험", "건강보험(요양포함)", "", calc.ded.health.amount]);
    aoa.push(["사회보험", "고용보험", "", calc.ded.employment.amount]);
    aoa.push(["사회보험", "국민연금", "", calc.ded.pension]);
    aoa.push(["기타", "기타공제", "", calc.ded.etcDed]);
    aoa.push([]);
    aoa.push(["공제액 계(B)", calc.sumB]);
    aoa.push([]);
    aoa.push(["실수령액(A-B)", calc.net]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 30 }, { wch: 14 }];
    return ws;
  }

  async function exportToPdf() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "p", unit: "pt", format: "a4" });

    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    // 1페이지: 출근내역+요약+달력
    const el1 = $("pdfPage1");
    const c1 = await html2canvas(el1, { scale: 2, backgroundColor: "#ffffff" });
    const img1 = c1.toDataURL("image/png");
    const r1 = Math.min(pageW / c1.width, pageH / c1.height);
    const w1 = c1.width * r1;
    const h1 = c1.height * r1;
    pdf.addImage(img1, "PNG", (pageW - w1) / 2, (pageH - h1) / 2, w1, h1);

    // 2페이지: 임금명세서
    pdf.addPage();
    const el2 = $("pdfPage2");
    const c2 = await html2canvas(el2, { scale: 2, backgroundColor: "#ffffff" });
    const img2 = c2.toDataURL("image/png");
    const r2 = Math.min(pageW / c2.width, pageH / c2.height);
    const w2 = c2.width * r2;
    const h2 = c2.height * r2;
    pdf.addImage(img2, "PNG", (pageW - w2) / 2, (pageH - h2) / 2, w2, h2);

    const info = getMonthInfo(state.meta.payMonth);
    const mpLabel = info ? `${info.y}.${info.m}월` : state.meta.payMonth;
    const fname = `${mpLabel}_출입문개폐전담원_출근내역_임금명세서.pdf`;
    pdf.save(fname);
  }

  // 방향키로 셀 이동(엑셀 느낌)
  function enableGridNavigation() {
    document.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLElement)) return;

      // cal/sched grid에만 적용
      const grid = t.dataset.grid;
      if (!grid) return;

      const r = Number(t.dataset.r);
      const c = Number(t.dataset.c);
      if (!Number.isFinite(r) || !Number.isFinite(c)) return;

      // 입력 중 커서 이동을 존중: 좌/우는 커서가 끝에 있을 때만 셀 이동
      if ((e.key === "ArrowLeft" || e.key === "ArrowRight") && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) {
        const inp = /** @type {HTMLInputElement|HTMLTextAreaElement} */ (t);
        if (inp.selectionStart != null && inp.selectionEnd != null) {
          const len = inp.value.length;
          const start = inp.selectionStart;
          const end = inp.selectionEnd;

          if (e.key === "ArrowLeft" && start > 0) return; // 커서 이동
          if (e.key === "ArrowRight" && end < len) return; // 커서 이동
        }
      }

      let nr = r, nc = c;
      if (e.key === "ArrowUp") nr = r - 1;
      else if (e.key === "ArrowDown") nr = r + 1;
      else if (e.key === "ArrowLeft") nc = c - 1;
      else if (e.key === "ArrowRight") nc = c + 1;
      else if (e.key === "Enter") nr = r + (e.shiftKey ? -1 : 1);
      else return;

      const next = document.querySelector(`[data-grid="${grid}"][data-r="${nr}"][data-c="${nc}"]`);
      if (next instanceof HTMLElement) {
        e.preventDefault();
        next.focus();
        if (next instanceof HTMLInputElement || next instanceof HTMLTextAreaElement) {
          next.select?.();
        }
      }
    });
  }

  function applyStateToUI() {
    // meta
    setInputValue("schoolName", state.meta.schoolName);
    setInputValue("docNo", state.meta.docNo);
    setInputValue("payMonth", state.meta.payMonth);
    setInputValue("payDate", state.meta.payDate);
    setInputValue("workerName", state.meta.workerName);
    setInputValue("jobTitle", state.meta.jobTitle);
    setInputValue("birthDate", state.meta.birthDate);
    setInputValue("firstHireDate", state.meta.firstHireDate);

    // rate
    setInputValue("hourlyRate", String(state.rate.hourly || 0));

    // options
    $("chkMonThuSame").checked = !!state.options.monThuSame;

    // schedule inputs are built later, but value will be set in buildScheduleTable()
    // deductions
    $("healthEnabled").checked = !!state.deductions.health.enabled;
    $("healthRate").value = String(state.deductions.health.rate ?? 0);
    $("ltcRate").value = String(state.deductions.health.ltcRate ?? 0);
    $("healthAuto").value = state.deductions.health.auto ? "Y" : "N";
    $("healthAmount").value = String(state.deductions.health.amount ?? 0);

    $("empEnabled").checked = !!state.deductions.employment.enabled;
    $("empRate").value = String(state.deductions.employment.rate ?? 0);
    $("empAuto").value = state.deductions.employment.auto ? "Y" : "N";
    $("empAmount").value = String(state.deductions.employment.amount ?? 0);

    // other deductions
    $("incomeTax").value = String(state.deductions.incomeTax ?? 0);
    $("localTax").value = String(state.deductions.localTax ?? 0);
    $("pension").value = String(state.deductions.pension ?? 0);
    $("etcDed").value = String(state.deductions.etcDed ?? 0);

    // allowances
    $("allowSeniority").value = String(state.payItems.allowSeniority ?? 0);
    $("allowMeal").value = String(state.payItems.allowMeal ?? 0);
    $("allowDanger").value = String(state.payItems.allowDanger ?? 0);
    $("allowLicense").value = String(state.payItems.allowLicense ?? 0);
    $("allowSpecial").value = String(state.payItems.allowSpecial ?? 0);
    $("allowMealOps").value = String(state.payItems.allowMealOps ?? 0);
    $("allowFamily").value = String(state.payItems.allowFamily ?? 0);
    $("allowHolidayBonus").value = String(state.payItems.allowHolidayBonus ?? 0);
  }

  function init() {
    // 초기 UI값 적용
    applyStateToUI();

    // 스케줄/달력 생성
    buildScheduleTable();
    buildCalendar();
    renderedMonth = state.meta.payMonth;

    // 바인딩
    bindMeta();
    bindRate();
    bindScheduleOptions();
    bindAllowances();
    bindDeductions();
    bindCalendarButtons();
    bindExcelAndPdf();

    // 방향키 이동
    enableGridNavigation();

    // 최초 렌더
    recomputeAndRender(true);
  }

  document.addEventListener("DOMContentLoaded", init);
})();
