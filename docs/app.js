/* Xiao Liu Ren Explorer
 * Full client-side application logic for a static GitHub Pages site.
 *
 * Features supported:
 * - tab navigation
 * - jump buttons and footer jumps
 * - theme toggle with persistence
 * - scroll to top
 * - live current date/time display
 * - current branch display
 * - current lunar display
 * - today view with cards and timeline
 * - lookup by date with filters
 * - manual classic calculator
 * - Gregorian -> lunar conversion calculator
 * - current moment calculation
 * - compare two times
 * - cycle diagram highlighting
 * - branch clock highlighting
 * - dataset browser with filters, sorting, pagination
 * - dataset detail panel
 * - JSON and CSV export of filtered rows
 * - statistics dashboard
 * - learning/FAQ accordions
 * - recent history via localStorage
 * - copy/share helpers
 * - toast notifications
 *
 * Notes:
 * - This app reads pre-generated JSON from xiao_liuren_data.json.
 * - Gregorian -> lunar conversion uses Intl Chinese calendar if available,
 *   and falls back to dataset date lookup when possible.
 * - This is an educational tool and not scientific proof.
 */

"use strict";

/* ==========================================================================
   CONSTANTS
   ========================================================================== */

const DATA_URL = "xiao_liuren_data.json";
const STORAGE_THEME_KEY = "xlr_theme";
const STORAGE_HISTORY_KEY = "xlr_history";
const STORAGE_LAST_TAB_KEY = "xlr_last_tab";
const STORAGE_FORM_STATE_KEY = "xlr_form_state";
const MAX_HISTORY_ITEMS = 20;

const RESULTS = ["大安", "留连", "速喜", "赤口", "小吉", "空亡"];
const BRANCHES = ["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"];
const FAVORABLE_RESULTS = new Set(["大安", "速喜", "小吉"]);
const CHALLENGING_RESULTS = new Set(["赤口", "空亡"]);
const DELAY_RESULTS = new Set(["留连"]);

const BRANCH_RANGES = {
  "子": "23:00-00:59",
  "丑": "01:00-02:59",
  "寅": "03:00-04:59",
  "卯": "05:00-06:59",
  "辰": "07:00-08:59",
  "巳": "09:00-10:59",
  "午": "11:00-12:59",
  "未": "13:00-14:59",
  "申": "15:00-16:59",
  "酉": "17:00-18:59",
  "戌": "19:00-20:59",
  "亥": "21:00-22:59"
};

const RESULT_MEANINGS = {
  "大安": {
    short: "Stable, calm, steady, broadly favorable.",
    simple: "This usually suggests a stable and calm situation.",
    advice: "Proceed steadily and avoid unnecessary disruption.",
    watchFor: "Do not overcomplicate something that is already stable."
  },
  "留连": {
    short: "Delay, repetition, lingering, entanglement.",
    simple: "This usually suggests waiting, repetition, or a slow outcome.",
    advice: "Be patient and expect delays or rework.",
    watchFor: "Do not mistake slow progress for total failure."
  },
  "速喜": {
    short: "Quick movement, favorable news, speed.",
    simple: "This usually suggests fast developments or encouraging news.",
    advice: "Act promptly while momentum is available.",
    watchFor: "Do not rush so much that you become careless."
  },
  "赤口": {
    short: "Conflict, friction, disagreement, sharpness.",
    simple: "This usually suggests tension, sharp communication, or conflict.",
    advice: "Stay calm, be precise, and reduce the chance of arguments.",
    watchFor: "Avoid impulsive reactions, especially in communication."
  },
  "小吉": {
    short: "Moderately good, manageable benefit, support.",
    simple: "This usually suggests a workable and moderately favorable outcome.",
    advice: "Proceed with measured confidence and practical judgment.",
    watchFor: "Do not expect perfect outcomes from a modestly good sign."
  },
  "空亡": {
    short: "Emptiness, non-arrival, unreliability, uncertainty.",
    simple: "This usually suggests that something may not fully materialize.",
    advice: "Verify facts carefully and do not rely on appearances alone.",
    watchFor: "Do not commit too early when the situation is still unclear."
  }
};

const INLINE_TEXT = {
  formula:
    "The site uses the standard compact formula: (lunar month + lunar day + branch hour − 3) mod 6. The result index maps to the fixed six-result cycle.",
  noDataToday:
    "No rows in the loaded dataset match today's date. This may mean the dataset range does not include today.",
  noDataLookup:
    "No rows matched the selected date and filters.",
  datasetLoadFail:
    "The dataset could not be loaded. Check that xiao_liuren_data.json exists in the docs folder.",
  copySuccess:
    "Copied to clipboard.",
  copyFail:
    "Copy failed in this browser.",
  exportNoRows:
    "There are no filtered rows to export.",
  invalidManual:
    "Please enter a lunar month from 1 to 12 and a lunar day from 1 to 30.",
  invalidGregorian:
    "Please enter a valid Gregorian date and time, or choose current time mode.",
  invalidCompare:
    "Please enter a valid date and time for both Time A and Time B."
};

/* ==========================================================================
   GLOBAL STATE
   ========================================================================== */

const state = {
  dataset: [],
  filteredDataset: [],
  currentTab: "intro",
  todayViewMode: "cards",
  datasetPage: 1,
  datasetRowsPerPage: 50,
  lastCalculatedResult: null
};

/* ==========================================================================
   DOM HELPERS
   ========================================================================== */

function byId(id) {
  return document.getElementById(id);
}

function query(selector) {
  return document.querySelector(selector);
}

function queryAll(selector) {
  return Array.from(document.querySelectorAll(selector));
}

function showElement(element) {
  if (element) {
    element.classList.remove("hidden");
  }
}

function hideElement(element) {
  if (element) {
    element.classList.add("hidden");
  }
}

function setText(id, value) {
  const el = byId(id);
  if (el) {
    el.textContent = value;
  }
}

function clearChildren(element) {
  if (!element) {
    return;
  }
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }
}

function createElement(tag, className = "", text = "") {
  const element = document.createElement(tag);
  if (className) {
    element.className = className;
  }
  if (text) {
    element.textContent = text;
  }
  return element;
}

/* ==========================================================================
   FORMATTING AND TIME HELPERS
   ========================================================================== */

function pad2(numberValue) {
  return String(numberValue).padStart(2, "0");
}

function formatDateISO(dateObject) {
  return [
    dateObject.getFullYear(),
    pad2(dateObject.getMonth() + 1),
    pad2(dateObject.getDate())
  ].join("-");
}

function formatTimeHHMM(dateObject) {
  return `${pad2(dateObject.getHours())}:${pad2(dateObject.getMinutes())}`;
}

function formatDateTimeHuman(dateObject) {
  return `${formatDateISO(dateObject)} ${formatTimeHHMM(dateObject)}`;
}

function extractHourFromDateTime(datetimeString) {
  return Number(String(datetimeString).slice(11, 13));
}

function extractTimeFromDateTime(datetimeString) {
  return String(datetimeString).slice(11, 16);
}

function getBranchIndexFromClockHour(hourValue) {
  return Math.floor((hourValue + 1) / 2) % 12;
}

function getBranchNumberFromClockHour(hourValue) {
  return getBranchIndexFromClockHour(hourValue) + 1;
}

function getBranchNameFromClockHour(hourValue) {
  return BRANCHES[getBranchIndexFromClockHour(hourValue)];
}

function getBranchNameFromBranchNumber(branchNumber) {
  return BRANCHES[(Number(branchNumber) - 1 + 12) % 12];
}

function classifyResult(resultName) {
  if (FAVORABLE_RESULTS.has(resultName)) {
    return "Favorable";
  }
  if (CHALLENGING_RESULTS.has(resultName)) {
    return "Challenging";
  }
  if (DELAY_RESULTS.has(resultName)) {
    return "Delay / Mixed";
  }
  return "Mixed";
}

function resultClass(resultName) {
  switch (resultName) {
    case "大安":
      return "result-da";
    case "留连":
      return "result-liu";
    case "速喜":
      return "result-su";
    case "赤口":
      return "result-chi";
    case "小吉":
      return "result-xiao";
    case "空亡":
      return "result-kong";
    default:
      return "";
  }
}

/* ==========================================================================
   TOASTS
   ========================================================================== */

function showToast(messageText) {
  const toast = byId("appToast");
  if (!toast) {
    return;
  }

  toast.textContent = messageText;
  toast.classList.remove("hidden");
  toast.classList.add("visible");

  window.clearTimeout(showToast._timer);
  showToast._timer = window.setTimeout(() => {
    toast.classList.remove("visible");
    window.setTimeout(() => {
      toast.classList.add("hidden");
    }, 250);
  }, 2200);
}

/* ==========================================================================
   THEME
   ========================================================================== */

function applyTheme(themeName) {
  document.documentElement.dataset.theme = themeName;
  localStorage.setItem(STORAGE_THEME_KEY, themeName);
}

function toggleTheme() {
  const currentTheme = document.documentElement.dataset.theme || "light";
  applyTheme(currentTheme === "dark" ? "light" : "dark");
}

function restoreTheme() {
  const savedTheme = localStorage.getItem(STORAGE_THEME_KEY);
  applyTheme(savedTheme || "light");
}

/* ==========================================================================
   TAB NAVIGATION
   ========================================================================== */

function switchTab(tabId) {
  state.currentTab = tabId;
  localStorage.setItem(STORAGE_LAST_TAB_KEY, tabId);

  queryAll(".tab-section").forEach((section) => {
    section.classList.toggle("hidden", section.id !== tabId);
    section.classList.toggle("active", section.id === tabId);
  });

  queryAll(".nav-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabId);
  });

  if (tabId === "dataset") {
    renderDatasetTable();
  }
}

function restoreLastTab() {
  const savedTab = localStorage.getItem(STORAGE_LAST_TAB_KEY);
  switchTab(savedTab || "intro");
}

/* ==========================================================================
   INLINE HELP AND ACCORDIONS
   ========================================================================== */

function setupInlineHelp() {
  queryAll(".inline-help-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const targetId = button.dataset.help;
      const target = byId(targetId);
      if (!target) {
        return;
      }
      const isHidden = target.classList.contains("hidden");
      target.classList.toggle("hidden", !isHidden);
      button.textContent = isHidden ? "Hide explanation" : "Show explanation";
    });
  });
}

function setupAccordions() {
  queryAll(".accordion-trigger").forEach((trigger) => {
    trigger.addEventListener("click", () => {
      const content = trigger.nextElementSibling;
      const expanded = trigger.getAttribute("aria-expanded") === "true";
      trigger.setAttribute("aria-expanded", expanded ? "false" : "true");
      if (content) {
        content.classList.toggle("hidden", expanded);
      }
    });
  });
}

/* ==========================================================================
   CURRENT STATUS HEADER
   ========================================================================== */

function updateHeaderStatus() {
  const now = new Date();
  const dateIso = formatDateISO(now);
  const timeText = formatTimeHHMM(now);
  const branch = getBranchNameFromClockHour(now.getHours());

  setText("currentDateDisplay", dateIso);
  setText("currentTimeDisplay", timeText);
  setText("currentBranchDisplay", `${branch} (${BRANCH_RANGES[branch]})`);

  getLunarForGregorian(dateIso)
    .then((lunarInfo) => {
      if (!lunarInfo) {
        setText("currentLunarDisplay", "Unavailable");
        return;
      }
      setText(
        "currentLunarDisplay",
        `${lunarInfo.lunarMonth}-${lunarInfo.lunarDay}`
      );
    })
    .catch(() => {
      setText("currentLunarDisplay", "Unavailable");
    });
}

function startClock() {
  updateHeaderStatus();
  window.setInterval(updateHeaderStatus, 1000);
}

/* ==========================================================================
   DATA LOADING
   ========================================================================== */

async function loadDataset() {
  try {
    const response = await fetch(DATA_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const rawData = await response.json();
    state.dataset = normalizeDataset(rawData);
    state.filteredDataset = [...state.dataset];

    setText("datasetTotalCount", String(state.dataset.length));
    setText("statsDatasetSize", String(state.dataset.length));

    initializeDateInputs();
    renderToday();
    renderLookupDefault();
    applyDatasetFiltersAndRender();
    renderStatistics(state.dataset, "Full Dataset");
  } catch (error) {
    console.error("Dataset load error:", error);
    showToast(INLINE_TEXT.datasetLoadFail);
  }
}

function normalizeDataset(rawRows) {
  return rawRows.map((row, index) => {
    const datetimeString = String(row.datetime || "");
    const datePart = datetimeString.slice(0, 10);
    const hourValue =
      typeof row.hour === "number" ? row.hour : extractHourFromDateTime(datetimeString);
    const branchName =
      row.branch || getBranchNameFromClockHour(hourValue);

    const resultName = row.result || "";
    const meaningSource = RESULT_MEANINGS[resultName] || {};

    return {
      id: index + 1,
      datetime: datetimeString,
      date: datePart,
      time: extractTimeFromDateTime(datetimeString),
      hour: hourValue,
      lunar_month: Number(row.lunar_month),
      lunar_day: Number(row.lunar_day),
      branch: branchName,
      result: resultName,
      meaning: row.meaning || meaningSource.short || "",
      advice: row.advice || meaningSource.advice || "",
      simple: meaningSource.simple || "",
      watchFor: meaningSource.watchFor || "",
      classification: classifyResult(resultName)
    };
  });
}

/* ==========================================================================
   LUNAR CONVERSION
   ========================================================================== */

/* Parse Chinese calendar using Intl if supported.
 * Fallback strategy:
 * 1) try Intl.DateTimeFormat with zh-u-ca-chinese and formatToParts
 * 2) if that fails, use dataset lookup by Gregorian date
 */
async function getLunarForGregorian(dateIso) {
  const intlResult = getLunarViaIntl(dateIso);
  if (intlResult) {
    return intlResult;
  }

  const datasetRow = state.dataset.find((row) => row.date === dateIso);
  if (datasetRow) {
    return {
      lunarMonth: datasetRow.lunar_month,
      lunarDay: datasetRow.lunar_day,
      source: "dataset"
    };
  }

  return null;
}

function getLunarViaIntl(dateIso) {
  try {
    const [yearValue, monthValue, dayValue] = dateIso.split("-").map(Number);
    const targetDate = new Date(yearValue, monthValue - 1, dayValue);

    const formatter = new Intl.DateTimeFormat("zh-u-ca-chinese", {
      year: "numeric",
      month: "long",
      day: "numeric"
    });

    const parts = formatter.formatToParts(targetDate);
    const monthPart = parts.find((part) => part.type === "month");
    const dayPart = parts.find((part) => part.type === "day");

    if (!monthPart || !dayPart) {
      return null;
    }

    const lunarMonth = parseChineseMonth(monthPart.value);
    const lunarDay = parseChineseDay(dayPart.value);

    if (!lunarMonth || !lunarDay) {
      return null;
    }

    return {
      lunarMonth,
      lunarDay,
      source: "intl"
    };
  } catch (error) {
    return null;
  }
}

function parseChineseMonth(textValue) {
  const cleaned = String(textValue).replace("月", "").replace("閏", "").replace("闰", "");
  const map = {
    "正": 1,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
    "十一": 11,
    "十二": 12,
    "冬": 11,
    "腊": 12
  };

  if (map[cleaned]) {
    return map[cleaned];
  }

  if (/^\d+$/.test(cleaned)) {
    const numeric = Number(cleaned);
    if (numeric >= 1 && numeric <= 12) {
      return numeric;
    }
  }

  return null;
}

function parseChineseDay(textValue) {
  const cleaned = String(textValue).replace("日", "");
  const map = {
    "初一": 1, "初二": 2, "初三": 3, "初四": 4, "初五": 5,
    "初六": 6, "初七": 7, "初八": 8, "初九": 9, "初十": 10,
    "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15,
    "十六": 16, "十七": 17, "十八": 18, "十九": 19, "二十": 20,
    "廿一": 21, "廿二": 22, "廿三": 23, "廿四": 24, "廿五": 25,
    "廿六": 26, "廿七": 27, "廿八": 28, "廿九": 29, "三十": 30
  };

  if (map[cleaned]) {
    return map[cleaned];
  }

  if (/^\d+$/.test(cleaned)) {
    const numeric = Number(cleaned);
    if (numeric >= 1 && numeric <= 30) {
      return numeric;
    }
  }

  return null;
}

/* ==========================================================================
   CORE XIAO LIU REN CALCULATION
   ========================================================================== */

function calculateXiaoLiuRen(lunarMonth, lunarDay, branchNumber) {
  const index = (Number(lunarMonth) + Number(lunarDay) + Number(branchNumber) - 3) % 6;
  return RESULTS[index];
}

function buildCalculationBreakdown(lunarMonth, lunarDay, branchNumber, resultName) {
  const rawTotal = Number(lunarMonth) + Number(lunarDay) + Number(branchNumber) - 3;
  const moduloValue = rawTotal % 6;
  const resultIndexOneBased = moduloValue + 1;

  return [
    `Lunar month = ${lunarMonth}`,
    `Lunar day = ${lunarDay}`,
    `Branch hour number = ${branchNumber} (${getBranchNameFromBranchNumber(branchNumber)})`,
    `Compute: (${lunarMonth} + ${lunarDay} + ${branchNumber} − 3)`,
    `Raw total = ${rawTotal}`,
    `Modulo 6 = ${moduloValue}`,
    `Index ${resultIndexOneBased} in the six-result cycle gives ${resultName}`
  ];
}

function buildResultSummary(resultName) {
  return RESULT_MEANINGS[resultName] || {
    short: "",
    simple: "",
    advice: "",
    watchFor: ""
  };
}

/* ==========================================================================
   VISUAL HIGHLIGHTING
   ========================================================================== */

function highlightCycle(resultName) {
  queryAll(".cycle-node").forEach((node) => {
    node.classList.toggle("active", node.dataset.result === resultName);
  });
}

function highlightBranchClock(branchName) {
  queryAll(".branch-segment").forEach((segment) => {
    segment.classList.toggle("active", segment.dataset.branch === branchName);
  });
}

/* ==========================================================================
   TODAY VIEW
   ========================================================================== */

function getRowsForDate(dateIso) {
  return state.dataset.filter((row) => row.date === dateIso);
}

function getDominantResult(rows) {
  if (!rows.length) {
    return "—";
  }

  const counts = {};
  rows.forEach((row) => {
    counts[row.result] = (counts[row.result] || 0) + 1;
  });

  return Object.keys(counts).sort((left, right) => counts[right] - counts[left])[0];
}

function renderToday() {
  const todayIso = formatDateISO(new Date());
  const rows = getRowsForDate(todayIso);
  const currentHour = new Date().getHours();
  const currentRow = rows.find((row) => row.hour === currentHour) || null;

  setText("todayDateDisplay", todayIso);
  setText("todayDominantResult", getDominantResult(rows));
  setText("todayCurrentHourResult", currentRow ? currentRow.result : "—");
  setText("todayCurrentHourBranch", currentRow ? currentRow.branch : getBranchNameFromClockHour(currentHour));

  const cardsView = byId("todayResults");
  const timelineView = byId("todayTimeline");
  const emptyState = byId("todayEmptyState");

  clearChildren(cardsView);
  clearChildren(timelineView);

  if (!rows.length) {
    showElement(emptyState);
    emptyState.textContent = INLINE_TEXT.noDataToday;
    return;
  }

  hideElement(emptyState);

  rows.forEach((row) => {
    const card = buildResultCard(row, row.hour === currentHour);
    cardsView.appendChild(card);

    const timelineItem = buildTimelineItem(row, row.hour === currentHour);
    timelineView.appendChild(timelineItem);
  });
}

function buildResultCard(row, isCurrent = false) {
  const article = createElement("article", `result-card fade-in ${isCurrent ? "current-card" : ""}`);
  article.innerHTML = `
    <div class="panel-header">
      <h3 class="${resultClass(row.result)}">${row.result}</h3>
      <span class="result-badge">${row.classification}</span>
    </div>
    <p><strong>Time:</strong> ${row.time}</p>
    <p><strong>Branch:</strong> ${row.branch} (${BRANCH_RANGES[row.branch] || ""})</p>
    <p><strong>Lunar:</strong> ${row.lunar_month}-${row.lunar_day}</p>
    <p><strong>Meaning:</strong> ${row.meaning}</p>
    <p><strong>Advice:</strong> ${row.advice}</p>
    ${isCurrent ? '<p class="current-hour-label">Current hour</p>' : ""}
  `;
  return article;
}

function buildTimelineItem(row, isCurrent = false) {
  const item = createElement("div", `timeline-item ${isCurrent ? "active" : ""}`);
  item.innerHTML = `
    <div class="timeline-time">${row.time}</div>
    <div class="timeline-body">
      <strong class="${resultClass(row.result)}">${row.result}</strong>
      <span>${row.branch}</span>
      <span>${row.meaning}</span>
    </div>
  `;
  return item;
}

function setTodayViewMode(modeName) {
  state.todayViewMode = modeName;
  queryAll("[data-today-view]").forEach((button) => {
    button.classList.toggle("active", button.dataset.todayView === modeName);
  });
  byId("todayCardsView").classList.toggle("hidden", modeName !== "cards");
  byId("todayTimelineView").classList.toggle("hidden", modeName !== "timeline");
}

/* ==========================================================================
   LOOKUP
   ========================================================================== */

function initializeDateInputs() {
  const todayIso = formatDateISO(new Date());

  if (byId("lookupDate")) {
    byId("lookupDate").value = todayIso;
  }
  if (byId("gregorianDateInput")) {
    byId("gregorianDateInput").value = todayIso;
  }
  if (byId("compareDateA")) {
    byId("compareDateA").value = todayIso;
  }
  if (byId("compareDateB")) {
    byId("compareDateB").value = todayIso;
  }

  const now = new Date();
  const timeValue = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
  if (byId("gregorianTimeInput")) {
    byId("gregorianTimeInput").value = timeValue;
  }
  if (byId("compareTimeA")) {
    byId("compareTimeA").value = timeValue;
  }
  if (byId("compareTimeB")) {
    byId("compareTimeB").value = timeValue;
  }
}

function renderLookupDefault() {
  const defaultDate = byId("lookupDate") ? byId("lookupDate").value : "";
  if (defaultDate) {
    renderLookupResults();
  }
}

function renderLookupResults() {
  const dateValue = byId("lookupDate").value;
  const resultFilter = byId("lookupResultFilter").value;
  const branchFilter = byId("lookupBranchFilter").value;
  const viewMode = byId("lookupViewMode").value;

  const cardsContainer = byId("lookupCardsView");
  const timelineContainer = byId("lookupTimelineView");
  const emptyState = byId("lookupEmptyState");

  clearChildren(cardsContainer);
  clearChildren(timelineContainer);

  if (!dateValue) {
    showElement(emptyState);
    emptyState.textContent = "Please select a date.";
    setText("lookupSelectedDateDisplay", "—");
    setText("lookupDominantResult", "—");
    setText("lookupCountDisplay", "0");
    return;
  }

  let rows = getRowsForDate(dateValue);

  if (resultFilter) {
    rows = rows.filter((row) => row.result === resultFilter);
  }
  if (branchFilter) {
    rows = rows.filter((row) => row.branch === branchFilter);
  }

  setText("lookupSelectedDateDisplay", dateValue);
  setText("lookupDominantResult", getDominantResult(rows));
  setText("lookupCountDisplay", String(rows.length));

  if (!rows.length) {
    showElement(emptyState);
    emptyState.textContent = INLINE_TEXT.noDataLookup;
    cardsContainer.classList.add("hidden");
    timelineContainer.classList.add("hidden");
    return;
  }

  hideElement(emptyState);
  cardsContainer.classList.toggle("hidden", viewMode !== "cards");
  timelineContainer.classList.toggle("hidden", viewMode !== "timeline");

  rows.forEach((row) => {
    const card = buildResultCard(row, false);
    cardsContainer.appendChild(card);

    const timelineItem = buildTimelineItem(row, false);
    timelineContainer.appendChild(timelineItem);
  });
}

/* ==========================================================================
   MANUAL CLASSIC CALCULATOR
   ========================================================================== */

function validateManualInputs(monthValue, dayValue) {
  return Number.isInteger(monthValue)
    && Number.isInteger(dayValue)
    && monthValue >= 1
    && monthValue <= 12
    && dayValue >= 1
    && dayValue <= 30;
}

function renderManualResult(lunarMonth, lunarDay, branchNumber) {
  const resultName = calculateXiaoLiuRen(lunarMonth, lunarDay, branchNumber);
  const summary = buildResultSummary(resultName);
  const breakdownItems = buildCalculationBreakdown(
    lunarMonth,
    lunarDay,
    branchNumber,
    resultName
  );

  showElement(byId("manualOutput"));
  setText("manualResultTitle", resultName);
  setText("manualResultBadge", classifyResult(resultName));
  setText("manualMeaning", summary.short);
  setText("manualAdvice", summary.advice);
  setText("manualSimpleExplanation", summary.simple);
  setText("manualWatchFor", summary.watchFor);
  setText(
    "formulaDisplay",
    `(${lunarMonth} + ${lunarDay} + ${branchNumber} − 3) mod 6`
  );

  const list = byId("manualBreakdownList");
  clearChildren(list);
  breakdownItems.forEach((itemText) => {
    const item = createElement("li", "", itemText);
    list.appendChild(item);
  });

  const badge = byId("manualResultBadge");
  badge.className = `result-badge ${resultClass(resultName)}`;

  highlightCycle(resultName);
  highlightBranchClock(getBranchNameFromBranchNumber(branchNumber));
  state.lastCalculatedResult = resultName;

  addHistoryItem({
    type: "Manual Classic",
    datetime: formatDateTimeHuman(new Date()),
    lunarMonth,
    lunarDay,
    branch: getBranchNameFromBranchNumber(branchNumber),
    result: resultName
  });
}

function handleManualCalculation(event) {
  event.preventDefault();

  const monthValue = Number(byId("monthInput").value);
  const dayValue = Number(byId("dayInput").value);
  const branchNumber = Number(byId("branchInput").value);

  if (!validateManualInputs(monthValue, dayValue)) {
    showElement(byId("manualError"));
    byId("manualError").textContent = INLINE_TEXT.invalidManual;
    return;
  }

  hideElement(byId("manualError"));
  renderManualResult(monthValue, dayValue, branchNumber);
}

/* ==========================================================================
   GREGORIAN -> LUNAR CALCULATOR
   ========================================================================== */

async function handleGregorianCalculation(event) {
  event.preventDefault();

  const dateValue = byId("gregorianDateInput").value;
  const timeValue = byId("gregorianTimeInput").value;
  const hourMode = byId("gregorianHourMode").value;

  if (!dateValue || (hourMode === "clock" && !timeValue)) {
    showElement(byId("gregorianError"));
    byId("gregorianError").textContent = INLINE_TEXT.invalidGregorian;
    return;
  }

  hideElement(byId("gregorianError"));

  const finalTime = hourMode === "current"
    ? formatTimeHHMM(new Date())
    : timeValue;

  const hourValue = Number(finalTime.slice(0, 2));
  const branchNumber = getBranchNumberFromClockHour(hourValue);
  const branchName = getBranchNameFromClockHour(hourValue);

  const lunarInfo = await getLunarForGregorian(dateValue);
  if (!lunarInfo) {
    showElement(byId("gregorianError"));
    byId("gregorianError").textContent =
      "Lunar conversion was unavailable for this date.";
    return;
  }

  const resultName = calculateXiaoLiuRen(
    lunarInfo.lunarMonth,
    lunarInfo.lunarDay,
    branchNumber
  );
  const summary = buildResultSummary(resultName);

  showElement(byId("gregorianOutput"));
  setText("gregorianOutputDate", dateValue);
  setText("gregorianOutputTime", finalTime);
  setText("gregorianOutputLunar", `${lunarInfo.lunarMonth}-${lunarInfo.lunarDay}`);
  setText("gregorianOutputBranch", `${branchName} (${BRANCH_RANGES[branchName]})`);
  setText("gregorianResultTitle", resultName);
  setText("gregorianResultBadge", classifyResult(resultName));
  setText("gregorianMeaning", summary.short);
  setText("gregorianAdvice", summary.advice);
  setText(
    "gregorianConversionExplanation",
    `Gregorian date ${dateValue} converts to lunar ${lunarInfo.lunarMonth}-${lunarInfo.lunarDay} using ${lunarInfo.source} conversion.`
  );
  setText(
    "gregorianFormulaDisplay",
    `(${lunarInfo.lunarMonth} + ${lunarInfo.lunarDay} + ${branchNumber} − 3) mod 6 = ${resultName}`
  );

  const badge = byId("gregorianResultBadge");
  badge.className = `result-badge ${resultClass(resultName)}`;

  highlightCycle(resultName);
  highlightBranchClock(branchName);
  state.lastCalculatedResult = resultName;

  addHistoryItem({
    type: "Gregorian Conversion",
    datetime: `${dateValue} ${finalTime}`,
    lunarMonth: lunarInfo.lunarMonth,
    lunarDay: lunarInfo.lunarDay,
    branch: branchName,
    result: resultName
  });
}

/* ==========================================================================
   CURRENT MOMENT
   ========================================================================== */

async function handleCurrentMomentCalculation() {
  const now = new Date();
  const dateIso = formatDateISO(now);
  const timeText = formatTimeHHMM(now);
  const branchName = getBranchNameFromClockHour(now.getHours());
  const branchNumber = getBranchNumberFromClockHour(now.getHours());

  const lunarInfo = await getLunarForGregorian(dateIso);
  if (!lunarInfo) {
    showToast("Lunar conversion was unavailable for the current date.");
    return;
  }

  const resultName = calculateXiaoLiuRen(
    lunarInfo.lunarMonth,
    lunarInfo.lunarDay,
    branchNumber
  );
  const summary = buildResultSummary(resultName);

  showElement(byId("currentMomentOutput"));
  setText("currentMomentDateTime", `${dateIso} ${timeText}`);
  setText("currentMomentLunar", `${lunarInfo.lunarMonth}-${lunarInfo.lunarDay}`);
  setText("currentMomentBranch", `${branchName} (${BRANCH_RANGES[branchName]})`);
  setText("currentMomentResult", resultName);
  setText("currentMomentMeaning", summary.short);
  setText("currentMomentAdvice", summary.advice);

  highlightCycle(resultName);
  highlightBranchClock(branchName);
  state.lastCalculatedResult = resultName;

  addHistoryItem({
    type: "Current Moment",
    datetime: `${dateIso} ${timeText}`,
    lunarMonth: lunarInfo.lunarMonth,
    lunarDay: lunarInfo.lunarDay,
    branch: branchName,
    result: resultName
  });
}

/* ==========================================================================
   COMPARE TIMES
   ========================================================================== */

async function resolveCompareInput(dateId, timeId) {
  const dateValue = byId(dateId).value;
  const timeValue = byId(timeId).value;

  if (!dateValue || !timeValue) {
    return null;
  }

  const lunarInfo = await getLunarForGregorian(dateValue);
  if (!lunarInfo) {
    return null;
  }

  const hourValue = Number(timeValue.slice(0, 2));
  const branchName = getBranchNameFromClockHour(hourValue);
  const branchNumber = getBranchNumberFromClockHour(hourValue);
  const resultName = calculateXiaoLiuRen(
    lunarInfo.lunarMonth,
    lunarInfo.lunarDay,
    branchNumber
  );

  return {
    datetime: `${dateValue} ${timeValue}`,
    lunar: `${lunarInfo.lunarMonth}-${lunarInfo.lunarDay}`,
    branch: `${branchName} (${BRANCH_RANGES[branchName]})`,
    result: resultName
  };
}

async function handleCompareSubmit(event) {
  event.preventDefault();

  const resultA = await resolveCompareInput("compareDateA", "compareTimeA");
  const resultB = await resolveCompareInput("compareDateB", "compareTimeB");

  if (!resultA || !resultB) {
    showElement(byId("compareError"));
    byId("compareError").textContent = INLINE_TEXT.invalidCompare;
    return;
  }

  hideElement(byId("compareError"));
  showElement(byId("compareOutput"));

  setText("compareADateTime", resultA.datetime);
  setText("compareALunar", resultA.lunar);
  setText("compareABranch", resultA.branch);
  setText("compareAResult", resultA.result);

  setText("compareBDateTime", resultB.datetime);
  setText("compareBLunar", resultB.lunar);
  setText("compareBBranch", resultB.branch);
  setText("compareBResult", resultB.result);
}

/* ==========================================================================
   HISTORY
   ========================================================================== */

function loadHistory() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_HISTORY_KEY) || "[]");
  } catch (error) {
    return [];
  }
}

function saveHistory(historyItems) {
  localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(historyItems));
}

function addHistoryItem(item) {
  const currentHistory = loadHistory();
  currentHistory.unshift(item);
  const trimmedHistory = currentHistory.slice(0, MAX_HISTORY_ITEMS);
  saveHistory(trimmedHistory);
  renderHistory();
}

function clearHistory() {
  saveHistory([]);
  renderHistory();
}

function renderHistory() {
  const list = byId("historyList");
  clearChildren(list);

  const historyItems = loadHistory();
  if (!historyItems.length) {
    const emptyItem = createElement("li", "history-empty", "No recent calculations yet.");
    list.appendChild(emptyItem);
    return;
  }

  historyItems.forEach((item) => {
    const li = createElement("li", "history-item");
    li.innerHTML = `
      <strong>${item.type}</strong>
      <span>${item.datetime}</span>
      <span>Lunar: ${item.lunarMonth}-${item.lunarDay}</span>
      <span>Branch: ${item.branch}</span>
      <span class="${resultClass(item.result)}">${item.result}</span>
    `;
    list.appendChild(li);
  });
}

/* ==========================================================================
   DATASET FILTERING, SORTING, PAGINATION
   ========================================================================== */

function getDatasetFilters() {
  return {
    search: byId("searchInput").value.trim().toLowerCase(),
    result: byId("datasetResultFilter").value,
    branch: byId("datasetBranchFilter").value,
    startDate: byId("datasetStartDate").value,
    endDate: byId("datasetEndDate").value,
    sortBy: byId("datasetSortBy").value,
    rowsPerPage: Number(byId("datasetRowsPerPage").value)
  };
}

function applyDatasetFilters(rows) {
  const filters = getDatasetFilters();

  let output = [...rows];

  if (filters.search) {
    output = output.filter((row) => {
      const haystack = [
        row.datetime,
        row.result,
        row.branch,
        row.meaning,
        String(row.lunar_month),
        String(row.lunar_day)
      ].join(" ").toLowerCase();
      return haystack.includes(filters.search);
    });
  }

  if (filters.result) {
    output = output.filter((row) => row.result === filters.result);
  }

  if (filters.branch) {
    output = output.filter((row) => row.branch === filters.branch);
  }

  if (filters.startDate) {
    output = output.filter((row) => row.date >= filters.startDate);
  }

  if (filters.endDate) {
    output = output.filter((row) => row.date <= filters.endDate);
  }

  output = sortDatasetRows(output, filters.sortBy);
  return output;
}

function sortDatasetRows(rows, sortBy) {
  const output = [...rows];

  switch (sortBy) {
    case "datetimeDesc":
      output.sort((a, b) => b.datetime.localeCompare(a.datetime));
      break;
    case "resultAsc":
      output.sort((a, b) => a.result.localeCompare(b.result, "zh"));
      break;
    case "hourAsc":
      output.sort((a, b) => a.hour - b.hour || a.datetime.localeCompare(b.datetime));
      break;
    case "datetimeAsc":
    default:
      output.sort((a, b) => a.datetime.localeCompare(b.datetime));
      break;
  }

  return output;
}

function applyDatasetFiltersAndRender() {
  const filters = getDatasetFilters();
  state.filteredDataset = applyDatasetFilters(state.dataset);
  state.datasetRowsPerPage = filters.rowsPerPage;
  state.datasetPage = 1;

  setText("datasetVisibleCount", String(state.filteredDataset.length));
  setText("datasetTotalCount", String(state.dataset.length));
  setText("datasetDominantResult", getDominantResult(state.filteredDataset));

  renderDatasetTable();
  renderStatistics(
    state.filteredDataset.length ? state.filteredDataset : state.dataset,
    state.filteredDataset.length === state.dataset.length ? "Full Dataset" : "Filtered Dataset"
  );
}

function renderDatasetTable() {
  const body = byId("datasetBody");
  const emptyState = byId("datasetEmptyState");
  clearChildren(body);

  const rows = state.filteredDataset;
  if (!rows.length) {
    showElement(emptyState);
    setText("paginationInfo", "0 / 0");
    return;
  }

  hideElement(emptyState);

  const pageCount = Math.max(1, Math.ceil(rows.length / state.datasetRowsPerPage));
  if (state.datasetPage > pageCount) {
    state.datasetPage = pageCount;
  }

  const startIndex = (state.datasetPage - 1) * state.datasetRowsPerPage;
  const endIndex = startIndex + state.datasetRowsPerPage;
  const visibleRows = rows.slice(startIndex, endIndex);

  visibleRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.datetime}</td>
      <td>${row.hour}</td>
      <td>${row.lunar_month}</td>
      <td>${row.lunar_day}</td>
      <td>${row.branch}</td>
      <td class="${resultClass(row.result)}">${row.result}</td>
      <td>${row.meaning}</td>
      <td><button type="button" class="row-detail-btn" data-row-id="${row.id}">View</button></td>
    `;
    body.appendChild(tr);
  });

  setText("paginationInfo", `Page ${state.datasetPage} / ${pageCount}`);

  queryAll(".row-detail-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const rowId = Number(button.dataset.rowId);
      const row = state.filteredDataset.find((item) => item.id === rowId);
      if (row) {
        openDatasetDetail(row);
      }
    });
  });
}

function openDatasetDetail(row) {
  const panel = byId("datasetRowDetail");
  const content = byId("datasetDetailContent");

  clearChildren(content);
  content.innerHTML = `
    <p><strong>DateTime:</strong> ${row.datetime}</p>
    <p><strong>Lunar:</strong> ${row.lunar_month}-${row.lunar_day}</p>
    <p><strong>Branch:</strong> ${row.branch} (${BRANCH_RANGES[row.branch] || ""})</p>
    <p><strong>Result:</strong> <span class="${resultClass(row.result)}">${row.result}</span></p>
    <p><strong>Meaning:</strong> ${row.meaning}</p>
    <p><strong>Advice:</strong> ${row.advice}</p>
    <p><strong>Classification:</strong> ${row.classification}</p>
  `;

  showElement(panel);
}

function closeDatasetDetail() {
  hideElement(byId("datasetRowDetail"));
}

function nextDatasetPage() {
  const pageCount = Math.max(
    1,
    Math.ceil(state.filteredDataset.length / state.datasetRowsPerPage)
  );
  if (state.datasetPage < pageCount) {
    state.datasetPage += 1;
    renderDatasetTable();
  }
}

function prevDatasetPage() {
  if (state.datasetPage > 1) {
    state.datasetPage -= 1;
    renderDatasetTable();
  }
}

/* ==========================================================================
   EXPORTS
   ========================================================================== */

function downloadBlob(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function exportFilteredJson() {
  if (!state.filteredDataset.length) {
    showToast(INLINE_TEXT.exportNoRows);
    return;
  }
  downloadBlob(
    "xiao_liuren_filtered.json",
    JSON.stringify(state.filteredDataset, null, 2),
    "application/json;charset=utf-8"
  );
}

function exportFilteredCsv() {
  if (!state.filteredDataset.length) {
    showToast(INLINE_TEXT.exportNoRows);
    return;
  }

  const headers = [
    "datetime",
    "hour",
    "lunar_month",
    "lunar_day",
    "branch",
    "result",
    "meaning",
    "advice"
  ];

  const rows = state.filteredDataset.map((row) => headers.map((header) => {
    const value = row[header] ?? "";
    const escaped = String(value).replace(/"/g, '""');
    return `"${escaped}"`;
  }).join(","));

  const csvContent = [headers.join(","), ...rows].join("\n");
  downloadBlob("xiao_liuren_filtered.csv", csvContent, "text/csv;charset=utf-8");
}

/* ==========================================================================
   STATISTICS
   ========================================================================== */

function countResults(rows) {
  const counts = {};
  RESULTS.forEach((resultName) => {
    counts[resultName] = 0;
  });
  rows.forEach((row) => {
    counts[row.result] = (counts[row.result] || 0) + 1;
  });
  return counts;
}

function renderStatistics(rows, scopeLabel) {
  const effectiveRows = rows || [];
  const counts = countResults(effectiveRows);
  const total = effectiveRows.length || 1;
  const favorableTotal = effectiveRows.filter((row) => FAVORABLE_RESULTS.has(row.result)).length;
  const challengingTotal = effectiveRows.filter((row) => CHALLENGING_RESULTS.has(row.result)).length;

  setText("statsDatasetSize", String(effectiveRows.length));
  setText("statsFavorableTotal", String(favorableTotal));
  setText("statsChallengingTotal", String(challengingTotal));
  setText("statsScopeLabel", scopeLabel);

  renderStatsCards(counts, total);
  renderPercentageBars(counts, total);
  renderHourlyDistribution(effectiveRows);
  renderBranchDistribution(effectiveRows);
  renderResultRanking(counts);
}

function renderStatsCards(counts, total) {
  const grid = byId("statsGrid");
  clearChildren(grid);

  RESULTS.forEach((resultName) => {
    const count = counts[resultName] || 0;
    const percentage = total ? ((count / total) * 100).toFixed(1) : "0.0";

    const card = createElement("article", "stat-card fade-in");
    card.innerHTML = `
      <h3 class="${resultClass(resultName)}">${resultName}</h3>
      <p><strong>${count}</strong> rows</p>
      <p>${percentage}%</p>
    `;
    grid.appendChild(card);
  });
}

function renderPercentageBars(counts, total) {
  const container = byId("percentageBars");
  clearChildren(container);

  RESULTS.forEach((resultName) => {
    const count = counts[resultName] || 0;
    const percentage = total ? (count / total) * 100 : 0;

    const row = createElement("div", "percentage-row");
    row.innerHTML = `
      <div class="percentage-label ${resultClass(resultName)}">${resultName}</div>
      <div class="percentage-track">
        <div class="percentage-fill ${resultClass(resultName)}" style="width:${percentage}%"></div>
      </div>
      <div class="percentage-value">${percentage.toFixed(1)}%</div>
    `;
    container.appendChild(row);
  });
}

function renderHourlyDistribution(rows) {
  const container = byId("hourlyDistribution");
  clearChildren(container);

  const counts = {};
  for (let hourValue = 0; hourValue < 24; hourValue += 1) {
    counts[hourValue] = 0;
  }
  rows.forEach((row) => {
    counts[row.hour] = (counts[row.hour] || 0) + 1;
  });

  const maxValue = Math.max(...Object.values(counts), 1);

  Object.keys(counts).forEach((hourKey) => {
    const count = counts[hourKey];
    const percentage = (count / maxValue) * 100;

    const bar = createElement("div", "chart-row");
    bar.innerHTML = `
      <span class="chart-label">${pad2(Number(hourKey))}:00</span>
      <span class="chart-bar-wrap">
        <span class="chart-bar" style="width:${percentage}%"></span>
      </span>
      <span class="chart-value">${count}</span>
    `;
    container.appendChild(bar);
  });
}

function renderBranchDistribution(rows) {
  const container = byId("branchDistribution");
  clearChildren(container);

  const counts = {};
  BRANCHES.forEach((branchName) => {
    counts[branchName] = 0;
  });
  rows.forEach((row) => {
    counts[row.branch] = (counts[row.branch] || 0) + 1;
  });

  const maxValue = Math.max(...Object.values(counts), 1);

  BRANCHES.forEach((branchName) => {
    const count = counts[branchName];
    const percentage = (count / maxValue) * 100;

    const bar = createElement("div", "chart-row");
    bar.innerHTML = `
      <span class="chart-label">${branchName}</span>
      <span class="chart-bar-wrap">
        <span class="chart-bar" style="width:${percentage}%"></span>
      </span>
      <span class="chart-value">${count}</span>
    `;
    container.appendChild(bar);
  });
}

function renderResultRanking(counts) {
  const list = byId("resultRankingList");
  clearChildren(list);

  Object.entries(counts)
    .sort((left, right) => right[1] - left[1])
    .forEach(([resultName, count]) => {
      const item = createElement("li", "ranking-item");
      item.innerHTML = `<span class="${resultClass(resultName)}">${resultName}</span><span>${count}</span>`;
      list.appendChild(item);
    });
}

/* ==========================================================================
   COPYING
   ========================================================================== */

async function copyText(textValue) {
  try {
    await navigator.clipboard.writeText(textValue);
    showToast(INLINE_TEXT.copySuccess);
  } catch (error) {
    console.error(error);
    showToast(INLINE_TEXT.copyFail);
  }
}

function buildTodaySummaryText() {
  const todayIso = formatDateISO(new Date());
  const rows = getRowsForDate(todayIso);
  if (!rows.length) {
    return `Today (${todayIso}): no matching dataset rows.`;
  }

  const dominant = getDominantResult(rows);
  const lines = rows.map((row) => `${row.time} ${row.branch} ${row.result}`);
  return [`Today: ${todayIso}`, `Dominant result: ${dominant}`, ...lines].join("\n");
}

function buildLookupSummaryText() {
  const dateValue = byId("lookupDate").value;
  const dominant = byId("lookupDominantResult").textContent;
  const count = byId("lookupCountDisplay").textContent;
  return `Lookup date: ${dateValue}\nDominant result: ${dominant}\nEntries: ${count}`;
}

function buildManualSummaryText() {
  return [
    `Manual result: ${byId("manualResultTitle").textContent}`,
    `Meaning: ${byId("manualMeaning").textContent}`,
    `Advice: ${byId("manualAdvice").textContent}`,
    `Formula: ${byId("formulaDisplay").textContent}`
  ].join("\n");
}

function buildGregorianSummaryText() {
  return [
    `Gregorian result: ${byId("gregorianResultTitle").textContent}`,
    `Date: ${byId("gregorianOutputDate").textContent}`,
    `Time: ${byId("gregorianOutputTime").textContent}`,
    `Lunar: ${byId("gregorianOutputLunar").textContent}`,
    `Branch: ${byId("gregorianOutputBranch").textContent}`
  ].join("\n");
}

function buildCurrentMomentSummaryText() {
  return [
    `Current Moment Result: ${byId("currentMomentResult").textContent}`,
    `DateTime: ${byId("currentMomentDateTime").textContent}`,
    `Lunar: ${byId("currentMomentLunar").textContent}`,
    `Branch: ${byId("currentMomentBranch").textContent}`
  ].join("\n");
}

/* ==========================================================================
   FORM STATE PERSISTENCE
   ========================================================================== */

function saveFormState() {
  const payload = {
    monthInput: byId("monthInput")?.value || "",
    dayInput: byId("dayInput")?.value || "",
    branchInput: byId("branchInput")?.value || "1",
    gregorianDateInput: byId("gregorianDateInput")?.value || "",
    gregorianTimeInput: byId("gregorianTimeInput")?.value || "",
    lookupDate: byId("lookupDate")?.value || ""
  };
  localStorage.setItem(STORAGE_FORM_STATE_KEY, JSON.stringify(payload));
}

function restoreFormState() {
  try {
    const payload = JSON.parse(localStorage.getItem(STORAGE_FORM_STATE_KEY) || "{}");
    if (payload.monthInput && byId("monthInput")) {
      byId("monthInput").value = payload.monthInput;
    }
    if (payload.dayInput && byId("dayInput")) {
      byId("dayInput").value = payload.dayInput;
    }
    if (payload.branchInput && byId("branchInput")) {
      byId("branchInput").value = payload.branchInput;
    }
    if (payload.gregorianDateInput && byId("gregorianDateInput")) {
      byId("gregorianDateInput").value = payload.gregorianDateInput;
    }
    if (payload.gregorianTimeInput && byId("gregorianTimeInput")) {
      byId("gregorianTimeInput").value = payload.gregorianTimeInput;
    }
    if (payload.lookupDate && byId("lookupDate")) {
      byId("lookupDate").value = payload.lookupDate;
    }
  } catch (error) {
    /* Ignore malformed stored state */
  }
}

function setupFormPersistence() {
  queryAll("input, select").forEach((field) => {
    field.addEventListener("change", saveFormState);
    field.addEventListener("input", saveFormState);
  });
}

/* ==========================================================================
   URL PARAMETER SUPPORT
   ========================================================================== */

function applyUrlState() {
  const params = new URLSearchParams(window.location.search);
  const dateValue = params.get("date");
  const hourValue = params.get("hour");
  const resultValue = params.get("result");
  const tabValue = params.get("tab");

  if (dateValue && byId("lookupDate")) {
    byId("lookupDate").value = dateValue;
  }
  if (resultValue && byId("lookupResultFilter")) {
    byId("lookupResultFilter").value = resultValue;
  }
  if (tabValue) {
    switchTab(tabValue);
  }
  if (hourValue && byId("compareTimeA")) {
    byId("compareTimeA").value = `${pad2(Number(hourValue))}:00`;
  }
}

/* ==========================================================================
   TOOL TAB SWITCHING IN CALCULATOR
   ========================================================================== */

function switchCalculatorPanel(panelId) {
  queryAll(".tool-tab").forEach((tabButton) => {
    tabButton.classList.toggle("active", tabButton.dataset.calcTab === panelId);
  });

  queryAll(".calc-panel").forEach((panel) => {
    panel.classList.toggle("hidden", panel.id !== panelId);
    panel.classList.toggle("active", panel.id === panelId);
  });
}

/* ==========================================================================
   RESET HELPERS
   ========================================================================== */

function resetManualForm() {
  byId("manualCalcForm").reset();
  byId("branchInput").value = "1";
  hideElement(byId("manualError"));
  hideElement(byId("manualOutput"));
}

function resetGregorianForm() {
  byId("gregorianCalcForm").reset();
  initializeDateInputs();
  hideElement(byId("gregorianError"));
  hideElement(byId("gregorianOutput"));
}

function resetLookupForm() {
  byId("lookupForm").reset();
  initializeDateInputs();
  renderLookupResults();
}

function resetCompareForm() {
  byId("compareForm").reset();
  initializeDateInputs();
  hideElement(byId("compareError"));
  hideElement(byId("compareOutput"));
}

/* ==========================================================================
   EVENT WIRING
   ========================================================================== */

function setupTabButtons() {
  queryAll(".nav-tab").forEach((button) => {
    button.addEventListener("click", () => {
      switchTab(button.dataset.tab);
    });
  });

  queryAll(".jump-btn, .footer-jump").forEach((button) => {
    button.addEventListener("click", () => {
      const targetTab = button.dataset.jump;
      switchTab(targetTab);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });
}

function setupButtons() {
  byId("themeToggleBtn")?.addEventListener("click", toggleTheme);
  byId("scrollTopBtn")?.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  byId("askNowBtn")?.addEventListener("click", () => {
    switchTab("calculator");
    switchCalculatorPanel("currentMoment");
    handleCurrentMomentCalculation();
  });

  byId("refreshTodayBtn")?.addEventListener("click", renderToday);
  byId("copyTodaySummaryBtn")?.addEventListener("click", () => {
    copyText(buildTodaySummaryText());
  });

  byId("lookupBtn")?.addEventListener("click", (event) => {
    event.preventDefault();
    renderLookupResults();
  });
  byId("lookupResetBtn")?.addEventListener("click", resetLookupForm);
  byId("lookupCopyBtn")?.addEventListener("click", () => {
    copyText(buildLookupSummaryText());
  });
  byId("lookupViewMode")?.addEventListener("change", renderLookupResults);
  byId("lookupResultFilter")?.addEventListener("change", renderLookupResults);
  byId("lookupBranchFilter")?.addEventListener("change", renderLookupResults);

  byId("manualCalcForm")?.addEventListener("submit", handleManualCalculation);
  byId("manualResetBtn")?.addEventListener("click", resetManualForm);
  byId("manualCopyBtn")?.addEventListener("click", () => {
    if (!byId("manualOutput").classList.contains("hidden")) {
      copyText(buildManualSummaryText());
    }
  });

  byId("gregorianCalcForm")?.addEventListener("submit", handleGregorianCalculation);
  byId("gregorianResetBtn")?.addEventListener("click", resetGregorianForm);
  byId("gregorianCopyBtn")?.addEventListener("click", () => {
    if (!byId("gregorianOutput").classList.contains("hidden")) {
      copyText(buildGregorianSummaryText());
    }
  });

  byId("currentMomentBtn")?.addEventListener("click", handleCurrentMomentCalculation);
  byId("currentMomentCopyBtn")?.addEventListener("click", () => {
    if (!byId("currentMomentOutput").classList.contains("hidden")) {
      copyText(buildCurrentMomentSummaryText());
    }
  });

  byId("compareForm")?.addEventListener("submit", handleCompareSubmit);
  byId("compareResetBtn")?.addEventListener("click", resetCompareForm);

  byId("clearHistoryBtn")?.addEventListener("click", clearHistory);

  byId("datasetFilterForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    applyDatasetFiltersAndRender();
  });
  byId("datasetClearFiltersBtn")?.addEventListener("click", () => {
    byId("datasetFilterForm").reset();
    state.datasetRowsPerPage = 50;
    byId("datasetRowsPerPage").value = "50";
    applyDatasetFiltersAndRender();
  });
  byId("downloadFilteredJsonBtn")?.addEventListener("click", exportFilteredJson);
  byId("downloadFilteredCsvBtn")?.addEventListener("click", exportFilteredCsv);
  byId("prevPageBtn")?.addEventListener("click", prevDatasetPage);
  byId("nextPageBtn")?.addEventListener("click", nextDatasetPage);
  byId("closeDatasetDetailBtn")?.addEventListener("click", closeDatasetDetail);

  ["searchInput", "datasetResultFilter", "datasetBranchFilter", "datasetStartDate", "datasetEndDate", "datasetSortBy", "datasetRowsPerPage"]
    .forEach((fieldId) => {
      byId(fieldId)?.addEventListener("input", applyDatasetFiltersAndRender);
      byId(fieldId)?.addEventListener("change", applyDatasetFiltersAndRender);
    });

  queryAll("[data-today-view]").forEach((button) => {
    button.addEventListener("click", () => {
      setTodayViewMode(button.dataset.todayView);
    });
  });

  queryAll(".tool-tab").forEach((button) => {
    button.addEventListener("click", () => {
      switchCalculatorPanel(button.dataset.calcTab);
    });
  });
}

/* ==========================================================================
   INITIAL VISUAL STATE
   ========================================================================== */

function setupInitialVisuals() {
  highlightCycle("大安");
  highlightBranchClock(getBranchNameFromClockHour(new Date().getHours()));
  setTodayViewMode("cards");
  switchCalculatorPanel("manualClassic");
}

/* ==========================================================================
   APP INITIALIZATION
   ========================================================================== */

async function initializeApp() {
  restoreTheme();
  setupTabButtons();
  setupButtons();
  setupInlineHelp();
  setupAccordions();
  setupFormPersistence();
  setupInitialVisuals();

  applyUrlState();
  restoreFormState();
  restoreLastTab();
  renderHistory();
  startClock();

  await loadDataset();
}

document.addEventListener("DOMContentLoaded", initializeApp);