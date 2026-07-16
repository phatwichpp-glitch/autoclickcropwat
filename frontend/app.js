"use strict";

const MONTH_NAMES = [
  "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
];
const MONTH_DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
const CANDIDATE_PRESETS = ["ทุกวัน", "ทุก 7 วัน", "วันที่ 1, 15", "กำหนดเอง"];
const SHOT_PRESETS = ["ไม่มี", "ทุกวันที่เลือก", "วันที่ 1,15"];

let settings = null;
// calendarState[month(1-12)] = { days: Set<number>, shotDays: Set<number> }
let calendarState = {};

// inline SVG สไตล์ lucide (เส้น outline 2px มุมมน) — ฝังตรงๆ ไม่พึ่ง CDN เพราะ
// โปรแกรมต้องใช้ได้แบบ offline สมบูรณ์
const IC = (paths, cls = "ic-xs") =>
  `<svg class="${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
const ICONS = {
  camera: IC('<path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z"/><circle cx="12" cy="13" r="3"/>'),
  thermometer: IC('<path d="M14 4v10.54a4 4 0 1 1-4 0V4a2 2 0 0 1 4 0Z"/>'),
  cloudRain: IC('<path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/><path d="M16 14v6"/><path d="M8 14v6"/><path d="M12 16v6"/>'),
  sprout: IC('<path d="M7 20h10"/><path d="M10 20c5.5-2.5.8-6.4 3-10"/><path d="M9.5 9.4c1.1.8 1.8 2.2 2.3 3.7-2 .4-3.5.4-4.8-.3-1.2-.6-2.3-1.9-3-4.2 2.8-.5 4.4 0 5.5.8z"/><path d="M14.1 6a7 7 0 0 0-1.1 4c1.9-.1 3.3-.6 4.3-1.4 1-1 1.6-2.3 1.7-4.6-2.7.1-4 1-4.9 2z"/>'),
  layers: IC('<path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/><path d="m22 17.65-9.17 4.16a2 2 0 0 1-1.66 0L2 17.65"/><path d="m22 12.65-9.17 4.16a2 2 0 0 1-1.66 0L2 12.65"/>'),
};

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------
document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.tab;
    if (!target) return; // กันปุ่มอื่นที่เผลอใส่ class นี้มาทำ view หายทั้งหน้า
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("view-dash").classList.toggle("active", target === "dash");
    document.getElementById("view-setup").classList.toggle("active", target === "setup");
    if (target === "setup") renderScanResults(lastScan);
  });
});

// ---------------------------------------------------------------------------
// Config load/save
// ---------------------------------------------------------------------------
async function loadConfig() {
  const res = await fetch("/api/config");
  settings = await res.json();

  document.getElementById("start-year").value = settings.default_start_year;
  document.getElementById("end-year").value = settings.default_end_year;
  document.getElementById("input-dir").value = settings.input_dir || "";
  document.getElementById("output-dir").value = settings.output_dir || "";
  document.getElementById("cropwat-exe-path").value = settings.cropwat_exe_path || "";
  document.getElementById("climate-station-dir").value = settings.climate_station_dir || "";
  document.getElementById("rain-station-dir").value = settings.rain_station_dir || "";
  document.getElementById("crop-file").value = settings.crop_file || "";
  document.getElementById("soil-file").value = settings.soil_file || "";
  document.getElementById("manual-per-candidate").value = settings.manual_minutes_per_candidate;
  document.getElementById("hidden-desktop-mode").checked = settings.hidden_desktop_mode !== false;
  document.getElementById("auto-build-outputs").checked = settings.auto_build_outputs !== false;
  // ปุ่ม "เปิด CropWat" มีประโยชน์เฉพาะโหมดคลาสสิก (ผู้ใช้ต้องเปิด CropWat เอง) —
  // โหมดเดสก์ท็อปซ่อนโปรแกรมเปิดให้เองอยู่แล้ว โชว์ไว้มีแต่ทำให้งง
  document.getElementById("btn-launch-cropwat").hidden = settings.hidden_desktop_mode !== false;
  // ปุ่มถ่ายภาพหน้าจอกลับกัน: มีประโยชน์เฉพาะโหมดเดสก์ท็อปซ่อน
  document.getElementById("btn-take-screenshot").hidden = settings.hidden_desktop_mode === false;
  document.getElementById("speed-preset").value = settings.speed_preset || "normal";
  document.getElementById("shift-year-per-candidate").checked = settings.shift_year_per_candidate !== false;
  document.getElementById("brand-sub").textContent = settings.input_dir
    ? settings.input_dir
    : "ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง";

  calendarState = {};
  for (let m = 1; m <= 12; m++) {
    const src = settings.planting_calendar[String(m)] || settings.planting_calendar[m] || { days: [], shot_days: [] };
    calendarState[m] = { days: new Set(src.days), shotDays: new Set(src.shot_days) };
  }
  renderMonthGrid();
}

function calendarStateToPayload() {
  const out = {};
  for (let m = 1; m <= 12; m++) {
    out[m] = {
      days: Array.from(calendarState[m].days).sort((a, b) => a - b),
      shot_days: Array.from(calendarState[m].shotDays).sort((a, b) => a - b),
    };
  }
  return out;
}

async function saveConfig(partial) {
  const payload = { ...settings, ...partial };
  const res = await fetch("/api/config", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    alert(err.detail || "บันทึกการตั้งค่าไม่สำเร็จ");
    return false;
  }
  settings = await res.json();
  return true;
}

// v0.5.20 — เดิมต้องกด "บันทึกปฏิทิน" แยกก่อนไปกด "เริ่มรันทั้งหมด" เสมอ เสี่ยง
// ลืมกดแล้วรันด้วยค่าเก่าโดยไม่รู้ตัว (feedback จาก UX review) — auto-save แทน
// ทุกครั้งที่แก้ปฏิทิน/ช่วงปี (debounce กันยิง API รัวตอนคลิกติดกันหลายครั้ง)
// พร้อม toast ยืนยันแทน alert() ที่บล็อกการใช้งาน
let toastEl = null;
function showToast(msg) {
  if (!toastEl) {
    toastEl = document.createElement("div");
    toastEl.className = "toast";
    document.body.appendChild(toastEl);
  }
  toastEl.textContent = msg;
  toastEl.classList.add("show");
  clearTimeout(toastEl._hideTimer);
  toastEl._hideTimer = setTimeout(() => toastEl.classList.remove("show"), 1800);
}

async function copyText(text, okMessage = "คัดลอกแล้ว") {
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    showToast(okMessage);
  } catch {
    showToast("คัดลอกไม่สำเร็จ — เบราว์เซอร์ไม่อนุญาต");
  }
}

let calendarSaveTimer = null;
function scheduleCalendarAutoSave() {
  clearTimeout(calendarSaveTimer);
  calendarSaveTimer = setTimeout(async () => {
    const ok = await saveConfig({
      default_start_year: Number(document.getElementById("start-year").value),
      default_end_year: Number(document.getElementById("end-year").value),
      planting_calendar: calendarStateToPayload(),
    });
    if (ok) showToast("บันทึกปฏิทินแล้ว");
  }, 700);
}

document.getElementById("btn-save-setup").addEventListener("click", async () => {
  const ok = await saveConfig({
    input_dir: document.getElementById("input-dir").value,
    output_dir: document.getElementById("output-dir").value,
    cropwat_exe_path: document.getElementById("cropwat-exe-path").value,
    climate_station_dir: document.getElementById("climate-station-dir").value,
    rain_station_dir: document.getElementById("rain-station-dir").value,
    crop_file: document.getElementById("crop-file").value,
    soil_file: document.getElementById("soil-file").value,
    manual_minutes_per_candidate: Number(document.getElementById("manual-per-candidate").value) || 0,
    hidden_desktop_mode: document.getElementById("hidden-desktop-mode").checked,
    auto_build_outputs: document.getElementById("auto-build-outputs").checked,
    speed_preset: document.getElementById("speed-preset").value,
    shift_year_per_candidate: document.getElementById("shift-year-per-candidate").checked,
  });
  updateSummary();
  if (ok) {
    document.getElementById("brand-sub").textContent = settings.input_dir || "ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง";
    alert("บันทึกการตั้งค่าแล้ว");
  }
});

// ---------------------------------------------------------------------------
// Month calendar instrument
// ---------------------------------------------------------------------------
function daysForPreset(preset, len) {
  const days = new Set();
  if (preset === "ทุกวัน") for (let d = 1; d <= len; d++) days.add(d);
  else if (preset === "ทุก 7 วัน") for (let d = 1; d <= len; d += 7) days.add(d);
  else if (preset === "วันที่ 1, 15") { days.add(1); if (len >= 15) days.add(15); }
  return days;
}

function shotDaysForPreset(preset, candidateDays) {
  if (preset === "ทุกวันที่เลือก") return new Set(candidateDays);
  if (preset === "วันที่ 1,15") return new Set([...candidateDays].filter((d) => d === 1 || d === 15));
  return new Set();
}

function detectCandidatePreset(days, len) {
  for (const preset of ["ทุกวัน", "ทุก 7 วัน", "วันที่ 1, 15"]) {
    const ref = daysForPreset(preset, len);
    if (ref.size === days.size && [...ref].every((d) => days.has(d))) return preset;
  }
  return days.size === 0 ? "ไม่รันเดือนนี้" : "กำหนดเอง";
}

function detectShotPreset(shotDays, candidateDays) {
  for (const preset of ["ไม่มี", "ทุกวันที่เลือก", "วันที่ 1,15"]) {
    const ref = shotDaysForPreset(preset, candidateDays);
    if (ref.size === shotDays.size && [...ref].every((d) => shotDays.has(d))) return preset;
  }
  return "กำหนดเอง";
}

const monthGrid = document.getElementById("month-grid");

function renderMonthGrid() {
  monthGrid.innerHTML = "";
  for (let m = 1; m <= 12; m++) {
    const len = MONTH_DAYS[m - 1];
    const s = calendarState[m];
    const candidatePreset = detectCandidatePreset(s.days, len);
    const shotPreset = detectShotPreset(s.shotDays, s.days);
    const skipped = s.days.size === 0;

    const card = document.createElement("div");
    card.className = "month-card" + (skipped ? " skipped" : "");
    card.dataset.month = m;

    card.innerHTML = `
      <div class="month-head">
        <span class="month-name">${MONTH_NAMES[m - 1]}</span>
        <label class="skip-toggle">
          <input type="checkbox" class="skipbox" ${skipped ? "checked" : ""} />
          ข้าม
        </label>
      </div>
      <div class="preset-chips">
        ${CANDIDATE_PRESETS.map(
          (p) => `<span class="chip" data-preset="${p}">${p}</span>`
        ).join("")}
      </div>
      <div class="day-grid"></div>
      <div class="month-count"><b class="cnt">${s.days.size}</b> วัน</div>
      <div class="shot-row">
        <span class="shot-lbl">${ICONS.camera} Capture:</span>
        ${SHOT_PRESETS.map(
          (p) => `<span class="chip shot-chip" data-shot="${p}">${p}</span>`
        ).join("")}
      </div>
    `;
    monthGrid.appendChild(card);

    const dayGrid = card.querySelector(".day-grid");
    for (let d = 1; d <= len; d++) {
      const cell = document.createElement("div");
      cell.className = "day-cell";
      cell.textContent = d;
      cell.dataset.day = d;
      dayGrid.appendChild(cell);
    }

    syncMonthCard(card, m, candidatePreset, shotPreset);
  }
  updateSummary();
}

function syncMonthCard(card, month, candidatePreset, shotPreset) {
  const s = calendarState[month];
  const len = MONTH_DAYS[month - 1];
  card.classList.toggle("skipped", s.days.size === 0);
  card.querySelector(".skipbox").checked = s.days.size === 0;
  card.querySelectorAll(".chip[data-preset]").forEach((c) => {
    c.classList.toggle("selected", c.dataset.preset === candidatePreset);
  });
  card.querySelectorAll(".chip[data-shot]").forEach((c) => {
    c.classList.toggle("selected", c.dataset.shot === shotPreset);
  });
  const dayGrid = card.querySelector(".day-grid");
  dayGrid.classList.toggle("open", candidatePreset === "กำหนดเอง" || s.days.size > 0);
  dayGrid.querySelectorAll(".day-cell").forEach((cell) => {
    const d = Number(cell.dataset.day);
    cell.classList.toggle("on", s.days.has(d));
    cell.classList.toggle("shot", s.shotDays.has(d));
  });
  card.querySelector(".cnt").textContent = s.days.size;
  void len;
}

monthGrid.addEventListener("click", (e) => {
  const card = e.target.closest(".month-card");
  if (!card) return;
  const month = Number(card.dataset.month);
  const len = MONTH_DAYS[month - 1];
  const s = calendarState[month];

  if (e.target.matches(".chip[data-shot]")) {
    s.shotDays = shotDaysForPreset(e.target.dataset.shot, s.days);
  } else if (e.target.matches(".chip[data-preset]")) {
    const preset = e.target.dataset.preset;
    if (preset !== "กำหนดเอง") s.days = daysForPreset(preset, len);
    s.shotDays = new Set([...s.shotDays].filter((d) => s.days.has(d)));
  } else if (e.target.matches(".day-cell")) {
    const d = Number(e.target.dataset.day);
    if (s.days.has(d)) { s.days.delete(d); s.shotDays.delete(d); }
    else s.days.add(d);
  } else {
    return;
  }
  const candidatePreset = detectCandidatePreset(s.days, len);
  const shotPreset = detectShotPreset(s.shotDays, s.days);
  syncMonthCard(card, month, candidatePreset, shotPreset);
  updateSummary();
  scheduleCalendarAutoSave();
});

monthGrid.addEventListener("change", (e) => {
  if (!e.target.matches(".skipbox")) return;
  const card = e.target.closest(".month-card");
  const month = Number(card.dataset.month);
  const len = MONTH_DAYS[month - 1];
  const s = calendarState[month];
  if (e.target.checked) {
    s.days = new Set();
    s.shotDays = new Set();
  } else {
    s.days = daysForPreset("วันที่ 1, 15", len);
  }
  const candidatePreset = detectCandidatePreset(s.days, len);
  const shotPreset = detectShotPreset(s.shotDays, s.days);
  syncMonthCard(card, month, candidatePreset, shotPreset);
  updateSummary();
  scheduleCalendarAutoSave();
});

function updateSummary() {
  let totalDays = 0;
  let totalShots = 0;
  for (let m = 1; m <= 12; m++) {
    totalDays += calendarState[m].days.size;
    totalShots += calendarState[m].shotDays.size;
  }
  document.getElementById("sum-days").textContent = totalDays;
  document.getElementById("sum-shots").textContent = totalShots;

  const startYear = Number(document.getElementById("start-year").value) || 0;
  const endYear = Number(document.getElementById("end-year").value) || 0;
  const years = Math.max(0, endYear - startYear + 1);
  document.getElementById("sum-total").textContent = (totalDays * years).toLocaleString("en-US");

  // เวลาทำมือเทียบเท่า = นาที/วันปลูก (ตัวเลขเดียว รวมทุกขั้นตอน) × วันปลูกทั้งหมด
  // แสดงเป็น "ชั่วโมง" — อ่านจากช่อง input ตรงๆ ให้เลขอัปเดตทันทีที่ผู้ใช้แก้ค่า
  const perCand = Number(document.getElementById("manual-per-candidate").value) || 0;
  const hours = (totalDays * years * perCand) / 60;
  document.getElementById("sum-manual-hours").textContent =
    hours >= 100 ? Math.round(hours).toLocaleString("en-US") : hours.toFixed(1);
}

// เลขชั่วโมงทำมือต้องอัปเดตทันทีที่แก้ค่านาทีในหน้าตั้งค่า (ไม่ต้องรอกดบันทึก)
document.getElementById("manual-per-candidate").addEventListener("input", updateSummary);

document.getElementById("start-year").addEventListener("input", () => { updateSummary(); scheduleCalendarAutoSave(); });
document.getElementById("end-year").addEventListener("input", () => { updateSummary(); scheduleCalendarAutoSave(); });

// ---------------------------------------------------------------------------
// Setup: scan
// ---------------------------------------------------------------------------
let lastScan = null;

document.getElementById("btn-scan").addEventListener("click", async () => {
  const ok = await saveConfig({ input_dir: document.getElementById("input-dir").value });
  if (!ok) return;
  const res = await fetch("/api/scan");
  lastScan = await res.json();
  renderScanResults(lastScan);
});

async function browseFolder(inputEl) {
  const res = await fetch("/api/browse-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ initial_dir: inputEl.value }),
  });
  const data = await res.json();
  if (data.path) {
    inputEl.value = data.path;
  }
}

document.getElementById("btn-browse-input").addEventListener("click", () => {
  browseFolder(document.getElementById("input-dir"));
});

document.getElementById("btn-browse-output").addEventListener("click", () => {
  browseFolder(document.getElementById("output-dir"));
});

async function pickFile(inputEl) {
  const res = await fetch("/api/browse-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ initial_dir: inputEl.value }),
  });
  const data = await res.json();
  if (data.path) {
    inputEl.value = data.path;
  }
}

document.getElementById("btn-browse-cropwat-exe").addEventListener("click", () => {
  pickFile(document.getElementById("cropwat-exe-path"));
});

document.getElementById("btn-launch-cropwat").addEventListener("click", async () => {
  const res = await fetch("/api/launch-cropwat", { method: "POST" });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    alert(err.detail || "เปิด CropWat ไม่สำเร็จ");
    return;
  }
  showToast("กำลังเปิด CropWat...");
});

async function openFolder(path, emptyMessage) {
  if (!path) {
    alert(emptyMessage);
    return;
  }
  const res = await fetch("/api/open-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    alert(err.detail || "เปิดโฟลเดอร์ไม่สำเร็จ");
  }
}

document.getElementById("btn-open-output").addEventListener("click", () => {
  openFolder(document.getElementById("output-dir").value || settings.output_dir, "ยังไม่ได้ตั้งค่าโฟลเดอร์ผลลัพธ์");
});

document.getElementById("btn-open-input").addEventListener("click", () => {
  openFolder(document.getElementById("input-dir").value || settings.input_dir, "ยังไม่ได้ตั้งค่าโฟลเดอร์ข้อมูลต้นทาง");
});

document.getElementById("btn-copy-input").addEventListener("click", () => {
  copyText(document.getElementById("input-dir").value, "คัดลอก path แล้ว");
});

document.getElementById("btn-copy-output").addEventListener("click", () => {
  copyText(document.getElementById("output-dir").value, "คัดลอก path แล้ว");
});

function renderScanResults(scan) {
  const badge = document.getElementById("scan-badge");
  const container = document.getElementById("scan-results");
  const strip = document.getElementById("readiness-strip");

  if (!scan) {
    badge.textContent = "ยังไม่ได้สแกน";
    badge.className = "badge neutral";
    container.innerHTML = "";
    strip.innerHTML = "";
    return;
  }

  if (scan.errors && scan.errors.length) {
    badge.textContent = "มีปัญหา";
    badge.className = "badge warn";
    container.innerHTML = `<div class="note">${scan.errors.join("<br>")}</div>`;
    strip.innerHTML = "";
    return;
  }

  badge.textContent = "สแกนแล้ว";
  badge.className = "badge ok";

  const yearStrip = (s) => {
    if (!s) return "";
    return s.years
      .map((y) => `<i class="${s.missing_years.includes(y) ? "miss" : "ok"}" title="${y}"></i>`)
      .join("");
  };

  container.innerHTML = `
    <div class="detect-row">
      <span class="dr-lbl">${ICONS.thermometer} สถานี Climate/ETo</span>
      <span class="dr-val">${scan.climate_station_folders.join(", ") || "ไม่พบ"}</span>
    </div>
    <div class="year-strip">${yearStrip(scan.climate)}</div>
    <div class="detect-row">
      <span class="dr-lbl">${ICONS.cloudRain} สถานี Rain</span>
      <span class="dr-val">${scan.rain_station_folders.join(", ") || "ไม่พบ"}</span>
    </div>
    <div class="year-strip">${yearStrip(scan.rain)}</div>
    <div class="detect-row">
      <span class="dr-lbl">${ICONS.sprout} ไฟล์ Crop</span>
      <span class="dr-val">${scan.crop_files.length} ไฟล์</span>
    </div>
    <div class="detect-row">
      <span class="dr-lbl">${ICONS.layers} ไฟล์ Soil</span>
      <span class="dr-val">${scan.soil_files.length} ไฟล์</span>
    </div>
  `;

  const climYears = scan.climate ? scan.climate.years.length - scan.climate.missing_years.length : 0;
  const rainYears = scan.rain ? scan.rain.years.length - scan.rain.missing_years.length : 0;
  strip.innerHTML = `
    <div class="stat-tile ok"><div class="stat-lbl">Climate ครบ</div><div class="stat-val">${climYears} ปี</div></div>
    <div class="stat-tile ${scan.rain && scan.rain.missing_years.length ? "warn" : "ok"}">
      <div class="stat-lbl">Rain ครบ</div><div class="stat-val">${rainYears} ปี</div>
    </div>
    <div class="stat-tile"><div class="stat-lbl">Crop</div><div class="stat-val" style="font-size:14px">${scan.crop_files.length ? "พบแล้ว" : "ไม่พบ"}</div></div>
    <div class="stat-tile"><div class="stat-lbl">Soil</div><div class="stat-val" style="font-size:14px">${scan.soil_files.length ? "พบแล้ว" : "ไม่พบ"}</div></div>
  `;
}

// ---------------------------------------------------------------------------
// Dashboard: run status (REST + WebSocket)
// ---------------------------------------------------------------------------
const STATUS_LABEL = { done: "เสร็จ", running: "กำลังรัน", queued: "รอคิว", error: "มีปัญหา" };
let wasRunning = false;

// จัดรูปวินาที -> ข้อความไทยอ่านง่าย ("เหลืออีกประมาณ ...") — ปัดเป็นหน่วยที่ใหญ่
// ที่สุด 2 หน่วย (เช่น "1 ชม. 20 นาที") พอให้กะเวลาได้ ไม่ต้องละเอียดเป็นวินาที
function formatEta(seconds) {
  if (seconds == null || !isFinite(seconds) || seconds < 1) return "";
  const totalMin = Math.round(seconds / 60);
  if (totalMin < 1) return "เหลืออีกไม่ถึง 1 นาที";
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  const parts = [];
  if (h > 0) parts.push(`${h} ชม.`);
  if (m > 0 || h === 0) parts.push(`${m} นาที`);
  return `เหลืออีกประมาณ ${parts.join(" ")}`;
}

function renderStatus(snapshot) {
  if (!snapshot) return;
  const years = snapshot.years || [];
  const doneCount = years.filter((y) => y.status === "done").length;
  const total = years.length;

  const currentTxt = snapshot.current_year ? ` · กำลังรันปี ${snapshot.current_year}` : "";
  const etaTxt = formatEta(snapshot.eta_seconds);
  const etaSuffix = etaTxt ? ` · ${etaTxt}` : "";
  // ใช้ progress ระดับ "วันปลูก" ถ้ามี (ละเอียดกว่าระดับปีมาก — 1 ปีมีหลายวันปลูก
  // การนับแค่ปีทำให้ bar กระโดดทีละก้าวใหญ่ดูเหมือนค้าง) fallback เป็นระดับปี
  if (snapshot.candidate_total > 0) {
    const pct = (snapshot.candidate_done / snapshot.candidate_total) * 100;
    document.getElementById("progress-fill").style.width = `${pct}%`;
    document.getElementById("progress-txt").textContent =
      `${snapshot.candidate_done} / ${snapshot.candidate_total} วันปลูก${currentTxt}${etaSuffix}`;
  } else {
    document.getElementById("progress-fill").style.width = total ? `${(doneCount / total) * 100}%` : "0%";
    document.getElementById("progress-txt").textContent = `${doneCount} / ${total} ปี${currentTxt}`;
  }

  const isRunning = snapshot.overall_state === "running";
  document.getElementById("btn-start").disabled = isRunning;
  document.getElementById("btn-stop").disabled = !isRunning;

  // v0.5.20 — เดิมปุ่ม "รันปีที่มีปัญหาซ้ำ" โชว์ตลอดแม้ไม่มีปีไหน error เลย ทำให้
  // ผู้ใช้กดแล้วงงว่าทำไมขึ้น "ไม่มีปีที่มีปัญหา" (เจอจริงจากผู้ใช้) — ซ่อนปุ่มไปเลย
  // เมื่อ error count = 0 พร้อมโชว์ badge จำนวนปีที่มีปัญหาให้เห็นชัดโดยไม่ต้อง
  // เลื่อนดู year-list ทั้งหมด
  const errorCount = years.filter((y) => y.status === "error").length;
  const retryBtn = document.getElementById("btn-retry");
  retryBtn.hidden = errorCount === 0;
  retryBtn.disabled = isRunning;
  const errorBadge = document.getElementById("error-count-badge");
  errorBadge.hidden = errorCount === 0;
  if (errorCount > 0) errorBadge.textContent = `⚠ ${errorCount} ปีมีปัญหา`;
  document.getElementById("btn-copy-all-errors").hidden = errorCount === 0;

  // พอรันจบ (idle) รีเฟรชความคืบหน้าไฟล์ output อัตโนมัติ ให้เห็นว่าทำต่อได้ถึงไหน
  if (!isRunning && wasRunning) fetchOutputProgress();
  wasRunning = isRunning;

  const list = document.getElementById("year-list");
  list.innerHTML = "";
  for (const y of years) {
    const row = document.createElement("div");
    row.className = "year-row";
    row.innerHTML = `
      <span class="yr tabular">${y.year}</span>
      <span></span>
      <span class="badge ${y.status}">${STATUS_LABEL[y.status] || y.status}</span>
      <span class="frac"></span>
    `;
    list.appendChild(row);
    if (y.status === "error" && y.error_message) {
      const detail = document.createElement("div");
      detail.className = "year-row";
      const wrap = document.createElement("div");
      wrap.className = "err-msg";
      const text = document.createElement("span");
      text.className = "err-text";
      text.textContent = y.error_message;
      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "btn-copy-err";
      copyBtn.title = "คัดลอกข้อความ error นี้";
      copyBtn.textContent = "คัดลอก";
      copyBtn.addEventListener("click", () => copyText(`ปี ${y.year}: ${y.error_message}`));
      wrap.appendChild(text);
      wrap.appendChild(copyBtn);
      detail.appendChild(wrap);
      list.appendChild(detail);
    }
  }
}

document.getElementById("btn-copy-all-errors").addEventListener("click", async () => {
  const res = await fetch("/api/status");
  const snapshot = await res.json();
  const lines = (snapshot.years || [])
    .filter((y) => y.status === "error" && y.error_message)
    .map((y) => `ปี ${y.year}: ${y.error_message}`);
  if (!lines.length) {
    showToast("ไม่มีปีที่มีปัญหา");
    return;
  }
  copyText(lines.join("\n\n"), `คัดลอก error ${lines.length} ปีแล้ว`);
});

async function fetchStatus() {
  const res = await fetch("/api/status");
  renderStatus(await res.json());
}

// ---------------------------------------------------------------------------
// Output scan — ทำถึงไหนแล้ว/เหลืออะไร (resume). สแกนไฟล์ .txt จริงเทียบกับแผน
// ---------------------------------------------------------------------------
async function fetchOutputProgress() {
  let data;
  try {
    const res = await fetch("/api/output-progress");
    data = await res.json();
  } catch {
    return;
  }
  const summary = document.getElementById("os-summary");
  const grid = document.getElementById("os-grid");
  const pct = data.total_expected ? Math.round((data.total_done / data.total_expected) * 100) : 0;
  // "เสร็จ" = ผ่านการตรวจเนื้อไฟล์จริง (parse .txt ครบทุกค่า + ภาพครบคู่) ไม่ใช่แค่มีไฟล์
  let html =
    `ตรวจเนื้อไฟล์แล้ว: สมบูรณ์ <b>${data.total_done.toLocaleString("en-US")}</b> / ${data.total_expected.toLocaleString("en-US")} วันปลูก (${pct}%)` +
    ` · ภาพ ${data.screenshot_count.toLocaleString("en-US")} ไฟล์`;
  if (data.invalid_count > 0) {
    html += `<br><span class="os-warn">⚠ พบไฟล์ไม่สมบูรณ์ ${data.invalid_count} ไฟล์ (เช่น ${data.invalid_files.slice(0, 3).join(", ")}) — จะถูกรันใหม่อัตโนมัติรอบหน้า</span>`;
  }
  summary.innerHTML = html;
  grid.innerHTML = data.years
    .map(
      (y) => `<div class="os-cell ${y.status}"><div class="y">${y.year}</div><div class="f">${y.done}/${y.expected}</div></div>`
    )
    .join("");
}

document.getElementById("btn-rescan").addEventListener("click", fetchOutputProgress);

const MONTH_ABBR_TH = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

// v0.5.24 — บังคับสแกนตรวจไฟล์ climate/rain ก่อนเริ่มรันทุกครั้ง (ตามคำขอผู้ใช้):
// กติกา shift-year ห้ามเดาใช้ไฟล์เดือน "อนาคต" แทนเด็ดขาด (ผิดหลักวิทยาศาสตร์ —
// เอาข้อมูลภูมิอากาศเดือนที่ยังไม่ถึงมาแทนวันปลูกที่ผ่านไปแล้วไม่ได้) ถ้าวันปลูก
// ไหนไม่มีไฟล์ "ก่อนหน้า" ให้ใช้เลย ต้องแจ้งเตือนรายละเอียด + ขอความยินยอมจาก
// ผู้ใช้ก่อนเริ่มรันเสมอ ห้ามปล่อยรันเงียบๆ แล้วค่อยพังทีละวันปลูก
async function checkShiftYearCoverage() {
  try {
    const res = await fetch("/api/shift-year-check");
    return await res.json();
  } catch {
    return { ok: true, problems: [] }; // เช็คไม่ได้ (ออฟไลน์ผิดปกติ) ไม่บล็อกการรัน
  }
}

function describeCoverageProblem(p) {
  if (p.year == null) return p.climate_error || p.rain_error || "เกิดข้อผิดพลาดไม่ทราบสาเหตุ";
  const missing = [];
  if (p.climate_error) missing.push("climate");
  if (p.rain_error) missing.push("rain");
  return `ปี ${p.year} เดือน ${MONTH_ABBR_TH[p.month] || p.month} — ไม่มีไฟล์ ${missing.join(" + ")} ให้ใช้`;
}

document.getElementById("btn-start").addEventListener("click", async () => {
  const check = await checkShiftYearCoverage();
  if (!check.ok && check.problems && check.problems.length) {
    const lines = check.problems.slice(0, 15).map(describeCoverageProblem);
    const more = check.problems.length > 15 ? `\n...และอีก ${check.problems.length - 15} รายการ` : "";
    const proceed = confirm(
      `⚠ พบ ${check.problems.length} จุดที่ไม่มีไฟล์ climate/rain "ก่อนหน้า" ให้ใช้ตามกติกา shift-year ` +
      `(วันปลูกเหล่านี้จะรันไม่สำเร็จ):\n\n${lines.join("\n")}${more}\n\n` +
      `ต้องการเริ่มรันต่อหรือไม่? (วันปลูกที่มีปัญหาจะขึ้นสถานะ "มีปัญหา" ส่วนวันปลูกอื่นรันตามปกติ)`
    );
    if (!proceed) return;
  }

  const start_year = Number(document.getElementById("start-year").value);
  const end_year = Number(document.getElementById("end-year").value);
  const res = await fetch("/api/run/start", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ start_year, end_year }),
  });
  if (!res.ok) {
    const err = await res.json();
    alert(err.detail || "เริ่มรันไม่สำเร็จ");
    return;
  }
  renderStatus(await res.json());
});

document.getElementById("btn-stop").addEventListener("click", async () => {
  const res = await fetch("/api/run/stop", { method: "POST" });
  renderStatus(await res.json());
});

document.getElementById("btn-force-close").addEventListener("click", async () => {
  const proceed = confirm(
    "จะบังคับปิดโปรแกรม CropWat ทันที (เหมือนสั่งปิดผ่าน Task Manager) แล้วรีเซ็ต " +
    "สถานะของโปรแกรมนี้กลับมาพร้อมเริ่มรันใหม่ได้เลย\n\n" +
    "ใช้เมื่อ CropWat ค้างสนิทเท่านั้น (กด X เองก็ไม่ติด) — ไฟล์ .txt ของวันปลูกที่ " +
    "ทำเสร็จไปแล้วจะไม่หายไปไหน กดเริ่มรันใหม่ได้ทันทีหลังเปิด CropWat ขึ้นมาใหม่\n\n" +
    "ดำเนินการต่อหรือไม่?"
  );
  if (!proceed) return;
  const res = await fetch("/api/run/force-close-cropwat", { method: "POST" });
  const data = await res.json().catch(() => ({}));
  showToast(data.killed ? `ปิด CropWat แล้ว (${data.killed} process) — เปิด CropWat ใหม่แล้วกดเริ่มรันได้เลย` : "ไม่พบ CropWat เปิดอยู่ — รีเซ็ตสถานะโปรแกรมนี้แล้ว");
  const status = await fetch("/api/status");
  renderStatus(await status.json());
});

document.getElementById("btn-take-screenshot").addEventListener("click", async (e) => {
  const btn = e.currentTarget;
  btn.disabled = true;
  try {
    const res = await fetch("/api/desktop/screenshot", { method: "POST" });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      alert(data.detail || "ถ่ายภาพหน้าจอไม่สำเร็จ");
      return;
    }
    document.getElementById("screenshot-img").src = data.url;
    document.getElementById("screenshot-modal").hidden = false;
  } finally {
    btn.disabled = false;
  }
});

document.getElementById("btn-minimize").addEventListener("click", async () => {
  showToast("ย่อไปที่ Tray แล้ว — โปรแกรมยังรันต่อเบื้องหลัง เปิดกลับได้จากไอคอนถาดระบบ");
  await fetch("/api/window/minimize", { method: "POST" });
});

document.getElementById("btn-retry").addEventListener("click", async () => {
  const res = await fetch("/api/run/retry", { method: "POST" });
  if (!res.ok) {
    const err = await res.json();
    alert(err.detail || "รันปีที่ค้างใหม่ไม่สำเร็จ");
    return;
  }
  renderStatus(await res.json());
});

document.getElementById("btn-build-excel").addEventListener("click", async (e) => {
  const btn = e.currentTarget;
  btn.disabled = true;
  try {
    const res = await fetch("/api/build-excel", { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      alert(data.detail || "สร้าง Excel ไม่สำเร็จ");
      return;
    }
    alert(`สร้าง Excel สำเร็จ (${data.years_written} ปี)`);
  } finally {
    btn.disabled = false;
  }
});

document.getElementById("btn-build-word").addEventListener("click", async (e) => {
  const btn = e.currentTarget;
  btn.disabled = true;
  try {
    const res = await fetch("/api/build-word", { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      alert(data.detail || "สร้างไฟล์ Word ไม่สำเร็จ");
      return;
    }
    alert(`สร้างไฟล์ Word สำเร็จ (${data.candidates_written} วันปลูก)`);
  } finally {
    btn.disabled = false;
  }
});

// ---------------------------------------------------------------------------
// WebSocket connection indicator
// ---------------------------------------------------------------------------
function setConn(online) {
  document.getElementById("conn-dot").classList.toggle("online", online);
  document.getElementById("conn-label").textContent = online ? "เชื่อมต่อแล้ว" : "ขาดการเชื่อมต่อ";
}

// v0.5.30 — บั๊กที่ผู้ใช้เจอจริง: "โปรแกรมชอบหลุดการเชื่อมต่อเองเวลารันค้าง" —
// ต้นเหตุคือโค้ดเดิม reload หน้าทั้งหน้าทุกครั้งที่ WS "ต่อกลับได้หลังหลุด" ไม่ว่า
// เหตุผลจะเป็นอะไร (สมมติไปเองว่าต่อกลับได้ = backend เพิ่ง restart จากการอัปเดต
// เวอร์ชันเท่านั้น) แต่จริงๆ WS หลุดได้จากหลายสาเหตุที่ไม่เกี่ยวกับ restart เลย
// เช่น เครื่องมีงานหนัก (ตอนรันค้าง automation thread แย่ง CPU/GIL) ทำให้ WS
// ค้างชั่วคราวเกิน timeout ของเบราว์เซอร์ — พอต่อกลับได้ก็ reload ทันที ถ้ายัง
// ค้างอยู่ก็หลุดซ้ำแล้ว reload วนอีก ดูเหมือน "หลุดการเชื่อมต่อเอง" ไม่มีที่สิ้นสุด
// แก้โดย reload เฉพาะตอนที่รู้แน่ชัดว่ากำลังอัปเดตเวอร์ชันอยู่จริง (ผ่าน
// updateInProgress ที่ applyUpdate() ตั้งไว้) เหตุผลอื่นๆ แค่ต่อกลับเงียบๆ พอ
let updateInProgress = false;

function connectWebSocket() {
  const proto = location.protocol === "https:" ? "wss" : "ws";
  const ws = new WebSocket(`${proto}://${location.host}/ws`);
  ws.onopen = () => {
    if (updateInProgress) { location.reload(); return; }
    setConn(true);
  };
  ws.onclose = () => { setConn(false); setTimeout(connectWebSocket, 2000); };
  ws.onerror = () => ws.close();
  ws.onmessage = (event) => renderStatus(JSON.parse(event.data));
}

// ---------------------------------------------------------------------------
// เช็คอัปเดตตอนเปิดโปรแกรม — มีเวอร์ชันใหม่ = โชว์ปุ่มอัปเดตที่มุมบน กดแล้ว
// backend ดาวน์โหลด+สลับไฟล์+restart ตัวเองให้ทั้งหมด (ดู backend/updater.py)
// ---------------------------------------------------------------------------
const updateModal = document.getElementById("update-modal");

async function checkUpdate(manual = false) {
  const btn = document.getElementById("btn-check-update");
  if (manual && btn) btn.disabled = true;
  try {
    const res = await fetch("/api/update/check");
    const info = await res.json();
    document.getElementById("app-version").textContent = `v${info.current}`;
    if (info.update_available) {
      const updBtn = document.getElementById("btn-update");
      updBtn.hidden = false;
      updBtn.lastChild.textContent = `อัปเดตเป็น v${info.latest}`;
      // เด้งแจ้งเตือนเต็มจอ "ทุกครั้ง" ที่เปิดโปรแกรมแล้วมีเวอร์ชันใหม่ (จงใจไม่มี
      // ปุ่ม "ไม่ต้องเตือนอีก" — กันคนใช้เวอร์ชันเก่าค้างไว้ทั้งที่ตัวแก้บั๊กออกแล้ว)
      document.getElementById("upd-ver").textContent = `v${info.latest}`;
      document.getElementById("upd-notes").textContent = (info.notes || "").trim();
      updateModal.hidden = false;
    } else if (manual) {
      showToast(`ใช้เวอร์ชันล่าสุดอยู่แล้ว (v${info.current})`);
    }
  } catch {
    if (manual) showToast("ตรวจอัปเดตไม่สำเร็จ — ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต");
  } finally {
    if (manual && btn) btn.disabled = false;
  }
}

document.getElementById("btn-check-update").addEventListener("click", () => checkUpdate(true));

const updatingModal = document.getElementById("updating-modal");

async function applyUpdate() {
  updateModal.hidden = true;
  updateInProgress = true;
  const btn = document.getElementById("btn-update");
  btn.disabled = true;

  // แสดงหน้าคำแนะนำทันทีที่กด — ขั้นดาวน์โหลด (~23MB) ใช้เวลาหลายวินาทีถึงหลาย
  // นาทีตามเน็ต ถ้าไม่มีอะไรบอกสถานะ ผู้ใช้จะคิดว่ากดแล้วไม่เกิดอะไรขึ้น
  updatingModal.hidden = false;
  document.getElementById("updating-status").textContent = "กำลังดาวน์โหลดเวอร์ชันใหม่...";
  document.getElementById("updating-hint").innerHTML =
    "อย่าปิดโปรแกรมระหว่างนี้ — เมื่อดาวน์โหลดเสร็จ โปรแกรมจะปิดและเปิดใหม่เอง หน้านี้จะกลับมาโดยอัตโนมัติ";

  try {
    const res = await fetch("/api/update/apply", { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      updateInProgress = false;
      updatingModal.hidden = true;
      alert(
        (data.detail || "อัปเดตไม่สำเร็จ") +
        "\n\nถ้ายังไม่สำเร็จ: ปิดโปรแกรม (ปุ่ม ✕ บนแถบลอย) แล้วเปิดใหม่ จากนั้นลองอัปเดตอีกครั้ง"
      );
      btn.disabled = false;
      return;
    }
    document.getElementById("updating-status").textContent = "ดาวน์โหลดเสร็จแล้ว — กำลังรีสตาร์ทโปรแกรม...";
    document.getElementById("conn-label").textContent = "กำลังอัปเดต...";
    // จากนี้ backend จะปิดตัวเอง → WS หลุด → พอตัวใหม่เปิด หน้าจะ reload เอง
    // ตาข่ายสุดท้าย: ถ้าเกิน 45 วินาทีแล้วหน้ายังไม่กลับมา แสดงวิธีกู้เอง
    setTimeout(() => {
      document.getElementById("updating-hint").innerHTML =
        "⚠ นานผิดปกติ — ถ้าหน้านี้ยังไม่กลับมาเอง: <b>ดับเบิลคลิกไอคอน CropWatAutoRunner เปิดโปรแกรมใหม่เอง</b> " +
        "(ไฟล์ถูกอัปเดตแล้ว แค่ตัวเปิดอัตโนมัติอาจไม่ทำงาน) แล้วเช็คเลขเวอร์ชันที่มุมขวาบน";
    }, 45000);
  } catch {
    updateInProgress = false;
    updatingModal.hidden = true;
    alert("อัปเดตไม่สำเร็จ (เชื่อมต่อ backend ไม่ได้)\n\nปิดโปรแกรมแล้วเปิดใหม่ จากนั้นลองอัปเดตอีกครั้ง");
    btn.disabled = false;
  }
}

document.getElementById("btn-update").addEventListener("click", () => {
  if (confirm("โปรแกรมจะปิดและเปิดขึ้นมาใหม่เป็นเวอร์ชันล่าสุดโดยอัตโนมัติ อัปเดตเลยหรือไม่?")) applyUpdate();
});
document.getElementById("upd-now").addEventListener("click", applyUpdate);
document.getElementById("upd-later").addEventListener("click", () => { updateModal.hidden = true; });
document.getElementById("upd-close-x").addEventListener("click", () => { updateModal.hidden = true; });

// ---------------------------------------------------------------------------
// Quick-start guide modal — โชว์อัตโนมัติตอนเปิดโปรแกรม จนกว่าผู้ใช้จะติ๊ก
// "ไม่ต้องแสดงอีก" (จำใน localStorage) — เรียกดูซ้ำได้ตลอดจากปุ่ม "วิธีใช้"
// ---------------------------------------------------------------------------
const guideModal = document.getElementById("guide-modal");

function showGuide() {
  guideModal.hidden = false;
  document.getElementById("guide-dontshow").checked =
    localStorage.getItem("cw-hide-guide") === "1";
}

function closeGuide() {
  localStorage.setItem(
    "cw-hide-guide",
    document.getElementById("guide-dontshow").checked ? "1" : "0"
  );
  guideModal.hidden = true;
}

document.getElementById("btn-help").addEventListener("click", showGuide);
document.getElementById("guide-close").addEventListener("click", closeGuide);
document.getElementById("guide-close-x").addEventListener("click", closeGuide);
guideModal.addEventListener("click", (e) => {
  if (e.target === guideModal) closeGuide();
});

if (localStorage.getItem("cw-hide-guide") !== "1") showGuide();

// ---------------------------------------------------------------------------
// Peek screenshot modal (v0.8.0) — เช็คสถานะเฉยๆ แยกจาก screenshot ที่ต้องส่งงาน
// ---------------------------------------------------------------------------
const screenshotModal = document.getElementById("screenshot-modal");

function closeScreenshotModal() {
  screenshotModal.hidden = true;
  document.getElementById("screenshot-img").src = "";
}

document.getElementById("screenshot-close").addEventListener("click", closeScreenshotModal);
document.getElementById("screenshot-close-x").addEventListener("click", closeScreenshotModal);
screenshotModal.addEventListener("click", (e) => {
  if (e.target === screenshotModal) closeScreenshotModal();
});

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
loadConfig();
fetchStatus();
fetchOutputProgress();
connectWebSocket();
checkUpdate();
