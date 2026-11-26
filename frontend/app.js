"use strict";

const MONTH_NAMES = [
  "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
];
const MONTH_DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
const CANDIDATE_PRESETS = ["ทุกวัน", "ทุก 7 วัน", "วันที่ 1, 15", "กำหนดเอง"];
const SHOT_PRESETS = ["ไม่มี", "ทุกวันที่เลือก", "วันที่ 1,15"];
// ข้อความบนชิป (สั้นลงให้อยู่บรรทัดเดียว) — ค่า preset ข้างบนยังเป็น key ภายในเหมือนเดิม
const CHIP_LABEL = { "วันที่ 1, 15": "1, 15", "ทุกวันที่เลือก": "ทั้งหมด", "วันที่ 1,15": "1, 15" };
const chipLabel = (p) => CHIP_LABEL[p] || p;

let settings = null;
// calendarState[month(1-12)] = { days: Set<number>, shotDays: Set<number> }
let calendarState = {};

// v1.0 — สถานะสำหรับแผงหลัก/แบนเนอร์นำทาง ใช้แสดงผลอย่างเดียว: อ่านจากผลของการเรียก
// API เดิม (ช่วงเวลา/จำนวนครั้งที่เรียกเหมือน v0.12 ทุกประการ — ปุ่ม "เปิด Excel ดูผล"
// ในแผงหลักแค่เรียกฟังก์ชัน buildAndOpen เดิมจากอีกปุ่มหนึ่ง)
let lastSnapshot = null;  // สถานะรันล่าสุดจาก /api/status หรือ WebSocket
let lastOutput = null;    // ผลตรวจไฟล์ในเครื่องล่าสุดจาก /api/output-progress
// แผน/ช่วงปี/โฟลเดอร์ผลลัพธ์เปลี่ยนหลังตรวจไฟล์ครั้งล่าสุด → ตัวเลขใน lastOutput เก่าแล้ว
let outputStale = false;
let exeState = "pending"; // "pending" | "found" | "missing" | "unknown" จาก /api/cropwat-exe
let exePath = "";
const fmtNum = (n) => Number(n || 0).toLocaleString("en-US");
const planCard = document.getElementById("plan-card");

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
function showTab(target) {
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === target));
  document.getElementById("view-dash").classList.toggle("active", target === "dash");
  document.getElementById("view-setup").classList.toggle("active", target === "setup");
  // จำแท็บไว้ใน URL — กด refresh แล้วยังอยู่แท็บเดิม
  history.replaceState(null, "", target === "setup" ? "#setup" : location.pathname);
  if (target === "setup") renderScanResults(lastScan);
}

document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (!btn.dataset.tab) return; // กันปุ่มอื่นที่เผลอใส่ class นี้มาทำ view หายทั้งหน้า
    showTab(btn.dataset.tab);
  });
});

// ---------------------------------------------------------------------------
// v1.1 — ธีม มืด / สว่าง / ตามระบบ (ค่าเริ่มต้นมืด) จำไว้ใน localStorage
// ค่าเริ่มต้นถูกตั้งตั้งแต่ <head> แล้ว (กันหน้ากะพริบ) — ส่วนนี้แค่สลับและตามธีมระบบ
// ---------------------------------------------------------------------------
const THEME_ORDER = ["dark", "light", "auto"];
const THEME_LABEL = { dark: "ธีมมืด", light: "ธีมสว่าง", auto: "ธีมตามระบบ (Windows)" };
const darkMq = window.matchMedia("(prefers-color-scheme: dark)");

function applyTheme(pref) {
  const root = document.documentElement;
  root.dataset.themePref = pref;
  root.dataset.theme = pref === "dark" || (pref === "auto" && darkMq.matches) ? "dark" : "light";
  document.getElementById("btn-theme").title = `${THEME_LABEL[pref]} — กดเพื่อสลับ`;
  // v1.2.1 — แถบชื่อหน้าต่าง (ของ Windows) ให้มืด/สว่างตามธีมด้วย
  fetch("/api/window/theme", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dark: root.dataset.theme === "dark" }),
  }).catch(() => { /* แถบชื่อเป็นของเสริม */ });
}

document.getElementById("btn-theme").addEventListener("click", () => {
  const cur = document.documentElement.dataset.themePref || "dark";
  const next = THEME_ORDER[(THEME_ORDER.indexOf(cur) + 1) % THEME_ORDER.length];
  try { localStorage.setItem("cw-theme", next); } catch { /* ใช้ได้แค่รอบนี้ */ }
  applyTheme(next);
  showToast(THEME_LABEL[next]);
});
darkMq.addEventListener("change", () => {
  if (document.documentElement.dataset.themePref === "auto") applyTheme("auto");
});
applyTheme(THEME_ORDER.includes(document.documentElement.dataset.themePref) ? document.documentElement.dataset.themePref : "dark");

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
  applyModeVisibility();
  document.getElementById("speed-preset").value = settings.speed_preset || "normal";
  document.getElementById("parallel-windows").value = String(settings.parallel_windows ?? 0);
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
  updateReadiness();
  updateSaveBar();
}

// ปุ่ม "เปิด CropWat" มีประโยชน์เฉพาะโหมดคลาสสิก (ผู้ใช้ต้องเปิด CropWat เอง) —
// โหมดเดสก์ท็อปซ่อนโปรแกรมเปิดให้เองอยู่แล้ว โชว์ไว้มีแต่ทำให้งง ส่วนปุ่ม
// Snapshot/Live View กลับกัน: มีประโยชน์เฉพาะโหมดเดสก์ท็อปเบื้องหลัง
// (v0.11.3 แยกเป็นฟังก์ชัน — เรียกซ้ำหลังบันทึกตั้งค่า ปุ่มเปลี่ยนทันทีไม่ต้องเปิดใหม่)
function applyModeVisibility() {
  const hiddenMode = settings.hidden_desktop_mode !== false;
  document.getElementById("btn-launch-cropwat").hidden = hiddenMode;
  document.getElementById("btn-take-screenshot").hidden = !hiddenMode;
  document.getElementById("btn-live-toggle").hidden = !hiddenMode;
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
function showToast(msg, isError = false) {
  if (!toastEl) {
    toastEl = document.createElement("div");
    toastEl.className = "toast";
    document.body.appendChild(toastEl);
  }
  toastEl.textContent = msg;
  toastEl.classList.toggle("error", isError);
  toastEl.classList.add("show");
  clearTimeout(toastEl._hideTimer);
  toastEl._hideTimer = setTimeout(() => toastEl.classList.remove("show"), isError ? 4000 : 2200);
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
  // แผน/ช่วงปีเปลี่ยน → ตัวเลข "เสร็จแล้ว X จาก Y" จากการตรวจไฟล์ครั้งก่อนใช้ไม่ได้แล้ว
  // (ไม่ยิงตรวจไฟล์ใหม่เอง — การเรียก API คงเดิมทุกประการ ผู้ใช้กด "ตรวจไฟล์ใหม่" ได้)
  outputStale = true;
  updateRunHero();
}

document.getElementById("btn-save-setup").addEventListener("click", async () => {
  const prevOutputDir = settings && settings.output_dir;
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
    parallel_windows: Number(document.getElementById("parallel-windows").value) || 0,
    shift_year_per_candidate: document.getElementById("shift-year-per-candidate").checked,
  });
  updateSummary();
  if (ok) {
    document.getElementById("brand-sub").textContent = settings.input_dir || "ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง";
    applyModeVisibility();
    refreshCropwatExeStatus(); // path CropWat อาจเพิ่งเปลี่ยน — อัปเดตสถานะทันที
    if (settings.output_dir !== prevOutputDir) outputStale = true; // ผลตรวจไฟล์เดิมเป็นของโฟลเดอร์เก่า
    updateReadiness();
    updateSaveBar();
    showToast("บันทึกการตั้งค่าแล้ว");
  }
});

// ---------------------------------------------------------------------------
// สถานะการค้นหาโปรแกรม CropWat 8.0 (v0.10.2) — ตอบคำถาม "ต้องตั้ง path เองไหม"
// ตรงๆ ในหน้าตั้งค่า: ติดตั้งที่ตำแหน่งมาตรฐาน ระบบหาเจอเอง ไม่ต้องตั้งอะไรเลย
// ---------------------------------------------------------------------------
let exeSource = null;
let exeDetailsAutoOpened = false;

async function refreshCropwatExeStatus() {
  try {
    const res = await fetch("/api/cropwat-exe");
    const data = await res.json();
    exeState = data.found ? "found" : "missing";
    exePath = data.path || "";
    exeSource = data.source;
  } catch {
    exeState = "unknown";
  }
  updateReadiness(); // วาดกล่องสถานะ (renderExeStatus) + เช็คลิสต์ + แผงหลัก
}

// กล่องสถานะในขั้นที่ 3 — แยกจากการ fetch เพราะต้องรู้โหมด (settings) ด้วย ซึ่งอาจ
// โหลดเสร็จทีหลัง: โหมดคลาสสิกไม่จำเป็นต้องรู้ตำแหน่งโปรแกรม (ผู้ใช้เปิด CropWat เอง)
function renderExeStatus() {
  const el = document.getElementById("cropwat-exe-status");
  const classic = !!settings && settings.hidden_desktop_mode === false;
  if (exeState === "found") {
    el.className = "exe-status ok";
    el.textContent = exeSource === "auto"
      ? `✓ พบโปรแกรมอัตโนมัติ: ${exePath} — ไม่ต้องตั้งค่าเพิ่ม`
      : `✓ ใช้ไฟล์ที่ระบุ: ${exePath}`;
  } else if (exeState === "missing" && classic) {
    el.className = "exe-status";
    el.textContent = "ไม่พบ CropWat 8.0 — โหมดคลาสสิกไม่จำเป็นต้องใช้ (ใช้เฉพาะปุ่ม \"เปิด CropWat\" ในหน้าภาพรวม)";
  } else if (exeState === "missing") {
    el.className = "exe-status missing";
    el.textContent = "✗ ไม่พบ CropWat 8.0 ในเครื่อง — ถ้าติดตั้งไว้แล้ว ให้ระบุตำแหน่ง CROPWAT.exe ด้านล่าง";
    // ช่องระบุตำแหน่งเองซ่อนไว้ไม่ให้รกตา — กางให้อัตโนมัติเฉพาะตอนหาไม่เจอ (ครั้งเดียว)
    if (!exeDetailsAutoOpened) {
      document.getElementById("exe-details").open = true;
      exeDetailsAutoOpened = true;
    }
  } else if (exeState === "unknown") {
    el.className = "exe-status";
    el.textContent = "ตรวจหาโปรแกรมไม่ได้ในตอนนี้ — ถ้าติดตั้งไว้นอกตำแหน่งมาตรฐาน ให้ระบุตำแหน่งเองด้านล่าง";
  } else {
    el.className = "exe-status";
    el.textContent = "กำลังตรวจสอบ...";
  }
}

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
// v0.11.3 — ตารางรายวันซ่อนไว้เป็นค่าเริ่มต้น (หน้าสั้นลงมาก) เปิดเฉพาะเดือนที่
// วันปลูกไม่ตรงแบบด่วนใดๆ หรือผู้ใช้กดชิป "กำหนดเอง" (กดซ้ำ = พับเก็บ)
const openMonths = new Set();

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

    // v0.13 — แถวเดือนแนวนอนกะทัดรัด: .month-inner เป็นกริดของแถว (ชื่อเดือน | ชิปวันปลูก
    // | ชิปถ่ายภาพ | ข้าม) ตารางรายวันกางเต็มความกว้างแถวเฉพาะเดือนที่กด "กำหนดเอง"
    card.innerHTML = `
      <div class="month-inner">
        <div class="month-head">
          <span class="month-name">${MONTH_NAMES[m - 1]}<span class="cnt-wrap"><span class="cnt">${s.days.size}</span> วัน</span></span>
          <label class="skip-toggle" title="ไม่ประมวลผลเดือนนี้">
            <input type="checkbox" class="skipbox" ${skipped ? "checked" : ""} />
            ข้าม
          </label>
        </div>
        <div class="preset-chips">
          ${CANDIDATE_PRESETS.map(
            (p) => `<span class="chip" data-preset="${p}">${chipLabel(p)}</span>`
          ).join("")}
        </div>
        <div class="shot-row" title="วันที่จะถ่ายภาพบันทึกไว้ส่งงาน (จุดสีทองบนตารางรายวัน)">
          <span class="shot-lbl">${ICONS.camera} ถ่ายภาพ</span>
          ${SHOT_PRESETS.map(
            (p) => `<span class="chip shot-chip" data-shot="${p}">${chipLabel(p)}</span>`
          ).join("")}
        </div>
        <div class="day-grid"></div>
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
  const gridOpen = s.days.size > 0 && (candidatePreset === "กำหนดเอง" || openMonths.has(month));
  card.querySelectorAll(".chip[data-preset]").forEach((c) => {
    const p = c.dataset.preset;
    c.classList.toggle("selected", p === candidatePreset || (p === "กำหนดเอง" && gridOpen));
  });
  card.querySelectorAll(".chip[data-shot]").forEach((c) => {
    c.classList.toggle("selected", c.dataset.shot === shotPreset);
  });
  const dayGrid = card.querySelector(".day-grid");
  dayGrid.classList.toggle("open", gridOpen);
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
    if (preset === "กำหนดเอง") {
      // แค่เปิด/ปิดตารางรายวัน ไม่แตะวันที่เลือกไว้ (เดือนที่ข้ามอยู่ → เริ่มจาก 1, 15)
      if (openMonths.has(month)) openMonths.delete(month);
      else openMonths.add(month);
      if (s.days.size === 0) s.days = daysForPreset("วันที่ 1, 15", len);
    } else {
      openMonths.delete(month);
      s.days = daysForPreset(preset, len);
    }
    s.shotDays = new Set([...s.shotDays].filter((d) => s.days.has(d)));
  } else if (e.target.matches(".day-cell")) {
    const d = Number(e.target.dataset.day);
    if (s.days.has(d)) { s.days.delete(d); s.shotDays.delete(d); }
    else s.days.add(d);
    openMonths.add(month); // กันตารางพับหายเองกลางคันตอนคลิกจนตรงแบบด่วนพอดี
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
    openMonths.delete(month);
  } else {
    s.days = daysForPreset("วันที่ 1, 15", len);
  }
  const candidatePreset = detectCandidatePreset(s.days, len);
  const shotPreset = detectShotPreset(s.shotDays, s.days);
  syncMonthCard(card, month, candidatePreset, shotPreset);
  updateSummary();
  scheduleCalendarAutoSave();
});

// ---------------------------------------------------------------------------
// v0.13 — แถบปุ่มลัด "ตั้งพร้อมกันทุกเดือน" เหนือแถวเดือน
// ใช้กับเฉพาะเดือนที่ไม่ได้ติ๊ก "ข้าม": ผู้ใช้มักข้ามบางเดือนโดยตั้งใจ ถ้าตั้งทับทั้ง 12
// เดือน เดือนที่ข้ามไว้จะกลับมารันเงียบๆ (เพิ่มงานหลายร้อยรอบโดยไม่รู้ตัว) —
// ยกเว้นตอนข้ามอยู่ทุกเดือน ให้ใช้กับทั้ง 12 เดือน (ทางเดียวที่จะเริ่มแผนใหม่ได้เร็ว)
// ---------------------------------------------------------------------------
const ALL_MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
const activeMonthList = () => ALL_MONTHS.filter((m) => calendarState[m].days.size > 0);

document.getElementById("bulk-bar").addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-bulk-days], button[data-bulk-shot]");
  if (!btn || !settings) return;
  const active = activeMonthList();

  if (btn.dataset.bulkDays) {
    const preset = btn.dataset.bulkDays;
    if (preset === "skip") {
      if (!active.length) return;
      if (!confirm("ข้ามทุกเดือน = ล้างวันปลูกที่เลือกไว้ทั้ง 12 เดือน\n\nต้องการดำเนินการต่อหรือไม่?")) return;
      for (const m of ALL_MONTHS) {
        calendarState[m].days = new Set();
        calendarState[m].shotDays = new Set();
      }
      openMonths.clear();
    } else {
      const targets = active.length ? active : ALL_MONTHS;
      const custom = targets.filter(
        (m) => detectCandidatePreset(calendarState[m].days, MONTH_DAYS[m - 1]) === "กำหนดเอง"
      );
      if (custom.length && !confirm(
        `มี ${custom.length} เดือนที่เลือกวันปลูกเองไว้ (${custom.map((m) => MONTH_ABBR_TH[m]).join(", ")}) ` +
        "— การตั้งพร้อมกันจะเขียนทับวันที่เลือกไว้\n\nต้องการดำเนินการต่อหรือไม่?"
      )) return;
      for (const m of targets) {
        const s = calendarState[m];
        // กติกาเดียวกับชิปรายเดือน: คงไว้เฉพาะวันถ่ายภาพเดิมที่ยังเป็นวันปลูกอยู่ — ห้าม
        // ขยายวันถ่ายภาพตามวันปลูกชุดใหม่ (เช่น "ทุกวัน" จะกลายเป็นภาพหลายพันใบโดยไม่ตั้งใจ)
        s.days = daysForPreset(preset, MONTH_DAYS[m - 1]);
        s.shotDays = new Set([...s.shotDays].filter((d) => s.days.has(d)));
        openMonths.delete(m);
      }
    }
  } else {
    if (!active.length) {
      showToast("ยังไม่มีเดือนที่เลือกวันปลูก — ตั้งวันปลูกก่อน", true);
      return;
    }
    for (const m of active) {
      calendarState[m].shotDays = shotDaysForPreset(btn.dataset.bulkShot, calendarState[m].days);
    }
  }
  renderMonthGrid();
  scheduleCalendarAutoSave();
});

// ไฮไลต์ปุ่มลัดที่ตรงกับแผนตอนนี้: ปุ่มติดเมื่อ "ทุกเดือนที่ไม่ได้ข้าม" ตรงกับปุ่มนั้นพอดี
// (เทียบชุดวันตรงๆ — แผนที่ตรงได้หลายแบบพร้อมกัน เช่นวันปลูก 1,15 ถ่ายภาพ 1,15 จะติดทั้ง
// "1, 15" และ "ทั้งหมด" ซึ่งถูกทั้งคู่ ผู้ใช้กดปุ่มไหนก็เห็นปุ่มนั้นติด)
const sameSet = (a, b) => a.size === b.size && [...a].every((x) => b.has(x));
function syncBulkBar() {
  const active = activeMonthList();
  document.querySelectorAll("#bulk-bar [data-bulk-days]").forEach((c) => {
    const p = c.dataset.bulkDays;
    const on = p === "skip"
      ? active.length === 0
      : active.length > 0 && active.every((m) => sameSet(calendarState[m].days, daysForPreset(p, MONTH_DAYS[m - 1])));
    c.classList.toggle("selected", on);
  });
  document.querySelectorAll("#bulk-bar [data-bulk-shot]").forEach((c) => {
    const p = c.dataset.bulkShot;
    const on = active.length > 0 &&
      active.every((m) => sameSet(calendarState[m].shotDays, shotDaysForPreset(p, calendarState[m].days)));
    c.classList.toggle("selected", on);
  });
}

function updateSummary() {
  let totalDays = 0;
  let totalShots = 0;
  for (let m = 1; m <= 12; m++) {
    totalDays += calendarState[m].days.size;
    totalShots += calendarState[m].shotDays.size;
  }
  document.getElementById("sum-days").textContent = totalDays;
  document.getElementById("sum-shots").textContent = totalShots;
  // สรุปบนหัวการ์ด "แผนวันปลูก" — เห็นได้แม้ตอนพับการ์ดเก็บ
  const activeMonths = Object.values(calendarState).filter((s) => s.days.size > 0).length;
  document.getElementById("plan-days").textContent = totalDays;
  document.getElementById("plan-shots").textContent = totalShots;
  document.getElementById("plan-months").textContent = activeMonths ? `${activeMonths} เดือน` : "ยังไม่ได้เลือกเดือน";
  syncBulkBar();

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
  // โฟลเดอร์ต้นทางเพิ่งถูกบันทึก — อัปเดตเช็คลิสต์/แถบบันทึก/ชื่อใต้โลโก้ทันที
  document.getElementById("brand-sub").textContent = settings.input_dir || "ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง";
  updateReadiness();
  updateSaveBar();
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
    updateSaveBar(); // ตั้งค่าผ่านโค้ดไม่ยิง event "input" — แจ้งแถบบันทึกเอง
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
    updateSaveBar();
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
    badge.textContent = "ยังไม่ได้ตรวจสอบ";
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

  badge.textContent = "ตรวจสอบแล้ว";
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
const STATUS_LABEL = { done: "เสร็จสมบูรณ์", running: "กำลังดำเนินการ", queued: "รอดำเนินการ", error: "มีปัญหา" };
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
  lastSnapshot = snapshot;
  const years = snapshot.years || [];
  const doneCount = years.filter((y) => y.status === "done").length;
  const total = years.length;

  // ใช้ progress ระดับ "วันปลูก" ถ้ามี (ละเอียดกว่าระดับปีมาก — 1 ปีมีหลายวันปลูก
  // การนับแค่ปีทำให้ bar กระโดดทีละก้าวใหญ่ดูเหมือนค้าง) fallback เป็นระดับปี
  // v0.13 — ปีที่กำลังทำ + เวลาที่เหลือย้ายไปเป็นหัวข้อ/ป้าย ETA ของแผงหลัก
  // (updateRunHero) บรรทัดใต้แถบเหลือแค่ตัวเลขนับ อ่านได้ในแวบเดียว
  let pct;
  if (snapshot.candidate_total > 0) {
    pct = (snapshot.candidate_done / snapshot.candidate_total) * 100;
    document.getElementById("progress-txt").textContent =
      `${fmtNum(snapshot.candidate_done)} / ${fmtNum(snapshot.candidate_total)} วันปลูก · เสร็จ ${doneCount} / ${total} ปี`;
  } else {
    pct = total ? (doneCount / total) * 100 : 0;
    document.getElementById("progress-txt").textContent = `เสร็จ ${doneCount} / ${total} ปี`;
  }
  document.getElementById("progress-fill").style.width = `${pct}%`;
  document.getElementById("run-pct").textContent = `${Math.floor(pct)}%`;
  // backend ส่ง eta = null ทั้งตอนยังประมาณไม่ได้ และตอนทำครบทุกวันปลูกแล้วแต่ยังสรุปผล/
  // สร้าง Excel+Word อยู่ (สถานะยังเป็น running) — แยกข้อความให้ตรงกับสิ่งที่เกิดจริง
  // (ตอน "กำลังหยุด" ป้ายนี้ถูกซ่อนด้วย CSS)
  const wrappingUp = snapshot.candidate_total > 0 && snapshot.candidate_done >= snapshot.candidate_total;
  document.getElementById("run-eta").textContent = wrappingUp
    ? "ครบทุกวันปลูกแล้ว — กำลังสรุปผล..."
    : formatEta(snapshot.eta_seconds) || "กำลังประมาณเวลาที่เหลือ...";

  const isRunning = snapshot.overall_state === "running";
  // ระหว่าง "กำลังหยุด" ก็กดเริ่มไม่ได้ (backend ยังไม่ว่าง) — ถ้าค้างนาน ใช้ Force Close
  document.getElementById("btn-start").disabled = isRunning || snapshot.overall_state === "stopping";
  document.getElementById("btn-stop").disabled = !isRunning;

  // v0.5.20 — เดิมปุ่ม "รันปีที่มีปัญหาซ้ำ" โชว์ตลอดแม้ไม่มีปีไหน error เลย ทำให้
  // ผู้ใช้กดแล้วงงว่าทำไมขึ้น "ไม่มีปีที่มีปัญหา" (เจอจริงจากผู้ใช้) — ซ่อนปุ่มไปเลย
  // เมื่อ error count = 0 พร้อมโชว์ badge จำนวนปีที่มีปัญหาให้เห็นชัดโดยไม่ต้อง
  // เลื่อนดู year-list ทั้งหมด
  const errorCount = years.filter((y) => y.status === "error").length;
  const retryBtn = document.getElementById("btn-retry");
  retryBtn.hidden = errorCount === 0;
  retryBtn.disabled = isRunning || snapshot.overall_state === "stopping";
  const errorBadge = document.getElementById("error-count-badge");
  errorBadge.hidden = errorCount === 0;
  if (errorCount > 0) errorBadge.textContent = `${errorCount} ปีมีปัญหา`;
  document.getElementById("btn-copy-all-errors").hidden = errorCount === 0;
  // v0.13 — กล่องปีที่มีปัญหาอยู่บนสุดของการ์ดผลลัพธ์ แสดงเฉพาะตอนมีปัญหา
  document.getElementById("error-zone").hidden = errorCount === 0;

  // v1.0 — เริ่มรัน: พับแผนวันปลูกเก็บ ดันผลลัพธ์/ภาพหน้าจอขึ้นมาในจอแรก (พับครั้งเดียว
  // ตอนเปลี่ยนเป็นกำลังรัน ผู้ใช้กดเปิดดูเองได้) — จงใจ "ไม่" กางคืนตอนรันจบ: ตอนนั้นคือ
  // จังหวะที่ผู้ใช้ต้องเห็นปีที่มีปัญหา/ปุ่มเปิดไฟล์ ซึ่งการ์ดแผนที่กางจะดันลงไปใต้จอ
  if (isRunning && !wasRunning && planCard.open) planCard.open = false;

  // พอรันจบ (idle) รีเฟรชความคืบหน้าไฟล์ output อัตโนมัติ ให้เห็นว่าทำต่อได้ถึงไหน
  // + ดึงภาพหน้าจอสุดท้ายมาแสดง (เผื่อภาพบันทึกใบสุดท้ายมาหลังรอบ poll ล่าสุด)
  if (!isRunning && wasRunning) {
    fetchOutputProgress();
    refreshLatestShot();
    stopLive("การประมวลผลสิ้นสุดแล้ว — หยุด Live View อัตโนมัติ");
  }
  wasRunning = isRunning;

  // จอแสดงภาพ: ระหว่างรันคอยเช็คภาพใหม่ทุก ~3 วิ (throttle ในตัว — renderStatus
  // ถูกยิงจาก WebSocket ถี่กว่านั้นมาก)
  maybePollShot(isRunning);

  // v0.11.3 — สถานะรายปีเป็นตารางช่องเล็กๆ (สีบอกสถานะ ชี้ค้างดูชื่อสถานะ)
  // แทนแถวยาวทีละปี ส่วนปีที่มีปัญหาแยกไปเป็นรายการ error ด้านล่างให้อ่านง่าย
  const list = document.getElementById("year-list");
  list.innerHTML = "";
  for (const y of years) {
    const tile = document.createElement("div");
    tile.className = `yt ${y.status}`;
    tile.title = `${y.year} · ${STATUS_LABEL[y.status] || y.status}`;
    tile.textContent = y.year;
    list.appendChild(tile);
  }

  const errorList = document.getElementById("error-list");
  errorList.innerHTML = "";
  for (const y of years) {
    if (y.status !== "error" || !y.error_message) continue;
    const item = document.createElement("div");
    item.className = "err-item";
    const yearEl = document.createElement("span");
    yearEl.className = "err-year";
    yearEl.textContent = y.year;
    const text = document.createElement("span");
    text.className = "err-text";
    text.textContent = y.error_message;
    const actions = document.createElement("span");
    actions.className = "err-actions";
    // ถ้า error_message อ้างถึงภาพ error ที่ระบบถ่ายเก็บไว้ (v0.11.0) เพิ่มปุ่ม
    // เปิดดูภาพนั้นในหน้าต่างขยาย — ผู้ใช้เห็นได้เลยว่า CropWat แจ้ง error อะไร
    const shotMatch = /([\w-]+\.png)/.exec(y.error_message);
    if (shotMatch) {
      const viewBtn = document.createElement("button");
      viewBtn.type = "button";
      viewBtn.className = "link-btn";
      viewBtn.title = "เปิดดูภาพหน้าจอ error ที่ระบบถ่ายเก็บไว้";
      viewBtn.textContent = "ดูภาพ";
      viewBtn.addEventListener("click", () => {
        document.getElementById("screenshot-img").src = `/api/error-shots/${shotMatch[1]}`;
        document.getElementById("screenshot-caption").textContent = `ปี ${y.year} · ${shotMatch[1]}`;
        document.getElementById("screenshot-modal").hidden = false;
      });
      actions.appendChild(viewBtn);
    }
    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.className = "link-btn";
    copyBtn.title = "คัดลอกข้อความ error นี้";
    copyBtn.textContent = "คัดลอก";
    copyBtn.addEventListener("click", () => copyText(`ปี ${y.year}: ${y.error_message}`));
    actions.appendChild(copyBtn);
    item.append(yearEl, text, actions);
    errorList.appendChild(item);
  }

  updateRunHero();
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
    `ไฟล์ผลลัพธ์สมบูรณ์ <b>${data.total_done.toLocaleString("en-US")}</b> / ${data.total_expected.toLocaleString("en-US")} วันปลูก (${pct}%)` +
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
  lastOutput = data;
  outputStale = false;
  updateRunHero();
}

document.getElementById("btn-rescan").addEventListener("click", fetchOutputProgress);

// ---------------------------------------------------------------------------
// v1.0 — ความพร้อม (แบนเนอร์นำทาง + ขั้นตอนหน้าตั้งค่า) และแผงสถานะหลัก
// ทั้งหมดเป็นการ "แสดงผล" จากข้อมูลที่โหลดมาแล้ว: settings, calendarState,
// lastSnapshot, lastOutput, exeState — ไม่มีการเรียก API ในส่วนนี้
// ---------------------------------------------------------------------------
function readinessItems() {
  const classic = !!settings && settings.hidden_desktop_mode === false;
  return {
    input: !!(settings && settings.input_dir),
    output: !!(settings && settings.output_dir),
    // โหมดคลาสสิกผู้ใช้เปิด CropWat เอง ไม่ต้องรู้ตำแหน่งโปรแกรม — และถ้ายังตรวจ
    // ไม่เสร็จ/ตรวจไม่ได้ ก็ไม่ถือว่าขาด กันแบนเนอร์เด้งผิดๆ (backend แจ้งเองตอนเริ่มถ้าไม่เจอ)
    exe: classic || exeState !== "missing",
    classic,
  };
}

function isReady() {
  if (!settings) return true; // ยังโหลดไม่เสร็จ — อย่าเพิ่งฟ้องว่าไม่พร้อม
  const r = readinessItems();
  return r.input && r.output && r.exe;
}

function setChk(id, ok, detail) {
  const li = document.getElementById(id);
  li.classList.toggle("ok", ok);
  const d = li.querySelector(".chk-detail");
  d.textContent = detail;
  d.title = detail;
}

function updateReadiness() {
  renderExeStatus();
  if (settings) {
    const r = readinessItems();
    setChk("chk-input", r.input, r.input ? settings.input_dir : "ยังไม่ได้เลือก — ขั้นที่ 1 ในหน้าตั้งค่า");
    setChk("chk-output", r.output, r.output ? settings.output_dir : "ยังไม่ได้เลือก — ขั้นที่ 2 ในหน้าตั้งค่า");
    // เครื่องหมายถูกเฉพาะตอนเจอจริง (หรือโหมดคลาสสิก) — ตรวจไม่ได้ขึ้นเป็นข้อที่ต้องดู
    const exeOk = r.classic || exeState === "found";
    setChk(
      "chk-exe",
      exeOk,
      r.classic ? "โหมดคลาสสิก — เปิด CropWat เองก่อนเริ่ม"
        : exeState === "found" ? exePath || "พบโปรแกรมแล้ว"
        : exeState === "missing" ? "ไม่พบโปรแกรม — ระบุตำแหน่งในขั้นที่ 3"
        : exeState === "unknown" ? "ตรวจหาโปรแกรมไม่ได้ — ดูขั้นที่ 3"
        : "กำลังตรวจสอบ...",
    );
    document.getElementById("onboard-banner").hidden = isReady();
    document.getElementById("step-1").dataset.done = String(r.input);
    document.getElementById("step-2").dataset.done = String(r.output);
    document.getElementById("step-3").dataset.done = String(exeOk);
  }
  updateRunHero();
}

// สถานะของแผงหลัก (เรียงตามลำดับความสำคัญ):
//   running/stopping → ยังโหลดไม่เสร็จ → ยังตั้งค่าไม่ครบ → แผนว่าง (ไม่มีวันปลูก)
//   → หยุดกลางคันโดยมีปีที่มีปัญหา (ยังเหลือปีรอคิว) → จบแล้วมีปีที่มีปัญหา
//   → เสร็จครบ → พร้อม
// "เสร็จครบ" ยึดไฟล์จริงในเครื่องเป็นหลัก (สถานะในหน่วยความจำหายเมื่อปิดโปรแกรม และอาจ
// เป็นของแผนเก่า) และไม่เชื่อผลตรวจไฟล์ที่เก่ากว่าแผนปัจจุบัน (outputStale)
function deriveRunState() {
  const overall = lastSnapshot ? lastSnapshot.overall_state : "idle";
  if (overall === "running") return "running";
  if (overall === "stopping") return "stopping";
  if (!settings) return "loading";
  if (!isReady()) return "setup";
  if (!ALL_MONTHS.some((m) => calendarState[m] && calendarState[m].days.size > 0)) return "noplan";
  const years = lastSnapshot ? lastSnapshot.years || [] : [];
  if (years.some((y) => y.status === "error")) {
    return years.some((y) => y.status === "queued") ? "stopped" : "error";
  }
  if (!outputStale) {
    const doneOnDisk = !!lastOutput && lastOutput.total_expected > 0 && lastOutput.total_done >= lastOutput.total_expected;
    const doneInRun = years.length > 0 && years.every((y) => y.status === "done");
    if (lastOutput ? doneOnDisk : doneInRun) return "done";
  }
  return "ready";
}

function updateRunHero() {
  const card = document.getElementById("run-card");
  const state = deriveRunState();
  const snap = lastSnapshot;
  const years = snap ? snap.years || [] : [];
  const doneYears = years.filter((y) => y.status === "done").length;
  const errorYears = years.filter((y) => y.status === "error").length;
  const queuedYears = years.filter((y) => y.status === "queued").length;
  card.dataset.state = state === "loading" ? "ready" : state;
  card.dataset.progress = snap && (snap.candidate_total > 0 || doneYears || errorYears) ? "1" : "0";
  document.getElementById("view-dash").dataset.run = state;

  let badge;
  let title;
  let sub;
  switch (state) {
    case "loading":
      badge = "กำลังโหลด";
      title = "กำลังโหลดสถานะ...";
      sub = " ";
      break;
    case "setup": {
      const r = readinessItems();
      const missing = [r.input, r.output, r.exe].filter((ok) => !ok).length;
      badge = "ยังไม่พร้อม";
      title = "ตั้งค่าให้ครบก่อนเริ่มประมวลผล";
      sub = `ยังขาดอีก ${missing} ขั้น — ดูรายการด้านบน แล้วกด "ไปที่หน้าตั้งค่า"`;
      break;
    }
    case "noplan":
      badge = "ยังไม่พร้อม";
      title = "ยังไม่ได้เลือกวันปลูก";
      sub = "ข้ามอยู่ทุกเดือน — เลือกวันปลูกในการ์ด \"แผนวันปลูก\" ด้านล่างก่อนเริ่มประมวลผล";
      break;
    case "running": {
      const wrappingUp = snap.candidate_total > 0 && snap.candidate_done >= snap.candidate_total;
      // v1.3.0 — รันหลายหน้าต่างพร้อมกัน: มีหลายปีสถานะ running ได้ในเวลาเดียวกัน
      const runningYears = (snap.years || []).filter((y) => y.status === "running").map((y) => y.year);
      badge = "กำลังประมวลผล";
      title = wrappingUp
        ? "กำลังสรุปผล..."
        : runningYears.length > 1
          ? `กำลังคำนวณปี ${runningYears.join(", ")} พร้อมกัน`
          : runningYears.length === 1 || snap.current_year
            ? `กำลังคำนวณปี ${runningYears[0] || snap.current_year}`
            : "กำลังเตรียมงาน...";
      sub = wrappingUp
        ? "ประมวลผลครบทุกวันปลูกแล้ว — รอสักครู่ ระบบกำลังเก็บงานและสร้างไฟล์ผลลัพธ์"
        : settings && settings.hidden_desktop_mode === false
          ? "CropWat ทำงานบนหน้าจอจริง — อย่าใช้เมาส์หรือคีย์บอร์ดจนกว่าจะเสร็จ"
          : (snap.windows_active || runningYears.length) > 1
            ? `CropWat ${snap.windows_active || runningYears.length} หน้าต่างทำงานบนเดสก์ท็อปเบื้องหลัง — ใช้เครื่องทำงานอื่นได้ตามปกติ`
            : "CropWat ทำงานบนเดสก์ท็อปเบื้องหลัง — ใช้เครื่องทำงานอื่นได้ตามปกติ";
      break;
    }
    case "stopping":
      badge = "กำลังหยุด";
      title = "กำลังหยุดการประมวลผล...";
      sub = "รอให้วันปลูกที่ทำอยู่เสร็จก่อน งานที่เสร็จแล้วไม่หาย — ถ้าค้างนานผิดปกติ ใช้ Force Close";
      break;
    case "stopped":
      badge = "หยุดอยู่";
      title = `หยุดการประมวลผลแล้ว — ยังเหลือ ${queuedYears} ปี`;
      sub = `กด "เริ่มประมวลผล" เพื่อทำต่อ (ระบบข้ามส่วนที่เสร็จแล้ว) · มี ${errorYears} ปีที่มีปัญหา — ` +
        "\"ประมวลผลซ้ำปีที่มีปัญหา\" ทำเฉพาะปีที่มีปัญหาเท่านั้น";
      break;
    case "done":
      badge = "เสร็จสมบูรณ์";
      title = "ประมวลผลเสร็จครบแล้ว";
      sub = (lastOutput
        ? `${fmtNum(lastOutput.total_done)} วันปลูก · ${lastOutput.years.length} ปี`
        : `${doneYears} ปี`) + " — กด \"เปิด Excel ดูผล\" (sheet สรุป บอกวันปลูกที่ดีที่สุด)";
      break;
    case "error":
      badge = "มีปัญหา";
      title = `ประมวลผลจบแล้ว แต่มี ${errorYears} ปีที่มีปัญหา`;
      sub = "ดูสาเหตุได้ที่การ์ดผลลัพธ์ — แก้ไขข้อมูลแล้วกด \"ประมวลผลซ้ำปีที่มีปัญหา\" ส่วนผลของปีอื่นเปิดดูใน Excel ได้เลย";
      break;
    default: {
      const resumable = !outputStale && !!lastOutput &&
        lastOutput.total_done > 0 && lastOutput.total_done < lastOutput.total_expected;
      badge = "พร้อม";
      title = resumable ? "พร้อมทำต่อจากเดิม" : "พร้อมเริ่มประมวลผล";
      sub = resumable
        ? `เสร็จแล้ว ${fmtNum(lastOutput.total_done)} จาก ${fmtNum(lastOutput.total_expected)} วันปลูก — กด "เริ่มประมวลผล" ระบบจะทำเฉพาะส่วนที่เหลือ`
        : outputStale
          ? "แผนเปลี่ยนแล้ว — กด \"เริ่มประมวลผล\" ได้เลย (ระบบข้ามส่วนที่เสร็จแล้วให้เอง) หรือกด \"ตรวจไฟล์ใหม่\" ในการ์ดผลลัพธ์เพื่อนับงานที่เสร็จใหม่"
          : "ตรวจช่วงปีและแผนวันปลูกด้านล่าง แล้วกด \"เริ่มประมวลผล\"";
    }
  }
  document.getElementById("run-state-badge").textContent = badge;
  document.getElementById("run-state-title").textContent = title;
  document.getElementById("run-state-sub").textContent = sub;
  updateWinCtl(state);
}

// ---------------------------------------------------------------------------
// v1.3.0 — จำนวนหน้าต่าง CropWat ที่รันพร้อมกัน: ปุ่ม − / + ระหว่างรัน (มีผลทันที
// เฉพาะรอบที่รันอยู่ ไม่บันทึกทับค่าในหน้าตั้งค่า) + ข้อความ "อัตโนมัติ" ในหน้าตั้งค่า
// ---------------------------------------------------------------------------
let maxWindows = 4;
let winBusy = false;

function updateWinCtl(state) {
  const snap = lastSnapshot;
  const ctl = document.getElementById("win-ctl");
  const target = snap ? snap.windows_target || 0 : 0;
  const show = state === "running" && target > 0;
  ctl.hidden = !show;
  if (!show) return;
  const active = snap.windows_active || 0;
  document.getElementById("win-count").textContent = String(target);
  document.getElementById("btn-win-minus").disabled = winBusy || target <= 1;
  document.getElementById("btn-win-plus").disabled = winBusy || target >= maxWindows;
  let note;
  if (active > target) note = `กำลังปิด ${active - target} หน้าต่าง หลังวันปลูกที่ทำอยู่เสร็จ`;
  else if (active < target) note = `กำลังเปิดเพิ่ม... (ทำงานอยู่ ${active})`;
  else if (target <= 1) note = "เครื่องทำงานเบาที่สุดแล้ว";
  else note = "จะใช้คอมทำงานอื่นไปด้วย? กด − ลดภาระเครื่อง";
  document.getElementById("win-note").textContent = note;
}

async function changeWindows(delta) {
  const snap = lastSnapshot;
  if (!snap || winBusy) return;
  const next = Math.max(1, Math.min(maxWindows, (snap.windows_target || 1) + delta));
  if (next === snap.windows_target) return;
  winBusy = true;
  updateRunHero();
  try {
    const res = await fetch("/api/run/windows", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ windows: next }),
    });
    if (res.ok) {
      renderStatus(await res.json());
      showToast(delta < 0
        ? `ลดเหลือ ${next} หน้าต่าง — หน้าต่างที่เกินจะปิดหลังวันปลูกที่ทำอยู่เสร็จ งานไม่หาย`
        : `เพิ่มเป็น ${next} หน้าต่าง — กำลังเปิด CropWat เพิ่ม`);
    } else {
      showToast("ปรับจำนวนหน้าต่างไม่สำเร็จ — การประมวลผลอาจจบไปแล้ว", true);
    }
  } catch {
    showToast("ปรับจำนวนหน้าต่างไม่สำเร็จ — ติดต่อโปรแกรมไม่ได้", true);
  } finally {
    winBusy = false;
    updateRunHero();
  }
}

document.getElementById("btn-win-minus").addEventListener("click", () => changeWindows(-1));
document.getElementById("btn-win-plus").addEventListener("click", () => changeWindows(1));

async function loadCapacity() {
  try {
    const res = await fetch("/api/system/capacity");
    if (!res.ok) return;
    const cap = await res.json();
    maxWindows = cap.max || maxWindows;
    document.getElementById("parallel-auto-opt").textContent =
      `อัตโนมัติ (เครื่องนี้ ${cap.recommended} หน้าต่าง)`;
    const note = document.getElementById("capacity-note");
    note.textContent = `ตรวจเครื่องนี้: ${cap.reason}`;
    note.hidden = false;
  } catch {
    /* ไม่มีข้อมูลเครื่อง = แสดงแค่ "อัตโนมัติ (แนะนำ)" */
  }
}

// ไปหน้าตั้งค่า → โฟกัสช่องของขั้นแรกที่ยังไม่เสร็จ (คีย์บอร์ด/โปรแกรมอ่านหน้าจอไม่หลง)
document.getElementById("btn-onboard-setup").addEventListener("click", () => {
  showTab("setup");
  window.scrollTo({ top: 0 });
  const r = readinessItems();
  const target = !r.input ? "input-dir" : !r.output ? "output-dir" : null;
  if (target) document.getElementById(target).focus({ preventScroll: true });
  else if (!r.exe) document.getElementById("exe-details").scrollIntoView({ block: "center" });
});

// ปุ่ม Excel มี 2 จุด (แผงหลัก + การ์ดผลลัพธ์) เรียก buildAndOpen เดิม — ปิดทั้งคู่ระหว่าง
// สร้างไฟล์ กันกดซ้อนแล้วสร้าง Result.xlsx พร้อมกัน 2 รอบ (เดิมมีปุ่มเดียวที่ปิดตัวเองอยู่แล้ว)
const excelBtns = ["btn-hero-excel", "btn-build-excel"].map((id) => document.getElementById(id));
async function openExcel(btn) {
  if (excelBtns.some((b) => b.disabled)) return;
  excelBtns.forEach((b) => { b.disabled = true; });
  try {
    await buildAndOpen(btn, "excel");
  } finally {
    excelBtns.forEach((b) => { b.disabled = false; });
  }
}
excelBtns.forEach((b) => b.addEventListener("click", (e) => openExcel(e.currentTarget)));

document.getElementById("btn-hero-errors").addEventListener("click", () => {
  const zone = document.getElementById("error-zone");
  zone.scrollIntoView({ behavior: "smooth", block: "start" });
  zone.focus({ preventScroll: true });
});

// ---------------------------------------------------------------------------
// v0.13 — แถบบันทึกหน้าตั้งค่า: บอกเมื่อมีการแก้ไขที่ยังไม่ได้กดบันทึก
// (เทียบค่าในฟอร์มกับ settings ที่บันทึกแล้ว + จุดเตือนบนแท็บ "ตั้งค่า")
// ---------------------------------------------------------------------------
const SETUP_FIELDS = [
  ["input-dir", "input_dir", "text"],
  ["output-dir", "output_dir", "text"],
  ["cropwat-exe-path", "cropwat_exe_path", "text"],
  ["climate-station-dir", "climate_station_dir", "text"],
  ["rain-station-dir", "rain_station_dir", "text"],
  ["crop-file", "crop_file", "text"],
  ["soil-file", "soil_file", "text"],
  ["manual-per-candidate", "manual_minutes_per_candidate", "number"],
  ["hidden-desktop-mode", "hidden_desktop_mode", "bool"],
  ["auto-build-outputs", "auto_build_outputs", "bool"],
  ["speed-preset", "speed_preset", "select"],
  ["parallel-windows", "parallel_windows", "int0"],
  ["shift-year-per-candidate", "shift_year_per_candidate", "bool"],
];

function isSetupDirty() {
  if (!settings) return false;
  return SETUP_FIELDS.some(([id, key, kind]) => {
    const el = document.getElementById(id);
    if (kind === "bool") return el.checked !== (settings[key] !== false);
    if (kind === "number") return (Number(el.value) || 0) !== (Number(settings[key]) || 0);
    if (kind === "select") return el.value !== (settings[key] || "normal");
    if (kind === "int0") return (Number(el.value) || 0) !== (Number(settings[key]) || 0);
    return el.value !== (settings[key] || "");
  });
}

function updateSaveBar() {
  const dirty = isSetupDirty();
  document.getElementById("save-bar").classList.toggle("dirty", dirty);
  document.getElementById("save-state").textContent = dirty
    ? "มีการแก้ไขที่ยังไม่ได้บันทึก — กด \"บันทึกการตั้งค่า\" เพื่อใช้ค่าใหม่"
    : "บันทึกแล้ว — การตั้งค่าเป็นปัจจุบัน";
  document.querySelector('.tab-btn[data-tab="setup"]').classList.toggle("dirty", dirty);
}

document.getElementById("view-setup").addEventListener("input", updateSaveBar);
document.getElementById("view-setup").addEventListener("change", updateSaveBar);

// ---------------------------------------------------------------------------
// v1.2.0 — ค่าตั้งของ CropWat (หน้า Settings → Options) — backend เขียน Cropwat.ini
// ตาราง/ป้าย/เงื่อนไขการแสดงช่อง มาจาก backend (cropwat_options.SECTIONS) ทั้งหมด
// ---------------------------------------------------------------------------
let cwopt = null;      // ผลล่าสุดจาก GET /api/cropwat-options
let cwoptDraft = {};   // ค่าที่กำลังแก้ (key -> value)
let cwoptTab = "rain";

const cwoptField = (key) => {
  for (const s of cwopt.sections) for (const f of s.fields) if (f.key === key) return f;
  return null;
};
const cwoptVisible = (f) =>
  !f.when || Object.entries(f.when).every(([k, vals]) => vals.includes(Number(cwoptDraft[k])));
const cwoptDirtyKeys = () =>
  cwopt ? Object.keys(cwoptDraft).filter((k) => String(cwoptDraft[k]) !== String(cwopt.values[k])) : [];

function cwoptChoiceLabel(key) {
  const f = cwoptField(key);
  const c = f && f.choices.find((x) => x[0] === Number(cwopt.values[key]));
  return c ? c[1] : "-";
}

async function loadCropwatOptions() {
  try {
    const res = await fetch("/api/cropwat-options");
    cwopt = await res.json();
  } catch {
    cwopt = null;
  }
  cwoptDraft = cwopt ? { ...cwopt.values } : {};
  renderCropwatOptions();
}

function renderCropwatOptions() {
  const badge = document.getElementById("cwopt-badge");
  const warn = document.getElementById("cwopt-warn");
  const fieldsEl = document.getElementById("cwopt-fields");
  const summaryEl = document.getElementById("cwopt-summary");
  if (!cwopt) {
    badge.textContent = "อ่านค่าไม่ได้";
    badge.className = "badge warn";
    fieldsEl.innerHTML = "";
    summaryEl.hidden = true;
    return;
  }
  const dirty = cwoptDirtyKeys().length > 0;
  badge.textContent = !cwopt.exists ? "ไม่พบไฟล์" : dirty ? "ยังไม่บันทึก" : "ตรงกับ CropWat";
  badge.className = `badge ${!cwopt.exists ? "warn" : dirty ? "running" : "ok"}`;
  const warnings = [];
  if (!cwopt.exists) warnings.push("ไม่พบไฟล์ค่าตั้งของ CropWat — เปิดโปรแกรม CropWat 8.0 สักครั้งก่อน แล้วกลับมาที่หน้านี้");
  if (cwopt.cropwat_running) warnings.push("มี CropWat เปิดอยู่ — ค่าที่บันทึกจะมีผลเมื่อปิดแล้วเปิด CropWat ใหม่ (และอย่ากด Save as default ใน CropWat ตัวนั้น ไม่งั้นจะทับค่าจากหน้านี้)");
  warn.hidden = warnings.length === 0;
  warn.textContent = warnings.join(" · ");
  document.getElementById("cwopt-path").textContent = cwopt.path;
  document.getElementById("btn-cwopt-save").disabled = !dirty || !cwopt.exists;
  document.getElementById("btn-cwopt-reset").disabled = !dirty;
  document.getElementById("btn-cwopt-restore").hidden = !cwopt.has_backup;

  // แท็บตามหน้า Options ของ CropWat
  document.getElementById("cwopt-tabs").innerHTML = cwopt.sections
    .map((s) => `<button type="button" role="tab" class="seg-btn${s.id === cwoptTab ? " active" : ""}" data-sec="${s.id}">${s.title}</button>`)
    .join("");
  const sec = cwopt.sections.find((s) => s.id === cwoptTab) || cwopt.sections[0];
  document.getElementById("cwopt-desc").textContent = sec.desc;

  fieldsEl.innerHTML = "";
  for (const f of sec.fields) {
    if (!cwoptVisible(f)) continue;
    const row = document.createElement("label");
    row.className = "cwopt-row" + (String(cwoptDraft[f.key]) !== String(cwopt.values[f.key]) ? " changed" : "");
    const lbl = document.createElement("span");
    lbl.className = "cwopt-lbl";
    lbl.textContent = f.label;
    if (f.note) {
      const n = document.createElement("small");
      n.textContent = f.note;
      lbl.appendChild(n);
    }
    let ctl;
    if (f.type === "choice") {
      ctl = document.createElement("select");
      for (const [code, text] of f.choices) {
        const o = document.createElement("option");
        o.value = code;
        o.textContent = text;
        ctl.appendChild(o);
      }
    } else {
      ctl = document.createElement("input");
      ctl.type = "number";
      ctl.step = f.type === "float" ? "0.1" : "1";
      if (f.min != null) ctl.min = f.min;
      if (f.max != null) ctl.max = f.max;
      ctl.className = "tabular";
    }
    ctl.dataset.key = f.key;
    ctl.value = cwoptDraft[f.key] ?? "";
    const unit = document.createElement("span");
    unit.className = "cwopt-unit";
    unit.textContent = f.unit || "";
    row.append(lbl, ctl, unit);
    fieldsEl.appendChild(row);
  }

  // บรรทัดสรุปในการ์ดแผนวันปลูก — เห็นก่อนกดเริ่มว่ารอบนี้จะใช้ค่าอะไร
  summaryEl.hidden = !cwopt.exists;
  summaryEl.textContent =
    `ค่าตั้ง CropWat ที่จะใช้: Effective rainfall = ${cwoptChoiceLabel("ermethod")} · ` +
    `พืชทั่วไป = ${cwoptChoiceLabel("timopt")} (efficiency ${cwopt.values.appleff ?? "-"}%) · แก้ได้ที่แท็บตั้งค่า`;
}

document.getElementById("cwopt-tabs").addEventListener("click", (e) => {
  const b = e.target.closest("[data-sec]");
  if (!b) return;
  cwoptTab = b.dataset.sec;
  renderCropwatOptions();
});

const cwoptFieldsEl = document.getElementById("cwopt-fields");
cwoptFieldsEl.addEventListener("change", (e) => {
  const key = e.target.dataset.key;
  if (!key) return;
  cwoptDraft[key] = e.target.tagName === "SELECT" ? Number(e.target.value) : e.target.value;
  renderCropwatOptions(); // ตัวเลือกเปลี่ยน → ช่องที่เกี่ยวข้องเปลี่ยนตาม
});
cwoptFieldsEl.addEventListener("input", (e) => {
  const key = e.target.dataset.key;
  if (!key || e.target.tagName === "SELECT") return;
  cwoptDraft[key] = e.target.value;
  // ไม่ render ใหม่ทั้งหมด (เคอร์เซอร์จะหลุด) — อัปเดตแค่ป้ายสถานะ/ปุ่ม
  e.target.closest(".cwopt-row").classList.toggle("changed", String(e.target.value) !== String(cwopt.values[key]));
  const dirty = cwoptDirtyKeys().length > 0;
  document.getElementById("btn-cwopt-save").disabled = !dirty || !cwopt.exists;
  document.getElementById("btn-cwopt-reset").disabled = !dirty;
  const badge = document.getElementById("cwopt-badge");
  badge.textContent = dirty ? "ยังไม่บันทึก" : "ตรงกับ CropWat";
  badge.className = `badge ${dirty ? "running" : "ok"}`;
});

async function cwoptApply(res, okMsg) {
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    showToast(data.detail || "บันทึกค่าตั้ง CropWat ไม่สำเร็จ", true);
    return;
  }
  cwopt = data;
  cwoptDraft = { ...cwopt.values };
  renderCropwatOptions();
  showToast(cwopt.cropwat_running ? `${okMsg} — ปิดแล้วเปิด CropWat ใหม่เพื่อใช้ค่านี้` : okMsg);
}

document.getElementById("btn-cwopt-save").addEventListener("click", async () => {
  const values = {};
  for (const k of cwoptDirtyKeys()) values[k] = cwoptDraft[k];
  const res = await fetch("/api/cropwat-options", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ values }),
  });
  await cwoptApply(res, "บันทึกค่าลง CropWat แล้ว");
});

document.getElementById("btn-cwopt-reset").addEventListener("click", () => {
  cwoptDraft = { ...cwopt.values };
  renderCropwatOptions();
});

document.getElementById("btn-cwopt-restore").addEventListener("click", async () => {
  if (!confirm("สลับกลับไปใช้ค่าตั้ง CropWat ก่อนบันทึกครั้งล่าสุดหรือไม่? (กดซ้ำเพื่อสลับกลับมาได้)")) return;
  const res = await fetch("/api/cropwat-options/restore", { method: "POST" });
  await cwoptApply(res, "ย้อนกลับค่าตั้ง CropWat แล้ว");
});

document.getElementById("btn-cwopt-export").addEventListener("click", async () => {
  const pick = await fetch("/api/browse-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ initial_dir: (settings && settings.output_dir) || "" }),
  });
  const { path } = await pick.json();
  if (!path) return;
  const res = await fetch("/api/cropwat-options/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) showToast(data.detail || "ส่งออกไม่สำเร็จ", true);
  else showToast(`ส่งออกแล้ว: ${data.path}`);
});

document.getElementById("btn-cwopt-import").addEventListener("click", async () => {
  const pick = await fetch("/api/browse-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ initial_dir: (settings && settings.output_dir) || "", kind: "ini" }),
  });
  const { path } = await pick.json();
  if (!path) return;
  if (!confirm(`ใช้ค่าตั้ง CropWat จากไฟล์นี้แทนค่าปัจจุบันหรือไม่?\n\n${path}\n\n(ค่าเดิมสำรองไว้ — ย้อนกลับได้)`)) return;
  const res = await fetch("/api/cropwat-options/import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path }),
  });
  await cwoptApply(res, "นำเข้าค่าตั้ง CropWat แล้ว");
});

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
      `ต้องการเริ่มประมวลผลต่อหรือไม่? (วันปลูกที่มีปัญหาจะแสดงสถานะ "มีปัญหา" ส่วนรายการอื่นดำเนินการตามปกติ)`
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
    alert(err.detail || "ไม่สามารถเริ่มประมวลผลได้");
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
    "ระบบจะบังคับปิดโปรแกรม CropWat ทันที และคืนสถานะระบบให้พร้อมเริ่มประมวลผลใหม่\n\n" +
    "ใช้เฉพาะกรณี CropWat ไม่ตอบสนองเท่านั้น — ข้อมูลของรายการที่เสร็จสมบูรณ์แล้ว " +
    "จะไม่สูญหาย สามารถเริ่มประมวลผลต่อจากเดิมได้ทันที\n\n" +
    "ยืนยันการดำเนินการหรือไม่?"
  );
  if (!proceed) return;
  const res = await fetch("/api/run/force-close-cropwat", { method: "POST" });
  const data = await res.json().catch(() => ({}));
  showToast(data.killed ? `ปิด CropWat เรียบร้อย (${data.killed} รายการ) — พร้อมเริ่มประมวลผลใหม่` : "ไม่พบ CropWat ที่เปิดอยู่ — คืนสถานะระบบเรียบร้อย");
  const status = await fetch("/api/status");
  renderStatus(await status.json());
});

// ---------------------------------------------------------------------------
// จอแสดงภาพหน้าจอ CropWat (v0.9.0) — แสดงภาพล่าสุดโดยอัตโนมัติ (ภาพบันทึกงาน
// ตามปฏิทิน หรือSnapshot อันที่ใหม่กว่า) เทียบ mtime ก่อนโหลดภาพจริงเสมอ
// จะได้ไม่โหลดภาพซ้ำโดยไม่จำเป็น — trigger จาก: เปิดหน้า, ทุก ~3 วิระหว่างรัน
// (เกาะ renderStatus ที่ WebSocket ยิงเข้ามาอยู่แล้ว), รันจบ, และกด "Snapshot"
// ---------------------------------------------------------------------------
const SHOT_KIND_LABEL = { work: "ภาพบันทึกงาน", live: "Snapshot" };
let shotMtime = 0;
let shotFetchBusy = false;
let shotLastPollAt = 0;

function displayShot(data) {
  if (!data || !data.exists || data.mtime === shotMtime) return;
  shotMtime = data.mtime;
  const img = document.getElementById("shot-img");
  img.src = data.url;
  img.hidden = false;
  document.getElementById("shot-empty").hidden = true;
  const badge = document.getElementById("shot-kind-badge");
  badge.hidden = false;
  badge.textContent = SHOT_KIND_LABEL[data.kind] || data.kind;
  badge.className = "badge " + (data.kind === "live" ? "gold" : "accent");
  const time = new Date(data.mtime * 1000).toLocaleTimeString("th-TH", {
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  });
  document.getElementById("shot-meta").textContent = `${data.name} · ${time}`;
}

async function refreshLatestShot() {
  if (shotFetchBusy) return;
  shotFetchBusy = true;
  try {
    const res = await fetch("/api/screenshots/latest");
    if (res.ok) displayShot(await res.json());
  } catch {
    /* เครือข่ายสะดุดชั่วคราว — รอบ poll ถัดไปลองใหม่เอง */
  } finally {
    shotFetchBusy = false;
  }
}

function maybePollShot(isRunning) {
  if (!isRunning) return;
  const now = Date.now();
  if (now - shotLastPollAt < 3000) return;
  shotLastPollAt = now;
  refreshLatestShot();
}

document.getElementById("btn-take-screenshot").addEventListener("click", async (e) => {
  const btn = e.currentTarget;
  btn.disabled = true;
  try {
    const res = await fetch("/api/desktop/screenshot", { method: "POST" });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      alert(data.detail || "ไม่สามารถบันทึกภาพหน้าจอได้");
      return;
    }
    await refreshLatestShot();
    showToast("บันทึก Snapshot เรียบร้อย");
  } finally {
    btn.disabled = false;
  }
});

// ---------------------------------------------------------------------------
// โหมดLive View (v0.10.0) — สตรีมภาพนิ่งต่อเนื่อง: ขอภาพใหม่ทุก ~2 วิผ่านกลไก
// Snapshotเดิม (PrintWindow ที่จุดปลอดภัย — ไม่แตะการสลับจอเลย จึงไม่มีความ
// เสี่ยง "Cannot make a visible window modal" ที่เคยทำให้ต้องถอดฟีเจอร์Live View
// แบบสลับจอออกไป) ได้ภาพเคลื่อนไหว ~0.5-1 fps พอเห็นว่า CropWat ทำอะไรอยู่
// — หยุดเองอัตโนมัติเมื่อ: การประมวลผลจบ, ถ่ายภาพไม่ได้ (409), หรือซ่อนหน้าต่าง
// ---------------------------------------------------------------------------
let liveTimer = null;
let liveBusy = false;

function isLiveOn() {
  return liveTimer !== null;
}

function stopLive(message) {
  if (!isLiveOn()) return;
  clearInterval(liveTimer);
  liveTimer = null;
  document.getElementById("shot-live-badge").hidden = true;
  const btn = document.getElementById("btn-live-toggle");
  btn.classList.remove("live-on");
  btn.innerHTML = `<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="6 3 20 12 6 21 6 3"/></svg>Live View`;
  if (message) showToast(message);
}

async function liveTick() {
  if (liveBusy) return; // เฟรมก่อนหน้ายังไม่เสร็จ — ข้ามรอบนี้ กันคำขอซ้อน
  liveBusy = true;
  try {
    const res = await fetch("/api/desktop/screenshot", { method: "POST" });
    if (!res.ok) {
      stopLive(res.status === 409 ? "หยุด Live View — ระบบไม่ได้กำลังประมวลผลอยู่" : "หยุด Live View — ไม่สามารถถ่ายภาพได้");
      return;
    }
    await refreshLatestShot();
  } catch {
    /* เครือข่ายสะดุด — ปล่อยรอบถัดไปลองใหม่ */
  } finally {
    liveBusy = false;
  }
}

document.getElementById("btn-live-toggle").addEventListener("click", () => {
  if (isLiveOn()) {
    stopLive();
    return;
  }
  document.getElementById("shot-live-badge").hidden = false;
  const btn = document.getElementById("btn-live-toggle");
  btn.classList.add("live-on");
  btn.innerHTML = `<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="5" width="14" height="14" rx="2"/></svg>หยุด Live View`;
  liveTick();
  liveTimer = setInterval(liveTick, 2000);
});

// ซ่อนหน้าต่าง/สลับไปแท็บอื่น → หยุด Live View อัตโนมัติ (ไม่เปลืองแรงเครื่องโดยไม่มีคนดู)
document.addEventListener("visibilitychange", () => {
  if (document.hidden) stopLive();
});

// คลิกที่ภาพ → เปิดดูขนาดเต็มใน modal (พร้อมชื่อไฟล์/เวลากำกับ)
document.getElementById("shot-img").addEventListener("click", () => {
  const img = document.getElementById("shot-img");
  if (!img.src) return;
  document.getElementById("screenshot-img").src = img.src;
  document.getElementById("screenshot-caption").textContent =
    document.getElementById("shot-meta").textContent;
  document.getElementById("screenshot-modal").hidden = false;
});

document.getElementById("btn-minimize").addEventListener("click", async () => {
  showToast("ย่อหน้าต่างลงถาดระบบแล้ว — โปรแกรมยังทำงานต่อเบื้องหลัง");
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

// v0.12.0 — ปุ่ม "เปิด Excel/Word": สร้างไฟล์ใหม่จากผลล่าสุดเสมอแล้วเปิดให้ทันที
// (ผู้ใช้ต้องการ "ดูผล" ไม่ใช่แค่ "สร้างไฟล์" แล้วต้องไปหาเองในโฟลเดอร์) — ถ้าไฟล์
// เปิดค้างอยู่ใน Excel/Word (เขียนทับไม่ได้ 409) ก็เปิดฉบับเดิมให้ พร้อมบอกวิธีอัปเดต
async function buildAndOpen(btn, kind) {
  const label = kind === "excel" ? "Excel" : "Word";
  btn.disabled = true;
  try {
    const res = await fetch(`/api/build-${kind}`, { method: "POST" });
    const data = await res.json().catch(() => ({}));
    let stale = false;
    if (res.status === 409) {
      stale = true;
    } else if (!res.ok) {
      showToast(data.detail || `สร้างไฟล์ ${label} ไม่สำเร็จ`, true);
      return;
    }
    const open = await fetch(`/api/open-output/${kind}`, { method: "POST" });
    if (!open.ok) {
      const err = await open.json().catch(() => ({}));
      showToast(err.detail || `เปิดไฟล์ ${label} ไม่สำเร็จ`, true);
      return;
    }
    if (stale) showToast(`ไฟล์ ${label} เปิดค้างอยู่ จึงอัปเดตไม่ได้ — ปิดไฟล์แล้วกดใหม่เพื่อดูผลล่าสุด`, true);
    else showToast(kind === "excel" ? `เปิด Excel (${data.years_written} ปี)` : `เปิด Word (${data.candidates_written} วันปลูก)`);
  } finally {
    btn.disabled = false;
  }
}

document.getElementById("btn-build-word").addEventListener("click", (e) => buildAndOpen(e.currentTarget, "word"));

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
    } else if (manual && info.error) {
      // v1.1.1 — เดิมไม่ดู info.error เลย: เช็คไม่ได้ (เช่น GitHub จำกัดจำนวนครั้ง) แต่ขึ้นว่า
      // "ใช้เวอร์ชันล่าสุดอยู่แล้ว" ทำให้ผู้ใช้ไม่รู้ว่ามีเวอร์ชันใหม่
      showToast("ตรวจอัปเดตไม่ได้ตอนนี้ — ลองใหม่อีกครั้งภายหลัง หรือดาวน์โหลดเองจากหน้า Releases ของโปรแกรม", true);
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
if (location.hash === "#setup") showTab("setup");
fetchStatus();
fetchOutputProgress();
refreshLatestShot(); // แสดงภาพหน้าจอล่าสุดทันทีที่เปิดหน้า (ถ้ามีจากรอบก่อน)
refreshCropwatExeStatus(); // สถานะ "หา CropWat เจอไหม" ในหน้าตั้งค่า
loadCapacity(); // v1.3.0 — จำนวนหน้าต่างที่โหมดอัตโนมัติจะใช้
loadCropwatOptions(); // v1.2.0 — ค่าตั้งของ CropWat (การ์ดในหน้าตั้งค่า + สรุปในแผนวันปลูก)
connectWebSocket();
checkUpdate();
