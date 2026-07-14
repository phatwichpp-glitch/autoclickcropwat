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
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    const target = btn.dataset.tab;
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
  document.getElementById("climate-station-dir").value = settings.climate_station_dir || "";
  document.getElementById("rain-station-dir").value = settings.rain_station_dir || "";
  document.getElementById("crop-file").value = settings.crop_file || "";
  document.getElementById("soil-file").value = settings.soil_file || "";
  document.getElementById("manual-per-candidate").value = settings.manual_minutes_per_candidate;
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

document.getElementById("btn-save-calendar").addEventListener("click", async () => {
  const ok = await saveConfig({
    default_start_year: Number(document.getElementById("start-year").value),
    default_end_year: Number(document.getElementById("end-year").value),
    planting_calendar: calendarStateToPayload(),
  });
  if (ok) alert("บันทึกปฏิทินแล้ว");
});

document.getElementById("btn-save-setup").addEventListener("click", async () => {
  const ok = await saveConfig({
    input_dir: document.getElementById("input-dir").value,
    output_dir: document.getElementById("output-dir").value,
    climate_station_dir: document.getElementById("climate-station-dir").value,
    rain_station_dir: document.getElementById("rain-station-dir").value,
    crop_file: document.getElementById("crop-file").value,
    soil_file: document.getElementById("soil-file").value,
    manual_minutes_per_candidate: Number(document.getElementById("manual-per-candidate").value) || 0,
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

document.getElementById("start-year").addEventListener("input", updateSummary);
document.getElementById("end-year").addEventListener("input", updateSummary);

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

function renderStatus(snapshot) {
  if (!snapshot) return;
  const years = snapshot.years || [];
  const doneCount = years.filter((y) => y.status === "done").length;
  const total = years.length;

  const currentTxt = snapshot.current_year ? ` · กำลังรันปี ${snapshot.current_year}` : "";
  // ใช้ progress ระดับ "วันปลูก" ถ้ามี (ละเอียดกว่าระดับปีมาก — 1 ปีมีหลายวันปลูก
  // การนับแค่ปีทำให้ bar กระโดดทีละก้าวใหญ่ดูเหมือนค้าง) fallback เป็นระดับปี
  if (snapshot.candidate_total > 0) {
    const pct = (snapshot.candidate_done / snapshot.candidate_total) * 100;
    document.getElementById("progress-fill").style.width = `${pct}%`;
    document.getElementById("progress-txt").textContent =
      `${snapshot.candidate_done} / ${snapshot.candidate_total} วันปลูก${currentTxt}`;
  } else {
    document.getElementById("progress-fill").style.width = total ? `${(doneCount / total) * 100}%` : "0%";
    document.getElementById("progress-txt").textContent = `${doneCount} / ${total} ปี${currentTxt}`;
  }

  const isRunning = snapshot.overall_state === "running";
  document.getElementById("btn-start").disabled = isRunning;
  document.getElementById("btn-stop").disabled = !isRunning;
  document.getElementById("btn-retry").disabled = isRunning;

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
      detail.innerHTML = `<div class="err-msg">${y.error_message}</div>`;
      list.appendChild(detail);
    }
  }
}

async function fetchStatus() {
  const res = await fetch("/api/status");
  renderStatus(await res.json());
}

document.getElementById("btn-start").addEventListener("click", async () => {
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

let wsDroppedOnce = false;

function connectWebSocket() {
  const proto = location.protocol === "https:" ? "wss" : "ws";
  const ws = new WebSocket(`${proto}://${location.host}/ws`);
  ws.onopen = () => {
    // ต่อกลับได้หลังหลุด = backend เพิ่ง restart (เคสหลักคือเพิ่งอัปเดตเวอร์ชัน
    // เสร็จ) — reload หน้าให้ได้ frontend เวอร์ชันใหม่จาก .exe ตัวใหม่อัตโนมัติ
    if (wsDroppedOnce) { location.reload(); return; }
    setConn(true);
  };
  ws.onclose = () => { wsDroppedOnce = true; setConn(false); setTimeout(connectWebSocket, 2000); };
  ws.onerror = () => ws.close();
  ws.onmessage = (event) => renderStatus(JSON.parse(event.data));
}

// ---------------------------------------------------------------------------
// เช็คอัปเดตตอนเปิดโปรแกรม — มีเวอร์ชันใหม่ = โชว์ปุ่มอัปเดตที่มุมบน กดแล้ว
// backend ดาวน์โหลด+สลับไฟล์+restart ตัวเองให้ทั้งหมด (ดู backend/updater.py)
// ---------------------------------------------------------------------------
async function checkUpdate() {
  try {
    const res = await fetch("/api/update/check");
    const info = await res.json();
    document.getElementById("app-version").textContent = `v${info.current}`;
    if (info.update_available) {
      const btn = document.getElementById("btn-update");
      btn.hidden = false;
      btn.lastChild.textContent = `อัปเดตเป็น v${info.latest}`;
    }
  } catch { /* ออฟไลน์/เช็คไม่ได้ = ไม่ต้องโชว์อะไร */ }
}

document.getElementById("btn-update").addEventListener("click", async (e) => {
  const btn = e.currentTarget;
  if (!confirm("โปรแกรมจะปิดและเปิดขึ้นมาใหม่เป็นเวอร์ชันล่าสุดโดยอัตโนมัติ อัปเดตเลยหรือไม่?")) return;
  btn.disabled = true;
  try {
    const res = await fetch("/api/update/apply", { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      alert(data.detail || "อัปเดตไม่สำเร็จ");
      btn.disabled = false;
      return;
    }
    document.getElementById("conn-label").textContent = "กำลังอัปเดต...";
    // จากนี้ backend จะปิดตัวเอง → WS หลุด → พอตัวใหม่เปิด หน้าจะ reload เอง
  } catch {
    alert("อัปเดตไม่สำเร็จ (เชื่อมต่อ backend ไม่ได้)");
    btn.disabled = false;
  }
});

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
// Init
// ---------------------------------------------------------------------------
loadConfig();
fetchStatus();
connectWebSocket();
checkUpdate();
