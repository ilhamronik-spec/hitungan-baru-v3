"use strict";

// V4 safety rules:
// 1) Sheet baru berstatus sheet kerja maksimal 2,5 jam sejak selesai dibuat.
// 2) Karyawan boleh menutup shift lebih cepat tanpa PIN; hasilnya full lock.
// 3) Setelah full lock, hanya aksi Admin/PIN yang dapat menjadikannya sheet kerja lagi.
// 4) Admin -> Buka Sheet Aktif tetap punya safety-net 15 menit.

const ADMIN_SHEET_AUTOLOCK_MS = 15 * 60 * 1000;
const ADMIN_UNLOCK_SHEET_ID_KEY = "hitunganV4.adminUnlockSheetId";
const ADMIN_UNLOCK_SHEET_NAME_KEY = "hitunganV4.adminUnlockSheetName";
const ADMIN_UNLOCK_UNTIL_KEY = "hitunganV4.adminUnlockUntil";

const SHIFT_WORK_WINDOW_MS = 2.5 * 60 * 60 * 1000;
const SHIFT_SHEET_ID_KEY = "hitunganV4.shiftWorkSheetId";
const SHIFT_SHEET_NAME_KEY = "hitunganV4.shiftWorkSheetName";
const SHIFT_CREATED_AT_KEY = "hitunganV4.shiftCreatedAt";
const SHIFT_UNTIL_KEY = "hitunganV4.shiftWorkUntil";

let adminSheetRelockTimer = null;
let shiftWorkRelockTimer = null;
let shiftCreateWrapped = false;
let shiftWrapTimer = null;
let shiftUiTimer = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  bindAdminUnlockSafety();
  ensureEmployeeShiftUi();
  waitAndWrapShiftCreate();

  restoreAdminSheetAutoLock().catch(() => {});
  restoreShiftWorkWindow().catch(() => {});

  if (shiftUiTimer) clearInterval(shiftUiTimer);
  shiftUiTimer = setInterval(() => refreshShiftUi().catch(() => {}), 30000);
  setTimeout(() => refreshShiftUi().catch(() => {}), 800);
});

function bindAdminUnlockSafety() {
  const unlockBtn = document.getElementById("adminUnlockSheetBtn");
  if (unlockBtn) {
    unlockBtn.addEventListener("click", () => {
      // Biarkan handler V4 utama validasi PIN dan membuka sheet lebih dulu.
      setTimeout(captureAdminUnlockedSheet, 700);
    });
  }

  ["adminWorkingLockBtn", "adminFullLockBtn", "adminSyncBtn"].forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) btn.addEventListener("click", () => setTimeout(clearPendingIfProtected, 900));
  });

  const fullLockBtn = document.getElementById("adminFullLockBtn");
  if (fullLockBtn) {
    fullLockBtn.addEventListener("click", () => setTimeout(clearShiftWindowIfItsSheetIsFullLocked, 1100));
  }

  const form = document.getElementById("newDayForm");
  if (form) form.addEventListener("submit", () => setTimeout(clearPendingIfProtected, 2500));
}

function ensureEmployeeShiftUi() {
  if (document.getElementById("employeeCloseShiftBtn")) return;

  const form = document.getElementById("newDayForm");
  const createBtn = document.getElementById("createButton");
  if (!form || !createBtn) return;

  const status = document.createElement("div");
  status.id = "shiftWindowStatus";
  status.className = "product-summary";
  status.style.marginTop = "14px";
  status.textContent = "Memeriksa waktu pengisian shift...";

  const btn = document.createElement("button");
  btn.id = "employeeCloseShiftBtn";
  btn.type = "button";
  btn.className = "secondary-button";
  btn.innerHTML = '<span class="button-icon" aria-hidden="true">&#128274;</span><span>SELESAI / KUNCI SHIFT</span>';
  btn.title = "Kunci sheet shift sekarang tanpa PIN. Setelah dikunci, hanya admin yang dapat membukanya kembali sebagai sheet kerja.";
  btn.addEventListener("click", employeeCloseShiftNow);

  createBtn.parentNode.insertBefore(status, createBtn);
  createBtn.parentNode.insertBefore(btn, createBtn);
}

function waitAndWrapShiftCreate() {
  if (shiftWrapTimer) clearInterval(shiftWrapTimer);
  let attempts = 0;

  shiftWrapTimer = setInterval(() => {
    attempts += 1;
    try {
      if (shiftCreateWrapped) {
        clearInterval(shiftWrapTimer);
        shiftWrapTimer = null;
        return;
      }

      if (typeof createNewDaySheet !== "function" || typeof requireSecret !== "function") return;

      // Tunggu wrapper Tambah Display selesai agar pencatatan 2,5 jam menjadi
      // lapisan terluar dan baru dimulai setelah seluruh proses pembuatan sheet selesai.
      if (typeof displayTransferWrapped !== "undefined" && !displayTransferWrapped && attempts < 80) return;

      const previousCreate = createNewDaySheet;
      createNewDaySheet = async function(person, sheetName, purchaseData) {
        const result = await previousCreate(person, sheetName, purchaseData);
        await registerNewShiftWorkWindow(sheetName);
        return result;
      };

      shiftCreateWrapped = true;
      clearInterval(shiftWrapTimer);
      shiftWrapTimer = null;
    } catch (_) {
      // Coba lagi pada tick berikutnya.
    }
  }, 250);
}

async function registerNewShiftWorkWindow(expectedSheetName) {
  const info = await Excel.run(async (context) => {
    const sh = context.workbook.worksheets.getActiveWorksheet();
    sh.load("id,name,protection/protected");
    await context.sync();
    return { id: sh.id, name: sh.name, protected: sh.protection.protected };
  });

  if (expectedSheetName && info.name !== expectedSheetName) {
    throw new Error(`Sheet baru seharusnya "${expectedSheetName}", tetapi sheet aktif sekarang "${info.name}".`);
  }

  const createdAt = Date.now();
  const until = createdAt + SHIFT_WORK_WINDOW_MS;

  setSetting(SHIFT_SHEET_ID_KEY, info.id);
  setSetting(SHIFT_SHEET_NAME_KEY, info.name);
  setSetting(SHIFT_CREATED_AT_KEY, createdAt);
  setSetting(SHIFT_UNTIL_KEY, until);
  await saveSettings();

  scheduleShiftWorkRelock(until);
  await refreshShiftUi();
  setSecurityStatus(`SHIFT AKTIF — sheet "${info.name}" dapat diisi maksimal 2 jam 30 menit. Setelah itu otomatis full lock.`, false);
}

async function restoreShiftWorkWindow() {
  const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
  const untilRaw = getSetting(SHIFT_UNTIL_KEY);

  if (!sheetId || !untilRaw) {
    await refreshShiftUi();
    return;
  }

  const until = Number(untilRaw);
  if (!Number.isFinite(until)) {
    await clearShiftWorkWindowState();
    return;
  }

  if (Date.now() >= until) await expireShiftWorkWindow();
  else scheduleShiftWorkRelock(until);

  await refreshShiftUi();
}

function scheduleShiftWorkRelock(until) {
  if (shiftWorkRelockTimer) clearTimeout(shiftWorkRelockTimer);
  const delay = Math.max(0, Number(until) - Date.now());
  shiftWorkRelockTimer = setTimeout(() => {
    expireShiftWorkWindow().catch(() => {});
  }, delay);
}

async function expireShiftWorkWindow() {
  const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
  const sheetName = getSetting(SHIFT_SHEET_NAME_KEY) || "sheet";
  const until = Number(getSetting(SHIFT_UNTIL_KEY));
  if (!sheetId) return;

  if (Number.isFinite(until) && Date.now() < until) {
    scheduleShiftWorkRelock(until);
    return;
  }

  try {
    const pw = requireSecret();
    const result = await Excel.run(async (context) => {
      const wb = context.workbook;
      const sh = wb.worksheets.getItem(sheetId);
      sh.load("name,protection/protected");
      wb.protection.load("protected");
      await context.sync();

      if (sh.protection.protected) {
        sh.protection.unprotect(pw);
        await context.sync();
      }

      fullLock(sh, pw);
      if (!wb.protection.protected) wb.protection.protect(pw);
      await context.sync();
      return sh.name;
    });

    await clearShiftWorkWindowState();
    await refreshShiftUi();
    setSecurityStatus(`BATAS WAKTU SELESAI — sheet "${result}" otomatis FULL LOCK setelah 2 jam 30 menit. Hanya admin/PIN yang dapat menjadikannya sheet kerja lagi.`, false);
  } catch (e) {
    // Pertahankan state agar V4 mencoba lagi saat panel dibuka berikutnya.
    setSecurityStatus(`AUTO-LOCK 2,5 jam belum berhasil untuk "${sheetName}". V4 akan mencoba lagi saat panel dibuka.`, true);
  }
}

async function employeeCloseShiftNow() {
  try {
    const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
    const sheetName = getSetting(SHIFT_SHEET_NAME_KEY) || "sheet";
    if (!sheetId) throw new Error("Tidak ada sheet kerja aktif yang bisa ditutup.");

    const pw = requireSecret();
    const result = await Excel.run(async (context) => {
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load("id,name,protection/protected");
      wb.protection.load("protected");
      await context.sync();

      if (active.id !== sheetId) {
        throw new Error(`Buka sheet kerja "${sheetName}" terlebih dahulu sebelum menutup shift.`);
      }

      if (active.protection.protected) {
        active.protection.unprotect(pw);
        await context.sync();
      }

      fullLock(active, pw);
      if (!wb.protection.protected) wb.protection.protect(pw);
      await context.sync();
      return active.name;
    });

    await clearShiftWorkWindowState();
    await refreshShiftUi();
    setSecurityStatus(`SHIFT DITUTUP — sheet "${result}" sudah FULL LOCK oleh karyawan. Perubahan berikutnya memerlukan Admin/PIN.`, false);
  } catch (e) {
    setSecurityStatus(errText(e), true);
  }
}

function isEmployeeWorkWindowActiveForSheet(sheetId) {
  const currentId = getSetting(SHIFT_SHEET_ID_KEY);
  const until = Number(getSetting(SHIFT_UNTIL_KEY));
  return Boolean(currentId && currentId === sheetId && Number.isFinite(until) && Date.now() < until);
}

async function refreshShiftUi() {
  ensureEmployeeShiftUi();
  const el = document.getElementById("shiftWindowStatus");
  const btn = document.getElementById("employeeCloseShiftBtn");
  if (!el || !btn) return;

  const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
  const sheetName = getSetting(SHIFT_SHEET_NAME_KEY) || "";
  const until = Number(getSetting(SHIFT_UNTIL_KEY));

  if (!sheetId || !Number.isFinite(until)) {
    el.textContent = "Tidak ada shift dalam masa pengisian 2,5 jam.";
    btn.disabled = true;
    return;
  }

  const remaining = until - Date.now();
  if (remaining <= 0) {
    el.textContent = `Waktu pengisian "${sheetName}" sudah habis. Menunggu full lock...`;
    btn.disabled = true;
    expireShiftWorkWindow().catch(() => {});
    return;
  }

  const totalMinutes = Math.ceil(remaining / 60000);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  const timeText = hours > 0 ? `${hours} jam ${minutes} menit` : `${minutes} menit`;
  el.textContent = `Sheet kerja: ${sheetName} — sisa waktu pengisian sekitar ${timeText}. Karyawan boleh menutup lebih cepat.`;
  btn.disabled = false;
}

async function clearShiftWorkWindowState() {
  if (shiftWorkRelockTimer) clearTimeout(shiftWorkRelockTimer);
  shiftWorkRelockTimer = null;
  const settings = Office.context.document.settings;
  settings.remove(SHIFT_SHEET_ID_KEY);
  settings.remove(SHIFT_SHEET_NAME_KEY);
  settings.remove(SHIFT_CREATED_AT_KEY);
  settings.remove(SHIFT_UNTIL_KEY);
  await saveSettings();
}

async function clearShiftWindowIfItsSheetIsFullLocked() {
  const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
  if (!sheetId) return;

  try {
    const state = await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getItem(sheetId);
      sh.load("protection/protected");
      await context.sync();
      return sh.protection.protected;
    });
    if (state) {
      await clearShiftWorkWindowState();
      await refreshShiftUi();
    }
  } catch (_) {}
}

// ---------------- Admin temporary unlock safety (15 minutes) ----------------

async function captureAdminUnlockedSheet() {
  try {
    const info = await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getActiveWorksheet();
      sh.load("id,name,protection/protected");
      await context.sync();
      return { id: sh.id, name: sh.name, protected: sh.protection.protected };
    });

    // PIN salah / gagal membuka: sheet tetap protected, jadi tidak membuat timer.
    if (info.protected) return;

    const until = Date.now() + ADMIN_SHEET_AUTOLOCK_MS;
    setSetting(ADMIN_UNLOCK_SHEET_ID_KEY, info.id);
    setSetting(ADMIN_UNLOCK_SHEET_NAME_KEY, info.name);
    setSetting(ADMIN_UNLOCK_UNTIL_KEY, until);
    await saveSettings();

    scheduleAdminSheetRelock(until);
    setSecurityStatus(`ADMIN: sheet "${info.name}" dibuka. Otomatis dikunci kembali dalam 15 menit.`, false);
  } catch (_) {}
}

async function restoreAdminSheetAutoLock() {
  const sheetId = getSetting(ADMIN_UNLOCK_SHEET_ID_KEY);
  const untilRaw = getSetting(ADMIN_UNLOCK_UNTIL_KEY);
  if (!sheetId || !untilRaw) return;

  const until = Number(untilRaw);
  if (!Number.isFinite(until)) {
    await clearAdminSheetAutoLockState();
    return;
  }

  if (Date.now() >= until) await relockPendingAdminSheet();
  else scheduleAdminSheetRelock(until);
}

function scheduleAdminSheetRelock(until) {
  if (adminSheetRelockTimer) clearTimeout(adminSheetRelockTimer);
  const delay = Math.max(0, Number(until) - Date.now());
  adminSheetRelockTimer = setTimeout(() => {
    relockPendingAdminSheet().catch(() => {});
  }, delay);
}

async function relockPendingAdminSheet() {
  const sheetId = getSetting(ADMIN_UNLOCK_SHEET_ID_KEY);
  const sheetName = getSetting(ADMIN_UNLOCK_SHEET_NAME_KEY) || "sheet";
  if (!sheetId) return;

  try {
    const pw = requireSecret();
    const shouldReturnToWorking = isEmployeeWorkWindowActiveForSheet(sheetId);

    const result = await Excel.run(async (context) => {
      const wb = context.workbook;
      const sh = wb.worksheets.getItem(sheetId);
      sh.load("name,protection/protected");
      wb.protection.load("protected");
      await context.sync();

      if (!sh.protection.protected) {
        if (shouldReturnToWorking) workingLock(sh, pw);
        else fullLock(sh, pw);
        if (!wb.protection.protected) wb.protection.protect(pw);
        await context.sync();
      }
      return sh.name;
    });

    await clearAdminSheetAutoLockState();
    const mode = shouldReturnToWorking ? "sheet kerja (masih dalam 2,5 jam)" : "full lock";
    setSecurityStatus(`AUTO-LOCK ADMIN: sheet "${result}" dikunci kembali sebagai ${mode}.`, false);
  } catch (e) {
    setSecurityStatus(`AUTO-LOCK admin belum berhasil untuk "${sheetName}". V4 akan mencoba lagi saat panel dibuka.`, true);
  }
}

async function clearPendingIfProtected() {
  const sheetId = getSetting(ADMIN_UNLOCK_SHEET_ID_KEY);
  if (!sheetId) return;

  try {
    const protectedNow = await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getItem(sheetId);
      sh.load("protection/protected");
      await context.sync();
      return sh.protection.protected;
    });
    if (protectedNow) await clearAdminSheetAutoLockState();
  } catch (_) {}
}

async function clearAdminSheetAutoLockState() {
  if (adminSheetRelockTimer) clearTimeout(adminSheetRelockTimer);
  adminSheetRelockTimer = null;
  const settings = Office.context.document.settings;
  settings.remove(ADMIN_UNLOCK_SHEET_ID_KEY);
  settings.remove(ADMIN_UNLOCK_SHEET_NAME_KEY);
  settings.remove(ADMIN_UNLOCK_UNTIL_KEY);
  await saveSettings();
}
