"use strict";

// Safety net for Admin / Koreksi -> Buka Sheet Aktif.
// A manually opened sheet is re-protected after 15 minutes.
// The pending deadline is stored in workbook settings so a later V4 open can
// recover and lock a sheet that was left open after the task pane was closed.

const ADMIN_SHEET_AUTOLOCK_MS = 15 * 60 * 1000;
const ADMIN_UNLOCK_SHEET_ID_KEY = "hitunganV4.adminUnlockSheetId";
const ADMIN_UNLOCK_SHEET_NAME_KEY = "hitunganV4.adminUnlockSheetName";
const ADMIN_UNLOCK_UNTIL_KEY = "hitunganV4.adminUnlockUntil";
let adminSheetRelockTimer = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  const unlockBtn = document.getElementById("adminUnlockSheetBtn");
  if (unlockBtn) {
    unlockBtn.addEventListener("click", () => {
      // Let the normal V4 admin handler validate the PIN and open the sheet first.
      setTimeout(captureAdminUnlockedSheet, 700);
    });
  }

  // If the admin manually locks/syncs before the 15-minute deadline, discard
  // the pending auto-lock once Excel confirms the sheet is protected again.
  ["adminWorkingLockBtn", "adminFullLockBtn", "adminSyncBtn"].forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) btn.addEventListener("click", () => setTimeout(clearPendingIfProtected, 900));
  });

  // Creating a new day normally re-locks the old sheet. Clear a stale pending
  // unlock if that has happened.
  const form = document.getElementById("newDayForm");
  if (form) form.addEventListener("submit", () => setTimeout(clearPendingIfProtected, 2500));

  restoreAdminSheetAutoLock().catch(() => {});
});

async function captureAdminUnlockedSheet() {
  try {
    const info = await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getActiveWorksheet();
      sh.load("id,name,protection/protected");
      await context.sync();
      return { id: sh.id, name: sh.name, protected: sh.protection.protected };
    });

    // Wrong PIN or another failure leaves the sheet protected; do nothing.
    if (info.protected) return;

    const until = Date.now() + ADMIN_SHEET_AUTOLOCK_MS;
    setSetting(ADMIN_UNLOCK_SHEET_ID_KEY, info.id);
    setSetting(ADMIN_UNLOCK_SHEET_NAME_KEY, info.name);
    setSetting(ADMIN_UNLOCK_UNTIL_KEY, until);
    await saveSettings();

    scheduleAdminSheetRelock(until);
    setSecurityStatus(`ADMIN: sheet "${info.name}" dibuka. Otomatis dikunci kembali dalam 15 menit.`, false);
  } catch (_) {
    // The normal admin handler remains authoritative; this safety helper must
    // never interfere with ordinary V4 use.
  }
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

  if (Date.now() >= until) {
    await relockPendingAdminSheet();
  } else {
    scheduleAdminSheetRelock(until);
  }
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
    const result = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sh = sheets.getItem(sheetId);
      sh.load("name,position,protection/protected");
      await context.sync();

      // If another V4 operation has already protected it, do not disturb it.
      if (!sh.protection.protected) {
        // V2 always places the current working sheet at the front. Therefore
        // position 0 returns to working-lock; older sheets return to full-lock.
        if (sh.position === 0) workingLock(sh, pw);
        else fullLock(sh, pw);
        await context.sync();
      }
      return { name: sh.name, position: sh.position };
    });

    await clearAdminSheetAutoLockState();
    const mode = result.position === 0 ? "sheet kerja" : "full lock";
    setSecurityStatus(`AUTO-LOCK: sheet "${result.name}" dikunci kembali sebagai ${mode}.`, false);
  } catch (e) {
    // Keep the pending state. The next time V4 is opened it will retry.
    setSecurityStatus(`AUTO-LOCK belum berhasil untuk "${sheetName}". V4 akan mencoba lagi saat panel dibuka.`, true);
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
  } catch (_) {
    // Sheet may have been deleted/renamed during an admin structure operation.
    // Do not throw into the main V4 UI.
  }
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
