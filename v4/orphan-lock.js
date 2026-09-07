"use strict";

// Fail-safe untuk sheet yang tertinggal terbuka dari versi/aksi lama.
// Jika tidak ada jendela shift 2,5 jam yang aktif dan tidak ada jendela
// Admin/Buka Sheet yang masih aktif, setiap sheet yang benar-benar UNPROTECTED
// akan langsung FULL LOCK saat panel V4 dibuka.

const ORPHAN_SHIFT_ID_KEY = "hitunganV4.shiftWorkSheetId";
const ORPHAN_SHIFT_UNTIL_KEY = "hitunganV4.shiftWorkUntil";
const ORPHAN_ADMIN_ID_KEY = "hitunganV4.adminUnlockSheetId";
const ORPHAN_ADMIN_UNTIL_KEY = "hitunganV4.adminUnlockUntil";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  setTimeout(() => lockOrphanUnprotectedSheets().catch(reportOrphanError), 1800);
});

async function lockOrphanUnprotectedSheets() {
  if (typeof getSetting !== "function" || typeof requireSecret !== "function" || typeof fullLock !== "function") return;

  const enabledRaw = getSetting("hitunganV4.enabled");
  const enabled = enabledRaw === true || String(enabledRaw).toLowerCase() === "true";
  if (!enabled) return;

  const now = Date.now();
  const shiftId = getSetting(ORPHAN_SHIFT_ID_KEY);
  const shiftUntil = Number(getSetting(ORPHAN_SHIFT_UNTIL_KEY));
  const adminId = getSetting(ORPHAN_ADMIN_ID_KEY);
  const adminUntil = Number(getSetting(ORPHAN_ADMIN_UNTIL_KEY));

  const validShiftId = shiftId && Number.isFinite(shiftUntil) && now < shiftUntil ? String(shiftId) : null;
  const validAdminId = adminId && Number.isFinite(adminUntil) && now < adminUntil ? String(adminId) : null;
  const pw = requireSecret();

  const lockedNames = await Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    sheets.load("items");
    wb.protection.load("protected");
    await context.sync();

    for (const sh of sheets.items) {
      sh.load("id,name,protection/protected");
    }
    await context.sync();

    const locked = [];
    for (const sh of sheets.items) {
      if (sh.protection.protected) continue;
      if (validShiftId && String(sh.id) === validShiftId) continue;
      if (validAdminId && String(sh.id) === validAdminId) continue;

      fullLock(sh, pw);
      locked.push(sh.name);
    }

    if (!wb.protection.protected) wb.protection.protect(pw);
    if (locked.length || !wb.protection.protected) await context.sync();
    return locked;
  });

  if (lockedNames.length && typeof setSecurityStatus === "function") {
    setSecurityStatus(`FAIL-SAFE: sheet yang tertinggal terbuka sudah FULL LOCK: ${lockedNames.join(", ")}.`, false);
  }
}

function reportOrphanError(e) {
  if (typeof setSecurityStatus !== "function") return;
  const text = e && e.message ? e.message : String(e || "gagal");
  setSecurityStatus(`FAIL-SAFE belum berhasil mengunci sheet terbuka: ${text}`, true);
}
