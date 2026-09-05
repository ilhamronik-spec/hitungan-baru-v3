"use strict";

// Hotfix V4 activation: the original setup preferred the stored V3 secret and
// ignored a manually entered old password. This version tries all available
// legacy passwords independently for workbook structure and each worksheet.

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  const btn = document.getElementById("setupSecurityBtn");
  if (!btn) return;

  // Capture phase so this runs before the original V4 click handler.
  btn.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
    await setupSecurityFixed();
  }, true);
});

async function setupSecurityFixed() {
  const p1 = document.getElementById("setupPin").value;
  const p2 = document.getElementById("setupPin2").value;
  const oldInput = document.getElementById("oldPassword").value;

  if (!p1 || p1.length < 8) {
    setSecurityStatus("PIN admin minimal 8 karakter.", true);
    return;
  }
  if (p1 !== p2) {
    setSecurityStatus("PIN admin dan ulangannya tidak sama.", true);
    return;
  }

  const legacy = getSetting(LEGACY_V3_SECRET_KEY);
  const candidates = uniquePasswords([oldInput || null, legacy || null, null]);
  const newPw = randomSecret();

  try {
    setSecurityStatus("Membuka proteksi lama…", false);

    await unlockWorkbookWithCandidates(candidates);
    await unlockAllSheetsWithCandidates(candidates);

    setSecurityStatus("Proteksi lama terbuka. Memasang pengaman V4…", false);

    await Excel.run(async (context) => {
      const wb = context.workbook;
      const sheets = wb.worksheets;
      const active = sheets.getActiveWorksheet();

      sheets.load("items/id,items/name,items/protection/protected");
      active.load("id,name");
      wb.protection.load("protected");
      await context.sync();

      // Safety check: migration must never continue while an old protection
      // remains active.
      if (wb.protection.protected) {
        throw new Error("Struktur workbook masih terproteksi setelah proses migrasi.");
      }
      const stillProtected = sheets.items.filter(s => s.protection.protected);
      if (stillProtected.length) {
        throw new Error(`Masih ada sheet terproteksi: ${stillProtected.map(s => s.name).join(", ")}`);
      }

      for (const sh of sheets.items) {
        if (sh.id === active.id) workingLock(sh, newPw);
        else fullLock(sh, newPw);
      }
      wb.protection.protect(newPw);
      await context.sync();
    });

    setSetting(CFG.pinHash, await sha256(p1));
    setSetting(CFG.secret, newPw);
    setSetting(CFG.enabled, true);
    setSetting(CFG.version, "4.0.2-migration-fix");
    await saveSettings();

    secret = newPw;
    document.getElementById("setupPin").value = "";
    document.getElementById("setupPin2").value = "";
    document.getElementById("oldPassword").value = "";
    setConfigured(true);
    setFormEnabled(true);
    setSecurityStatus("AMAN — V4 aktif. Sheet aktif = kerja; sheet lama = full lock; struktur = lock.", false);
  } catch (e) {
    setSecurityStatus(friendlyMigrationError(e), true);
  }
}

function uniquePasswords(values) {
  const out = [];
  const seen = new Set();
  for (const value of values) {
    const key = value === null ? "__NO_PASSWORD__" : `pw:${value}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(value);
  }
  return out;
}

async function isWorkbookProtected() {
  return Excel.run(async (context) => {
    const wb = context.workbook;
    wb.protection.load("protected");
    await context.sync();
    return wb.protection.protected;
  });
}

async function unlockWorkbookWithCandidates(candidates) {
  if (!(await isWorkbookProtected())) return;
  let lastError = null;

  for (const pw of candidates) {
    try {
      await Excel.run(async (context) => {
        const wb = context.workbook;
        wb.protection.load("protected");
        await context.sync();
        if (!wb.protection.protected) return;
        if (pw === null) wb.protection.unprotect();
        else wb.protection.unprotect(pw);
        await context.sync();
      });
      if (!(await isWorkbookProtected())) return;
    } catch (e) {
      lastError = e;
    }
  }

  throw new Error("Struktur workbook tidak bisa dibuka dengan password proteksi yang tersedia." + debugSuffix(lastError));
}

async function getProtectedSheets() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,items/name,items/protection/protected");
    await context.sync();
    return sheets.items
      .filter(s => s.protection.protected)
      .map(s => ({ id: s.id, name: s.name }));
  });
}

async function sheetIsProtected(sheetId) {
  return Excel.run(async (context) => {
    const sh = context.workbook.worksheets.getItem(sheetId);
    sh.load("protection/protected");
    await context.sync();
    return sh.protection.protected;
  });
}

async function unlockAllSheetsWithCandidates(candidates) {
  const protectedSheets = await getProtectedSheets();

  for (const item of protectedSheets) {
    let unlocked = false;
    let lastError = null;

    for (const pw of candidates) {
      try {
        await Excel.run(async (context) => {
          const sh = context.workbook.worksheets.getItem(item.id);
          sh.load("protection/protected");
          await context.sync();
          if (!sh.protection.protected) return;
          if (pw === null) sh.protection.unprotect();
          else sh.protection.unprotect(pw);
          await context.sync();
        });
        if (!(await sheetIsProtected(item.id))) {
          unlocked = true;
          break;
        }
      } catch (e) {
        lastError = e;
      }
    }

    if (!unlocked && await sheetIsProtected(item.id)) {
      throw new Error(`Sheet "${item.name}" tidak bisa dibuka dengan password proteksi yang tersedia.` + debugSuffix(lastError));
    }
  }
}

function debugSuffix(e) {
  if (!e) return "";
  const loc = e.debugInfo && e.debugInfo.errorLocation ? ` @ ${e.debugInfo.errorLocation}` : "";
  const code = e.code ? ` [${e.code}]` : "";
  return `${code}${loc}`;
}

function friendlyMigrationError(e) {
  if (!e) return "Aktivasi gagal.";
  const base = e.message || String(e);
  const loc = e.debugInfo && e.debugInfo.errorLocation ? ` (${e.debugInfo.errorLocation})` : "";
  return `${base}${loc}`;
}
