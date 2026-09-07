"use strict";

// Hotfix: tombol karyawan harus tetap bisa menutup sheet terbaru walaupun
// sheet tersebut dibuat sebelum pencatatan timer 2,5 jam tersedia.

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  setTimeout(bindShiftLockFallback, 1200);
  setInterval(() => refreshFallbackShiftButton().catch(() => {}), 5000);
});

function bindShiftLockFallback() {
  const btn = document.getElementById("employeeCloseShiftBtn");
  if (!btn) {
    setTimeout(bindShiftLockFallback, 700);
    return;
  }
  if (btn.dataset.fallbackBound === "1") return;
  btn.dataset.fallbackBound = "1";

  btn.addEventListener("click", async (event) => {
    const shiftId = typeof getSetting === "function" ? getSetting("hitunganV4.shiftWorkSheetId") : null;
    const shiftUntil = typeof getSetting === "function" ? Number(getSetting("hitunganV4.shiftWorkUntil")) : NaN;
    const trackedAndValid = Boolean(shiftId && Number.isFinite(shiftUntil) && Date.now() < shiftUntil);
    if (trackedAndValid) return;

    event.preventDefault();
    event.stopImmediatePropagation();
    await employeeFallbackCloseLatestSheet();
  }, true);

  refreshFallbackShiftButton().catch(() => {});
}

async function refreshFallbackShiftButton() {
  const btn = document.getElementById("employeeCloseShiftBtn");
  const status = document.getElementById("shiftWindowStatus");
  if (!btn || !status || typeof getSetting !== "function") return;

  const shiftId = getSetting("hitunganV4.shiftWorkSheetId");
  const shiftUntil = Number(getSetting("hitunganV4.shiftWorkUntil"));
  if (shiftId && Number.isFinite(shiftUntil) && Date.now() < shiftUntil) return;

  const state = await Excel.run(async (context) => {
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name,position,protection/protected");
    await context.sync();
    return { name: active.name, position: active.position, protected: active.protection.protected };
  });

  if (state.position === 0) {
    btn.disabled = false;
    status.textContent = state.protected
      ? `Sheet terbaru: ${state.name}. Tombol kunci tetap tersedia untuk memastikan FULL LOCK.`
      : `Sheet terbaru: ${state.name}. Belum ada timer 2,5 jam tersimpan; karyawan bisa kunci sekarang.`;
  }
}

async function employeeFallbackCloseLatestSheet() {
  try {
    if (typeof requireSecret !== "function" || typeof fullLock !== "function") {
      throw new Error("Mesin pengaman V4 belum siap.");
    }

    const pw = requireSecret();
    const result = await Excel.run(async (context) => {
      const wb = context.workbook;
      const active = wb.worksheets.getActiveWorksheet();
      active.load("name,position,protection/protected");
      wb.protection.load("protected");
      await context.sync();

      if (active.position !== 0) {
        throw new Error("Tombol ini hanya untuk sheet shift terbaru.");
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

    if (typeof clearShiftWorkWindowState === "function") {
      try { await clearShiftWorkWindowState(); } catch (_) {}
    }

    const btn = document.getElementById("employeeCloseShiftBtn");
    if (btn) btn.disabled = true;
    if (typeof setSecurityStatus === "function") {
      setSecurityStatus(`SHIFT DITUTUP — sheet "${result}" sudah FULL LOCK. Untuk membuka lagi harus Admin/PIN.`, false);
    }
  } catch (e) {
    if (typeof setSecurityStatus === "function") {
      const text = e && e.message ? e.message : String(e || "Terjadi kesalahan.");
      setSecurityStatus(text, true);
    }
  }
}
