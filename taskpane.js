const CFG = {
  enabled: "hitunganV3.enabled",
  pinHash: "hitunganV3.pinHash",
  secret: "hitunganV3.secret",
  version: "hitunganV3.version"
};

const WORK_AREA = "A1:AH400";
const LOCK_RANGES = [
  "B267","B269","B271:B275","G272:G275","F178:F238",
  "K178:K238","L178:L238","M178:M238","N178:N238","O178:O238",
  "N163","B6:B75","C6:C75","D6:D75","F6:F75","G6:G75","H6:H75",
  "M23","L50","K75","L75","H76:H79","B114:B116","B145:B147",
  "B163:B165","J168:J170","F171:F173","R249:R251","K145:K147",
  "W143:W145","W157:W159","W129:W131","Y178","Y237"
];

const IDLE_MS = 5000;
const CREATE_WINDOW_MS = 90000;

let runtimeSecret = null;
let handlersRegistered = false;
let pendingSheetId = null;
let lastChangeAt = 0;
let settleTimer = null;
let creationTimer = null;
let createWindowOpen = false;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  bindUi();

  try {
    const enabled = getSetting(CFG.enabled) === true;
    const secret = getSetting(CFG.secret);

    if (enabled && secret) {
      runtimeSecret = secret;
      await registerWorkbookEvents();
      await enforceFailClosed(secret);
      setConfiguredUi(true);
      setStatus("on", "Fail-closed aktif. Struktur workbook terkunci.");
      msg("V3 aktif. Karyawan boleh membuat hari baru tanpa PIN; sheet lama tetap terlindungi.");
    } else {
      setConfiguredUi(false);
      setStatus("off", "Belum dikonfigurasi.");
    }
  } catch (e) {
    setStatus("off", "Gagal memulai V3.");
    msg(errorText(e));
  }
});

function bindUi() {
  byId("setupBtn").addEventListener("click", setupV3);
  byId("beginDayBtn").addEventListener("click", beginCreateDayWindow);
  byId("unlockBtn").addEventListener("click", unlockActiveSheet);
  byId("partialLockBtn").addEventListener("click", lockActiveAsWorkingSheet);
  byId("fullLockBtn").addEventListener("click", fullLockActiveSheet);
  byId("syncBtn").addEventListener("click", syncAllSheets);
  byId("lockNowBtn").addEventListener("click", lockStructureNow);
  byId("hideBtn").addEventListener("click", hidePane);
}

async function setupV3() {
  const pin1 = byId("setupPin").value;
  const pin2 = byId("setupPin2").value;
  const oldPassword = byId("oldPassword").value;

  if (!pin1 || pin1.length < 8) return msg("PIN admin minimal 8 karakter.");
  if (pin1 !== pin2) return msg("PIN admin dan pengulangannya tidak sama.");

  try {
    setStatus("wait", "Menyiapkan fail-closed…");
    const pinHash = await sha256(pin1);
    const secret = generateSecret();

    await migrateAndSync(secret, oldPassword || null);

    setSetting(CFG.pinHash, pinHash);
    setSetting(CFG.secret, secret);
    setSetting(CFG.enabled, true);
    setSetting(CFG.version, "3.1.1-failclosed");
    await saveSettings();

    runtimeSecret = secret;
    await registerWorkbookEvents();

    try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (_) {}

    byId("setupPin").value = "";
    byId("setupPin2").value = "";
    byId("oldPassword").value = "";

    setConfiguredUi(true);
    setStatus("on", "Fail-closed aktif. Struktur workbook terkunci.");
    msg("Berhasil. Sheet aktif = sheet kerja; sheet lain = full lock; struktur workbook = lock.");
  } catch (e) {
    setStatus("off", "Setup gagal.");
    msg("Setup gagal.\nJika ada proteksi lama, isi password proteksi lama yang benar.\n\n" + errorText(e));
  }
}

async function migrateAndSync(newSecret, oldPassword) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    const active = sheets.getActiveWorksheet();

    wb.load("protection/protected");
    sheets.load("items/id,items/name,items/protection/protected");
    active.load("id");
    await context.sync();

    if (wb.protection.protected) {
      if (!oldPassword) throw new Error("Workbook sudah diproteksi. Masukkan password proteksi lama.");
      wb.protection.unprotect(oldPassword);
    }

    for (const sheet of sheets.items) {
      if (sheet.protection.protected) {
        if (!oldPassword) throw new Error(`Sheet "${sheet.name}" sudah diproteksi. Masukkan password proteksi lama.`);
        sheet.protection.unprotect(oldPassword);
      }
    }
    await context.sync();

    for (const sheet of sheets.items) {
      if (sheet.id === active.id) prepareWorkingSheet(sheet, newSecret);
      else prepareFullLock(sheet, newSecret);
    }

    wb.protection.protect(newSecret);
    await context.sync();
  });
}

async function enforceFailClosed(secret) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    wb.load("protection/protected");
    await context.sync();
    if (!wb.protection.protected) {
      wb.protection.protect(secret);
      await context.sync();
    }
  });
}

async function registerWorkbookEvents() {
  if (handlersRegistered) return;

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.onAdded.add(onWorksheetAdded);
    sheets.onChanged.add(onWorksheetChanged);
    await context.sync();
  });

  handlersRegistered = true;
}

async function beginCreateDayWindow() {
  try {
    const secret = requireSecret();

    if (createWindowOpen) {
      msg("Jendela pembuatan hari baru masih aktif. Jalankan BUAT HARI BARU V2 sekarang.");
      return;
    }

    await Excel.run(async (context) => {
      const wb = context.workbook;
      wb.load("protection/protected");
      await context.sync();

      // HANYA struktur workbook yang dibuka. Isi sheet lama tetap diproteksi.
      if (wb.protection.protected) wb.protection.unprotect(secret);
      await context.sync();
    });

    createWindowOpen = true;
    pendingSheetId = null;
    setStatus("wait", "Jendela hari baru aktif 90 detik. Jalankan BUAT HARI BARU V2.");
    msg("Tidak perlu PIN. Isi sheet lama tetap terkunci. Struktur workbook akan dikunci kembali otomatis.");

    if (creationTimer) clearTimeout(creationTimer);
    creationTimer = setTimeout(async () => {
      if (!createWindowOpen || pendingSheetId) return;
      try {
        await relockWorkbookStructure(secret);
        createWindowOpen = false;
        setStatus("on", "Waktu habis. Struktur workbook dikunci kembali.");
        msg("Tidak ada sheet baru dalam 90 detik. Silakan klik IZINKAN BUAT HARI BARU lagi jika diperlukan.");
      } catch (e) {
        setStatus("off", "Gagal mengunci kembali struktur workbook.");
        msg(errorText(e));
      }
    }, CREATE_WINDOW_MS);
  } catch (e) {
    setStatus("off", "Tidak dapat membuka jendela hari baru.");
    msg(errorText(e));
  }
}

function onWorksheetAdded(event) {
  if (!runtimeSecret) return;

  pendingSheetId = event.worksheetId;
  lastChangeAt = Date.now();

  setStatus("wait", "Sheet baru terdeteksi. Menunggu BUAT HARI BARU V2 selesai…");

  // Jika V2 menyalin sheet lama beserta proteksinya, buka HANYA sheet baru agar V2 dapat menulis.
  makeNewSheetWritable(event.worksheetId).catch((e) => msg(errorText(e)));
  startSettleTimer();
}

function onWorksheetChanged(event) {
  if (!pendingSheetId) return;
  if (event.worksheetId === pendingSheetId) lastChangeAt = Date.now();
}

async function makeNewSheetWritable(sheetId) {
  const secret = requireSecret();
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    sheet.load("protection/protected");
    await context.sync();
    if (sheet.protection.protected) {
      sheet.protection.unprotect(secret);
      await context.sync();
    }
  });
}

function startSettleTimer() {
  if (settleTimer) return;

  settleTimer = setInterval(async () => {
    if (!pendingSheetId || !runtimeSecret) return;
    if (Date.now() - lastChangeAt < IDLE_MS) return;

    const sheetId = pendingSheetId;
    pendingSheetId = null;

    clearInterval(settleTimer);
    settleTimer = null;
    if (creationTimer) {
      clearTimeout(creationTimer);
      creationTimer = null;
    }

    try {
      await finalizeNewSheet(sheetId, runtimeSecret);
      createWindowOpen = false;
      setStatus("on", "Hari baru aman. Sheet lama full lock; struktur terkunci.");
      msg("Proteksi otomatis selesai. Karyawan dapat melanjutkan input pada sheet terbaru.");
    } catch (e) {
      createWindowOpen = false;
      try { await relockWorkbookStructure(runtimeSecret); } catch (_) {}
      setStatus("off", "Proteksi otomatis gagal; struktur sudah dicoba dikunci kembali.");
      msg(errorText(e));
    }
  }, 1000);
}

async function finalizeNewSheet(newSheetId, secret) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    const newSheet = sheets.getItem(newSheetId);

    wb.load("protection/protected");
    sheets.load("items/id,items/name,items/protection/protected");
    newSheet.load("id,name,protection/protected");
    await context.sync();

    for (const sheet of sheets.items) {
      if (sheet.id === newSheetId) continue;
      if (sheet.protection.protected) sheet.protection.unprotect(secret);
      prepareFullLock(sheet, secret);
    }

    if (newSheet.protection.protected) newSheet.protection.unprotect(secret);
    prepareWorkingSheet(newSheet, secret);

    if (!wb.protection.protected) wb.protection.protect(secret);
    await context.sync();
  });
}

function prepareWorkingSheet(sheet, secret) {
  const work = sheet.getRange(WORK_AREA);
  work.format.protection.locked = false;
  work.format.protection.formulaHidden = false;

  for (const address of LOCK_RANGES) {
    const r = sheet.getRange(address);
    r.format.protection.locked = true;
    r.format.protection.formulaHidden = true;
  }

  sheet.protection.protect({
    selectionMode: Excel.ProtectionSelectionMode.unlocked,
    allowFormatCells: false,
    allowFormatColumns: false,
    allowFormatRows: false,
    allowInsertColumns: false,
    allowInsertRows: false,
    allowDeleteColumns: false,
    allowDeleteRows: false
  }, secret);
}

function prepareFullLock(sheet, secret) {
  const work = sheet.getRange(WORK_AREA);
  work.format.protection.locked = true;
  work.format.protection.formulaHidden = true;

  sheet.protection.protect({
    selectionMode: Excel.ProtectionSelectionMode.none,
    allowFormatCells: false,
    allowFormatColumns: false,
    allowFormatRows: false,
    allowInsertColumns: false,
    allowInsertRows: false,
    allowDeleteColumns: false,
    allowDeleteRows: false
  }, secret);
}

async function relockWorkbookStructure(secret) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    wb.load("protection/protected");
    await context.sync();
    if (!wb.protection.protected) wb.protection.protect(secret);
    await context.sync();
  });
}

async function lockStructureNow() {
  try {
    await requireAdminPin();
    const secret = requireSecret();
    await relockWorkbookStructure(secret);
    createWindowOpen = false;
    if (creationTimer) clearTimeout(creationTimer);
    creationTimer = null;
    setStatus("on", "Struktur workbook terkunci.");
    msg("Struktur workbook sudah dikunci sekarang.");
  } catch (e) { msg(errorText(e)); }
}

async function unlockActiveSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();
      if (sheet.protection.protected) sheet.protection.unprotect(secret);
      await context.sync();
      msg(`Sheet "${sheet.name}" dibuka untuk admin. Setelah koreksi, kunci kembali.`);
    });
  } catch (e) { msg(errorText(e)); }
}

async function lockActiveAsWorkingSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();
      if (sheet.protection.protected) sheet.protection.unprotect(secret);
      prepareWorkingSheet(sheet, secret);
      await context.sync();
      msg(`Sheet "${sheet.name}" dikunci sebagai SHEET KERJA.`);
    });
  } catch (e) { msg(errorText(e)); }
}

async function fullLockActiveSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();
      if (sheet.protection.protected) sheet.protection.unprotect(secret);
      prepareFullLock(sheet, secret);
      await context.sync();
      msg(`Sheet "${sheet.name}" dikunci PENUH.`);
    });
  } catch (e) { msg(errorText(e)); }
}

async function syncAllSheets() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const wb = context.workbook;
      const sheets = wb.worksheets;
      const active = sheets.getActiveWorksheet();

      wb.load("protection/protected");
      sheets.load("items/id,items/name,items/protection/protected");
      active.load("id,name");
      await context.sync();

      if (wb.protection.protected) wb.protection.unprotect(secret);

      for (const sheet of sheets.items) {
        if (sheet.protection.protected) sheet.protection.unprotect(secret);
        if (sheet.id === active.id) prepareWorkingSheet(sheet, secret);
        else prepareFullLock(sheet, secret);
      }

      wb.protection.protect(secret);
      await context.sync();
      msg(`Sinkron selesai. "${active.name}" = sheet kerja; semua sheet lain = full lock; struktur = lock.`);
    });
  } catch (e) { msg(errorText(e)); }
}

async function requireAdminPin() {
  const pin = byId("adminPin").value;
  if (!pin) throw new Error("Masukkan PIN admin.");

  const savedHash = getSetting(CFG.pinHash);
  const actualHash = await sha256(pin);
  if (!savedHash || actualHash !== savedHash) throw new Error("PIN admin salah.");

  byId("adminPin").value = "";
}

function requireSecret() {
  const secret = runtimeSecret || getSetting(CFG.secret);
  if (!secret) throw new Error("Pengaman V3 belum dikonfigurasi.");
  runtimeSecret = secret;
  return secret;
}

async function hidePane() {
  try { await Office.addin.hide(); }
  catch (_) { msg("Panel boleh ditutup manual. Proteksi yang sudah dipasang tetap berada di workbook."); }
}

function getSetting(key) { return Office.context.document.settings.get(key); }
function setSetting(key, value) { Office.context.document.settings.set(key, value); }

function saveSettings() {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error || new Error("Gagal menyimpan pengaturan workbook."));
    });
  });
}

async function sha256(text) {
  const data = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, "0")).join("");
}

function generateSecret() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return Array.from(bytes).map(b => b.toString(16).padStart(2, "0")).join("");
}

function setConfiguredUi(configured) {
  byId("setupCard").classList.toggle("hidden", configured);
  byId("employeeCard").classList.toggle("hidden", !configured);
  byId("adminCard").classList.toggle("hidden", !configured);
}

function setStatus(mode, text) {
  const dot = byId("statusDot");
  dot.className = `dot ${mode}`;
  byId("statusText").textContent = text;
}

function msg(text) { byId("message").textContent = text || ""; }
function byId(id) { return document.getElementById(id); }

function errorText(e) {
  if (!e) return "Terjadi kesalahan.";
  if (e.debugInfo && e.debugInfo.errorLocation) return `${e.message || e} (${e.debugInfo.errorLocation})`;
  return e.message || String(e);
}
