/* Pengaman Hitungan V3
 * Companion add-in untuk BUAT HARI BARU V2.
 * ExcelApi 1.7 + SharedRuntime 1.1
 */

const CFG = {
  enabled: "hitunganV3.enabled",
  pinHash: "hitunganV3.pinHash",
  secret: "hitunganV3.secret",
  version: "hitunganV3.version"
};

const WORK_AREA = "A1:AH400";
const LOCK_RANGES = [
  "B267",
  "B269",
  "B271:B275",
  "G272:G275",
  "F178:F238",
  "K178:K238",
  "L178:L238",
  "M178:M238",
  "N178:N238",
  "O178:O238",
  "N163",
  "B6:B75",
  "C6:C75",
  "D6:D75",
  "F6:F75",
  "G6:G75",
  "H6:H75",
  "M23",
  "L50",
  "K75",
  "L75",
  "H76:H79",
  "B114:B116",
  "B145:B147",
  "B163:B165",
  "J168:J170",
  "F171:F173",
  "R249:R251",
  "K145:K147",
  "W143:W145",
  "W157:W159",
  "W129:W131",
  "Y178",
  "Y237"
];

const IDLE_MS = 5000;

let runtimeSecret = null;
let handlersRegistered = false;
let pendingSheetId = null;
let lastChangeAt = 0;
let settleTimer = null;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  bindUi();

  try {
    const enabled = getSetting(CFG.enabled) === true;
    const secret = getSetting(CFG.secret);

    if (enabled && secret) {
      runtimeSecret = secret;
      await registerWorkbookEvents();
      setConfiguredUi(true);
      setStatus("on", "Pengaman V3 aktif. Menunggu sheet baru.");
      msg("V3 dimuat otomatis dari workbook.");
    } else {
      setConfiguredUi(false);
      setStatus("off", "Belum dikonfigurasi.");
    }
  } catch (e) {
    setStatus("off", "Gagal memulai Pengaman V3.");
    msg(errorText(e));
  }
});

function bindUi() {
  byId("setupBtn").addEventListener("click", setupV3);
  byId("unlockBtn").addEventListener("click", unlockActiveSheet);
  byId("partialLockBtn").addEventListener("click", lockActiveAsWorkingSheet);
  byId("fullLockBtn").addEventListener("click", fullLockActiveSheet);
  byId("syncBtn").addEventListener("click", syncAllSheets);
  byId("hideBtn").addEventListener("click", hidePane);
}

async function setupV3() {
  const pin1 = byId("setupPin").value;
  const pin2 = byId("setupPin2").value;
  const oldPassword = byId("oldPassword").value;

  if (!pin1 || pin1.length < 8) {
    msg("PIN admin minimal 8 karakter.");
    return;
  }
  if (pin1 !== pin2) {
    msg("PIN admin dan pengulangannya tidak sama.");
    return;
  }

  try {
    setStatus("wait", "Menyiapkan proteksi…");
    const pinHash = await sha256(pin1);
    const secret = generateSecret();

    await migrateAndSync(secret, oldPassword || null);

    setSetting(CFG.pinHash, pinHash);
    setSetting(CFG.secret, secret);
    setSetting(CFG.enabled, true);
    setSetting(CFG.version, "3.0.0");
    await saveSettings();

    runtimeSecret = secret;
    await registerWorkbookEvents();

    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    } catch (_) {
      // Tetap bisa dipakai walau host tidak mendukung startup behavior.
    }

    byId("setupPin").value = "";
    byId("setupPin2").value = "";
    byId("oldPassword").value = "";

    setConfiguredUi(true);
    setStatus("on", "Pengaman V3 aktif. Menunggu sheet baru.");
    msg("Berhasil. Sheet aktif = sheet kerja. Sheet lain = terkunci penuh.");
  } catch (e) {
    setStatus("off", "Setup gagal.");
    msg(
      "Setup gagal.\n" +
      "Jika ada sheet yang sudah diproteksi, isi 'Password proteksi lama' dengan password yang benar.\n\n" +
      errorText(e)
    );
  }
}

async function migrateAndSync(newSecret, oldPassword) {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();

    sheets.load("items/id,items/name,items/protection/protected");
    active.load("id");
    await context.sync();

    for (const sheet of sheets.items) {
      if (sheet.protection.protected) {
        if (!oldPassword) {
          throw new Error(`Sheet "${sheet.name}" sudah diproteksi. Masukkan password proteksi lama.`);
        }
        sheet.protection.unprotect(oldPassword);
      }
    }
    await context.sync();

    for (const sheet of sheets.items) {
      if (sheet.id === active.id) {
        prepareWorkingSheet(sheet, newSecret);
      } else {
        prepareFullLock(sheet, newSecret);
      }
    }

    await context.sync();
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

function onWorksheetAdded(event) {
  if (!runtimeSecret) return;

  pendingSheetId = event.worksheetId;
  lastChangeAt = Date.now();

  setStatus("wait", "Sheet baru terdeteksi. Menunggu add-in V2 selesai mengisi data…");
  startSettleTimer();
}

function onWorksheetChanged(event) {
  if (!pendingSheetId) return;
  if (event.worksheetId === pendingSheetId) {
    lastChangeAt = Date.now();
  }
}

function startSettleTimer() {
  if (settleTimer) return;

  settleTimer = setInterval(async () => {
    if (!pendingSheetId || !runtimeSecret) return;

    const idleFor = Date.now() - lastChangeAt;
    if (idleFor < IDLE_MS) return;

    const sheetId = pendingSheetId;
    pendingSheetId = null;

    clearInterval(settleTimer);
    settleTimer = null;

    try {
      await finalizeNewSheet(sheetId, runtimeSecret);
      setStatus("on", "Sheet baru aman. Sheet lama terkunci penuh.");
      msg("Proteksi otomatis selesai.");
    } catch (e) {
      setStatus("off", "Proteksi otomatis gagal.");
      msg(errorText(e));
    }
  }, 1000);
}

async function finalizeNewSheet(newSheetId, secret) {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const newSheet = sheets.getItem(newSheetId);

    sheets.load("items/id,items/name,items/protection/protected");
    newSheet.load("id,name,protection/protected");
    await context.sync();

    for (const sheet of sheets.items) {
      if (sheet.id === newSheetId) continue;

      if (sheet.protection.protected) {
        sheet.protection.unprotect(secret);
      }
      prepareFullLock(sheet, secret);
    }

    if (newSheet.protection.protected) {
      newSheet.protection.unprotect(secret);
    }
    prepareWorkingSheet(newSheet, secret);

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

  sheet.protection.protect(
    {
      selectionMode: "Unlocked",
      allowFormatCells: false,
      allowFormatColumns: false,
      allowFormatRows: false,
      allowInsertColumns: false,
      allowInsertRows: false,
      allowDeleteColumns: false,
      allowDeleteRows: false
    },
    secret
  );
}

function prepareFullLock(sheet, secret) {
  const work = sheet.getRange(WORK_AREA);
  work.format.protection.locked = true;
  work.format.protection.formulaHidden = true;

  sheet.protection.protect(
    {
      selectionMode: "None",
      allowFormatCells: false,
      allowFormatColumns: false,
      allowFormatRows: false,
      allowInsertColumns: false,
      allowInsertRows: false,
      allowDeleteColumns: false,
      allowDeleteRows: false
    },
    secret
  );
}

async function unlockActiveSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect(secret);
        await context.sync();
      }

      msg(`Sheet "${sheet.name}" sudah dibuka untuk admin.`);
    });
  } catch (e) {
    msg(errorText(e));
  }
}

async function lockActiveAsWorkingSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect(secret);
      }
      prepareWorkingSheet(sheet, secret);
      await context.sync();

      msg(`Sheet "${sheet.name}" dikunci sebagai SHEET KERJA.`);
    });
  } catch (e) {
    msg(errorText(e));
  }
}

async function fullLockActiveSheet() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect(secret);
      }
      prepareFullLock(sheet, secret);
      await context.sync();

      msg(`Sheet "${sheet.name}" dikunci PENUH.`);
    });
  } catch (e) {
    msg(errorText(e));
  }
}

async function syncAllSheets() {
  try {
    await requireAdminPin();
    const secret = requireSecret();

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const active = sheets.getActiveWorksheet();

      sheets.load("items/id,items/name,items/protection/protected");
      active.load("id,name");
      await context.sync();

      for (const sheet of sheets.items) {
        if (sheet.protection.protected) {
          sheet.protection.unprotect(secret);
        }

        if (sheet.id === active.id) {
          prepareWorkingSheet(sheet, secret);
        } else {
          prepareFullLock(sheet, secret);
        }
      }

      await context.sync();
      msg(`Sinkron selesai. "${active.name}" = sheet kerja, semua sheet lain = full lock.`);
    });
  } catch (e) {
    msg(errorText(e));
  }
}

async function requireAdminPin() {
  const pin = byId("adminPin").value;
  if (!pin) throw new Error("Masukkan PIN admin.");

  const savedHash = getSetting(CFG.pinHash);
  const actualHash = await sha256(pin);

  if (!savedHash || actualHash !== savedHash) {
    throw new Error("PIN admin salah.");
  }
  byId("adminPin").value = "";
}

function requireSecret() {
  const s = runtimeSecret || getSetting(CFG.secret);
  if (!s) throw new Error("Pengaman V3 belum dikonfigurasi.");
  runtimeSecret = s;
  return s;
}

async function hidePane() {
  try {
    await Office.addin.hide();
  } catch (e) {
    msg("Panel tidak dapat disembunyikan di host ini. Anda boleh menutupnya manual.");
  }
}

function getSetting(key) {
  return Office.context.document.settings.get(key);
}

function setSetting(key, value) {
  Office.context.document.settings.set(key, value);
}

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
  return Array.from(new Uint8Array(digest))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

function generateSecret() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return Array.from(bytes).map(b => b.toString(16).padStart(2, "0")).join("");
}

function setConfiguredUi(configured) {
  byId("setupCard").classList.toggle("hidden", configured);
  byId("adminCard").classList.toggle("hidden", !configured);
}

function setStatus(mode, text) {
  const dot = byId("statusDot");
  dot.className = `dot ${mode}`;
  byId("statusText").textContent = text;
}

function msg(text) {
  byId("message").textContent = text || "";
}

function byId(id) {
  return document.getElementById(id);
}

function errorText(e) {
  if (!e) return "Terjadi kesalahan.";
  if (e.debugInfo && e.debugInfo.errorLocation) {
    return `${e.message || e} (${e.debugInfo.errorLocation})`;
  }
  return e.message || String(e);
}
