const CFG = {
  enabled: "hitunganV3.enabled",
  pinHash: "hitunganV3.pinHash",
  secret: "hitunganV3.secret",
  version: "hitunganV3.version"
};

Office.onReady((info) => {
  const status = document.getElementById("status");
  if (info.host !== Office.HostType.Excel) {
    status.textContent = "Buka add-in ini dari Excel.";
    return;
  }
  status.textContent = "Siap. Masukkan PIN admin V3.";
  document.getElementById("unlock").addEventListener("click", recover);
});

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
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, "0")).join("");
}
function errText(e) {
  if (!e) return "Terjadi kesalahan.";
  if (e.debugInfo && e.debugInfo.errorLocation) return `${e.message || e} (${e.debugInfo.errorLocation})`;
  return e.message || String(e);
}

async function recover() {
  const status = document.getElementById("status");
  const btn = document.getElementById("unlock");
  const pin = document.getElementById("pin").value;
  btn.disabled = true;
  try {
    if (!pin) throw new Error("Masukkan PIN admin V3.");
    const secret = getSetting(CFG.secret);
    const savedHash = getSetting(CFG.pinHash);
    if (!secret || !savedHash) throw new Error("Pengaturan V3 tidak ditemukan di workbook ini.");
    const actualHash = await sha256(pin);
    if (actualHash !== savedHash) throw new Error("PIN admin salah.");

    status.textContent = "Membuka proteksi workbook dan semua sheet…";

    await Excel.run(async (context) => {
      const wb = context.workbook;
      const sheets = wb.worksheets;
      wb.load("protection/protected");
      sheets.load("items/name,items/protection/protected");
      await context.sync();

      if (wb.protection.protected) wb.protection.unprotect(secret);
      for (const sheet of sheets.items) {
        if (sheet.protection.protected) sheet.protection.unprotect(secret);
      }
      await context.sync();
    });

    setSetting(CFG.enabled, false);
    setSetting(CFG.version, "recovery-disabled");
    await saveSettings();

    try { await Office.addin.setStartupBehavior(Office.StartupBehavior.none); } catch (_) {}

    document.getElementById("pin").value = "";
    status.textContent = "BERHASIL. Proteksi V3 sudah dibuka dan auto-protection dimatikan.\n\nSekarang: File → Info → Check for Issues → Inspect Document. Hapus hanya 'Task Pane Add-ins'.";
  } catch (e) {
    status.textContent = "Gagal: " + errText(e);
  } finally {
    btn.disabled = false;
  }
}
