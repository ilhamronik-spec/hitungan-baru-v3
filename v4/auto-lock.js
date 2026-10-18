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
const GLOBAL_UNLOCK_WINDOW_MS = 90 * 60 * 1000;
const GLOBAL_UNLOCK_UNTIL_KEY = "hitunganV4.globalUnlockUntil";
let globalUnlockRelockTimer = null;
const SHIFT_SHEET_ID_KEY = "hitunganV4.shiftWorkSheetId";
const SHIFT_SHEET_NAME_KEY = "hitunganV4.shiftWorkSheetName";
const SHIFT_CREATED_AT_KEY = "hitunganV4.shiftCreatedAt";
const SHIFT_UNTIL_KEY = "hitunganV4.shiftWorkUntil";

let adminSheetRelockTimer = null;
let shiftWorkRelockTimer = null;
let shiftCreateWrapped = false;
let shiftWrapTimer = null;
let shiftUiTimer = null;
// Shared across existing V4 modules: delay nonessential background writes
// until all sheet creation / display-transfer / price checks have finished.
let hbShiftPipelineBusy = false;
let shiftAutoLockBusy = false;
let shiftAutoLockRetryTimer = null;
const SHIFT_AUTOLOCK_RETRY_MS = 20000;

function persistentAdminUnlockActive() {
  if (typeof isGlobalUnlockMode === "function") return isGlobalUnlockMode();
  if (typeof getSetting !== "function") return false;
  const value = getSetting("hitunganV4.globalUnlockAll");
  return value === true || String(value).toLowerCase() === "true";
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  bindAdminUnlockSafety();
  ensureEmployeeShiftUi();
  waitAndWrapShiftCreate();

  restoreGlobalUnlockWindow().catch(() => {});
  restoreAdminSheetAutoLock().catch(() => {});
  restoreShiftWorkWindow().catch((e) => shiftSecurityError("Pemulihan timer shift", e));

  if (shiftUiTimer) clearInterval(shiftUiTimer);
  shiftUiTimer = setInterval(() => {
    auditShiftAutoLock().catch(e => shiftSecurityError("Pemeriksaan Full Lock",e));
  }, 30000);
  setTimeout(() => auditShiftAutoLock().catch(e => shiftSecurityError("Pemeriksaan awal",e)), 1100);
});

function shiftSecurityError(label,e){
  const reason=e&&e.message?e.message:String(e||"Kesalahan tidak diketahui");
  if(typeof setSecurityStatus==="function")
    setSecurityStatus(label+" belum selesai: "+reason+". Sistem akan mencoba lagi; jangan anggap sheet sudah FULL LOCK.",true);
}
function scheduleShiftAutoLockRetry(delay=SHIFT_AUTOLOCK_RETRY_MS){
  if(shiftAutoLockRetryTimer)clearTimeout(shiftAutoLockRetryTimer);
  shiftAutoLockRetryTimer=setTimeout(()=>{
    shiftAutoLockRetryTimer=null;
    auditShiftAutoLock().catch(e=>shiftSecurityError("Percobaan penguncian ulang",e));
  },delay);
}
async function auditShiftAutoLock(){
  if(typeof getSetting!=="function"||typeof requireSecret!=="function")return;
  const id=getSetting(SHIFT_SHEET_ID_KEY);
  const until=Number(getSetting(SHIFT_UNTIL_KEY));
  // Proactively recover from an expired timer, including after browser sleep.
  if(id && Number.isFinite(until)&&until>0&&Date.now()>=until){
    await expireShiftWorkWindow();
  }else if(id && Number.isFinite(until)&&until>0){
    scheduleShiftWorkRelock(until);
  }else if(id){
    await restoreShiftWorkWindow();
  }else if(!persistentAdminUnlockActive() && !hbShiftPipelineBusy){
    await failCloseUntrackedLatestShift();
  }
  if(persistentAdminUnlockActive() && typeof restoreGlobalUnlockWindow==="function"){
    await restoreGlobalUnlockWindow();
  }
  if(typeof refreshShiftUi==="function")await refreshShiftUi();
}
// A legacy sheet may have been created before the timer was registered.
// It must never remain writable indefinitely simply because the timer key
// is absent. Only a work-mode sheet (A1 unlocked) is treated as orphaned.
let shiftTimerRecoveryBusy=false;
async function refreshShiftSettingsFromWorkbook(){
  const settings=Office.context.document.settings;
  if(typeof settings.refreshAsync!=="function")return false;
  return new Promise((resolve,reject)=>{
    settings.refreshAsync(result=>{
      if(result && result.status===Office.AsyncResultStatus.Succeeded)resolve(true);
      else reject(result&&result.error||new Error("Pengaturan waktu shift belum tersinkron antarperangkat."));
    });
  });
}
// Never hard-lock a newly cloned shift merely because this device has not
// received the timestamp saved by the device that created it.
// Re-read workbook settings and recover the timer while keeping sensitive
// cells protected by the existing workingLock.
async function failCloseUntrackedLatestShift(){
  if(shiftTimerRecoveryBusy||hbShiftPipelineBusy||
     (typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress))return;
  const adminUntil=Number(getSetting(ADMIN_UNLOCK_UNTIL_KEY));
  if(Number.isFinite(adminUntil)&&adminUntil>Date.now())return;
  shiftTimerRecoveryBusy=true;
  try{
    await refreshShiftSettingsFromWorkbook();
    const freshId=getSetting(SHIFT_SHEET_ID_KEY);
    const freshUntil=Number(getSetting(SHIFT_UNTIL_KEY));
    if(freshId){
      if(Number.isFinite(freshUntil)&&freshUntil>Date.now())
        scheduleShiftWorkRelock(freshUntil);
      // An expired existing timestamp is handled by expireShiftWorkWindow.
      // Do NOT replace its original deadline.
      return;
    }
    const info=await Excel.run(async ctx=>{
      const collection=ctx.workbook.worksheets;
      collection.load("items/id,items/name,items/position");
      await ctx.sync();
      if(!collection.items.length)return null;
      const latest=collection.items.reduce((a,b)=>
        Number(b.position)<Number(a.position)?b:a);
      const sh=collection.getItem(latest.id);
      sh.load("id,name,protection/protected");
      const marker=sh.getRange("A1");
      marker.format.protection.load("locked");
      await ctx.sync();
      return {id:sh.id,name:sh.name,protected:sh.protection.protected,
        working:marker.format.protection.locked===false};
    });
    if(!info || !info.protected || !info.working)return;
    // If the identity/timestamp survived while the ID did not, reuse its
    // original clock. Otherwise create a recovery deadline and warn Admin.
    const savedName=String(getSetting(SHIFT_SHEET_NAME_KEY)||"");
    const originalStarted=Number(getSetting(SHIFT_CREATED_AT_KEY));
    const hasOriginal=savedName===info.name &&
      Number.isFinite(originalStarted)&&originalStarted>0 &&
      originalStarted<=Date.now();
    const createdAt=hasOriginal?originalStarted:Date.now();
    const until=createdAt+SHIFT_WORK_WINDOW_MS;
    setSetting(SHIFT_SHEET_ID_KEY,info.id);
    setSetting(SHIFT_SHEET_NAME_KEY,info.name);
    setSetting(SHIFT_CREATED_AT_KEY,createdAt);
    setSetting(SHIFT_UNTIL_KEY,until);
    await saveSettings();
    if(until<=Date.now()){
      await expireShiftWorkWindow();
    }else{
      scheduleShiftWorkRelock(until);
      setSecurityStatus("TIMER SHIFT DIPULIHKAN untuk \""+info.name+
        "\". "+(hasOriginal?"Waktu awal asli tetap dipakai.":
        "Jam mulai asli tidak tersedia di perangkat ini; batas 2,5 jam dihitung sejak pemulihan.")+
        " Sheet tidak ditutup prematur dan sel penting tetap terlindungi.",true);
    }
  }finally{
    shiftTimerRecoveryBusy=false;
  }
}

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
        if(hbShiftPipelineBusy)
          throw new Error("Pembuatan shift lain masih diproses. Tunggu hingga seluruh stok dan harga selesai disimpan.");
        hbShiftPipelineBusy=true;
        try{
        // Take a snapshot before any inner wrapper can clear settings.
        const scheduledPrices=(typeof getNextShiftPriceChanges==="function")
          ? getNextShiftPriceChanges().map(item=>({...item}))
          : [];
        const result=await previousCreate(person,sheetName,purchaseData);
        if(!getSetting(SHIFT_SHEET_ID_KEY) || String(getSetting(SHIFT_SHEET_NAME_KEY))!==String(sheetName))
          await registerNewShiftWorkWindow(sheetName);
        if(scheduledPrices.length && typeof verifyAndFinalizeNextShiftPricing==="function"){
          try{
            const outcome=await verifyAndFinalizeNextShiftPricing(sheetName,scheduledPrices);
            if(outcome.skipped.length){
              setSecurityStatus("SHEET BARU BERHASIL dibuat. "+outcome.verified+
                " harga otomatis terverifikasi, "+outcome.skipped.length+
                " harga BELUM berlaku karena nama/baris tidak cocok: "+
                outcome.skipped.join("; ")+". Periksa Menu Admin.",true);
            }else{
              setSecurityStatus("SHEET BARU BERHASIL dibuat. "+outcome.verified+
                " harga jual shift baru diterapkan dan DIVERIFIKASI otomatis"+
                (outcome.repaired?" ("+outcome.repaired+" dikoreksi setelah pembuatan).":".")+
                " Tidak perlu klik tombol pemulihan manual.",false);
            }
          }catch(e){
            // The worksheet already exists; avoid encouraging employees to
            // press Create again and accidentally create duplicate shifts.
            setSecurityStatus("SHEET BARU sudah dibuat, tetapi penerapan harga otomatis BELUM berhasil: "+
              (e&&e.message||String(e))+
              ". Harga pending tetap disimpan. Jangan buat sheet baru lagi; gunakan Menu Admin untuk perbaikan.",true);
          }
        }
        return result;
        }finally{
          hbShiftPipelineBusy=false;
        }
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

  const existingId=getSetting(SHIFT_SHEET_ID_KEY);
  const existingUntil=Number(getSetting(SHIFT_UNTIL_KEY));
  if(String(existingId)===String(info.id) && existingUntil>Date.now())return;
  const createdAt = Date.now();
  const until = createdAt + SHIFT_WORK_WINDOW_MS;

  setSetting(SHIFT_SHEET_ID_KEY, info.id);
  setSetting(SHIFT_SHEET_NAME_KEY, info.name);
  setSetting(SHIFT_CREATED_AT_KEY, createdAt);
  setSetting(SHIFT_UNTIL_KEY, until);
  await saveSettings();

  scheduleShiftWorkRelock(until);
  await refreshShiftUi();
  setSecurityStatus(`SHIFT AKTIF — sheet "${info.name}" dapat diisi maksimal 2,5 jam. Setelah itu otomatis full lock.`, false);
}

async function restoreShiftWorkWindow() {
  const sheetId = getSetting(SHIFT_SHEET_ID_KEY);
  const untilRaw = getSetting(SHIFT_UNTIL_KEY);

  if (!sheetId || !untilRaw) {
    await failCloseUntrackedLatestShift();
    await refreshShiftUi();
    return;
  }

  const until = Number(untilRaw);
  if (!Number.isFinite(until)||until<=0) {
    // Invalid/missing time is NOT proof that 2.5 hours have passed.
    await refreshShiftSettingsFromWorkbook();
    const restoredUntil=Number(getSetting(SHIFT_UNTIL_KEY));
    if(Number.isFinite(restoredUntil)&&restoredUntil>0){
      if(restoredUntil>Date.now())scheduleShiftWorkRelock(restoredUntil);
      else await expireShiftWorkWindow();
    }else{
      setSecurityStatus("Timer sheet belum valid. Jangan anggap shift sudah habis; pemeriksaan waktu akan dicoba lagi.",true);
    }
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
    auditShiftAutoLock().catch(e=>{
      shiftSecurityError("Timer 2,5 jam",e);
      scheduleShiftAutoLockRetry();
    });
  }, delay);
}

async function expireShiftWorkWindow() {
  if(shiftAutoLockBusy)return;
  const sheetId=getSetting(SHIFT_SHEET_ID_KEY);
  if(!sheetId)return;
  const until=Number(getSetting(SHIFT_UNTIL_KEY));
  if(Number.isFinite(until)&&until>Date.now()){
    scheduleShiftWorkRelock(until);
    return;
  }
  // An admin's temporary unlock must not prevent the shift deadline. Retry
  // as soon as the 90-minute admin window is over, without clearing the shift.
  if(persistentAdminUnlockActive()){
    scheduleShiftAutoLockRetry();
    return;
  }
  shiftAutoLockBusy=true;
  const sheetName=getSetting(SHIFT_SHEET_NAME_KEY)||"sheet";
  try{
    const pw=requireSecret();
    const result=await Excel.run(async ctx=>{
      const wb=ctx.workbook,sh=wb.worksheets.getItem(sheetId);
      sh.load("name,protection/protected");
      wb.protection.load("protected");
      const proof=sh.getRange("A1");
      proof.format.protection.load("locked");
      await ctx.sync();
      // Protected alone does not mean full lock. WorkingLock also protects.
      if(!sh.protection.protected || proof.format.protection.locked!==true){
        if(sh.protection.protected){
          sh.protection.unprotect(pw);
          await ctx.sync();
        }
        fullLock(sh,pw);
      }
      if(!wb.protection.protected)wb.protection.protect(pw);
      await ctx.sync();
      sh.protection.load("protected");
      proof.format.protection.load("locked");
      await ctx.sync();
      if(!sh.protection.protected || proof.format.protection.locked!==true)
        throw new Error("Excel belum mengonfirmasi FULL LOCK; penguncian akan diulang.");
      return sh.name;
    });
    await clearShiftWorkWindowState();
    if(shiftAutoLockRetryTimer)clearTimeout(shiftAutoLockRetryTimer);
    shiftAutoLockRetryTimer=null;
    await refreshShiftUi();
    setSecurityStatus('FULL LOCK TERVERIFIKASI — sheet "'+result+'" sudah ditutup setelah batas 2,5 jam.',false);
  }catch(e){
    shiftSecurityError('AUTO-LOCK 2,5 jam "'+sheetName+'"',e);
    scheduleShiftAutoLockRetry();
  }finally{
    shiftAutoLockBusy=false;
  }
}

async function employeeCloseShiftNow() {
  try {
    if (persistentAdminUnlockActive()) throw new Error("Mode ADMIN BEBAS sedang aktif. Sheet hanya dapat dikunci kembali melalui tombol KUNCI SEMUA KEMBALI.");
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

  if (persistentAdminUnlockActive()) {
    el.textContent = "ADMIN BEBAS aktif — semua sheet tetap terbuka sampai admin mengunci kembali.";
    btn.disabled = true;
    return;
  }

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

async function clearShiftWindowIfItsSheetIsFullLocked(){
  const id=getSetting(SHIFT_SHEET_ID_KEY);
  if(!id)return;
  try{
    const full=await Excel.run(async ctx=>{
      const sh=ctx.workbook.worksheets.getItem(id);
      sh.load("protection/protected");
      const marker=sh.getRange("A1");
      marker.format.protection.load("locked");
      await ctx.sync();
      return sh.protection.protected&&marker.format.protection.locked===true;
    });
    // Only manually confirmed FULL LOCK can cancel an unexpired timer.
    if(full){
      await clearShiftWorkWindowState();
      await refreshShiftUi();
    }
  }catch(e){shiftSecurityError("Periksa Full Lock manual",e);}
}

// ---------------- Admin temporary unlock safety (15 minutes) ----------------

async function captureAdminUnlockedSheet() {
  try {
    if (persistentAdminUnlockActive()) return;
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
  if (persistentAdminUnlockActive()) {
    await clearAdminSheetAutoLockState();
    return;
  }
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
  if (persistentAdminUnlockActive()) return;
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


// Global unlock is always temporary (90 minutes), persisted with workbook settings.
async function beginGlobalUnlockWindow() {
  const until = Date.now() + GLOBAL_UNLOCK_WINDOW_MS;
  setSetting(GLOBAL_UNLOCK_UNTIL_KEY, until);
  await saveSettings();
  scheduleGlobalUnlockRelock(until);
  return until;
}
function scheduleGlobalUnlockRelock(until) {
  if (globalUnlockRelockTimer) clearTimeout(globalUnlockRelockTimer);
  globalUnlockRelockTimer = setTimeout(() => {
    expireGlobalUnlockWindow().catch(e => setSecurityStatus(errText(e), true));
  }, Math.max(0, Number(until) - Date.now()));
}
async function restoreGlobalUnlockWindow() {
  if (!persistentAdminUnlockActive()) return clearGlobalUnlockWindowState();
  let until = Number(getSetting(GLOBAL_UNLOCK_UNTIL_KEY));
  if (!Number.isFinite(until) || until <= 0) {
    // Migration from unlimited mode: give existing unlocked workbooks a single
    // 90-minute safety window starting the first time this version opens.
    until = await beginGlobalUnlockWindow();
  }
  if (Date.now() >= until) await expireGlobalUnlockWindow();
  else scheduleGlobalUnlockRelock(until);
}
async function expireGlobalUnlockWindow() {
  if (!persistentAdminUnlockActive()) return clearGlobalUnlockWindowState();
  const until = Number(getSetting(GLOBAL_UNLOCK_UNTIL_KEY));
  if (Number.isFinite(until) && until > Date.now()) return scheduleGlobalUnlockRelock(until);
  if (typeof lockAllSheetsCore !== "function") throw new Error("Mesin penguncian belum siap.");
  await lockAllSheetsCore("WAKTU ADMIN HABIS — 1,5 jam berlalu. Sheet kerja yang masih dalam 2,5 jam tetap bisa diisi.");
}
async function clearGlobalUnlockWindowState() {
  if (globalUnlockRelockTimer) clearTimeout(globalUnlockRelockTimer);
  globalUnlockRelockTimer = null;
  if (typeof getSetting !== "function") return;
  if (getSetting(GLOBAL_UNLOCK_UNTIL_KEY) == null) return;
  Office.context.document.settings.remove(GLOBAL_UNLOCK_UNTIL_KEY);
  await saveSettings();
}
