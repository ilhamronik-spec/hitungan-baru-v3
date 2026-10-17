"use strict";

const V2_SCRIPT = "https://hitungan-baru-excel-ilham.bienbachthuatz2.chatgpt.site/taskpane.js";
const CFG = { enabled:"hitunganV4.enabled", pinHash:"hitunganV4.pinHash", secret:"hitunganV4.secret", version:"hitunganV4.version", globalUnlock:"hitunganV4.globalUnlockAll" };
const LEGACY_V3_SECRET_KEY = "hitunganV3.secret";
const WORK_AREA = "A1:AH400";
const LOCK_RANGES = [
  "B267","B269","B271:B275","G272:G275","B178:B238","C178:C238","F178:F238","K178:K238","L178:L238","M178:M238","N178:N238","O178:O238",
  "N163","B6:B75","C6:C75","D6:D75","F6:F75","G6:G75","H6:H75","M23","L50","K75","L75","H76:H79","B114:B116",
  "B145:B147","B163:B165","J168:J170","F171:F173","R249:R251","K145:K147","W143:W145","W157:W159","W129:W131","Y178","Y237"
];
let secret = null;
let relockTimer = null;
let v2Loaded = false;
let shiftCreationInProgress = false;
let adminSessionAuthenticated = false;
let adminGlobalUnmaskBusy = false;

loadV2();

function loadV2(){
  const s=document.createElement("script");
  s.src=V2_SCRIPT;
  s.onload=()=>{ v2Loaded=true; installWrappers(); startSecurityUi(); };
  s.onerror=()=>setSecurityStatus("Gagal memuat mesin Hitungan Baru V2 dari hosting lama.",true);
  document.head.appendChild(s);
}

function installWrappers(){
  if(typeof createNewDaySheet!=="function") return setSecurityStatus("Fungsi V2 tidak ditemukan.",true);
  const originalCreate=createNewDaySheet;
  createNewDaySheet=async function(person,sheetName,purchaseData){
    if(shiftCreationInProgress)throw new Error("Pembuatan sheet lain masih berlangsung. Tunggu sampai selesai.");
    shiftCreationInProgress=true;
    try{
    const pw=requireSecret();
    let sourceName="";
    await Excel.run(async c=>{
      const wb=c.workbook, src=wb.worksheets.getActiveWorksheet();
      wb.protection.load("protected"); src.load("name,protection/protected");
      await c.sync(); sourceName=src.name;
      if(wb.protection.protected) wb.protection.unprotect(pw);
      if(src.protection.protected) src.protection.unprotect(pw);
      await c.sync();
    });
    try{
      const result=await originalCreate(person,sheetName,purchaseData);
      let pricingResult={applied:[],skipped:[],alerts:[]};
      await Excel.run(async c=>{
        const wb=c.workbook, sheets=wb.worksheets, newest=sheets.getActiveWorksheet(), old=sheets.getItem(sourceName);
        wb.protection.load("protected"); newest.load("name,protection/protected"); old.load("protection/protected");
        await c.sync();
        if(newest.name!==sheetName)throw new Error("Sheet baru yang aktif bukan '"+sheetName+"'; harga pending tidak diterapkan untuk menghindari salah sheet.");
        if(newest.protection.protected) newest.protection.unprotect(pw);
        if(old.protection.protected) old.protection.unprotect(pw);
        if(typeof applyNextShiftPricing==="function"){
          pricingResult=await applyNextShiftPricing(c,newest,sourceName,purchaseData);
        }
        // The new sheet may be cloned from a full-locked sheet with ";;;".
        // Restore readable number text and unmask only the current work shift.
        await restoreMinesReadableNumbers(c,newest,true);
        newest.calculate(true);
        if(isGlobalUnlockMode()){
          // Mode admin persisten: sheet lama dan baru harus tetap terbuka.
          if(!wb.protection.protected) wb.protection.protect(pw);
        }else{
          fullLock(old,pw); workingLock(newest,pw);
          if(!wb.protection.protected) wb.protection.protect(pw);
        }
        await c.sync();
      });
      // Register the shift in the primary creation pipeline, rather than relying
      // on a later optional wrapper that might not be installed yet.
      if(typeof registerNewShiftWorkWindow==="function"){
        await registerNewShiftWorkWindow(sheetName);
      }
      // Visual formatting must never turn successful sheet creation into
      // "creation failed". Apply afterwards and offer repair on any error.
      let marginColorWarning="";
      if(typeof hbMarginInstallNewSheet==="function"){
        try { await hbMarginInstallNewSheet(sheetName); }
        catch(e) {
          marginColorWarning=errText(e);
          if(typeof hbMarginQueueRetry==="function")hbMarginQueueRetry(12000);
        }
      }
      // If the outer shift wrapper is installed, keep pending entries intact
      // until the LAST worksheet readback verifies prices after all wrappers.
      // This avoids deleting schedules based on an intermediate Excel state.
      const finalPriceCheckPending=
        typeof shiftCreateWrapped!=="undefined"&&shiftCreateWrapped&&
        typeof verifyAndFinalizeNextShiftPricing==="function";
      if(pricingResult.applied.length && !finalPriceCheckPending &&
         typeof consumeEffectivePrices==="function"){
        await consumeEffectivePrices(pricingResult.applied,sheetName);
      }
      if(!finalPriceCheckPending && !pricingResult.applied.length &&
         typeof getNextShiftPriceChanges==="function" &&
         getNextShiftPriceChanges().length && !pricingResult.skipped.length){
        setSecurityStatus("PERINGATAN: "+getNextShiftPriceChanges().length+
          " harga masih pending sesudah sheet '"+sheetName+
          "' dibuat. Periksa status di Admin > Kelola Harga Jual.",true);
      }else if(marginColorWarning){
        setSecurityStatus("Sheet baru berhasil dibuat, tetapi warna margin belum diterapkan: "+marginColorWarning+". Periksa ulang melalui Menu Admin.",true);
      }else if(pricingResult.skipped.length){
        setSecurityStatus("PERINGATAN HARGA: "+pricingResult.skipped.join(", ")+" belum diterapkan; harga pending tetap tersimpan.",true);
      }else if(pricingResult.alerts.length){
        setSecurityStatus("PERINGATAN MARGIN < Rp1.000: "+pricingResult.alerts.join("; "),true);
      }else setSecurityStatus(
        isGlobalUnlockMode()
          ? "ADMIN BEBAS — hari baru selesai dan semua sheet tetap terbuka sampai dikunci kembali."
          : "AMAN — hari baru selesai, sheet lama full lock, struktur terkunci.",
        false
      );
      return result;
    }catch(err){
      await failClosedRecovery(sourceName,sheetName,pw);
      throw err;
    }
    }finally{
      shiftCreationInProgress=false;
    }
  };

  if(typeof repairActiveStorageFormulas==="function"){
    const originalRepair=repairActiveStorageFormulas;
    repairActiveStorageFormulas=async function(){
      const pw=requireSecret();
      await Excel.run(async c=>{ const sh=c.workbook.worksheets.getActiveWorksheet(); sh.load("protection/protected"); await c.sync(); if(sh.protection.protected) sh.protection.unprotect(pw); await c.sync(); });
      try{ return await originalRepair(); }
      finally{ await Excel.run(async c=>{ const sh=c.workbook.worksheets.getActiveWorksheet(); sh.load("protection/protected"); await c.sync(); if(sh.protection.protected) sh.protection.unprotect(pw); if(!isGlobalUnlockMode()) workingLock(sh,pw); await c.sync(); }); }
    };
  }
}

async function failClosedRecovery(sourceName,newName,pw){
  try{
    await Excel.run(async c=>{
      const wb=c.workbook,sheets=wb.worksheets; wb.protection.load("protected"); sheets.load("items/name,items/protection/protected"); await c.sync();
      for(const sh of sheets.items){
        if(sh.protection.protected){ try{ sh.protection.unprotect(pw); }catch(_){} }
        if(!isGlobalUnlockMode()){
          if(sh.name===sourceName) workingLock(sh,pw);
          else if(sh.name===newName) fullLock(sh,pw);
        }
      }
      if(!wb.protection.protected) wb.protection.protect(pw);
      await c.sync();
    });
  }catch(_){ }
}

function startSecurityUi(){
  Office.onReady(async info=>{
    if(info.host!==Office.HostType.Excel) return;
    by("setupSecurityBtn").addEventListener("click",setupSecurity);
    by("adminUnlockAllSheetsBtn").addEventListener("click",adminUnlockAllSheets);
    by("adminLockAllSheetsBtn").addEventListener("click",adminLockAllSheets);
    by("adminUnlockSheetBtn").addEventListener("click",adminUnlockSheet);
    by("adminWorkingLockBtn").addEventListener("click",adminWorkingLock);
    by("adminFullLockBtn").addEventListener("click",adminFullLock);
    by("adminUnlockStructureBtn").addEventListener("click",adminUnlockStructure);
    by("adminLockStructureBtn").addEventListener("click",adminLockStructure);
    by("adminSyncBtn").addEventListener("click",adminSync);
    by("adminViewMinusBtn").addEventListener("click",adminViewMinus);
    by("adminHideMinusBtn").addEventListener("click",hideMinusPanel);
    setupRoleMenus();
    const enabled=getSetting(CFG.enabled)===true, saved=getSetting(CFG.secret);
    if(enabled&&saved){
      secret=saved;
      // Ask Office to reopen this task pane with this workbook when the
      // existing sideloaded add-in remains registered in the browser.
      // This does NOT sideload missing manifests and NEVER unprotects a sheet.
      configureDocumentAutoShow().catch(()=>{});
      setConfigured(true);
      setFormEnabled(true);
      try{
        if(isGlobalUnlockMode()){
          await restoreGlobalUnlockWindow();
          if(isGlobalUnlockMode()) await ensureAllSheetsUnlocked();
          await enforceStructure();
          setGlobalUnlockUi(isGlobalUnlockMode());
          setSecurityStatus(isGlobalUnlockMode()?"ADMIN BEBAS — sheet terbuka maksimal 1,5 jam.":"AMAN — batas buka semua sheet selesai; proteksi kembali aktif.",false);
        }else{
          setGlobalUnlockUi(false);
          await enforceStructure();
          setSecurityStatus("AMAN — struktur workbook terkunci. Pemeriksaan mines tersedia pada panel Penutup Angka Mines.",false);
        }
      }catch(e){setSecurityStatus(errText(e),true);}
    }
    else{ setConfigured(false); setFormEnabled(false); setGlobalUnlockUi(false); setSecurityStatus("Belum diaktifkan. Buat PIN admin terlebih dahulu.",true); }
  });
}

async function configureDocumentAutoShow(){
  const key="Office.AutoShowTaskpaneWithDocument";
  if(getSetting(key)===true)return;
  setSetting(key,true);
  await saveSettings();
}

async function setupSecurity(){
  const p1=by("setupPin").value,p2=by("setupPin2").value,old=by("oldPassword").value;
  if(!p1||p1.length<8)return setSecurityStatus("PIN admin minimal 8 karakter.",true);
  if(p1!==p2)return setSecurityStatus("PIN admin dan ulangannya tidak sama.",true);
  const legacy=getSetting(LEGACY_V3_SECRET_KEY), oldPw=legacy||old||null, newPw=randomSecret();
  try{
    await Excel.run(async c=>{
      const wb=c.workbook,sheets=wb.worksheets,active=sheets.getActiveWorksheet();
      wb.protection.load("protected"); sheets.load("items/id,items/name,items/protection/protected"); active.load("id"); await c.sync();
      if(wb.protection.protected){ if(!oldPw)throw new Error("Workbook sudah diproteksi. Isi password proteksi lama."); wb.protection.unprotect(oldPw); }
      for(const sh of sheets.items){ if(sh.protection.protected){ if(!oldPw)throw new Error(`Sheet "${sh.name}" sudah diproteksi. Isi password proteksi lama.`); sh.protection.unprotect(oldPw); } }
      await c.sync();
      for(const sh of sheets.items){ if(sh.id===active.id)workingLock(sh,newPw); else fullLock(sh,newPw); }
      wb.protection.protect(newPw); await c.sync();
    });
    setSetting(CFG.pinHash,await sha256(p1)); setSetting(CFG.secret,newPw); setSetting(CFG.enabled,true); setSetting(CFG.globalUnlock,false); setSetting(CFG.version,"4.3.0"); await saveSettings(); secret=newPw;
    by("setupPin").value="";by("setupPin2").value="";by("oldPassword").value="";setConfigured(true);setFormEnabled(true);setSecurityStatus("AMAN — V4 aktif. Sheet aktif=kerja, sheet lama=full lock, struktur=lock.",false);
  }catch(e){setSecurityStatus(errText(e),true);}
}

function workingLock(sh,pw){
  const w=sh.getRange(WORK_AREA); w.format.protection.locked=false; w.format.protection.formulaHidden=false;
  for(const a of LOCK_RANGES){const r=sh.getRange(a);r.format.protection.locked=true;r.format.protection.formulaHidden=true;}
  // Do not conceal mines from external Excel integrations while a shift is active.
  for(const address of MINES_CELLS){sh.getRange(address).format.protection.formulaHidden=false;}
  sh.protection.protect({selectionMode:Excel.ProtectionSelectionMode.unlocked,allowFormatCells:false,allowFormatColumns:false,allowFormatRows:false,allowInsertColumns:false,allowInsertRows:false,allowDeleteColumns:false,allowDeleteRows:false},pw);
}
const MINES_CELLS = ["D275","D276","G275","J274"];
// Format ";;;" was incorrectly installed by previous builds: it hides the
// formatted TEXT read by external apps such as SmartSDM, not the raw number.
// This readable fallback is applied ONLY to cells already damaged by ";;;".
const MINES_RECOVERY_NUMBER_FORMAT = '"Rp" #,##0;-"Rp" #,##0;"Rp" 0';
function minesHasInvisibleNumberFormat(format){
  return typeof format==="string" && format.replace(/\s/g,"")===";;;";
}
// Only the authenticated "BUKA SEMUA SHEET" flow calls this.
// The worksheet's original numeric values/formulas must never be replaced.
function revealMinusCellsForAdmin(sh){
  for(const address of MINES_CELLS){
    const r=sh.getRange(address);
    r.format.fill.color="#FFFFFF";
    r.format.font.color="#000000";
    r.format.protection.formulaHidden=false;
  }
}
async function revealAllMinusCellsForAdmin(ctx,sheets){
  const cells=[];
  for(const sh of sheets){
    for(const address of MINES_CELLS){
      const cell=sh.getRange(address);
      cell.load("numberFormat");
      cells.push(cell);
    }
  }
  await ctx.sync();
  // Old V4 installations may have added ";;;", which also hides TEXT;
  // fix only those formats so SmartSDM can read them as well.
  for(const cell of cells){
    if(minesHasInvisibleNumberFormat(cell.numberFormat[0][0])){
      cell.numberFormat=[[MINES_RECOVERY_NUMBER_FORMAT]];
    }
  }
  for(const sh of sheets)revealMinusCellsForAdmin(sh);
  await ctx.sync();
}
function maskMinusCells(sh){
  // VISUAL ONLY: Never replace number formats, values or formulas here.
  for(const address of MINES_CELLS){
    const r=sh.getRange(address);
    r.format.fill.color="#000000";
    r.format.font.color="#000000";
    r.format.protection.locked=true;
    r.format.protection.formulaHidden=false;
  }
}
async function restoreMinesReadableNumbers(context,sh,activeSheet=false){
  const ranges=MINES_CELLS.map(address=>{
    const r=sh.getRange(address);
    r.load("numberFormat");
    if(activeSheet)r.format.fill.load("color");
    return r;
  });
  await context.sync();
  let restored=0,shown=0;
  for(const r of ranges){
    if(minesHasInvisibleNumberFormat(r.numberFormat[0][0])){
      r.numberFormat=[[MINES_RECOVERY_NUMBER_FORMAT]];
      restored++;
    }
    if(activeSheet && String(r.format.fill.color||"").toUpperCase()==="#000000"){
      // The new active shift must not inherit the previous closed sheet's mask.
      r.format.fill.clear();
      r.format.font.color="#000000";
      shown++;
    }
  }
  if(restored||shown)await context.sync();
  return {restored,shown};
}
function fullLock(sh,pw){
  const w=sh.getRange(WORK_AREA); w.format.protection.locked=true; w.format.protection.formulaHidden=true;
  maskMinusCells(sh);
  sh.protection.protect({selectionMode:Excel.ProtectionSelectionMode.none,allowFormatCells:false,allowFormatColumns:false,allowFormatRows:false,allowInsertColumns:false,allowInsertRows:false,allowDeleteColumns:false,allowDeleteRows:false},pw);
}
// Reconcile older worksheets already locked before the masking feature existed.
// Never clear or replace source values/formulas; only apply display formatting.
async function maskExistingClosedSheets(){
  if(isGlobalUnlockMode())return;
  const pw=requireSecret();
  const shiftId=String(getSetting("hitunganV4.shiftWorkSheetId")||"");
  const shiftUntil=Number(getSetting("hitunganV4.shiftWorkUntil"));
  const shiftActive=Boolean(shiftId && Number.isFinite(shiftUntil) && Date.now()<shiftUntil);
  await Excel.run(async c=>{
    const sheets=c.workbook.worksheets;
    sheets.load("items/id,items/name,items/protection/protected");
    await c.sync();
    for(const sh of sheets.items){
      if(shiftActive && String(sh.id)===shiftId)continue;
      // Do not change an intentionally unlocked admin correction sheet.
      if(!sh.protection.protected)continue;
      sh.protection.unprotect(pw);
      maskMinusCells(sh);
      fullLock(sh,pw);
    }
    await c.sync();
  });
}

async function enforceStructure(){const pw=requireSecret();await Excel.run(async c=>{const wb=c.workbook;wb.protection.load("protected");await c.sync();if(!wb.protection.protected){wb.protection.protect(pw);await c.sync();}});}
async function checkPin(){if(adminSessionAuthenticated)return;throw new Error("Masuk ke Menu Admin menggunakan PIN terlebih dahulu.");}

function showRoleMenu(role){
  const choice=by("menuChoice"),employee=by("employeeWorkspace"),gate=by("adminGate"),admin=by("adminWorkspace");
  choice.hidden=role!=="home";
  employee.hidden=role!=="employee";
  gate.hidden=role!=="admin" || adminSessionAuthenticated;
  admin.hidden=role!=="admin" || !adminSessionAuthenticated;
  if(role!=="admin"){
    adminSessionAuthenticated=false;
    gate.hidden=true;
    admin.hidden=true;
    const pin=by("adminGatePin");if(pin)pin.value="";
    const err=by("adminGateStatus");if(err)err.textContent="";
    if(typeof hideMinusPanel==="function")hideMinusPanel();
    if(typeof sellingEditorRows!=="undefined")sellingEditorRows=[];
    const list=by("sellingPriceRows");if(list)list.replaceChildren();
  }
}
function setupRoleMenus(){
  by("menuEmployeeBtn").addEventListener("click",()=>showRoleMenu("employee"));
  by("menuAdminBtn").addEventListener("click",()=>showRoleMenu("admin"));
  by("employeeBackBtn").addEventListener("click",()=>showRoleMenu("home"));
  by("adminLogoutBtn").addEventListener("click",()=>showRoleMenu("home"));
  by("adminGateBackBtn").addEventListener("click",()=>showRoleMenu("home"));
  by("adminGateLoginBtn").addEventListener("click",unlockAdminMenu);
  by("adminGatePin").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();unlockAdminMenu();}});
  showRoleMenu("home");
}
async function unlockAdminMenu(){
  const status=by("adminGateStatus");
  try{
    if(!getSetting(CFG.enabled))throw new Error("Pengaman admin belum diaktifkan.");
    const pin=by("adminGatePin").value;
    if(!pin)throw new Error("Masukkan PIN admin.");
    if(await sha256(pin)!==getSetting(CFG.pinHash))throw new Error("PIN admin salah.");
    by("adminGatePin").value="";
    adminSessionAuthenticated=true;
    by("adminGate").hidden=true;
    by("adminWorkspace").hidden=false;
    by("adminSecurity").open=true;
    status.textContent="";
    if(typeof loadSellingPriceSheets==="function")await loadSellingPriceSheets();
  }catch(e){
    // A failure loading the optional editor must not leave a partially-open gate.
    if(adminSessionAuthenticated){
      setSellingStatus("Menu admin terbuka, tetapi daftar harga belum dimuat: "+errText(e),true);
    }else status.textContent=errText(e);
  }
}

async function adminUnlockSheet(){try{await checkPin();if(isGlobalUnlockMode())return setSecurityStatus("Semua sheet sudah terbuka dalam mode ADMIN BEBAS.",false);const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);await c.sync();setSecurityStatus(`ADMIN: sheet "${sh.name}" dibuka.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminWorkingLock(){try{await checkPin();assertGlobalUnlockOff();const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);workingLock(sh,pw);await c.sync();setSecurityStatus(`Sheet "${sh.name}" dikunci sebagai sheet kerja.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminFullLock(){try{await checkPin();assertGlobalUnlockOff();const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);fullLock(sh,pw);await c.sync();setSecurityStatus(`Sheet "${sh.name}" full lock.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminUnlockStructure(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const wb=c.workbook;wb.protection.load("protected");await c.sync();if(wb.protection.protected)wb.protection.unprotect(pw);await c.sync();});if(relockTimer)clearTimeout(relockTimer);relockTimer=setTimeout(()=>enforceStructure().catch(()=>{}),120000);setSecurityStatus("ADMIN: struktur terbuka maksimal 2 menit. Delete/Rename aktif.",false);}catch(e){setSecurityStatus(errText(e),true);}}
async function adminLockStructure(){try{await checkPin();if(relockTimer)clearTimeout(relockTimer);relockTimer=null;await enforceStructure();setSecurityStatus("Struktur workbook dikunci kembali.",false);}catch(e){setSecurityStatus(errText(e),true);}}
async function adminSync(){
  try{
    await checkPin();assertGlobalUnlockOff();
    const pw=requireSecret();
    const shiftId=String(getSetting("hitunganV4.shiftWorkSheetId")||"");
    const shiftUntil=Number(getSetting("hitunganV4.shiftWorkUntil"));
    const shiftValid=Boolean(shiftId&&Number.isFinite(shiftUntil)&&Date.now()<shiftUntil);
    const current=await Excel.run(async c=>{
      const wb=c.workbook,sheets=wb.worksheets;
      wb.protection.load("protected");
      sheets.load("items/id,items/name,items/protection/protected");
      await c.sync();
      if(wb.protection.protected)wb.protection.unprotect(pw);
      let workName="";
      for(const sh of sheets.items){
        if(sh.protection.protected)sh.protection.unprotect(pw);
        if(shiftValid&&String(sh.id)===shiftId){
          workingLock(sh,pw);workName=sh.name;
        }else{
          fullLock(sh,pw);
        }
      }
      wb.protection.protect(pw);
      await c.sync();
      return workName;
    });
    setSecurityStatus(current?
      "Sinkron selesai: '"+current+"' tetap sheet kerja karena 2,5 jam belum habis. Lainnya full lock.":
      "Sinkron selesai: tidak ada shift aktif; seluruh sheet full lock.",false);
  }catch(e){setSecurityStatus(errText(e),true);}
}

async function adminUnlockAllSheets(){
  let operationStarted=false;
  try{
    await checkPin();
    const pw=requireSecret();
    setSecurityStatus("ADMIN: membuka sheet dan menampilkan kembali angka mines…",false);
    adminGlobalUnmaskBusy=true;
    operationStarted=true;
    try{
      await Excel.run(async c=>{
        const wb=c.workbook,sheets=wb.worksheets;
        wb.protection.load("protected");
        sheets.load("items/name,items/protection/protected");
        await c.sync();
        for(const sh of sheets.items){
          if(sh.protection.protected)sh.protection.unprotect(pw);
        }
        await c.sync();
        await revealAllMinusCellsForAdmin(c,sheets.items);
        // Keep workbook STRUCTURE protected; only sheet access is opened.
        if(!wb.protection.protected)wb.protection.protect(pw);
        await c.sync();
      });
      setSetting(CFG.globalUnlock,true);
      await saveSettings();
    }finally{
      adminGlobalUnmaskBusy=false;
    }
    await beginGlobalUnlockWindow();
    if(typeof clearAdminSheetAutoLockState==="function"){
      try{await clearAdminSheetAutoLockState();}catch(_){}
    }
    setGlobalUnlockUi(true);
    setSecurityStatus("ADMIN BEBAS — SEMUA SHEET TERBUKA maksimal 1,5 jam. Tidak ada auto-lock sampai admin memilih KUNCI SEMUA KEMBALI.",false);
  }catch(e){
    if(operationStarted){
      // Partial Excel failure must not leave the workbook unlocked indefinitely.
      try{
        await lockAllSheetsCore();
      }catch(relockError){
        setSecurityStatus("GAGAL membuka seluruh sheet DAN mengunci ulang: "+errText(relockError)+
          ". Segera amankan workbook dari perangkat Admin.",true);
        return;
      }
    }
    setSecurityStatus("Gagal membuka seluruh sheet: "+errText(e)+
      ". Penguncian sheet sudah dipulihkan.",true);
  }
}

async function adminLockAllSheets(){
  try{
    await checkPin();
    await lockAllSheetsCore();
  }catch(e){setSecurityStatus(errText(e),true);}
}

async function lockAllSheetsCore(statusOverride){
  try{
    const pw=requireSecret();
    const shiftId=getSetting("hitunganV4.shiftWorkSheetId");
    const shiftUntil=Number(getSetting("hitunganV4.shiftWorkUntil"));
    const validShift=Boolean(shiftId&&Number.isFinite(shiftUntil)&&Date.now()<shiftUntil);
    setSecurityStatus("ADMIN: mengunci kembali seluruh sheet…",false);
    const result=await Excel.run(async c=>{
      const wb=c.workbook,sheets=wb.worksheets;
      wb.protection.load("protected");
      sheets.load("items/id,items/name,items/protection/protected");
      await c.sync();
      if(wb.protection.protected) wb.protection.unprotect(pw);
      let workingName="";
      for(const sh of sheets.items){
        if(sh.protection.protected) sh.protection.unprotect(pw);
        if(validShift&&String(sh.id)===String(shiftId)){
          // Current shift remains readable while its normal input window lasts.
          revealMinusCellsForAdmin(sh);
          workingLock(sh,pw);
          workingName=sh.name;
        }else{
          fullLock(sh,pw);
        }
      }
      wb.protection.protect(pw);
      await c.sync();
      return workingName;
    });
    setSetting(CFG.globalUnlock,false);
    await saveSettings();
    if(typeof clearGlobalUnlockWindowState==="function") await clearGlobalUnlockWindowState();
    if(!validShift&&typeof clearShiftWorkWindowState==="function"){
      try{await clearShiftWorkWindowState();}catch(_){}
    }
    setGlobalUnlockUi(false);
    setSecurityStatus(
      statusOverride || (result
        ? `MODE NORMAL — semua sheet dikunci kembali; "${result}" tetap sebagai sheet kerja karena shift masih aktif.`
        : "MODE NORMAL — seluruh sheet sudah FULL LOCK dan struktur workbook terkunci kembali."),
      false
    );
  }catch(e){setSecurityStatus(errText(e),true);throw e;}
}

async function ensureAllSheetsUnlocked(){
  if(!isGlobalUnlockMode()) return;
  const pw=requireSecret();
  await Excel.run(async c=>{
    const sheets=c.workbook.worksheets;
    sheets.load("items/protection/protected");
    await c.sync();
    for(const sh of sheets.items){
      if(sh.protection.protected)sh.protection.unprotect(pw);
    }
    await c.sync();
    await revealAllMinusCellsForAdmin(c,sheets.items);
  });
}

function isGlobalUnlockMode(){
  const v=getSetting(CFG.globalUnlock);
  return v===true||String(v).toLowerCase()==="true";
}

function assertGlobalUnlockOff(){
  if(isGlobalUnlockMode()) throw new Error('Mode BUKA SEMUA SHEET masih aktif. Gunakan tombol "KUNCI SEMUA KEMBALI" untuk mengakhiri mode admin bebas.');
}

function setGlobalUnlockUi(on){
  const open=by("adminUnlockAllSheetsBtn"),close=by("adminLockAllSheetsBtn"),hint=by("globalUnlockHint");
  if(open) open.disabled=Boolean(on);
  if(close) close.disabled=!on;
  ["adminUnlockSheetBtn","adminWorkingLockBtn","adminFullLockBtn","adminSyncBtn"].forEach(id=>{const el=by(id);if(el)el.disabled=Boolean(on);});
  if(hint) hint.textContent=on
    ? "AKTIF: semua sheet terbuka maksimal 1,5 jam, lalu otomatis dikunci. Shift yang masih aktif tetap dapat diisi."
    : "Dengan PIN admin, buka semua sheet maksimal 1,5 jam. Setelahnya sistem mengunci otomatis.";
}

function requireSecret(){const s=secret||getSetting(CFG.secret);if(!s)throw new Error("Pengaman V4 belum diaktifkan.");secret=s;return s;}
function setConfigured(ok){by("securitySetup").hidden=ok;by("adminSecurity").hidden=!ok;}
function setFormEnabled(ok){const f=by("newDayForm");f.classList.toggle("security-disabled",!ok);f.querySelectorAll("input,select,button").forEach(el=>el.disabled=!ok);}
function setSecurityStatus(t,e){const x=by("securityStatus");if(!x)return;x.textContent=t;x.className=e?"security-status error":"security-status ok";}
function by(id){return document.getElementById(id);}
function getSetting(k){return Office.context.document.settings.get(k);}
function setSetting(k,v){Office.context.document.settings.set(k,v);}
function saveSettings(){return new Promise((res,rej)=>Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?res():rej(r.error||new Error("Gagal menyimpan pengaturan."))));}
async function sha256(t){const d=new TextEncoder().encode(t),h=await crypto.subtle.digest("SHA-256",d);return Array.from(new Uint8Array(h)).map(b=>b.toString(16).padStart(2,"0")).join("");}
function randomSecret(){const b=new Uint8Array(24);crypto.getRandomValues(b);return Array.from(b).map(x=>x.toString(16).padStart(2,"0")).join("");}
function errText(e){return e&&e.message?e.message:String(e||"Terjadi kesalahan.");}

let minusPreviewTimer = null;
function hideMinusPanel(){
  const panel=by("minusPreview"), content=by("minusPreviewContent");
  if(minusPreviewTimer){clearTimeout(minusPreviewTimer);minusPreviewTimer=null;}
  if(content) content.replaceChildren();
  if(panel) panel.hidden=true;
}
async function adminViewMinus(){
  try{
    hideMinusPanel();
    await checkPin();
    const data=await Excel.run(async context=>{
      const sh=context.workbook.worksheets.getActiveWorksheet();
      sh.load("name");
      const ranges=MINES_CELLS.map(address=>sh.getRange(address));
      ranges.forEach(r=>r.load("values,text"));
      await context.sync();
      return {name:sh.name, rows:MINES_CELLS.map((address,i)=>({
        address, value:ranges[i].text[0][0],
        raw:ranges[i].values[0][0]
      }))};
    });
    const target=by("minusPreviewContent");
    const heading=document.createElement("strong");
    heading.textContent="Sheet: "+data.name;
    target.appendChild(heading);
    for(const row of data.rows){
      const line=document.createElement("div");
      line.style.padding="5px 0";
      // Read the actual value; previous builds hid formatted text with ';;;'.
      line.textContent=row.address+" : "+(row.raw===null||row.raw===""?"(kosong)":String(row.raw));
      target.appendChild(line);
    }
    by("minusPreview").hidden=false;
    minusPreviewTimer=setTimeout(hideMinusPanel,5*60*1000);
    setSecurityStatus("Nilai mines ditampilkan untuk admin selama maksimal 5 menit. Angka tetap tersembunyi di worksheet.",false);
  }catch(e){hideMinusPanel();setSecurityStatus(errText(e),true);}
}
