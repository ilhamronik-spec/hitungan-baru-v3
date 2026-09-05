"use strict";

const V2_SCRIPT = "https://hitungan-baru-excel-ilham.bienbachthuatz2.chatgpt.site/taskpane.js";
const CFG = { enabled:"hitunganV4.enabled", pinHash:"hitunganV4.pinHash", secret:"hitunganV4.secret", version:"hitunganV4.version" };
const LEGACY_V3_SECRET_KEY = "hitunganV3.secret";
const WORK_AREA = "A1:AH400";
const LOCK_RANGES = [
  "B267","B269","B271:B275","G272:G275","F178:F238","K178:K238","L178:L238","M178:M238","N178:N238","O178:O238",
  "N163","B6:B75","C6:C75","D6:D75","F6:F75","G6:G75","H6:H75","M23","L50","K75","L75","H76:H79","B114:B116",
  "B145:B147","B163:B165","J168:J170","F171:F173","R249:R251","K145:K147","W143:W145","W157:W159","W129:W131","Y178","Y237"
];
let secret = null;
let relockTimer = null;
let v2Loaded = false;

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
      await Excel.run(async c=>{
        const wb=c.workbook, sheets=wb.worksheets, newest=sheets.getActiveWorksheet(), old=sheets.getItem(sourceName);
        wb.protection.load("protected"); newest.load("protection/protected"); old.load("protection/protected");
        await c.sync();
        if(newest.protection.protected) newest.protection.unprotect(pw);
        if(old.protection.protected) old.protection.unprotect(pw);
        fullLock(old,pw); workingLock(newest,pw);
        if(!wb.protection.protected) wb.protection.protect(pw);
        await c.sync();
      });
      setSecurityStatus("AMAN — hari baru selesai, sheet lama full lock, struktur terkunci.",false);
      return result;
    }catch(err){
      await failClosedRecovery(sourceName,sheetName,pw);
      throw err;
    }
  };

  if(typeof repairActiveStorageFormulas==="function"){
    const originalRepair=repairActiveStorageFormulas;
    repairActiveStorageFormulas=async function(){
      const pw=requireSecret();
      await Excel.run(async c=>{ const sh=c.workbook.worksheets.getActiveWorksheet(); sh.load("protection/protected"); await c.sync(); if(sh.protection.protected) sh.protection.unprotect(pw); await c.sync(); });
      try{ return await originalRepair(); }
      finally{ await Excel.run(async c=>{ const sh=c.workbook.worksheets.getActiveWorksheet(); sh.load("protection/protected"); await c.sync(); if(sh.protection.protected) sh.protection.unprotect(pw); workingLock(sh,pw); await c.sync(); }); }
    };
  }
}

async function failClosedRecovery(sourceName,newName,pw){
  try{
    await Excel.run(async c=>{
      const wb=c.workbook,sheets=wb.worksheets; wb.protection.load("protected"); sheets.load("items/name,items/protection/protected"); await c.sync();
      for(const sh of sheets.items){
        if(sh.protection.protected){ try{ sh.protection.unprotect(pw); }catch(_){} }
        if(sh.name===sourceName) workingLock(sh,pw);
        else if(sh.name===newName) fullLock(sh,pw);
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
    by("adminUnlockSheetBtn").addEventListener("click",adminUnlockSheet);
    by("adminWorkingLockBtn").addEventListener("click",adminWorkingLock);
    by("adminFullLockBtn").addEventListener("click",adminFullLock);
    by("adminUnlockStructureBtn").addEventListener("click",adminUnlockStructure);
    by("adminLockStructureBtn").addEventListener("click",adminLockStructure);
    by("adminSyncBtn").addEventListener("click",adminSync);
    const enabled=getSetting(CFG.enabled)===true, saved=getSetting(CFG.secret);
    if(enabled&&saved){ secret=saved; setConfigured(true); setFormEnabled(true); try{await enforceStructure();setSecurityStatus("AMAN — struktur workbook terkunci.",false);}catch(e){setSecurityStatus(errText(e),true);} }
    else{ setConfigured(false); setFormEnabled(false); setSecurityStatus("Belum diaktifkan. Buat PIN admin terlebih dahulu.",true); }
  });
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
    setSetting(CFG.pinHash,await sha256(p1)); setSetting(CFG.secret,newPw); setSetting(CFG.enabled,true); setSetting(CFG.version,"4.0.1"); await saveSettings(); secret=newPw;
    by("setupPin").value="";by("setupPin2").value="";by("oldPassword").value="";setConfigured(true);setFormEnabled(true);setSecurityStatus("AMAN — V4 aktif. Sheet aktif=kerja, sheet lama=full lock, struktur=lock.",false);
  }catch(e){setSecurityStatus(errText(e),true);}
}

function workingLock(sh,pw){
  const w=sh.getRange(WORK_AREA); w.format.protection.locked=false; w.format.protection.formulaHidden=false;
  for(const a of LOCK_RANGES){const r=sh.getRange(a);r.format.protection.locked=true;r.format.protection.formulaHidden=true;}
  sh.protection.protect({selectionMode:Excel.ProtectionSelectionMode.unlocked,allowFormatCells:false,allowFormatColumns:false,allowFormatRows:false,allowInsertColumns:false,allowInsertRows:false,allowDeleteColumns:false,allowDeleteRows:false},pw);
}
function fullLock(sh,pw){
  const w=sh.getRange(WORK_AREA); w.format.protection.locked=true; w.format.protection.formulaHidden=true;
  sh.protection.protect({selectionMode:Excel.ProtectionSelectionMode.none,allowFormatCells:false,allowFormatColumns:false,allowFormatRows:false,allowInsertColumns:false,allowInsertRows:false,allowDeleteColumns:false,allowDeleteRows:false},pw);
}
async function enforceStructure(){const pw=requireSecret();await Excel.run(async c=>{const wb=c.workbook;wb.protection.load("protected");await c.sync();if(!wb.protection.protected){wb.protection.protect(pw);await c.sync();}});}
async function checkPin(){const p=by("adminPin").value;if(!p)throw new Error("Masukkan PIN admin.");if(await sha256(p)!==getSetting(CFG.pinHash))throw new Error("PIN admin salah.");by("adminPin").value="";}
async function adminUnlockSheet(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);await c.sync();setSecurityStatus(`ADMIN: sheet "${sh.name}" dibuka.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminWorkingLock(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);workingLock(sh,pw);await c.sync();setSecurityStatus(`Sheet "${sh.name}" dikunci sebagai sheet kerja.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminFullLock(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const sh=c.workbook.worksheets.getActiveWorksheet();sh.load("name,protection/protected");await c.sync();if(sh.protection.protected)sh.protection.unprotect(pw);fullLock(sh,pw);await c.sync();setSecurityStatus(`Sheet "${sh.name}" full lock.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}
async function adminUnlockStructure(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const wb=c.workbook;wb.protection.load("protected");await c.sync();if(wb.protection.protected)wb.protection.unprotect(pw);await c.sync();});if(relockTimer)clearTimeout(relockTimer);relockTimer=setTimeout(()=>enforceStructure().catch(()=>{}),120000);setSecurityStatus("ADMIN: struktur terbuka maksimal 2 menit. Delete/Rename aktif.",false);}catch(e){setSecurityStatus(errText(e),true);}}
async function adminLockStructure(){try{await checkPin();if(relockTimer)clearTimeout(relockTimer);relockTimer=null;await enforceStructure();setSecurityStatus("Struktur workbook dikunci kembali.",false);}catch(e){setSecurityStatus(errText(e),true);}}
async function adminSync(){try{await checkPin();const pw=requireSecret();await Excel.run(async c=>{const wb=c.workbook,sheets=wb.worksheets,active=sheets.getActiveWorksheet();wb.protection.load("protected");sheets.load("items/id,items/name,items/protection/protected");active.load("id,name");await c.sync();if(wb.protection.protected)wb.protection.unprotect(pw);for(const sh of sheets.items){if(sh.protection.protected)sh.protection.unprotect(pw);if(sh.id===active.id)workingLock(sh,pw);else fullLock(sh,pw);}wb.protection.protect(pw);await c.sync();setSecurityStatus(`Sinkron selesai. "${active.name}"=kerja; lainnya=full lock.`,false);});}catch(e){setSecurityStatus(errText(e),true);}}

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
