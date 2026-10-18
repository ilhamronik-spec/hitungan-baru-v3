"use strict";

// Separate, verifiable recovery for already-protected historical worksheets.
// This does not change cell values or formulas. The existing V4 fullLock handles
// protection, while this routine verifies the actual displayed cell formatting.
const MINES_MASK_BUILD = "5.5.0";
const MINES_RETRY_DELAY = 20000;
let minesRecoveryBusy = false;
let minesRecoveryTimer = null;
let minesRecoveryAttempts = 0;

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  const btn = document.getElementById("repairMinusMasksBtn");
  if (btn) btn.addEventListener("click", () => reconcileMinesMasks(true));
  updateMinesRecoveryStatus("Versi 5.5.0 — memulihkan teks mines agar terbaca SmartSDM...",false);
  scheduleMinesRecovery(2000);
});

function updateMinesRecoveryStatus(message, error) {
  const el = document.getElementById("minesMaskStatus");
  if (!el) return;
  el.textContent = message;
  el.style.color = error ? "#ae2020" : "#23643f";
}

function scheduleMinesRecovery(delay) {
  if (minesRecoveryTimer) clearTimeout(minesRecoveryTimer);
  minesRecoveryTimer = setTimeout(() => {
    minesRecoveryTimer = null;
    reconcileMinesMasks(false).catch(() => {});
  },delay);
}

function minesExpectedExcludedIds() {
  const now = Date.now();
  const shiftId = getSetting("hitunganV4.shiftWorkSheetId");
  const shiftUntil = Number(getSetting("hitunganV4.shiftWorkUntil"));
  const adminId = getSetting("hitunganV4.adminUnlockSheetId");
  const adminUntil = Number(getSetting("hitunganV4.adminUnlockUntil"));
  const result = new Set();
  if (shiftId && Number.isFinite(shiftUntil) && shiftUntil > now) result.add(String(shiftId));
  if (adminId && Number.isFinite(adminUntil) && adminUntil > now) result.add(String(adminId));
  return result;
}

function minesCellReadable(r){
  const f=String(r.numberFormat[0][0]||"");
  const number=r.values[0][0];
  return !minesHasInvisibleNumberFormat(f) &&
    (number===null||number===""||String(r.text[0][0]||"").trim()!=="");
}
async function scanMinesMasks(){
  const excluded=minesExpectedExcludedIds();
  return Excel.run(async c=>{
    const sheets=c.workbook.worksheets;
    sheets.load("items");
    await c.sync();
    const found=[];
    for(const sh of sheets.items){
      sh.load("id,name,protection/protected");
      const probes=MINES_CELLS.map(address=>{
        const r=sh.getRange(address);
        r.format.fill.load("color");
        r.format.font.load("color");
        r.load("numberFormat,values,text");
        return r;
      });
      found.push({sh,probes});
    }
    await c.sync();
    return found.map(({sh,probes})=>{
      const visibleToApps=probes.every(minesCellReadable);
      const fillCorrect=probes.every(r=>
        String(r.format.fill.color||"").toUpperCase()==="#000000"&&
        String(r.format.font.color||"").toUpperCase()==="#000000"
      );
      return {id:String(sh.id),name:sh.name,
        protected:Boolean(sh.protection.protected),
        masked:visibleToApps&&fillCorrect,
        unreadable:!visibleToApps,
        excluded:excluded.has(String(sh.id))};
    });
  });
}

async function maskOneClosedSheet(entry,pw){
  return Excel.run(async c=>{
    const sh=c.workbook.worksheets.getItem(entry.name);
    sh.load("id,name,protection/protected");
    await c.sync();
    if(minesExpectedExcludedIds().has(String(sh.id))||isGlobalUnlockMode())return "skipped";
    if(sh.protection.protected){
      sh.protection.unprotect(pw);
      await c.sync();
    }
    // Repair ONLY previous hidden-number-format damage; do not change data.
    await restoreMinesReadableNumbers(c,sh,false);
    maskMinusCells(sh);
    await c.sync();
    fullLock(sh,pw);
    await c.sync();

    sh.protection.load("protected");
    const checks=MINES_CELLS.map(address=>{
      const r=sh.getRange(address);
      r.format.fill.load("color");
      r.format.font.load("color");
      r.load("numberFormat,values,text");
      return r;
    });
    await c.sync();
    if(!sh.protection.protected||
       checks.some(r=>String(r.format.fill.color||"").toUpperCase()!=="#000000"||
                      String(r.format.font.color||"").toUpperCase()!=="#000000"||
                      !minesCellReadable(r))){
      throw new Error('Teks angka / format hitam sheet "'+entry.name+'" belum berhasil diverifikasi.');
    }
    return "masked";
  });
}

// Existing active shifts from older builds must also become readable, while
// preserving their original remaining 2.5-hour window and protection mode.
async function restoreActiveShiftMines(){
  if(isGlobalUnlockMode())return 0;
  const id=String(getSetting("hitunganV4.shiftWorkSheetId")||"");
  const until=Number(getSetting("hitunganV4.shiftWorkUntil"));
  if(!id||!Number.isFinite(until)||until<=Date.now())return 0;
  const pw=requireSecret();
  return Excel.run(async c=>{
    const sh=c.workbook.worksheets.getItem(id);
    sh.load("protection/protected");
    const ranges=MINES_CELLS.map(a=>{
      const r=sh.getRange(a);
      r.load("numberFormat");
      return r;
    });
    await c.sync();
    if(!ranges.some(r=>minesHasInvisibleNumberFormat(r.numberFormat[0][0])))return 0;
    if(sh.protection.protected){
      sh.protection.unprotect(pw);
      await c.sync();
    }
    const result=await restoreMinesReadableNumbers(c,sh,true);
    workingLock(sh,pw);
    await c.sync();
    return result.restored;
  });
}

async function reconcileMinesMasks(manual=false){
  if((typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
     (typeof hbFormulaAuditBusy!=="undefined"&&hbFormulaAuditBusy)){
    updateMinesRecoveryStatus("Menunggu proses sheet baru / pemeriksaan rumus selesai...",false);
    scheduleMinesRecovery(12000);
    return;
  }
  // Don't re-apply black masking in the middle of an authenticated
  // BUKA SEMUA SHEET operation before the global-unlock flag is saved.
  if(typeof adminGlobalUnmaskBusy!=="undefined"&&adminGlobalUnmaskBusy){
    updateMinesRecoveryStatus("Menunggu pembukaan semua sheet oleh Admin...",false);
    scheduleMinesRecovery(9000);
    return;
  }
  if(minesRecoveryBusy){
    if(manual) updateMinesRecoveryStatus("Pemeriksaan sedang berjalan; tunggu hasilnya.",false);
    return;
  }
  if(typeof hbMarginColorBusy !== "undefined" && hbMarginColorBusy){
    updateMinesRecoveryStatus("Menunggu pembaruan warna margin selesai...", false);
    scheduleMinesRecovery(9000);
    return;
  }
  minesRecoveryBusy = true;
  const btn=document.getElementById("repairMinusMasksBtn");
  if(btn)btn.disabled=true;
  try{
    if(getSetting("hitunganV4.enabled")!==true){
      updateMinesRecoveryStatus("Pengaman belum aktif; penutup mines belum dijalankan.",true);
      return;
    }
    if(isGlobalUnlockMode()){
      updateMinesRecoveryStatus("Mode BUKA SEMUA aktif. Penguncian dan penutup mines berjalan setelah mode ini berakhir.",false);
      return;
    }
    updateMinesRecoveryStatus("Memeriksa angka mines dan memulihkan format teks yang disembunyikan versi lama...",false);
    const activeRestored=await restoreActiveShiftMines();
    const list=await scanMinesMasks();
    const targets=list.filter(s=>!s.excluded && (!s.masked || !s.protected));
    let completed=0;
    const failed=[];
    const pw=requireSecret();
    for(const entry of targets){
      if(isGlobalUnlockMode()||
         (typeof adminGlobalUnmaskBusy!=="undefined"&&adminGlobalUnmaskBusy))break;
      try {
        if(await maskOneClosedSheet(entry,pw)==="masked")completed++;
      }catch(e){
        failed.push(entry.name+": "+(e&&e.message||String(e)));
      }
    }
    // A separate readback proves that the black fill actually exists.
    const final=await scanMinesMasks();
    const unfinished=final.filter(s=>!s.excluded && (!s.masked || !s.protected));
    if(unfinished.length){
      const names=unfinished.slice(0,5).map(s=>s.name).join(", ");
      throw new Error("Masih gagal: "+names+(failed.length?" — "+failed[0]:""));
    }
    minesRecoveryAttempts=0;
    updateMinesRecoveryStatus("BERHASIL — "+(final.length-final.filter(s=>s.excluded).length)+
       " sheet diperiksa; "+activeRestored+" format pada shift aktif dipulihkan. Nilai tetap ada, teks angka bisa dibaca aplikasi, dan warna hitam tetap menutup sheet terkunci. Build "+MINES_MASK_BUILD,false);
  }catch(e){
    minesRecoveryAttempts++;
    const reason=e&&e.message?e.message:String(e);
    updateMinesRecoveryStatus("BELUM BERHASIL — "+reason+
      ". Tekan Esc/Enter di Excel, lalu klik PERBAIKI PENUTUP MINES. Percobaan ulang otomatis berjalan.",true);
    if(minesRecoveryAttempts<=10) scheduleMinesRecovery(MINES_RETRY_DELAY);
  }finally{
    minesRecoveryBusy=false;
    if(btn)btn.disabled=false;
  }
}
