"use strict";

// Separate, verifiable recovery for already-protected historical worksheets.
// This does not change cell values or formulas. The existing V4 fullLock handles
// protection, while this routine verifies the actual displayed cell formatting.
const MINES_MASK_BUILD = "4.8.0";
const MINES_RETRY_DELAY = 20000;
let minesRecoveryBusy = false;
let minesRecoveryTimer = null;
let minesRecoveryAttempts = 0;

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  const btn = document.getElementById("repairMinusMasksBtn");
  if (btn) btn.addEventListener("click", () => reconcileMinesMasks(true));
  updateMinesRecoveryStatus("Versi 4.8.0 — memeriksa penutup angka sheet lama...",false);
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

async function scanMinesMasks() {
  const excluded = minesExpectedExcludedIds();
  return Excel.run(async c => {
    const sheets=c.workbook.worksheets;
    sheets.load("items");
    await c.sync();
    const data=[];
    for (const sh of sheets.items) {
      sh.load("id,name,protection/protected");
      const probes=MINES_CELLS.map(address=>{
        const r=sh.getRange(address);
        r.format.fill.load("color");
        r.load("numberFormat");
        return r;
      });
      data.push({sh,probes});
    }
    await c.sync();
    return data.map(({sh,probes})=>{
      const masked=probes.every(r=>
        String(r.format.fill.color||"").toUpperCase()==="#000000" &&
        String(r.numberFormat[0][0]||"").trim()===";;;"
      );
      return {id:String(sh.id),name:sh.name,
        protected:Boolean(sh.protection.protected),masked,
        excluded:excluded.has(String(sh.id))};
    });
  });
}

async function maskOneClosedSheet(entry, pw) {
  return Excel.run(async c=>{
    const sh=c.workbook.worksheets.getItem(entry.name);
    sh.load("id,name,protection/protected");
    await c.sync();
    if(minesExpectedExcludedIds().has(String(sh.id)) || isGlobalUnlockMode())return "skipped";
    // A protected sheet must be unlocked and synchronized before formatting.
    // Excel Online may reject changing formats in the same unprotect batch.
    if(sh.protection.protected){
      sh.protection.unprotect(pw);
      await c.sync();
    }
    maskMinusCells(sh);
    await c.sync();
    fullLock(sh,pw);
    await c.sync();
    sh.protection.load("protected");
    const checks=MINES_CELLS.map(address=>{
      const r=sh.getRange(address);
      r.format.fill.load("color");
      r.load("numberFormat");
      return r;
    });
    await c.sync();
    if(!sh.protection.protected ||
       checks.some(r=>String(r.format.fill.color||"").toUpperCase()!=="#000000" ||
                      String(r.numberFormat[0][0]||"").trim()!==";;;")){
      throw new Error('Format hitam/proteksi sheet "'+entry.name+'" belum terverifikasi');
    }
    return "masked";
  });
}

async function reconcileMinesMasks(manual=false){
  if(minesRecoveryBusy){
    if(manual) updateMinesRecoveryStatus("Pemeriksaan sedang berjalan; tunggu hasilnya.",false);
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
    updateMinesRecoveryStatus("Memeriksa empat sel mines pada sheet lama...",false);
    const list=await scanMinesMasks();
    const targets=list.filter(s=>!s.excluded && (!s.masked || !s.protected));
    let completed=0;
    const failed=[];
    const pw=requireSecret();
    for(const entry of targets){
      if(isGlobalUnlockMode())break;
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
       " sheet lama sudah diverifikasi. Empat sel mines berwarna hitam dan terkunci. Build "+MINES_MASK_BUILD,false);
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
