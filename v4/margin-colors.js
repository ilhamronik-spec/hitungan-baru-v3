"use strict";

// V5 — Margins are evaluated on every worksheet from that worksheet's own
// B (harga dasar) and C (harga jual). Color is applied only to item names in A.
// Conditional formatting is dynamic and persists when Excel recalculates.
const HB_MARGIN_BUILD = "5.0.0";
const HB_MARGIN_TAG = "HB_MARGIN_V50";
const HB_MARGIN_GROUPS = [
  { name:"Paket", first:6, last:75 },
  { name:"Rokok", first:178, last:238 }
];
const HB_MARGIN_RULES = [
  {key:"NEGATIVE", test:"d<0", fill:"#000000", font:"#FFFFFF", name:"Minus"},
  {key:"RED", test:"AND(d>=0,d<500)", fill:"#EF4444", font:"#FFFFFF", name:"Di bawah Rp500"},
  {key:"GREY", test:"AND(d>=500,d<1000)", fill:"#9CA3AF", font:"#111827", name:"Rp500–999"},
  {key:"BLUE", test:"AND(d>1500,d<=3000)", fill:"#2563EB", font:"#FFFFFF", name:"Di atas Rp1.500"},
  {key:"PURPLE", test:"d>3000", fill:"#7C3AED", font:"#FFFFFF", name:"Di atas Rp3.000"}
];
const HB_MARGIN_RETRY_MS = 25000;
let hbMarginColorBusy = false;
let hbMarginTimer = null;
let hbMarginFailures = 0;
let hbMarginRuleScanInProgress = false;

function hbMarginBand(margin) {
  if(!Number.isFinite(margin))return "NONE";
  if(margin<0)return "NEGATIVE";
  if(margin<500)return "RED";
  if(margin<1000)return "GREY";
  if(margin>3000)return "PURPLE";
  if(margin>1500)return "BLUE";
  return "NONE";
}

// Formula uses an explicit marker through N("text") so that duplicated rules
// can be recognized without deleting pre-existing workbook formatting.
function hbMarginRuleFormula(first, rule) {
  const diff="$C"+first+"-$B"+first;
  const guard='AND($A'+first+'<>"",ISNUMBER($B'+first+'),ISNUMBER($C'+first+'),$B'+first+'>0)';
  const test=rule.test.replace(/\bd\b/g,diff);
  const tag=HB_MARGIN_TAG+"_"+first+"_"+rule.key;
  return "=AND("+guard+","+test+',N("'+tag+'")=0)';
}

function hbMarginStatus(text,isError=false) {
  const el=document.getElementById("marginColorStatus");
  if(!el)return;
  el.textContent=text;
  el.style.color=isError?"#B42318":"#17643B";
}

function hbMarginQueueRetry(delay=HB_MARGIN_RETRY_MS){
  if(hbMarginTimer)clearTimeout(hbMarginTimer);
  hbMarginTimer=setTimeout(()=>{
    hbMarginTimer=null;
    hbMarginReconcileAll(false).catch(()=>{});
  },delay);
}

function hbMarginExcludedSheetIds(){
  const excluded=new Set();
  const now=Date.now();
  const shiftId=getSetting("hitunganV4.shiftWorkSheetId");
  const until=Number(getSetting("hitunganV4.shiftWorkUntil"));
  if(shiftId&&Number.isFinite(until)&&until>now)excluded.add(String(shiftId));
  const adminId=getSetting("hitunganV4.adminUnlockSheetId");
  const adminUntil=Number(getSetting("hitunganV4.adminUnlockUntil"));
  if(adminId&&Number.isFinite(adminUntil)&&adminUntil>now)excluded.add(String(adminId));
  return excluded;
}

async function hbEnsureMarginRules(context,sheet){
  // No destructive clearAll(): preserve rules the user already has.
  const descriptors=[];
  for(const g of HB_MARGIN_GROUPS){
    const range=sheet.getRange("A"+g.first+":A"+g.last);
    const formats=range.conditionalFormats;
    formats.load("type");
    descriptors.push({g,range,formats});
  }
  await context.sync();
  for(const x of descriptors){
    x.custom=x.formats.items.filter(format=>format.type===Excel.ConditionalFormatType.custom||format.type==="Custom");
    x.custom.forEach(format=>format.custom.load("rule"));
  }
  await context.sync();

  let added=0;
  for(const x of descriptors){
    const formulas=x.custom.map(format=>String(format.custom.rule.formula||"").toUpperCase());
    for(const rule of HB_MARGIN_RULES){
      const marker=(HB_MARGIN_TAG+"_"+x.g.first+"_"+rule.key).toUpperCase();
      if(formulas.some(s=>s.includes(marker)))continue;
      const cf=x.range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula=hbMarginRuleFormula(x.g.first,rule);
      cf.custom.format.fill.color=rule.fill;
      cf.custom.format.font.color=rule.font;
      cf.stopIfTrue=true;
      added++;
    }
  }
  if(added)await context.sync();
  return added;
}


async function hbMarginRulesInstalled(context,sheet){
  // Read before touching protection: restarting the pane must not repeatedly
  // unprotect and re-protect 16+ historical sheets unnecessarily.
  const groups=HB_MARGIN_GROUPS.map(g=>{
    const formats=sheet.getRange("A"+g.first+":A"+g.last).conditionalFormats;
    formats.load("type");
    return {first:g.first,formats};
  });
  await context.sync();
  for(const g of groups){
    g.custom=g.formats.items.filter(x=>x.type===Excel.ConditionalFormatType.custom||x.type==="Custom");
    g.custom.forEach(x=>x.custom.load("rule"));
  }
  await context.sync();
  for(const g of groups){
    const formulas=g.custom.map(f=>String(f.custom.rule.formula||"").toUpperCase());
    for(const rule of HB_MARGIN_RULES){
      const token=(HB_MARGIN_TAG+"_"+g.first+"_"+rule.key).toUpperCase();
      if(!formulas.some(f=>f.includes(token)))return false;
    }
  }
  return true;
}

async function hbMarginApplyToSheet(entry,allowTrackedShift=false){
  return Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(entry.name);
    sh.load("id,name,protection/protected");
    await ctx.sync();
    if(String(sh.id)!==String(entry.id))throw new Error("Identitas sheet berubah, batalkan pengeditan.");
    if(isGlobalUnlockMode())return 0;
    const excluded=hbMarginExcludedSheetIds();
    if(excluded.has(String(sh.id))&&!allowTrackedShift)return 0;
    if(await hbMarginRulesInstalled(ctx,sh))return 0;
    const originallyProtected=Boolean(sh.protection.protected);
    const secretPw=requireSecret();
    if(originallyProtected){
      sh.protection.unprotect(secretPw);
      await ctx.sync();
    }
    let added=0;
    try{
      added=await hbEnsureMarginRules(ctx,sh);
    }finally{
      // If the admin opened all sheets while we were processing, preserve
      // that mode. Never return from a finally block (it would hide errors).
      if(!isGlobalUnlockMode()){
        // Restore protection even if a format update failed.
        if(originallyProtected){
          const fresh=hbMarginExcludedSheetIds();
          if(fresh.has(String(sh.id))){
            workingLock(sh,secretPw);
          }else{
            fullLock(sh,secretPw);
          }
          await ctx.sync();
        }else if(!excluded.has(String(sh.id))){
          // Old unprotected sheets without an admin/shift exception are closed.
          fullLock(sh,secretPw);
          await ctx.sync();
        }
      }
    }
    return added;
  });
}

async function hbMarginInstallNewSheet(name){
  if(isGlobalUnlockMode())return 0;
  const entry=await Excel.run(async ctx=>{
    const sheet=ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("id,name,protection/protected");
    await ctx.sync();
    return {id:String(sheet.id),name:sheet.name,protected:sheet.protection.protected};
  });
  if(entry.name!==name)throw new Error("Sheet terbaru berubah; warna margin tidak dipaksakan ke sheet lain.");
  return hbMarginApplyToSheet(entry,true);
}

async function hbMarginReadSheets(){
  return Excel.run(async ctx=>{
    const sheets=ctx.workbook.worksheets;
    sheets.load("items/id,items/name,items/protection/protected");
    await ctx.sync();
    return sheets.items.map(s=>({id:String(s.id),name:s.name,protected:s.protection.protected}));
  });
}

async function hbMarginReconcileAll(manual=false){
  if(hbMarginColorBusy){
    if(manual)hbMarginStatus("Pemeriksaan warna masih berjalan. Tunggu hingga selesai.");
    return;
  }
  if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)return;
  if(isGlobalUnlockMode()){
    hbMarginStatus("Mode BUKA SEMUA sedang aktif. Pemasangan warna menunggu proteksi normal.");
    return;
  }
  if(typeof minesRecoveryBusy!=="undefined"&&minesRecoveryBusy){
    hbMarginQueueRetry(7000);
    hbMarginStatus("Menunggu proses penutup angka mines selesai...");
    return;
  }
  hbMarginColorBusy=true;
  const btn=document.getElementById("marginColorRefreshBtn");
  if(btn)btn.disabled=true;
  const failures=[];
  let completed=0,created=0,skipped=0;
  try{
    const sheets=await hbMarginReadSheets();
    const excluded=hbMarginExcludedSheetIds();
    hbMarginStatus("Memasang/memeriksa aturan warna pada "+sheets.length+" sheet...");
    for(const sh of sheets){
      if(isGlobalUnlockMode())break;
      if(excluded.has(sh.id)){skipped++;continue;}
      try{
        created+=await hbMarginApplyToSheet(sh);
        completed++;
      }catch(e){
        failures.push(sh.name+" — "+(e&&e.message||String(e)));
      }
    }
    if(failures.length)throw new Error(failures.slice(0,2).join("; "));
    hbMarginFailures=0;
    hbMarginStatus("BERHASIL — "+completed+" sheet diperiksa; "+
      created+" aturan warna baru dipasang. "+skipped+" sheet aktif/admin dilewati. "+
      "Warna nama barang akan mengikuti margin otomatis. Versi "+HB_MARGIN_BUILD+".");
  }catch(e){
    hbMarginFailures++;
    hbMarginStatus("BELUM SELESAI — "+(e&&e.message||String(e))+
      ". Coba keluar dari edit sel (Esc/Enter), lalu tekan Periksa Ulang Warna.",true);
    if(hbMarginFailures<=8)hbMarginQueueRetry();
  }finally{
    hbMarginColorBusy=false;
    if(btn)btn.disabled=false;
  }
}

Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel)return;
  const btn=document.getElementById("marginColorRefreshBtn");
  if(btn)btn.addEventListener("click",()=>hbMarginReconcileAll(true));
  // Startup also migrates historical sheets without rebuilding or rewriting data.
  hbMarginQueueRetry(6500);
});
