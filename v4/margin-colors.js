"use strict";

// V5 — Margins are evaluated on every worksheet from that worksheet's own
// B (harga dasar) and C (harga jual). Color is applied only to item names in A.
// Conditional formatting is dynamic and persists when Excel recalculates.
const HB_MARGIN_BUILD = "5.2.0";
const HB_MARGIN_TAG = "HB_MARGIN_V51";
const HB_MARGIN_OLD_TAG = "HB_MARGIN_V50";
const HB_MARGIN_CIGARETTE_OLD_TAG = "HB_MARGIN_V51_178_";
const HB_CIGARETTE_ROWS_START = 178;
const HB_CIGARETTE_ROWS_END = 238;
const HB_CIGARETTE_NORMAL = {fill:"#FFFFFF",font:"#111827"};
const HB_MARGIN_GROUPS = [
  { name:"Paket", first:6, last:75, mode:"packet" },
  { name:"Rokok", first:178, last:238, mode:"cigarette" }
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
function hbMarginRuleFormula(group, rule) {
  const first=group.first;
  // Paket: cost B, selling C. Rokok: native margin is already calculated in E.
  // Never compute C-B on cigarette rows: it reverses their margin sign.
  const d=group.mode==="cigarette"?"$E"+first:"($C"+first+"-$B"+first+")";
  const guard=group.mode==="cigarette"
    ? 'AND($A'+first+'<>"",ISNUMBER($B'+first+'),ISNUMBER($C'+first+'),ISNUMBER($E'+first+'),$B'+first+'>0,$C'+first+'>0)'
    : 'AND($A'+first+'<>"",ISNUMBER($B'+first+'),ISNUMBER($C'+first+'),$B'+first+'>0,$C'+first+'>0)';
  const test=rule.test.replace(/\bd\b/g,d);
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

async function hbMarginInspectRules(context,sheet){
  const groups=HB_MARGIN_GROUPS.map(g=>{
    const formats=sheet.getRange("A"+g.first+":A"+g.last).conditionalFormats;
    formats.load("type");
    return {g,formats};
  });
  await context.sync();
  for(const group of groups){
    group.custom=group.formats.items.filter(f=>f.type===Excel.ConditionalFormatType.custom || f.type==="Custom");
    group.custom.forEach(f=>f.custom.load("rule"));
  }
  await context.sync();
  return groups;
}


function hbMarginIsLegacyCigaretteRule(formula) {
  const s=String(formula||"").toUpperCase();
  return s.includes("HB_MARGIN_V50_178_") || s.includes("HB_MARGIN_V51_178_");
}
function hbMarginNeedsMigration(groups) {
  return groups.some(({g,custom})=>{
    const formulas=custom.map(f=>String(f.custom.rule.formula||"").toUpperCase());
    if(g.mode==="cigarette")return formulas.some(hbMarginIsLegacyCigaretteRule);
    if(formulas.some(f=>f.includes(HB_MARGIN_OLD_TAG)))return true;
    return HB_MARGIN_RULES.some(rule=>!formulas.some(f=>f.includes((HB_MARGIN_TAG+"_"+g.first+"_"+rule.key).toUpperCase())));
  });
}
function hbCigaretteRowStyle(row) {
  const [name,selling,cost,unused,marginRaw]=row;
  if(!String(name??"").trim())return null;
  if(typeof selling!=="number" || !Number.isFinite(selling) || selling<=0)return null;
  if(typeof cost!=="number" || !Number.isFinite(cost) || cost<=0)return null;
  // The sheet's own E column is the authoritative cigarette margin.
  const margin=typeof marginRaw==="number"&&Number.isFinite(marginRaw)?marginRaw:NaN;
  const band=hbMarginBand(margin);
  const matched=HB_MARGIN_RULES.find(rule=>rule.key===band);
  return {fill:matched?.fill||HB_CIGARETTE_NORMAL.fill,
          font:matched?.font||HB_CIGARETTE_NORMAL.font,
          band,margin};
}
async function hbReadCigaretteDisplay(context,sheet){
  const data=sheet.getRange("A"+HB_CIGARETTE_ROWS_START+":E"+HB_CIGARETTE_ROWS_END);
  data.load("values");
  const names=[];
  for(let row=HB_CIGARETTE_ROWS_START;row<=HB_CIGARETTE_ROWS_END;row++){
    const cell=sheet.getRange("A"+row);
    cell.format.fill.load("color");
    cell.format.font.load("color");
    names.push(cell);
  }
  await context.sync();
  return {values:data.values,names};
}
function hbColorCompare(actual,expected){
  return String(actual||"").trim().toUpperCase()===String(expected||"").trim().toUpperCase();
}
function hbCigaretteNeedsColor(display){
  return display.values.some((row,i)=>{
    const desired=hbCigaretteRowStyle(row);
    if(!desired)return false;
    const cell=display.names[i];
    return !hbColorCompare(cell.format.fill.color,desired.fill)
      || !hbColorCompare(cell.format.font.color,desired.font);
  });
}
async function hbMarginRulesInstalled(context,sheet){
  const groups=await hbMarginInspectRules(context,sheet);
  if(hbMarginNeedsMigration(groups))return false;
  const rokok=await hbReadCigaretteDisplay(context,sheet);
  return !hbCigaretteNeedsColor(rokok);
}
async function hbEnsureMarginRules(context,sheet){
  const groups=await hbMarginInspectRules(context,sheet);
  let removed=0,added=0, recolored=0;
  for(const group of groups){
    for(const format of group.custom){
      const formula=String(format.custom.rule.formula||"").toUpperCase();
      const remove=group.g.mode==="cigarette"
        ? hbMarginIsLegacyCigaretteRule(formula)
        : formula.includes(HB_MARGIN_OLD_TAG);
      if(remove){format.delete();removed++;}
    }
  }
  if(removed)await context.sync();
  // Keep dynamic Conditional Formatting for Paket (it is already working).
  const packageGroup=groups.find(g=>g.g.mode==="packet");
  if(packageGroup){
    const existing=packageGroup.custom.map(f=>String(f.custom.rule.formula||"").toUpperCase())
       .filter(f=>!f.includes(HB_MARGIN_OLD_TAG));
    for(const rule of HB_MARGIN_RULES){
      const marker=(HB_MARGIN_TAG+"_"+packageGroup.g.first+"_"+rule.key).toUpperCase();
      if(existing.some(f=>f.includes(marker)))continue;
      const cf=packageGroup.formats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula=hbMarginRuleFormula(packageGroup.g,rule);
      cf.custom.format.fill.color=rule.fill;
      cf.custom.format.font.color=rule.font;
      cf.stopIfTrue=true;
      added++;
    }
  }
  if(added)await context.sync();

  // Cigarette template has BLACK base fill even when there is NO loss.
  // Explicitly replace its base fill and font using E, not C-B.
  const rokok=await hbReadCigaretteDisplay(context,sheet);
  for(let i=0;i<rokok.values.length;i++){
    const style=hbCigaretteRowStyle(rokok.values[i]);
    if(!style)continue; // Leave separators/headings and blank products unchanged.
    const nameCell=rokok.names[i];
    if(!hbColorCompare(nameCell.format.fill.color,style.fill) ||
       !hbColorCompare(nameCell.format.font.color,style.font)){
      nameCell.format.fill.color=style.fill;
      nameCell.format.font.color=style.font;
      recolored++;
    }
  }
  if(recolored)await context.sync();

  const verified=await hbReadCigaretteDisplay(context,sheet);
  if(hbCigaretteNeedsColor(verified)){
    throw new Error('Warna Rokok masih berbeda dari nilai margin kolom E setelah penyimpanan.');
  }
  return added+recolored+removed;
}

async function hbMarginApplyToSheet(entry,allowTrackedShift=false){
  return Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(entry.name);
    sh.load("id,name,position,protection/protected");
    const workingMarker=sh.getRange("A1");
    workingMarker.format.protection.load("locked");
    await ctx.sync();
    if(String(sh.id)!==String(entry.id))throw new Error("Identitas sheet berubah, batalkan pengeditan.");
    if(isGlobalUnlockMode())return 0;
    const excluded=hbMarginExcludedSheetIds();
    if(excluded.has(String(sh.id))&&!allowTrackedShift)return 0;
    if(await hbMarginRulesInstalled(ctx,sh))return 0;
    const originallyProtected=Boolean(sh.protection.protected);
    const wasLatestWorking=originallyProtected&&Number(sh.position)===0&&
      workingMarker.format.protection.locked===false;
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
          // A newest sheet can be in protected working mode while another
          // device's shift timestamp is still syncing. Never promote it to
          // FULL LOCK just because settings do not yet contain its ID.
          const tracked=String(getSetting("hitunganV4.shiftWorkSheetId")||"");
          const deadline=Number(getSetting("hitunganV4.shiftWorkUntil"));
          const expiredHere=tracked===String(sh.id)&&
            Number.isFinite(deadline)&&deadline>0&&deadline<=Date.now();
          if(fresh.has(String(sh.id))||(wasLatestWorking&&!expiredHere)){
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
  if((typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
     (typeof hbFormulaAuditBusy!=="undefined"&&hbFormulaAuditBusy)){
    hbMarginStatus("Menunggu pembuatan shift / pemeriksaan rumus selesai sebelum memeriksa warna.");
    hbMarginQueueRetry(12000);
    return;
  }
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
    const now=Date.now();
    const shiftId=String(getSetting("hitunganV4.shiftWorkSheetId")||"");
    const shiftUntil=Number(getSetting("hitunganV4.shiftWorkUntil"));
    const validShift=shiftId&&Number.isFinite(shiftUntil)&&shiftUntil>now;
    const adminId=String(getSetting("hitunganV4.adminUnlockSheetId")||"");
    const adminUntil=Number(getSetting("hitunganV4.adminUnlockUntil"));
    const validAdmin=adminId&&Number.isFinite(adminUntil)&&adminUntil>now;
    hbMarginStatus("Memasang/memeriksa aturan warna pada "+sheets.length+" sheet...");
    for(const sh of sheets){
      if(isGlobalUnlockMode())break;
      const working=Boolean(validShift&&sh.id===shiftId);
      if(validAdmin&&sh.id===adminId){skipped++;continue;}
      if(excluded.has(sh.id)&&!working){skipped++;continue;}
      try{
        created+=await hbMarginApplyToSheet(sh,working);
        completed++;
      }catch(e){
        failures.push(sh.name+" — "+(e&&e.message||String(e)));
      }
    }
    if(failures.length)throw new Error(failures.slice(0,2).join("; "));
    hbMarginFailures=0;
    hbMarginStatus("BERHASIL — "+completed+" sheet diperiksa; "+
      created+" pembaruan aturan/warna. "+skipped+" sheet admin dilewati. "+
      "Warna dasar Rokok diverifikasi langsung dari margin E, aturan paket tetap dinamis. Versi "+HB_MARGIN_BUILD+".");
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
