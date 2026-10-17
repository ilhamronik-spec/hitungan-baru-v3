"use strict";

/*
 * Pengaman Rumus V6.2 — acuan BENAR dibekukan dari WORKSHEET OKTOBER(5).xlsx,
 * sheet "17 oktober", 18 Oktober 2026.
 *
 * Read-only audit of formulas/values; only cell colors are changed on errors.
 * NEVER fixes/replaces formulas, sales prices, quantities, or stock.
 *
 * IMPORTANT: the reference is a committed JavaScript snapshot, NOT the live
 * source sheet. A damaged source sheet cannot silently rewrite the standard.
 */

const HB_FORMULA_GUARD_BUILD="6.2.0";
const HB_FORMULA_GUARD_JOURNAL_KEY="hitunganV4.formulaGuardColorJournalV1";
const HB_FORMULA_GUARD_BLACK="#000000";
const HB_FORMULA_GUARD_RED="#FF3030";
const HB_FORMULA_MINES=new Set(["D275","D276","G275"]);
const HB_FORMULA_F_BLANKS=new Set([19,20,32,33,43,52,53,58,59,60,65,66]);

// Snapshot formulas are stored as Excel A1 formulas. {r} is the row number.
// Source patterns were checked against all 17 October reference formulas.
const HB_FORMULA_PATTERNS=[
  {from:6,to:75,columns:{
    F:"=D{r}-E{r}",G:"=C{r}*F{r}",
    H:"=(C{r}-B{r})*F{r}",I:"=B{r}*E{r}"}},
  {from:93,to:113,columns:{
    B:"commission:A",F:"commission:E",
    J:"commission:I",N:"commission:M"}},
  {from:178,to:239,columns:{
    D:"=C{r}*G{r}",E:"=B{r}-C{r}",
    H:"=F{r}-G{r}",J:"=E{r}*H{r}",
    L:"=F{r}-K{r}",N:"=M{r}-L{r}+O{r}",
    P:"=N{r}*C{r}"}}
];
// Preserve this complete 17 October commission calculation, including its
// threshold ordering; never substitute a revised financial calculation.
const HB_COMMISSION_93="=IF(A93>2000000,20000,IF(A93>25000000,25000,IF(A93>3000000,30000,IF(A93>1500000,15000,IF(A93>999000,10000,IF(A93>500000,8000,IF(A93>300000,5000,IF(A93>100000,3000,IF(A93>19000,2000,IF(A93>1000,1000,IF(A93<1000,0)))))))))))";
const HB_FORMULA_FIXED={
  B273:"=SUM(B248:B271)+B272",
  B274:"=B273-B272",
  C274:"=B275+G273",
  C275:"=B274-C274-N163+K279",
  D275:"=C275+M23-L50",
  D276:"=D275+L79+K79",
  G275:"=B273-(G272+G273+N163-K279)+L79+K79",
  G273:"=S8+L75-K75",
  J273:"=G273",
  B271:"=D241+P240",
  B242:"=D241+I240-J241",
  D241:"=SUM(D178:D240)",
  J241:"=SUM(J178:J240)",
  P240:"=SUM(P178:P239)",
  W129:"=SUM(X122:X128)",
  W143:"=SUM(X136:X142)",
  W157:"=SUM(X150:X156)",
  R249:"=SUM(S122:S248)",
  J170:"=SUM(K151:K167)",
  K145:"=SUM(L122:L144)",
  B165:"=SUM(D150:D162*1800)",
  F173:"=SUM(G151:G170)",
  B115:"=SUM(B93:B113)+SUM(F93:F113)+SUM(J93:J113)+SUM(N93:N113)+SUM(R93:R113)",
  B145:"=SUM(B122:B144)+SUM(F122:F144)",
  H76:"=SUM(H6:H75)",
  H79:"=H77-H76",
  I76:"=SUM(I6:I75)",
  K75:"=SUM(K55:K74)",
  L75:"=SUM(L55:L74)",
  L50:"=SUM(L31:L49)",
  M23:"=SUM(M7:M22)",
  S8:"=H76+B115+B145+K145+J170+F173+R249+B165+W129+N71+J241+W157+W143+N151+N157+N162",
  B272:"=SUM(B278:B427)+SUM(G279:G399)"
};

function hbFormulaColumnNumber(letter){
  let n=0;
  for(const char of letter)n=n*26+char.charCodeAt(0)-64;
  return n-1;
}
function hbFormulaColumnLetter(index){
  let s="";
  for(let n=index+1;n;n=Math.floor((n-1)/26))
    s=String.fromCharCode(65+(n-1)%26)+s;
  return s;
}
function hbCommissionFormula(sourceCol,row){
  // Only A93 tokens are translated. Numeric commission thresholds stay fixed.
  return HB_COMMISSION_93.replace(/A93/g,sourceCol+row);
}
function hbExpectedPattern(template,row){
  return template.indexOf("commission:")===0
    ? hbCommissionFormula(template.slice("commission:".length),row)
    : template.replace(/\{r\}/g,String(row));
}
function hbReferenceFormulas(){
  const map=new Map();
  for(const group of HB_FORMULA_PATTERNS){
    for(const [col,template] of Object.entries(group.columns)){
      for(let row=group.from;row<=group.to;row++){
        const address=col+row;
        const intentionallyBlank=(col==="F"&&row>=6&&row<=75&&HB_FORMULA_F_BLANKS.has(row))||
          (row===239&&["H","L","N"].includes(col));
        map.set(address,intentionallyBlank?null:hbExpectedPattern(template,row));
      }
    }
  }
  for(const [addr,formula] of Object.entries(HB_FORMULA_FIXED))map.set(addr,formula);
  return map;
}
const HB_FORMULA_REFERENCE=hbReferenceFormulas();

function hbFormulaMatches(actual,expected){
  if(expected===null)return actual===null||actual===undefined||actual==="";
  return typeof actual==="string"&&actual===expected;
}
function hbFormulaGuardJournal(){
  if(typeof getSetting!=="function")throw new Error("Pengaturan pengaman workbook belum siap.");
  const raw=getSetting(HB_FORMULA_GUARD_JOURNAL_KEY);
  return raw&&typeof raw==="object"&&!Array.isArray(raw)?JSON.parse(JSON.stringify(raw)):{};
}
function hbFormulaIssueDetail(issue){
  if(issue.expected===null)
    return issue.address+" — SEHARUSNYA KOSONG, tetapi sekarang terisi";
  const actual=issue.actual;
  if(typeof actual!=="string"||!actual.startsWith("="))
    return issue.address+" — RUMUS HILANG / DIGANTI ANGKA ATAU TEKS"+
      (issue.expected.length<=70?" (acuan "+issue.expected+")":"");
  return issue.address+" — RUMUS BERUBAH DARI POLA / ACUAN"+
    (issue.expected.length<=70?" (seharusnya "+issue.expected+")":"");
}
function hbFormulaGuardError(issues,sheetName){
  const details=issues.slice(0,45).map(hbFormulaIssueDetail);
  const remainder=issues.length-details.length;
  const prefix='PEMBUATAN SHEET DIBLOKIR — '+issues.length+
    ' kesalahan rumus pada sheet sumber "'+(sheetName||"aktif")+'".';
  return prefix+"\nLOKASI KESALAHAN:\n"+details.join("\n")+
    (remainder?"\n... dan "+remainder+" alamat lainnya.":"")+
    "\nSel bermasalah ditandai FILL HITAM dan FONT MERAH."+
    "\nBetulkan rumus pada sheet sumber terlebih dahulu; rumus/stok tidak diperbaiki otomatis.";
}
function hbShowFormulaNote(message){
  if(typeof setSecurityStatus==="function")setSecurityStatus(message,true);
  // Reuse existing panel/form status. No extra menus or controls.
  const el=typeof document!=="undefined"?document.getElementById("securityStatus"):null;
  if(el){
    el.style.whiteSpace="pre-line";
    el.style.maxHeight="290px";
    el.style.overflowY="auto";
  }
  const form=typeof document!=="undefined"?document.getElementById("status"):null;
  if(form){
    form.textContent=message;
    form.style.whiteSpace="pre-line";
  }
}
async function hbScanActiveSheetFormulas(){
  const groupRanges=[
    {range:"F6:I75",fromRow:6,fromCol:hbFormulaColumnNumber("F")},
    {range:"B93:N113",fromRow:93,fromCol:hbFormulaColumnNumber("B")},
    {range:"D178:P239",fromRow:178,fromCol:hbFormulaColumnNumber("D")}
  ];
  return Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getActiveWorksheet();
    sh.load("id,name,protection/protected");
    const chunks=groupRanges.map(group=>{
      const range=sh.getRange(group.range);
      range.load("formulas");
      return {...group,range};
    });
    const fixed=Object.entries(HB_FORMULA_FIXED).map(([address,expected])=>{
      const range=sh.getRange(address);range.load("formulas");
      return {address,expected,range};
    });
    await ctx.sync();
    const actuals=new Map();
    for(const group of chunks){
      const data=group.range.formulas;
      for(let r=0;r<data.length;r++){
        for(let col=0;col<data[r].length;col++){
          const address=hbFormulaColumnLetter(group.fromCol+col)+(group.fromRow+r);
          if(HB_FORMULA_REFERENCE.has(address))actuals.set(address,data[r][col]);
        }
      }
    }
    for(const {address,range} of fixed)actuals.set(address,range.formulas[0][0]);
    const issues=[];
    for(const [address,expected] of HB_FORMULA_REFERENCE){
      if(!actuals.has(address))throw new Error("Tidak dapat membaca rumus acuan di "+address);
      const actual=actuals.get(address);
      if(!hbFormulaMatches(actual,expected)){
        issues.push({address,expected,actual});
      }
    }
    return {sheetId:sh.id,sheetName:sh.name,issues,checked:HB_FORMULA_REFERENCE.size};
  });
}

function hbFormulaColorReadSafe(v){
  return typeof v==="string"&&v.trim()!==""?v:null;
}
async function hbPaintFormulaIssues(result){
  const previous=hbFormulaGuardJournal();
  const oldSheet=previous[result.sheetId]||{name:result.sheetName,cells:{}};
  const recorded=oldSheet.cells||{};
  const toIssue=new Set(result.issues.map(x=>x.address));
  const toStart=result.issues.map(x=>x.address).filter(addr=>!recorded[addr]);
  const toRestore=Object.keys(recorded).filter(addr=>!toIssue.has(addr));
  const adminOpen=typeof isGlobalUnlockMode==="function"&&isGlobalUnlockMode();
  const shouldPaint=address=>!HB_FORMULA_MINES.has(address)||adminOpen;
  const starts=toStart.filter(shouldPaint);
  const restores=toRestore.filter(shouldPaint);
  const marked=Object.keys(recorded).filter(addr=>toIssue.has(addr)&&shouldPaint(addr));
  if(!starts.length&&!restores.length&&!marked.length)return;
  const colors=await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(result.sheetId);
    const inspect=[...new Set(starts.concat(marked))].map(address=>{
      const range=sh.getRange(address);
      range.format.fill.load("color");
      range.format.font.load("color");
      return {address,range};
    });
    await ctx.sync();
    return inspect.map(({address,range})=>({
      address,fill:hbFormulaColorReadSafe(range.format.fill.color),
      font:hbFormulaColorReadSafe(range.format.font.color)
    }));
  });
  const toRepaint=colors.filter(color=>
    marked.includes(color.address) &&
    (String(color.fill||"").toUpperCase()!==HB_FORMULA_GUARD_BLACK ||
     String(color.font||"").toUpperCase()!==HB_FORMULA_GUARD_RED)
  ).map(color=>color.address);
  const toPaint=[...new Set(starts.concat(toRepaint))];
  if(!toPaint.length&&!restores.length)return;
  // Persist original colors BEFORE editing Excel so a refresh can restore them.
  const entry={name:result.sheetName,cells:{...recorded}};
  for(const color of colors){
    if(starts.includes(color.address))
      entry.cells[color.address]={fill:color.fill,font:color.font};
  }
  if(starts.length){
    previous[result.sheetId]=entry;
    setSetting(HB_FORMULA_GUARD_JOURNAL_KEY,previous);
    await saveSettings();
  }
  await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(result.sheetId);
    sh.load("id,protection/protected");
    const marker=sh.getRange("A1");
    marker.format.protection.load("locked");
    await ctx.sync();
    const protectedBefore=sh.protection.protected;
    const wasFull=protectedBefore&&marker.format.protection.locked===true;
    const pw=protectedBefore?requireSecret():null;
    let unprotected=false,failed=null;
    try{
      if(protectedBefore){
        // Mark before the Office batch is sent: a failed sync may still have
        // unlocked Excel, and finally MUST attempt to restore protection.
        unprotected=true;
        sh.protection.unprotect(pw);
        await ctx.sync();
      }
      for(const address of toPaint){
        const cell=sh.getRange(address);
        cell.format.fill.color=HB_FORMULA_GUARD_BLACK;
        cell.format.font.color=HB_FORMULA_GUARD_RED;
      }
      for(const address of restores){
        const cell=sh.getRange(address),original=recorded[address];
        if(original.fill)cell.format.fill.color=original.fill;
        else cell.format.fill.clear();
        if(original.font)cell.format.font.color=original.font;
        else cell.format.font.color="#000000";
      }
      await ctx.sync();
    }catch(e){failed=e;}
    finally{
      if(unprotected){
        try{
          if(wasFull)fullLock(sh,pw);
          else workingLock(sh,pw);
          await ctx.sync();
        }catch(relockError){
          throw new Error("Tidak dapat memulihkan proteksi setelah penanda rumus: "+
            (relockError.message||String(relockError)));
        }
      }
    }
    if(failed)throw failed;
  });
  for(const address of restores)delete entry.cells[address];
  if(!Object.keys(entry.cells).length)delete previous[result.sheetId];
  setSetting(HB_FORMULA_GUARD_JOURNAL_KEY,previous);
  await saveSettings();
}

let hbFormulaRunning=null;
async function hbVerifySourceFormulas(blockCreation=false){
  if(hbFormulaRunning){
    await hbFormulaRunning;
    // In a concurrent preflight, rescan rather than reusing a stale result.
  }
  const run=(async()=>{
    const result=await hbScanActiveSheetFormulas();
    let formatFailure=null;
    try{await hbPaintFormulaIssues(result);}catch(e){formatFailure=e;}
    if(result.issues.length){
      const msg=hbFormulaGuardError(result.issues,result.sheetName);
      const notice=msg+(formatFailure?
        "\nPERINGATAN: Sebagian warna belum berhasil diterapkan ("+formatFailure.message+").":"");
      hbShowFormulaNote(notice);
      if(blockCreation)throw new Error(notice);
    }else if(formatFailure){
      // Valid source formulas must not be blocked merely because an old
      // warning color could not be restored by Excel.
      if(typeof setSecurityStatus==="function")
        setSecurityStatus("Rumus sumber BENAR, tetapi pemulihan warna peringatan belum selesai: "+
          formatFailure.message,true);
    }
    return result;
  })();
  hbFormulaRunning=run;
  try{return await run;}finally{if(hbFormulaRunning===run)hbFormulaRunning=null;}
}

Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel)return;
  const passiveAudit=()=>{
    if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)return;
    if(typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress)return;
    if(typeof adminGlobalUnmaskBusy!=="undefined"&&adminGlobalUnmaskBusy)return;
    hbVerifySourceFormulas(false).catch(e=>{
      if(typeof setSecurityStatus==="function")
        setSecurityStatus("Pengaman rumus belum dapat diperiksa: "+e.message,true);
    });
  };
  // Existing panel only: audit on load and then periodically as long as the
  // panel remains open; no new menus, no workbook formula/value edits.
  setTimeout(passiveAudit,4500);
  setInterval(passiveAudit,60000);
});
