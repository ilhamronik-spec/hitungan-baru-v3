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

const HB_FORMULA_GUARD_BUILD="6.3.8";
const HB_FORMULA_GUARD_JOURNAL_KEY="hitunganV4.formulaGuardColorJournalV1";
const HB_CORRECTION_ACCESS_KEY="hitunganV4.anomalyCorrectionAccessV1";
const HB_FORMULA_GUARD_BLACK="#000000";
const HB_FORMULA_GUARD_RED="#FF3030";
const HB_FORMULA_MINES=new Set(["D275","D276","G275"]);
const HB_FORMULA_F_BLANKS=new Set([19,20,32,33,43,52,53,58,59,60,65,66]);
// Extra STOCK-CONSISTENCY GATE, newest sheet only. Values of L178:L238
// are differences (F-K), not individual stock-transfer inputs.
// Zero is valid (no difference); positive values must be multiples of ten.
// A negative value, fraction, nonnumber, or Excel error blocks new sheets.
const HB_ROKOK_L_FIRST=178;
const HB_ROKOK_L_LAST=238;
const HB_ROKOK_L_CHECK_COUNT=HB_ROKOK_L_LAST-HB_ROKOK_L_FIRST+1;
function hbRokokLIsAllowed(value){
  return typeof value==="number"&&Number.isSafeInteger(value)&&
    value>=0&&value%10===0;
}
function hbRokokLReadable(value){
  if(value===null||value===undefined||value==="")return "(kosong)";
  return String(value);
}

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
// The same price/margin calculation that makes the product NAME black.
 // Paket: C (harga jual) - B (harga beli); Rokok: native margin E.
 // Never flag a black-filled mines cell or a black/red formula alert as a
 // negative-margin product. The price guards mirror margin-colors.js.
function hbNegativeProductMargin(name,selling,cost,margin,category,row,address){
  if(!String(name??"").trim()||
     typeof selling!=="number"||!Number.isFinite(selling)||selling<=0||
     typeof cost!=="number"||!Number.isFinite(cost)||cost<=0||
     typeof margin!=="number"||!Number.isFinite(margin)||margin>=0)return null;
  return {category,name:String(name).trim(),selling,cost,margin,row,address};
}
function hbMarginValueLabel(value){
  return new Intl.NumberFormat("id-ID",{maximumFractionDigits:2}).format(value);
}
function hbMarginIssueText(info){
  return info.address+" — "+info.category+" "+info.name+
    " (baris "+info.row+") HITAM karena MARGIN MINUS "+
    hbMarginValueLabel(info.margin)+
    ". Harga jual "+hbMarginValueLabel(info.selling)+
    ", harga beli "+hbMarginValueLabel(info.cost)+".";
}
// Generate the exact, minimal editable INPUT addresses for each anomaly.
// Do not open name cells merely because they are black due to margin.
// Do not make calculated L or confidential mines editable to disguise errors.
function hbCorrectionTargets(issues){
  const targets=new Set();
  for(const issue of issues||[]){
    if(issue.marginIssue){
      const info=issue.marginIssue,row=Number(info.row);
      if(Number.isInteger(row)){
        if(info.category==="PAKET"&&row>=6&&row<=75)
          targets.add("C"+row); // Paket: selling price, not product name A
        else if(info.category==="ROKOK"&&row>=178&&row<=238)
          targets.add("B"+row); // Rokok: selling price; E is calculated
      }
    }
    if(issue.kind==="L_MULTIPLE_10" ||
       Object.prototype.hasOwnProperty.call(issue,"stock10Error")){
      const row=Number(String(issue.address||"").replace(/^L/,""));
      if(Number.isInteger(row)&&row>=178&&row<=238)
        targets.add("K"+row); // L=F-K: allow correction of ending display K
    }
    // Actual formula anomaly: allow only that broken formula cell, never
    // sensitive mines values. Any new A12 anomaly would unlock only A12.
    if(Object.prototype.hasOwnProperty.call(issue,"expected") &&
       !HB_FORMULA_MINES.has(issue.address) &&
       /^[A-Z]{1,2}[1-9][0-9]*$/.test(String(issue.address||""))){
      // Do not open calculated L for a stock-multiple issue unless
      // the L formula itself is independently incorrect.
      targets.add(issue.address);
    }
  }
  return [...targets].sort((a,b)=>{
    const ra=Number(a.match(/\d+/)[0]),rb=Number(b.match(/\d+/)[0]);
    return ra-rb||a.localeCompare(b);
  });
}
function hbCorrectionAccessJournal(){
  const raw=getSetting(HB_CORRECTION_ACCESS_KEY);
  return raw&&typeof raw==="object"&&!Array.isArray(raw)
    ?JSON.parse(JSON.stringify(raw)):{};
}

// A correction is COPYABLE only when the committed snapshot contains
// a literal reference formula for this exact cell. Never synthesize
// formulas from an adjacent (possibly damaged/locked) worksheet.
function hbCopyableExpectedFormula(issue){
  if(!issue||HB_FORMULA_MINES.has(issue.address))return null;
  return typeof issue.expected==="string"&&issue.expected.startsWith("=")
    ?issue.expected:null;
}
async function hbCopyFormulaToClipboard(input,button,status){
  const formula=String(input.value||"");
  if(!formula.startsWith("="))return;
  let copied=false;
  try{
    if(typeof navigator!=="undefined"&&navigator.clipboard&&
       typeof navigator.clipboard.writeText==="function"){
      await navigator.clipboard.writeText(formula);
      copied=true;
    }
  }catch(_){}
  if(!copied){
    // Excel Online can deny navigator.clipboard in the add-in iframe.
    // Always leave the entire exact formula selected for Ctrl+C.
    input.focus();
    input.select();
    try{
      if(typeof document!=="undefined"&&typeof document.execCommand==="function")
        copied=document.execCommand("copy")===true;
    }catch(_){}
  }
  if(status)status.textContent=copied
    ?"Rumus lengkap tersalin. Tempel (Ctrl+V) di sel yang ditandai."
    :"Rumus lengkap sudah dipilih. Tekan Ctrl+C lalu tempel di sel tujuan.";
  if(button&&copied)button.textContent="Tersalin";
}
function hbRenderFormulaCorrectionList(result){
  const container=typeof document!=="undefined"
    ?document.getElementById("formulaCorrectionCopies"):null;
  if(!container)return;
  container.replaceChildren();
  const issues=(result&&Array.isArray(result.issues))?result.issues:[];
  const formulas=issues.filter(issue=>hbCopyableExpectedFormula(issue));
  const blankCells=issues.filter(issue=>issue&&issue.expected===null);
  if(!issues.length){
    container.hidden=true;
    return;
  }
  container.hidden=false;
  const heading=document.createElement("div");
  heading.className="formula-correction-heading";
  heading.textContent="RUMUS BENAR — SALIN KE SEL YANG SALAH";
  container.appendChild(heading);
  const info=document.createElement("p");
  info.className="formula-correction-intro";
  info.textContent=formulas.length
    ?("Ada "+formulas.length+" rumus perbaikan pada sheet terbaru. "+
      "Tekan Salin Rumus, pilih sel tujuannya di Excel, lalu Ctrl+V. "+
      "Rumus diambil dari acuan tetap, bukan sheet di atas.")
    :"Tidak ada rumus yang perlu disalin. Periksa keterangan anomali stok/harga di bawah.";
  container.appendChild(info);
  for(const issue of formulas){
    const formula=hbCopyableExpectedFormula(issue);
    const item=document.createElement("div");
    item.className="formula-correction-item";
    const label=document.createElement("label");
    label.className="formula-correction-label";
    label.textContent="Sel "+issue.address+" — rumus yang benar:";
    const input=document.createElement("textarea");
    input.className="formula-correction-value";
    input.readOnly=true;
    input.rows=formula.length>100?4:2;
    input.value=formula;
    input.setAttribute("aria-label","Rumus benar untuk "+issue.address);
    const copy=document.createElement("button");
    copy.type="button";
    copy.className="secondary-button formula-correction-copy";
    copy.textContent="Salin Rumus "+issue.address;
    const status=document.createElement("small");
    status.className="formula-correction-copy-status";
    copy.addEventListener("click",()=>hbCopyFormulaToClipboard(input,copy,status));
    label.appendChild(input);
    item.appendChild(label);
    item.appendChild(copy);
    item.appendChild(status);
    container.appendChild(item);
  }
  for(const issue of blankCells){
    if(HB_FORMULA_MINES.has(issue.address))continue;
    const hint=document.createElement("p");
    hint.className="formula-correction-intro";
    hint.textContent="Sel "+issue.address+
      " seharusnya KOSONG — hapus isi sel tersebut, jangan tempel rumus.";
    container.appendChild(hint);
  }
  if(issues.some(issue=>HB_FORMULA_MINES.has(issue.address))){
    const note=document.createElement("p");
    note.className="formula-correction-intro";
    note.textContent="Koreksi sel mines rahasia harus melalui Admin; rumusnya tidak ditampilkan kepada karyawan.";
    container.appendChild(note);
  }
  // Entries are read-only documentation. This UI never modifies Excel cells,
  // their protection, existing formulas, prices or stock.
}

function hbFormulaIssueDetail(issue){
  if(issue.kind==="NEGATIVE_MARGIN")return hbMarginIssueText(issue.marginIssue);
  if(issue.marginIssue){
    const {marginIssue,...rest}=issue;
    return hbFormulaIssueDetail(rest)+"; "+hbMarginIssueText(marginIssue);
  }
  if(issue.kind==="L_MULTIPLE_10"){
    return issue.address+" — NILAI "+hbRokokLReadable(issue.actual)+
      " BUKAN KELIPATAN 10 NONNEGATIF (harus 0, 10, 20, 30, 40, 50, dst).";
  }
  if(issue.stock10Error){
    const {stock10Error,...formulaIssue}=issue;
    return hbFormulaIssueDetail(formulaIssue)+
      "; NILAI "+hbRokokLReadable(stock10Error)+
      " BUKAN KELIPATAN 10 NONNEGATIF.";
  }
  if(issue.expected===null)
    return issue.address+" — SEHARUSNYA KOSONG, tetapi sekarang terisi";
  if(HB_FORMULA_MINES.has(issue.address))
    return issue.address+" — RUMUS RAHASIA SALAH; hanya Admin boleh memperbaiki.";
  const actual=issue.actual;
  if(typeof actual!=="string"||!actual.startsWith("="))
    return issue.address+" — RUMUS HILANG / DIGANTI ANGKA ATAU TEKS. "+
      "RUMUS BENAR: "+issue.expected;
  return issue.address+" — RUMUS BERUBAH DARI POLA / ACUAN. "+
    "RUMUS BENAR: "+issue.expected;
}
function hbFormulaGuardError(issues,sheetName){
  const details=issues.slice(0,45).map(hbFormulaIssueDetail);
  const remainder=issues.length-details.length;
  const prefix='PEMBUATAN SHEET DIBLOKIR — '+issues.length+
    ' anomali rumus, stok Rokok, atau margin minus Paket/Rokok pada sheet terbaru "'+
    (sheetName||"aktif")+'".';
  return prefix+"\nLOKASI KESALAHAN:\n"+details.join("\n")+
    (remainder?"\nALAMAT TAMBAHAN ("+remainder+"): "+
      issues.slice(details.length).map(x=>x.address).join(", "):"")+
    "\nSel bermasalah ditandai FILL HITAM dan FONT MERAH."+
    "\nPerbaiki rumus, selisih stok, atau harga yang menyebabkan margin minus pada SHEET TERBARU."+
    "\nPenanda hitam-merah dan warna hitam margin akan hilang ketika anomali sudah benar."+
    "\nKOREKSI KARYAWAN: hanya "+(hbCorrectionTargets(issues).join(", ")||"tidak ada sel yang aman untuk dibuka")+
    " yang boleh diedit. Nama barang hitam tidak otomatis boleh diedit."+
    "\nSel lain tetap terkunci sesuai mode proteksi; setelah perbaikan lolos, akses koreksi ditutup otomatis."+
    "\nAplikasi tidak mengubah rumus, harga, atau stok secara otomatis.";
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

function hbFormulaSummaryText(result,formatFailure){
  const issues=result.issues||[];
  const header=issues.length?"HASIL PEMERIKSAAN: "+issues.length+" ANOMALI":"HASIL PEMERIKSAAN: LOLOS";
  const lines=[
    header,
    "Sheet terbaru: "+result.sheetName,
    "Diperiksa: "+result.checked+" posisi rumus + "+(result.stockCheckCount||0)+
      " nilai L178:L238 + "+(result.marginCheckCount||0)+
      " baris margin Paket/Rokok | Anomali: "+issues.length,
    "Acuan tetap: 17 oktober",
    "Waktu: "+new Date().toLocaleString("id-ID")
  ];
  if(issues.length){
    lines.push("KEPUTUSAN: DIBLOKIR sampai rumus, kelipatan 10, dan margin Paket/Rokok tidak minus.");
    const corrections=hbCorrectionTargets(issues);
    lines.push("SEL KOREKSI KARYAWAN (tanpa PIN): "+(corrections.join(", ")||"tidak ada; perlu Admin"));
    lines.push("Hanya sel itu yang dikecualikan dari Full Lock. Setelah lolos, otomatis terkunci kembali.");
    lines.push("LOKASI KESALAHAN:");
    for(const issue of issues.slice(0,12))lines.push(hbFormulaIssueDetail(issue));
    if(issues.length>12)
      lines.push("Sel lain ("+(issues.length-12)+"): "+
        issues.slice(12).map(x=>x.address).join(", "));
    const confidential=issues.filter(x=>HB_FORMULA_MINES.has(x.address)).length;
    lines.push("Penanda: fill HITAM + font MERAH di sel anomali."+
      (confidential?" "+confidential+" sel mines rahasia hanya ditampilkan merah saat Admin membuka semua sheet.":""));
  }else{
    lines.push("KEPUTUSAN: Rumus sesuai pola, L178:L238 kelipatan 10, tidak ada Paket/Rokok bermargin minus. Sheet baru boleh dibuat.");
  }
  if(formatFailure)
    lines.push("PERHATIAN: pengaturan warna belum berhasil/terverifikasi: "+
      (formatFailure.message||String(formatFailure))+".");
  return lines.join("\n");
}
function hbRenderFormulaSummary(result,formatFailure){
  const el=typeof document!=="undefined"?document.getElementById("formulaAuditSummary"):null;
  if(!el)return;
  const hasIssues=Boolean(result.issues&&result.issues.length);
  el.textContent=hbFormulaSummaryText(result,formatFailure);
  el.className="formula-audit-summary "+(hasIssues||formatFailure?"error":"ok");
  hbRenderFormulaCorrectionList(result);
}
function hbRenderFormulaAuditUnavailable(reason){
  const el=typeof document!=="undefined"?document.getElementById("formulaAuditSummary"):null;
  if(!el)return;
  el.className="formula-audit-summary error";
  const copies=document.getElementById("formulaCorrectionCopies");
  if(copies){copies.hidden=true;copies.replaceChildren();}
  el.textContent="PEMERIKSAAN RUMUS BELUM BERHASIL\n"+
    "Belum ada hasil pemeriksaan yang dapat dipercaya.\n"+
    "Penyebab: "+(reason&&reason.message||String(reason))+
    "\nJangan membuat sheet baru sampai pemeriksaan berhasil.";
}

async function hbScanLatestSheetFormulas(){
  const groupRanges=[
    {range:"F6:I75",fromRow:6,fromCol:hbFormulaColumnNumber("F")},
    {range:"B93:N113",fromRow:93,fromCol:hbFormulaColumnNumber("B")},
    {range:"D178:P239",fromRow:178,fromCol:hbFormulaColumnNumber("D")}
  ];
  return Excel.run(async ctx=>{
    // The only formula gate is the newest (leftmost, position 0) worksheet.
    // The user may be viewing a historical tab; it must never become
    // the formula source merely because it is selected in the UI.
    const collection=ctx.workbook.worksheets;
    collection.load("items/id,items/name,items/position");
    await ctx.sync();
    if(!collection.items.length)
      throw new Error("Workbook tidak memiliki sheet untuk diperiksa.");
    const latest=collection.items.reduce((a,b)=>Number(b.position)<Number(a.position)?b:a);
    const sh=collection.getItem(latest.id);
    sh.load("id,name,protection/protected");
    const chunks=groupRanges.map(group=>{
      const range=sh.getRange(group.range);
      range.load("formulas");
      return {...group,range};
    });
    const rokokL=sh.getRange("L178:L238");
    rokokL.load("values");
    // Match the negative-margin/black-name coloring rules exactly.
    const packetMargins=sh.getRange("A6:C75");
    packetMargins.load("values");
    const cigaretteMargins=sh.getRange("A178:E238");
    cigaretteMargins.load("values");
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
    // Audit CALCULATED values, not the L formula strings. The formula-pattern
    // guard above independently ensures L formulas were not replaced/removed.
    // Merge an invalid value into an existing formula issue at the same cell,
    // avoiding duplicate markers and preserving one address per anomaly.
    const lValues=rokokL.values;
    if(!Array.isArray(lValues)||lValues.length!==HB_ROKOK_L_CHECK_COUNT)
      throw new Error("Tidak dapat memeriksa 61 nilai L178:L238. Buat Sheet Baru DIBLOKIR demi keamanan.");
    for(let index=0;index<HB_ROKOK_L_CHECK_COUNT;index++){
      if(!Array.isArray(lValues[index])||lValues[index].length!==1)
        throw new Error("Nilai L"+(HB_ROKOK_L_FIRST+index)+" tidak terbaca. Buat Sheet Baru DIBLOKIR.");
      const actual=lValues[index][0];
      if(hbRokokLIsAllowed(actual))continue;
      const address="L"+(HB_ROKOK_L_FIRST+index);
      const existing=issues.find(x=>x.address===address);
      if(existing)existing.stock10Error=actual;
      else issues.push({address,kind:"L_MULTIPLE_10",actual});
    }
    const packet=packetMargins.values,cigarettes=cigaretteMargins.values;
    if(!Array.isArray(packet)||packet.length!==70||
       !Array.isArray(cigarettes)||cigarettes.length!==61)
      throw new Error("Pemeriksaan margin Paket/Rokok belum lengkap. Buat Sheet Baru DIBLOKIR.");
    const pushMarginIssue=info=>{
      if(!info)return;
      const original=issues.find(x=>x.address===info.address);
      if(original)original.marginIssue=info;
      else issues.push({address:info.address,kind:"NEGATIVE_MARGIN",marginIssue:info});
    };
    for(let i=0;i<packet.length;i++){
      const row=6+i,values=packet[i];
      if(!Array.isArray(values)||values.length!==3)
        throw new Error("Pemeriksaan harga Paket baris "+row+" tidak lengkap.");
      const [name,buy,sell]=values;
      pushMarginIssue(hbNegativeProductMargin(name,sell,buy,
        typeof sell==="number"&&typeof buy==="number"?sell-buy:NaN,
        "PAKET",row,"C"+row));
    }
    for(let i=0;i<cigarettes.length;i++){
      const row=178+i,values=cigarettes[i];
      if(!Array.isArray(values)||values.length!==5)
        throw new Error("Pemeriksaan harga Rokok baris "+row+" tidak lengkap.");
      const [name,sell,buy,,margin]=values;
      pushMarginIssue(hbNegativeProductMargin(name,sell,buy,margin,
        "ROKOK",row,"E"+row));
    }
    return {sheetId:sh.id,sheetName:sh.name,issues,
      checked:HB_FORMULA_REFERENCE.size,stockCheckCount:HB_ROKOK_L_CHECK_COUNT,
      marginCheckCount:131};
  });
}

function hbFormulaColorReadSafe(v){
  return typeof v==="string"&&v.trim()!==""?v:null;
}
const HB_FORMULA_VISUAL_SWEPT=new Set();
function hbIsBlackRedMarker(fill,font){
  return String(fill||"").toUpperCase()===HB_FORMULA_GUARD_BLACK&&
    String(font||"").toUpperCase()===HB_FORMULA_GUARD_RED;
}
function hbNormalMarkerColors(address,journal,current){
  const candidates=[];
  if(current&&current.cells&&current.cells[address])
    candidates.push(current.cells[address]);
  for(const source of Object.values(journal||[])){
    if(source&&source.cells&&source.cells[address])
      candidates.push(source.cells[address]);
  }
  const found=candidates.find(x=>x&&!hbIsBlackRedMarker(x.fill,x.font));
  // Unknown inherited markers: remove hardcoded error fill/font so Excel's
  // underlying cell/conditional styles can show. Never guess a neighboring
  // sheet's formula or value to restore formatting.
  return found?{fill:hbFormulaColorReadSafe(found.fill),
    font:hbFormulaColorReadSafe(found.font)}:{fill:null,font:"#000000"};
}
async function hbFindUntrackedBlackRedMarkers(sheetId){
  const addresses=[...HB_FORMULA_REFERENCE.keys()];
  for(let row=6;row<=75;row++)addresses.push("C"+row);
  for(let row=178;row<=238;row++)addresses.push("E"+row);
  const unique=[...new Set(addresses)].filter(x=>!HB_FORMULA_MINES.has(x));
  const matched=[];
  // Excel Online can reject large range-format loads. Read individual cells
  // in bounded batches ONCE per latest worksheet, not on each keystroke.
  for(let i=0;i<unique.length;i+=100){
    const group=unique.slice(i,i+100);
    const found=await Excel.run(async ctx=>{
      const sh=ctx.workbook.worksheets.getItem(sheetId);
      const cells=group.map(address=>{
        const r=sh.getRange(address);
        r.format.fill.load("color");
        r.format.font.load("color");
        return {address,r};
      });
      await ctx.sync();
      return cells.filter(({r})=>hbIsBlackRedMarker(
        r.format.fill.color,r.format.font.color)).map(x=>x.address);
    });
    matched.push(...found);
  }
  return matched;
}
async function hbPaintFormulaIssues(result){
  const previous=hbFormulaGuardJournal();
  const oldSheet=previous[result.sheetId]||{name:result.sheetName,cells:{}};
  const recorded=oldSheet.cells||{};
  const adminOpen=typeof isGlobalUnlockMode==="function"&&isGlobalUnlockMode();
  const shouldPaint=address=>!HB_FORMULA_MINES.has(address)||adminOpen;
  const active=new Set(result.issues.map(x=>x.address));
  const toStart=[...active].filter(address=>shouldPaint(address)&&!recorded[address]);
  const toRestore=Object.keys(recorded).filter(address=>!active.has(address)&&shouldPaint(address));
  let inherited=[];
  if(!HB_FORMULA_VISUAL_SWEPT.has(result.sheetId)){
    inherited=await hbFindUntrackedBlackRedMarkers(result.sheetId);
  }
  const stale=inherited.filter(address=>!active.has(address)&&
    !toRestore.includes(address)&&shouldPaint(address));
  const marked=Object.keys(recorded).filter(address=>active.has(address)&&shouldPaint(address));
  const needRead=[...new Set([...toStart,...toRestore,...marked])];
  if(!needRead.length&&!stale.length){
    HB_FORMULA_VISUAL_SWEPT.add(result.sheetId);
    return;
  }
  const inspected=needRead.length?await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(result.sheetId);
    const cells=needRead.map(address=>{
      const range=sh.getRange(address);
      range.format.fill.load("color");
      range.format.font.load("color");
      return {address,range};
    });
    await ctx.sync();
    return cells.map(({address,range})=>({
      address,fill:hbFormulaColorReadSafe(range.format.fill.color),
      font:hbFormulaColorReadSafe(range.format.font.color)
    }));
  }):[];
  const actual=new Map(inspected.map(x=>[x.address,x]));
  const repaint=marked.filter(address=>{
    const c=actual.get(address);
    return c&&!hbIsBlackRedMarker(c.fill,c.font);
  });
  const paint=[...new Set(toStart.concat(repaint))];
  // Do not alter a cell whose marker was already cleared by another module.
  const restore=[...new Set([...toRestore.filter(address=>{
    const c=actual.get(address);
    return c&&hbIsBlackRedMarker(c.fill,c.font);
  }),...stale])];
  const next={name:result.sheetName,cells:{...recorded}};
  for(const address of toStart){
    const c=actual.get(address);
    const base=c&&hbIsBlackRedMarker(c.fill,c.font)
      ?hbNormalMarkerColors(address,previous,oldSheet)
      :{fill:c?c.fill:null,font:c?c.font:"#000000"};
    next.cells[address]=base;
  }
  if(toStart.length){
    previous[result.sheetId]=next;
    setSetting(HB_FORMULA_GUARD_JOURNAL_KEY,previous);
    await saveSettings(); // Save original colors BEFORE painting.
  }
  if(paint.length||restore.length){
    await Excel.run(async ctx=>{
      const sh=ctx.workbook.worksheets.getItem(result.sheetId);
      sh.load("id,protection/protected");
      const marker=sh.getRange("A1");
      marker.format.protection.load("locked");
      await ctx.sync();
      const wasProtected=Boolean(sh.protection.protected);
      const wasFull=wasProtected&&marker.format.protection.locked===true;
      const pw=wasProtected?requireSecret():null;
      let unlocked=false,failed=null;
      try{
        if(wasProtected){
          unlocked=true;
          sh.protection.unprotect(pw);
          await ctx.sync();
        }
        for(const address of paint){
          const r=sh.getRange(address);
          r.format.fill.color=HB_FORMULA_GUARD_BLACK;
          r.format.font.color=HB_FORMULA_GUARD_RED;
        }
        for(const address of restore){
          const r=sh.getRange(address);
          const original=next.cells[address]||
            hbNormalMarkerColors(address,previous,oldSheet);
          if(original.fill)r.format.fill.color=original.fill;
          else r.format.fill.clear();
          r.format.font.color=original.font||"#000000";
        }
        await ctx.sync();
      }catch(e){failed=e;}
      finally{
        if(unlocked){
          try{
            if(wasFull)fullLock(sh,pw);
            else workingLock(sh,pw);
            await ctx.sync();
          }catch(e){
            throw new Error("Proteksi gagal dipulihkan setelah pembaruan warna: "+
              (e.message||String(e)));
          }
        }
      }
      if(failed)throw failed;
    });
  }
  if(paint.length||restore.length){
    // Excel may acknowledge a format operation before another loaded add-in
    // overrides it. Read each changed CELL again; never report a clean sheet
    // while a restored black/red alert still remains on it.
    const confirmation=await Excel.run(async ctx=>{
      const sh=ctx.workbook.worksheets.getItem(result.sheetId);
      const changes=[...new Set([...paint,...restore])].map(address=>{
        const cell=sh.getRange(address);
        cell.format.fill.load("color");
        cell.format.font.load("color");
        return {address,cell};
      });
      await ctx.sync();
      return changes.map(({address,cell})=>({
        address,marked:hbIsBlackRedMarker(
          cell.format.fill.color,cell.format.font.color)
      }));
    });
    const paints=new Set(paint),restores=new Set(restore);
    const unexpected=confirmation.filter(x=>
      (paints.has(x.address)&&!x.marked)||
      (restores.has(x.address)&&x.marked));
    if(unexpected.length)
      throw new Error("Perubahan warna anomali belum tersimpan di: "+
        unexpected.map(x=>x.address).join(", ")+
        ". Tekan PERIKSA ULANG ANOMALI & WARNA.");
  }
  for(const address of toRestore)delete next.cells[address];
  if(Object.keys(next.cells).length)previous[result.sheetId]=next;
  else delete previous[result.sheetId];
  if(toRestore.length){
    setSetting(HB_FORMULA_GUARD_JOURNAL_KEY,previous);
    await saveSettings();
  }
  HB_FORMULA_VISUAL_SWEPT.add(result.sheetId);
}

// Grant a worksheet-level exception ONLY to cells needed for correcting
// currently detected issues. The worksheet itself stays protected; every
// other cell retains the previously applied Full Lock / Working Lock rules.
// Original cell locked/formulaHidden flags are journaled per sheet so that
// fixing the issue can restore the precise prior protection.
async function hbSyncCorrectionOnlyAccess(result){
  if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)return;
  if(typeof isGlobalUnlockMode==="function"&&isGlobalUnlockMode())return;
  // Refresh workbook-backed close intent, not just this browser's cached
  // settings. Failure to refresh must fail closed instead of granting access.
  if(typeof refreshShiftSettingsFromWorkbook==="function")
    await refreshShiftSettingsFromWorkbook();
  if(typeof isGlobalUnlockMode==="function"&&isGlobalUnlockMode())return;
  // Explicit employee/admin close is FINAL: formula/margin corrections cannot
  // silently turn a deliberately closed shift back into an editable sheet.
  if(typeof hbIsExplicitlyClosedShift==="function"&&
     hbIsExplicitlyClosedShift(result.sheetId))return;
  const desired=hbCorrectionTargets(result.issues);
  const desiredSet=new Set(desired);
  const journal=hbCorrectionAccessJournal();
  const existing=journal[result.sheetId]||{name:result.sheetName,cells:{}};
  const recorded=existing.cells||{};
  const inspected=[...new Set([...desired,...Object.keys(recorded)])];
  if(!inspected.length)return;
  // Read the actual locking state every time; another Excel Online device
  // may have run Full Lock after this task pane last audited the sheet.
  const before=await Excel.run(async ctx=>{
    const sheets=ctx.workbook.worksheets;
    sheets.load("items/id,items/position");
    await ctx.sync();
    if(!sheets.items.length)return null;
    const latest=sheets.items.reduce((a,b)=>
      Number(a.position)<Number(b.position)?a:b);
    if(String(latest.id)!==String(result.sheetId))
      throw new Error("Sheet terbaru berubah saat memberi akses koreksi; pemeriksaan harus diulang.");
    const sh=sheets.getItem(result.sheetId);
    sh.load("id,protection/protected");
    const marker=sh.getRange("A1");
    marker.format.protection.load("locked");
    const ranges=inspected.map(address=>{
      const r=sh.getRange(address);
      r.format.protection.load("locked,formulaHidden");
      return {address,r};
    });
    await ctx.sync();
    return {
      protectedBefore:Boolean(sh.protection.protected),
      fullLock:marker.format.protection.locked===true,
      cells:ranges.map(({address,r})=>({
        address,
        locked:r.format.protection.locked===true,
        formulaHidden:r.format.protection.formulaHidden===true
      }))
    };
  });
  if(!before)return;
  const read=new Map(before.cells.map(x=>[x.address,x]));
  const nextCells={};
  // On a first exception preserve the original protection. On later audits,
  // retain that original snapshot even if the cell is currently editable.
  for(const address of desired){
    nextCells[address]=recorded[address]||{
      locked:read.get(address).locked,
      formulaHidden:read.get(address).formulaHidden
    };
  }
  const changes=[];
  for(const address of inspected){
    const current=read.get(address),old=recorded[address];
    if(desiredSet.has(address)){
      if(current.locked||current.formulaHidden)
        changes.push({address,locked:false,formulaHidden:false});
    }else if(old &&
             (current.locked!==Boolean(old.locked) ||
              current.formulaHidden!==Boolean(old.formulaHidden))){
      changes.push({address,locked:Boolean(old.locked),
        formulaHidden:Boolean(old.formulaHidden)});
    }
  }
  // Store the journal before any edit, allowing the original protected
  // attributes to be recovered after an Excel save/reload interruption.
  const oldSerialized=JSON.stringify(recorded);
  const newSerialized=JSON.stringify(nextCells);
  if(oldSerialized!==newSerialized){
    if(Object.keys(nextCells).length)
      journal[result.sheetId]={name:result.sheetName,cells:nextCells};
    else delete journal[result.sheetId];
    setSetting(HB_CORRECTION_ACCESS_KEY,journal);
    await saveSettings();
  }
  if(!changes.length)return;
  await Excel.run(async ctx=>{
    const sheets=ctx.workbook.worksheets;
    sheets.load("items/id,items/position");
    await ctx.sync();
    const latest=sheets.items.reduce((a,b)=>
      Number(a.position)<Number(b.position)?a:b);
    if(String(latest.id)!==String(result.sheetId))
      throw new Error("Sheet berubah sebelum akses koreksi diterapkan.");
    const sh=sheets.getItem(result.sheetId);
    sh.load("id,protection/protected");
    const marker=sh.getRange("A1");
    marker.format.protection.load("locked");
    await ctx.sync();
    const originallyProtected=Boolean(sh.protection.protected);
    const wasFull=marker.format.protection.locked===true;
    const pw=originallyProtected?requireSecret():null;
    let startedUnprotect=false,failed=null;
    try{
      if(originallyProtected){
        startedUnprotect=true;
        sh.protection.unprotect(pw);
        await ctx.sync();
      }
      for(const item of changes){
        const correction=sh.getRange(item.address);
        correction.format.protection.locked=item.locked;
        correction.format.protection.formulaHidden=item.formulaHidden;
      }
      await ctx.sync();
    }catch(e){failed=e;}
    finally{
      if(startedUnprotect){
        try{
          // Do not call fullLock/workingLock: both reset protection on many
          // unrelated cells. Reprotect without changing any other cell flags.
          sh.protection.protect({
            selectionMode:(wasFull&&desired.length===0)
              ?Excel.ProtectionSelectionMode.none
              :Excel.ProtectionSelectionMode.unlocked,
            allowFormatCells:false,allowFormatColumns:false,
            allowFormatRows:false,allowInsertColumns:false,
            allowInsertRows:false,allowDeleteColumns:false,
            allowDeleteRows:false
          },pw);
          await ctx.sync();
        }catch(e){
          // Never leave the whole worksheet unprotected if Office rejects
          // the narrow correction-mode protection options.
          let fallbackError=null;
          try{
            fullLock(sh,pw);
            await ctx.sync();
          }catch(restoreError){fallbackError=restoreError;}
          throw new Error("Gagal mengunci ulang sheet setelah memberi akses koreksi: "+
            (e&&e.message||String(e))+
            (fallbackError?"; Full Lock darurat juga gagal: "+
              (fallbackError.message||String(fallbackError)):"; Full Lock darurat diterapkan."));
        }
      }
    }
    if(failed)throw failed;
  });
}

let hbFormulaRunning=null;
let hbFormulaAuditBusy=false;
let hbFormulaWatchedSheetId=null;
let hbFormulaChangeTimer=null;
// Worksheet.onChanged is supported by ExcelApi 1.7 (the current manifest).
// Recheck only after a committed edit; no recurring per-minute workbook writes.
async function hbWatchLatestFormulaCorrections(id){
  if(hbFormulaWatchedSheetId===String(id))return;
  const nextId=String(id);
  await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(nextId);
    if(!sh.onChanged||typeof sh.onChanged.add!=="function")return;
    sh.onChanged.add(event=>{
      if(hbFormulaWatchedSheetId!==nextId)return;
      if(hbFormulaChangeTimer)clearTimeout(hbFormulaChangeTimer);
      hbFormulaChangeTimer=setTimeout(()=>{
        hbFormulaChangeTimer=null;
        const busy=
          (typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
          (typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress)||
          (typeof adminGlobalUnmaskBusy!=="undefined"&&adminGlobalUnmaskBusy)||
          (typeof minesRecoveryBusy!=="undefined"&&minesRecoveryBusy)||
          (typeof hbMarginColorBusy!=="undefined"&&hbMarginColorBusy);
        if(busy)return;
        hbVerifySourceFormulas(false).catch(e=>hbRenderFormulaAuditUnavailable(e));
      },2400);
    });
    await ctx.sync();
    hbFormulaWatchedSheetId=nextId;
  });
}
function hbClearStaleFormulaBlockMessage(){
  if(typeof document==="undefined")return;
  const panel=document.getElementById("securityStatus");
  if(panel&&/PEMBUATAN SHEET DIBLOKIR/.test(panel.textContent||"")&&
     typeof setSecurityStatus==="function")
    setSecurityStatus("PEMERIKSAAN ULANG LOLOS — anomali telah diperbaiki. Warna peringatan dipulihkan.",false);
  const form=document.getElementById("status");
  if(form&&/PEMBUATAN SHEET DIBLOKIR/.test(form.textContent||""))
    form.textContent="Anomali sudah diperbaiki; pemeriksaan rumus dan stok sekarang lolos.";
}
async function hbVerifySourceFormulas(blockCreation=false){
  if(hbFormulaRunning){
    try{await hbFormulaRunning;}catch(_){
      // A preceding blocked preflight does not prevent a fresh rescan
      // after the operator repairs an incorrect cell.
    }
  }
  hbFormulaAuditBusy=true;
  const run=(async()=>{
    const result=await hbScanLatestSheetFormulas();
    // Event subscription is best-effort and must never block preflight.
    try{await hbWatchLatestFormulaCorrections(result.sheetId);}catch(_){}
    let formatFailure=null,markerCleanupFailure=null;
    try{await hbPaintFormulaIssues(result);}
    catch(e){markerCleanupFailure=e;formatFailure=e;}
    // Cigarette NAME colors use static base fills (unlike dynamic Paket
    // formatting); they must be refreshed after a negative E margin is fixed.
    if(typeof hbMarginRefreshLatestRokokNames==="function"){
      try{
        await hbMarginRefreshLatestRokokNames(result.sheetId);
      }catch(e){
        formatFailure=formatFailure?
          new Error(formatFailure.message+"; perbaikan warna Rokok: "+(e.message||e)):
          new Error("Pemulihan warna Rokok: "+(e.message||e));
      }
    }
    // Open ONLY the exact correction input cells on the protected latest
    // sheet, then restore those original locks when its anomalies disappear.
    // Access errors must be visible and must never silently grant Full Access.
    let accessFailure=null;
    try{await hbSyncCorrectionOnlyAccess(result);}catch(e){accessFailure=e;}
    const warning=accessFailure
      ?new Error((formatFailure?formatFailure.message+"; ":"")+
        "akses koreksi terbatas: "+(accessFailure.message||String(accessFailure)))
      :formatFailure;
    // Keep this visible regardless of later workbook-protection status messages.
    hbRenderFormulaSummary(result,warning);
    if(result.issues.length){
      const msg=hbFormulaGuardError(result.issues,result.sheetName);
      const notice=msg+(warning?
        "\nPERINGATAN: Sebagian warna / akses koreksi belum berhasil ("+warning.message+").":"");
      hbShowFormulaNote(notice);
      if(blockCreation)throw new Error(notice);
    }else{
      // A new shift inherits colors from the source. If legacy black/red
      // cleanup failed, block cloning rather than copying obsolete alerts
      // and confusing the next shift's employee.
      if(markerCleanupFailure&&blockCreation)
        throw new Error("PEMBUATAN SHEET DIBLOKIR: penanda hitam-merah lama "+
          "belum berhasil diperiksa/dipulihkan pada sheet terbaru. "+
          (markerCleanupFailure.message||String(markerCleanupFailure))+
          ". Gunakan PERIKSA ULANG ANOMALI & WARNA.");
      if(!warning)hbClearStaleFormulaBlockMessage();
      if(accessFailure&&blockCreation)
        throw new Error("Pembuatan sheet DIBLOKIR: akses koreksi sebelumnya belum bisa dikunci kembali. "+
          (accessFailure.message||String(accessFailure)));
    }
    if(!result.issues.length&&warning){
      // Valid source formulas must not be blocked merely because an old
      // warning color could not be restored by Excel.
      if(typeof setSecurityStatus==="function")
        setSecurityStatus("Rumus sumber BENAR, tetapi pemulihan warna peringatan belum selesai: "+
          warning.message,true);
    }
    return result;
  })();
  hbFormulaRunning=run;
  try{return await run;}finally{
    if(hbFormulaRunning===run)hbFormulaRunning=null;
    hbFormulaAuditBusy=false;
  }
}

Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel)return;
  const recheck=document.getElementById("formulaRecheckBtn");
  if(recheck)recheck.addEventListener("click",async()=>{
    if(hbFormulaAuditBusy)return;
    recheck.disabled=true;
    try{
      const result=await hbVerifySourceFormulas(false);
      if(typeof setSecurityStatus==="function"&&!result.issues.length)
        setSecurityStatus("PEMERIKSAAN ULANG LOLOS — semua penanda hitam-merah lama yang ditemukan telah dipulihkan.",false);
    }catch(e){
      hbRenderFormulaAuditUnavailable(e);
      if(typeof setSecurityStatus==="function")
        setSecurityStatus("Pemeriksaan ulang belum selesai: "+(e.message||String(e)),true);
    }finally{recheck.disabled=false;}
  });
  let attempts=0;
  const passiveAudit=()=>{
    if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)return;
    // Background auditing must not repaint cells while Excel is saving a
    // newly created worksheet, stock transfers, admin unlocking, or historical
    // margin/mines repair. The create button still runs a strict preflight.
    const busy=
      (typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
      (typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress)||
      (typeof adminGlobalUnmaskBusy!=="undefined"&&adminGlobalUnmaskBusy)||
      (typeof minesRecoveryBusy!=="undefined"&&minesRecoveryBusy)||
      (typeof hbMarginColorBusy!=="undefined"&&hbMarginColorBusy)||
      (typeof hbHistoricalLockAuditBusy!=="undefined"&&hbHistoricalLockAuditBusy)||
      Boolean(hbFormulaRunning);
    if(busy){
      if(++attempts<=8)setTimeout(passiveAudit,10000);
      else hbRenderFormulaAuditUnavailable(new Error(
        "Excel sedang menyimpan/memperbarui sheet; periksa lagi dengan membuka ulang add-in, atau saat membuat shift berikutnya."));
      return;
    }
    hbVerifySourceFormulas(false).catch(e=>{
      hbRenderFormulaAuditUnavailable(e);
      if(typeof setSecurityStatus==="function")
        setSecurityStatus("Pengaman rumus belum dapat diperiksa: "+e.message,true);
    });
  };
  // One delayed startup audit, never a recurring background write every
  // minute. The next shift is always validated immediately before cloning.
  setTimeout(passiveAudit,11000);
  // A newly created shift becomes the newest tab. Attach its own onChanged
  // watcher only AFTER creation settles so future corrections can automatically
  // restore black/red flags and cigarette margin colors in that new tab.
  Excel.run(async ctx=>{
    const sheets=ctx.workbook.worksheets;
    if(sheets.onAdded&&typeof sheets.onAdded.add==="function"){
      sheets.onAdded.add(()=>setTimeout(passiveAudit,15000));
      await ctx.sync();
    }
  }).catch(()=>{}); // Optional Excel event support must not block creating sheets.
});
