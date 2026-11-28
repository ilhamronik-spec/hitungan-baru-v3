"use strict";
/*
 * Koreksi Belanja — independent opt-in V4 module.
 *
 * Existing create/buy/transfer/protection/formula handlers remain untouched.
 * The last worksheet with evidence of purchase is located by:
 *   Paket: D(current opening) - E(previous ending) > 0
 *   Rokok: O(current purchase quantity) > 0.
 * No assumed "purchase quantity" is read from display transfer F.
 *
 * Only D6:D75 / O178:O238 on the target, plus M178:M238 on newer sheets
 * if the existing warehouse carry-forward chain is EXACT, may be written.
 * All workbook balance formulas stay intact and recalculate themselves.
 */
const HB_PC_BUILD="6.3.9";
const HB_PC_JOURNAL_KEY="hitunganV4.purchaseCorrectionHistoryV1";
const HB_PC_MAX_SHEETS=180;
const hbPcState={loaded:null,working:false,showAll:false,filter:""};
const hbPcDom=id=>document.getElementById(id);
const hbPcNum=v=>typeof v==="number"&&Number.isSafeInteger(v)?v:NaN;
const hbPcId=x=>String(x??"");
const hbPcRp=n=>new Intl.NumberFormat("id-ID",{maximumFractionDigits:0}).format(Number(n)||0);
function hbPcStatus(message,error=false){
  const el=hbPcDom("purchaseCorrectionStatus");
  if(!el)return;
  el.textContent=message;
  el.style.color=error?"#b42318":"#334155";
  el.style.whiteSpace="pre-wrap";
}
function hbPcJournal(){
  const raw=typeof getSetting==="function"?getSetting(HB_PC_JOURNAL_KEY):null;
  return Array.isArray(raw)?raw.filter(x=>x&&typeof x==="object"):[];
}
async function hbPcSaveJournal(entry){
  const all=hbPcJournal().filter(x=>x.id!==entry.id);
  // The Office document-settings store is small. Pending writes need the
  // full cell journal for idempotent recovery; verified entries keep a
  // compact audit summary, not a huge historical copy of every cell.
  const record=entry.status==="pending"?entry:{
    id:entry.id,status:entry.status,targetId:entry.targetId,
    targetName:entry.targetName,createdAt:entry.createdAt,
    verifiedAt:entry.verifiedAt||0,rows:entry.rows||[],
    cellCount:entry.cells?.length||entry.cellCount||0,
    sheetCount:entry.cells
      ?new Set(entry.cells.map(x=>x.sheetId)).size:(entry.sheetCount||0)
  };
  all.push(record);
  setSetting(HB_PC_JOURNAL_KEY,all.slice(-8));
  await saveSettings();
}
function hbPcPrevious(sheetItems,target){
  const index=sheetItems.findIndex(x=>hbPcId(x.id)===hbPcId(target.id));
  return index<0?null:sheetItems[index+1]||null;
}
function hbPcInt(v){
  const n=hbPcNum(v);
  return Number.isFinite(n)&&n>=0?n:NaN;
}
function hbPcCompare(a,b){return Number.isFinite(a)&&Number.isFinite(b)&&a===b;}
function hbPcItemsFromRanges(targetId,targetName,previousId,previousName,packet,previousEnding,cigarette){
  const items=[];
  for(let i=0;i<70;i++){
    const row=6+i,a=packet[i],v=previousEnding[i][0];
    const name=String(a[0]??"").trim();
    if(!name)continue;
    const stock=hbPcInt(a[3]),ending=hbPcInt(a[4]),previous=hbPcInt(v);
    if(!Number.isFinite(stock)||!Number.isFinite(ending)||!Number.isFinite(previous))continue;
    const qty=stock-previous;
    // Negative carry differences are not proof of a purchase. Do not
    // expose them as editable "quantity purchased" inputs.
    if(!Number.isSafeInteger(qty)||qty<0)continue;
    items.push({group:"paket",row,name,address:"D"+row,previous,opening:stock,ending,
      originalQty:qty,cost:Number(a[1])||0,qtyText:String(Math.max(0,qty))});
  }
  for(let i=0;i<61;i++){
    const row=178+i,a=cigarette[i],name=String(a[0]??"").trim();
    if(!name)continue;
    const qty=hbPcInt(a[14]);
    if(!Number.isFinite(qty))continue;
    items.push({group:"rokok",row,name,address:"O"+row,
      originalQty:qty,cost:Number(a[2])||0,
      warehouseOpening:hbPcInt(a[12]),warehouseEnding:hbPcInt(a[13]),
      transferred:Number(a[11])||0,qtyText:String(qty)});
  }
  return {targetId,targetName,previousId,previousName,items,
    hasPurchase:items.some(x=>x.originalQty>0)};
}
// Read-only scan. Never use the currently ACTIVE tab as a shortcut;
// the most recent sheet with actual purchase evidence may be older.
async function hbPcInspectLastPurchased(){
  return Excel.run(async ctx=>{
    const collection=ctx.workbook.worksheets;
    collection.load("items/id,items/name,items/position");
    await ctx.sync();
    const sorted=[...collection.items].sort((a,b)=>Number(a.position)-Number(b.position));
    if(sorted.length<2)throw new Error("Perlu sedikitnya dua shift untuk membuktikan riwayat belanja Paket.");
    if(sorted.length>HB_PC_MAX_SHEETS)
      throw new Error("Workbook memiliki terlalu banyak shift untuk audit koreksi aman; hubungi Admin.");
    const comparisons=[];
    for(let i=0;i<sorted.length-1;i++){
      const sh=collection.getItem(sorted[i].id);
      const prior=collection.getItem(sorted[i+1].id);
      const d=sh.getRange("D6:D75"),e=prior.getRange("E6:E75"),o=sh.getRange("O178:O238");
      d.load("values");e.load("values");o.load("values");
      comparisons.push({current:sorted[i],prior:sorted[i+1],d,e,o});
    }
    await ctx.sync();
    let selected=null;
    for(const entry of comparisons){
      let evidence=false;
      for(let i=0;i<70&&!evidence;i++){
        const d=hbPcNum(entry.d.values[i][0]),e=hbPcNum(entry.e.values[i][0]);
        if(Number.isFinite(d)&&Number.isFinite(e)&&d>e)evidence=true;
      }
      for(let i=0;i<61&&!evidence;i++)
        if(hbPcNum(entry.o.values[i][0])>0)evidence=true;
      if(evidence){selected=entry;break;}
    }
    if(!selected)return null;
    const sh=collection.getItem(selected.current.id),prior=collection.getItem(selected.prior.id);
    const packet=sh.getRange("A6:E75"),previousEnding=prior.getRange("E6:E75"),
      cigarette=sh.getRange("A178:O238");
    packet.load("values");packet.load("formulas");
    previousEnding.load("values");cigarette.load("values");cigarette.load("formulas");
    await ctx.sync();
    const result=hbPcItemsFromRanges(selected.current.id,selected.current.name,
      selected.prior.id,selected.prior.name,packet.values,previousEnding.values,cigarette.values);
    // Source purchase cells must be literal inputs, not formula cells.
    result.items=result.items.filter(item=>{
      const f=item.group==="paket"?packet.formulas[item.row-6][3]:
        cigarette.formulas[item.row-178][14];
      return !(typeof f==="string"&&f.startsWith("="));
    });
    result.newerSheetIds=sorted.slice(0,sorted.findIndex(x=>x.id===selected.current.id)).map(x=>x.id);
    return result;
  });
}
function hbPcRender(){
  const el=hbPcDom("purchaseCorrectionRows");
  if(!el)return;
  el.replaceChildren();
  const current=hbPcState.loaded;
  if(!current)return;
  const show=hbPcState.showAll,query=hbPcState.filter.toLowerCase();
  let displayed=0;
  for(const item of current.items){
    if(!show&&item.originalQty<=0&&item.qtyText==="0")continue;
    if(query&&!item.name.toLowerCase().includes(query))continue;
    displayed++;
    const row=document.createElement("div");
    row.className="purchase-correction-row";
    const title=document.createElement("div");
    title.className="purchase-correction-title";
    title.textContent=(item.group==="paket"?"PAKET":"ROKOK")+" | "+item.name;
    const note=document.createElement("small");
    note.textContent="Baris "+item.row+" • jumlah belanja tercatat: "+item.originalQty+
      " • harga dasar Rp"+hbPcRp(item.cost);
    const label=document.createElement("label");
    label.textContent="Jumlah belanja yang benar (pcs)";
    const input=document.createElement("input");
    input.type="number";input.min="0";input.step="1";input.inputMode="numeric";
    input.value=item.qtyText;input.dataset.group=item.group;input.dataset.row=String(item.row);
    input.addEventListener("input",()=>{
      item.qtyText=input.value;
      hbPcRefreshDiff();
    });
    label.appendChild(input);row.append(title,note,label);el.appendChild(row);
  }
  if(!displayed){
    const empty=document.createElement("p");
    empty.className="security-note";
    empty.textContent="Tidak ada barang pada filter ini. Pilih 'Tampilkan semua barang' untuk menambahkan belanja yang belum tercatat.";
    el.appendChild(empty);
  }
  hbPcRefreshDiff();
}
function hbPcChanges(){
  const obj=hbPcState.loaded;
  if(!obj)return [];
  const changes=[];
  for(const item of obj.items){
    const raw=String(item.qtyText??"").trim();
    if(!/^\d+$/.test(raw))throw new Error("Jumlah "+item.name+" harus angka bulat mulai 0.");
    const next=Number(raw);
    if(!Number.isSafeInteger(next))throw new Error("Jumlah "+item.name+" terlalu besar.");
    if(next!==item.originalQty)changes.push({group:item.group,row:item.row,name:item.name,
      oldQty:item.originalQty,newQty:next,delta:next-item.originalQty});
  }
  return changes;
}
function hbPcRefreshDiff(){
  const element=hbPcDom("purchaseCorrectionPreview");
  if(!element)return;
  try{
    const changes=hbPcChanges();
    element.textContent=changes.length
      ?changes.length+" jenis akan dikoreksi: "+changes.slice(0,12).map(x=>
        x.name+" ("+x.oldQty+" → "+x.newQty+" pcs)").join("; ")+
        (changes.length>12?"; dan "+(changes.length-12)+" lainnya":"")+
        ". Rumus stok, modal dan balance akan dihitung ulang."
      :"Belum ada perubahan jumlah belanja.";
  }catch(e){element.textContent=e.message;}
}
function hbPcFindItem(target,change){
  const item=target.items.find(x=>x.group===change.group&&x.row===change.row);
  if(!item||item.name!==change.name||item.originalQty!==change.oldQty)
    throw new Error("Riwayat "+change.name+" telah berubah. Muat ulang Koreksi Belanja.");
  return item;
}
async function hbPcPlan(changes,expectedTarget){
  const target=await hbPcInspectLastPurchased();
  if(!target||hbPcId(target.targetId)!==hbPcId(expectedTarget))
    throw new Error("Sheet belanja terakhir sudah berubah. Muat ulang daftar sebelum mengoreksi.");
  const cells=[],adjustments=[];
  const add=(sheetId,address,before,after)=>{
    if(!Number.isSafeInteger(before)||!Number.isSafeInteger(after))
      throw new Error("Nilai stok tidak valid di "+address+". Koreksi dibatalkan.");
    if(before!==after)cells.push({sheetId:hbPcId(sheetId),address,before,after});
  };
  for(const change of changes){
    const item=hbPcFindItem(target,change);
    if(change.group==="paket"){
      const nextOpening=item.previous+change.newQty;
      if(nextOpening<item.ending)
        throw new Error("Paket "+item.name+": stok awal baru "+nextOpening+
          " lebih kecil dari stok akhir terhitung "+item.ending+
          ". Perbaiki jumlah koreksi atau periksa stok akhir terlebih dahulu.");
      add(target.targetId,"D"+item.row,item.opening,nextOpening);
    }else{
      if(item.warehouseEnding+change.delta<0)
        throw new Error("Rokok "+item.name+": stok gudang akhir akan minus.");
      add(target.targetId,"O"+item.row,item.originalQty,change.newQty);
      adjustments.push({row:item.row,name:item.name,delta:change.delta});
    }
  }
  // Newer warehouse openings M must continue to equal previous ending N.
  // If ANY link was manually modified, refuse rather than invent a rollover.
  if(adjustments.length&&target.newerSheetIds.length){
    const chain=await Excel.run(async ctx=>{
      const ordered=[target.targetId,...target.newerSheetIds.slice().reverse()];
      const readings=[];
      for(const id of ordered){
        const sh=ctx.workbook.worksheets.getItem(id);
        const probes=adjustments.map(x=>{
          const a=sh.getRange("A"+x.row),m=sh.getRange("M"+x.row),
            n=sh.getRange("N"+x.row);
          a.load("values");m.load("values");n.load("values");m.load("formulas");
          return {row:x.row,a,m,n};
        });
        readings.push({id,probes});
      }
      await ctx.sync();
      return readings.map(x=>({id:x.id,rows:x.probes.map(p=>({
        row:p.row,name:String(p.a.values[0][0]??"").trim(),
        m:hbPcNum(p.m.values[0][0]),n:hbPcNum(p.n.values[0][0]),
        mFormula:p.m.formulas[0][0]
      }))}));
    });
    for(const item of adjustments){
      for(let j=1;j<chain.length;j++){
        const older=chain[j-1].rows.find(x=>x.row===item.row);
        const newer=chain[j].rows.find(x=>x.row===item.row);
        if(!older||!newer||newer.name!==item.name||
           !hbPcCompare(older.n,newer.m)||
           (typeof newer.mFormula==="string"&&newer.mFormula.startsWith("=")))
          throw new Error("Koreksi Rokok "+item.name+" DIBLOKIR: stok gudang antarshift "+
            "tidak tersambung di baris "+item.row+" (shift berikutnya). "+
            "Tidak ada data yang diubah; perlu pemeriksaan Admin.");
        if(newer.m+item.delta<0||newer.n+item.delta<0)
          throw new Error("Stok gudang Rokok "+item.name+" akan minus pada shift lanjutan.");
        add(chain[j].id,"M"+item.row,newer.m,newer.m+item.delta);
      }
    }
  }
  // Balance and warehouse formulas MUST exist BEFORE changing any stock.
  // This prevents a "successful" quantity edit that cannot update totals.
  if(cells.length){
    const affected=[...new Set(cells.map(x=>x.sheetId))];
    await Excel.run(async ctx=>{
      const refs=[];
      for(const id of affected){
        const sh=ctx.workbook.worksheets.getItem(id);
        for(const address of ["G275","D275"]){
          const r=sh.getRange(address);r.load("formulas");
          refs.push({id,address,r});
        }
        for(const a of adjustments){
          const address="N"+a.row;
          const r=sh.getRange(address);r.load("formulas");
          refs.push({id,address,r});
        }
      }
      await ctx.sync();
      for(const {id,address,r} of refs){
        if(!String(r.formulas[0][0]||"").startsWith("="))
          throw new Error("Rumus "+address+" pada sheet terkait hilang. "+
            "Koreksi BELUM DILAKUKAN agar balance/stok tidak salah; perlu Admin.");
      }
    });
  }
  if(cells.length>400)
    throw new Error("Koreksi menyentuh lebih dari 400 sel stok lintas shift. "+
      "Dibatalkan agar jurnal tidak terlalu besar; perlu rekonsiliasi Admin.");
  return {targetId:hbPcId(target.targetId),targetName:target.targetName,
    previousId:hbPcId(target.previousId),newerSheetIds:target.newerSheetIds.map(hbPcId),
    rows:changes,cells,createdAt:Date.now()};
}
function hbPcBusy(){
  return hbPcState.working||
    (typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
    (typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress)||
    (typeof hbFormulaAuditBusy!=="undefined"&&hbFormulaAuditBusy)||
    (typeof hbHistoricalLockAuditBusy!=="undefined"&&hbHistoricalLockAuditBusy)||
    (typeof minesRecoveryBusy!=="undefined"&&minesRecoveryBusy)||
    (typeof hbMarginColorBusy!=="undefined"&&hbMarginColorBusy);
}
async function hbPcApplyCells(plan){
  const pw=requireSecret();
  // One synchronized Excel batch. First validate ALL values, then write.
  await Excel.run(async ctx=>{
    // Guard against a newly created shift (on this or another device) since
    // preview. Never silently mutate a no-longer-latest purchase record.
    const collection=ctx.workbook.worksheets;
    collection.load("items/id,items/position");
    await ctx.sync();
    const sorted=[...collection.items].sort((a,b)=>Number(a.position)-Number(b.position));
    const selectedIndex=sorted.findIndex(x=>hbPcId(x.id)===plan.targetId);
    if(selectedIndex<0||!sorted[selectedIndex+1]||
       hbPcId(sorted[selectedIndex+1].id)!==plan.previousId||
       JSON.stringify(sorted.slice(0,selectedIndex).map(x=>hbPcId(x.id)))!==
       JSON.stringify(plan.newerSheetIds))
      throw new Error("Susunan shift berubah saat koreksi; tidak ada stok yang diubah.");
    const sheetIds=[...new Set(plan.cells.map(x=>x.sheetId))];
    const sheets=sheetIds.map(id=>{
      const sh=ctx.workbook.worksheets.getItem(id);
      sh.load("id,name,protection/protected");
      const a1=sh.getRange("A1");a1.format.protection.load("locked");
      const changes=plan.cells.filter(x=>x.sheetId===id).map(x=>{
        const cell=sh.getRange(x.address);
        cell.load("values");cell.load("formulas");
        return {...x,cell};
      });
      return {sh,a1,changes,unlocked:false};
    });
    await ctx.sync();
    for(const entry of sheets){
      if(!entry.sh.protection.protected)
        throw new Error("Sheet "+entry.sh.name+" sedang tidak terproteksi. "+
          "Minta Admin sinkronkan proteksi sebelum koreksi.");
      for(const p of entry.changes){
        const actual=hbPcNum(p.cell.values[0][0]);
        if(actual!==p.before&&actual!==p.after)
          throw new Error("Stok di "+entry.sh.name+"!"+p.address+
            " berubah dari hasil pemeriksaan. Koreksi dibatalkan.");
        if(typeof p.cell.formulas[0][0]==="string"&&
           p.cell.formulas[0][0].startsWith("="))
          throw new Error("Sel "+p.address+" merupakan rumus; koreksi ditolak.");
      }
    }
    let failure=null;
    try{
      for(const entry of sheets){
        if(entry.sh.protection.protected){
          entry.unlocked=true;
          entry.sh.protection.unprotect(pw);
        }
      }
      await ctx.sync();
      for(const entry of sheets)
        for(const item of entry.changes)
          if(hbPcNum(item.cell.values[0][0])!==item.after)
            item.cell.values=[[item.after]];
      await ctx.sync();
      for(const entry of sheets)entry.sh.calculate(true);
      await ctx.sync();
    }catch(e){failure=e;}
    finally{
      for(const entry of sheets){
        if(!entry.unlocked)continue;
        try{
          // Restore the same protection mode WITHOUT rewriting any cell's
          // original locked/formulaHidden flags, formulas, fills or stock.
          const mode=entry.a1.format.protection.locked===true
            ?Excel.ProtectionSelectionMode.none
            :Excel.ProtectionSelectionMode.unlocked;
          entry.sh.protection.protect({
            selectionMode:mode,allowFormatCells:false,allowFormatColumns:false,
            allowFormatRows:false,allowInsertColumns:false,allowInsertRows:false,
            allowDeleteColumns:false,allowDeleteRows:false
          },pw);
          await ctx.sync();
        }catch(e){
          throw new Error("Koreksi dihentikan: gagal memulihkan proteksi sheet "+
            entry.sh.name+"; "+(e.message||String(e))+
            ". Jangan mengulang koreksi sebelum diperiksa Admin.");
        }
      }
    }
    if(failure)throw failure;
  });
}
async function hbPcVerifyPlan(plan){
  return Excel.run(async ctx=>{
    const ids=[...new Set(plan.cells.map(x=>x.sheetId))];
    const reads=plan.cells.map(p=>{
      const r=ctx.workbook.worksheets.getItem(p.sheetId).getRange(p.address);
      r.load("values");return {p,r};
    });
    const formulas=[];
    for(const id of ids){
      const sh=ctx.workbook.worksheets.getItem(id);
      const balance=sh.getRange("G275"),cash=sh.getRange("D275");
      balance.load("formulas");cash.load("formulas");
      balance.load("values");cash.load("values");
      formulas.push({sh,balance,cash});
    }
    await ctx.sync();
    const failed=reads.filter(({p,r})=>hbPcNum(r.values[0][0])!==p.after);
    if(failed.length)throw new Error("Koreksi belum terverifikasi: "+
      failed.slice(0,7).map(x=>x.p.address).join(", ")+
      ". Jangan mengulang input; periksa status koreksi.");
    for(const {sh,balance,cash} of formulas){
      if(!String(balance.formulas[0][0]||"").startsWith("=")||
         !String(cash.formulas[0][0]||"").startsWith("="))
        throw new Error("Rumus balance/modal hilang pada sheet "+sh.name+
          ". Perubahan stok tersimpan tetapi balance perlu Admin.");
      if(typeof balance.values[0][0]!=="number"||
         !Number.isFinite(balance.values[0][0])||
         typeof cash.values[0][0]!=="number"||
         !Number.isFinite(cash.values[0][0]))
        throw new Error("Perhitungan balance/modal pada sheet "+sh.name+
          " tidak menghasilkan angka. Stok sudah diperbarui, tetapi perlu Admin.");
    }
    return {changed:reads.length,sheets:ids.length};
  });
}
async function hbPcRecoverPending(){
  const pending=hbPcJournal().filter(x=>x.status==="pending").slice(-1)[0];
  if(!pending)return null;
  try{
    await hbPcVerifyPlan(pending);
    pending.status="verified";pending.verifiedAt=Date.now();
    await hbPcSaveJournal(pending);
    return "Koreksi sebelumnya terverifikasi setelah pemeriksaan ulang.";
  }catch(e){
    // If absolutely no cells were written, discard the old intent safely.
    // A partially applied plan remains blocked for explicit Admin review.
    try{
      const status=await Excel.run(async ctx=>{
        const cells=pending.cells.map(x=>{
          const r=ctx.workbook.worksheets.getItem(x.sheetId).getRange(x.address);
          r.load("values");return {r,x};
        });
        await ctx.sync();
        return cells.every(({r,x})=>hbPcNum(r.values[0][0])===x.before);
      });
      if(status){
        pending.status="cancelled";pending.verifiedAt=Date.now();
        await hbPcSaveJournal(pending);
        return "Koreksi tertunda dibatalkan dengan aman: tidak ada nilai stok yang berubah.";
      }
    }catch(_){}
    return "Ada koreksi BELUM TERVERIFIKASI ("+pending.targetName+"): "+
      (e.message||String(e))+". Koreksi baru diblokir sampai Admin memeriksa.";
  }
}
async function hbPcLoad(){
  if(hbPcState.working)return;
  hbPcStatus("Memeriksa riwayat belanja terakhir...",false);
  try{
    if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)
      throw new Error("Pengaman workbook belum aktif.");
    const pending=await hbPcRecoverPending();
    if(hbPcJournal().some(x=>x.status==="pending")){
      hbPcState.loaded=null;hbPcRender();
      throw new Error(pending||"Koreksi tertunda harus ditinjau Admin.");
    }
    hbPcState.loaded=await hbPcInspectLastPurchased();
    const data=hbPcState.loaded;
    if(!data){
      hbPcStatus("Tidak ditemukan riwayat belanja terverifikasi dari stok pada workbook ini. Tidak ada sheet yang dibuka.",true);
      hbPcRender();return;
    }
    hbPcStatus("SHEET BELANJA TERAKHIR: "+data.targetName+
      " | sumber pembanding: "+data.previousName+
      ". Koreksi jumlah Paket menyesuaikan D (stok awal), Rokok menyesuaikan O (belanja). "+
      "Harga dan rumus tetap. Balance dihitung ulang otomatis.",false);
    hbPcRender();
  }catch(e){hbPcStatus(e.message||String(e),true);}
}
async function hbPcSubmit(){
  if(hbPcBusy())return hbPcStatus("Excel masih menjalankan proses lain. Koreksi belum dimulai.",true);
  if(!hbPcState.loaded)return hbPcStatus("Muat dahulu sheet riwayat belanja.",true);
  if(!hbPcDom("purchaseCorrectionConfirm")?.checked)
    return hbPcStatus("Centang konfirmasi koreksi terlebih dahulu.",true);
  let changes;
  try{changes=hbPcChanges();}catch(e){return hbPcStatus(e.message,true);}
  if(!changes.length)return hbPcStatus("Belum ada jumlah belanja yang berubah.",true);
  hbPcState.working=true;
  const button=hbPcDom("purchaseCorrectionApply");
  if(button)button.disabled=true;
  let ownedBusy=false;
  try{
    const plan=await hbPcPlan(changes,hbPcState.loaded.targetId);
    if(!plan.cells.length)throw new Error("Tidak ditemukan perubahan stok yang dapat diterapkan.");
    if(typeof hbShiftPipelineBusy!=="undefined"){
      if(hbShiftPipelineBusy)throw new Error("Pipeline sheet masih aktif.");
      hbShiftPipelineBusy=true;ownedBusy=true;
    }
    plan.id="hbpc-"+plan.targetId+"-"+plan.createdAt;
    plan.status="pending";
    await hbPcSaveJournal(plan); // before ANY worksheet mutation
    hbPcStatus("Menerapkan "+plan.rows.length+" koreksi; memverifikasi stok dan balance...",false);
    await hbPcApplyCells(plan);
    await hbPcVerifyPlan(plan);
    plan.status="verified";plan.verifiedAt=Date.now();
    await hbPcSaveJournal(plan);
    hbPcDom("purchaseCorrectionConfirm").checked=false;
    // Refresh the visible quantities without calling hbPcLoad() while
    // hbPcState.working is still true (that function intentionally defers).
    try{
      hbPcState.loaded=await hbPcInspectLastPurchased();
      hbPcRender();
    }catch(_){hbPcState.loaded=null;}
    hbPcStatus("BERHASIL TERVERIFIKASI: "+plan.rows.length+" jenis belanja pada "+
      plan.targetName+". "+plan.cells.length+" sel stok diperbarui dalam "+
      [...new Set(plan.cells.map(x=>x.sheetId))].length+
      " sheet. Rumus modal dan balance tetap dan otomatis dihitung ulang.",false);
  }catch(e){
    hbPcStatus("KOREKSI BELUM TERKONFIRMASI: "+(e.message||String(e))+
      ". Jangan mengulangi sebelum status riwayat diperiksa.",true);
  }finally{
    if(ownedBusy)hbShiftPipelineBusy=false;
    hbPcState.working=false;
    if(button)button.disabled=false;
  }
}
Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel)return;
  const toggle=hbPcDom("purchaseCorrectionToggle"),body=hbPcDom("purchaseCorrectionBody");
  if(!toggle||!body)return;
  toggle.addEventListener("click",async()=>{
    body.hidden=!body.hidden;
    toggle.textContent=body.hidden?"Buka Koreksi":"Tutup Koreksi";
    if(!body.hidden)await hbPcLoad();
  });
  hbPcDom("purchaseCorrectionReload")?.addEventListener("click",hbPcLoad);
  hbPcDom("purchaseCorrectionApply")?.addEventListener("click",hbPcSubmit);
  hbPcDom("purchaseCorrectionSearch")?.addEventListener("input",event=>{
    hbPcState.filter=String(event.target.value||"").trim();hbPcRender();
  });
  hbPcDom("purchaseCorrectionShowAll")?.addEventListener("change",event=>{
    hbPcState.showAll=Boolean(event.target.checked);hbPcRender();
  });
});
