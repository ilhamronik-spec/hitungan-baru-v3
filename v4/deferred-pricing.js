"use strict";

// Prices are staged for the next sheet created from the selected SOURCE sheet.
// Historical/current sheets are never changed by saving a scheduled price.
const PRICE_PENDING_KEY="hitunganV4.pendingSellingPricesV1";
const LOW_MARGIN_FILL="#FFC7CE";
const LOW_MARGIN_FONT="#9C0006";
const MIN_MARGIN_RUPIAH=1000;
const SELLING_PRICE_GROUPS={
  paket:{start:6,end:75,label:"Paket"},
  rokok:{start:178,end:238,label:"Rokok"}
};
let sellingEditorRows=[];
let sellingEditorSheet="";
let sellingEditorCategory="paket";

function sellingPriceQueue(){
  const raw=getSetting(PRICE_PENDING_KEY);
  return Array.isArray(raw)?raw.filter(p=>p&&typeof p==="object"):[];
}
function numericSellingPrice(value){
  const n=Number(value);
  return Number.isFinite(n)?n:NaN;
}
function scheduledPriceId(source,category,row){
  return source+"|"+category+"|"+row;
}
function sellingName(v){return String(v===undefined||v===null?"":v).trim().toLocaleLowerCase("id-ID");}
function setSellingStatus(t,error=false){
  const el=document.getElementById("sellingPriceStatus");
  if(!el)return;
  el.textContent=t;
  el.style.color=error?"#a4262c":"#15613b";
}
function getSourcePriceChanges(sheetName){
  return sellingPriceQueue().filter(item=>item.source===sheetName);
}

async function loadSellingPriceSheets(){
  if(typeof adminSessionAuthenticated!=="undefined"&&!adminSessionAuthenticated)return;
  const sheets=await Excel.run(async ctx=>{
    const collection=ctx.workbook.worksheets;
    collection.load("items/name,items/position");
    const active=collection.getActiveWorksheet();
    active.load("name");
    await ctx.sync();
    return {all:collection.items.map(s=>({name:s.name,position:s.position})),active:active.name};
  });
  const select=document.getElementById("sellingSourceSheet");
  if(!select)return;
  const old=select.value;
  select.replaceChildren();
  for(const sh of sheets.all.sort((a,b)=>a.position-b.position)){
    const opt=document.createElement("option");
    opt.value=sh.name;opt.textContent=sh.name;select.appendChild(opt);
  }
  select.value=sheets.all.some(sh=>sh.name===old)?old:sheets.active;
  await loadSellingPriceRows();
}

async function loadSellingPriceRows(){
  if(typeof adminSessionAuthenticated!=="undefined"&&!adminSessionAuthenticated)return;
  const source=document.getElementById("sellingSourceSheet");
  const cat=document.getElementById("sellingCategory");
  if(!source||!cat)return;
  sellingEditorSheet=source.value;
  sellingEditorCategory=cat.value;
  const g=SELLING_PRICE_GROUPS[sellingEditorCategory];
  if(!sellingEditorSheet||!g)return;
  const rows=await Excel.run(async ctx=>{
    const sh=ctx.workbook.worksheets.getItem(sellingEditorSheet);
    const block=sh.getRange("A"+g.start+":C"+g.end);
    block.load("values");
    await ctx.sync();
    return block.values.map((v,i)=>({
      row:g.start+i,name:String(v[0]??"").trim(),oldPrice:numericSellingPrice(v[2])
    })).filter(r=>r.name);
  });
  sellingEditorRows=rows;
  renderSellingPriceRows();
}

function renderSellingPriceRows(){
  const container=document.getElementById("sellingPriceRows");
  if(!container)return;
  container.replaceChildren();
  const q=String(document.getElementById("sellingPriceSearch")?.value||"").toLowerCase().trim();
  const pending=sellingPriceQueue();
  const matches=sellingEditorRows.filter(x=>x.name.toLowerCase().includes(q));
  for(const item of matches){
    const row=document.createElement("div");
    row.className="product-row selling-edit-row";
    const label=document.createElement("label");
    label.className="product-name";
    label.textContent=item.name+" — baris "+item.row;
    const info=document.createElement("div");
    info.className="product-meta";
    info.textContent="Harga sekarang: Rp "+(Number.isFinite(item.oldPrice)?item.oldPrice.toLocaleString("id-ID"):"-");
    const input=document.createElement("input");
    input.type="number"; input.min="0";input.step="1";
    input.placeholder="Harga jual baru (shift selanjutnya)";
    input.setAttribute("aria-label","Harga baru "+item.name);
    const found=pending.find(p=>p.source===sellingEditorSheet&&p.group===sellingEditorCategory&&p.row===item.row);
    if(found)input.value=String(found.price);
    const save=document.createElement("button");
    save.type="button";save.className="secondary-button selling-save";
    save.textContent=found?"Perbarui harga pending":"Jadwalkan harga";
    save.addEventListener("click",async ()=>{
      try{
        if(!adminSessionAuthenticated)throw new Error("Masuk sebagai admin terlebih dahulu.");
        const p=Number(input.value);
        if(!Number.isInteger(p)||p<=0)throw new Error("Harga harus rupiah bulat positif.");
        const all=sellingPriceQueue();
        const key=scheduledPriceId(sellingEditorSheet,sellingEditorCategory,item.row);
        const clean=all.filter(x=>scheduledPriceId(x.source,x.group,x.row)!==key);
        clean.push({source:sellingEditorSheet,group:sellingEditorCategory,row:item.row,name:item.name,
          price:p,savedAt:Date.now(),id:key});
        setSetting(PRICE_PENDING_KEY,clean);
        await saveSettings();
        setSellingStatus(item.name+": harga Rp "+p.toLocaleString("id-ID")+
          " DIJADWALKAN untuk sheet setelah '"+sellingEditorSheet+"'. Sheet saat ini tidak berubah.");
        renderSellingPriceRows();
      }catch(e){setSellingStatus(errText(e),true);}
    });
    row.append(label,info,input,save);
    if(found){
      const cancel=document.createElement("button");
      cancel.className="secondary-button";cancel.type="button";cancel.textContent="Batalkan pending";
      cancel.addEventListener("click",async ()=>{
        try{
          if(!adminSessionAuthenticated)throw new Error("Masuk sebagai admin.");
          const clean=sellingPriceQueue().filter(p=>scheduledPriceId(p.source,p.group,p.row)!==found.id);
          setSetting(PRICE_PENDING_KEY,clean);await saveSettings();renderSellingPriceRows();
          setSellingStatus("Harga pending "+item.name+" dibatalkan; sheet belum diubah.");
        }catch(e){setSellingStatus(errText(e),true);}
      });
      row.appendChild(cancel);
    }
    container.appendChild(row);
  }
  if(!matches.length)container.textContent="Barang tidak ditemukan.";
}

function readPurchaseRowDetails(purchaseData){
  const results=[];
  if(!purchaseData||typeof purchaseData!=="object")return results;
  function numberOf(obj,keys){
    for(const key of keys){
      const n=Number(obj[key]);
      if(obj[key]!==null&&obj[key]!==undefined&&obj[key]!==""&&Number.isFinite(n))return n;
    }
    return NaN;
  }
  function add(items,group){
    if(!Array.isArray(items))return;
    for(const item of items){
      if(!item||typeof item!=="object")continue;
      const row=Number(item.row);
      const limits=SELLING_PRICE_GROUPS[group];
      if(!Number.isInteger(row)||row<limits.start||row>limits.end)continue;
      const qty=numberOf(item,["qty","quantity","pcs","jumlah"]);
      let cost=NaN;
      if(group==="paket"){
        cost=numberOf(item,["newCost","basePrice","costPrice","unitCost","cost","buyPrice","hargaDasar","price","harga"]);
      }else{
        const total=numberOf(item,["totalCost","totalPurchase","purchaseTotal","amount","total","spend","buyTotal","totalPrice","hargaBelanja","hargaTotal"]);
        if(Number.isFinite(total)&&Number.isFinite(qty)&&qty>0)cost=total/qty;
        else cost=numberOf(item,["newCost","unitCost","costPrice","basePrice","hargaDasar"]);
      }
      if(Number.isFinite(cost)&&cost>0)results.push({group,row,cost,qty});
    }
  }
  add(purchaseData.packages||purchaseData.packets||purchaseData.paket,"paket");
  add(purchaseData.cigarettes||purchaseData.rokok,"rokok");
  return results;
}

// Called by the central V4 create wrapper ONLY on the new worksheet.
// The source sheet is used to validate when a pending price becomes effective.
async function applyNextShiftPricing(context,newSheet,sourceName,purchaseData){
  const queue=getSourcePriceChanges(sourceName);
  const purchases=readPurchaseRowDetails(purchaseData);
  const alerts=[];
  const applied=[];
  const skipped=[];
  // The newly copied sheet may inherit alert styling from the source.
  const sellCells=[];
  for(const group of Object.keys(SELLING_PRICE_GROUPS)){
    const range=SELLING_PRICE_GROUPS[group];
    for(let row=range.start;row<=range.end;row++){
      const cell=newSheet.getRange("C"+row);
      cell.format.fill.load("color");
      sellCells.push(cell);
    }
  }
  const checks=queue.map(item=>{
    const cell=newSheet.getRange("A"+item.row);
    cell.load("values");
    return {item,cell};
  });
  await context.sync();
  for(const cell of sellCells){
    if(String(cell.format.fill.color||"").toUpperCase()===LOW_MARGIN_FILL){
      cell.format.fill.clear();
      cell.format.font.color="#000000";
    }
  }
  for(const {item,cell} of checks){
    const cfg=SELLING_PRICE_GROUPS[item.group];
    if(!cfg||item.row<cfg.start||item.row>cfg.end||sellingName(cell.values[0][0])!==sellingName(item.name)){
      skipped.push(item.name+" (nama/baris tidak cocok)");
      continue;
    }
    newSheet.getRange("C"+item.row).values=[[item.price]];
    applied.push(item.id);
  }
  const sell= purchases.map(item=>({
    item, cell:newSheet.getRange("C"+item.row)
  }));
  sell.forEach(x=>x.cell.load("values"));
  await context.sync();
  for(const {item,cell} of sell){
    const sellPrice=numericSellingPrice(cell.values[0][0]);
    if(Number.isFinite(sellPrice)&&sellPrice-item.cost<MIN_MARGIN_RUPIAH){
      // Alert is informational; visual category now belongs exclusively to
      // conditional formatting on the product NAME (column A).
      alerts.push(item.group+" baris "+item.row+": margin "+(sellPrice-item.cost));
    }
  }
  await context.sync();
  return {applied,skipped,alerts};
}

async function consumeEffectivePrices(ids){
  if(!ids||!ids.length)return;
  const done=new Set(ids);
  const rest=sellingPriceQueue().filter(p=>!done.has(p.id));
  setSetting(PRICE_PENDING_KEY,rest);
  await saveSettings();
}
function initSellingPriceEditor(){
  const source=document.getElementById("sellingSourceSheet");
  const cat=document.getElementById("sellingCategory");
  const search=document.getElementById("sellingPriceSearch");
  const refresh=document.getElementById("refreshSellingPrices");
  if(!source||!cat||!search||!refresh)return;
  source.addEventListener("change",()=>loadSellingPriceRows().catch(e=>setSellingStatus(errText(e),true)));
  cat.addEventListener("change",()=>loadSellingPriceRows().catch(e=>setSellingStatus(errText(e),true)));
  search.addEventListener("input",renderSellingPriceRows);
  refresh.addEventListener("click",()=>loadSellingPriceSheets().catch(e=>setSellingStatus(errText(e),true)));
}
Office.onReady(info=>{
  if(info.host!==Office.HostType.Excel)return;
  initSellingPriceEditor();
});
