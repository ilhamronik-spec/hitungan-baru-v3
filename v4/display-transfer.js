"use strict";

// Tambahan untuk V4 yang sudah terpasang.
// Fitur ini BUKAN belanja: jumlah yang diinput dipindahkan dari gudang ke display.
// V2 tetap mengurus alur hari baru dan belanja. Setelah sheet baru selesai dibuat,
// script ini menambahkan jumlah transfer ke kolom F (JUMLAH AWAL display).
// Rumus workbook yang sudah ada kemudian menghitung L = perpindahan dan
// N = stok gudang akhir, sehingga stok gudang berkurang otomatis.

const displayTransferState = {
  open: false,
  sourceSheetName: "",
  items: []
};

let displayTransferWrapped = false;
let displayWrapTimer = null;
const ROKOK_DISPLAY_AUDIT_KEY="hitunganV4.rokokDisplayTransferAuditV1";
const ROKOK_DISPLAY_AUDIT_MAX=5; // Keep Office document settings small; last five shifts suffice for recovery.

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  bindDisplayTransferUi();
  waitAndWrapCreateNewDay();
  // Recover ONLY an already-journaled, unfinished transfer for the LATEST
  // worksheet. Never create a new sheet or guess quantities from stock.
  setTimeout(()=>resumePendingRokokDisplayTransfer().catch(e=>
    showDisplayTransferMessage("Transfer Rokok belum terpulihkan: "+displayErrorText(e),true)),16000);
});

function bindDisplayTransferUi() {
  const toggle = document.getElementById("displayTransferToggle");
  const body = document.getElementById("displayTransferBody");
  const search = document.getElementById("displayTransferSearch");
  const list = document.getElementById("displayTransferList");

  if (!toggle || !body || !search || !list) return;

  toggle.addEventListener("click", async () => {
    const opening = body.hidden;
    if (opening) {
      body.hidden = false;
      displayTransferState.open = true;
      toggle.textContent = "Tutup";
      try {
        await loadDisplayTransferItems(false);
      } catch (e) {
        showDisplayTransferMessage(displayErrorText(e), true);
      }
    } else {
      // Hiding this panel must NOT delete amounts already entered.
      // Transfers are cleared only after a verified new sheet is created.
      captureDisplayTransferDraftFromDOM();
      body.hidden=true;
      displayTransferState.open=false;
      const count=displayTransferState.items.filter(item=>Number(item.qtyText)>0).length;
      toggle.textContent=count?"Buka ("+count+" terisi)":"Buka";
    }
  });

  search.addEventListener("input", renderDisplayTransferItems);
  list.addEventListener("input", onDisplayTransferInput);
  list.addEventListener("change", onDisplayTransferInput);
  const form=document.getElementById("newDayForm");
  if(form)form.addEventListener("submit",event=>{
    captureDisplayTransferDraftFromDOM();
    const pending=displayTransferState.items.some(x=>Number(x.qtyText)>0);
    if(pending&&!displayTransferWrapped){
      event.preventDefault();
      event.stopImmediatePropagation();
      showDisplayTransferMessage(
        "Transfer Rokok masih disiapkan oleh add-in. Pembuatan sheet DIBATALKAN supaya input tidak hilang; coba ulang setelah add-in siap.",true);
    }
  },true);
}

async function loadDisplayTransferItems(forceReload) {
  const list = document.getElementById("displayTransferList");
  if (list) list.innerHTML = '<div class="empty-list">Memuat daftar rokok...</div>';

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const markers = ["A4", "I271", "A177"].map((address) => {
      const r = sheet.getRange(address);
      r.load("values");
      return r;
    });
    const range = sheet.getRange("A178:O238");

    sheet.load("name");
    range.load("values");
    await context.sync();

    if (typeof validateWorkbookTemplate === "function") {
      validateWorkbookTemplate(markers.map((r) => r.values[0][0]));
    }

    if (!forceReload && displayTransferState.sourceSheetName === sheet.name && displayTransferState.items.length) {
      return;
    }

    displayTransferState.sourceSheetName = sheet.name;
    displayTransferState.items = range.values
      .map((values, index) => ({
        row: index + 178,
        name: String(values[0] ?? "").trim(),
        oldCost: displayNumber(values[2]),
        displayEnding: displayNumber(values[6]),
        warehouseEnding: displayNumber(values[13]),
        qtyText: ""
      }))
      .filter((item) => item.name && item.oldCost > 0);
  });

  renderDisplayTransferItems();
  showDisplayTransferMessage(`Sumber stok: ${displayTransferState.sourceSheetName}`, false);
}

function renderDisplayTransferItems() {
  const list = document.getElementById("displayTransferList");
  const search = document.getElementById("displayTransferSearch");
  if (!list || !search) return;

  const query = String(search.value || "").trim().toLowerCase();
  const visible = displayTransferState.items.filter((item) => item.name.toLowerCase().includes(query));
  list.replaceChildren();

  if (!visible.length) {
    const empty = document.createElement("div");
    empty.className = "empty-list";
    empty.textContent = displayTransferState.items.length ? "Rokok tidak ditemukan." : "Belum ada daftar rokok.";
    list.appendChild(empty);
    updateDisplayTransferSummary();
    return;
  }

  visible.forEach((item) => {
    const row = document.createElement("article");
    row.className = "product-row";

    const name = document.createElement("div");
    name.className = "product-name";
    name.textContent = item.name;
    row.appendChild(name);

    const meta = document.createElement("div");
    meta.className = "product-meta";
    meta.textContent = `Display akhir: ${formatDisplayNumber(item.displayEnding)} | Gudang akhir: ${formatDisplayNumber(item.warehouseEnding)}`;
    row.appendChild(meta);

    const fields = document.createElement("div");
    fields.className = "purchase-fields";
    fields.style.gridTemplateColumns = "1fr";

    const label = document.createElement("label");
    label.textContent = "Jumlah dari gudang ke display";

    const input = document.createElement("input");
    input.type = "text";
    input.inputMode = "numeric";
    input.autocomplete = "off";
    input.placeholder = "0";
    input.value = item.qtyText;
    input.dataset.displayRow = String(item.row);
    input.setAttribute("aria-label", `Jumlah ${item.name} dari gudang ke display`);

    label.appendChild(input);
    fields.appendChild(label);
    row.appendChild(fields);
    list.appendChild(row);
  });

  updateDisplayTransferSummary();
}

function onDisplayTransferInput(event) {
  const input = event.target;
  if (!input.matches("input[data-display-row]")) return;

  const row = Number(input.dataset.displayRow);
  const item = displayTransferState.items.find((x) => x.row === row);
  if (!item) return;
  item.qtyText = input.value;
  updateDisplayTransferSummary();
}

function updateDisplayTransferSummary() {
  const summary = document.getElementById("displayTransferSummary");
  if (!summary) return;

  let count = 0;
  let total = 0;
  for (const item of displayTransferState.items) {
    const text = String(item.qtyText || "").trim();
    if (!text || !/^\d+$/.test(text)) continue;
    const qty = Number(text);
    if (qty > 0) {
      count += 1;
      total += qty;
    }
  }
  summary.textContent = count ? `Terisi: ${count} jenis rokok, total ${formatDisplayNumber(total)} ke display` : "Belum ada rokok yang dipindahkan ke display.";
}

function captureDisplayTransferDraftFromDOM(){
  const list=document.getElementById("displayTransferList");
  if(!list||typeof list.querySelectorAll!=="function")return;
  for(const input of list.querySelectorAll('input[data-display-row]')){
    const row=Number(input.dataset.displayRow);
    const item=displayTransferState.items.find(x=>x.row===row);
    if(item)item.qtyText=String(input.value||"");
  }
}

function collectDisplayTransfers(sourceName, purchaseData) {
  // Preserve pending transfers even when the panel is collapsed. Read the
  // current DOM too, so the last typed quantity cannot be silently omitted.
  captureDisplayTransferDraftFromDOM();
  if (!displayTransferState.items.some(item=>String(item.qtyText||"").trim()!=="")) return [];

  if (displayTransferState.sourceSheetName && displayTransferState.sourceSheetName !== sourceName) {
    throw new Error(`Daftar tambah display berasal dari "${displayTransferState.sourceSheetName}", tetapi sheet aktif sekarang "${sourceName}". Tutup lalu buka kembali bagian Tambah Display.`);
  }

  const purchaseByRow = new Map();
  if (purchaseData && Array.isArray(purchaseData.cigarettes)) {
    purchaseData.cigarettes.forEach((item) => {
      // V2 can submit quantities as numeric input strings. Accept whole
      // decimal digits, but never silently coerce invalid/free-form amounts.
      const raw=String(item.qty??"").trim();
      const qty=/^\d+$/.test(raw)?Number(raw):0;
      if(Number.isSafeInteger(qty)&&qty>0)
        purchaseByRow.set(Number(item.row),qty);
    });
  }

  const transfers = [];
  for (const item of displayTransferState.items) {
    const text = String(item.qtyText || "").trim();
    if (!text) continue;
    if (!/^\d+$/.test(text)) {
      throw new Error(`Jumlah tambah display rokok "${item.name}" harus berupa angka bulat.`);
    }

    const qty = Number(text);
    if (qty <= 0) continue;

    const available = displayNumber(item.warehouseEnding) + displayNumber(purchaseByRow.get(item.row));
    if (qty > available) {
      throw new Error(`Stok gudang "${item.name}" tidak cukup. Tersedia ${formatDisplayNumber(available)}, diminta ${formatDisplayNumber(qty)}.`);
    }

    transfers.push({ row: item.row, name: item.name, qty });
  }
  return transfers;
}

function waitAndWrapCreateNewDay() {
  if (displayWrapTimer) clearInterval(displayWrapTimer);

  displayWrapTimer = setInterval(() => {
    try {
      if (displayTransferWrapped) {
        clearInterval(displayWrapTimer);
        displayWrapTimer = null;
        return;
      }

      // V4 mengatur v2Loaded=true lalu memasang wrapper pengaman dalam callback
      // yang sama. Timer ini baru berjalan setelah callback tersebut selesai.
      if (typeof createNewDaySheet !== "function" || typeof requireSecret !== "function") return;
      if (typeof v2Loaded !== "undefined" && !v2Loaded) return;

      const securedCreateNewDay = createNewDaySheet;
      createNewDaySheet = async function(person, sheetName, purchaseData) {
        let sourceName = "";

        await Excel.run(async (context) => {
          const source = context.workbook.worksheets.getActiveWorksheet();
          source.load("name");
          await context.sync();
          sourceName = source.name;
        });

        const transfers = collectDisplayTransfers(sourceName, purchaseData);
        const result = await securedCreateNewDay(person, sheetName, purchaseData);

        if (transfers.length) {
          try {
            await applyDisplayTransfersToNewSheet(sheetName, transfers, sourceName);
          } catch (e) {
            showDisplayTransferMessage(displayErrorText(e),true);
            throw new Error(`Sheet "${sheetName}" sudah dibuat tetapi pindah stok Rokok ke display belum terverifikasi: ${displayErrorText(e)}. JANGAN buat sheet lagi; periksa stok pada sheet yang sudah dibuat sebelum melanjutkan.`);
          }
        }

        resetDisplayTransferDrafts();
        return result;
      };

      displayTransferWrapped = true;
      clearInterval(displayWrapTimer);
      displayWrapTimer = null;
    } catch (_) {
      // Coba lagi pada tick berikutnya.
    }
  }, 250);
}


async function resumePendingRokokDisplayTransfer(attempt=0){
  if(typeof getSetting!=="function"||getSetting("hitunganV4.enabled")!==true)return;
  const journal=readRokokTransferAudit().filter(x=>x.status==="pending");
  if(!journal.length)return;
  if((typeof hbShiftPipelineBusy!=="undefined"&&hbShiftPipelineBusy)||
     (typeof shiftCreationInProgress!=="undefined"&&shiftCreationInProgress)||
     (typeof minesRecoveryBusy!=="undefined"&&minesRecoveryBusy)||
     (typeof hbMarginColorBusy!=="undefined"&&hbMarginColorBusy)){
    if(attempt<4)setTimeout(()=>
      resumePendingRokokDisplayTransfer(attempt+1).catch(e=>
        showDisplayTransferMessage("Pemulihan transfer tertunda: "+displayErrorText(e),true)),12000);
    return;
  }
  const latest=await Excel.run(async ctx=>{
    const ws=ctx.workbook.worksheets;
    ws.load("items/name,items/position");
    await ctx.sync();
    if(!ws.items.length)return null;
    return ws.items.reduce((a,b)=>Number(a.position)<Number(b.position)?a:b).name;
  });
  const planned=journal.find(x=>x.target===latest);
  if(!planned||!Array.isArray(planned.rows)||!planned.rows.length)return;
  await applyDisplayTransfersToNewSheet(planned.target,
    planned.rows.map(x=>({row:x.row,name:x.name,qty:x.qty})),planned.source);
  if(typeof setSecurityStatus==="function"){
    setSecurityStatus("TRANSFER ROKOK DIPULIHKAN pada sheet terbaru '"+latest+
      "' dan F diverifikasi. Tidak ada penambahan ganda.",false);
  }
}

function readRokokTransferAudit(){
  const raw=typeof getSetting==="function"?getSetting(ROKOK_DISPLAY_AUDIT_KEY):null;
  return Array.isArray(raw)?raw.filter(x=>x&&typeof x==="object"):[];
}
async function saveRokokTransferAudit(entry){
  if(typeof setSetting!=="function"||typeof saveSettings!=="function")
    throw new Error("Mesin penyimpanan audit stok belum siap.");
  const all=readRokokTransferAudit().filter(x=>x.target!==entry.target);
  all.push(entry);
  setSetting(ROKOK_DISPLAY_AUDIT_KEY,all.slice(-ROKOK_DISPLAY_AUDIT_MAX));
  await saveSettings();
}
function getRokokTransferAudit(target){
  return readRokokTransferAudit().find(x=>x.target===target)||null;
}
function renderRokokTransferAudit(){
  const el=document.getElementById("rokokTransferAuditStatus");
  if(!el)return;
  const rows=readRokokTransferAudit().slice(-5).reverse();
  el.textContent=rows.length?rows.map(x=>
    x.target+": "+(x.status==="verified"?"TERVERIFIKASI":"PERLU DIPERIKSA")+
    " ("+x.rows.length+" barang, "+x.rows.reduce((s,r)=>s+Number(r.qty||0),0)+" pcs)"
  ).join(" | "):"Belum ada catatan transfer Rokok versi terbaru.";
}
// A target-specific journal is written BEFORE any F-cell changes. If Excel
// rejects a later operation, retry checks baseline vs expected to prevent
// silently applying the same transfer twice.
async function applyDisplayTransfersToNewSheet(sheetName,transfers,sourceName){
  if(!Array.isArray(transfers)||!transfers.length)return;
  const pw=requireSecret();
  const existing=getRokokTransferAudit(sheetName);
  if(existing&&existing.status==="verified"&&
     (existing.source!==sourceName ||
      transfers.some(x=>!existing.rows.some(y=>y.row===x.row&&y.qty===x.qty))))
    throw new Error("Transfer Rokok ke sheet ini sudah tercatat; jumlah berbeda. Hubungi admin.");
  // Always read Excel again, even for a previously verified transfer.
  // A cached verification must not hide a later workbook synchronization error.
  let planned=existing;
  if(!planned){
    planned=await Excel.run(async ctx=>{
      const sheet=ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("id,name");
      const values=transfers.map(item=>{
        const cell=sheet.getRange("F"+item.row);
        cell.load("values");
        const name=sheet.getRange("A"+item.row);name.load("values");
        return {item,cell,name};
      });
      await ctx.sync();
      if(sheet.name!==sheetName)
        throw new Error('Sheet tujuan berubah. Harus "'+sheetName+'", aktif "'+sheet.name+'".');
      return {target:sheetName,source:sourceName,createdAt:Date.now(),status:"pending",
        rows:values.map(({item,cell,name})=>{
          const before=displayNumber(cell.values[0][0]);
          if(String(name.values[0][0]||"").trim()!==String(item.name).trim())
            throw new Error("Nama Rokok di sheet tujuan berbeda pada baris "+item.row+".");
          return {row:item.row,name:item.name,qty:item.qty,before,expected:before+item.qty};
        })};
    });
    await saveRokokTransferAudit(planned);
  }
  if(planned.source!==sourceName)
    throw new Error("Sumber transfer tidak sama dengan catatan sebelumnya.");
  const finalRows=await Excel.run(async ctx=>{
    const sheet=ctx.workbook.worksheets.getItem(planned.target);
    sheet.load("id,name,protection/protected");
    const workMarker=sheet.getRange("A1");
    workMarker.format.protection.load("locked");
    const targets=planned.rows.map(item=>{
      const cell=sheet.getRange("F"+item.row);cell.load("values");
      const name=sheet.getRange("A"+item.row);name.load("values");
      return {item,cell,name};
    });
    await ctx.sync();
    for(const {item,cell,name} of targets){
      if(String(name.values[0][0]||"").trim()!==String(item.name).trim())
        throw new Error("Nama Rokok berubah pada baris "+item.row+". Transfer dihentikan.");
      const current=displayNumber(cell.values[0][0]);
      if(current!==item.before && current!==item.expected)
        throw new Error("F"+item.row+" berubah menjadi "+current+
          "; diharapkan "+item.before+" atau "+item.expected+
          ". Jangan tambah dua kali; periksa manual.");
    }
    const needsWrite=targets.filter(({item,cell})=>
      displayNumber(cell.values[0][0])===item.before && item.expected!==item.before);
    const originallyProtected=sheet.protection.protected;
    const wasFullLock=originallyProtected && workMarker.format.protection.locked===true;
    let failure=null;
    try{
      if(needsWrite.length){
        if(originallyProtected){sheet.protection.unprotect(pw);await ctx.sync();}
        for(const {item,cell} of needsWrite)cell.values=[[item.expected]];
        await ctx.sync(); // Values must persist before recalculating formulas.
      }
      // Also recalculate when F is already correct but G was stale because
      // Excel had not refreshed the sheet's dependent stock formulas.
      sheet.calculate(true);
      await ctx.sync();
      for(const {cell} of targets)cell.load("values");
      await ctx.sync();
      for(const {item,cell} of targets){
        if(displayNumber(cell.values[0][0])!==item.expected)
          throw new Error("Stok display Rokok F"+item.row+" belum sesuai: diharapkan "+
            item.expected+", terbaca "+cell.values[0][0]+".");
        // Column G is the independently entered ENDING DISPLAY STOCK,
        // not a calculated quantity that must rise with a warehouse transfer.
        // H=F-G and L=F-K handle the downstream calculations in the workbook.
        // Only F is the transferred quantity destination; never write G.
      }
    }catch(e){failure=e;}
    finally{
      if(needsWrite.length && originallyProtected && !isGlobalUnlockMode()){
        // Restore the EXACT protection mode seen before this transfer.
        // A temporarily unsynchronized 2.5-hour timer must never force the
        // newest working sheet into Full Lock just for changing F.
        if(wasFullLock)fullLock(sheet,pw);
        else workingLock(sheet,pw);
        await ctx.sync();
      }
    }
    if(failure)throw failure;
    return targets.map(({item})=>({row:item.row,qty:item.qty,expected:item.expected}));
  });
  // Independently read back F AFTER the write/recalculation/protection.
  // A successful worksheet operation alone does not prove stock persisted.
  const confirmed=await Excel.run(async ctx=>{
    const sheet=ctx.workbook.worksheets.getItem(planned.target);
    const cells=planned.rows.map(item=>{
      const r=sheet.getRange("F"+item.row);r.load("values");
      return {item,r};
    });
    await ctx.sync();
    return cells.map(({item,r})=>({
      row:item.row,expected:item.expected,actual:displayNumber(r.values[0][0])
    }));
  });
  const missing=confirmed.filter(item=>item.actual!==item.expected);
  if(missing.length)throw new Error("Transfer Rokok belum tersimpan di sheet baru: "+
    missing.map(x=>"F"+x.row+" diharapkan "+x.expected+", terbaca "+x.actual).join("; ")+
    ". JANGAN buat sheet kedua untuk shift yang sama.");
  planned.status="verified";
  planned.verifiedAt=Date.now();
  await saveRokokTransferAudit(planned);
  renderRokokTransferAudit();
  showDisplayTransferMessage(finalRows.length+" jenis Rokok masuk display dan stok F sudah diverifikasi.",false);
}

function resetDisplayTransferDrafts() {
  displayTransferState.open = false;
  displayTransferState.sourceSheetName = "";
  displayTransferState.items = [];

  const body = document.getElementById("displayTransferBody");
  const toggle = document.getElementById("displayTransferToggle");
  const search = document.getElementById("displayTransferSearch");
  const list = document.getElementById("displayTransferList");
  const summary = document.getElementById("displayTransferSummary");
  const message = document.getElementById("displayTransferMessage");

  if (body) body.hidden = true;
  if (toggle) toggle.textContent = "Buka";
  if (search) search.value = "";
  if (list) list.replaceChildren();
  if (summary) summary.textContent = "";
  if (message) message.textContent = "";
}

function showDisplayTransferMessage(text, isError) {
  const el = document.getElementById("displayTransferMessage");
  if (!el) return;
  el.textContent = text || "";
  el.style.color = isError ? "#b42318" : "#5b6472";
}

function displayNumber(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
}

function formatDisplayNumber(value) {
  return new Intl.NumberFormat("id-ID", { maximumFractionDigits: 2 }).format(displayNumber(value));
}

function displayErrorText(error) {
  if (error && typeof error.message === "string") return error.message;
  return String(error || "Terjadi kesalahan.");
}
