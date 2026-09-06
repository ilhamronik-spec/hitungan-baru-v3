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

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  bindDisplayTransferUi();
  waitAndWrapCreateNewDay();
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
      resetDisplayTransferDrafts();
    }
  });

  search.addEventListener("input", renderDisplayTransferItems);
  list.addEventListener("input", onDisplayTransferInput);
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

function collectDisplayTransfers(sourceName, purchaseData) {
  if (!displayTransferState.open) return [];

  if (displayTransferState.sourceSheetName && displayTransferState.sourceSheetName !== sourceName) {
    throw new Error(`Daftar tambah display berasal dari "${displayTransferState.sourceSheetName}", tetapi sheet aktif sekarang "${sourceName}". Tutup lalu buka kembali bagian Tambah Display.`);
  }

  const purchaseByRow = new Map();
  if (purchaseData && Array.isArray(purchaseData.cigarettes)) {
    purchaseData.cigarettes.forEach((item) => {
      const qty = displayNumber(item.qty);
      if (qty > 0) purchaseByRow.set(Number(item.row), qty);
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
            await applyDisplayTransfersToNewSheet(sheetName, transfers);
          } catch (e) {
            throw new Error(`Sheet "${sheetName}" sudah dibuat, tetapi tambah stok display gagal: ${displayErrorText(e)} Jangan buat sheet baru lagi; buka Admin/Koreksi.`);
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

async function applyDisplayTransfersToNewSheet(sheetName, transfers) {
  const pw = requireSecret();

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name,protection/protected");

    const targetRanges = transfers.map((item) => {
      const range = sheet.getRange(`F${item.row}`);
      range.load("values");
      return { item, range };
    });

    await context.sync();

    if (sheet.name !== sheetName) {
      throw new Error(`Sheet aktif berubah menjadi "${sheet.name}".`);
    }

    if (sheet.protection.protected) {
      sheet.protection.unprotect(pw);
      await context.sync();
    }

    targetRanges.forEach(({ item, range }) => {
      const current = displayNumber(range.values[0][0]);
      range.values = [[current + item.qty]];
    });

    sheet.calculate(true);
    workingLock(sheet, pw);
    await context.sync();
  });
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
