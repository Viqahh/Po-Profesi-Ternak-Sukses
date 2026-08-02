const MASTER_SPREADSHEET_ID = "";
const INDEX_SHEET_NAME = "Daftar Link";
const DATA_SHEET_NAME = "PO OVK";
const REGION_NAMES = ["Sengkang", "Maros", "Takalar", "Kendari", "Masamba"];

// Isi ID spreadsheet tujuan kalau data OVK harus masuk ke file yang sudah ada.
// Contoh link:
// https://docs.google.com/spreadsheets/d/INI_SPREADSHEET_ID_NYA/edit
const TARGET_SPREADSHEET_IDS = {
  Sengkang: "1_lvgfz60dZXnvOr1rZAJBqG9PbUeEKnprXW-Jx4-NtM",
  Maros: "1rqEIM_DpCK0He-h57WdQJDAjZoKzk7LAVQwX0cKShyQ",
  Takalar: "1YJssHcTLMRaaM6JFFW4balFO5qultpwQTBbVAwkIrmU",
  Kendari: "1HifqQGwfl6alzwtd1DYn94yMD2yopNb4fLleE0mIHj8",
  Masamba: "1wbVap6EB7QSKrBxa4Lz4WzmRkjXEpyQ0b1OCJIyzSRk"
};

const REGION_PASSWORDS = {
  Maros: "Maros010",
  Takalar: "Takalar020",
  Sengkang: "Sengkang030",
  Kendari: "Kendari040",
  Masamba: "Masamba050"
};

const ADMIN_ID = "PTS100";
const ADMIN_PASSWORD = "170304";

const HEADERS = [
  "Group Submit",
  "Timestamp Server",
  "Timestamp Submit",
  "Unit / Daerah",
  "Kode Unit",
  "No PO",
  "Tgl PO",
  "Tgl Check In",
  "Nama Peternak",
  "Populasi",
  "Alamat Peternak",
  "Nama PPL",
  "Admin",
  "Download PDF",
  "Timestamp Download PDF",
  "No Barang",
  "Nama Barang",
  "Qty",
  "Satuan",
  "Suplier"
];

function doPost(e) {
  const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
  const form = payload.form || {};
  const region = normalizeRegion_(form.unit || payload.daerah || payload.unit || payload.targetSheet);
  const authError = validateSubmitAuth_(payload, region);
  if (authError) {
    return jsonOutput_({
      ok: false,
      error: authError
    });
  }

  const regionSpreadsheet = getOrCreateRegionSpreadsheet_(region);
  const sheet = setupRegionSpreadsheet_(regionSpreadsheet, region);
  const rows = buildDataRows_(payload, region, new Date());

  if (rows.length) {
    const orderRow = buildOrderHeaderRow_(payload, region, rows.length);
    const allRows = [orderRow, ...rows];
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, allRows.length, HEADERS.length).setValues(allRows);
    applyDownloadStatus_(sheet, startRow + 1, rows.length, Boolean(payload.pdfDownloaded));
    formatSubmitGroup_(sheet, startRow, allRows.length, (payload.form || {}).noPo || "");
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  updateIndexSheet_(region, regionSpreadsheet);

  return jsonOutput_({
    ok: true,
    region,
    spreadsheetId: regionSpreadsheet.getId(),
    spreadsheetUrl: regionSpreadsheet.getUrl(),
    rows: rows.length
  });
}

function doGet() {
  bootstrapRegionSpreadsheets();
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: `PO endpoint aktif. Data masuk ke tab ${DATA_SHEET_NAME}.` }))
    .setMimeType(ContentService.MimeType.JSON);
}

function bootstrapRegionSpreadsheets() {
  REGION_NAMES.forEach((region) => {
    const spreadsheet = getOrCreateRegionSpreadsheet_(region);
    setupRegionSpreadsheet_(spreadsheet, region);
    updateIndexSheet_(region, spreadsheet);
  });
}

function getOrCreateRegionSpreadsheet_(region) {
  const props = PropertiesService.getScriptProperties();
  const configuredId = TARGET_SPREADSHEET_IDS[region];
  if (configuredId) {
    return SpreadsheetApp.openById(configuredId);
  }

  const key = `OVK_REGION_SPREADSHEET_${region.toUpperCase()}`;
  const savedId = props.getProperty(key);

  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (error) {
      props.deleteProperty(key);
    }
  }

  const spreadsheet = SpreadsheetApp.create(`PO ${region}`);
  props.setProperty(key, spreadsheet.getId());
  return spreadsheet;
}

function setupRegionSpreadsheet_(spreadsheet, region) {
  const sheet = spreadsheet.getSheetByName(DATA_SHEET_NAME) || spreadsheet.insertSheet(DATA_SHEET_NAME);
  setupHeader_(sheet);
  return sheet;
}

function setupHeader_(sheet) {
  const currentHeader = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeader = currentHeader.join("") !== HEADERS.join("");

  if (needsHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#f4b183")
      .setHorizontalAlignment("center");
  }

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), HEADERS.length).createFilter();
  }
}

function updateIndexSheet_(region, regionSpreadsheet) {
  const master = getMasterSpreadsheet_();
  if (!master) return;

  const sheet = master.getSheetByName(INDEX_SHEET_NAME) || master.insertSheet(INDEX_SHEET_NAME);
  const headers = ["Daerah", "Tab Data", "Link Google Sheets", "Spreadsheet ID", "Last Updated"];
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  if (firstRow.join("") !== headers.join("")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#bdd7ee")
      .setHorizontalAlignment("center");
  }

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const index = values.indexOf(region);
  const targetRow = index >= 0 ? index + 2 : lastRow + 1;
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([[
    region,
    DATA_SHEET_NAME,
    regionSpreadsheet.getUrl(),
    regionSpreadsheet.getId(),
    new Date()
  ]]);
  sheet.autoResizeColumns(1, headers.length);
}

function getMasterSpreadsheet_() {
  if (MASTER_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  }

  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    return null;
  }
}

function buildDataRows_(payload, region, serverTimestamp) {
  const form = payload.form || {};
  const items = Array.isArray(payload.items) && payload.items.length
    ? payload.items
    : [{ no: "", namaBarang: "", qty: "", satuan: "", supplier: "" }];

  return items.map((item, index) => [
    form.noPo || payload.submitTimestamp || serverTimestamp,
    serverTimestamp,
    payload.submitTimestamp || "",
    region,
    form.kodeUnit || payload.kodeUnit || "",
    form.noPo || "",
    form.tglPo || "",
    form.tglCheckIn || "",
    form.namaPeternak || "",
    form.populasi || "",
    form.alamatPeternak || "",
    form.namaPpl || "",
    payload.admin || "Ananda Thufail",
    payload.pdfDownloaded ? true : "Belum",
    payload.downloadTimestamp || "Belum",
    item.no || index + 1,
    item.namaBarang || "",
    item.qty || "",
    item.satuan || "",
    item.supplier || ""
  ]);
}

function buildOrderHeaderRow_(payload, region, itemCount) {
  const form = payload.form || {};
  return [
    `ORDER ${form.noPo || payload.submitTimestamp || ""}`,
    "",
    payload.submitTimestamp || "",
    region,
    form.kodeUnit || payload.kodeUnit || "",
    form.noPo || "",
    form.tglPo || "",
    form.tglCheckIn || "",
    form.namaPeternak || "",
    form.populasi || "",
    form.alamatPeternak || "",
    form.namaPpl || "",
    payload.admin || "Ananda Thufail",
    payload.pdfDownloaded ? true : "Belum",
    payload.downloadTimestamp || "Belum",
    "",
    `${itemCount} barang`,
    "",
    "",
    ""
  ];
}

function applyDownloadStatus_(sheet, startRow, rowCount, downloaded) {
  const downloadCol = HEADERS.indexOf("Download PDF") + 1;
  const range = sheet.getRange(startRow, downloadCol, rowCount, 1);

  if (downloaded) {
    range.insertCheckboxes();
    range.setValue(true);
  } else {
    range.setValue("Belum");
  }
}

function formatSubmitGroup_(sheet, startRow, rowCount, groupKey) {
  const range = sheet.getRange(startRow, 1, rowCount, HEADERS.length);
  const orderRange = sheet.getRange(startRow, 1, 1, HEADERS.length);
  const itemRange = rowCount > 1 ? sheet.getRange(startRow + 1, 1, rowCount - 1, HEADERS.length) : null;

  orderRange
    .setBackground("#1f4e78")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      "#17365d",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

  if (itemRange) {
    itemRange.setBackground("#fce4d6");
  }

  range
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      true,
      "#c55a11",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

  const groupCol = HEADERS.indexOf("Group Submit") + 1;
  sheet.getRange(startRow, groupCol, rowCount, 1).setFontWeight("bold");

  if (rowCount > 2) {
    try {
      sheet.groupRows(startRow + 1, rowCount - 1);
    } catch (error) {
      // Row grouping can fail when the sheet already has overlapping groups.
    }
  }
}

function groupColor_(value) {
  const colors = ["#fff2cc", "#d9ead3", "#d0e0e3", "#fce5cd", "#eadcf8"];
  const text = String(value || "");
  let total = 0;
  for (let i = 0; i < text.length; i += 1) total += text.charCodeAt(i);
  return colors[total % colors.length];
}

function normalizeRegion_(value) {
  const clean = String(value || "").trim();
  const known = REGION_NAMES.find((name) => name.toLowerCase() === clean.toLowerCase());
  return known || "PO";
}

function validateSubmitAuth_(payload, region) {
  const auth = payload.auth || {};
  const role = String(auth.role || "").toLowerCase();

  if (!REGION_NAMES.includes(region)) {
    return "Daerah tidak terdaftar.";
  }

  if (role === "admin") {
    if (String(auth.adminId || "") === ADMIN_ID && String(auth.password || "") === ADMIN_PASSWORD) {
      return "";
    }
    return "Admin ID atau password salah.";
  }

  if (role === "user") {
    const authRegion = normalizeRegion_(auth.region);
    if (authRegion === region && String(auth.password || "") === REGION_PASSWORDS[region]) {
      return "";
    }
    return "Password daerah salah atau tidak sesuai daerah.";
  }

  return "Auth login wajib dikirim.";
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
