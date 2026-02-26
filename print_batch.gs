/**
 * Google Apps Script: In hàng loạt theo mẫu in + xuất PDF A6 ngang.
 */
const CONFIG = {
  DATA_SHEET: 'Table_FBS',
  TEMPLATE_SHEET: 'Mau_in',
  OUTPUT_SHEET: 'IN_HANG_LOAT',
  OUTPUT_MIN_ROWS: 30000,

  // Mặc định theo mẫu hiện tại: A6 ngang, vùng A1:E10 (đã bỏ dòng người in/ngày in).
  TEMPLATE_RANGE_A1: 'A1:E10',
  BLOCK_HEIGHT: 10,

  // Header aliases để đọc dữ liệu bền vững (không phụ thuộc hoa/thường/dấu/_/space).
  HEADERS: {
    sku: ['Ma_sku', 'Mã SKU', 'sku'],
    tenSp: ['Ten_sp', 'Tên hàng', 'Ten hang'],
    phanLoai: ['Phan_loai', 'Phân loại', 'Phan loai'],
    soLuong: ['So_luong', 'Số lượng', 'So luong'],
    maKien: ['Ma_kien', 'Mã kiện', 'Ma kien'],
    maDon: ['Ma_don', 'Mã đơn', 'Ma don'],
    supplier: ['Supplier_name', 'Supplier name', 'Supplier'],
  },

  // Toạ độ ô theo đúng layout mẫu trong ảnh.
  CELLS: {
    maDon: 'B3',
    supplier: 'B4',
    maThung: 'C5',
    sku: 'A8',
    tenSp: 'B8',
    phanLoai: 'D8',
    sl: 'E8',
    tong: 'E9',
  },

  PDF: {
    filePrefix: 'FBS_LABEL_A6',
    paperSize: 'A6',
    landscape: true,
  },
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖨️ In tem')
    .addItem('Tạo sheet in hàng loạt', 'inHangLoatTheoMau')
    .addItem('Tạo PDF A6 ngang (download)', 'inHangLoatVaXuatPdf')
    .addToUi();
}

function inHangLoatVaXuatPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!templateSheet) throw new Error(`Không tìm thấy sheet mẫu: ${CONFIG.TEMPLATE_SHEET}`);

  const jobs = buildPrintJobs_();
  if (!jobs.length) {
    SpreadsheetApp.getUi().alert('Không có dòng hợp lệ (Ma_kien > 0) để in.');
    return;
  }

  const startedAt = Date.now();
  const out = recreateOutputSheet_(ss, CONFIG.OUTPUT_SHEET);
  const renderedArea = renderBatch_(templateSheet, out, jobs);
  const renderSeconds = ((Date.now() - startedAt) / 1000).toFixed(2);

  const fileName = `${CONFIG.PDF.filePrefix}_${formatDateTime_(new Date())}.pdf`;
  const pdfFile = exportRenderedSheetToPdf_(out, fileName, renderedArea);
  showPdfDialog_(pdfFile.getUrl(), pdfFile.getName(), jobs.length, renderSeconds);
}

function inHangLoatTheoMau() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!templateSheet) throw new Error(`Không tìm thấy sheet mẫu: ${CONFIG.TEMPLATE_SHEET}`);

  const jobs = buildPrintJobs_();
  if (!jobs.length) {
    SpreadsheetApp.getUi().alert('Không có dòng hợp lệ (Ma_kien > 0) để in.');
    return;
  }

  const startedAt = Date.now();
  const out = recreateOutputSheet_(ss, CONFIG.OUTPUT_SHEET);
  renderBatch_(templateSheet, out, jobs);
  const renderSeconds = ((Date.now() - startedAt) / 1000).toFixed(2);

  SpreadsheetApp.getUi().alert(
    `Đã tạo ${jobs.length} trang in tại sheet "${CONFIG.OUTPUT_SHEET}" trong ${renderSeconds}s.
` +
      'Bạn có thể Ctrl+P để in hoặc dùng menu tạo PDF A6 ngang.'
  );
}

function buildPrintJobs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!dataSheet) throw new Error(`Không tìm thấy sheet dữ liệu: ${CONFIG.DATA_SHEET}`);

  const data = dataSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map((h) => String(h || ''));
  const idx = createHeaderIndex_(headers);

  const jobs = [];
  for (let r = 1; r < data.length; r += 1) {
    const row = data[r];
    const sku = String(getByHeaderAliases_(row, idx, CONFIG.HEADERS.sku) || '');
    if (!sku) continue;

    const tongKien = Number(getByHeaderAliases_(row, idx, CONFIG.HEADERS.maKien)) || 0;
    const soLuong = Number(getByHeaderAliases_(row, idx, CONFIG.HEADERS.soLuong)) || 0;
    const tenSp = String(getByHeaderAliases_(row, idx, CONFIG.HEADERS.tenSp) || '');
    const phanLoai = String(getByHeaderAliases_(row, idx, CONFIG.HEADERS.phanLoai) || '');
    const maDon = String(getByHeaderAliases_(row, idx, CONFIG.HEADERS.maDon) || '');
    const supplier = String(getByHeaderAliases_(row, idx, CONFIG.HEADERS.supplier) || '');

    if (tongKien <= 0) continue;

    const slMoiKien = Math.ceil(soLuong / tongKien);
    for (let i = 1; i <= tongKien; i += 1) {
      const slKien = i === tongKien ? soLuong - slMoiKien * (tongKien - 1) : slMoiKien;
      jobs.push({
        maDon,
        supplier,
        sku,
        tenSp,
        phanLoai,
        tongKien,
        soTrang: i,
        sl: slKien,
      });
    }
  }

  return jobs;
}

function renderBatch_(templateSheet, outSheet, jobs) {
  const tplRange = templateSheet.getRange(CONFIG.TEMPLATE_RANGE_A1);
  const tplRows = tplRange.getNumRows();
  const tplCols = tplRange.getNumColumns();
  const totalRows = jobs.length * tplRows;

  for (let c = 1; c <= tplCols; c += 1) {
    outSheet.setColumnWidth(c, templateSheet.getColumnWidth(c));
  }

  // Tăng tốc: copy mẫu 1 lần cho toàn bộ vùng đích (Google Sheets tự lặp block mẫu).
  const targetRange = outSheet.getRange(1, 1, totalRows, tplCols);
  tplRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // Tăng tốc: ghi dữ liệu động bằng 1 lần setValues thay vì set từng ô.
  const values = targetRange.getValues();
  jobs.forEach((job, i) => fillTemplateValuesInMemory_(values, i * tplRows, job));
  targetRange.setValues(values);

  return { totalRows, totalCols: tplCols };
}

function fillTemplateValuesInMemory_(values, blockStartIndex, job) {
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.maDon, job.maDon);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.supplier, job.supplier);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.maThung, `${job.soTrang} / ${job.tongKien}`);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.sku, job.sku);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.tenSp, job.tenSp);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.phanLoai, job.phanLoai);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.sl, job.sl);
  setValueInMemory_(values, blockStartIndex, CONFIG.CELLS.tong, job.sl);
}

function setValueInMemory_(values, blockStartIndex, cellA1, value) {
  const parsed = parseA1_(cellA1);
  const rowIndex = blockStartIndex + parsed.row - 1;
  const colIndex = parsed.col - 1;
  values[rowIndex][colIndex] = value;
}

function exportRenderedSheetToPdf_(renderedSheet, fileName, renderedArea) {
  const tempSs = SpreadsheetApp.create(`TMP_PRINT_${Date.now()}`);
  const copiedSheet = renderedSheet.copyTo(tempSs).setName(CONFIG.OUTPUT_SHEET);

  const defaultSheet = tempSs.getSheets().find((sh) => sh.getSheetId() !== copiedSheet.getSheetId());
  if (defaultSheet) tempSs.deleteSheet(defaultSheet);

  const maxRows = copiedSheet.getMaxRows();
  if (renderedArea && renderedArea.totalRows < maxRows) {
    copiedSheet.deleteRows(renderedArea.totalRows + 1, maxRows - renderedArea.totalRows);
  }

  const maxCols = copiedSheet.getMaxColumns();
  if (renderedArea && renderedArea.totalCols < maxCols) {
    copiedSheet.deleteColumns(renderedArea.totalCols + 1, maxCols - renderedArea.totalCols);
  }

  SpreadsheetApp.flush();

  const blob = exportSpreadsheetAsPdfBlob_(tempSs.getId(), fileName);
  const pdfFile = DriveApp.createFile(blob);
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  return pdfFile;
}

function exportSpreadsheetAsPdfBlob_(spreadsheetId, fileName) {
  const base = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export`;
  const params = {
    format: 'pdf',
    size: CONFIG.PDF.paperSize,
    portrait: String(!CONFIG.PDF.landscape),
    fitw: 'true',
    sheetnames: 'false',
    printtitle: 'false',
    pagenum: 'UNDEFINED',
    gridlines: 'false',
    fzr: 'false',
    top_margin: '0.25',
    bottom_margin: '0.25',
    left_margin: '0.25',
    right_margin: '0.25',
  };

  const query = Object.keys(params)
    .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join('&');

  const response = UrlFetchApp.fetch(`${base}?${query}`, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`Xuất PDF thất bại. HTTP ${code}: ${response.getContentText()}`);
  }

  return response.getBlob().setName(fileName);
}

function showPdfDialog_(fileUrl, fileName, pageCount, renderSeconds) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;padding:8px">` +
      `<p>Đã tạo <b>${pageCount}</b> trang tem trong <b>${renderSeconds}s</b> và xuất PDF <b>${fileName}</b>.</p>` +
      `<p><a href="${fileUrl}" target="_blank">Bấm vào đây để mở / tải PDF</a></p>` +
      `<p>File cũng đã lưu trong Google Drive của bạn.</p>` +
      `</div>`
  ).setWidth(420).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xuất PDF A6 ngang thành công');
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
}

function parseA1_(a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+)$/i);
  if (!m) throw new Error(`A1 không hợp lệ: ${a1}`);
  return { row: Number(m[2]), col: colToNumber_(m[1].toUpperCase()) };
}

function colToNumber_(col) {
  let n = 0;
  for (let i = 0; i < col.length; i += 1) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

function recreateOutputSheet_(ss, name) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);

  const sheet = ss.insertSheet(name);
  const missingRows = CONFIG.OUTPUT_MIN_ROWS - sheet.getMaxRows();
  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  }

  return sheet;
}

function createHeaderIndex_(headers) {
  return headers.reduce((acc, h, i) => {
    acc[normalizeHeader_(h)] = i;
    return acc;
  }, {});
}

function getByHeaderAliases_(row, headerIndex, aliases) {
  for (let i = 0; i < aliases.length; i += 1) {
    const idx = headerIndex[normalizeHeader_(aliases[i])];
    if (idx !== undefined) return row[idx];
  }
  return '';
}

function normalizeHeader_(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase();
}
