/**
 * Google Apps Script: In hàng loạt theo mẫu in + xuất PDF A6 ngang.
 */
const CONFIG = {
  DATA_SHEET: 'Table_FBS',
  TEMPLATE_SHEET: 'Mau_in',
  OUTPUT_SHEET: 'IN_HANG_LOAT',

  TEMPLATE_RANGE_A1: 'A1:H12',
  BLOCK_HEIGHT: 12,

  HEADERS: {
    sku: 'Ma_sku',
    tenSp: 'Ten_sp',
    phanLoai: 'Phan_loai',
    soLuong: 'So_luong',
    maKien: 'Ma_kien',
    maDon: 'Ma_don',
    supplier: 'Supplier_name',
  },

  CELLS: {
    maDon: 'C3',
    supplier: 'C4',
    maThung: 'E5',
    sku: 'A9',
    tenSp: 'B9',
    phanLoai: 'G9',
    sl: 'H9',
    tong: 'H10',
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

  const out = recreateOutputSheet_(ss, CONFIG.OUTPUT_SHEET);
  renderBatch_(templateSheet, out, jobs);

  const pdfFile = exportJobsToPdfA6Landscape_(templateSheet, jobs);
  showPdfDialog_(pdfFile.getUrl(), pdfFile.getName(), jobs.length);
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

  const out = recreateOutputSheet_(ss, CONFIG.OUTPUT_SHEET);
  renderBatch_(templateSheet, out, jobs);

  SpreadsheetApp.getUi().alert(
    `Đã tạo ${jobs.length} trang in tại sheet "${CONFIG.OUTPUT_SHEET}".\n` +
      'Bạn có thể Ctrl+P để in hoặc dùng menu tạo PDF A6 ngang.'
  );
}

function buildPrintJobs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!dataSheet) throw new Error(`Không tìm thấy sheet dữ liệu: ${CONFIG.DATA_SHEET}`);

  const data = dataSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map((h) => String(h).trim());
  const idx = createHeaderIndex_(headers);

  const jobs = [];
  for (let r = 1; r < data.length; r += 1) {
    const row = data[r];
    const sku = String(getByHeader_(row, idx, CONFIG.HEADERS.sku));
    if (!sku) continue;

    const tongKien = Number(getByHeader_(row, idx, CONFIG.HEADERS.maKien)) || 0;
    const soLuong = Number(getByHeader_(row, idx, CONFIG.HEADERS.soLuong)) || 0;
    const tenSp = String(getByHeader_(row, idx, CONFIG.HEADERS.tenSp) || '');
    const phanLoai = String(getByHeader_(row, idx, CONFIG.HEADERS.phanLoai) || '');
    const maDon = String(getByHeader_(row, idx, CONFIG.HEADERS.maDon) || '');
    const supplier = String(getByHeader_(row, idx, CONFIG.HEADERS.supplier) || '');

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

  for (let c = 1; c <= tplRange.getNumColumns(); c += 1) {
    outSheet.setColumnWidth(c, templateSheet.getColumnWidth(c));
  }

  jobs.forEach((job, i) => {
    const blockStart = i * CONFIG.BLOCK_HEIGHT + 1;
    const target = outSheet.getRange(blockStart, 1, tplRange.getNumRows(), tplRange.getNumColumns());
    tplRange.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    for (let rr = 0; rr < CONFIG.BLOCK_HEIGHT; rr += 1) {
      outSheet.setRowHeight(blockStart + rr, templateSheet.getRowHeight(rr + 1));
    }

    fillTemplateCells_(outSheet, blockStart, job);
  });
}

function fillTemplateCells_(sheet, blockStartRow, job) {
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.maDon, job.maDon);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.supplier, job.supplier);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.maThung, `${job.soTrang} / ${job.tongKien}`);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.sku, job.sku);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.tenSp, job.tenSp);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.phanLoai, job.phanLoai);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.sl, job.sl);
  setCellInBlock_(sheet, blockStartRow, CONFIG.CELLS.tong, job.sl);
}

function exportJobsToPdfA6Landscape_(templateSheet, jobs) {
  const tempSs = SpreadsheetApp.create(`TMP_PRINT_${Date.now()}`);
  const defaultSheet = tempSs.getSheets()[0];
  tempSs.deleteSheet(defaultSheet);

  jobs.forEach((job, i) => {
    const pageSheet = templateSheet.copyTo(tempSs).setName(`Tem_${i + 1}`);
    fillTemplateCells_(pageSheet, 1, job);
  });

  SpreadsheetApp.flush();

  const fileName = `${CONFIG.PDF.filePrefix}_${formatDateTime_(new Date())}.pdf`;
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

function showPdfDialog_(fileUrl, fileName, pageCount) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;padding:8px">` +
      `<p>Đã tạo <b>${pageCount}</b> trang tem và xuất PDF <b>${fileName}</b>.</p>` +
      `<p><a href="${fileUrl}" target="_blank">Bấm vào đây để mở / tải PDF</a></p>` +
      `<p>File cũng đã lưu trong Google Drive của bạn.</p>` +
      `</div>`
  ).setWidth(420).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xuất PDF A6 ngang thành công');
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
}

function setCellInBlock_(sheet, blockStartRow, cellA1, value) {
  const parsed = parseA1_(cellA1);
  const row = blockStartRow + parsed.row - 1;
  sheet.getRange(row, parsed.col).setValue(value);
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
  return ss.insertSheet(name);
}

function createHeaderIndex_(headers) {
  return headers.reduce((acc, h, i) => {
    acc[String(h).trim()] = i;
    return acc;
  }, {});
}

function getByHeader_(row, headerIndex, headerName) {
  const idx = headerIndex[headerName];
  if (idx === undefined) return '';
  return row[idx];
}
