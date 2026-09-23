import { getCellValue, isFormulaCell } from './helpers.js';
import { runInExcel, runWorkbookAction } from './runner.js';
import { getCopiedSource, setCopiedSource } from './clipboard-state.js';

export async function highlightDuplicates() {
  await runInExcel('Đã đánh dấu dữ liệu trùng lặp.', async (context, range) => {
    range.format.fill.clear();

    const values = range.values || [];
    const counts = new Map();

    for (const row of values) {
      for (const cell of row) {
        const value = getCellValue(cell).trim();
        if (!value) continue;
        counts.set(value, (counts.get(value) || 0) + 1);
      }
    }

    for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
      for (let colIndex = 0; colIndex < range.columnCount; colIndex += 1) {
        const current = getCellValue(values[rowIndex]?.[colIndex]).trim();
        if (!current) continue;

        if (counts.get(current) > 1) {
          const cell = range.getCell(rowIndex, colIndex);
          cell.format.fill.color = '#FDE7E7';
        }
      }
    }
  });
}

function transformTextCells(range, transform) {
  const values = range.values || [];
  const formulas = range.formulas || [];

  for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
    for (let colIndex = 0; colIndex < range.columnCount; colIndex += 1) {
      const raw = values[rowIndex]?.[colIndex];
      if (typeof raw !== 'string' || isFormulaCell(formulas, rowIndex, colIndex)) {
        continue;
      }

      const next = transform(raw);
      if (next !== raw) {
        range.getCell(rowIndex, colIndex).values = [[next]];
      }
    }
  }
}

export async function trimSpaces() {
  await runInExcel('Đã xoá khoảng trắng thừa.', async (context, range) => {
    transformTextCells(range, (text) => text.replace(/\s+/g, ' ').trim());
  });
}

function convertTextCase(text, mode) {
  if (mode === 'upper') return text.toUpperCase();
  if (mode === 'lower') return text.toLowerCase();
  return text.replace(/\S+/g, (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());
}

const CASE_LABELS = {
  upper: 'Đã chuyển sang chữ hoa.',
  lower: 'Đã chuyển sang chữ thường.',
  title: 'Đã viết hoa chữ cái đầu mỗi từ.',
};

export async function convertCase(mode) {
  await runInExcel(CASE_LABELS[mode], async (context, range) => {
    transformTextCells(range, (text) => convertTextCase(text, mode));
  });
}

export async function applyBorder() {
  await runInExcel('Đã kẻ border theo kiểu outline và inside.', async (context, range) => {
    const borderColor = '#2D3C4D';

    ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideVertical', 'InsideHorizontal'].forEach((borderName) => {
      const border = range.format.borders.getItem(borderName);
      border.style = borderName === 'InsideHorizontal' ? 'Dash' : 'Continuous';
      border.color = borderColor;
      border.weight = 'Thin';
    });
  });
}

const ROW_HEIGHT_STEP = 15;
const ROW_HEIGHT_BASE = 20;

export async function autoHeight() {
  await runInExcel('Đã chuẩn hóa chiều cao hàng.', async (context, range) => {
    range.format.wrapText = true;
    range.format.autofitRows();
    await context.sync();

    const rows = [];
    for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
      const row = range.getRow(rowIndex);
      row.format.load('rowHeight');
      rows.push(row);
    }
    await context.sync();

    rows.forEach((row) => {
      const lineCount = Math.max(1, Math.round(row.format.rowHeight / ROW_HEIGHT_STEP));
      row.format.rowHeight = ROW_HEIGHT_BASE + (lineCount - 1) * ROW_HEIGHT_STEP;
    });
  });
}

const ALIGNMENT_LABELS = {
  center: 'Đã căn giữa (Center Across Selection).',
  justify: 'Đã căn đều (Justify).',
};

export async function applyAlignment(type) {
  await runInExcel(ALIGNMENT_LABELS[type], async (context, range) => {
    range.format.horizontalAlignment = type === 'center' ? 'CenterAcrossSelection' : 'Justify';
    range.format.verticalAlignment = 'Center';
    range.format.wrapText = true;
  });
}

export async function clearFormatting() {
  await runInExcel('Đã xoá toàn bộ định dạng.', async (context, range) => {
    range.clear(Excel.ClearApplyTo.formats);
  });
}

export async function applyZebraStripes() {
  await runInExcel('Đã tô màu xen kẽ dòng.', async (context, range) => {
    range.format.fill.clear();

    for (let rowIndex = 1; rowIndex < range.rowCount; rowIndex += 2) {
      range.getRow(rowIndex).format.fill.color = '#F3E0D3';
    }
  });
}

export async function setPrintArea() {
  await runInExcel('Đã đặt vùng in.', async (context, range) => {
    range.worksheet.pageLayout.setPrintArea(range);
  });
}

export async function setRepeatingHeaderRows() {
  await runInExcel('Đã đặt dòng tiêu đề lặp lại khi in.', async (context, range) => {
    range.worksheet.pageLayout.setPrintTitleRows(range.getEntireRow());
  });
}

export async function convertFormulasToValues() {
  await runInExcel('Đã chuyển công thức thành giá trị.', async (context, range) => {
    range.values = range.values;
  });
}

export async function copySelection() {
  await runInExcel('Đã sao chép vùng chọn.', async (context, range) => {
    range.load('address');
    range.worksheet.load('name');
    await context.sync();

    const address = range.address.includes('!') ? range.address.split('!')[1] : range.address;
    setCopiedSource({ sheetName: range.worksheet.name, address });
  });
}

async function pasteFromCopiedSource(copyType, actionLabel) {
  await runInExcel(actionLabel, async (context, range) => {
    const source = getCopiedSource();
    if (!source) {
      throw new Error('Chưa sao chép vùng nào. Hãy bấm "Sao chép" trước.');
    }

    const sourceSheet = context.workbook.worksheets.getItem(source.sheetName);
    const sourceRange = sourceSheet.getRange(source.address);
    range.copyFrom(sourceRange, copyType);
  });
}

export async function pasteValuesOnly() {
  await pasteFromCopiedSource(Excel.RangeCopyType.values, 'Đã dán chỉ giá trị.');
}

export async function pasteFormulasOnly() {
  await pasteFromCopiedSource(Excel.RangeCopyType.formulas, 'Đã dán công thức.');
}

export async function insertPageNumbers() {
  await runWorkbookAction('Đã chèn số trang vào chân trang khi in.', async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.pageLayout.headersFooters.defaultForAllPages.centerFooter = 'Trang &P/&N';
  });
}

export async function insertPrintHeader() {
  await runWorkbookAction('Đã chèn tiêu đề trang khi in.', async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerFooter = sheet.pageLayout.headersFooters.defaultForAllPages;
    headerFooter.leftHeader = '&F';
    headerFooter.centerHeader = '&"-,Bold"&A';
    headerFooter.rightHeader = '&D';
  });
}

export async function normalizeWorkbookFont() {
  await runWorkbookAction('Đã chuẩn hóa font Cambria cho toàn bộ workbook.', async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();

    sheets.items.forEach((sheet) => {
      sheet.getRange().format.font.name = 'Cambria';
    });
  });
}
