export function ensureOfficeReady() {
  if (window.Office && typeof Office.onReady === 'function') {
    return new Promise((resolve) => {
      Office.onReady(() => resolve());
    });
  }

  return Promise.resolve();
}

export function getCellValue(cell) {
  if (cell === null || cell === undefined) {
    return '';
  }

  if (typeof cell === 'string') {
    return cell;
  }

  if (typeof cell === 'number') {
    return String(cell);
  }

  return '';
}

export function isFormulaCell(formulas, rowIndex, colIndex) {
  const formula = formulas?.[rowIndex]?.[colIndex];
  return typeof formula === 'string' && formula.startsWith('=');
}
