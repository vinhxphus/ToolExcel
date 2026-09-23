import { setStatus } from './ui/status.js';
import {
  highlightDuplicates,
  trimSpaces,
  convertCase,
  applyBorder,
  autoHeight,
  applyAlignment,
  clearFormatting,
  applyZebraStripes,
  setPrintArea,
  setRepeatingHeaderRows,
  convertFormulasToValues,
  copySelection,
  pasteValuesOnly,
  pasteFormulasOnly,
  insertPageNumbers,
  insertPrintHeader,
  normalizeWorkbookFont,
} from './excel/actions.js';

const ACTIONS = {
  duplicate: highlightDuplicates,
  trim: trimSpaces,
  upper: () => convertCase('upper'),
  lower: () => convertCase('lower'),
  title: () => convertCase('title'),
  formulavalue: convertFormulasToValues,
  copy: copySelection,
  pastevalues: pasteValuesOnly,
  pasteformulas: pasteFormulasOnly,
  border: applyBorder,
  height: autoHeight,
  center: () => applyAlignment('center'),
  justify: () => applyAlignment('justify'),
  clear: clearFormatting,
  zebra: applyZebraStripes,
  printarea: setPrintArea,
  printtitle: setRepeatingHeaderRows,
  pagenumber: insertPageNumbers,
  printheader: insertPrintHeader,
  font: normalizeWorkbookFont,
};

document.addEventListener('DOMContentLoaded', () => {
  Object.entries(ACTIONS).forEach(([action, handler]) => {
    const button = document.querySelector(`[data-action="${action}"]`);
    if (!button) return;

    button.addEventListener('click', () => handler());
  });

  setStatus('Sẵn sàng. Chọn vùng dữ liệu trong Excel.');
});
