import { ensureOfficeReady } from './helpers.js';
import { setStatus, setBusy } from '../ui/status.js';

export async function runInExcel(action, callback) {
  await ensureOfficeReady();

  setBusy(true);
  try {
    let executed = false;

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('rowCount,columnCount,values,formulas');
      await context.sync();

      if (!range.rowCount || !range.columnCount) {
        setStatus('Vui lòng chọn một khối dữ liệu trước khi thực hiện.', true);
        return;
      }

      await callback(context, range);
      await context.sync();
      executed = true;
    });

    if (executed) {
      setStatus(action);
    }
  } catch (error) {
    console.error(error);
    setStatus(`Lỗi: ${error.message || 'Không thể thực hiện thao tác.'}`, true);
  } finally {
    setBusy(false);
  }
}

export async function runWorkbookAction(action, callback) {
  await ensureOfficeReady();

  setBusy(true);
  try {
    await Excel.run(async (context) => {
      await callback(context);
      await context.sync();
    });

    setStatus(action);
  } catch (error) {
    console.error(error);
    setStatus(`Lỗi: ${error.message || 'Không thể thực hiện thao tác.'}`, true);
  } finally {
    setBusy(false);
  }
}
