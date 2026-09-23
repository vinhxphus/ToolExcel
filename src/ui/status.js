const statusEl = document.getElementById('status');

export function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? '#b42318' : '#364152';
}

export function setBusy(isBusy) {
  document.querySelectorAll('button[data-action]').forEach((button) => {
    button.disabled = isBusy;
  });
}
