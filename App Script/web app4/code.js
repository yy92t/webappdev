(() => {
  const dateTimeFormatter = new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'medium',
  });

  /** @type {{ statusEl: HTMLElement, inputEl: HTMLInputElement, runBtn: HTMLElement } | null} */
  let dom = null;

  function mustGetById(id) {
    const el = document.getElementById(id);
    if (!el) throw new Error(`Missing element: ${id}`);
    return el;
  }

  function setStatus(message) {
    dom.statusEl.textContent = message;
  }

  function run() {
    const value = dom.inputEl.value.trim();
    if (!value) {
      setStatus('Enter a value, then click Run.');
      return;
    }

    const now = new Date();
    setStatus(`You entered: ${value}\nTime: ${dateTimeFormatter.format(now)}`);
  }

  document.addEventListener('DOMContentLoaded', () => {
    dom = {
      statusEl: mustGetById('status'),
      inputEl: /** @type {HTMLInputElement} */ (mustGetById('nameInput')),
      runBtn: mustGetById('runBtn'),
    };

    dom.runBtn.addEventListener('click', run);
    dom.inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') run();
    });

    setStatus('Ready.');
  });
})();
