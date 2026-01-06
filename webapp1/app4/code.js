(() => {
  function byId(id) {
    const el = document.getElementById(id);
    if (!el) throw new Error(`Missing element: ${id}`);
    return el;
  }

  function setStatus(message) {
    byId('status').textContent = message;
  }

  function run() {
    const value = byId('nameInput').value.trim();
    if (!value) {
      setStatus('Enter a value, then click Run.');
      return;
    }

    const now = new Date();
    setStatus(`You entered: ${value}\nTime: ${now.toLocaleString()}`);
  }

  document.addEventListener('DOMContentLoaded', () => {
    byId('runBtn').addEventListener('click', run);
    byId('nameInput').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') run();
    });

    setStatus('Ready.');
  });
})();
