const form = document.querySelector('#generator-form');
const areaInput = document.querySelector('#area');
const errorNode = document.querySelector('#area-error');
const decrease = document.querySelector('#decrease');
const increase = document.querySelector('#increase');
const submitButton = form.querySelector('button[type="submit"]');
const result = document.querySelector('#result');
const resultMeta = document.querySelector('#result-meta');
const downloadList = document.querySelector('#download-list');
const rates = [25000, 27000, 29000];

function readArea() {
  const normalized = areaInput.value.trim().replace(/\s/g, '').replace(',', '.');
  if (!normalized || !/^\d+(\.\d{1,2})?$/.test(normalized)) return null;
  const value = Number(normalized);
  if (!Number.isFinite(value) || value <= 0 || value > 100000) return null;
  return value;
}

function formatNumber(value, maximumFractionDigits = 2) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits }).format(value);
}

function updateSummary() {
  const area = readArea();
  const totals = document.querySelectorAll('[id^="total-"]');
  totals.forEach((node, index) => {
    node.textContent = area === null ? '—' : `${formatNumber(area * rates[index])} ₸`;
  });
}

function showError(message) {
  errorNode.textContent = message;
  errorNode.hidden = !message;
  areaInput.setAttribute('aria-invalid', message ? 'true' : 'false');
  if (message) areaInput.focus();
}

function stepArea(direction) {
  const current = readArea() ?? 0;
  areaInput.value = formatNumber(Math.max(1, current + direction));
  showError('');
  updateSummary();
}

decrease.addEventListener('click', () => stepArea(-1));
increase.addEventListener('click', () => stepArea(1));
areaInput.addEventListener('input', () => { showError(''); updateSummary(); });
areaInput.addEventListener('keydown', (event) => {
  if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
  event.preventDefault();
  stepArea(event.key === 'ArrowUp' ? 1 : -1);
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const area = readArea();
  if (area === null) {
    showError('Введите площадь от 0,01 до 100 000 м², не более двух знаков после запятой.');
    return;
  }

  showError('');
  result.hidden = true;
  submitButton.disabled = true;
  submitButton.classList.add('is-loading');
  submitButton.setAttribute('aria-busy', 'true');

  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        area: areaInput.value,
        ktru: new FormData(form).get('ktru'),
      }),
    });
    if (response.status === 401) {
      window.location.assign('/login');
      return;
    }
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Не удалось сформировать документы');

    downloadList.replaceChildren();
    payload.files.forEach((file) => {
      const link = document.createElement('a');
      link.className = 'download-link';
      link.href = file.url;
      link.download = file.name;
      const label = document.createElement('span');
      label.textContent = file.description;
      const badge = document.createElement('em');
      badge.textContent = file.kind;
      link.append(label, badge);
      downloadList.append(link);
    });
    resultMeta.textContent = `${payload.ktru} · ${payload.area} м²`;
    result.hidden = false;
    result.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } catch (error) {
    showError(error.message);
  } finally {
    submitButton.disabled = false;
    submitButton.classList.remove('is-loading');
    submitButton.removeAttribute('aria-busy');
  }
});

updateSummary();
