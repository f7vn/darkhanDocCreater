const form = document.querySelector('#generator-form');
const areaInput = document.querySelector('#area');
const errorNode = document.querySelector('#area-error');
const decrease = document.querySelector('#decrease');
const increase = document.querySelector('#increase');
const standFields = document.querySelector('#stand-fields');
const standCountInput = document.querySelector('#stand-count');
const standError = document.querySelector('#stand-error');
const standPriceInputs = [
  document.querySelector('#stand-price-best'),
  document.querySelector('#stand-price-nalibai'),
  document.querySelector('#stand-price-archi'),
];
const vendorNodes = [
  document.querySelector('#vendor-1'),
  document.querySelector('#vendor-2'),
  document.querySelector('#vendor-3'),
];
const rateNodes = [
  document.querySelector('#rate-1'),
  document.querySelector('#rate-2'),
  document.querySelector('#rate-3'),
];
const submitButton = form.querySelector('button[type="submit"]');
const result = document.querySelector('#result');
const resultStatus = document.querySelector('#result-status');
const resultMeta = document.querySelector('#result-meta');
const downloadList = document.querySelector('#download-list');
const rates = [25000, 27000, 29000];
const defaultVendors = ['ИП Абдихалыков', 'ИП Әбдіқұл', 'ТОО Архиспецснабмонтаж'];
const combinedVendors = ['Best Production', 'ИП Нәлібай', 'ТОО Архи'];

function selectedKtru() {
  return new FormData(form).get('ktru');
}

function readArea() {
  const normalized = areaInput.value.trim().replace(/\s/g, '').replace(',', '.');
  if (!normalized || !/^\d+(\.\d{1,2})?$/.test(normalized)) return null;
  const value = Number(normalized);
  if (!Number.isFinite(value) || value <= 0 || value > 100000) return null;
  return value;
}

function readStandCount() {
  const normalized = standCountInput.value.trim().replace(/\s/g, '');
  if (!/^\d+$/.test(normalized)) return null;
  const value = Number(normalized);
  if (!Number.isInteger(value) || value <= 0 || value > 10000) return null;
  return value;
}

function readMoney(input) {
  const normalized = input.value.trim().replace(/\s/g, '').replace(',', '.');
  if (!normalized || !/^\d+(\.\d{1,2})?$/.test(normalized)) return null;
  const value = Number(normalized);
  if (!Number.isFinite(value) || value <= 0 || value > 100000000) return null;
  return value;
}

function formatNumber(value, maximumFractionDigits = 2) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits }).format(value);
}

function updateMode() {
  const combined = selectedKtru() === 'combined';
  standFields.hidden = !combined;
  standFields.querySelectorAll('input').forEach((input) => { input.disabled = !combined; });
  vendorNodes.forEach((node, index) => { node.textContent = (combined ? combinedVendors : defaultVendors)[index]; });
  updateSummary();
}

function updateSummary() {
  const area = readArea();
  const combined = selectedKtru() === 'combined';
  const standCount = readStandCount();
  const totals = document.querySelectorAll('[id^="total-"]');

  totals.forEach((node, index) => {
    if (area === null) {
      node.textContent = '—';
      return;
    }
    if (!combined) {
      node.textContent = `${formatNumber(area * rates[index])} ₸`;
      rateNodes[index].textContent = `${formatNumber(rates[index])} ₸ / м²`;
      return;
    }
    const standPrice = readMoney(standPriceInputs[index]);
    rateNodes[index].textContent = standPrice === null
      ? `${formatNumber(rates[index])} ₸ / м² + цена стенда`
      : `${formatNumber(rates[index])} ₸ / м² + ${formatNumber(standPrice)} ₸ / шт.`;
    node.textContent = standPrice === null || standCount === null
      ? '—'
      : `${formatNumber(area * rates[index] + standCount * standPrice)} ₸`;
  });
}

function showError(message) {
  errorNode.textContent = message;
  errorNode.hidden = !message;
  areaInput.setAttribute('aria-invalid', message ? 'true' : 'false');
  if (message) areaInput.focus();
}

function showStandError(message, input = null) {
  standError.textContent = message;
  standError.hidden = !message;
  [standCountInput, ...standPriceInputs].forEach((node) => node.removeAttribute('aria-invalid'));
  if (message && input) {
    input.setAttribute('aria-invalid', 'true');
    input.focus();
  }
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
form.querySelectorAll('input[name="ktru"]').forEach((input) => input.addEventListener('change', updateMode));
[standCountInput, ...standPriceInputs].forEach((input) => {
  input.addEventListener('input', () => { showStandError(''); updateSummary(); });
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const area = readArea();
  if (area === null) {
    showError('Введите площадь от 0,01 до 100 000 м², не более двух знаков после запятой.');
    return;
  }

  const ktru = selectedKtru();
  const combined = ktru === 'combined';
  const standCount = readStandCount();
  if (combined && standCount === null) {
    showStandError('Укажите целое количество стендов от 1 до 10 000.', standCountInput);
    return;
  }
  if (combined) {
    const missingPriceIndex = standPriceInputs.findIndex((input) => readMoney(input) === null);
    if (missingPriceIndex !== -1) {
      showStandError('Укажите цену одного стенда для каждого поставщика.', standPriceInputs[missingPriceIndex]);
      return;
    }
  }

  showError('');
  showStandError('');
  result.hidden = true;
  submitButton.disabled = true;
  submitButton.classList.add('is-loading');
  submitButton.setAttribute('aria-busy', 'true');

  try {
    const body = { area: areaInput.value, ktru };
    if (combined) {
      body.standCount = standCountInput.value;
      body.standPriceBest = standPriceInputs[0].value;
      body.standPriceNalibai = standPriceInputs[1].value;
      body.standPriceArchi = standPriceInputs[2].value;
    }
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (response.status === 401) {
      window.location.assign('/login');
      return;
    }
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Не удалось сформировать документы');

    downloadList.replaceChildren();
    const downloadLinks = payload.files.map((file) => {
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
      return link;
    });
    resultStatus.textContent = 'Готово — загрузка началась';
    resultMeta.textContent = combined
      ? `${payload.ktru} · ${payload.area} м² · ${payload.standCount} шт.`
      : `${payload.ktru} · ${payload.area} м²`;
    result.hidden = false;
    result.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    downloadLinks.forEach((link, index) => {
      window.setTimeout(() => link.click(), index * 350);
    });
  } catch (error) {
    showError(error.message);
  } finally {
    submitButton.disabled = false;
    submitButton.classList.remove('is-loading');
    submitButton.removeAttribute('aria-busy');
  }
});

updateMode();
