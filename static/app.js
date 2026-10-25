const form = document.querySelector('#generator-form');
const areaFields = document.querySelector('#area-fields');
const areaInput = document.querySelector('#area');
const areaPriceInput = document.querySelector('#area-price');
const standFields = document.querySelector('#stand-fields');
const standCountInput = document.querySelector('#stand-count');
const standPriceInput = document.querySelector('#stand-price');
const markupInputs = [document.querySelector('#markup-1'), document.querySelector('#markup-2')];
const errorNode = document.querySelector('#form-error');
const submitButton = form.querySelector('button[type="submit"]');
const result = document.querySelector('#result');
const resultStatus = document.querySelector('#result-status');
const resultMeta = document.querySelector('#result-meta');
const downloadList = document.querySelector('#download-list');
const allInputs = [areaInput, areaPriceInput, standCountInput, standPriceInput, ...markupInputs];
const vendors = {
  default: ['ИП Абдихалыков', 'ИП Әбдіқұл', 'ТОО Архиспецснабмонтаж'],
  combined: ['Best Production', 'ИП Нәлібай', 'ТОО Архи'],
  stands: ['ИП Әбдіқұл', 'ТОО Торговая сеть Казахстан', 'ИП Андрей и Ко'],
};

function selectedKtru() {
  return form.querySelector('input[name="ktru"]:checked').value;
}

function readDecimal(input, minimum, maximum, integer = false) {
  const text = input.value.trim().replace(/\s/g, '').replace(',', '.');
  if (!(integer ? /^\d+$/ : /^\d+(\.\d{1,2})?$/).test(text)) return null;
  const value = Number(text);
  return Number.isFinite(value) && value >= minimum && value <= maximum ? value : null;
}

function formatNumber(value) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 2 }).format(value);
}

// Calculate percentage prices in integer kopecks, using half-up rounding like the server.
function markedPrice(base, percent) {
  const baseCents = Math.round(base * 100);
  const percentHundredths = Math.round(percent * 100);
  return Math.floor((baseCents * (10000 + percentHundredths) + 5000) / 10000) / 100;
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function showError(message = '', input = null) {
  errorNode.textContent = message;
  errorNode.hidden = !message;
  allInputs.forEach((field) => field.removeAttribute('aria-invalid'));
  if (input) {
    input.setAttribute('aria-invalid', 'true');
    input.focus();
  }
}

function updateMode() {
  const mode = selectedKtru();
  areaFields.hidden = mode === 'stands';
  areaFields.querySelectorAll('input, button').forEach((input) => { input.disabled = mode === 'stands'; });
  standFields.hidden = !['stands', 'combined'].includes(mode);
  standFields.classList.toggle('standalone', mode === 'stands');
  standFields.querySelectorAll('input').forEach((input) => { input.disabled = standFields.hidden; });
  const names = vendors[mode] || vendors.default;
  names.forEach((name, index) => { document.querySelector(`#vendor-${index + 1}`).textContent = name; });
  markupInputs.forEach((input, index) => { input.setAttribute('aria-label', `Надбавка: ${names[index + 1]}`); });
  showError();
  result.hidden = true;
  updateSummary();
}

function updateSummary() {
  const mode = selectedKtru();
  const area = readDecimal(areaInput, 0.01, 100000);
  const areaBase = readDecimal(areaPriceInput, 0.01, 100000000);
  const count = readDecimal(standCountInput, 1, 10000, true);
  const standBase = readDecimal(standPriceInput, 0.01, 100000000);
  const percents = [0, ...markupInputs.map((input) => readDecimal(input, 0, 1000))];
  percents.forEach((percent, index) => {
    const areaRate = areaBase !== null && percent !== null ? markedPrice(areaBase, percent) : null;
    const standRate = standBase !== null && percent !== null ? markedPrice(standBase, percent) : null;
    const rateText = [];
    let total = 0;
    let valid = percent !== null;
    if (mode !== 'stands') {
      valid &&= area !== null && areaRate !== null;
      rateText.push(areaRate === null ? 'Укажите цену за м²' : `${formatNumber(areaRate)} ₸ / м²`);
      if (valid) total += roundMoney(area * areaRate);
    }
    if (mode === 'stands' || mode === 'combined') {
      valid &&= count !== null && standRate !== null;
      rateText.push(standRate === null ? 'Укажите цену за стенд' : `${formatNumber(standRate)} ₸ / шт.`);
      if (valid) total += count * standRate;
    }
    document.querySelector(`#rate-${index + 1}`).textContent = rateText.join(' + ');
    document.querySelector(`#total-${index + 1}`).textContent = valid ? `${formatNumber(roundMoney(total))} ₸` : '—';
  });
}

function stepArea(direction) {
  const current = readDecimal(areaInput, 0.01, 100000) ?? 0;
  areaInput.value = formatNumber(Math.min(100000, Math.max(0.01, current + direction)));
  showError();
  result.hidden = true;
  updateSummary();
}

document.querySelector('#decrease').addEventListener('click', () => stepArea(-1));
document.querySelector('#increase').addEventListener('click', () => stepArea(1));
areaInput.addEventListener('keydown', (event) => {
  if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
  event.preventDefault();
  stepArea(event.key === 'ArrowUp' ? 1 : -1);
});
form.querySelectorAll('input[name="ktru"]').forEach((input) => input.addEventListener('change', updateMode));
allInputs.forEach((input) => {
  let selectOnNextClick = false;
  input.closest('label').addEventListener('pointerdown', () => {
    selectOnNextClick = document.activeElement !== input;
  });
  input.addEventListener('focus', () => input.select());
  input.addEventListener('click', () => {
    if (selectOnNextClick) {
      input.select();
      selectOnNextClick = false;
    }
  });
});
allInputs.forEach((input) => input.addEventListener('input', () => {
  showError();
  result.hidden = true;
  updateSummary();
}));

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const mode = selectedKtru();
  const fields = [];
  if (mode !== 'stands') {
    fields.push([areaInput, 0.01, 100000, false, 'Введите площадь от 0,01 до 100 000 м².']);
    fields.push([areaPriceInput, 0.01, 100000000, false, 'Укажите цену за м² больше нуля.']);
  }
  if (mode === 'stands' || mode === 'combined') {
    fields.push([standCountInput, 1, 10000, true, 'Укажите целое количество стендов от 1 до 10 000.']);
    fields.push([standPriceInput, 0.01, 100000000, false, 'Укажите цену за один стенд больше нуля.']);
  }
  markupInputs.forEach((input) => fields.push([input, 0, 1000, false, 'Введите надбавку от 0 до 1 000%. Допустимы два знака после запятой.']));
  const invalid = fields.find(([input, min, max, integer]) => readDecimal(input, min, max, integer) === null);
  if (invalid) {
    showError(invalid[4], invalid[0]);
    return;
  }
  const body = Object.fromEntries(new FormData(form));
  showError();
  result.hidden = true;
  submitButton.disabled = true;
  form.querySelectorAll('fieldset').forEach((fieldset) => { fieldset.disabled = true; });
  markupInputs.forEach((input) => { input.disabled = true; });
  submitButton.classList.add('is-loading');
  submitButton.setAttribute('aria-busy', 'true');
  try {
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
    const quantities = [];
    if (payload.area) quantities.push(`${payload.area} м²`);
    if (payload.standCount) quantities.push(`${payload.standCount} шт.`);
    resultMeta.textContent = [payload.ktru, ...quantities].join(' · ');
    result.hidden = false;
    result.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    downloadLinks.forEach((link, index) => window.setTimeout(() => link.click(), index * 350));
  } catch (error) {
    showError(error.message);
  } finally {
    form.querySelectorAll('fieldset').forEach((fieldset) => { fieldset.disabled = false; });
    markupInputs.forEach((input) => { input.disabled = false; });
    submitButton.disabled = false;
    submitButton.classList.remove('is-loading');
    submitButton.removeAttribute('aria-busy');
  }
});

updateMode();
