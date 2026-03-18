const form = document.getElementById('generator-form');
const inputText = document.getElementById('input-text');
const briefFile = document.getElementById('brief-file');
const descriptionFile = document.getElementById('description-file');
const submitBtn = document.getElementById('submit-btn');
const clearBtn = document.getElementById('clear-btn');

const statusBox = document.getElementById('status');
const warningsBox = document.getElementById('warnings');
const sourceChip = document.getElementById('source-chip');
const REQUEST_TIMEOUT_MS = 180000;

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.style.borderColor = isError ? '#d27b7b' : '';
  statusBox.style.background = isError ? '#fff1f1' : '#fff';
}

function setSourceLabel(source) {
  const map = {
    openai: 'Режим: OpenAI',
    fallback: 'Режим: Fallback'
  };
  sourceChip.textContent = map[source] || 'Режим: очікування';
}

function renderWarnings(warnings = []) {
  warningsBox.innerHTML = '';
  warnings.forEach((warning) => {
    const el = document.createElement('div');
    el.className = 'warning-item';
    el.textContent = warning;
    warningsBox.appendChild(el);
  });
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  const text = inputText.value.trim();
  if (!text && !briefFile.files[0]) {
    setStatus('Надайте текст брифу або завантажте файл.', true);
    return;
  }

  const formData = new FormData();
  formData.append('input_text', text);

  if (briefFile.files[0]) {
    formData.append('brief_file', briefFile.files[0]);
  }

  if (descriptionFile.files[0]) {
    formData.append('description_file', descriptionFile.files[0]);
  }

  submitBtn.disabled = true;
  renderWarnings([]);
  setStatus('Формування документа ТЗ...');
  let timeoutId;

  try {
    const controller = new AbortController();
    timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    const response = await fetch('/api/generate-docx', {
      method: 'POST',
      body: formData,
      signal: controller.signal
    });

    if (!response.ok) {
      const maybeJson = await response.json().catch(() => ({}));
      throw new Error(maybeJson.error || 'Невідома помилка сервера');
    }

    const source = response.headers.get('X-Generator-Source') || '';
    const warningCount = Number(response.headers.get('X-Generator-Warning-Count') || '0');
    const warningMessage = response.headers.get('X-Generator-Warning') || '';
    const duration = response.headers.get('X-Generator-Duration-Seconds') || '';
    setSourceLabel(source);

    if (warningCount > 0) {
      renderWarnings([warningMessage || 'Генерація виконана у fallback-режимі. Перевірте OPENAI_API_KEY та модель.']);
    }

    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="?([^";]+)"?/i);
    const fileName = (match && match[1]) ? match[1] : 'TZ_output.docx';

    downloadBlob(blob, fileName);
    setStatus(`Документ сформовано і завантажено${duration ? ` за ${duration} c` : ''}.`);
  } catch (error) {
    if (error.name === 'AbortError') {
      setStatus('Помилка: перевищено ліміт очікування 180 секунд. Спробуйте ще раз або зменште обсяг брифу.', true);
    } else {
      setStatus(`Помилка: ${error.message}`, true);
    }
    setSourceLabel();
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
    submitBtn.disabled = false;
  }
});

clearBtn.addEventListener('click', () => {
  inputText.value = '';
  briefFile.value = '';
  descriptionFile.value = '';
  warningsBox.innerHTML = '';
  setSourceLabel();
  setStatus('Поля очищено.');
});
