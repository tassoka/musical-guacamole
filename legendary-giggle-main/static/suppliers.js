const addHandlerButton = document.getElementById('add-handler-button');
const handlersContainer = document.getElementById('supplier-handlers');
const tabsContainer = document.getElementById('supplier-tabs');
const emptyStateCard = document.getElementById('supplier-empty');
const handlerDialog = document.getElementById('handler-dialog');
const handlerForm = document.getElementById('handler-form');
const previewButton = document.getElementById('preview-source-button');
const mappingSection = document.getElementById('mapping-section');
const sheetList = document.getElementById('sheet-list');
const sourceFileInput = handlerForm.elements.source_file;
const wholesaleMultiplierInput = handlerForm.elements.wholesale_multiplier;
const recommendedMultiplierInput = handlerForm.elements.recommended_multiplier;
const saveHandlerButton = document.getElementById('save-handler-button');
const cancelButton = handlerForm.querySelector('button[value="cancel"]');
const handlerDialogTitle = handlerDialog.querySelector('h3');
const toast = document.getElementById('supplier-toast');
const addDialog = document.getElementById('supplier-add-dialog');
const addForm = document.getElementById('supplier-add-form');
const addHeading = document.getElementById('supplier-add-heading');
const addInfo = document.getElementById('supplier-add-info');
const addSubmitButton = document.getElementById('supplier-add-submit');
const addCancelButton = addForm.querySelector('button[value="cancel"]');
const searchInput = document.getElementById('supplier-search');
const refreshAllButton = document.getElementById('refresh-all-button');
const negativeFilterToggle = document.getElementById('supplier-filter-negative');
const changesDialog = document.getElementById('changes-dialog');
const changesDialogList = document.getElementById('changes-dialog-list');
const changesDialogEmpty = document.getElementById('changes-dialog-empty');

let previewSheets = [];
let previewContext = null;
let autoRefreshTimer;
let toastTimer;
let addItemContext = null;
let editingHandlerId = null;
const sheetControls = new Map();

const pageRoot = document.querySelector('[data-active-handler-id]');
const initialActiveHandlerId = pageRoot ? pageRoot.dataset.activeHandlerId : '';

const handlersState = {
  list: [],
  activeId: initialActiveHandlerId ? String(initialActiveHandlerId) : null
};

const AUTO_RELOAD_INTERVAL = 5 * 60 * 1000;

const autoMappingKeywords = {
  name: ['наименование', 'товар', 'название', 'product', 'item'],
  wholesale_price: ['опт', 'закуп', 'закупочная', 'оптовая', 'wholesale'],
  recommended_price: ['ррц', 'розница', 'цена', 'price', 'rrc'],
  stock: ['остаток', 'кол-во', 'количество', 'stock', 'qty']
};

const changeFieldLabels = {
  wholesale_price: 'Опт',
  recommended_price: 'РРЦ',
  stock: 'Остаток'
};

const placeholderValuePattern = /^[\-\u2010\u2012\u2013\u2014\u2212]+$/;

let currentSearchTerm = '';
const filtersState = {
  hideNegativeStock: false
};

if (searchInput) {
  searchInput.disabled = true;
}

if (negativeFilterToggle) {
  negativeFilterToggle.disabled = true;
}

if (refreshAllButton) {
  refreshAllButton.disabled = true;
}

function showToast(message, type = 'info') {
  toast.textContent = message;
  toast.dataset.type = type;
  toast.style.backgroundColor = type === 'error' ? '#dc2626' : type === 'success' ? '#16a34a' : '#111827';
  toast.hidden = false;
  requestAnimationFrame(() => {
    toast.classList.add('show');
  });
  clearTimeout(toastTimer);
  toastTimer = setTimeout(hideToast, 3000);
}

function hideToast() {
  toast.classList.remove('show');
  toast.addEventListener(
    'transitionend',
    () => {
      toast.hidden = true;
    },
    { once: true }
  );
}

function normalizePriceValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  const normalized = String(value).replace(/\s+/g, '').replace(',', '.');
  const parsed = Number.parseFloat(normalized);
  if (Number.isFinite(parsed)) {
    return Math.round(parsed * 100) / 100;
  }
  return '';
}

function toInputNumber(value) {
  const normalized = normalizePriceValue(value);
  if (normalized === '') {
    return '';
  }
  return normalized.toString();
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const text = String(value)
    .replace(/\u00a0/g, ' ')
    .trim();
  if (!text) {
    return '';
  }
  const collapsed = text.replace(/\s+/g, '');
  if (placeholderValuePattern.test(collapsed)) {
    return '';
  }
  return text;
}

function hasMeaningfulContent(value) {
  return normalizeCellValue(value) !== '';
}

function createTextCell(content, className = '') {
  const td = document.createElement('td');
  if (className) {
    td.className = className;
  }
  const normalized = normalizeCellValue(content);
  const display = normalized || '—';
  td.textContent = display;
  if (normalized) {
    td.title = normalized;
  }
  return td;
}

function createCategoryRow(category) {
  const tr = document.createElement('tr');
  tr.className = 'supplier-table__category-row';
  const td = document.createElement('td');
  td.colSpan = 5;
  td.textContent = normalizeCellValue(category) || 'Без категории';
  tr.appendChild(td);
  return tr;
}

function createItemRow(handler, row) {
  const name = normalizeCellValue(row.name);
  const hasValues = [row.wholesale_price, row.recommended_price, row.stock]
    .some((value) => hasMeaningfulContent(value));
  if (!name && !hasValues) {
    return null;
  }

  const tr = document.createElement('tr');
  tr.className = 'supplier-table__item-row';
  tr.appendChild(createTextCell(name, 'supplier-table__name'));
  tr.appendChild(createTextCell(row.wholesale_price, 'supplier-table__number'));
  tr.appendChild(createTextCell(row.recommended_price, 'supplier-table__number'));
  tr.appendChild(createTextCell(row.stock, 'supplier-table__number'));

  const actions = createTextCell('', 'supplier-table__actions');
  const addButton = document.createElement('button');
  addButton.type = 'button';
  addButton.className = 'supplier-table__add-button';
  addButton.textContent = '+';
  addButton.title = 'Добавить товар на склад';
  addButton.setAttribute('aria-label', name ? `Добавить «${name}» на склад` : 'Добавить товар на склад');
  addButton.addEventListener('click', () => {
    openAddItemDialog(handler, row);
  });
  if (!name) {
    addButton.disabled = true;
    addButton.title = 'В прайс-листе нет наименования';
  }
  actions.appendChild(addButton);
  tr.appendChild(actions);

  return tr;
}

function resetPreview() {
  previewSheets = [];
  previewContext = null;
  mappingSection.hidden = true;
  sheetControls.clear();
  if (sheetList) {
    sheetList.innerHTML = '';
  }
  updateSaveButtonState();
}

function createMappingSelect(columns, field, existingValue = '') {
  const select = document.createElement('select');
  select.dataset.field = field;
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Не выбрано';
  select.appendChild(placeholder);
  columns.forEach((column) => {
    const option = document.createElement('option');
    option.value = column.letter;
    const header = column.header ? ` — ${column.header}` : '';
    option.textContent = `${column.letter}${header}`;
    select.appendChild(option);
  });
  select.value = existingValue || '';
  select.addEventListener('change', updateSaveButtonState);
  return select;
}

function suggestColumn(sheet, field) {
  const keywords = autoMappingKeywords[field];
  if (!keywords || !sheet || !Array.isArray(sheet.columns)) {
    return '';
  }
  const match = sheet.columns.find((column) => {
    const header = (column.header || '').toLowerCase();
    return keywords.some((keyword) => header.includes(keyword));
  });
  return match ? match.letter : '';
}

function buildSheetCard(sheet, index, existingMapping = null) {
  const container = document.createElement('div');
  container.className = 'sheet-card';
  container.dataset.sheetName = sheet.name;

  const header = document.createElement('label');
  header.className = 'sheet-card__header';

  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.className = 'sheet-card__checkbox';

  const title = document.createElement('span');
  title.className = 'sheet-card__title';
  title.textContent = sheet.name || `Лист ${index + 1}`;

  header.append(checkbox, title);
  container.appendChild(header);

  const body = document.createElement('div');
  body.className = 'sheet-card__body';

  const grid = document.createElement('div');
  grid.className = 'mapping-grid mapping-grid--sheet';

  const fields = [
    { key: 'name', label: 'Наименование', required: true },
    { key: 'wholesale_price', label: 'Опт', required: false },
    { key: 'recommended_price', label: 'РРЦ', required: false },
    { key: 'stock', label: 'Остаток', required: false },
  ];

  const selects = {};
  fields.forEach((field) => {
    const wrapper = document.createElement('label');
    const caption = document.createElement('span');
    caption.textContent = field.label + (field.required ? '' : ' (необязательно)');
    const savedValue = existingMapping && existingMapping[field.key] ? existingMapping[field.key] : '';
    const suggestedValue = savedValue || suggestColumn(sheet, field.key);
    const select = createMappingSelect(sheet.columns || [], field.key, suggestedValue);
    wrapper.append(caption, select);
    grid.appendChild(wrapper);
    selects[field.key] = select;
  });

  body.appendChild(grid);
  container.appendChild(body);

  checkbox.addEventListener('change', () => {
    body.hidden = !checkbox.checked;
    if (checkbox.checked) {
      container.classList.remove('sheet-card--disabled');
    } else {
      container.classList.add('sheet-card--disabled');
    }
    updateSaveButtonState();
  });

  const shouldCheck = Boolean(existingMapping) || index === 0;
  checkbox.checked = shouldCheck;
  body.hidden = !shouldCheck;
  if (!shouldCheck) {
    container.classList.add('sheet-card--disabled');
  }

  sheetControls.set(sheet.name, {
    checkbox,
    selects,
    container,
    sheet,
  });

  if (sheetList) {
    sheetList.appendChild(container);
  }
}

function applySheetMappings(previewData, handlerSheets = []) {
  if (!sheetList) {
    return;
  }
  sheetList.innerHTML = '';
  sheetControls.clear();

  const mappingByName = new Map();
  handlerSheets.forEach((item) => {
    if (item && item.sheet_name) {
      mappingByName.set(item.sheet_name, item.mapping || {});
    }
  });

  previewData.forEach((sheet, index) => {
    const mapping = mappingByName.get(sheet.name) || null;
    buildSheetCard(sheet, index, mapping);
  });
}

function collectSelectedSheets() {
  const configs = [];
  let valid = true;
  sheetControls.forEach((control, sheetName) => {
    if (!control.checkbox.checked) {
      return;
    }
    const mapping = {};
    const nameValue = (control.selects.name.value || '').trim().toUpperCase();
    if (!nameValue) {
      valid = false;
    }
    mapping.name = nameValue;
    ['wholesale_price', 'recommended_price', 'stock'].forEach((field) => {
      const value = (control.selects[field].value || '').trim().toUpperCase();
      mapping[field] = value;
    });
    configs.push({
      sheet_name: sheetName,
      position: configs.length,
      mapping,
    });
  });
  return { configs, valid };
}

function getSelectedFileDescriptor() {
  if (!sourceFileInput || !sourceFileInput.files || sourceFileInput.files.length === 0) {
    return null;
  }
  const file = sourceFileInput.files[0];
  return {
    name: file.name || '',
    size: Number.isFinite(file.size) ? file.size : 0,
    lastModified: Number.isFinite(file.lastModified) ? file.lastModified : 0
  };
}

function isPreviewCurrent() {
  if (!previewContext) {
    return false;
  }
  const mode = handlerDialog.dataset.mode || 'create';
  if (previewContext.type === 'url') {
    return handlerForm.elements.source_url.value.trim() === previewContext.value;
  }
  if (previewContext.type === 'file') {
    const descriptor = getSelectedFileDescriptor();
    if (!descriptor) {
      return false;
    }
    return (
      descriptor.name === previewContext.name &&
      descriptor.size === previewContext.size &&
      descriptor.lastModified === previewContext.lastModified
    );
  }
  if (previewContext.type === 'handler') {
    return mode === 'edit' && String(previewContext.id) === String(editingHandlerId || '');
  }
  return false;
}

function updateSaveButtonState() {
  const titleFilled = handlerForm.elements.title.value.trim() !== '';
  const mode = handlerDialog.dataset.mode || 'create';
  const sourceUrl = handlerForm.elements.source_url.value.trim();
  const fileDescriptor = getSelectedFileDescriptor();
  const hasExistingFile = handlerDialog.dataset.hasFile === 'true';
  const hasSource = Boolean(sourceUrl) || Boolean(fileDescriptor) || (mode === 'edit' && hasExistingFile);
  const selection = collectSelectedSheets();
  const hasSheets = selection.configs.length > 0;
  const mappingsSelected = selection.valid;
  const previewValid = previewSheets.length > 0 && isPreviewCurrent();
  const wholesaleValue = Number.parseFloat(wholesaleMultiplierInput.value);
  const recommendedValue = Number.parseFloat(recommendedMultiplierInput.value);
  const multipliersValid =
    Number.isFinite(wholesaleValue) && wholesaleValue > 0 &&
    Number.isFinite(recommendedValue) && recommendedValue > 0;
  const isReady = titleFilled && hasSource && hasSheets && mappingsSelected && previewValid && multipliersValid;
  saveHandlerButton.disabled = !isReady;
}

async function requestPreview({ sourceUrl = '', file = null, handlerId = null } = {}) {
  const formData = new FormData();
  if (sourceUrl) {
    formData.append('source_url', sourceUrl);
  }
  if (file) {
    formData.append('source_file', file);
  }
  if (handlerId) {
    formData.append('handler_id', handlerId);
  }
  const response = await fetch('/api/suppliers/preview-source', {
    method: 'POST',
    body: formData
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || 'Не удалось загрузить таблицу');
  }
  return payload;
}

function applyPreviewResult(payload, options = {}) {
  previewSheets = Array.isArray(payload.sheets) ? payload.sheets : [];
  if (!previewSheets.length) {
    throw new Error('Не удалось обнаружить листы в файле');
  }
  mappingSection.hidden = false;
  applySheetMappings(previewSheets, options.handlerSheets || []);

  if (options.contextType === 'file') {
    const descriptor = getSelectedFileDescriptor();
    if (descriptor) {
      previewContext = {
        type: 'file',
        name: descriptor.name,
        size: descriptor.size,
        lastModified: descriptor.lastModified
      };
    } else {
      previewContext = null;
    }
  } else if (options.contextType === 'url') {
    previewContext = { type: 'url', value: options.sourceUrl || '' };
  } else if (options.contextType === 'handler') {
    previewContext = { type: 'handler', id: String(options.handlerId || '') };
  }

  updateSaveButtonState();
}

function formatDateTime(value) {
  if (!value) {
    return 'Ожидает обновления';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString('ru-RU');
}

function createEmptyRow(message = 'Данные прайс-листа не найдены.') {
  const tr = document.createElement('tr');
  tr.className = 'empty-row';
  const td = document.createElement('td');
  td.colSpan = 5;
  td.textContent = message;
  tr.appendChild(td);
  return tr;
}

function formatChangeDate(value) {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString('ru-RU', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function formatChangeValue(value) {
  if (value === null || value === undefined || value === '') {
    return '—';
  }
  return String(value);
}

function isNegativeStockValue(value) {
  const normalized = normalizeCellValue(value);
  if (!normalized) {
    return false;
  }
  const numeric = Number(normalized.replace(',', '.'));
  if (Number.isFinite(numeric)) {
    return numeric < 0;
  }
  return /^-+$/.test(normalized);
}

function matchesSearch(row, term) {
  const normalizedTerm = (term || '').trim().toLowerCase();
  if (!normalizedTerm) {
    return true;
  }
  const nameValue = normalizeCellValue(row.name).toLowerCase();
  return nameValue.includes(normalizedTerm);
}

function getFilteredRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return [];
  }

  const filtered = [];
  let pendingCategory = null;
  let categoryAdded = false;
  let lastSyntheticCategory = '';

  rows.forEach((row) => {
    const kind = row.kind || 'item';
    if (kind === 'category') {
      pendingCategory = row;
      categoryAdded = false;
      lastSyntheticCategory = '';
      return;
    }

    if (!hasMeaningfulContent(row.name) &&
        !hasMeaningfulContent(row.wholesale_price) &&
        !hasMeaningfulContent(row.recommended_price) &&
        !hasMeaningfulContent(row.stock)) {
      return;
    }

    if (filtersState.hideNegativeStock && isNegativeStockValue(row.stock)) {
      return;
    }

    if (!matchesSearch(row, currentSearchTerm)) {
      return;
    }

    if (pendingCategory && !categoryAdded) {
      filtered.push(pendingCategory);
      categoryAdded = true;
    } else if (!pendingCategory && hasMeaningfulContent(row.category)) {
      const categoryName = normalizeCellValue(row.category);
      if (categoryName && lastSyntheticCategory !== categoryName) {
        filtered.push({
          kind: 'category',
          category: categoryName,
          sheet: row.sheet
        });
        lastSyntheticCategory = categoryName;
      }
    }

    filtered.push(row);
  });

  return filtered;
}

function renderHandlerRows(card, rows) {
  if (!card._ui) {
    return;
  }
  const { tbody, sheetTabs, sheetTabsList } = card._ui;
  const sourceRows = Array.isArray(rows) ? rows : [];
  const grouped = new Map();
  sourceRows.forEach((row) => {
    const sheetName = row.sheet || '';
    if (!grouped.has(sheetName)) {
      grouped.set(sheetName, []);
    }
    grouped.get(sheetName).push(row);
  });

  const handlerSheets = (card._handler && Array.isArray(card._handler.sheets)) ? card._handler.sheets : [];
  handlerSheets.forEach((sheet) => {
    const name = sheet.sheet_name || '';
    if (!grouped.has(name)) {
      grouped.set(name, []);
    }
  });

  const orderedNames = [];
  handlerSheets.forEach((sheet) => {
    const name = sheet.sheet_name || '';
    if (grouped.has(name) && !orderedNames.includes(name)) {
      orderedNames.push(name);
    }
  });
  const remainingNames = Array.from(grouped.keys()).filter((name) => !orderedNames.includes(name));
  remainingNames.sort((a, b) => a.localeCompare(b || '', 'ru', { sensitivity: 'base' }));
  const sheetNames = orderedNames.concat(remainingNames);

  let activeSheet = card._activeSheet;
  if (!activeSheet || !grouped.has(activeSheet)) {
    activeSheet = sheetNames[0] || null;
  }
  card._activeSheet = activeSheet;

  if (sheetTabs && sheetTabsList) {
    sheetTabsList.innerHTML = '';
    if (!sheetNames.length || sheetNames.length === 1) {
      sheetTabs.hidden = true;
    } else {
      sheetTabs.hidden = false;
      sheetNames.forEach((name) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'sheet-tabs__button';
        button.textContent = name || 'Лист';
        if (name === activeSheet) {
          button.classList.add('is-active');
          button.setAttribute('aria-current', 'page');
        }
        button.addEventListener('click', () => {
          if (card._activeSheet === name) {
            return;
          }
          card._activeSheet = name;
          renderHandlerRows(card, rows);
        });
        sheetTabsList.appendChild(button);
      });
    }
  }

  const activeRows = activeSheet ? grouped.get(activeSheet) || [] : [];
  const filteredRows = getFilteredRows(activeRows);
  tbody.innerHTML = '';

  let hasVisibleRows = false;
  filteredRows.forEach((row) => {
    const kind = row.kind || 'item';
    if (kind === 'category') {
      tbody.appendChild(createCategoryRow(row.category));
      hasVisibleRows = true;
      return;
    }
    const itemRow = createItemRow(card._handler, row);
    if (itemRow) {
      tbody.appendChild(itemRow);
      hasVisibleRows = true;
    }
  });

  if (!hasVisibleRows) {
    const message = currentSearchTerm || filtersState.hideNegativeStock
      ? 'По запросу ничего не найдено.'
      : 'Данные прайс-листа не найдены.';
    tbody.appendChild(createEmptyRow(message));
  }
}

function renderHandlerChanges(card, changes) {
  // Changes are now shown in a dialog on demand.
  return;
}

function populateChangesDialog(changes) {
  if (!changesDialogList) return;

  changesDialogList.innerHTML = '';
  const items = Array.isArray(changes) ? changes : [];

  if (!items.length) {
    if (changesDialogEmpty) changesDialogEmpty.hidden = false;
    changesDialogList.hidden = true;
    return;
  }
  if (changesDialogEmpty) changesDialogEmpty.hidden = true;
  changesDialogList.hidden = false;

  items.forEach((change) => {
    const listItem = document.createElement('li');
    listItem.className = 'supplier-card__change-item';

    const date = document.createElement('span');
    date.className = 'supplier-card__change-date';
    const formattedDate = formatChangeDate(change.changed_at);
    date.textContent = formattedDate || '';

    const text = document.createElement('span');
    text.className = 'supplier-card__change-text';
    const fieldLabel = changeFieldLabels[change.field] || change.field || 'Показатель';
    const itemName = (change.item_name || '').trim();
    const category = (change.category || '').trim();
    const sheetName = (change.sheet_name || '').trim();
    const namePart = itemName ? ` «${itemName}»` : '';
    const categoryPart = category ? ` (${category})` : '';
    const sheetPart = sheetName ? `[${sheetName}] ` : '';
    text.textContent = `${sheetPart}${fieldLabel}${namePart}${categoryPart}: ${formatChangeValue(change.old_value)} → ${formatChangeValue(change.new_value)}`;

    listItem.append(date, text);
    changesDialogList.appendChild(listItem);
  });
}

function updateControlsAvailability() {
  const hasHandlers = Array.isArray(handlersState.list) && handlersState.list.length > 0;
  if (searchInput) {
    searchInput.disabled = !hasHandlers;
    if (!hasHandlers) {
      searchInput.value = '';
      currentSearchTerm = '';
    }
  }
  if (negativeFilterToggle) {
    negativeFilterToggle.disabled = !hasHandlers;
    if (!hasHandlers) {
      negativeFilterToggle.checked = false;
      filtersState.hideNegativeStock = false;
    }
  }
  if (refreshAllButton) {
    refreshAllButton.disabled = !hasHandlers;
  }
}

function renderTabs(handlers) {
  if (!tabsContainer) {
    return;
  }
  tabsContainer.innerHTML = '';
  if (!handlers.length) {
    tabsContainer.hidden = true;
    return;
  }

  tabsContainer.hidden = false;
  handlers.forEach((handler) => {
    const link = document.createElement('a');
    link.className = 'tabs__button';
    link.dataset.handlerId = String(handler.id);
    link.href = `/suppliers/${handler.id}`;
    link.textContent = handler.title || `Прайс-лист ${handler.id}`;
    const isActive = String(handler.id) === String(handlersState.activeId);
    if (isActive) {
      link.classList.add('is-active');
      link.setAttribute('aria-current', 'page');
    }
    tabsContainer.appendChild(link);
  });
}

function getActiveCard() {
  if (!handlersState.activeId) {
    return null;
  }
  return handlersContainer.querySelector(`[data-handler-id="${handlersState.activeId}"]`);
}

function renderActiveCardRows() {
  const card = getActiveCard();
  if (!card) {
    return;
  }
  renderHandlerRows(card, card._rows || []);
}

function buildSupplierInfo(row) {
  const parts = [];
  const category = (row.category || '').trim();
  if (category) {
    parts.push(category);
  }
  if (row.wholesale_price) {
    parts.push(`Опт: ${row.wholesale_price}`);
  }
  if (row.recommended_price) {
    parts.push(`РРЦ: ${row.recommended_price}`);
  }
  if (row.stock) {
    parts.push(`Остаток: ${row.stock}`);
  }
  return parts.join(' • ');
}

function openAddItemDialog(handler, row) {
  addItemContext = { handler, row };
  const name = (row.name || '').trim();
  addForm.elements.name.value = name;
  addForm.elements.category.value = row.category || '';
  addForm.elements.quantity.value = 1;
  addForm.elements.wholesale_price.value = toInputNumber(row.wholesale_price);
  const recommendedValue = toInputNumber(row.recommended_price);
  addForm.elements.sale_price.value = recommendedValue;
  addHeading.textContent = name || 'Добавить на склад';
  const infoParts = [];
  if (handler && handler.title) {
    infoParts.push(handler.title);
  }
  const infoText = buildSupplierInfo(row);
  if (infoText) {
    infoParts.push(infoText);
  }
  addInfo.textContent = infoParts.join(' • ') || 'Укажите параметры товара для добавления на склад.';
  addDialog.showModal();
  addForm.elements.quantity.focus();
}

function closeAddItemDialog() {
  addDialog.close('cancel');
  addItemContext = null;
  addForm.reset();
  addHeading.textContent = 'Добавить на склад';
  addInfo.textContent = '';
}

async function submitAddItem(event) {
  event.preventDefault();
  if (!addItemContext) {
    showToast('Выберите товар из прайс-листа', 'error');
    return;
  }

  const name = addForm.elements.name.value.trim();
  if (!name) {
    showToast('Наименование обязательно.', 'error');
    return;
  }

  const payload = {
    name,
    category: addForm.elements.category.value || null,
    wholesale_price: addForm.elements.wholesale_price.value,
    sale_price: addForm.elements.sale_price.value,
    quantity: addForm.elements.quantity.value,
    date_received: new Date().toISOString().slice(0, 10)
  };

  if (!payload.wholesale_price) {
    showToast('Заполните поле «Опт».', 'error');
    return;
  }

  if (!payload.sale_price) {
    delete payload.sale_price;
  }

  const originalText = addSubmitButton.textContent;
  addSubmitButton.disabled = true;
  addSubmitButton.textContent = 'Добавление...';

  try {
    const response = await fetch('/items', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error || 'Не удалось добавить товар');
    }
    showToast('Товар добавлен на склад', 'success');
    closeAddItemDialog();
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    addSubmitButton.disabled = false;
    addSubmitButton.textContent = originalText;
  }
}

function buildHandlerCard(handler) {
  const layout = document.createElement('div');
  layout.className = 'supplier-layout';
  layout.dataset.handlerId = handler.id;

  const card = document.createElement('section');
  card.className = 'card supplier-card';

  const header = document.createElement('header');
  header.className = 'supplier-card__header';

  const meta = document.createElement('div');
  meta.className = 'supplier-card__meta';

  const title = document.createElement('h2');
  title.textContent = handler.title;

  const link = document.createElement('a');
  link.href = handler.source_url || '#';
  link.target = '_blank';
  link.rel = 'noopener';
  link.className = 'supplier-card__link';
  link.textContent = 'Открыть источник';
  if (!handler.source_url) {
    link.hidden = true;
  }

  const fileInfo = document.createElement('p');
  fileInfo.className = 'supplier-card__file muted';
  fileInfo.hidden = true;

  const status = document.createElement('p');
  status.className = 'supplier-card__status';

  const errorMessage = document.createElement('p');
  errorMessage.className = 'supplier-card__error';
  errorMessage.hidden = true;

  meta.append(title, link, fileInfo, status, errorMessage);

  const actions = document.createElement('div');
  actions.className = 'supplier-card__actions';

  const refreshButton = document.createElement('button');
  refreshButton.type = 'button';
  refreshButton.className = 'btn btn-secondary';
  refreshButton.textContent = 'Обновить';
  actions.appendChild(refreshButton);

  const historyButton = document.createElement('button');
  historyButton.type = 'button';
  historyButton.className = 'btn btn-secondary';
  historyButton.textContent = 'История';
  actions.appendChild(historyButton);

  const editButton = document.createElement('button');
  editButton.type = 'button';
  editButton.className = 'btn btn-secondary';
  editButton.textContent = 'Изменить';
  actions.appendChild(editButton);

  header.append(meta, actions);

  const tableWrapper = document.createElement('div');
  tableWrapper.className = 'table-wrapper supplier-card__table';

  const table = document.createElement('table');
  table.className = 'supplier-table';
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th class="supplier-table__col-name">Наименование</th>
      <th class="supplier-table__col-number">Опт</th>
      <th class="supplier-table__col-number">РРЦ</th>
      <th class="supplier-table__col-number">Остаток</th>
      <th class="supplier-table__col-actions">Добавить</th>
    </tr>
  `;
  const tbody = document.createElement('tbody');

  table.append(thead, tbody);
  tableWrapper.appendChild(table);

  const body = document.createElement('div');
  body.className = 'supplier-card__body';
  const sheetTabs = document.createElement('div');
  sheetTabs.className = 'sheet-tabs';
  sheetTabs.hidden = true;
  const sheetTabsList = document.createElement('div');
  sheetTabsList.className = 'sheet-tabs__list';
  sheetTabs.appendChild(sheetTabsList);
  body.append(sheetTabs, tableWrapper);

  card.append(header, body);

  layout.append(card);
  layout._ui = {
    title,
    link,
    fileInfo,
    status,
    errorMessage,
    refreshButton,
    editButton,
    historyButton,
    sheetTabs,
    sheetTabsList,
    tbody
  };

  historyButton.addEventListener('click', () => {
    populateChangesDialog(card._changes || []);
    if (changesDialog) changesDialog.showModal();
  });

  refreshButton.addEventListener('click', () => {
    refreshHandler(handler.id, layout);
  });

  editButton.addEventListener('click', () => {
    if (layout._handler) {
      openEditDialog(layout._handler);
    }
  });

  return layout;
}

function updateHandlerCard(card, handler) {
  if (!card._ui) {
    return;
  }
  const { title, link, fileInfo, status, errorMessage } = card._ui;
  title.textContent = handler.title;
  if (handler.source_url) {
    link.href = handler.source_url;
    link.hidden = false;
  } else {
    link.hidden = true;
  }
  if (fileInfo) {
    if (handler.source_file_name) {
      fileInfo.textContent = `Файл: ${handler.source_file_name}`;
      fileInfo.hidden = false;
    } else if (!handler.source_url) {
      fileInfo.textContent = 'Источник не указан';
      fileInfo.hidden = false;
    } else {
      fileInfo.hidden = true;
    }
  }
  status.textContent = `Обновлено: ${formatDateTime(handler.last_refreshed_at)}`;
  if (handler.error) {
    errorMessage.textContent = handler.error;
    errorMessage.hidden = false;
  } else {
    errorMessage.hidden = true;
  }

  card._handler = handler;
  card.dataset.hasFile = handler.has_file ? 'true' : 'false';
  card._rows = handler.rows || [];
  card._changes = handler.changes || [];

  renderHandlerRows(card, card._rows);
  renderHandlerChanges(card, card._changes);
  card.hidden = false;
}

async function refreshHandler(handlerId, card) {
  const { refreshButton } = card._ui;
  const originalText = refreshButton.textContent;
  refreshButton.disabled = true;
  refreshButton.textContent = 'Обновление...';
  try {
    const response = await fetch(`/api/suppliers/${handlerId}/refresh`, {
      method: 'POST'
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось обновить прайс-лист');
    }
    updateHandlerCard(card, payload.handler);
    showToast('Прайс-лист обновлён', 'success');
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    refreshButton.disabled = false;
    refreshButton.textContent = originalText;
  }
}

async function refreshAllHandlers() {
  if (!refreshAllButton) {
    return;
  }
  const originalText = refreshAllButton.textContent;
  refreshAllButton.disabled = true;
  refreshAllButton.textContent = 'Обновление...';
  try {
    const response = await fetch('/api/suppliers/refresh-all', {
      method: 'POST'
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось обновить прайс-листы');
    }
    const updated = Number(payload.updated) || 0;
    const failed = Array.isArray(payload.failed) ? payload.failed : [];
    await loadHandlersSummary();
    if (failed.length && !updated) {
      showToast('Не удалось обновить прайс-листы', 'error');
    } else if (failed.length) {
      showToast(`Обновлено: ${updated}. Ошибок: ${failed.length}`, 'error');
    } else {
      showToast(`Обновлено прайс-листов: ${updated}`, 'success');
    }
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    refreshAllButton.textContent = originalText;
    updateControlsAvailability();
  }
}

function showEmptyState() {
  emptyStateCard.hidden = false;
  handlersContainer.hidden = true;
  handlersContainer.innerHTML = '';
  if (tabsContainer) {
    tabsContainer.hidden = true;
  }
  handlersState.activeId = null;
  currentSearchTerm = '';
  filtersState.hideNegativeStock = false;
  handlersState.list = [];
  updateControlsAvailability();
}

function showLoadingState() {
  if (searchInput) {
    searchInput.disabled = true;
  }
  if (negativeFilterToggle) {
    negativeFilterToggle.disabled = true;
  }
  if (refreshAllButton) {
    refreshAllButton.disabled = true;
  }
  emptyStateCard.hidden = true;
  handlersContainer.hidden = false;
  handlersContainer.innerHTML = '';
  const loadingCard = document.createElement('section');
  loadingCard.className = 'card supplier-card supplier-card--placeholder';
  const text = document.createElement('p');
  text.className = 'muted';
  text.textContent = 'Загрузка прайс-листа...';
  loadingCard.appendChild(text);
  handlersContainer.appendChild(loadingCard);
}

function showErrorState(message) {
  emptyStateCard.hidden = true;
  handlersContainer.hidden = false;
  handlersContainer.innerHTML = '';
  const errorCard = document.createElement('section');
  errorCard.className = 'card supplier-card supplier-card--placeholder';
  const text = document.createElement('p');
  text.className = 'muted';
  text.textContent = message || 'Не удалось загрузить прайс-лист.';
  errorCard.appendChild(text);
  handlersContainer.appendChild(errorCard);
  if (negativeFilterToggle) {
    negativeFilterToggle.disabled = true;
  }
  if (refreshAllButton) {
    refreshAllButton.disabled = handlersState.list.length === 0;
  }
}

async function loadActiveHandler(handlerId, options = {}) {
  const { silent = false } = options;
  if (handlerId === null || handlerId === undefined || handlerId === '') {
    return;
  }

  const numericId = Number.parseInt(handlerId, 10);
  if (!Number.isFinite(numericId)) {
    return;
  }

  handlersState.activeId = String(numericId);
  renderTabs(handlersState.list);

  if (!silent) {
    showLoadingState();
  }

  try {
    const response = await fetch(`/api/suppliers/${numericId}`);
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось загрузить прайс-лист');
    }
    const handler = payload.handler || {};
    if (!handler.id) {
      throw new Error('Не удалось загрузить прайс-лист');
    }
    let card = handlersContainer.querySelector(`[data-handler-id="${handler.id}"]`);
    if (!card) {
      handlersContainer.innerHTML = '';
      card = buildHandlerCard(handler);
      handlersContainer.appendChild(card);
    }
    updateHandlerCard(card, handler);
    handlersContainer.hidden = false;
    emptyStateCard.hidden = true;
    updateControlsAvailability();

    const summaryIndex = handlersState.list.findIndex(
      (item) => String(item.id) === String(handler.id)
    );
    if (summaryIndex >= 0) {
      handlersState.list[summaryIndex] = {
        ...handlersState.list[summaryIndex],
        title: handler.title,
        last_refreshed_at: handler.last_refreshed_at
      };
      renderTabs(handlersState.list);
    }
  } catch (error) {
    if (!silent) {
      showErrorState(error.message);
    }
    showToast(error.message, 'error');
  }
}

async function loadHandlersSummary() {
  try {
    const response = await fetch('/api/suppliers?summary=1');
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось загрузить прайс-листы');
    }

    const handlers = payload.handlers || [];
    handlersState.list = handlers;
    updateControlsAvailability();

    if (handlers.length === 0) {
      showEmptyState();
      return;
    }

    renderTabs(handlers);

    if (!handlers.some((handler) => String(handler.id) === String(handlersState.activeId))) {
      const firstHandler = handlers[0];
      if (firstHandler) {
        window.location.replace(`/suppliers/${firstHandler.id}`);
        return;
      }
    }

    await loadActiveHandler(handlersState.activeId || handlers[0].id);
  } catch (error) {
    showToast(error.message, 'error');
    if (!handlersState.list.length) {
      showErrorState(error.message);
    }
  }
}

async function previewSource() {
  const sourceUrl = handlerForm.elements.source_url.value.trim();
  const file = sourceFileInput && sourceFileInput.files ? sourceFileInput.files[0] : null;
  if (!sourceUrl && !file) {
    showToast('Укажите ссылку или загрузите файл', 'error');
    return;
  }

  previewButton.disabled = true;
  const originalText = previewButton.textContent;
  previewButton.textContent = 'Загрузка...';

  try {
    const payload = await requestPreview({
      sourceUrl,
      file,
      handlerId: editingHandlerId
    });
    applyPreviewResult(payload, {
      contextType: file ? 'file' : 'url',
      sourceUrl,
      handlerId: editingHandlerId
    });
    showToast('Таблица загружена', 'success');
  } catch (error) {
    resetPreview();
    showToast(error.message, 'error');
  } finally {
    previewButton.disabled = false;
    previewButton.textContent = originalText;
  }
}

async function submitHandler(event) {
  event.preventDefault();
  if (saveHandlerButton.disabled) {
    return;
  }

  const mode = handlerDialog.dataset.mode || 'create';
  const formData = new FormData();
  const title = handlerForm.elements.title.value.trim();
  const sourceUrl = handlerForm.elements.source_url.value.trim();
  const file = sourceFileInput && sourceFileInput.files ? sourceFileInput.files[0] : null;
  const currentEditingId = editingHandlerId;
  const sheetSelection = collectSelectedSheets();

  if (!sheetSelection.configs.length || !sheetSelection.valid) {
    showToast('Выберите листы и укажите колонку с наименованием.', 'error');
    return;
  }

  formData.append('title', title);
  if (sourceUrl) {
    formData.append('source_url', sourceUrl);
  }
  if (file) {
    formData.append('source_file', file);
  }
  formData.append('sheets', JSON.stringify(sheetSelection.configs));
  formData.append('wholesale_multiplier', wholesaleMultiplierInput.value || '1');
  formData.append('recommended_multiplier', recommendedMultiplierInput.value || '1');

  const originalText = saveHandlerButton.textContent;
  saveHandlerButton.disabled = true;
  saveHandlerButton.textContent = 'Сохранение...';

  try {
    const endpoint = mode === 'edit' && currentEditingId ? `/api/suppliers/${currentEditingId}` : '/api/suppliers';
    const response = await fetch(endpoint, {
      method: mode === 'edit' ? 'PUT' : 'POST',
      body: formData
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error || 'Не удалось сохранить прайс-лист');
    }

    const resultHandler = result.handler || null;
    showToast(mode === 'edit' ? 'Прайс-лист обновлён' : 'Прайс-лист добавлен', 'success');

    if (mode === 'create') {
      closeDialog();
      if (resultHandler && resultHandler.id) {
        window.location.assign(`/suppliers/${resultHandler.id}`);
        return;
      }
      await loadHandlersSummary();
      return;
    }

    const handlerId = resultHandler && resultHandler.id ? resultHandler.id : currentEditingId;
    closeDialog();
    if (resultHandler && resultHandler.id) {
      const handlerKey = String(resultHandler.id);
      handlersState.activeId = handlerKey;
      let card = handlersContainer.querySelector(`[data-handler-id="${handlerKey}"]`);
      if (!card) {
        handlersContainer.innerHTML = '';
        card = buildHandlerCard(resultHandler);
        handlersContainer.appendChild(card);
      }
      updateHandlerCard(card, resultHandler);
      const summaryIndex = handlersState.list.findIndex((item) => String(item.id) === handlerKey);
      if (summaryIndex >= 0) {
        handlersState.list[summaryIndex] = {
          ...handlersState.list[summaryIndex],
          title: resultHandler.title,
          last_refreshed_at: resultHandler.last_refreshed_at
        };
      }
      renderTabs(handlersState.list);
    } else if (handlerId) {
      await loadActiveHandler(handlerId, { silent: false });
    }
  } catch (error) {
    showToast(error.message, 'error');
  } finally {
    saveHandlerButton.textContent = originalText;
    saveHandlerButton.disabled = false;
  }
}

function openCreateDialog() {
  handlerForm.reset();
  if (sourceFileInput) {
    sourceFileInput.value = '';
  }
  editingHandlerId = null;
  handlerDialog.dataset.mode = 'create';
  handlerDialog.dataset.handlerId = '';
  handlerDialog.dataset.hasFile = 'false';
  handlerDialog.dataset.originalHasFile = 'false';
  handlerDialogTitle.textContent = 'Добавить прайс-лист';
  resetPreview();
  handlerDialog.showModal();
  updateSaveButtonState();
}

async function loadExistingHandlerPreview(handler) {
  if (!handler || !handler.id) {
    return;
  }
  const originalText = previewButton.textContent;
  previewButton.disabled = true;
  previewButton.textContent = 'Загрузка...';
  try {
    const payload = await requestPreview({ handlerId: handler.id });
    let handlerSheets = (payload.handler && payload.handler.sheets) || [];
    if (!Array.isArray(handlerSheets) || !handlerSheets.length) {
      handlerSheets = [
        {
          sheet_name: handler.sheet_name,
          mapping: {
            name: handler.column_name,
            wholesale_price: handler.column_wholesale,
            recommended_price: handler.column_recommended,
            stock: handler.column_stock
          }
        }
      ];
    }
    applyPreviewResult(payload, {
      contextType: 'handler',
      handlerId: handler.id,
      handlerSheets
    });
  } catch (error) {
    resetPreview();
    showToast(error.message, 'error');
  } finally {
    previewButton.disabled = false;
    previewButton.textContent = originalText;
  }
}

function openEditDialog(handler) {
  if (!handler) {
    return;
  }
  handlerForm.reset();
  if (sourceFileInput) {
    sourceFileInput.value = '';
  }
  editingHandlerId = handler.id;
  handlerDialog.dataset.mode = 'edit';
  handlerDialog.dataset.handlerId = String(handler.id);
  handlerDialog.dataset.hasFile = handler.has_file ? 'true' : 'false';
  handlerDialog.dataset.originalHasFile = handler.has_file ? 'true' : 'false';
  handlerDialogTitle.textContent = 'Редактировать прайс-лист';
  handlerForm.elements.title.value = handler.title || '';
  handlerForm.elements.source_url.value = handler.source_url || '';
  wholesaleMultiplierInput.value = toInputNumber(handler.wholesale_multiplier) || '1';
  recommendedMultiplierInput.value = toInputNumber(handler.recommended_multiplier) || '1';
  resetPreview();
  handlerDialog.showModal();
  loadExistingHandlerPreview(handler);
  updateSaveButtonState();
}

function closeDialog() {
  handlerDialog.close('cancel');
  handlerForm.reset();
  if (sourceFileInput) {
    sourceFileInput.value = '';
  }
  handlerDialog.dataset.mode = 'create';
  handlerDialog.dataset.handlerId = '';
  handlerDialog.dataset.hasFile = 'false';
  handlerDialog.dataset.originalHasFile = 'false';
  editingHandlerId = null;
  handlerDialogTitle.textContent = 'Добавить прайс-лист';
  resetPreview();
  updateSaveButtonState();
}

function handleUrlInput() {
  if (previewContext) {
    resetPreview();
  } else {
    updateSaveButtonState();
  }
}

function handleFileInputChange() {
  if (sourceFileInput && sourceFileInput.files && sourceFileInput.files.length > 0) {
    handlerDialog.dataset.hasFile = 'true';
  } else {
    handlerDialog.dataset.hasFile = handlerDialog.dataset.originalHasFile === 'true' ? 'true' : 'false';
  }
  resetPreview();
}

function initEventListeners() {
  addHandlerButton.addEventListener('click', openCreateDialog);
  cancelButton.addEventListener('click', closeDialog);
  previewButton.addEventListener('click', previewSource);
  handlerForm.addEventListener('submit', submitHandler);
  handlerForm.elements.title.addEventListener('input', updateSaveButtonState);
  handlerForm.elements.source_url.addEventListener('input', handleUrlInput);
  if (sourceFileInput) {
    sourceFileInput.addEventListener('change', handleFileInputChange);
  }
  wholesaleMultiplierInput.addEventListener('input', updateSaveButtonState);
  recommendedMultiplierInput.addEventListener('input', updateSaveButtonState);
  handlerDialog.addEventListener('cancel', (event) => {
    event.preventDefault();
    closeDialog();
  });
  addCancelButton.addEventListener('click', closeAddItemDialog);
  addForm.addEventListener('submit', submitAddItem);
  addDialog.addEventListener('cancel', (event) => {
    event.preventDefault();
    closeAddItemDialog();
  });
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      currentSearchTerm = (searchInput.value || '').trim().toLowerCase();
      renderActiveCardRows();
    });
    searchInput.addEventListener('search', () => {
      currentSearchTerm = (searchInput.value || '').trim().toLowerCase();
      renderActiveCardRows();
    });
  }
  if (negativeFilterToggle) {
    negativeFilterToggle.addEventListener('change', () => {
      filtersState.hideNegativeStock = negativeFilterToggle.checked;
      renderActiveCardRows();
    });
  }
  if (refreshAllButton) {
    refreshAllButton.addEventListener('click', refreshAllHandlers);
  }
}

function startAutoReload() {
  clearInterval(autoRefreshTimer);
  autoRefreshTimer = setInterval(() => {
    if (handlersState.activeId) {
      loadActiveHandler(handlersState.activeId, { silent: true });
    }
  }, AUTO_RELOAD_INTERVAL);
}

initEventListeners();
loadHandlersSummary().then(startAutoReload);
