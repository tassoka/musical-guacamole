const tableBody = document.querySelector('#inventory-table tbody');
const newItemForm = document.getElementById('new-item-form');
const importForm = document.getElementById('import-form');
const inventorySearchInput = document.getElementById('inventory-search');
const inventoryCategoryFilter = document.getElementById('inventory-category-filter');
const historyModal = document.getElementById('history-modal');
const historyTableBody = historyModal.querySelector('tbody');
const modalCloseButton = historyModal.querySelector('.modal-close');
const toast = document.getElementById('toast');
const editDialog = document.getElementById('edit-item-dialog');
const editForm = document.getElementById('edit-item-form');
const editCancelButton = editForm ? editForm.querySelector('button[value="cancel"]') : null;
let editingRow = null;
let toastTimer;

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

function formatPrice(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  const numeric = Number(value);
  if (Number.isFinite(numeric)) {
    return numeric.toFixed(2);
  }
  return '';
}

function displayValue(value) {
  if (value === null || value === undefined) {
    return '—';
  }
  const normalized = String(value).trim();
  return normalized ? normalized : '—';
}

function refreshCategoryOptions() {
  if (!inventoryCategoryFilter) {
    return;
  }
  const previousValue = inventoryCategoryFilter.value;
  const categories = new Set();
  if (tableBody) {
    tableBody.querySelectorAll('tr[data-item-id]').forEach((row) => {
      const category = (row.dataset.category || '').trim();
      if (category) {
        categories.add(category);
      }
    });
  }
  const sorted = Array.from(categories).sort((a, b) => a.localeCompare(b, 'ru', { sensitivity: 'base' }));
  inventoryCategoryFilter.innerHTML = '';
  const allOption = document.createElement('option');
  allOption.value = '';
  allOption.textContent = 'Все категории';
  inventoryCategoryFilter.appendChild(allOption);
  sorted.forEach((category) => {
    const option = document.createElement('option');
    option.value = category;
    option.textContent = category;
    inventoryCategoryFilter.appendChild(option);
  });
  if (sorted.includes(previousValue)) {
    inventoryCategoryFilter.value = previousValue;
  }
}

function applyInventoryFilters() {
  if (!tableBody) {
    return;
  }
  const searchTerm = (inventorySearchInput ? inventorySearchInput.value : '').trim().toLowerCase();
  const selectedCategory = (inventoryCategoryFilter ? inventoryCategoryFilter.value : '').trim().toLowerCase();
  const dataRows = Array.from(tableBody.querySelectorAll('tr[data-item-id]'));
  let visibleCount = 0;

  dataRows.forEach((row) => {
    const name = (row.dataset.name || '').toLowerCase();
    const category = (row.dataset.category || '').toLowerCase();
    const matchesSearch = !searchTerm || name.includes(searchTerm);
    const matchesCategory = !selectedCategory || category === selectedCategory;
    const shouldShow = matchesSearch && matchesCategory;
    row.hidden = !shouldShow;
    if (shouldShow) {
      visibleCount += 1;
    }
  });

  const emptyRow = tableBody.querySelector('.empty-row');
  if (emptyRow) {
    emptyRow.hidden = dataRows.length > 0;
  }

  let filterRow = tableBody.querySelector('.filter-empty-row');
  if (!dataRows.length) {
    if (filterRow) {
      filterRow.remove();
    }
    return;
  }

  if (!visibleCount) {
    if (!filterRow) {
      filterRow = document.createElement('tr');
      filterRow.className = 'filter-empty-row';
      const cell = document.createElement('td');
      cell.colSpan = 8;
      cell.textContent = 'По запросу ничего не найдено.';
      filterRow.appendChild(cell);
      tableBody.appendChild(filterRow);
    }
    filterRow.hidden = false;
  } else if (filterRow) {
    filterRow.remove();
  }
}

function setFieldValue(row, field, value) {
  const span = row.querySelector(`.inventory-table__value[data-field="${field}"]`);
  if (!span) {
    return;
  }
  const shown = displayValue(value);
  span.textContent = shown;
  if (shown && shown !== '—') {
    span.title = shown;
  } else {
    span.removeAttribute('title');
  }
}

function applyItemData(row, item) {
  const previousCategory = row.dataset.category || '';
  row.dataset.itemId = item.id;
  const category = item.category ?? '';
  const name = item.name ?? '';
  const wholesale = formatPrice(item.wholesale_price ?? 0) || '0.00';
  const sale = item.sale_price === null || item.sale_price === undefined ? '' : formatPrice(item.sale_price);
  const quantity = Number.isFinite(Number(item.quantity)) ? String(item.quantity) : '0';
  const gtin = item.gtin ?? '';
  const dateReceived = item.date_received ?? '';

  row.dataset.category = category;
  row.dataset.name = name;
  row.dataset.wholesalePrice = wholesale;
  row.dataset.salePrice = sale || '';
  row.dataset.quantity = quantity;
  row.dataset.gtin = gtin;
  row.dataset.dateReceived = dateReceived;

  setFieldValue(row, 'category', category);
  setFieldValue(row, 'name', name);
  setFieldValue(row, 'wholesale_price', wholesale);
  setFieldValue(row, 'sale_price', sale);
  setFieldValue(row, 'gtin', gtin);
  setFieldValue(row, 'date_received', dateReceived);

  const quantityInput = row.querySelector('.quantity-input');
  if (quantityInput) {
    quantityInput.value = quantity;
    quantityInput.title = quantity;
  }

  if (row.isConnected) {
    if (previousCategory !== category) {
      refreshCategoryOptions();
    }
    applyInventoryFilters();
  }
}

function createValueCell(field, value, extraClass = '') {
  const td = document.createElement('td');
  const span = document.createElement('span');
  span.className = 'inventory-table__value';
  if (extraClass) {
    span.classList.add(extraClass);
  }
  span.dataset.field = field;
  span.textContent = displayValue(value);
  if (value) {
    span.title = String(value);
  }
  td.appendChild(span);
  return td;
}

function createRow(item) {
  const row = document.createElement('tr');
  row.dataset.itemId = item.id;

  row.appendChild(createValueCell('category', item.category ?? ''));
  row.appendChild(createValueCell('name', item.name ?? '', 'inventory-table__value--name'));
  row.appendChild(createValueCell('wholesale_price', formatPrice(item.wholesale_price ?? 0) || '0.00', 'inventory-table__value--number'));
  const saleValue = item.sale_price === null || item.sale_price === undefined ? '' : formatPrice(item.sale_price);
  row.appendChild(createValueCell('sale_price', saleValue, 'inventory-table__value--number'));

  const quantityCell = document.createElement('td');
  quantityCell.classList.add('quantity-cell');
  const controls = document.createElement('div');
  controls.className = 'quantity-controls';

  const minusButton = document.createElement('button');
  minusButton.type = 'button';
  minusButton.className = 'adjust-button';
  minusButton.dataset.delta = '-1';
  minusButton.textContent = '−';
  minusButton.setAttribute('aria-label', 'Уменьшить остаток');

  const quantityInput = document.createElement('input');
  quantityInput.type = 'number';
  quantityInput.className = 'quantity-input';
  quantityInput.step = '1';
  quantityInput.value = Number.isFinite(Number(item.quantity)) ? item.quantity : 0;
  quantityInput.setAttribute('aria-label', 'Остаток');

  const plusButton = document.createElement('button');
  plusButton.type = 'button';
  plusButton.className = 'adjust-button';
  plusButton.dataset.delta = '1';
  plusButton.textContent = '+';
  plusButton.setAttribute('aria-label', 'Увеличить остаток');

  controls.append(minusButton, quantityInput, plusButton);
  quantityCell.appendChild(controls);
  row.appendChild(quantityCell);

  row.appendChild(createValueCell('gtin', item.gtin ?? ''));
  row.appendChild(createValueCell('date_received', item.date_received ?? ''));

  const actionsCell = document.createElement('td');
  actionsCell.className = 'row-actions';

  const editButton = document.createElement('button');
  editButton.type = 'button';
  editButton.className = 'icon-button row-menu-button';
  editButton.setAttribute('aria-label', 'Редактировать товар');
  editButton.textContent = '⋯';

  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'icon-button row-delete-button';
  deleteButton.setAttribute('aria-label', 'Удалить товар');
  deleteButton.textContent = '🗑';

  const historyButton = document.createElement('button');
  historyButton.type = 'button';
  historyButton.className = 'button small secondary history-button';
  historyButton.textContent = 'История';

  actionsCell.append(editButton, deleteButton, historyButton);
  row.appendChild(actionsCell);

  applyItemData(row, item);
  attachRowHandlers(row);
  return row;
}

function setQuantityControlsDisabled(row, disabled) {
  row.querySelectorAll('.adjust-button, .quantity-input').forEach((element) => {
    element.disabled = disabled;
  });
}

function attachRowHandlers(row) {
  row.querySelectorAll('.adjust-button').forEach((button) => {
    if (!button.dataset.bound) {
      button.dataset.bound = 'true';
      button.addEventListener('click', async () => {
        const delta = Number(button.dataset.delta);
        if (!Number.isFinite(delta) || delta === 0) {
          return;
        }
        await adjustQuantity(row, delta);
      });
    }
  });

  const quantityInput = row.querySelector('.quantity-input');
  if (quantityInput && !quantityInput.dataset.bound) {
    quantityInput.dataset.bound = 'true';
    quantityInput.addEventListener('change', async () => {
      const newValue = Number.parseInt(quantityInput.value, 10);
      const current = Number.parseInt(row.dataset.quantity || '0', 10);
      if (!Number.isFinite(newValue)) {
        quantityInput.value = current;
        return;
      }
      const delta = newValue - current;
      if (delta === 0) {
        return;
      }
      await adjustQuantity(row, delta, { silentToast: true });
      showToast('Остаток обновлён', 'success');
    });
  }

  const editButton = row.querySelector('.row-menu-button');
  if (editButton && !editButton.dataset.bound) {
    editButton.dataset.bound = 'true';
    editButton.addEventListener('click', () => {
      openEditDialog(row);
    });
  }

  const historyButton = row.querySelector('.history-button');
  if (historyButton && !historyButton.dataset.bound) {
    historyButton.dataset.bound = 'true';
    historyButton.addEventListener('click', async () => {
      await openHistoryModal(row.dataset.itemId, row.dataset.name || 'Товар');
    });
  }

  const deleteButton = row.querySelector('.row-delete-button');
  if (deleteButton && !deleteButton.dataset.bound) {
    deleteButton.dataset.bound = 'true';
    deleteButton.addEventListener('click', async () => {
      const itemName = row.dataset.name || 'товар';
      const confirmed = window.confirm(`Удалить «${itemName}»?`);
      if (!confirmed) {
        return;
      }
      try {
        const response = await fetch(`/items/${row.dataset.itemId}`, {
          method: 'DELETE'
        });
        let payload = {};
        try {
          payload = await response.json();
        } catch (error) {
          // ignore JSON parse errors for empty responses
        }
        if (!response.ok) {
          throw new Error(payload.error || 'Не удалось удалить товар');
        }
        row.remove();
        refreshCategoryOptions();
        applyInventoryFilters();
        const emptyRow = tableBody ? tableBody.querySelector('.empty-row') : null;
        if (tableBody && !tableBody.querySelector('tr[data-item-id]')) {
          if (emptyRow) {
            emptyRow.hidden = false;
          } else {
            const placeholder = document.createElement('tr');
            placeholder.className = 'empty-row';
            const cell = document.createElement('td');
            cell.colSpan = 8;
            cell.textContent = 'Добавьте товары с помощью формы выше.';
            placeholder.appendChild(cell);
            tableBody.appendChild(placeholder);
          }
        }
        showToast('Товар удалён', 'success');
      } catch (error) {
        showToast(error.message, 'error');
      }
    });
  }
}

async function adjustQuantity(row, delta, options = {}) {
  if (!delta) {
    return;
  }
  const { silentToast = false } = options;
  const id = row.dataset.itemId;
  setQuantityControlsDisabled(row, true);
  try {
    const response = await fetch(`/items/${id}/adjust`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ delta }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось обновить количество');
    }
    row.dataset.quantity = String(payload.quantity);
    const quantityInput = row.querySelector('.quantity-input');
    if (quantityInput) {
      quantityInput.value = payload.quantity;
      quantityInput.title = String(payload.quantity);
    }
    if (!silentToast) {
      showToast(delta > 0 ? 'Количество увеличено' : 'Количество уменьшено', 'success');
    }
  } catch (error) {
    showToast(error.message, 'error');
    const quantityInput = row.querySelector('.quantity-input');
    if (quantityInput) {
      quantityInput.value = row.dataset.quantity || '0';
    }
  } finally {
    setQuantityControlsDisabled(row, false);
  }
}

async function openHistoryModal(itemId, itemName) {
  try {
    const response = await fetch(`/items/${itemId}/history`);
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || 'Не удалось получить историю');
    }
    historyTableBody.innerHTML = '';
    if (!payload.history.length) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 3;
      cell.textContent = 'Пока нет записей об изменениях.';
      row.appendChild(cell);
      historyTableBody.appendChild(row);
    } else {
      payload.history.forEach((entry) => {
        const row = document.createElement('tr');
        const createdAtCell = document.createElement('td');
        const changeCell = document.createElement('td');
        const actionCell = document.createElement('td');

        createdAtCell.textContent = new Date(entry.created_at).toLocaleString();
        changeCell.textContent = `${entry.change_amount > 0 ? '+' : ''}${entry.change_amount}`;
        actionCell.textContent = entry.action;

        row.append(createdAtCell, changeCell, actionCell);
        historyTableBody.appendChild(row);
      });
    }
    historyModal.querySelector('h3').textContent = `История движений — ${itemName}`;
    historyModal.hidden = false;
  } catch (error) {
    showToast(error.message, 'error');
  }
}

function closeHistoryModal() {
  historyModal.hidden = true;
}

modalCloseButton.addEventListener('click', closeHistoryModal);

historyModal.addEventListener('click', (event) => {
  if (event.target === historyModal) {
    closeHistoryModal();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && !historyModal.hidden) {
    closeHistoryModal();
  }
});

if (newItemForm) {
  const dateInput = newItemForm.querySelector('input[name="date_received"]');
  if (dateInput && !dateInput.value) {
    dateInput.value = new Date().toISOString().slice(0, 10);
  }

  newItemForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(newItemForm);
    const payload = Object.fromEntries(formData.entries());
    if (!payload.name || !payload.wholesale_price) {
      showToast('Заполните обязательные поля', 'error');
      return;
    }
    if (!payload.date_received) {
      payload.date_received = new Date().toISOString().slice(0, 10);
    }
    try {
      const response = await fetch('/items', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || 'Не удалось добавить товар');
      }
      if (data.item) {
        const newRow = createRow(data.item);
        const emptyRow = tableBody.querySelector('.empty-row');
        if (emptyRow) {
          emptyRow.remove();
        }
        tableBody.appendChild(newRow);
        refreshCategoryOptions();
        applyInventoryFilters();
        showToast('Товар добавлен', 'success');
      }
      newItemForm.reset();
      if (dateInput) {
        dateInput.value = new Date().toISOString().slice(0, 10);
      }
    } catch (error) {
      showToast(error.message, 'error');
    }
  });
}

if (importForm) {
  importForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const fileInput = importForm.querySelector('input[type="file"]');
    if (!fileInput.files.length) {
      showToast('Выберите файл для импорта', 'error');
      return;
    }
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
      const response = await fetch('/import/items', {
        method: 'POST',
        body: formData,
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Ошибка импорта');
      }
      showToast(`Импортировано записей: ${payload.imported}`, 'success');
      setTimeout(() => {
        window.location.reload();
      }, 1000);
    } catch (error) {
      showToast(error.message, 'error');
    }
  });
}

function closeEditDialog() {
  if (!editDialog) {
    return;
  }
  if (editDialog.open) {
    editDialog.close('cancel');
  }
  editingRow = null;
  if (editForm) {
    editForm.reset();
  }
}

function openEditDialog(row) {
  if (!editDialog || !editForm) {
    return;
  }
  editingRow = row;
  editForm.elements.id.value = row.dataset.itemId;
  editForm.elements.category.value = row.dataset.category || '';
  editForm.elements.name.value = row.dataset.name || '';
  editForm.elements.wholesale_price.value = row.dataset.wholesalePrice || '0.00';
  editForm.elements.sale_price.value = row.dataset.salePrice || '';
  editForm.elements.quantity.value = row.dataset.quantity || '0';
  editForm.elements.gtin.value = row.dataset.gtin || '';
  editForm.elements.date_received.value = row.dataset.dateReceived || '';
  editDialog.showModal();
  editForm.elements.name.focus();
}

if (editForm) {
  editForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (!editingRow) {
      closeEditDialog();
      return;
    }
    const id = editForm.elements.id.value;
    const payload = {
      category: editForm.elements.category.value || null,
      name: editForm.elements.name.value.trim(),
      wholesale_price: editForm.elements.wholesale_price.value,
      sale_price: editForm.elements.sale_price.value || null,
      quantity: editForm.elements.quantity.value || '0',
      gtin: editForm.elements.gtin.value || null,
      date_received: editForm.elements.date_received.value || editingRow.dataset.dateReceived || new Date().toISOString().slice(0, 10),
    };

    if (!payload.name) {
      showToast('Наименование обязательно.', 'error');
      return;
    }

    const submitButton = editForm.querySelector('button[type="submit"]');
    const originalText = submitButton.textContent;
    submitButton.disabled = true;
    submitButton.textContent = 'Сохранение...';

    try {
      const response = await fetch(`/items/${id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || 'Не удалось сохранить изменения');
      }
      if (data.item) {
        applyItemData(editingRow, data.item);
        showToast('Товар обновлён', 'success');
      }
      closeEditDialog();
    } catch (error) {
      showToast(error.message, 'error');
    } finally {
      submitButton.disabled = false;
      submitButton.textContent = originalText;
    }
  });
}

if (editCancelButton) {
  editCancelButton.addEventListener('click', () => {
    closeEditDialog();
  });
}

if (editDialog) {
  editDialog.addEventListener('cancel', () => {
    closeEditDialog();
  });
}

if (inventorySearchInput) {
  inventorySearchInput.addEventListener('input', applyInventoryFilters);
  inventorySearchInput.addEventListener('search', applyInventoryFilters);
}

if (inventoryCategoryFilter) {
  inventoryCategoryFilter.addEventListener('change', applyInventoryFilters);
}

Array.from(tableBody.querySelectorAll('tr[data-item-id]')).forEach((row) => {
  attachRowHandlers(row);
});

refreshCategoryOptions();
applyInventoryFilters();
