/**
 * BILL TRACKER — Frontend logic
 * -----------------------------
 * Talks to the Google Apps Script Web App using the Fetch API.
 * POST bodies are sent as text/plain (containing JSON) so the
 * browser does NOT issue a CORS preflight request — Apps Script
 * web apps don't respond to OPTIONS requests.
 */

// ⚠️ REPLACE with your deployed Apps Script Web App URL (ends in /exec)
const API_URL = 'https://script.google.com/macros/s/AKfycbzCKHis1hywcj9vixvf98VlC1F9V2ufP1lBAMNggvRo9V4uiuzoB3au7gGDvxlnfHakBg/exec';

// ---------------------------------------------------------------------
// STATE
// ---------------------------------------------------------------------

const state = {
  todayStr: formatDateDDMMYYYY(new Date()),
  currentDate: formatDateDDMMYYYY(new Date()), // date currently being viewed
  selectedMethod: null,                        // 'Cash' | 'GPay' for the add-bill form
  isToday: true,
  isSubmitting: false,
  allDates: [],                                // list of sheet names for dropdown / nav
  editingBill: null                             // bill object currently open in edit modal
};

// ---------------------------------------------------------------------
// DOM REFS
// ---------------------------------------------------------------------

const el = {
  headerLabel: document.getElementById('headerLabel'),
  headerDate: document.getElementById('headerDate'),

  sumTotal: document.getElementById('sumTotal'),
  sumCash: document.getElementById('sumCash'),
  sumGpay: document.getElementById('sumGpay'),
  sumCount: document.getElementById('sumCount'),

  addBillCard: document.getElementById('addBillCard'),
  btnCash: document.getElementById('btnCash'),
  btnGpay: document.getElementById('btnGpay'),
  amountInput: document.getElementById('amountInput'),
  addBillBtn: document.getElementById('addBillBtn'),

  prevDayBtn: document.getElementById('prevDayBtn'),
  nextDayBtn: document.getElementById('nextDayBtn'),
  navDateLabel: document.getElementById('navDateLabel'),
  billsTitle: document.getElementById('billsTitle'),

  billsTableBody: document.getElementById('billsTableBody'),
  actionHeader: document.getElementById('actionHeader'),
  emptyState: document.getElementById('emptyState'),

  dateSelect: document.getElementById('dateSelect'),

  editModalOverlay: document.getElementById('editModalOverlay'),
  editBtnCash: document.getElementById('editBtnCash'),
  editBtnGpay: document.getElementById('editBtnGpay'),
  editAmountInput: document.getElementById('editAmountInput'),
  saveEditBtn: document.getElementById('saveEditBtn'),
  cancelEditBtn: document.getElementById('cancelEditBtn'),
  deleteBillBtn: document.getElementById('deleteBillBtn'),

  deleteModalOverlay: document.getElementById('deleteModalOverlay'),
  confirmDeleteBtn: document.getElementById('confirmDeleteBtn'),
  cancelDeleteBtn: document.getElementById('cancelDeleteBtn'),

  toast: document.getElementById('toast'),
  loadingOverlay: document.getElementById('loadingOverlay')
};

// ---------------------------------------------------------------------
// INIT
// ---------------------------------------------------------------------

window.addEventListener('DOMContentLoaded', init);

function init() {
  el.headerDate.textContent = formatDatePretty(state.todayStr);
  updateNavDateLabel();

  bindEvents();
  loadDates();
  loadBillsForCurrentDate();
}

function bindEvents() {
  el.btnCash.addEventListener('click', () => selectMethod('Cash'));
  el.btnGpay.addEventListener('click', () => selectMethod('GPay'));
  el.addBillBtn.addEventListener('click', handleAddBill);

  el.prevDayBtn.addEventListener('click', goToPreviousDay);
  el.nextDayBtn.addEventListener('click', goToNextDay);

  el.dateSelect.addEventListener('change', () => {
    if (el.dateSelect.value) {
      state.currentDate = el.dateSelect.value;
      loadBillsForCurrentDate();
    }
  });

  el.editBtnCash.addEventListener('click', () => selectEditMethod('Cash'));
  el.editBtnGpay.addEventListener('click', () => selectEditMethod('GPay'));
  el.saveEditBtn.addEventListener('click', handleSaveEdit);
  el.cancelEditBtn.addEventListener('click', closeEditModal);
  el.deleteBillBtn.addEventListener('click', openDeleteModal);

  el.confirmDeleteBtn.addEventListener('click', handleConfirmDelete);
  el.cancelDeleteBtn.addEventListener('click', closeDeleteModal);

  // Allow Enter key to submit amount
  el.amountInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') handleAddBill();
  });
}

// ---------------------------------------------------------------------
// ADD BILL (today only)
// ---------------------------------------------------------------------

function selectMethod(method) {
  state.selectedMethod = method;
  el.btnCash.classList.toggle('selected', method === 'Cash');
  el.btnGpay.classList.toggle('selected', method === 'GPay');
}

async function handleAddBill() {
  if (state.isSubmitting) return;

  if (!state.selectedMethod) {
    showToast('Please select Cash or GPay', 'error');
    return;
  }
  const amount = parseFloat(el.amountInput.value);
  if (isNaN(amount) || amount <= 0) {
    showToast('Please enter a valid amount', 'error');
    return;
  }

  state.isSubmitting = true;
  el.addBillBtn.disabled = true;
  el.addBillBtn.textContent = 'ADDING...';

  try {
    const result = await postToApi({
      action: 'addBill',
      paymentMethod: state.selectedMethod,
      amount: amount
    });

    if (result.success) {
      showToast('✓ Bill Added', 'success');
      el.amountInput.value = '';
      // Keep selected payment method ready for next entry (per spec)
      await loadBillsForCurrentDate(); // refresh totals/list without page reload
      await loadDates(); // in case a brand-new day's sheet was just created
    } else {
      showToast(result.message || 'Failed to add bill', 'error');
    }
  } catch (err) {
    showToast('Network error — please try again', 'error');
  } finally {
    state.isSubmitting = false;
    el.addBillBtn.disabled = false;
    el.addBillBtn.textContent = '+ ADD BILL';
  }
}

// ---------------------------------------------------------------------
// LOAD / RENDER BILLS
// ---------------------------------------------------------------------

async function loadBillsForCurrentDate() {
  state.isToday = state.currentDate === state.todayStr;
  updateNavDateLabel();
  toggleAddBillVisibility();

  showLoading(true);
  try {
    const data = await getFromApi({ action: 'getBills', date: state.currentDate });
    if (data.success) {
      renderSummary(data.summary);
      renderBillsTable(data.bills);
    } else {
      showToast(data.message || 'Failed to load bills', 'error');
    }
  } catch (err) {
    showToast('Network error — could not load bills', 'error');
  } finally {
    showLoading(false);
  }
}

function toggleAddBillVisibility() {
  el.addBillCard.classList.toggle('hidden', !state.isToday);
  el.billsTitle.textContent = state.isToday ? "Today's Bills" : 'Bills';
  el.actionHeader.classList.toggle('hidden', state.isToday); // edit/delete only for past days per spec
}

function renderSummary(summary) {
  el.sumTotal.textContent = formatRupees(summary.totalSales);
  el.sumCash.textContent = formatRupees(summary.cash);
  el.sumGpay.textContent = formatRupees(summary.gpay);
  el.sumCount.textContent = summary.totalBills;
}

function renderBillsTable(bills) {
  el.billsTableBody.innerHTML = '';

  if (!bills || bills.length === 0) {
    el.emptyState.classList.remove('hidden');
    return;
  }
  el.emptyState.classList.add('hidden');

  bills.forEach((bill) => {
    const tr = document.createElement('tr');

    tr.innerHTML = `
      <td>${bill.sno}</td>
      <td>${escapeHtml(bill.time)}</td>
      <td><span class="method-pill ${bill.paymentMethod}">${bill.paymentMethod}</span></td>
      <td class="amount-cell">${formatRupees(bill.amount)}</td>
      <td><button type="button" class="edit-link" data-sno="${bill.sno}">Edit</button></td>
    `;
    el.billsTableBody.appendChild(tr);

    if (!state.isToday) {
      tr.querySelector('.edit-link').addEventListener('click', () => openEditModal(bill));
    }
  });
}

// ---------------------------------------------------------------------
// DATE NAVIGATION + PREVIOUS DAYS DROPDOWN
// ---------------------------------------------------------------------

async function loadDates() {
  try {
    const data = await getFromApi({ action: 'getDates' });
    if (data.success) {
      state.allDates = data.dates; // newest first
      populateDateSelect();
    }
  } catch (err) {
    // Non-fatal — dropdown just won't populate
  }
}

function populateDateSelect() {
  const current = el.dateSelect.value;
  el.dateSelect.innerHTML = '<option value="">— Choose a date —</option>';
  state.allDates
    .filter(d => d !== state.todayStr)
    .forEach((d) => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = formatDatePretty(d);
      el.dateSelect.appendChild(opt);
    });
  if (current) el.dateSelect.value = current;
}

function updateNavDateLabel() {
  el.navDateLabel.textContent = state.currentDate;
  el.headerLabel.textContent = state.isToday === false ? 'Viewing Past Day' : "Today's Bills";
  // Disable "next day" beyond today (no future browsing)
  el.nextDayBtn.disabled = state.currentDate === state.todayStr;
}

function goToPreviousDay() {
  state.currentDate = shiftDate(state.currentDate, -1);
  el.dateSelect.value = state.currentDate === state.todayStr ? '' : state.currentDate;
  loadBillsForCurrentDate();
}

function goToNextDay() {
  if (state.currentDate === state.todayStr) return; // no future dates
  state.currentDate = shiftDate(state.currentDate, 1);
  el.dateSelect.value = state.currentDate === state.todayStr ? '' : state.currentDate;
  loadBillsForCurrentDate();
}

// ---------------------------------------------------------------------
// EDIT MODAL
// ---------------------------------------------------------------------

function openEditModal(bill) {
  state.editingBill = bill;
  selectEditMethod(bill.paymentMethod);
  el.editAmountInput.value = bill.amount;
  el.editModalOverlay.classList.remove('hidden');
}

function closeEditModal() {
  state.editingBill = null;
  el.editModalOverlay.classList.add('hidden');
}

function selectEditMethod(method) {
  state.editSelectedMethod = method;
  el.editBtnCash.classList.toggle('selected', method === 'Cash');
  el.editBtnGpay.classList.toggle('selected', method === 'GPay');
}

async function handleSaveEdit() {
  if (!state.editingBill) return;

  const amount = parseFloat(el.editAmountInput.value);
  if (isNaN(amount) || amount <= 0) {
    showToast('Please enter a valid amount', 'error');
    return;
  }
  if (!state.editSelectedMethod) {
    showToast('Please select a payment method', 'error');
    return;
  }

  showLoading(true);
  try {
    const result = await postToApi({
      action: 'updateBill',
      date: state.currentDate,
      sno: state.editingBill.sno,
      paymentMethod: state.editSelectedMethod,
      amount: amount
    });

    if (result.success) {
      showToast('✓ Bill Updated', 'success');
      closeEditModal();
      await loadBillsForCurrentDate();
    } else {
      showToast(result.message || 'Failed to update bill', 'error');
    }
  } catch (err) {
    showToast('Network error — please try again', 'error');
  } finally {
    showLoading(false);
  }
}

// ---------------------------------------------------------------------
// DELETE MODAL
// ---------------------------------------------------------------------

function openDeleteModal() {
  el.deleteModalOverlay.classList.remove('hidden');
}
function closeDeleteModal() {
  el.deleteModalOverlay.classList.add('hidden');
}

async function handleConfirmDelete() {
  if (!state.editingBill) return;

  showLoading(true);
  try {
    const result = await postToApi({
      action: 'deleteBill',
      date: state.currentDate,
      sno: state.editingBill.sno
    });

    if (result.success) {
      showToast('✓ Bill Deleted', 'success');
      closeDeleteModal();
      closeEditModal();
      await loadBillsForCurrentDate();
    } else {
      showToast(result.message || 'Failed to delete bill', 'error');
    }
  } catch (err) {
    showToast('Network error — please try again', 'error');
  } finally {
    showLoading(false);
  }
}

// ---------------------------------------------------------------------
// API HELPERS
// ---------------------------------------------------------------------

async function getFromApi(params) {
  const url = new URL(API_URL);
  Object.keys(params).forEach(key => url.searchParams.append(key, params[key]));

  const response = await fetch(url.toString(), { method: 'GET' });
  if (!response.ok) throw new Error('Network response was not ok');
  return response.json();
}

async function postToApi(payload) {
  // text/plain avoids a CORS preflight (OPTIONS) request, which
  // Apps Script web apps do not handle.
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw new Error('Network response was not ok');
  return response.json();
}

// ---------------------------------------------------------------------
// UI HELPERS
// ---------------------------------------------------------------------

let toastTimer = null;
function showToast(message, type) {
  el.toast.textContent = message;
  el.toast.className = `toast show ${type || ''}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.toast.classList.remove('show');
  }, 2200);
}

function showLoading(isLoading) {
  el.loadingOverlay.classList.toggle('hidden', !isLoading);
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ---------------------------------------------------------------------
// DATE / CURRENCY FORMATTING
// ---------------------------------------------------------------------

function formatDateDDMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function formatDatePretty(ddmmyyyy) {
  const [dd, mm, yyyy] = ddmmyyyy.split('-').map(Number);
  const date = new Date(yyyy, mm - 1, dd);
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${String(dd).padStart(2, '0')} ${months[mm - 1]} ${yyyy}`;
}

function shiftDate(ddmmyyyy, deltaDays) {
  const [dd, mm, yyyy] = ddmmyyyy.split('-').map(Number);
  const date = new Date(yyyy, mm - 1, dd);
  date.setDate(date.getDate() + deltaDays);
  return formatDateDDMMYYYY(date);
}

// Indian number formatting, e.g. ₹1,25,000
function formatRupees(amount) {
  const num = Number(amount) || 0;
  const formatted = num.toLocaleString('en-IN', {
    maximumFractionDigits: 2,
    minimumFractionDigits: num % 1 === 0 ? 0 : 2
  });
  return `₹${formatted}`;
}
