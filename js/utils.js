// ==================== Date Utilities ====================

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function toISODate(ddmmyyyy) {
  if (!ddmmyyyy) return '';
  if (ddmmyyyy.includes('-')) return ddmmyyyy;
  const [day, month, year] = ddmmyyyy.split('/');
  return `${year}-${month}-${day}`;
}

function todayISO() {
  return new Date().toISOString().split('T')[0];
}

function todayDisplay() {
  return formatDate(todayISO());
}

function getCurrentMonth() {
  return new Date().getMonth() + 1;
}

function getCurrentYear() {
  return new Date().getFullYear();
}

function getMonthName(month) {
  const names = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return names[month - 1] || '';
}

function getMonthShort(month) {
  const names = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
    'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return names[month - 1] || '';
}

// ==================== Number / Currency Utilities ====================

function formatCurrency(amount) {
  if (amount === null || amount === undefined || amount === '') return '฿0.00';
  return '฿' + parseFloat(amount).toLocaleString('th-TH', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatNumber(num) {
  if (num === null || num === undefined) return '0';
  return parseFloat(num).toLocaleString('th-TH');
}

function parseCurrency(str) {
  if (!str) return 0;
  return parseFloat(String(str).replace(/[฿,]/g, '')) || 0;
}

// ==================== ID Generation ====================

function generateMemberId(existingMembers) {
  const ids = existingMembers
    .map(m => parseInt((m.member_id || '').replace('MB', ''), 10))
    .filter(n => !isNaN(n));
  const max = ids.length > 0 ? Math.max(...ids) : 0;
  return 'MB' + String(max + 1).padStart(3, '0');
}

function generateLoanId(existingLoans) {
  const ids = existingLoans
    .map(l => parseInt((l.loan_id || '').replace('LN', ''), 10))
    .filter(n => !isNaN(n));
  const max = ids.length > 0 ? Math.max(...ids) : 0;
  return 'LN' + String(max + 1).padStart(3, '0');
}

function generateId(prefix, existingList, idField) {
  const ids = existingList
    .map(item => parseInt((item[idField] || '').replace(prefix, ''), 10))
    .filter(n => !isNaN(n));
  const max = ids.length > 0 ? Math.max(...ids) : 0;
  return prefix + String(max + 1).padStart(3, '0');
}

// ==================== Toast Notifications ====================

function showToast(message, type = 'success', duration = 3500) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  const icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };
  const colors = {
    success: 'bg-green-500',
    error: 'bg-red-500',
    warning: 'bg-yellow-500',
    info: 'bg-blue-500'
  };

  toast.className = `flex items-center gap-3 px-4 py-3 rounded-lg shadow-lg text-white ${colors[type]} transform translate-x-full transition-transform duration-300 max-w-sm`;
  toast.innerHTML = `
    <span class="text-lg font-bold">${icons[type] || 'ℹ'}</span>
    <span class="flex-1 text-sm">${escapeHtml(message)}</span>
    <button onclick="this.parentElement.remove()" class="text-white opacity-75 hover:opacity-100 text-lg leading-none">&times;</button>
  `;

  container.appendChild(toast);
  requestAnimationFrame(() => {
    toast.classList.remove('translate-x-full');
    toast.classList.add('translate-x-0');
  });

  setTimeout(() => {
    toast.classList.add('translate-x-full');
    toast.classList.remove('translate-x-0');
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ==================== Loading State ====================

function showLoading(show = true) {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) {
    overlay.classList.toggle('hidden', !show);
  }
}

function setButtonLoading(btn, loading, originalText) {
  if (!btn) return;
  if (loading) {
    btn.disabled = true;
    btn.dataset.originalText = btn.innerHTML;
    btn.innerHTML = '<svg class="animate-spin h-4 w-4 inline mr-2" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"></path></svg>กำลังดำเนินการ...';
  } else {
    btn.disabled = false;
    btn.innerHTML = btn.dataset.originalText || originalText || 'บันทึก';
  }
}

// ==================== Confirmation Dialog ====================

function confirmDialog(message, title = 'ยืนยันการดำเนินการ') {
  return new Promise((resolve) => {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50';
    modal.innerHTML = `
      <div class="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full mx-4">
        <div class="flex items-center gap-3 mb-4">
          <div class="w-10 h-10 rounded-full bg-yellow-100 flex items-center justify-center">
            <svg class="w-5 h-5 text-yellow-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
            </svg>
          </div>
          <h3 class="text-lg font-semibold text-gray-800">${escapeHtml(title)}</h3>
        </div>
        <p class="text-gray-600 mb-6">${escapeHtml(message)}</p>
        <div class="flex gap-3 justify-end">
          <button id="confirmNo" class="px-4 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50 transition-colors">ยกเลิก</button>
          <button id="confirmYes" class="px-4 py-2 rounded-lg bg-red-500 text-white hover:bg-red-600 transition-colors">ยืนยัน</button>
        </div>
      </div>
    `;
    document.body.appendChild(modal);
    modal.querySelector('#confirmYes').onclick = () => { modal.remove(); resolve(true); };
    modal.querySelector('#confirmNo').onclick = () => { modal.remove(); resolve(false); };
    modal.onclick = (e) => { if (e.target === modal) { modal.remove(); resolve(false); } };
  });
}

// ==================== Modal Utilities ====================

function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.remove('hidden');
    modal.classList.add('flex');
    document.body.style.overflow = 'hidden';
  }
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.add('hidden');
    modal.classList.remove('flex');
    document.body.style.overflow = '';
  }
}

function closeAllModals() {
  document.querySelectorAll('[id$="Modal"]').forEach(m => {
    m.classList.add('hidden');
    m.classList.remove('flex');
  });
  document.body.style.overflow = '';
}

// ==================== Security ====================

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ==================== Local Storage Cache ====================

const Cache = {
  set(key, data) {
    try {
      localStorage.setItem(key, JSON.stringify({ data, ts: Date.now() }));
    } catch (e) {
      console.warn('Cache write failed:', e);
    }
  },
  get(key, maxAge = CONFIG.CACHE_TTL_MS) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const { data, ts } = JSON.parse(raw);
      if (Date.now() - ts > maxAge) return null;
      return data;
    } catch (e) {
      return null;
    }
  },
  clear(key) {
    if (key) {
      localStorage.removeItem(key);
    } else {
      Object.keys(localStorage)
        .filter(k => k.startsWith('cb_'))
        .forEach(k => localStorage.removeItem(k));
    }
  },
  invalidate(sheetName) {
    localStorage.removeItem(`cb_${sheetName}`);
  }
};

// ==================== Table Search/Filter ====================

function filterTable(tableId, searchValue) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = table.querySelectorAll('tbody tr');
  const val = searchValue.toLowerCase();
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(val) ? '' : 'none';
  });
}

// ==================== Deposit Interest Rate ====================

function getDepositInterestRate(month) {
  // Jan = 3%, Feb = 2.75%, ..., Dec = 0.25%
  return CONFIG.INTEREST.DEPOSIT_BASE_RATE - (month - 1) * CONFIG.INTEREST.DEPOSIT_DECREMENT;
}

// ==================== Loan Calculations ====================

function calculateLoanPayment(principal, monthlyRate, totalMonths) {
  if (principal <= 0 || monthlyRate <= 0 || totalMonths <= 0) return 0;
  const r = monthlyRate / 100;
  return (principal * r * Math.pow(1 + r, totalMonths)) / (Math.pow(1 + r, totalMonths) - 1);
}

function calculateLoanSchedule(loanAmount, durationMonths, startDate) {
  const monthlyRate = CONFIG.INTEREST.LOAN_MONTHLY_RATE / 100;
  const monthlyPayment = calculateLoanPayment(loanAmount, CONFIG.INTEREST.LOAN_MONTHLY_RATE, durationMonths);
  const schedule = [];
  let balance = loanAmount;
  const start = new Date(startDate);

  for (let i = 1; i <= durationMonths; i++) {
    const paymentDate = new Date(start);
    paymentDate.setMonth(paymentDate.getMonth() + i);
    const interest = balance * monthlyRate;
    const principal = monthlyPayment - interest;
    balance = Math.max(0, balance - principal);
    schedule.push({
      month_number: i,
      payment_date: paymentDate.toISOString().split('T')[0],
      monthly_payment: monthlyPayment,
      principal: principal,
      interest: interest,
      remaining_balance: balance
    });
  }
  return schedule;
}

// ==================== Form Validation ====================

function validateRequired(form) {
  const required = form.querySelectorAll('[required]');
  let valid = true;
  required.forEach(field => {
    if (!field.value.trim()) {
      field.classList.add('border-red-500');
      valid = false;
    } else {
      field.classList.remove('border-red-500');
    }
  });
  return valid;
}

function clearValidation(form) {
  form.querySelectorAll('.border-red-500').forEach(f => f.classList.remove('border-red-500'));
}

// ==================== Pagination ====================

function paginateArray(arr, page, pageSize) {
  const start = (page - 1) * pageSize;
  return arr.slice(start, start + pageSize);
}

function renderPagination(container, total, currentPage, pageSize, onPageChange) {
  const totalPages = Math.ceil(total / pageSize);
  if (totalPages <= 1) { container.innerHTML = ''; return; }

  let html = '<div class="flex items-center gap-1">';
  html += `<button onclick="(${onPageChange})(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''} class="px-3 py-1 rounded border text-sm disabled:opacity-40 hover:bg-gray-50">‹</button>`;

  for (let i = 1; i <= totalPages; i++) {
    if (i === currentPage) {
      html += `<button class="px-3 py-1 rounded border text-sm bg-blue-600 text-white">${i}</button>`;
    } else if (Math.abs(i - currentPage) <= 2 || i === 1 || i === totalPages) {
      html += `<button onclick="(${onPageChange})(${i})" class="px-3 py-1 rounded border text-sm hover:bg-gray-50">${i}</button>`;
    } else if (Math.abs(i - currentPage) === 3) {
      html += '<span class="px-2 text-gray-400">...</span>';
    }
  }

  html += `<button onclick="(${onPageChange})(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''} class="px-3 py-1 rounded border text-sm disabled:opacity-40 hover:bg-gray-50">›</button>`;
  html += '</div>';
  container.innerHTML = html;
}

// ==================== Export Utilities ====================

function exportTableToCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = Array.from(table.querySelectorAll('tr'));
  const csv = rows.map(row => {
    const cells = Array.from(row.querySelectorAll('th, td'));
    return cells.map(cell => `"${cell.textContent.replace(/"/g, '""').trim()}"`).join(',');
  }).join('\n');

  const bom = '\uFEFF';
  const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename || 'export.csv';
  a.click();
  URL.revokeObjectURL(url);
}

function printSection(sectionId, title) {
  const content = document.getElementById(sectionId);
  if (!content) return;
  const printWindow = window.open('', '_blank');
  printWindow.document.write(`
    <!DOCTYPE html><html><head>
    <meta charset="utf-8">
    <title>${escapeHtml(title)}</title>
    <style>
      body { font-family: 'Sarabun', sans-serif; font-size: 12px; margin: 20px; }
      table { width: 100%; border-collapse: collapse; }
      th, td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; }
      th { background: #f5f5f5; font-weight: bold; }
      h1 { font-size: 16px; margin-bottom: 10px; }
      @media print { button { display: none; } }
    </style>
    </head><body>
    <h1>${escapeHtml(title)}</h1>
    <div>${content.innerHTML}</div>
    </body></html>
  `);
  printWindow.document.close();
  printWindow.print();
}

// ==================== Debounce ====================

function debounce(fn, delay) {
  let timer;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}
