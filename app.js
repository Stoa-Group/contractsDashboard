"use strict";
// ============================================================
//  Contracts Dashboard — Full CRUD with API Integration
//  Stoa Group Property Operations
// ============================================================

// ============================================================
//  DOMO BOOTSTRAP
// ============================================================

function getDomoQuick() {
  try {
    if (typeof window !== 'undefined' && window.domo) return window.domo;
  } catch(e) {}
  try {
    if (typeof window !== 'undefined') {
      const hasDomo = Object.prototype.hasOwnProperty.call(window, 'domo');
      if (hasDomo) return window.domo;
    }
  } catch(e) {}
  return null;
}

async function waitForDomo(maxWait = 5000) {
  let waited = 0;
  const interval = 100;
  while (waited < maxWait) {
    const obj = getDomoQuick();
    if (obj) { console.log('Domo object found after', waited, 'ms'); return obj; }
    await new Promise(r => setTimeout(r, interval));
    waited += interval;
  }
  console.warn('Domo object not found after', maxWait, 'ms');
  return null;
}

let DOMO = getDomoQuick();

(function detectLocalDev() {
  try {
    const h = window.location.hostname;
    if (h === 'localhost' || h === '127.0.0.1') {
      if (typeof API !== 'undefined' && API.setApiBaseUrl) {
        API.setApiBaseUrl('http://localhost:3000');
        console.log('[Local Dev] API base URL set to http://localhost:3000');
      }
    }
  } catch(e) {}
})();

// ============================================================
//  DOM HELPERS & STATE
// ============================================================

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

let allContracts = [];
let allProjects = [];
let allPersons = [];
let allCategories = [];
let allVendors = [];
let analyticsSummary = null;
let spendByCategoryData = [];
let spendByPropertyData = [];
let spendOverTimeData = [];

let chartInstances = {};
let currentTab = 'dashboard';
let renderedTabs = new Set();
let sortState = {};
let activeFilters = {
  search: '',
  status: [],
  propertyIds: [],
  category: ''
};
let confirmCallback = null;
let editingContractId = null;
let editingVendorId = null;
let renewingContractId = null;
let detailContractId = null;
let pendingFiles = [];

let isAuthenticated = false;
let isEditMode = false;
let currentUser = null;

// ============================================================
//  UTILITY FUNCTIONS
// ============================================================

function formatCurrency(n) {
  const v = parseFloat(n);
  if (!Number.isFinite(v)) return '$0.00';
  return '$' + v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatCurrencyWhole(n) {
  const v = parseFloat(n);
  if (!Number.isFinite(v)) return '$0';
  return '$' + Math.round(v).toLocaleString('en-US');
}

function formatCurrencyShort(n) {
  const v = parseFloat(n);
  if (!Number.isFinite(v)) return '$0';
  if (Math.abs(v) >= 1000000) return '$' + (v / 1000000).toFixed(1) + 'M';
  if (Math.abs(v) >= 1000) return '$' + (v / 1000).toFixed(1) + 'K';
  return '$' + v.toFixed(0);
}

function formatDate(d) {
  if (!d) return '—';
  const dt = new Date(d);
  if (!Number.isFinite(dt.getTime())) return '—';
  return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function formatDateShort(d) {
  if (!d) return '—';
  const dt = new Date(d);
  if (!Number.isFinite(dt.getTime())) return '—';
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  const yy = String(dt.getFullYear()).slice(-2);
  return `${mm}/${dd}/${yy}`;
}

function formatDateISO(d) {
  if (!d) return '';
  const dt = new Date(d);
  if (!Number.isFinite(dt.getTime())) return '';
  return dt.toISOString().split('T')[0];
}

function daysUntil(date) {
  if (!date) return Infinity;
  const dt = new Date(date);
  if (!Number.isFinite(dt.getTime())) return Infinity;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  dt.setHours(0, 0, 0, 0);
  return Math.ceil((dt - today) / (1000 * 60 * 60 * 24));
}

function getUrgencyClass(days) {
  if (days <= 0) return 'urgency-critical';
  if (days <= 30) return 'urgency-critical';
  if (days <= 60) return 'urgency-warning';
  return 'urgency-ok';
}

function getUrgencyLabel(days) {
  if (days <= 0) return 'Expired';
  if (days <= 7) return 'Urgent';
  if (days <= 30) return 'Critical';
  if (days <= 60) return 'Warning';
  if (days <= 90) return 'Upcoming';
  return 'OK';
}

function getStatusBadge(status) {
  const s = (status || 'Unknown').toLowerCase();
  let cls = 'badge-default';
  if (s === 'active') cls = 'badge-success';
  else if (s === 'pending' || s === 'under review') cls = 'badge-warning';
  else if (s === 'expired' || s === 'terminated' || s === 'cancelled') cls = 'badge-danger';
  else if (s === 'archived') cls = 'badge-archived';
  else if (s === 'renewed') cls = 'badge-neutral';
  return `<span class="badge ${cls}">${escapeHtml(status || 'Unknown')}</span>`;
}

function getAutoRenewBadge(autoRenew) {
  const isAuto = autoRenew === true || autoRenew === 1 || autoRenew === 'Yes' || autoRenew === 'true';
  return isAuto
    ? '<span class="badge badge-success">Auto-Renew</span>'
    : '<span class="badge badge-default">No Auto-Renew</span>';
}

function isAutoRenew(c) {
  return c.AutoRenew === true || c.AutoRenew === 1 || c.AutoRenew === 'Yes' || c.AutoRenew === 'true';
}

function escapeHtml(s) {
  if (s == null) return '';
  const d = document.createElement('div');
  d.textContent = String(s);
  return d.innerHTML;
}
const esc = escapeHtml;

function showToast(message, type = 'info') {
  const container = $('#toast-container');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  const iconMap = { success: '&#10003;', error: '&#10007;', warning: '&#9888;', info: '&#8505;' };
  const icon = iconMap[type] || iconMap.info;

  toast.innerHTML = `
    <span class="toast-icon">${icon}</span>
    <span class="toast-message">${escapeHtml(message)}</span>
    <button class="toast-close" aria-label="Dismiss">&times;</button>
  `;
  container.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('show'));

  const dismiss = () => {
    toast.classList.remove('show');
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 300);
  };
  toast.querySelector('.toast-close').addEventListener('click', dismiss);
  setTimeout(dismiss, 5000);
}

function showLoading(statusText) {
  const el = $('#loadingOverlay');
  if (el) el.style.display = 'flex';
  if (statusText) setLoadingStatus(statusText);
}

function hideLoading() {
  const el = $('#loadingOverlay');
  if (el) el.style.display = 'none';
}

function setLoadingStatus(text) {
  const el = $('#loadingStatus');
  if (el) el.textContent = text;
}

function setLoadingProgress(pct) {
  const fill = $('#loadingProgressFill');
  const txt = $('#loadingProgressText');
  if (fill) fill.style.width = Math.min(100, Math.max(0, pct)) + '%';
  if (txt) txt.textContent = Math.round(pct) + '%';
}

function debounce(fn, ms = 300) {
  let t;
  return function(...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), ms);
  };
}

function throttle(fn, ms = 100) {
  let last = 0;
  return function(...args) {
    const now = Date.now();
    if (now - last >= ms) {
      last = now;
      fn.apply(this, args);
    }
  };
}

function num(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[$,%"\u00a0]/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function clamp(val, min, max) {
  return Math.min(max, Math.max(min, val));
}

// ============================================================
//  LOOKUP HELPERS
// ============================================================

function lookupProject(id) {
  if (!id) return null;
  return allProjects.find(p =>
    (p.ProjectId != null && String(p.ProjectId) === String(id)) ||
    (p.Id != null && String(p.Id) === String(id))
  );
}

function lookupVendor(id) {
  if (!id) return null;
  return allVendors.find(v =>
    (v.VendorId != null && String(v.VendorId) === String(id)) ||
    (v.Id != null && String(v.Id) === String(id))
  );
}

function lookupCategory(id) {
  if (!id) return null;
  return allCategories.find(c =>
    (c.CategoryId != null && String(c.CategoryId) === String(id)) ||
    (c.Id != null && String(c.Id) === String(id))
  );
}

function lookupPerson(id) {
  if (!id) return null;
  return allPersons.find(p =>
    (p.PersonId != null && String(p.PersonId) === String(id)) ||
    (p.Id != null && String(p.Id) === String(id))
  );
}

function projectName(id) {
  const p = lookupProject(id);
  return p ? (p.ProjectName || p.Name || 'Unknown') : 'Unknown';
}

function projectCity(id) {
  const p = lookupProject(id);
  if (!p) return '';
  const parts = [p.City, p.State].filter(Boolean);
  return parts.join(', ');
}

function projectUnits(id) {
  const p = lookupProject(id);
  return p ? num(p.Units || p.UnitCount || 0) : 0;
}

function vendorName(id) {
  const v = lookupVendor(id);
  return v ? (v.VendorName || v.Name || 'Unknown') : 'Unknown';
}

function categoryName(id) {
  const c = lookupCategory(id);
  return c ? (c.CategoryName || c.Name || 'Unknown') : 'Uncategorized';
}

function personName(id) {
  const p = lookupPerson(id);
  return p ? (p.FullName || p.Name || 'Unknown') : '—';
}

function getActiveContracts() {
  return allContracts.filter(c => (c.Status || '').toLowerCase() === 'active');
}

function getContractMonthly(c) {
  return num(c.MonthlyCost);
}

function getContractAnnual(c) {
  const annual = num(c.AnnualCost);
  if (annual > 0) return annual;
  return getContractMonthly(c) * 12;
}

function contractMatchesSearch(c, query) {
  if (!query) return true;
  const q = query.toLowerCase();
  const fields = [
    c.Description,
    projectName(c.ProjectId),
    vendorName(c.VendorId),
    categoryName(c.CategoryId),
    c.ContractType,
    c.Status,
    c.Notes,
    c.ContractNumber,
    c.SignedBy,
    c.ContractScope,
    c.AccountRepresentative,
    c.RenewalTermType,
    c.ServiceFrequency
  ];
  return fields.some(f => f && String(f).toLowerCase().includes(q));
}

function closeModalAnimated(modalEl) {
  if (!modalEl || !modalEl.classList.contains('active')) return;
  modalEl.classList.add('closing');
  modalEl.classList.remove('active');
  setTimeout(() => {
    modalEl.classList.remove('closing');
  }, 200);
}

function formatFileSize(bytes) {
  if (!bytes || bytes <= 0) return '0 B';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function pluralize(n, singular, plural) {
  return n === 1 ? singular : (plural || singular + 's');
}

// ============================================================
//  DATA LOADING
// ============================================================

async function loadAllData() {
  setLoadingStatus('Fetching data from server…');
  setLoadingProgress(10);

  const promises = [
    API.getAllContracts(),
    API.getAllProjects(),
    API.getAllPersons(),
    API.getAllCategories(),
    API.getAllVendors(),
    API.getAnalyticsSummary(),
    API.getSpendByCategory(),
    API.getSpendByProperty(),
    API.getSpendOverTime()
  ];

  const labels = [
    'contracts', 'projects', 'persons', 'categories', 'vendors',
    'analytics summary', 'spend by category', 'spend by property', 'spend over time'
  ];

  let completed = 0;
  const total = promises.length;

  const tracked = promises.map((p, i) =>
    p.then(result => {
      completed++;
      const pct = 10 + Math.round((completed / total) * 60);
      setLoadingProgress(pct);
      setLoadingStatus(`Loaded ${labels[i]}… (${completed}/${total})`);
      return { status: 'fulfilled', value: result };
    }).catch(err => {
      completed++;
      const pct = 10 + Math.round((completed / total) * 60);
      setLoadingProgress(pct);
      console.warn(`Failed to load ${labels[i]}:`, err);
      return { status: 'rejected', reason: err };
    })
  );

  const results = await Promise.all(tracked);

  setLoadingProgress(80);
  setLoadingStatus('Processing data…');

  const extract = (r) => {
    if (r.status === 'fulfilled' && r.value && r.value.success) {
      return r.value.data || [];
    }
    return [];
  };

  allContracts = Array.isArray(extract(results[0])) ? extract(results[0]) : [];
  allProjects = Array.isArray(extract(results[1])) ? extract(results[1]) : [];
  allPersons = Array.isArray(extract(results[2])) ? extract(results[2]) : [];
  allCategories = Array.isArray(extract(results[3])) ? extract(results[3]) : [];
  allVendors = Array.isArray(extract(results[4])) ? extract(results[4]) : [];

  allCategories.forEach(c => { if (c.ContractCategoryId && !c.CategoryId) c.CategoryId = c.ContractCategoryId; });
  allContracts.forEach(c => {
    if (c.ContractCategoryId && !c.CategoryId) c.CategoryId = c.ContractCategoryId;
    if (c.DaysUntilExpiry == null && c.ExpirationDate) c.DaysUntilExpiry = daysUntil(c.ExpirationDate);
  });

  if (results[5].status === 'fulfilled' && results[5].value && results[5].value.success) {
    analyticsSummary = results[5].value.data;
  } else {
    analyticsSummary = buildLocalAnalyticsSummary();
  }

  spendByCategoryData = Array.isArray(extract(results[6])) ? extract(results[6]) : [];
  spendByPropertyData = Array.isArray(extract(results[7])) ? extract(results[7]) : [];
  spendOverTimeData = Array.isArray(extract(results[8])) ? extract(results[8]) : [];

  setLoadingProgress(90);

  const failedCount = results.filter(r => r.status === 'rejected').length;
  if (failedCount > 0) {
    console.warn(`${failedCount} of ${total} API calls failed.`);
  }

  console.log(
    `[Data] ${allContracts.length} contracts, ${allProjects.length} projects, ` +
    `${allPersons.length} persons, ${allCategories.length} categories, ${allVendors.length} vendors`
  );
}

function buildLocalAnalyticsSummary() {
  const active = getActiveContracts();
  const totalMonthlySpend = active.reduce((s, c) => s + getContractMonthly(c), 0);
  const totalAnnualSpend = active.reduce((s, c) => s + getContractAnnual(c), 0);

  const expiring30 = allContracts.filter(c => {
    const d = daysUntil(c.ExpirationDate);
    return d >= 0 && d <= 30 && (c.Status || '').toLowerCase() === 'active';
  }).length;

  const expiring60 = allContracts.filter(c => {
    const d = daysUntil(c.ExpirationDate);
    return d >= 0 && d <= 60 && (c.Status || '').toLowerCase() === 'active';
  }).length;

  const expired = allContracts.filter(c => (c.Status || '').toLowerCase() === 'expired').length;

  return {
    totalActive: active.length,
    totalMonthlySpend,
    totalAnnualSpend,
    expiring30,
    expiring60,
    expired
  };
}

async function refreshData() {
  try {
    showLoading('Refreshing data…');
    await loadAllData();
    renderedTabs.clear();
    setLoadingProgress(95);
    setLoadingStatus('Rendering…');
    renderCurrentTab();
    setLoadingProgress(100);
    setTimeout(hideLoading, 200);
    showToast('Data refreshed successfully', 'success');
  } catch (err) {
    hideLoading();
    showToast('Failed to refresh data: ' + (err.message || 'Unknown error'), 'error');
    console.error('Refresh failed:', err);
  }
}

// ============================================================
//  AUTHENTICATION
// ============================================================

async function initializeAuth() {
  const storedToken = localStorage.getItem('authToken');
  if (storedToken) {
    try {
      API.setAuthToken(storedToken);
      const result = await API.verifyAuth(storedToken);
      if (result && result.success) {
        isAuthenticated = true;
        currentUser = result.data?.user || result.user || { username: 'Admin' };
        updateAuthUI();
        console.log('[Auth] Restored session for', currentUser.username || currentUser.email);
        return;
      }
    } catch (e) {
      console.warn('[Auth] Stored token invalid, clearing.');
      localStorage.removeItem('authToken');
      API.clearAuthToken();
    }
  }

  if (DOMO) {
    try {
      const domoUser = await getDomoCurrentUser();
      if (domoUser && domoUser.email) {
        const result = await API.loginWithDomo(domoUser);
        if (result && result.success && result.data?.token) {
          localStorage.setItem('authToken', result.data.token);
          isAuthenticated = true;
          currentUser = result.data.user || { username: domoUser.name || domoUser.email };
          updateAuthUI();
          console.log('[Auth] Domo SSO login for', currentUser.username || currentUser.email);
          return;
        }
      }
    } catch (e) {
      console.warn('[Auth] Domo SSO failed:', e.message);
    }
  }

  updateAuthUI();
}

async function getDomoCurrentUser() {
  const params = new URLSearchParams(window.location.search);
  const email = params.get('userEmail');
  const name = params.get('userName');
  const userId = params.get('userId');
  if (email) return { email, name, userId };

  if (DOMO && DOMO.env) {
    const domoUserId = DOMO.env.userId;
    const domoEmail = DOMO.env.email;
    if (domoEmail) return { email: domoEmail, name: '', userId: domoUserId };
    if (domoUserId) {
      try {
        const resp = await fetch(`/api/content/v1/users/${domoUserId}`);
        const data = await resp.json();
        if (data.emailAddress) return { email: data.emailAddress, name: data.displayName || '', userId: domoUserId };
      } catch (e) {}
    }
  }
  return null;
}

async function handleLogin(e) {
  e.preventDefault();
  const username = ($('#loginUsername') || {}).value || '';
  const password = ($('#loginPassword') || {}).value || '';
  const errorEl = $('#loginError');

  if (!username || !password) {
    if (errorEl) { errorEl.textContent = 'Username and password required.'; errorEl.style.display = 'block'; }
    return;
  }

  try {
    if (errorEl) errorEl.style.display = 'none';
    const result = await API.login(username, password);
    if (result && result.success && result.data?.token) {
      localStorage.setItem('authToken', result.data.token);
      isAuthenticated = true;
      currentUser = result.data.user || { username };
      updateAuthUI();
      closeModalAnimated($('#login-modal'));
      showToast('Logged in as ' + (currentUser.username || currentUser.email || 'Admin'), 'success');
      if ($('#loginForm')) $('#loginForm').reset();
    } else {
      throw new Error(result?.error?.message || result?.message || 'Login failed');
    }
  } catch (err) {
    if (errorEl) {
      errorEl.textContent = err.message || 'Invalid credentials.';
      errorEl.style.display = 'block';
    }
  }
}

function handleLogout() {
  isAuthenticated = false;
  isEditMode = false;
  currentUser = null;
  localStorage.removeItem('authToken');
  API.clearAuthToken();
  updateAuthUI();
  renderedTabs.clear();
  renderCurrentTab();
  showToast('Logged out successfully.', 'info');
}

function toggleEditMode() {
  if (!isAuthenticated) return;
  isEditMode = !isEditMode;
  updateAuthUI();
  renderedTabs.clear();
  renderCurrentTab();
}

function updateAuthUI() {
  const adminBadge = $('#adminBadge');
  const loginBtn = $('#loginBtn');
  const editModeBtn = $('#editModeBtn');

  if (isAuthenticated) {
    if (adminBadge) adminBadge.style.display = '';
    if (loginBtn) { loginBtn.textContent = 'Logout'; loginBtn.title = 'Click to log out'; }
    if (editModeBtn) editModeBtn.style.display = '';
    document.body.classList.add('admin-authenticated');
  } else {
    if (adminBadge) adminBadge.style.display = 'none';
    if (loginBtn) { loginBtn.textContent = 'Login'; loginBtn.title = 'Click to log in'; }
    if (editModeBtn) editModeBtn.style.display = 'none';
    document.body.classList.remove('admin-authenticated');
  }

  if (isEditMode && isAuthenticated) {
    if (editModeBtn) { editModeBtn.textContent = 'Exit Edit Mode'; editModeBtn.classList.add('active'); }
    document.body.classList.add('admin-edit-mode');
  } else {
    if (editModeBtn) { editModeBtn.textContent = 'Edit Mode'; editModeBtn.classList.remove('active'); }
    document.body.classList.remove('admin-edit-mode');
  }
}

function canEdit() {
  return isAuthenticated && isEditMode;
}

// ============================================================
//  INITIALIZATION
// ============================================================

async function init() {
  try {
    showLoading('Initializing dashboard…');
    setLoadingProgress(0);

    if (!DOMO) {
      DOMO = await waitForDomo(3000);
    }

    setLoadingProgress(5);

    await loadAllData();

    setLoadingProgress(85);
    setLoadingStatus('Checking authentication…');

    await initializeAuth();

    setLoadingProgress(92);
    setLoadingStatus('Rendering dashboard…');

    bindEventListeners();
    renderCurrentTab();

    setLoadingProgress(100);
    setTimeout(hideLoading, 350);

    console.log('[Contracts Dashboard] Initialized successfully.');
  } catch (err) {
    console.error('[Contracts Dashboard] Init failed:', err);
    hideLoading();
    showToast('Failed to load dashboard: ' + (err.message || 'Unknown error'), 'error');
  }
}

// ============================================================
//  TAB SWITCHING
// ============================================================

function switchTab(tabName) {
  if (currentTab === tabName) return;
  currentTab = tabName;

  $$('.main-tab').forEach(btn => {
    const isActive = btn.getAttribute('data-tab') === tabName;
    btn.classList.toggle('active', isActive);
    btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  $$('.view').forEach(v => {
    const isActive = v.id === 'view-' + tabName;
    v.classList.toggle('active', isActive);
  });

  renderCurrentTab();
}

function renderCurrentTab() {
  const q = activeFilters.search;
  switch (currentTab) {
    case 'dashboard': renderDashboard(q); break;
    case 'property':  renderPropertyView(q); break;
    case 'category':  renderCategoryView(q); break;
    case 'vendor':    renderVendorView(q); break;
    case 'expiry':    renderExpiryView(q); break;
    case 'analytics': renderAnalyticsView(); break;
  }
  renderedTabs.add(currentTab);
}

// ============================================================
//  KPI RENDERING
// ============================================================

function renderKPIs() {
  const active = getActiveContracts();

  const totalActive = analyticsSummary
    ? (analyticsSummary.totalContracts ?? analyticsSummary.totalActive ?? active.length)
    : active.length;

  const monthlySpend = analyticsSummary
    ? num(analyticsSummary.totalMonthlySpend)
    : active.reduce((s, c) => s + getContractMonthly(c), 0);

  const annualSpend = analyticsSummary
    ? num(analyticsSummary.totalAnnualSpend)
    : active.reduce((s, c) => s + getContractAnnual(c), 0);

  const expiring30 = analyticsSummary
    ? (analyticsSummary.expiringIn30 ?? analyticsSummary.expiring30 ?? 0)
    : allContracts.filter(c => {
        const d = daysUntil(c.ExpirationDate);
        return d >= 0 && d <= 30 && (c.Status || '').toLowerCase() === 'active';
      }).length;

  const expiring60 = analyticsSummary
    ? (analyticsSummary.expiringIn60 ?? analyticsSummary.expiring60 ?? 0)
    : allContracts.filter(c => {
        const d = daysUntil(c.ExpirationDate);
        return d >= 0 && d <= 60 && (c.Status || '').toLowerCase() === 'active';
      }).length;

  const expired = analyticsSummary
    ? (analyticsSummary.expiredCount ?? analyticsSummary.expired ?? 0)
    : allContracts.filter(c => (c.Status || '').toLowerCase() === 'expired').length;

  const archived = allContracts.filter(c => (c.Status || '').toLowerCase() === 'archived').length;

  const nonCancellable = active.filter(c => c.IsNonCancellable === true || c.IsNonCancellable === 1).length;
  const withEscalation = active.filter(c => c.AnnualEscalation).length;
  const totalSetupFees = active.reduce((s, c) => s + (parseFloat(c.OneTimeSetupFee) || 0), 0);

  setKPI('#kpi-total', totalActive.toLocaleString());
  setKPI('#kpi-monthly-spend', formatCurrencyWhole(monthlySpend));
  setKPI('#kpi-annual-spend', formatCurrencyWhole(annualSpend));
  setKPI('#kpi-expiring-30', expiring30, expiring30 > 0 ? 'kpi-alert' : '');
  setKPI('#kpi-expiring-60', expiring60, expiring60 > 0 ? 'kpi-attention' : '');
  setKPI('#kpi-expired', expired, expired > 0 ? 'kpi-danger' : '');
  setKPI('#kpi-archived', archived);
  setKPI('#kpi-non-cancellable', nonCancellable, nonCancellable > 0 ? 'kpi-attention' : '');
  setKPI('#kpi-with-escalation', withEscalation);
  setKPI('#kpi-setup-fees', formatCurrencyWhole(totalSetupFees));
}

function setKPI(selector, value, extraClass) {
  const card = $(selector);
  if (!card) return;
  const valEl = card.querySelector('.kpi-value');
  if (valEl) valEl.textContent = value;
  card.classList.remove('kpi-alert', 'kpi-danger', 'kpi-attention');
  if (extraClass) card.classList.add(extraClass);
}

// ============================================================
//  DASHBOARD TAB
// ============================================================

function renderDashboard(searchQuery) {
  renderKPIs();
  renderRenewalsList(searchQuery);
  renderRecentActivity(searchQuery);
}

function renderRenewalsList(searchQuery) {
  const el = $('#renewals-list');
  if (!el) return;

  let upcoming = allContracts
    .filter(c => {
      const d = daysUntil(c.ExpirationDate);
      return d >= -7 && d <= 60 && (c.Status || '').toLowerCase() === 'active';
    });

  if (searchQuery) {
    upcoming = upcoming.filter(c => contractMatchesSearch(c, searchQuery));
  }

  upcoming.sort((a, b) => daysUntil(a.ExpirationDate) - daysUntil(b.ExpirationDate));

  if (upcoming.length === 0) {
    el.innerHTML = '<div class="empty-state"><p>No contracts expiring in the next 60 days.</p></div>';
    return;
  }

  el.innerHTML = `<div class="renewals-table">
    <div class="renewals-header">
      <span>Property</span>
      <span>Vendor</span>
      <span>Description</span>
      <span>Category</span>
      <span>Expiration</span>
      <span>Days</span>
      <span>Auto-Renew</span>
    </div>
    ${upcoming.map(c => {
      const days = daysUntil(c.ExpirationDate);
      const urg = getUrgencyClass(days);
      return `<div class="renewals-row ${urg}" data-contract-id="${c.ContractId || c.Id}" role="button" tabindex="0" title="Click to view details">
        <span class="renewals-cell">${escapeHtml(projectName(c.ProjectId))}</span>
        <span class="renewals-cell">${escapeHtml(vendorName(c.VendorId))}</span>
        <span class="renewals-cell">${escapeHtml(c.Description || '—')}</span>
        <span class="renewals-cell">${escapeHtml(categoryName(c.CategoryId))}</span>
        <span class="renewals-cell">${formatDateShort(c.ExpirationDate)}</span>
        <span class="renewals-cell"><span class="days-badge ${urg}">${days <= 0 ? 'EXPIRED' : days + 'd'}</span></span>
        <span class="renewals-cell">${getAutoRenewBadge(c.AutoRenew)}</span>
      </div>`;
    }).join('')}
  </div>`;

  el.querySelectorAll('.renewals-row').forEach(row => {
    row.addEventListener('click', () => {
      const id = parseInt(row.dataset.contractId, 10);
      if (id) openDetailModal(id);
    });
    row.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); row.click(); }
    });
  });
}

function renderRecentActivity(searchQuery) {
  const el = $('#recent-activity');
  if (!el) return;

  let sorted = [...allContracts]
    .sort((a, b) => {
      const da = new Date(a.UpdatedAt || a.CreatedAt || 0).getTime();
      const db = new Date(b.UpdatedAt || b.CreatedAt || 0).getTime();
      return db - da;
    });

  if (searchQuery) {
    sorted = sorted.filter(c => contractMatchesSearch(c, searchQuery));
  }

  sorted = sorted.slice(0, 10);

  if (sorted.length === 0) {
    el.innerHTML = '<div class="empty-state"><p>No recent activity.</p></div>';
    return;
  }

  el.innerHTML = `<div class="activity-list">
    ${sorted.map(c => {
      const dateStr = c.UpdatedAt || c.CreatedAt;
      const isUpdate = c.UpdatedAt && c.CreatedAt && new Date(c.UpdatedAt).getTime() > new Date(c.CreatedAt).getTime();
      const action = isUpdate ? 'Updated' : 'Created';
      const descText = c.Description || vendorName(c.VendorId);
      const propText = projectName(c.ProjectId);
      return `<div class="activity-item" data-contract-id="${c.ContractId || c.Id}" role="button" tabindex="0">
        <div class="activity-dot ${isUpdate ? 'dot-update' : 'dot-create'}"></div>
        <div class="activity-content">
          <div class="activity-desc">${escapeHtml(descText)} — <em>${escapeHtml(propText)}</em></div>
          <div class="activity-meta">${action} ${formatDate(dateStr)} &middot; ${getStatusBadge(c.Status)}</div>
        </div>
      </div>`;
    }).join('')}
  </div>`;

  el.querySelectorAll('.activity-item').forEach(item => {
    item.addEventListener('click', () => {
      const id = parseInt(item.dataset.contractId, 10);
      if (id) openDetailModal(id);
    });
    item.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); item.click(); }
    });
  });
}

// ============================================================
//  CHART CONFIGURATION
// ============================================================

const STOA_COLORS = [
  '#7e8a6b', '#a6ad8a', '#757270', '#bdc2ce',
  '#333333', '#6b7a5a', '#9ca3af', '#e8ebe5'
];

const STOA_COLORS_EXTENDED = [
  '#7e8a6b', '#a6ad8a', '#757270', '#bdc2ce',
  '#333333', '#6b7a5a', '#9ca3af', '#e8ebe5',
  '#5a6b4e', '#8b9178', '#636160', '#a4aab8',
  '#4d4d4d', '#556844', '#7d8590', '#d0d4c9'
];

function getChartDefaults() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          font: { family: "'Inter', 'system-ui', -apple-system, sans-serif", size: 12 },
          color: '#555',
          padding: 14,
          usePointStyle: true,
          pointStyleWidth: 10
        }
      },
      tooltip: {
        backgroundColor: 'rgba(51,51,51,0.95)',
        titleFont: { family: "'Inter', 'system-ui', sans-serif", size: 13, weight: '600' },
        bodyFont: { family: "'Inter', 'system-ui', sans-serif", size: 12 },
        padding: 12,
        cornerRadius: 6,
        displayColors: true,
        boxPadding: 4
      }
    }
  };
}

function destroyChart(key) {
  if (chartInstances[key]) {
    try { chartInstances[key].destroy(); } catch(e) {}
    delete chartInstances[key];
  }
}

function createChart(canvasId, config, key) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  key = key || canvasId;
  destroyChart(key);

  const parent = canvas.parentElement;
  if (parent && !parent.style.position) {
    parent.style.position = 'relative';
  }

  const ctx = canvas.getContext('2d');
  try {
    chartInstances[key] = new Chart(ctx, config);
  } catch (err) {
    console.error(`[Chart] Failed to create "${key}":`, err);
    return null;
  }
  return chartInstances[key];
}

// ============================================================
//  DASHBOARD CHARTS
// ============================================================

function renderSpendByCategoryChart() {
  let labels, values;

  if (spendByCategoryData.length > 0) {
    const sorted = [...spendByCategoryData].sort((a, b) =>
      num(b.TotalMonthlySpend || b.totalMonthly || b.total || 0) -
      num(a.TotalMonthlySpend || a.totalMonthly || a.total || 0)
    );
    labels = sorted.map(d => d.CategoryName || d.category || 'Other');
    values = sorted.map(d => num(d.TotalMonthlySpend || d.totalMonthly || d.total || 0));
  } else {
    const grouped = {};
    getActiveContracts().forEach(c => {
      const cat = categoryName(c.CategoryId);
      grouped[cat] = (grouped[cat] || 0) + getContractMonthly(c);
    });
    const entries = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
    labels = entries.map(e => e[0]);
    values = entries.map(e => e[1]);
  }

  if (labels.length === 0) { labels = ['No Data']; values = [0]; }

  const colors = STOA_COLORS_EXTENDED.slice(0, labels.length);

  createChart('chart-spend-category', {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors,
        borderWidth: 2,
        borderColor: '#fff',
        hoverOffset: 8
      }]
    },
    options: {
      ...getChartDefaults(),
      cutout: '58%',
      plugins: {
        ...getChartDefaults().plugins,
        legend: {
          ...getChartDefaults().plugins.legend,
          position: window.innerWidth < 1024 ? 'bottom' : 'right'
        },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: {
            label: function(ctx) {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : '0.0';
              return ` ${ctx.label}: ${formatCurrency(ctx.raw)} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderSpendByPropertyChart() {
  let data;

  if (spendByPropertyData.length > 0) {
    data = [...spendByPropertyData]
      .map(d => ({
        name: d.ProjectName || d.property || 'Unknown',
        value: num(d.TotalMonthlySpend || d.totalMonthly || d.total || 0)
      }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 12);
  } else {
    const grouped = {};
    getActiveContracts().forEach(c => {
      const pName = projectName(c.ProjectId);
      grouped[pName] = (grouped[pName] || 0) + getContractMonthly(c);
    });
    data = Object.entries(grouped)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 12);
  }

  const labels = data.map(d => d.name);
  const values = data.map(d => d.value);

  if (labels.length === 0) { labels.push('No Data'); values.push(0); }

  createChart('chart-spend-property', {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Monthly Spend',
        data: values,
        backgroundColor: STOA_COLORS[0] + 'cc',
        borderColor: STOA_COLORS[0],
        borderWidth: 1,
        borderRadius: 4,
        barPercentage: 0.7
      }]
    },
    options: {
      ...getChartDefaults(),
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        },
        y: {
          grid: { display: false },
          ticks: {
            font: { size: 11 },
            callback: function(value) {
              const label = this.getLabelForValue(value);
              return label.length > 22 ? label.substring(0, 20) + '…' : label;
            }
          }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        legend: { display: false },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: { label: ctx => ' Monthly: ' + formatCurrency(ctx.raw) }
        }
      }
    }
  });
}

function renderSpendTrendChart() {
  let labels, values;

  if (spendOverTimeData.length > 0) {
    const sorted = [...spendOverTimeData].sort((a, b) => {
      const da = a.Month || a.month || a.period || '';
      const db = b.Month || b.month || b.period || '';
      return da.localeCompare(db);
    });
    labels = sorted.map(d => {
      const raw = d.Month || d.month || d.period || '';
      try {
        const dt = new Date(raw + '-01');
        return Number.isFinite(dt.getTime())
          ? dt.toLocaleDateString('en-US', { month: 'short', year: '2-digit' })
          : raw;
      } catch(e) { return raw; }
    });
    values = sorted.map(d => num(d.TotalSpend || d.totalSpend || d.total || 0));
  } else {
    const now = new Date();
    labels = [];
    values = [];
    const monthlyTotal = getActiveContracts().reduce((s, c) => s + getContractMonthly(c), 0);
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      labels.push(d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }));
      values.push(monthlyTotal);
    }
  }

  if (labels.length === 0) { labels = ['No Data']; values = [0]; }

  createChart('chart-spend-trend', {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Monthly Spend',
        data: values,
        borderColor: STOA_COLORS[0],
        backgroundColor: STOA_COLORS[0] + '1a',
        fill: true,
        tension: 0.35,
        pointRadius: 3,
        pointHoverRadius: 6,
        pointBackgroundColor: '#fff',
        pointBorderColor: STOA_COLORS[0],
        pointBorderWidth: 2,
        borderWidth: 2.5
      }]
    },
    options: {
      ...getChartDefaults(),
      scales: {
        x: {
          grid: { color: 'rgba(0,0,0,0.04)' },
          ticks: { font: { size: 11 }, maxRotation: 45 }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        legend: { display: false },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: { label: ctx => ' ' + formatCurrency(ctx.raw) }
        }
      }
    }
  });
}

// ============================================================
//  BY PROPERTY TAB
// ============================================================

function renderPropertyView(globalSearch) {
  const tbody = $('#property-table tbody');
  if (!tbody) return;

  const localSearch = ($('#propertySearch') || {}).value || '';
  const statusFilters = activeFilters.status;
  const searchTerm = globalSearch || localSearch;

  const grouped = {};
  allContracts.forEach(c => {
    const pid = c.ProjectId;
    if (!grouped[pid]) grouped[pid] = [];
    grouped[pid].push(c);
  });

  let propertyRows = Object.entries(grouped).map(([pid, contracts]) => {
    const pidNum = parseInt(pid, 10) || pid;
    const proj = lookupProject(pidNum);
    const name = proj ? (proj.ProjectName || proj.Name || 'Unknown') : 'Unknown';
    const city = proj ? (proj.City || '') : '';
    const state = proj ? (proj.State || '') : '';
    const units = projectUnits(pidNum);

    let filteredContracts = contracts;
    if (statusFilters.length > 0 && statusFilters.length < 5) {
      filteredContracts = contracts.filter(c => statusFilters.includes((c.Status || '').toLowerCase()));
    }

    const activeContracts = filteredContracts.filter(c => (c.Status || '').toLowerCase() === 'active');
    const monthlySpend = activeContracts.reduce((s, c) => s + getContractMonthly(c), 0);
    const annualSpend = activeContracts.reduce((s, c) => s + getContractAnnual(c), 0);

    const expiringSoon = filteredContracts.filter(c => {
      const d = daysUntil(c.ExpirationDate);
      return d >= 0 && d <= 60 && (c.Status || '').toLowerCase() === 'active';
    }).length;

    const worstDays = filteredContracts
      .filter(c => (c.Status || '').toLowerCase() === 'active' && c.ExpirationDate)
      .map(c => daysUntil(c.ExpirationDate))
      .sort((a, b) => a - b)[0];

    let statusText = 'OK';
    if (worstDays != null && worstDays <= 30) statusText = 'critical';
    else if (worstDays != null && worstDays <= 60) statusText = 'warning';

    return {
      ProjectId: pidNum,
      Property: name,
      City: city,
      State: state,
      Units: units,
      ContractCount: filteredContracts.length,
      MonthlySpend: monthlySpend,
      AnnualSpend: annualSpend,
      Expiring: expiringSoon,
      Status: statusText,
      contracts: filteredContracts
    };
  });

  if (searchTerm) {
    const q = searchTerm.toLowerCase();
    propertyRows = propertyRows.filter(r =>
      r.Property.toLowerCase().includes(q) ||
      r.City.toLowerCase().includes(q) ||
      r.State.toLowerCase().includes(q) ||
      r.contracts.some(c => contractMatchesSearch(c, q))
    );
  }

  const sortKey = sortState['property-table']?.key || 'Property';
  const sortDir = sortState['property-table']?.dir || 'asc';
  propertyRows = sortData(propertyRows, sortKey, sortDir);
  updateSortIndicators('property-table', sortKey, sortDir);

  if (propertyRows.length === 0) {
    tbody.innerHTML = '<tr><td class="empty" colspan="6">No properties found.</td></tr>';
    return;
  }

  tbody.innerHTML = propertyRows.map(r => {
    const statusBadge = r.Expiring > 0
      ? `<span class="badge ${r.Status === 'critical' ? 'badge-danger' : 'badge-warning'}">Needs Attention</span>`
      : '<span class="badge badge-success">OK</span>';

    const locationLine = [r.City, r.State].filter(Boolean).join(', ');
    const unitsLine = r.Units > 0 ? ` &middot; ${r.Units} ${pluralize(r.Units, 'unit')}` : '';

    return `<tr class="property-row expandable" data-project-id="${r.ProjectId}">
      <td data-label="Property">
        <strong>${escapeHtml(r.Property)}</strong>
        ${locationLine ? `<br><small class="subtle">${escapeHtml(locationLine)}${unitsLine}</small>` : ''}
      </td>
      <td class="num" data-label="Contracts">${r.ContractCount}</td>
      <td class="num" data-label="Monthly Spend">${formatCurrency(r.MonthlySpend)}</td>
      <td class="num" data-label="Annual Spend">${formatCurrency(r.AnnualSpend)}</td>
      <td class="num ${r.Expiring > 0 ? getUrgencyClass(r.Status === 'critical' ? 15 : 45) : ''}" data-label="Expiring Soon">${r.Expiring > 0 ? r.Expiring : '—'}</td>
      <td data-label="Status">${statusBadge}</td>
    </tr>
    <tr class="detail-row" data-project-detail="${r.ProjectId}" style="display:none;">
      <td colspan="6">
        <div class="sub-table-wrap">
          ${buildContractSubTable(r.contracts)}
        </div>
      </td>
    </tr>`;
  }).join('');

  tbody.querySelectorAll('.property-row').forEach(row => {
    row.addEventListener('click', () => {
      const pid = row.dataset.projectId;
      const detailRow = tbody.querySelector(`[data-project-detail="${pid}"]`);
      if (detailRow) {
        const isOpen = detailRow.style.display !== 'none';
        const showVal = window.innerWidth <= 768 ? 'block' : 'table-row';
        detailRow.style.display = isOpen ? 'none' : showVal;
        detailRow.classList.toggle('open', !isOpen);
        row.classList.toggle('expanded', !isOpen);
      }
    });
  });

  bindSubTableClicks(tbody);
}

function buildContractSubTable(contracts) {
  if (!contracts || contracts.length === 0) {
    return '<div class="empty-state-sm">No contracts.</div>';
  }

  const sorted = [...contracts].sort((a, b) => {
    const da = daysUntil(a.ExpirationDate);
    const db = daysUntil(b.ExpirationDate);
    if (da === Infinity && db === Infinity) return 0;
    if (da === Infinity) return 1;
    if (db === Infinity) return -1;
    return da - db;
  });

  return `<table class="table sub-table">
    <thead>
      <tr>
        <th>Description</th>
        <th>Vendor</th>
        <th>Category</th>
        <th>Status</th>
        <th class="num">Monthly Cost</th>
        <th>Billing</th>
        <th>Expiration</th>
        <th>Days</th>
        <th>Auto-Renew</th>
        <th>Escalation</th>
        <th>Signed By</th>
      </tr>
    </thead>
    <tbody>
      ${sorted.map(c => {
        const days = daysUntil(c.ExpirationDate);
        const isActive = (c.Status || '').toLowerCase() === 'active';
        const urg = isActive && days !== Infinity ? getUrgencyClass(days) : '';
        const daysDisplay = days === Infinity ? '—' : (days <= 0 ? 'Expired' : days + 'd');
        const ncBadge = (c.IsNonCancellable === true || c.IsNonCancellable === 1) ? ' <span class="badge badge-danger badge-sm">NC</span>' : '';
        return `<tr class="contract-row ${urg}" data-contract-id="${c.ContractId || c.Id}" role="button" tabindex="0" title="Click for details">
          <td data-label="Description">${escapeHtml(c.Description || '—')}${ncBadge}</td>
          <td data-label="Vendor">${escapeHtml(vendorName(c.VendorId))}</td>
          <td data-label="Category">${escapeHtml(categoryName(c.CategoryId))}</td>
          <td data-label="Status">${getStatusBadge(c.Status)}</td>
          <td class="num" data-label="Monthly Cost">${formatCurrency(c.MonthlyCost)}</td>
          <td data-label="Billing">${escapeHtml(c.BillingFrequency || '—')}</td>
          <td data-label="Expiration">${formatDateShort(c.ExpirationDate)}</td>
          <td class="${urg}" data-label="Days">${daysDisplay}</td>
          <td data-label="Auto-Renew">${getAutoRenewBadge(c.AutoRenew)}</td>
          <td data-label="Escalation">${escapeHtml(c.AnnualEscalation || '—')}</td>
          <td data-label="Signed By">${escapeHtml(c.SignedBy || '—')}</td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>`;
}

function bindSubTableClicks(container) {
  container.querySelectorAll('.contract-row').forEach(row => {
    if (row.dataset._bound) return;
    row.dataset._bound = '1';
    row.addEventListener('click', e => {
      e.stopPropagation();
      const id = parseInt(row.dataset.contractId, 10);
      if (id) openDetailModal(id);
    });
    row.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); row.click(); }
    });
  });
}

// ============================================================
//  BY CATEGORY TAB
// ============================================================

function renderCategoryView(globalSearch) {
  const grid = $('#category-grid');
  if (!grid) return;

  const grouped = {};
  allContracts.forEach(c => {
    if (globalSearch && !contractMatchesSearch(c, globalSearch)) return;
    const cid = c.CategoryId || 0;
    if (!grouped[cid]) grouped[cid] = [];
    grouped[cid].push(c);
  });

  const categories = Object.entries(grouped).map(([cid, contracts]) => {
    const catObj = lookupCategory(parseInt(cid, 10) || cid);
    const name = catObj ? (catObj.CategoryName || catObj.Name || 'Uncategorized') : 'Uncategorized';
    const activeContracts = contracts.filter(c => (c.Status || '').toLowerCase() === 'active');
    const monthlySpend = activeContracts.reduce((s, c) => s + getContractMonthly(c), 0);
    const annualSpend = activeContracts.reduce((s, c) => s + getContractAnnual(c), 0);
    const expiringSoon = activeContracts.filter(c => daysUntil(c.ExpirationDate) <= 60 && daysUntil(c.ExpirationDate) >= 0).length;

    return {
      CategoryId: parseInt(cid, 10) || cid,
      name,
      count: contracts.length,
      activeCount: activeContracts.length,
      monthlySpend,
      annualSpend,
      expiringSoon,
      contracts
    };
  }).sort((a, b) => b.monthlySpend - a.monthlySpend);

  if (categories.length === 0) {
    grid.innerHTML = '<div class="empty-state"><p>No categories found.</p></div>';
    return;
  }

  grid.innerHTML = categories.map(cat => {
    const pctOfTotal = (() => {
      const total = categories.reduce((s, c) => s + c.monthlySpend, 0);
      return total > 0 ? ((cat.monthlySpend / total) * 100).toFixed(1) : '0.0';
    })();

    return `<div class="category-card" data-category-id="${cat.CategoryId}">
      <div class="category-card-header">
        <h3 class="category-card-title">${escapeHtml(cat.name)}</h3>
        <div class="category-card-badges">
          <span class="badge badge-default">${cat.count} ${pluralize(cat.count, 'contract')}</span>
          ${cat.expiringSoon > 0 ? `<span class="badge badge-warning">${cat.expiringSoon} expiring</span>` : ''}
        </div>
      </div>
      <div class="category-card-stats">
        <div class="stat">
          <span class="stat-label">Monthly</span>
          <span class="stat-value">${formatCurrencyWhole(cat.monthlySpend)}</span>
        </div>
        <div class="stat">
          <span class="stat-label">Annual</span>
          <span class="stat-value">${formatCurrencyWhole(cat.annualSpend)}</span>
        </div>
        <div class="stat">
          <span class="stat-label">Active</span>
          <span class="stat-value">${cat.activeCount}</span>
        </div>
        <div class="stat">
          <span class="stat-label">% of Spend</span>
          <span class="stat-value">${pctOfTotal}%</span>
        </div>
      </div>
      <div class="category-card-expand" data-cat-detail="${cat.CategoryId}" style="display:none;">
        ${buildContractSubTable(cat.contracts)}
      </div>
      <button class="category-card-toggle" type="button" aria-expanded="false">
        <span class="toggle-text">Show Contracts</span>
        <span class="toggle-icon">&#9660;</span>
      </button>
    </div>`;
  }).join('');

  grid.querySelectorAll('.category-card-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      const card = btn.closest('.category-card');
      const detail = card.querySelector('.category-card-expand');
      const isOpen = detail.style.display !== 'none';
      detail.style.display = isOpen ? 'none' : 'block';
      btn.querySelector('.toggle-text').textContent = isOpen ? 'Show Contracts' : 'Hide Contracts';
      btn.querySelector('.toggle-icon').textContent = isOpen ? '\u25BC' : '\u25B2';
      btn.setAttribute('aria-expanded', isOpen ? 'false' : 'true');
    });
  });

  bindSubTableClicks(grid);
}

// ============================================================
//  BY VENDOR TAB
// ============================================================

function renderVendorView(globalSearch) {
  const tbody = $('#vendor-table tbody');
  if (!tbody) return;

  const localSearch = ($('#vendorSearch') || {}).value || '';
  const searchTerm = globalSearch || localSearch;

  const contractsByVendor = {};
  allContracts.forEach(c => {
    const vid = c.VendorId || 0;
    if (!contractsByVendor[vid]) contractsByVendor[vid] = [];
    contractsByVendor[vid].push(c);
  });

  let vendorRows = allVendors.map(v => {
    const vid = v.VendorId || v.Id;
    const contracts = contractsByVendor[vid] || [];
    const activeContracts = contracts.filter(c => (c.Status || '').toLowerCase() === 'active');
    const totalMonthly = activeContracts.reduce((s, c) => s + getContractMonthly(c), 0);
    const totalAnnual = activeContracts.reduce((s, c) => s + getContractAnnual(c), 0);

    return {
      VendorId: vid,
      VendorName: v.VendorName || v.Name || 'Unknown',
      ContactName: v.ContactName || '—',
      Email: v.ContactEmail || v.Email || '—',
      Phone: v.ContactPhone || v.Phone || '—',
      Website: v.Website || '',
      Address: v.Address || '',
      ContractCount: contracts.length,
      ActiveCount: activeContracts.length,
      TotalSpend: totalAnnual,
      MonthlySpend: totalMonthly,
      contracts
    };
  });

  const unlinkedVids = new Set(allVendors.map(v => String(v.VendorId || v.Id)));
  Object.entries(contractsByVendor).forEach(([vid, contracts]) => {
    if (!unlinkedVids.has(String(vid)) && vid !== '0' && vid !== 'undefined' && vid !== 'null') {
      const totalAnnual = contracts.filter(c => (c.Status || '').toLowerCase() === 'active')
        .reduce((s, c) => s + getContractAnnual(c), 0);
      vendorRows.push({
        VendorId: vid,
        VendorName: 'Unknown Vendor #' + vid,
        ContactName: '—',
        Email: '—',
        Phone: '—',
        Website: '',
        Address: '',
        ContractCount: contracts.length,
        ActiveCount: contracts.filter(c => (c.Status || '').toLowerCase() === 'active').length,
        TotalSpend: totalAnnual,
        MonthlySpend: 0,
        contracts
      });
    }
  });

  if (searchTerm) {
    const q = searchTerm.toLowerCase();
    vendorRows = vendorRows.filter(r =>
      r.VendorName.toLowerCase().includes(q) ||
      r.ContactName.toLowerCase().includes(q) ||
      r.Email.toLowerCase().includes(q) ||
      r.Phone.toLowerCase().includes(q) ||
      r.contracts.some(c => contractMatchesSearch(c, q))
    );
  }

  const sortKey = sortState['vendor-table']?.key || 'VendorName';
  const sortDir = sortState['vendor-table']?.dir || 'asc';
  vendorRows = sortData(vendorRows, sortKey, sortDir);
  updateSortIndicators('vendor-table', sortKey, sortDir);

  if (vendorRows.length === 0) {
    tbody.innerHTML = '<tr><td class="empty" colspan="7">No vendors found.</td></tr>';
    return;
  }

  tbody.innerHTML = vendorRows.map(r => `
    <tr class="vendor-row expandable" data-vendor-id="${r.VendorId}">
      <td data-label="Vendor">
        <strong>${escapeHtml(r.VendorName)}</strong>
        ${r.Website ? `<br><a href="${escapeHtml(r.Website)}" target="_blank" rel="noopener" class="subtle-link">${escapeHtml(r.Website.replace(/^https?:\/\//, ''))}</a>` : ''}
        ${r.Address ? `<br><small class="subtle">${escapeHtml(r.Address)}</small>` : ''}
      </td>
      <td data-label="Contact">${escapeHtml(r.ContactName)}</td>
      <td data-label="Email">${r.Email !== '—' ? `<a href="mailto:${escapeHtml(r.Email)}" class="subtle-link">${escapeHtml(r.Email)}</a>` : '—'}</td>
      <td data-label="Phone">${escapeHtml(r.Phone)}</td>
      <td class="num" data-label="Contracts">${r.ContractCount}${r.ActiveCount < r.ContractCount ? ` <small class="subtle">(${r.ActiveCount} active)</small>` : ''}</td>
      <td class="num" data-label="Annual Spend">${formatCurrency(r.TotalSpend)}</td>
      <td class="actions-cell" data-label="" onclick="event.stopPropagation()">
        ${canEdit() ? `<button class="btn btn-xs btn-edit-vendor" data-vendor-id="${r.VendorId}" title="Edit vendor">Edit</button>
        <button class="btn btn-xs btn-danger btn-delete-vendor" data-vendor-id="${r.VendorId}" title="Delete vendor">Del</button>` : ''}
      </td>
    </tr>
    <tr class="detail-row" data-vendor-detail="${r.VendorId}" style="display:none;">
      <td colspan="7">
        <div class="sub-table-wrap">
          ${buildContractSubTable(r.contracts)}
        </div>
      </td>
    </tr>
  `).join('');

  tbody.querySelectorAll('.vendor-row').forEach(row => {
    row.addEventListener('click', e => {
      if (e.target.closest('.actions-cell')) return;
      const vid = row.dataset.vendorId;
      const detailRow = tbody.querySelector(`[data-vendor-detail="${vid}"]`);
      if (detailRow) {
        const isOpen = detailRow.style.display !== 'none';
        const showVal = window.innerWidth <= 768 ? 'block' : 'table-row';
        detailRow.style.display = isOpen ? 'none' : showVal;
        detailRow.classList.toggle('open', !isOpen);
        row.classList.toggle('expanded', !isOpen);
      }
    });
  });

  tbody.querySelectorAll('.btn-edit-vendor').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      openVendorModal(parseInt(btn.dataset.vendorId, 10));
    });
  });

  tbody.querySelectorAll('.btn-delete-vendor').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      deleteVendor(parseInt(btn.dataset.vendorId, 10));
    });
  });

  bindSubTableClicks(tbody);
}

// ============================================================
//  EXPIRY TRACKER TAB
// ============================================================

function renderExpiryView(globalSearch) {
  renderExpiryTable(globalSearch);
  renderExpiryTimeline(globalSearch);
}

function renderExpiryTable(globalSearch) {
  const tbody = $('#expiry-table tbody');
  if (!tbody) return;

  const timeframe = parseInt(($('#expiryTimeframe') || {}).value || '90', 10);
  const autoRenewFilter = ($('#expiryAutoRenewFilter') || {}).value || 'all';

  let expiring = allContracts.filter(c => {
    const days = daysUntil(c.ExpirationDate);
    const isActive = (c.Status || '').toLowerCase() === 'active';
    return isActive && days >= -30 && days <= timeframe;
  });

  if (autoRenewFilter === 'auto') {
    expiring = expiring.filter(c => isAutoRenew(c));
  } else if (autoRenewFilter === 'manual') {
    expiring = expiring.filter(c => !isAutoRenew(c));
  }

  if (globalSearch) {
    expiring = expiring.filter(c => contractMatchesSearch(c, globalSearch));
  }

  expiring = expiring.map(c => ({
    ...c,
    Property: projectName(c.ProjectId),
    Vendor: vendorName(c.VendorId),
    Category: categoryName(c.CategoryId),
    DaysUntil: daysUntil(c.ExpirationDate),
    _ExpirationDate: c.ExpirationDate,
    _AutoRenew: isAutoRenew(c) ? 1 : 0,
    _NoticePeriod: num(c.NoticePeriodDays || c.NoticePeriod || 0)
  }));

  const sortKey = sortState['expiry-table']?.key || 'DaysUntil';
  const sortDir = sortState['expiry-table']?.dir || 'asc';
  expiring = sortData(expiring, sortKey, sortDir);
  updateSortIndicators('expiry-table', sortKey, sortDir);

  if (expiring.length === 0) {
    tbody.innerHTML = `<tr><td class="empty" colspan="9">No contracts expiring in the next ${timeframe} days.</td></tr>`;
    return;
  }

  tbody.innerHTML = expiring.map(c => {
    const days = c.DaysUntil;
    const urg = getUrgencyClass(days);
    const noticeDays = num(c.NoticePeriodDays || c.NoticePeriod || 0);
    const noticeWarning = noticeDays > 0 && days <= noticeDays && days > 0;

    const ncBadge = (c.IsNonCancellable === true || c.IsNonCancellable === 1) ? ' <span class="badge badge-danger badge-sm">NC</span>' : '';
    const renewType = c.RenewalTermType || '—';

    return `<tr class="${urg}">
      <td data-label="Property">${escapeHtml(c.Property)}</td>
      <td data-label="Vendor">${escapeHtml(c.Vendor)}</td>
      <td data-label="Category">${escapeHtml(c.Category)}</td>
      <td data-label="Expiration">${formatDateShort(c.ExpirationDate)}</td>
      <td class="num" data-label="Days Left">
        <span class="days-badge ${urg}">${days <= 0 ? 'EXPIRED' : days + 'd'}</span>
        ${noticeWarning ? '<span class="notice-flag" title="Within notice period">!</span>' : ''}
      </td>
      <td data-label="Auto-Renew">${getAutoRenewBadge(c.AutoRenew)}${ncBadge}</td>
      <td class="num" data-label="Notice">${noticeDays > 0 ? noticeDays + 'd' : '—'}</td>
      <td data-label="Renewal Type">${escapeHtml(renewType)}</td>
      <td class="actions-cell" data-label="">
        <button class="btn btn-xs" onclick="openDetailModal(${c.ContractId || c.Id})">View</button>
        ${canEdit() ? `<button class="btn btn-xs" onclick="openRenewModal(${c.ContractId || c.Id})">Renew</button>` : ''}
      </td>
    </tr>`;
  }).join('');
}

function renderExpiryTimeline(globalSearch) {
  const el = $('#expiry-timeline');
  if (!el) return;

  const timeframe = parseInt(($('#expiryTimeframe') || {}).value || '90', 10);

  let expiring = allContracts
    .filter(c => {
      const days = daysUntil(c.ExpirationDate);
      return (c.Status || '').toLowerCase() === 'active' && days >= 0 && days <= timeframe;
    });

  if (globalSearch) {
    expiring = expiring.filter(c => contractMatchesSearch(c, globalSearch));
  }

  expiring.sort((a, b) => daysUntil(a.ExpirationDate) - daysUntil(b.ExpirationDate));

  if (expiring.length === 0) {
    el.innerHTML = '<div class="empty-state"><p>No upcoming expirations to display on timeline.</p></div>';
    return;
  }

  const monthGroups = {};
  expiring.forEach(c => {
    const dt = new Date(c.ExpirationDate);
    const key = dt.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    if (!monthGroups[key]) monthGroups[key] = [];
    monthGroups[key].push(c);
  });

  el.innerHTML = `
    <h3 class="timeline-title">Expiration Timeline</h3>
    <div class="timeline">
      ${Object.entries(monthGroups).map(([month, contracts]) => {
        const monthUrgency = (() => {
          const minDays = Math.min(...contracts.map(c => daysUntil(c.ExpirationDate)));
          return getUrgencyClass(minDays);
        })();

        return `<div class="timeline-group">
          <div class="timeline-month ${monthUrgency}">${escapeHtml(month)} <small>(${contracts.length})</small></div>
          <div class="timeline-items">
            ${contracts.map(c => {
              const days = daysUntil(c.ExpirationDate);
              const urg = getUrgencyClass(days);
              const desc = c.Description || vendorName(c.VendorId);
              return `<div class="timeline-item ${urg}" data-contract-id="${c.ContractId || c.Id}" role="button" tabindex="0" title="Click for details">
                <div class="timeline-date">${formatDateShort(c.ExpirationDate)}</div>
                <div class="timeline-info">
                  <strong>${escapeHtml(desc)}</strong>
                  <span>${escapeHtml(projectName(c.ProjectId))}</span>
                </div>
                <div class="timeline-days">
                  <span class="days-badge ${urg}">${days}d</span>
                  ${isAutoRenew(c) ? '<span class="auto-renew-indicator" title="Auto-renew">&#8635;</span>' : ''}
                </div>
              </div>`;
            }).join('')}
          </div>
        </div>`;
      }).join('')}
    </div>
  `;

  el.querySelectorAll('.timeline-item').forEach(item => {
    item.addEventListener('click', () => {
      const id = parseInt(item.dataset.contractId, 10);
      if (id) openDetailModal(id);
    });
    item.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); item.click(); }
    });
  });
}

// ============================================================
//  ANALYTICS TAB
// ============================================================

let analyticsSubView = 'spend-by-category';

function renderAnalyticsView() {
  showAnalyticsSubView(analyticsSubView);
}

function showAnalyticsSubView(view) {
  analyticsSubView = view;

  $$('.view-switch').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === view);
  });

  $$('.analytics-charts .panel').forEach(panel => {
    const chartType = panel.dataset.chart;
    panel.style.display = chartType === view ? 'block' : 'none';
  });

  switch (view) {
    case 'spend-by-category': renderSpendByCategoryChart(); break;
    case 'spend-by-property': renderSpendByPropertyChart(); break;
    case 'spend-over-time':   renderAnalyticsSpendOverTime(); break;
    case 'cost-per-unit':     renderAnalyticsCostPerUnit(); break;
    case 'category-trends':   renderAnalyticsCategoryTrends(); break;
    case 'vendor-analysis':   renderAnalyticsVendorConcentration(); break;
    case 'service-matrix':    renderServiceMatrix(); break;
    case 'completeness':      renderCompleteness(); break;
    case 'service-coverage':  renderServiceCoverage(); break;
    case 'vendor-dependency': renderVendorDependency(); break;
    case 'renewal-pipeline':  renderRenewalPipeline(); break;
    case 'cost-benchmarks':   renderCostBenchmarks(); break;
  }
}

function renderAnalyticsSpendOverTime() {
  let labels, values;

  if (spendOverTimeData.length > 0) {
    const sorted = [...spendOverTimeData].sort((a, b) => {
      const da = a.Month || a.month || a.period || '';
      const db = b.Month || b.month || b.period || '';
      return da.localeCompare(db);
    });
    labels = sorted.map(d => {
      const raw = d.Month || d.month || d.period || '';
      try {
        const dt = new Date(raw + '-01');
        return Number.isFinite(dt.getTime())
          ? dt.toLocaleDateString('en-US', { month: 'short', year: '2-digit' })
          : raw;
      } catch(e) { return raw; }
    });
    values = sorted.map(d => num(d.TotalSpend || d.totalSpend || d.total || 0));
  } else {
    const now = new Date();
    labels = [];
    values = [];
    const monthlyTotal = getActiveContracts().reduce((s, c) => s + getContractMonthly(c), 0);
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      labels.push(d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }));
      values.push(monthlyTotal);
    }
  }

  if (labels.length === 0) { labels = ['No Data']; values = [0]; }

  const avg = values.reduce((a, b) => a + b, 0) / (values.length || 1);
  const avgLine = values.map(() => avg);

  createChart('chart-spend-over-time', {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'Total Monthly Spend',
          data: values,
          borderColor: STOA_COLORS[0],
          backgroundColor: STOA_COLORS[0] + '1a',
          fill: true,
          tension: 0.35,
          pointRadius: 4,
          pointHoverRadius: 7,
          pointBackgroundColor: '#fff',
          pointBorderColor: STOA_COLORS[0],
          pointBorderWidth: 2,
          borderWidth: 2.5,
          order: 1
        },
        {
          label: 'Average',
          data: avgLine,
          borderColor: STOA_COLORS[2],
          borderWidth: 1.5,
          borderDash: [6, 4],
          pointRadius: 0,
          fill: false,
          order: 2
        }
      ]
    },
    options: {
      ...getChartDefaults(),
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: {
          grid: { color: 'rgba(0,0,0,0.04)' },
          ticks: { font: { size: 11 }, maxRotation: 45 }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatCurrency(ctx.raw)}` }
        }
      }
    }
  });
}

function renderAnalyticsCostPerUnit() {
  const grouped = {};
  getActiveContracts().forEach(c => {
    const pid = c.ProjectId;
    if (!grouped[pid]) grouped[pid] = { monthly: 0 };
    grouped[pid].monthly += getContractMonthly(c);
  });

  const rows = Object.entries(grouped).map(([pid, data]) => {
    const name = projectName(parseInt(pid, 10) || pid);
    const units = projectUnits(parseInt(pid, 10) || pid);
    const perUnit = units > 0 ? data.monthly / units : 0;
    return { name, perUnit, units, monthly: data.monthly };
  }).filter(r => r.perUnit > 0 && r.units > 0).sort((a, b) => b.perUnit - a.perUnit).slice(0, 15);

  const labels = rows.map(r => r.name);
  const values = rows.map(r => Math.round(r.perUnit * 100) / 100);

  if (labels.length === 0) {
    labels.push('No unit data');
    values.push(0);
  }

  const avgCost = rows.length > 0
    ? rows.reduce((s, r) => s + r.perUnit, 0) / rows.length
    : 0;

  createChart('chart-cost-per-unit', {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Cost Per Unit (Monthly)',
        data: values,
        backgroundColor: values.map(v => v > avgCost * 1.3 ? STOA_COLORS[4] + 'aa' : STOA_COLORS[1] + 'cc'),
        borderColor: values.map(v => v > avgCost * 1.3 ? STOA_COLORS[4] : STOA_COLORS[1]),
        borderWidth: 1,
        borderRadius: 4,
        barPercentage: 0.7
      }]
    },
    options: {
      ...getChartDefaults(),
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        },
        y: {
          grid: { display: false },
          ticks: {
            font: { size: 11 },
            callback: function(value) {
              const label = this.getLabelForValue(value);
              return label.length > 25 ? label.substring(0, 23) + '…' : label;
            }
          }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        legend: { display: false },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: {
            label: ctx => {
              const r = rows[ctx.dataIndex];
              if (!r) return ` ${formatCurrency(ctx.raw)}/unit`;
              return ` ${formatCurrency(ctx.raw)}/unit  (${r.units} units, ${formatCurrency(r.monthly)} total)`;
            }
          }
        }
      }
    }
  });
}

function renderAnalyticsCategoryTrends() {
  const catNames = [...new Set(allContracts.map(c => categoryName(c.CategoryId)))].sort();
  const now = new Date();
  const months = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }));
  }

  const topCats = catNames.slice(0, 8);

  const datasets = topCats.map((cat, i) => {
    const contracts = getActiveContracts().filter(c => categoryName(c.CategoryId) === cat);
    const monthly = contracts.reduce((s, c) => s + getContractMonthly(c), 0);
    const color = STOA_COLORS_EXTENDED[i % STOA_COLORS_EXTENDED.length];

    return {
      label: cat,
      data: months.map(() => monthly),
      backgroundColor: color + '99',
      borderColor: color,
      borderWidth: 1
    };
  });

  createChart('chart-category-trends', {
    type: 'bar',
    data: { labels: months, datasets },
    options: {
      ...getChartDefaults(),
      scales: {
        x: {
          stacked: true,
          grid: { display: false },
          ticks: { font: { size: 11 } }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          mode: 'index',
          callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatCurrency(ctx.raw)}` }
        }
      }
    }
  });
}

function renderAnalyticsVendorConcentration() {
  const grouped = {};
  getActiveContracts().forEach(c => {
    const vName = vendorName(c.VendorId);
    grouped[vName] = (grouped[vName] || 0) + getContractAnnual(c);
  });

  const sorted = Object.entries(grouped)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const labels = sorted.map(s => s[0]);
  const values = sorted.map(s => s[1]);
  const totalAnnual = Object.values(grouped).reduce((a, b) => a + b, 0);

  if (labels.length === 0) { labels.push('No Data'); values.push(0); }

  createChart('chart-vendor-concentration', {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Annual Spend',
        data: values,
        backgroundColor: sorted.map((_, i) => {
          const pct = totalAnnual > 0 ? (_[1] / totalAnnual) * 100 : 0;
          return pct > 20 ? STOA_COLORS[4] + 'cc' : STOA_COLORS[0] + 'cc';
        }),
        borderColor: sorted.map((_, i) => {
          const pct = totalAnnual > 0 ? (_[1] / totalAnnual) * 100 : 0;
          return pct > 20 ? STOA_COLORS[4] : STOA_COLORS[0];
        }),
        borderWidth: 1,
        borderRadius: 4,
        barPercentage: 0.7
      }]
    },
    options: {
      ...getChartDefaults(),
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: { callback: v => formatCurrencyShort(v), font: { size: 11 } }
        },
        y: {
          grid: { display: false },
          ticks: {
            font: { size: 11 },
            callback: function(value) {
              const label = this.getLabelForValue(value);
              return label.length > 25 ? label.substring(0, 23) + '…' : label;
            }
          }
        }
      },
      plugins: {
        ...getChartDefaults().plugins,
        legend: { display: false },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: {
            label: ctx => {
              const pct = totalAnnual > 0 ? ((ctx.raw / totalAnnual) * 100).toFixed(1) : '0.0';
              return ` ${formatCurrency(ctx.raw)} (${pct}% of total)`;
            }
          }
        }
      }
    }
  });
}

// ============================================================
//  ENHANCED ANALYTICS – SERVICE MATRIX
// ============================================================

async function renderServiceMatrix() {
  const container = document.getElementById('service-matrix-container');
  if (!container) return;
  container.innerHTML = '<p class="text-muted" style="padding:16px;">Loading service matrix…</p>';

  try {
    const res = await API.getServiceMatrix();
    const rows = (res && res.data) || [];

    if (rows.length === 0) {
      container.innerHTML = '<p class="text-muted" style="padding:16px;">No active contracts found.</p>';
      return;
    }

    const categories = [...new Set(rows.map(r => r.CategoryName || 'Uncategorized'))].sort();
    const properties = [...new Set(rows.map(r => r.ProjectName || 'Unknown'))].sort();

    const lookup = {};
    rows.forEach(r => {
      const key = (r.CategoryName || 'Uncategorized') + '||' + (r.ProjectName || 'Unknown');
      if (!lookup[key]) lookup[key] = [];
      lookup[key].push(r);
    });

    let html = `
      <div class="matrix-legend">
        <span class="matrix-legend-item"><span class="matrix-legend-swatch" style="background:#dcfce7;"></span> Covered</span>
        <span class="matrix-legend-item"><span class="matrix-legend-swatch" style="background:#fee2e2;"></span> Missing</span>
        <span class="matrix-legend-item"><span class="matrix-legend-swatch" style="background:#fef9c3;"></span> No cost data</span>
      </div>
      <table class="service-matrix">
        <thead><tr><th>Service</th>`;

    properties.forEach(p => {
      const short = p.length > 18 ? p.substring(0, 16) + '…' : p;
      html += `<th title="${esc(p)}">${esc(short)}</th>`;
    });
    html += '</tr></thead><tbody>';

    categories.forEach(cat => {
      html += `<tr><td>${esc(cat)}</td>`;
      properties.forEach(prop => {
        const key = cat + '||' + prop;
        const contracts = lookup[key];
        if (!contracts || contracts.length === 0) {
          html += '<td class="matrix-cell matrix-cell-empty">—</td>';
        } else {
          const c = contracts[0];
          const cost = parseFloat(c.MonthlyCost);
          const hasCost = Number.isFinite(cost) && cost > 0;
          const bgClass = hasCost ? '' : 'style="background:#fef9c3;"';
          html += `<td class="matrix-cell" ${bgClass}>
            <div class="matrix-cell-vendor">${esc(c.VendorName || '—')}</div>
            <div class="matrix-cell-cost">${hasCost ? formatCurrency(cost) + '/mo' : 'No cost'}</div>
            ${contracts.length > 1 ? '<div class="matrix-cell-cost">+' + (contracts.length - 1) + ' more</div>' : ''}
          </td>`;
        }
      });
      html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
  } catch (err) {
    console.error('Service Matrix error:', err);
    container.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load service matrix.</p>';
  }
}

// ============================================================
//  ENHANCED ANALYTICS – COMPLETENESS SCORECARD
// ============================================================

async function renderCompleteness() {
  const container = document.getElementById('completeness-container');
  if (!container) return;
  container.innerHTML = '<p class="text-muted" style="padding:16px;">Loading completeness data…</p>';

  try {
    const res = await API.getCompleteness();
    const rows = (res && res.data) || [];

    if (rows.length === 0) {
      container.innerHTML = '<p class="text-muted" style="padding:16px;">No data available.</p>';
      return;
    }

    const fields = ['HasCost', 'HasStartDate', 'HasEndDate', 'HasNotice', 'HasBilling'];
    const fieldLabels = { HasCost: 'Cost', HasStartDate: 'Start Date', HasEndDate: 'End Date', HasNotice: 'Notice', HasBilling: 'Billing' };

    let totalScore = 0;
    let totalMax = 0;
    let fullyComplete = 0;

    const cards = rows.map(row => {
      const total = parseInt(row.TotalContracts) || 1;
      let score = 0;
      const bars = fields.map(f => {
        const val = parseInt(row[f]) || 0;
        const pct = Math.round((val / total) * 100);
        score += pct;
        const cls = pct >= 80 ? 'fill-high' : pct >= 40 ? 'fill-med' : 'fill-low';
        return `<div class="completeness-bar-row">
          <span class="completeness-bar-label">${fieldLabels[f]}</span>
          <div class="completeness-bar-track"><div class="completeness-bar-fill ${cls}" style="width:${pct}%"></div></div>
          <span class="completeness-bar-val">${val}/${total}</span>
        </div>`;
      }).join('');

      const avg = Math.round(score / fields.length);
      totalScore += avg;
      totalMax += 100;
      if (avg >= 95) fullyComplete++;
      const scoreCls = avg >= 80 ? 'score-high' : avg >= 40 ? 'score-med' : 'score-low';

      return `<div class="completeness-card">
        <div class="completeness-card-header">
          <span class="completeness-card-title">${esc(row.ProjectName)}</span>
          <span class="completeness-card-score ${scoreCls}">${avg}%</span>
        </div>
        ${bars}
      </div>`;
    }).join('');

    const portfolioAvg = rows.length > 0 ? Math.round(totalScore / rows.length) : 0;

    const summary = `<div class="completeness-summary">
      <div class="completeness-summary-stat">
        <div class="stat-value">${portfolioAvg}%</div>
        <div class="stat-label">Portfolio Avg</div>
      </div>
      <div class="completeness-summary-stat">
        <div class="stat-value">${fullyComplete}</div>
        <div class="stat-label">Fully Complete</div>
      </div>
      <div class="completeness-summary-stat">
        <div class="stat-value">${rows.length}</div>
        <div class="stat-label">Properties</div>
      </div>
    </div>`;

    container.innerHTML = summary + '<div class="completeness-grid">' + cards + '</div>';
  } catch (err) {
    console.error('Completeness error:', err);
    container.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load completeness data.</p>';
  }
}

// ============================================================
//  ENHANCED ANALYTICS – SERVICE COVERAGE
// ============================================================

async function renderServiceCoverage() {
  const container = document.getElementById('service-coverage-container');
  if (!container) return;
  container.innerHTML = '<p class="text-muted" style="padding:16px;">Loading service coverage…</p>';

  try {
    const res = await API.getServiceCoverage();
    const rows = (res && res.data) || [];

    if (rows.length === 0) {
      container.innerHTML = '<p class="text-muted" style="padding:16px;">No data available.</p>';
      return;
    }

    const categories = [...new Set(rows.map(r => r.CategoryName))].sort();
    const properties = [...new Set(rows.map(r => r.ProjectName))].sort();

    const lookup = {};
    rows.forEach(r => {
      lookup[r.CategoryName + '||' + r.ProjectName] = r;
    });

    const catStats = {};
    categories.forEach(cat => {
      let covered = 0, total = properties.length, costs = [];
      properties.forEach(prop => {
        const d = lookup[cat + '||' + prop];
        if (d && parseInt(d.ContractCount) > 0) {
          covered++;
          const spend = parseFloat(d.MonthlySpend) || 0;
          if (spend > 0) costs.push(spend);
        }
      });
      catStats[cat] = { covered, total, costs };
    });

    let totalCovered = 0, totalCells = 0;
    Object.values(catStats).forEach(s => { totalCovered += s.covered; totalCells += s.total; });

    const stats = `<div class="coverage-stats">
      <div class="coverage-stat-card">
        <div class="stat-value">${Math.round((totalCovered / Math.max(totalCells, 1)) * 100)}%</div>
        <div class="stat-label">Overall Coverage</div>
      </div>
      <div class="coverage-stat-card">
        <div class="stat-value">${totalCells - totalCovered}</div>
        <div class="stat-label">Gaps</div>
      </div>
      <div class="coverage-stat-card">
        <div class="stat-value">${categories.length}</div>
        <div class="stat-label">Service Types</div>
      </div>
      <div class="coverage-stat-card">
        <div class="stat-value">${properties.length}</div>
        <div class="stat-label">Properties</div>
      </div>
    </div>`;

    let html = '<div class="coverage-wrap"><table class="coverage-table"><thead><tr><th>Service</th>';
    properties.forEach(p => {
      const short = p.length > 14 ? p.substring(0, 12) + '…' : p;
      html += `<th title="${esc(p)}">${esc(short)}</th>`;
    });
    html += '<th>Coverage</th></tr></thead><tbody>';

    categories.forEach(cat => {
      const s = catStats[cat];
      html += `<tr><td>${esc(cat)}</td>`;
      properties.forEach(prop => {
        const d = lookup[cat + '||' + prop];
        const count = d ? parseInt(d.ContractCount) || 0 : 0;
        const spend = d ? parseFloat(d.MonthlySpend) || 0 : 0;

        if (count === 0) {
          html += '<td class="coverage-cell missing">—</td>';
        } else if (spend === 0) {
          html += '<td class="coverage-cell underfunded">No $</td>';
        } else {
          html += `<td class="coverage-cell covered">${formatCurrencyShort(spend)}</td>`;
        }
      });
      const pct = Math.round((s.covered / Math.max(s.total, 1)) * 100);
      html += `<td style="font-weight:600;">${pct}%</td></tr>`;
    });

    html += '</tbody></table></div>';
    container.innerHTML = stats + html;
  } catch (err) {
    console.error('Service Coverage error:', err);
    container.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load coverage data.</p>';
  }
}

// ============================================================
//  ENHANCED ANALYTICS – VENDOR DEPENDENCY
// ============================================================

async function renderVendorDependency() {
  const container = document.getElementById('vendor-dependency-container');
  if (!container) return;
  container.innerHTML = '<p class="text-muted" style="padding:16px;">Loading vendor dependency…</p>';

  try {
    const res = await API.getVendorDependency();
    const rows = (res && res.data) || [];

    if (rows.length === 0) {
      container.innerHTML = '<p class="text-muted" style="padding:16px;">No vendor data available.</p>';
      return;
    }

    const totalProperties = [...new Set(rows.flatMap(r => (r.Properties || '').split(', ').filter(Boolean)))].length || 1;
    const totalMonthly = rows.reduce((s, r) => s + (parseFloat(r.TotalMonthly) || 0), 0);

    const cards = rows.slice(0, 20).map(row => {
      const propCount = parseInt(row.PropertyCount) || 0;
      const monthly = parseFloat(row.TotalMonthly) || 0;
      const annual = parseFloat(row.TotalAnnual) || 0;
      const contractCount = parseInt(row.ContractCount) || 0;
      const concentrationPct = Math.round((propCount / totalProperties) * 100);
      const spendPct = totalMonthly > 0 ? Math.round((monthly / totalMonthly) * 100) : 0;

      let riskLevel, riskLabel;
      if (concentrationPct >= 70) { riskLevel = 'risk-high'; riskLabel = 'High Risk'; }
      else if (concentrationPct >= 40) { riskLevel = 'risk-med'; riskLabel = 'Medium'; }
      else { riskLevel = 'risk-low'; riskLabel = 'Low'; }

      return `<div class="vendor-dep-card">
        <div class="vendor-dep-card-header">
          <span class="vendor-dep-name">${esc(row.VendorName)}</span>
          <span class="vendor-dep-badge ${riskLevel}">${riskLabel}</span>
        </div>
        <div class="vendor-dep-stats">
          <span class="vendor-dep-stat-label">Properties</span>
          <span class="vendor-dep-stat-value">${propCount} / ${totalProperties}</span>
          <span class="vendor-dep-stat-label">Contracts</span>
          <span class="vendor-dep-stat-value">${contractCount}</span>
          <span class="vendor-dep-stat-label">Monthly</span>
          <span class="vendor-dep-stat-value">${formatCurrencyShort(monthly)}</span>
          <span class="vendor-dep-stat-label">Annual</span>
          <span class="vendor-dep-stat-value">${formatCurrencyShort(annual)}</span>
        </div>
        <div class="vendor-dep-concentration-bar">
          <div class="vendor-dep-concentration-fill" style="width:${concentrationPct}%;background:${concentrationPct >= 70 ? 'var(--error)' : concentrationPct >= 40 ? 'var(--warning)' : 'var(--success)'}"></div>
        </div>
        <div class="vendor-dep-properties">${esc(row.Properties || '—')}</div>
      </div>`;
    }).join('');

    container.innerHTML = '<div class="vendor-dep-grid">' + cards + '</div>';
  } catch (err) {
    console.error('Vendor Dependency error:', err);
    container.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load vendor dependency data.</p>';
  }
}

// ============================================================
//  ENHANCED ANALYTICS – RENEWAL PIPELINE
// ============================================================

async function renderRenewalPipeline() {
  const detailsContainer = document.getElementById('renewal-pipeline-details');
  if (!detailsContainer) return;
  detailsContainer.innerHTML = '<p class="text-muted" style="padding:16px;">Loading renewal pipeline…</p>';

  try {
    const res = await API.getRenewalPipeline();
    const data = (res && res.data) || {};
    const pipeline = data.pipeline || [];
    const noExpiryCount = data.noExpiryCount || 0;

    if (pipeline.length === 0) {
      detailsContainer.innerHTML = `<p class="text-muted" style="padding:16px;">No upcoming renewals found.${noExpiryCount > 0 ? ' ' + noExpiryCount + ' contracts have no expiry date.' : ''}</p>`;
      return;
    }

    const totalAtRisk = pipeline.reduce((s, r) => s + (parseFloat(r.MonthlyAtRisk) || 0), 0);
    const totalContracts = pipeline.reduce((s, r) => s + (parseInt(r.ContractCount) || 0), 0);
    const totalNotice = pipeline.reduce((s, r) => s + (parseInt(r.NoticeDeadlineSoon) || 0), 0);

    const summary = `<div class="pipeline-summary">
      <div class="pipeline-summary-stat">
        <div class="stat-value">${totalContracts}</div>
        <div class="stat-label">Expiring</div>
      </div>
      <div class="pipeline-summary-stat">
        <div class="stat-value">${formatCurrencyShort(totalAtRisk * 12)}</div>
        <div class="stat-label">Annual at Risk</div>
      </div>
      <div class="pipeline-summary-stat">
        <div class="stat-value">${totalNotice}</div>
        <div class="stat-label">Notice Deadline Soon</div>
      </div>
      <div class="pipeline-summary-stat">
        <div class="stat-value">${noExpiryCount}</div>
        <div class="stat-label">No Expiry Date</div>
      </div>
    </div>`;

    const labels = pipeline.map(r => r.ExpiryMonth);
    const autoData = pipeline.map(r => parseInt(r.AutoRenewCount) || 0);
    const manualData = pipeline.map(r => parseInt(r.ManualRenewCount) || 0);

    createChart('chart-renewal-pipeline', {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label: 'Auto-Renew',
            data: autoData,
            backgroundColor: STOA_COLORS[0] + 'cc',
            borderColor: STOA_COLORS[0],
            borderWidth: 1,
            borderRadius: 4,
          },
          {
            label: 'Manual Renewal',
            data: manualData,
            backgroundColor: '#f59e0b' + 'cc',
            borderColor: '#f59e0b',
            borderWidth: 1,
            borderRadius: 4,
          }
        ]
      },
      options: {
        ...getChartDefaults(),
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { font: { size: 11 } } },
          y: { stacked: true, beginAtZero: true, grid: { color: 'rgba(0,0,0,0.06)' }, ticks: { stepSize: 1, font: { size: 11 } } }
        },
        plugins: {
          ...getChartDefaults().plugins,
          legend: { display: true, position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
        }
      }
    });

    let tableHtml = `<table class="pipeline-details-table">
      <thead><tr>
        <th>Month</th><th>Contracts</th><th>Monthly at Risk</th><th>Annual at Risk</th>
        <th>Auto / Manual</th><th>Notice Soon</th>
      </tr></thead><tbody>`;

    pipeline.forEach(r => {
      const auto = parseInt(r.AutoRenewCount) || 0;
      const manual = parseInt(r.ManualRenewCount) || 0;
      const total = auto + manual;
      const autoPct = total > 0 ? Math.round((auto / total) * 100) : 0;
      const noticeCount = parseInt(r.NoticeDeadlineSoon) || 0;

      tableHtml += `<tr>
        <td>${esc(r.ExpiryMonth)}</td>
        <td>${r.ContractCount}</td>
        <td>${formatCurrency(r.MonthlyAtRisk)}</td>
        <td>${formatCurrency(r.AnnualAtRisk)}</td>
        <td>
          <div class="pipeline-bar" title="${auto} auto, ${manual} manual" style="width:120px;">
            <div class="pipeline-bar-auto" style="width:${autoPct}%"></div>
            <div class="pipeline-bar-manual" style="width:${100 - autoPct}%"></div>
          </div>
        </td>
        <td>${noticeCount > 0 ? '<span class="pipeline-notice-flag">⚠ ' + noticeCount + '</span>' : '—'}</td>
      </tr>`;
    });

    tableHtml += '</tbody></table>';
    detailsContainer.innerHTML = summary + tableHtml;
  } catch (err) {
    console.error('Renewal Pipeline error:', err);
    detailsContainer.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load renewal pipeline.</p>';
  }
}

// ============================================================
//  ENHANCED ANALYTICS – COST BENCHMARKS
// ============================================================

async function renderCostBenchmarks() {
  const container = document.getElementById('cost-benchmarks-container');
  if (!container) return;
  container.innerHTML = '<p class="text-muted" style="padding:16px;">Loading cost benchmarks…</p>';

  try {
    const res = await API.getCostBenchmarks();
    const rows = (res && res.data) || [];

    if (rows.length === 0) {
      container.innerHTML = '<p class="text-muted" style="padding:16px;">No benchmark data available.</p>';
      return;
    }

    const byCategory = {};
    rows.forEach(r => {
      const cat = r.CategoryName || 'Uncategorized';
      if (!byCategory[cat]) byCategory[cat] = [];
      byCategory[cat].push(r);
    });

    let html = '<div class="benchmarks-wrap">';

    Object.keys(byCategory).sort().forEach(cat => {
      const contracts = byCategory[cat];
      const cpuValues = contracts
        .map(c => parseFloat(c.CostPerUnit))
        .filter(v => Number.isFinite(v) && v > 0);

      const avg = cpuValues.length > 0 ? cpuValues.reduce((a, b) => a + b, 0) / cpuValues.length : 0;
      const min = cpuValues.length > 0 ? Math.min(...cpuValues) : 0;
      const max = cpuValues.length > 0 ? Math.max(...cpuValues) : 0;
      const median = cpuValues.length > 0 ? (() => { const s = [...cpuValues].sort((a, b) => a - b); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; })() : 0;
      const outlierThreshold = avg * 2;

      html += `<div class="benchmark-category">
        <div class="benchmark-category-header">
          <span class="benchmark-category-name">${esc(cat)}</span>
          <span class="benchmark-category-avg">${contracts.length} contracts · Avg: ${formatCurrency(avg)}/unit</span>
        </div>
        <div class="benchmark-category-body">
          <table class="benchmark-table">
            <thead><tr>
              <th>Property</th><th>Vendor</th><th>Monthly</th><th>Units</th><th>Cost/Unit</th><th>vs. Avg</th>
            </tr></thead><tbody>`;

      contracts
        .sort((a, b) => ((parseFloat(b.CostPerUnit) || 0) - (parseFloat(a.CostPerUnit) || 0)))
        .forEach(c => {
          const cpu = parseFloat(c.CostPerUnit);
          const hasCPU = Number.isFinite(cpu) && cpu > 0;
          const isOutlier = hasCPU && cpu > outlierThreshold;
          const vsAvg = hasCPU && avg > 0 ? Math.round(((cpu - avg) / avg) * 100) : null;
          const barWidth = hasCPU && max > 0 ? Math.round((cpu / max) * 100) : 0;

          html += `<tr class="${isOutlier ? 'benchmark-outlier' : ''}">
            <td>${esc(c.ProjectName || '—')}</td>
            <td>${esc(c.VendorName || '—')}</td>
            <td>${c.MonthlyCost ? formatCurrency(c.MonthlyCost) : '—'}</td>
            <td>${c.Units || '—'}</td>
            <td>
              <div class="benchmark-bar-container">
                <div class="benchmark-bar-track">
                  <div class="benchmark-bar-fill ${isOutlier ? 'outlier' : ''}" style="width:${barWidth}%"></div>
                </div>
                <span>${hasCPU ? formatCurrency(cpu) : '—'}</span>
              </div>
            </td>
            <td>${vsAvg !== null ? (vsAvg >= 0 ? '+' : '') + vsAvg + '%' : '—'}</td>
          </tr>`;
        });

      html += `</tbody></table>
        <div class="benchmark-stats-row">
          <span>Min: <strong>${formatCurrency(min)}</strong></span>
          <span>Avg: <strong>${formatCurrency(avg)}</strong></span>
          <span>Median: <strong>${formatCurrency(median)}</strong></span>
          <span>Max: <strong>${formatCurrency(max)}</strong></span>
        </div>
      </div></div>`;
    });

    html += '</div>';
    container.innerHTML = html;
  } catch (err) {
    console.error('Cost Benchmarks error:', err);
    container.innerHTML = '<p class="text-muted" style="padding:16px;">Failed to load benchmark data.</p>';
  }
}

// ============================================================
//  SORTING
// ============================================================

function sortData(data, field, direction) {
  if (!data || data.length === 0) return data;
  const dir = direction === 'desc' ? -1 : 1;
  return [...data].sort((a, b) => {
    let va = a[field] ?? a['_' + field] ?? '';
    let vb = b[field] ?? b['_' + field] ?? '';

    if (typeof va === 'string' && typeof vb === 'string') {
      return va.localeCompare(vb, undefined, { sensitivity: 'base' }) * dir;
    }

    const na = parseFloat(va);
    const nb = parseFloat(vb);
    if (Number.isFinite(na) && Number.isFinite(nb)) return (na - nb) * dir;

    if (va < vb) return -1 * dir;
    if (va > vb) return 1 * dir;
    return 0;
  });
}

function handleSortClick(tableId, key) {
  const current = sortState[tableId];
  if (current && current.key === key) {
    sortState[tableId] = { key, dir: current.dir === 'asc' ? 'desc' : 'asc' };
  } else {
    sortState[tableId] = { key, dir: 'asc' };
  }
  renderCurrentTab();
}

function updateSortIndicators(tableId, activeKey, dir) {
  const table = document.getElementById(tableId);
  if (!table) return;
  table.querySelectorAll('.th-sort').forEach(th => {
    const key = th.dataset.key;
    th.classList.remove('sort-asc', 'sort-desc');
    if (key === activeKey) {
      th.classList.add(dir === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  });
}

// ============================================================
//  SEARCH & FILTERS
// ============================================================

function handleGlobalSearch(query) {
  activeFilters.search = (query || '').trim();
  renderCurrentTab();
}

function initStatusMulti() {
  const multi = $('#statusMulti');
  if (!multi) return;

  const statuses = ['Active', 'Pending', 'Expired', 'Terminated', 'Under Review'];
  const listEl = multi.querySelector('.multi-list');
  if (!listEl) return;

  listEl.innerHTML = statuses.map(s => `
    <label class="multi-item">
      <input type="checkbox" value="${s.toLowerCase()}" checked />
      <span>${escapeHtml(s)}</span>
    </label>
  `).join('');

  const toggleBtn = multi.querySelector('.multi-btn');
  const menu = multi.querySelector('.multi-menu');

  if (toggleBtn && menu) {
    toggleBtn.addEventListener('click', e => {
      e.stopPropagation();
      const isOpen = menu.getAttribute('aria-hidden') === 'false';
      menu.setAttribute('aria-hidden', isOpen ? 'true' : 'false');
      toggleBtn.setAttribute('aria-expanded', isOpen ? 'false' : 'true');
    });
  }

  multi.querySelectorAll('[data-act]').forEach(btn => {
    btn.addEventListener('click', () => {
      const act = btn.dataset.act;
      listEl.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = act === 'all';
      });
      applyStatusFilter();
    });
  });

  listEl.addEventListener('change', () => applyStatusFilter());

  document.addEventListener('click', e => {
    if (!multi.contains(e.target) && menu) {
      menu.setAttribute('aria-hidden', 'true');
      if (toggleBtn) toggleBtn.setAttribute('aria-expanded', 'false');
    }
  });
}

function applyStatusFilter() {
  const multi = $('#statusMulti');
  if (!multi) return;
  const checked = Array.from(multi.querySelectorAll('.multi-list input:checked')).map(cb => cb.value);
  activeFilters.status = checked;

  const label = multi.querySelector('.multi-label');
  if (label) {
    const total = multi.querySelectorAll('.multi-list input[type="checkbox"]').length;
    if (checked.length === total) label.textContent = 'All Statuses';
    else if (checked.length === 0) label.textContent = 'None Selected';
    else label.textContent = checked.length + ' selected';
  }

  if (currentTab === 'property') renderPropertyView();
}

// ============================================================
//  CONTRACT MODAL (ADD / EDIT)
// ============================================================

function openContractModal(contractId) {
  if (!canEdit()) { showToast('Login and enable Edit Mode to make changes.', 'warning'); return; }
  editingContractId = contractId || null;
  pendingFiles = [];

  const modal = $('#contract-modal');
  if (!modal) return;

  const title = $('#contractModalTitle');
  if (title) title.textContent = contractId ? 'Edit Contract' : 'Add Contract';

  const form = $('#contractForm');
  if (form) form.reset();

  populateContractDropdowns();

  const attachmentsEl = $('#contract-attachments');
  if (attachmentsEl) attachmentsEl.innerHTML = '';

  if (contractId) {
    const c = allContracts.find(x => (x.ContractId || x.Id) === contractId);
    if (c) {
      setFormVal('#contractProperty', c.ProjectId);
      setFormVal('#contractVendor', c.VendorId);
      setFormVal('#contractCategory', c.CategoryId);
      setFormVal('#contractDescription', c.Description || '');
      setFormVal('#contractType', c.ContractType || '');
      setFormVal('#contractStatus', c.Status || '');
      setFormVal('#contractMonthlyCost', c.MonthlyCost || '');
      setFormVal('#contractAnnualCost', c.AnnualCost || '');
      setFormVal('#contractBillingFrequency', c.BillingFrequency || '');
      setFormVal('#contractTerm', c.ContractTermMonths || c.TermMonths || '');
      setFormVal('#contractStartDate', formatDateISO(c.StartDate));
      setFormVal('#contractExpirationDate', formatDateISO(c.ExpirationDate));
      setFormVal('#contractNoticePeriod', c.NoticePeriodDays || c.NoticePeriod || '');
      setFormChecked('#contractAutoRenew', isAutoRenew(c));
      setFormChecked('#contractTerminationFee', c.TerminationFee === true || c.TerminationFee === 1);
      setFormVal('#contractTerminationFeeAmount', c.TerminationFeeAmount || '');
      setFormVal('#contractResponsiblePerson', c.ResponsiblePersonId || '');
      setFormVal('#contractAssignmentLanguage', c.AssignmentLanguage || '');
      setFormVal('#contractNotes', c.Notes || '');
      setFormVal('#contractContractNumber', c.ContractNumber || '');
      setFormVal('#contractSignedBy', c.SignedBy || '');
      setFormVal('#contractSignedDate', formatDateISO(c.SignedDate));
      setFormVal('#contractSetupFee', c.OneTimeSetupFee || '');
      setFormVal('#contractEscalation', c.AnnualEscalation || '');
      setFormVal('#contractEarlyTermTerms', c.EarlyTerminationTerms || '');
      setFormVal('#contractServiceFrequency', c.ServiceFrequency || '');
      setFormChecked('#contractInsuranceRequired', c.InsuranceRequired === true || c.InsuranceRequired === 1);
      setFormVal('#contractPaymentMethod', c.PaymentMethod || '');
      setFormVal('#contractPerUnitCost', c.PerUnitCost || '');
      setFormVal('#contractScope', c.ContractScope || '');
      setFormChecked('#contractNonCancellable', c.IsNonCancellable === true || c.IsNonCancellable === 1);
      setFormVal('#contractRenewalTermType', c.RenewalTermType || '');
      setFormVal('#contractAccountRep', c.AccountRepresentative || '');
      setFormVal('#contractReminderDays', c.ReminderDaysBefore || 60);

      if (attachmentsEl) loadContractAttachments(contractId, attachmentsEl, false);
    }
  }

  modal.classList.add('active');
  setTimeout(() => { const first = modal.querySelector('select, input'); if (first) first.focus(); }, 100);
}

function populateContractDropdowns() {
  populateSelect('#contractProperty', allProjects, p => p.ProjectId || p.Id, p => p.ProjectName || p.Name || 'Unknown', '-- Select Property --');
  populateSelect('#contractCategory', allCategories, c => c.CategoryId || c.Id, c => c.CategoryName || c.Name || 'Unknown', '-- Select Category --');
  populateSelect('#contractResponsiblePerson', allPersons, p => p.PersonId || p.Id, p => p.FullName || p.Name || 'Unknown', '-- Select --');

  const vendorSelect = $('#contractVendor');
  if (vendorSelect) {
    const current = vendorSelect.value;
    vendorSelect.innerHTML = '<option value="">-- Select Vendor --</option><option value="__new__">+ Add New Vendor</option>';
    allVendors.forEach(v => {
      const id = v.VendorId || v.Id;
      const name = v.VendorName || v.Name || 'Unknown';
      vendorSelect.innerHTML += `<option value="${id}">${escapeHtml(name)}</option>`;
    });
    if (current) vendorSelect.value = current;
  }
}

function populateSelect(selector, items, getVal, getLabel, placeholder) {
  const el = $(selector);
  if (!el) return;
  const current = el.value;
  el.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>`;
  items.forEach(item => {
    const val = getVal(item);
    const label = getLabel(item);
    el.innerHTML += `<option value="${val}">${escapeHtml(label)}</option>`;
  });
  if (current) el.value = current;
}

function setFormVal(selector, value) {
  const el = $(selector);
  if (el) el.value = value ?? '';
}

function setFormChecked(selector, checked) {
  const el = $(selector);
  if (el) el.checked = !!checked;
}

async function saveContract(e) {
  if (e) e.preventDefault();
  if (!canEdit()) { showToast('Login and enable Edit Mode to make changes.', 'warning'); return; }

  const propId = ($('#contractProperty') || {}).value;
  const vendorVal = ($('#contractVendor') || {}).value;
  const catId = ($('#contractCategory') || {}).value;
  const status = ($('#contractStatus') || {}).value;

  if (!propId) {
    showToast('Please select a property.', 'warning');
    return;
  }
  if (!catId) {
    showToast('Please select a category.', 'warning');
    return;
  }
  if (!status) {
    showToast('Please select a status.', 'warning');
    return;
  }

  if (vendorVal === '__new__') {
    openVendorModal();
    showToast('Please create the vendor first, then re-save the contract.', 'info');
    return;
  }

  const saveBtn = $('#saveContractBtn');
  if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'Saving…'; }

  const data = {
    ProjectId: parseInt(propId, 10),
    VendorId: vendorVal ? parseInt(vendorVal, 10) : null,
    ContractCategoryId: parseInt(catId, 10),
    Description: ($('#contractDescription') || {}).value || null,
    ContractType: ($('#contractType') || {}).value || null,
    Status: status,
    MonthlyCost: parseFloat(($('#contractMonthlyCost') || {}).value) || null,
    AnnualCost: parseFloat(($('#contractAnnualCost') || {}).value) || null,
    BillingFrequency: ($('#contractBillingFrequency') || {}).value || null,
    ContractTermMonths: parseInt(($('#contractTerm') || {}).value, 10) || null,
    StartDate: ($('#contractStartDate') || {}).value || null,
    ExpirationDate: ($('#contractExpirationDate') || {}).value || null,
    NoticePeriodDays: parseInt(($('#contractNoticePeriod') || {}).value, 10) || null,
    AutoRenew: !!($('#contractAutoRenew') || {}).checked,
    TerminationFee: !!($('#contractTerminationFee') || {}).checked,
    TerminationFeeAmount: parseFloat(($('#contractTerminationFeeAmount') || {}).value) || null,
    ResponsiblePersonId: parseInt(($('#contractResponsiblePerson') || {}).value, 10) || null,
    AssignmentLanguage: ($('#contractAssignmentLanguage') || {}).value || null,
    Notes: ($('#contractNotes') || {}).value || null,
    ContractNumber: ($('#contractContractNumber') || {}).value || null,
    SignedBy: ($('#contractSignedBy') || {}).value || null,
    SignedDate: ($('#contractSignedDate') || {}).value || null,
    OneTimeSetupFee: parseFloat(($('#contractSetupFee') || {}).value) || null,
    AnnualEscalation: ($('#contractEscalation') || {}).value || null,
    EarlyTerminationTerms: ($('#contractEarlyTermTerms') || {}).value || null,
    ServiceFrequency: ($('#contractServiceFrequency') || {}).value || null,
    InsuranceRequired: !!($('#contractInsuranceRequired') || {}).checked,
    PaymentMethod: ($('#contractPaymentMethod') || {}).value || null,
    PerUnitCost: parseFloat(($('#contractPerUnitCost') || {}).value) || null,
    ContractScope: ($('#contractScope') || {}).value || null,
    IsNonCancellable: !!($('#contractNonCancellable') || {}).checked,
    RenewalTermType: ($('#contractRenewalTermType') || {}).value || null,
    AccountRepresentative: ($('#contractAccountRep') || {}).value || null,
    ReminderDaysBefore: parseInt(($('#contractReminderDays') || {}).value, 10) || 60
  };

  try {
    let result;
    if (editingContractId) {
      result = await API.updateContract(editingContractId, data);
    } else {
      result = await API.createContract(data);
    }

    if (result.success) {
      const savedId = editingContractId || (result.data && (result.data.ContractId || result.data.Id));

      if (savedId && pendingFiles.length > 0) {
        let uploadOk = 0;
        let uploadFail = 0;
        for (const file of pendingFiles) {
          try {
            await API.uploadContractAttachment(savedId, file);
            uploadOk++;
          } catch (uploadErr) {
            uploadFail++;
            console.warn('File upload failed:', file.name, uploadErr);
          }
        }
        if (uploadFail > 0) {
          showToast(`${uploadOk} file(s) uploaded, ${uploadFail} failed.`, 'warning');
        }
      }

      showToast(
        editingContractId ? 'Contract updated successfully.' : 'Contract created successfully.',
        'success'
      );
      closeModalAnimated($('#contract-modal'));
      await refreshData();
    } else {
      showToast(result.error?.message || 'Failed to save contract.', 'error');
    }
  } catch (err) {
    showToast('Error saving contract: ' + (err.message || 'Unknown error'), 'error');
  } finally {
    if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = 'Save Contract'; }
  }
}

// ============================================================
//  RENEW MODAL
// ============================================================

function openRenewModal(contractId) {
  if (!canEdit()) { showToast('Login and enable Edit Mode to renew contracts.', 'warning'); return; }
  renewingContractId = contractId;
  const modal = $('#renew-modal');
  if (!modal) return;

  const c = allContracts.find(x => (x.ContractId || x.Id) === contractId);
  if (!c) {
    showToast('Contract not found.', 'error');
    return;
  }

  const infoEl = $('#renewCurrentInfo');
  if (infoEl) {
    infoEl.innerHTML = `
      <div class="renew-info-grid">
        <div><strong>Property:</strong> ${escapeHtml(projectName(c.ProjectId))}</div>
        <div><strong>Vendor:</strong> ${escapeHtml(vendorName(c.VendorId))}</div>
        <div><strong>Description:</strong> ${escapeHtml(c.Description || '—')}</div>
        <div><strong>Current Expiration:</strong> ${formatDate(c.ExpirationDate)}</div>
        <div><strong>Current Monthly Cost:</strong> ${formatCurrency(c.MonthlyCost)}</div>
        <div><strong>Current Annual Cost:</strong> ${formatCurrency(c.AnnualCost)}</div>
      </div>
    `;
  }

  setFormVal('#renewMonthlyCost', c.MonthlyCost || '');
  setFormVal('#renewAnnualCost', c.AnnualCost || '');
  setFormVal('#renewExpirationDate', '');
  setFormVal('#renewNotes', '');

  modal.classList.add('active');
  setTimeout(() => { const inp = $('#renewExpirationDate'); if (inp) inp.focus(); }, 100);
}

async function saveRenewal(e) {
  if (e) e.preventDefault();
  if (!canEdit()) { showToast('Login and enable Edit Mode to make changes.', 'warning'); return; }

  const newExpiration = ($('#renewExpirationDate') || {}).value;
  if (!newExpiration) {
    showToast('New expiration date is required.', 'warning');
    return;
  }

  const saveBtn = $('#saveRenewBtn');
  if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'Saving…'; }

  const data = {
    NewExpirationDate: newExpiration,
    NewMonthlyCost: parseFloat(($('#renewMonthlyCost') || {}).value) || undefined,
    NewAnnualCost: parseFloat(($('#renewAnnualCost') || {}).value) || undefined,
    Notes: ($('#renewNotes') || {}).value || undefined
  };

  try {
    const result = await API.renewContract(renewingContractId, data);
    if (result.success) {
      showToast('Contract renewed successfully.', 'success');
      closeModalAnimated($('#renew-modal'));
      await refreshData();
    } else {
      showToast(result.error?.message || 'Failed to renew contract.', 'error');
    }
  } catch (err) {
    showToast('Error renewing contract: ' + (err.message || 'Unknown error'), 'error');
  } finally {
    if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = 'Save Renewal'; }
  }
}

// ============================================================
//  DETAIL MODAL
// ============================================================

async function openDetailModal(contractId) {
  detailContractId = contractId;
  const modal = $('#detail-modal');
  if (!modal) return;

  const c = allContracts.find(x => (x.ContractId || x.Id) === contractId);
  if (!c) {
    showToast('Contract not found.', 'error');
    return;
  }

  const title = $('#detailModalTitle');
  if (title) title.textContent = c.Description || vendorName(c.VendorId) || 'Contract Details';

  renderDetailFields(c);

  const attachmentsEl = $('#detail-attachments');
  if (attachmentsEl) {
    attachmentsEl.innerHTML = '<div class="loading-inline">Loading attachments…</div>';
    loadContractAttachments(contractId, attachmentsEl, true);
  }

  const historyEl = $('#detail-history');
  if (historyEl) {
    historyEl.innerHTML = '<div class="loading-inline">Loading history…</div>';
    loadContractHistory(contractId, historyEl);
  }

  modal.classList.add('active');
}

function renderDetailFields(c) {
  const fields = $('#detailFields');
  if (!fields) return;

  const days = daysUntil(c.ExpirationDate);
  const hasExpiry = days !== Infinity;

  const detailRows = [
    ['Property', escapeHtml(projectName(c.ProjectId))],
    ['Vendor', escapeHtml(vendorName(c.VendorId))],
    ['Category', escapeHtml(categoryName(c.CategoryId))],
    ['Status', getStatusBadge(c.Status)],
    ['Contract Type', escapeHtml(c.ContractType || '—')],
    ['Contract #', escapeHtml(c.ContractNumber || '—')],
    ['Responsible', escapeHtml(personName(c.ResponsiblePersonId))],
    ['Monthly Cost', formatCurrency(c.MonthlyCost)],
    ['Annual Cost', formatCurrency(c.AnnualCost)],
    ['Per-Unit Cost', c.PerUnitCost ? formatCurrency(c.PerUnitCost) + '/unit' : '—'],
    ['Billing Frequency', escapeHtml(c.BillingFrequency || '—')],
    ['Service Frequency', escapeHtml(c.ServiceFrequency || '—')],
    ['Term', (c.ContractTermMonths || c.TermMonths) ? (c.ContractTermMonths || c.TermMonths) + ' months' : '—'],
    ['Start Date', formatDate(c.StartDate)],
    [
      'Expiration Date',
      formatDate(c.ExpirationDate) +
      (hasExpiry ? ` <span class="days-badge ${getUrgencyClass(days)}">${days <= 0 ? 'EXPIRED' : days + ' days'}</span>` : '')
    ],
    ['Notice Period', (c.NoticePeriodDays || c.NoticePeriod) ? (c.NoticePeriodDays || c.NoticePeriod) + ' days' : '—'],
    ['Auto-Renew', getAutoRenewBadge(c.AutoRenew)],
    ['Renewal Type', escapeHtml(c.RenewalTermType || '—')],
    ['Non-Cancellable', (c.IsNonCancellable === true || c.IsNonCancellable === 1) ? '<span class="badge badge-danger">Non-Cancellable</span>' : 'No'],
    [
      'Termination Fee',
      (c.TerminationFee === true || c.TerminationFee === 1)
        ? (c.TerminationFeeAmount ? formatCurrency(c.TerminationFeeAmount) : 'Yes')
        : 'None'
    ],
    ['Setup / One-Time Fee', c.OneTimeSetupFee ? formatCurrency(c.OneTimeSetupFee) : '—'],
    ['Annual Escalation', escapeHtml(c.AnnualEscalation || '—')],
    ['Payment Method', escapeHtml(c.PaymentMethod || '—')],
    ['Signed By', escapeHtml(c.SignedBy || '—')],
    ['Signed Date', formatDate(c.SignedDate)],
    ['Account Rep', escapeHtml(c.AccountRepresentative || '—')],
  ];

  const fullWidthRows = [];
  if (c.Description) fullWidthRows.push(['Description', escapeHtml(c.Description)]);
  if (c.ContractScope) fullWidthRows.push(['Scope / Coverage', escapeHtml(c.ContractScope)]);
  if (c.EarlyTerminationTerms) fullWidthRows.push(['Early Termination Terms', escapeHtml(c.EarlyTerminationTerms)]);
  if (c.AssignmentLanguage) fullWidthRows.push(['Assignment Language', escapeHtml(c.AssignmentLanguage)]);
  if (c.Notes) fullWidthRows.push(['Notes', escapeHtml(c.Notes)]);

  fields.innerHTML = `
    <div class="detail-grid">
      ${detailRows.map(([label, value]) => `
        <div class="detail-field">
          <span class="detail-label">${label}</span>
          <span class="detail-value">${value}</span>
        </div>
      `).join('')}
      ${fullWidthRows.map(([label, value]) => `
        <div class="detail-field detail-field-full">
          <span class="detail-label">${label}</span>
          <span class="detail-value">${value}</span>
        </div>
      `).join('')}
    </div>
  `;
}

async function loadContractAttachments(contractId, container, showDeleteButton) {
  try {
    const result = await API.getContractAttachments(contractId);
    if (result.success && Array.isArray(result.data) && result.data.length > 0) {
      container.innerHTML = result.data.map(att => {
        const attId = att.AttachmentId || att.Id;
        const fileName = att.FileName || att.filename || att.OriginalName || 'File';
        const fileSize = att.FileSize ? ` (${formatFileSize(att.FileSize)})` : '';
        return `<div class="attachment-item" data-attachment-id="${attId}">
          <span class="attachment-name">${escapeHtml(fileName)}${fileSize}</span>
          <div class="attachment-actions">
            <button class="btn btn-xs btn-download-att" data-att-id="${attId}" title="Download">Download</button>
            ${showDeleteButton ? `<button class="btn btn-xs btn-danger btn-delete-att" data-att-id="${attId}" title="Delete">Delete</button>` : ''}
          </div>
        </div>`;
      }).join('');

      container.querySelectorAll('.btn-download-att').forEach(btn => {
        btn.addEventListener('click', async (e) => {
          e.stopPropagation();
          try {
            btn.disabled = true;
            btn.textContent = '…';
            const url = await API.downloadContractAttachment(parseInt(btn.dataset.attId, 10));
            const a = document.createElement('a');
            a.href = url;
            a.download = '';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
          } catch (err) {
            showToast('Download failed: ' + (err.message || 'Unknown error'), 'error');
          } finally {
            btn.disabled = false;
            btn.textContent = 'Download';
          }
        });
      });

      container.querySelectorAll('.btn-delete-att').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const attId = parseInt(btn.dataset.attId, 10);
          showConfirm('Delete this attachment? This cannot be undone.', async () => {
            try {
              const result = await API.deleteContractAttachment(attId);
              if (result.success) {
                const item = btn.closest('.attachment-item');
                if (item) item.remove();
                showToast('Attachment deleted.', 'success');
              } else {
                showToast('Failed to delete attachment.', 'error');
              }
            } catch (err) {
              showToast('Error: ' + (err.message || 'Unknown error'), 'error');
            }
          });
        });
      });
    } else {
      container.innerHTML = '<div class="empty-state-sm">No attachments.</div>';
    }
  } catch (err) {
    console.warn('Failed to load attachments:', err);
    container.innerHTML = '<div class="empty-state-sm">Could not load attachments.</div>';
  }
}

async function loadContractHistory(contractId, container) {
  try {
    const result = await API.getContractHistory(contractId);
    if (result.success && Array.isArray(result.data) && result.data.length > 0) {
      const items = [...result.data].sort((a, b) => {
        const da = new Date(b.ChangedAt || b.CreatedAt || 0).getTime();
        const db = new Date(a.ChangedAt || a.CreatedAt || 0).getTime();
        return da - db;
      });

      container.innerHTML = `<div class="history-list">
        ${items.map((h, idx) => {
          const changeType = h.ChangeType || h.Action || 'Change';
          const description = h.ChangeDescription || h.Description || h.Details || '';
          const dateStr = h.ChangedAt || h.CreatedAt;
          const changedBy = h.ChangedBy || h.User || '';
          const isFirst = idx === 0;

          return `<div class="history-item ${isFirst ? 'history-item-latest' : ''}">
            <div class="history-dot"></div>
            <div class="history-content">
              <div class="history-action">${escapeHtml(changeType)}</div>
              ${description ? `<div class="history-detail">${escapeHtml(description)}</div>` : ''}
              <div class="history-meta">
                ${formatDate(dateStr)}${changedBy ? ' &middot; ' + escapeHtml(changedBy) : ''}
              </div>
            </div>
          </div>`;
        }).join('')}
      </div>`;
    } else {
      container.innerHTML = '<div class="empty-state-sm">No history available.</div>';
    }
  } catch (err) {
    console.warn('Failed to load history:', err);
    container.innerHTML = '<div class="empty-state-sm">Could not load history.</div>';
  }
}

// ============================================================
//  VENDOR MODAL
// ============================================================

function openVendorModal(vendorId) {
  if (!canEdit()) { showToast('Login and enable Edit Mode to manage vendors.', 'warning'); return; }
  editingVendorId = vendorId || null;
  const modal = $('#vendor-modal');
  if (!modal) return;

  const title = $('#vendorModalTitle');
  if (title) title.textContent = vendorId ? 'Edit Vendor' : 'Add Vendor';

  const form = $('#vendorForm');
  if (form) form.reset();

  if (vendorId) {
    const v = allVendors.find(x => (x.VendorId || x.Id) === vendorId);
    if (v) {
      setFormVal('#vendorName', v.VendorName || v.Name || '');
      setFormVal('#vendorContactName', v.ContactName || '');
      setFormVal('#vendorContactEmail', v.ContactEmail || v.Email || '');
      setFormVal('#vendorContactPhone', v.ContactPhone || v.Phone || '');
      setFormVal('#vendorWebsite', v.Website || '');
      setFormVal('#vendorAddress', v.Address || '');
      setFormVal('#vendorNotes', v.Notes || '');
    }
  }

  modal.classList.add('active');
  setTimeout(() => { const inp = $('#vendorName'); if (inp) inp.focus(); }, 100);
}

async function saveVendor(e) {
  if (e) e.preventDefault();
  if (!canEdit()) { showToast('Login and enable Edit Mode to make changes.', 'warning'); return; }

  const name = ($('#vendorName') || {}).value;
  if (!name || !name.trim()) {
    showToast('Vendor name is required.', 'warning');
    return;
  }

  const saveBtn = $('#saveVendorBtn');
  if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'Saving…'; }

  const data = {
    VendorName: name.trim(),
    ContactName: ($('#vendorContactName') || {}).value || null,
    ContactEmail: ($('#vendorContactEmail') || {}).value || null,
    ContactPhone: ($('#vendorContactPhone') || {}).value || null,
    Website: ($('#vendorWebsite') || {}).value || null,
    Address: ($('#vendorAddress') || {}).value || null,
    Notes: ($('#vendorNotes') || {}).value || null
  };

  try {
    let result;
    if (editingVendorId) {
      result = await API.updateVendor(editingVendorId, data);
    } else {
      result = await API.createVendor(data);
    }

    if (result.success) {
      showToast(
        editingVendorId ? 'Vendor updated successfully.' : 'Vendor created successfully.',
        'success'
      );
      closeModalAnimated($('#vendor-modal'));
      await refreshData();
    } else {
      showToast(result.error?.message || 'Failed to save vendor.', 'error');
    }
  } catch (err) {
    showToast('Error saving vendor: ' + (err.message || 'Unknown error'), 'error');
  } finally {
    if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = 'Save Vendor'; }
  }
}

// ============================================================
//  CONFIRM MODAL
// ============================================================

function showConfirm(message, onConfirm) {
  const modal = $('#confirm-modal');
  if (!modal) return;

  const msgEl = $('#confirmMessage');
  if (msgEl) msgEl.textContent = message;

  confirmCallback = onConfirm;
  modal.classList.add('active');

  setTimeout(() => {
    const okBtn = $('#confirmOkBtn');
    if (okBtn) okBtn.focus();
  }, 100);
}

function handleConfirmOk() {
  closeModalAnimated($('#confirm-modal'));
  if (typeof confirmCallback === 'function') {
    const cb = confirmCallback;
    confirmCallback = null;
    cb();
  }
}

function handleConfirmCancel() {
  closeModalAnimated($('#confirm-modal'));
  confirmCallback = null;
}

// ============================================================
//  DELETE OPERATIONS
// ============================================================

function deleteContract(id) {
  if (!canEdit()) { showToast('Login and enable Edit Mode to delete.', 'warning'); return; }
  const c = allContracts.find(x => (x.ContractId || x.Id) === id);
  const desc = c ? (c.Description || vendorName(c.VendorId) || 'Contract #' + id) : 'this contract';
  showConfirm(`Delete "${desc}"? This action cannot be undone.`, async () => {
    try {
      const result = await API.deleteContract(id);
      if (result.success) {
        showToast('Contract deleted.', 'success');
        closeModalAnimated($('#detail-modal'));
        await refreshData();
      } else {
        showToast(result.error?.message || 'Failed to delete contract.', 'error');
      }
    } catch (err) {
      showToast('Error: ' + (err.message || 'Unknown error'), 'error');
    }
  });
}

function deleteVendor(id) {
  if (!canEdit()) { showToast('Login and enable Edit Mode to delete.', 'warning'); return; }
  const v = allVendors.find(x => (x.VendorId || x.Id) === id);
  const name = v ? (v.VendorName || v.Name || 'Vendor #' + id) : 'this vendor';
  const contractCount = allContracts.filter(c => String(c.VendorId) === String(id)).length;
  const msg = contractCount > 0
    ? `Delete vendor "${name}"? This vendor has ${contractCount} associated ${pluralize(contractCount, 'contract')} that will be unlinked.`
    : `Delete vendor "${name}"?`;

  showConfirm(msg, async () => {
    try {
      const result = await API.deleteVendor(id);
      if (result.success) {
        showToast('Vendor deleted.', 'success');
        await refreshData();
      } else {
        showToast(result.error?.message || 'Failed to delete vendor.', 'error');
      }
    } catch (err) {
      showToast('Error: ' + (err.message || 'Unknown error'), 'error');
    }
  });
}

// ============================================================
//  FILE UPLOAD HANDLING
// ============================================================

function initFileUpload() {
  const zone = $('#contractUploadZone');
  const input = $('#contractFileInput');
  if (!zone || !input) return;

  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    e.stopPropagation();
    zone.classList.add('dragover');
  });

  zone.addEventListener('dragenter', e => {
    e.preventDefault();
    e.stopPropagation();
    zone.classList.add('dragover');
  });

  zone.addEventListener('dragleave', e => {
    e.preventDefault();
    e.stopPropagation();
    if (!zone.contains(e.relatedTarget)) {
      zone.classList.remove('dragover');
    }
  });

  zone.addEventListener('drop', e => {
    e.preventDefault();
    e.stopPropagation();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      addPendingFiles(e.dataTransfer.files);
    }
  });

  input.addEventListener('change', () => {
    if (input.files && input.files.length > 0) {
      addPendingFiles(input.files);
      input.value = '';
    }
  });
}

function addPendingFiles(fileList) {
  const maxSize = 25 * 1024 * 1024;
  for (const file of fileList) {
    if (file.size > maxSize) {
      showToast(`"${file.name}" exceeds 25 MB limit.`, 'warning');
      continue;
    }
    if (pendingFiles.some(f => f.name === file.name && f.size === file.size)) {
      showToast(`"${file.name}" is already added.`, 'info');
      continue;
    }
    pendingFiles.push(file);
  }
  renderPendingFiles();
}

function renderPendingFiles() {
  const container = $('#contract-attachments');
  if (!container) return;

  const existingEls = container.querySelectorAll('.attachment-item[data-attachment-id]');
  const existingHtml = Array.from(existingEls).map(el => el.outerHTML).join('');

  const pendingHtml = pendingFiles.map((f, i) => `
    <div class="attachment-item pending-file" data-pending-index="${i}">
      <span class="attachment-name">
        <span class="pending-badge">NEW</span>
        ${escapeHtml(f.name)} <small class="subtle">(${formatFileSize(f.size)})</small>
      </span>
      <button class="btn btn-xs btn-danger btn-remove-pending" data-index="${i}" title="Remove">&times;</button>
    </div>
  `).join('');

  container.innerHTML = existingHtml + pendingHtml;

  container.querySelectorAll('.btn-remove-pending').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const idx = parseInt(btn.dataset.index, 10);
      pendingFiles.splice(idx, 1);
      renderPendingFiles();
    });
  });
}

// ============================================================
//  BULK OPERATIONS
// ============================================================

async function sendBulkReminders() {
  showConfirm('Send expiry reminder emails for all contracts expiring soon?', async () => {
    try {
      const result = await API.sendExpiryReminders();
      if (result.success) {
        const count = result.data?.sentCount || result.data?.count || '';
        showToast(
          count ? `${count} expiry ${pluralize(count, 'reminder')} sent.` : 'Expiry reminders sent successfully.',
          'success'
        );
      } else {
        showToast(result.error?.message || 'Failed to send reminders.', 'error');
      }
    } catch (err) {
      showToast('Error sending reminders: ' + (err.message || 'Unknown error'), 'error');
    }
  });
}

// ============================================================
//  EVENT LISTENERS
// ============================================================

function bindEventListeners() {
  $$('.main-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      if (tab) switchTab(tab);
    });
  });

  const loginBtn = $('#loginBtn');
  if (loginBtn) {
    loginBtn.addEventListener('click', () => {
      if (isAuthenticated) {
        handleLogout();
      } else {
        const modal = $('#login-modal');
        if (modal) modal.classList.add('active');
        const usernameField = $('#loginUsername');
        if (usernameField) setTimeout(() => usernameField.focus(), 200);
      }
    });
  }

  const loginForm = $('#loginForm');
  if (loginForm) loginForm.addEventListener('submit', handleLogin);

  const closeLoginBtn = $('#closeLoginBtn');
  if (closeLoginBtn) closeLoginBtn.addEventListener('click', () => closeModalAnimated($('#login-modal')));

  const cancelLoginBtn = $('#cancelLoginBtn');
  if (cancelLoginBtn) cancelLoginBtn.addEventListener('click', () => closeModalAnimated($('#login-modal')));

  const editModeBtn = $('#editModeBtn');
  if (editModeBtn) editModeBtn.addEventListener('click', toggleEditMode);

  const addBtn = $('#addContractBtn');
  if (addBtn) addBtn.addEventListener('click', () => openContractModal());

  const refreshBtn = $('#refreshBtn');
  if (refreshBtn) refreshBtn.addEventListener('click', () => refreshData());

  const searchInput = $('#q');
  if (searchInput) {
    const debouncedSearch = debounce(e => handleGlobalSearch(e.target.value), 250);
    searchInput.addEventListener('input', debouncedSearch);
    searchInput.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        searchInput.value = '';
        handleGlobalSearch('');
      }
    });
  }

  const mobileSearchBtn = $('#mobileSearchBtn');
  const mobileSearchBar = $('#mobileSearchBar');
  const mobileQ = $('#mobileQ');
  if (mobileSearchBtn && mobileSearchBar && mobileQ) {
    mobileSearchBtn.addEventListener('click', () => {
      mobileSearchBar.classList.toggle('active');
      if (mobileSearchBar.classList.contains('active')) {
        mobileQ.focus();
      }
    });
    const debouncedMobileSearch = debounce(e => {
      handleGlobalSearch(e.target.value);
      if (searchInput) searchInput.value = e.target.value;
    }, 250);
    mobileQ.addEventListener('input', debouncedMobileSearch);
    mobileQ.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        mobileQ.value = '';
        handleGlobalSearch('');
        mobileSearchBar.classList.remove('active');
      }
    });
  }

  const propSearch = $('#propertySearch');
  if (propSearch) {
    propSearch.addEventListener('input', debounce(() => renderPropertyView(), 250));
  }

  const vendorSearch = $('#vendorSearch');
  if (vendorSearch) {
    vendorSearch.addEventListener('input', debounce(() => renderVendorView(), 250));
  }

  const addVendorBtn = $('#addVendorBtn');
  if (addVendorBtn) addVendorBtn.addEventListener('click', () => openVendorModal());

  const contractForm = $('#contractForm');
  if (contractForm) contractForm.addEventListener('submit', saveContract);

  const renewForm = $('#renewForm');
  if (renewForm) renewForm.addEventListener('submit', saveRenewal);

  const vendorForm = $('#vendorForm');
  if (vendorForm) vendorForm.addEventListener('submit', saveVendor);

  const cancelContractBtn = $('#cancelContractBtn');
  if (cancelContractBtn) cancelContractBtn.addEventListener('click', () => closeModalAnimated($('#contract-modal')));

  const cancelRenewBtn = $('#cancelRenewBtn');
  if (cancelRenewBtn) cancelRenewBtn.addEventListener('click', () => closeModalAnimated($('#renew-modal')));

  const cancelVendorBtn = $('#cancelVendorBtn');
  if (cancelVendorBtn) cancelVendorBtn.addEventListener('click', () => closeModalAnimated($('#vendor-modal')));

  const confirmOkBtn = $('#confirmOkBtn');
  if (confirmOkBtn) confirmOkBtn.addEventListener('click', handleConfirmOk);

  const confirmCancelBtn = $('#confirmCancelBtn');
  if (confirmCancelBtn) confirmCancelBtn.addEventListener('click', handleConfirmCancel);

  const detailCloseBtn = $('#detailCloseBtn');
  if (detailCloseBtn) detailCloseBtn.addEventListener('click', () => closeModalAnimated($('#detail-modal')));

  const detailEditBtn = $('#detailEditBtn');
  if (detailEditBtn) {
    detailEditBtn.addEventListener('click', () => {
      closeModalAnimated($('#detail-modal'));
      if (detailContractId) setTimeout(() => openContractModal(detailContractId), 220);
    });
  }

  const detailRenewBtn = $('#detailRenewBtn');
  if (detailRenewBtn) {
    detailRenewBtn.addEventListener('click', () => {
      closeModalAnimated($('#detail-modal'));
      if (detailContractId) setTimeout(() => openRenewModal(detailContractId), 220);
    });
  }

  $$('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('mousedown', e => {
      if (e.target === overlay) closeModalAnimated(overlay);
    });
  });

  $$('.modal-close').forEach(btn => {
    btn.addEventListener('click', () => {
      const overlay = btn.closest('.modal-overlay');
      if (overlay) closeModalAnimated(overlay);
    });
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      const openModals = $$('.modal-overlay').filter(m =>
        m.classList.contains('active')
      );
      if (openModals.length > 0) {
        closeModalAnimated(openModals[openModals.length - 1]);
        e.preventDefault();
      }
    }
  });

  $$('.th-sort').forEach(th => {
    th.style.cursor = 'pointer';
    th.addEventListener('click', () => {
      const table = th.closest('table');
      if (!table || !table.id) return;
      const key = th.dataset.key;
      if (key) handleSortClick(table.id, key);
    });
  });

  const expiryTimeframe = $('#expiryTimeframe');
  if (expiryTimeframe) expiryTimeframe.addEventListener('change', () => renderExpiryView());

  const expiryAutoRenewFilter = $('#expiryAutoRenewFilter');
  if (expiryAutoRenewFilter) expiryAutoRenewFilter.addEventListener('change', () => renderExpiryView());

  const bulkReminderBtn = $('#bulkReminderBtn');
  if (bulkReminderBtn) bulkReminderBtn.addEventListener('click', sendBulkReminders);

  $$('.view-switch').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.view;
      if (view) showAnalyticsSubView(view);
    });
  });

  const contractVendorSelect = $('#contractVendor');
  if (contractVendorSelect) {
    contractVendorSelect.addEventListener('change', () => {
      if (contractVendorSelect.value === '__new__') {
        openVendorModal();
      }
    });
  }

  const monthlyCostInput = $('#contractMonthlyCost');
  const annualCostInput = $('#contractAnnualCost');
  if (monthlyCostInput && annualCostInput) {
    monthlyCostInput.addEventListener('input', () => {
      const monthly = parseFloat(monthlyCostInput.value);
      if (Number.isFinite(monthly) && (!annualCostInput.value || annualCostInput.dataset.autoCalc === 'true')) {
        annualCostInput.value = (monthly * 12).toFixed(2);
        annualCostInput.dataset.autoCalc = 'true';
      }
    });
    annualCostInput.addEventListener('input', () => {
      annualCostInput.dataset.autoCalc = 'false';
    });
  }

  initFileUpload();
  initStatusMulti();

  const resizeHandler = debounce(() => {
    const isMobile = window.innerWidth <= 768;
    const wasMobile = document.documentElement.getAttribute('data-mobile') === 'true';
    if (isMobile !== wasMobile) {
      document.documentElement.setAttribute('data-mobile', isMobile ? 'true' : 'false');
      window.IS_MOBILE = isMobile;
      renderedTabs.clear();
      renderCurrentTab();
    }

    Object.values(chartInstances).forEach(chart => {
      if (chart && typeof chart.resize === 'function') {
        try { chart.resize(); } catch(e) {}
      }
    });
  }, 250);
  window.addEventListener('resize', resizeHandler);

  window.addEventListener('orientationchange', () => {
    setTimeout(() => {
      resizeHandler();
      Object.values(chartInstances).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
          try { chart.resize(); } catch(e) {}
        }
      });
    }, 300);
  });
}

// ============================================================
//  GLOBAL SCOPE EXPORTS (for inline onclick handlers)
// ============================================================

window.openDetailModal = openDetailModal;
window.openContractModal = openContractModal;
window.openRenewModal = openRenewModal;
window.openVendorModal = openVendorModal;
window.deleteContract = deleteContract;
window.deleteVendor = deleteVendor;
window.refreshData = refreshData;
window.showConfirm = showConfirm;

// ============================================================
//  BOOTSTRAP
// ============================================================

document.addEventListener('DOMContentLoaded', init);
