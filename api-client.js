(function(global) {
  'use strict';

  let API_BASE_URL = (typeof window !== 'undefined' && window.API_BASE_URL)
    ? window.API_BASE_URL
    : 'https://stoagroupdb-ddre.onrender.com';

  let authToken = null;

  const API = {};

  function setApiBaseUrl(url) {
    API_BASE_URL = url;
    console.log(`API Base URL updated to: ${API_BASE_URL}`);
  }

  function getApiBaseUrl() {
    return API_BASE_URL;
  }

  function setAuthToken(token) {
    authToken = token;
  }

  function getAuthToken() {
    return authToken;
  }

  function clearAuthToken() {
    authToken = null;
  }

  API.setApiBaseUrl = setApiBaseUrl;
  API.getApiBaseUrl = getApiBaseUrl;
  API.setAuthToken = setAuthToken;
  API.getAuthToken = getAuthToken;
  API.clearAuthToken = clearAuthToken;

  async function apiRequest(endpoint, method = 'GET', data = null, token = null) {
    try {
      const options = {
        method,
        headers: {},
      };

      const authTokenToUse = token || authToken;
      if (authTokenToUse) {
        options.headers['Authorization'] = `Bearer ${authTokenToUse}`;
      }

      if (data && (method === 'POST' || method === 'PUT')) {
        if (data instanceof FormData) {
          options.body = data;
        } else {
          options.headers['Content-Type'] = 'application/json';
          options.body = JSON.stringify(data);
        }
      }

      const response = await fetch(`${API_BASE_URL}${endpoint}`, options);
      const text = await response.text();
      const result = text ? (function() { try { return JSON.parse(text); } catch (_) { return {}; } })() : {};

      if (!response.ok) {
        throw new Error(result.error?.message || result.message || `API Error: ${response.status}`);
      }

      return result;
    } catch (error) {
      console.error('API Request Error:', error);
      throw error;
    }
  }

  // Fetch a binary response and return a blob URL
  async function apiBlobRequest(endpoint, token = null) {
    try {
      const options = { method: 'GET', headers: {} };
      const authTokenToUse = token || authToken;
      if (authTokenToUse) {
        options.headers['Authorization'] = `Bearer ${authTokenToUse}`;
      }
      const response = await fetch(`${API_BASE_URL}${endpoint}`, options);
      if (!response.ok) {
        throw new Error(`API Error: ${response.status}`);
      }
      const blob = await response.blob();
      return URL.createObjectURL(blob);
    } catch (error) {
      console.error('API Blob Request Error:', error);
      throw error;
    }
  }

  // ============================================================
  // AUTHENTICATION
  // ============================================================

  async function login(username, password) {
    const result = await apiRequest('/api/auth/login', 'POST', { username, password });
    if (result.success && result.data && result.data.token) {
      authToken = result.data.token;
    }
    return result;
  }

  async function loginWithDomo(domoUser) {
    if (!domoUser || !domoUser.email || !String(domoUser.email).trim()) {
      return { success: false, error: { message: 'Email is required for Domo SSO' } };
    }
    const body = {
      email: String(domoUser.email).trim(),
      name: domoUser.name ? String(domoUser.name).trim() : undefined,
      userId: domoUser.userId ? String(domoUser.userId).trim() : undefined
    };
    const result = await apiRequest('/api/auth/domo', 'POST', body);
    if (result.success && result.data && result.data.token) {
      authToken = result.data.token;
    }
    return result;
  }

  // ============================================================
  // CORE DATA
  // ============================================================

  async function getAllProjects() {
    return apiRequest('/api/core/projects');
  }

  async function getAllPersons() {
    return apiRequest('/api/core/persons');
  }

  // ============================================================
  // CATEGORIES
  // ============================================================

  async function getAllCategories() {
    return apiRequest('/api/contracts/categories');
  }

  async function getCategoryById(id) {
    return apiRequest(`/api/contracts/categories/${id}`);
  }

  async function createCategory(data) {
    return apiRequest('/api/contracts/categories', 'POST', data);
  }

  async function updateCategory(id, data) {
    return apiRequest(`/api/contracts/categories/${id}`, 'PUT', data);
  }

  async function deleteCategory(id) {
    return apiRequest(`/api/contracts/categories/${id}`, 'DELETE');
  }

  // ============================================================
  // VENDORS
  // ============================================================

  async function getAllVendors() {
    return apiRequest('/api/contracts/vendors');
  }

  async function getVendorById(id) {
    return apiRequest(`/api/contracts/vendors/${id}`);
  }

  async function createVendor(data) {
    return apiRequest('/api/contracts/vendors', 'POST', data);
  }

  async function updateVendor(id, data) {
    return apiRequest(`/api/contracts/vendors/${id}`, 'PUT', data);
  }

  async function deleteVendor(id) {
    return apiRequest(`/api/contracts/vendors/${id}`, 'DELETE');
  }

  // ============================================================
  // CONTRACTS
  // ============================================================

  async function getAllContracts() {
    return apiRequest('/api/contracts');
  }

  async function getContractById(id) {
    return apiRequest(`/api/contracts/${id}`);
  }

  async function createContract(data) {
    return apiRequest('/api/contracts', 'POST', data);
  }

  async function updateContract(id, data) {
    return apiRequest(`/api/contracts/${id}`, 'PUT', data);
  }

  async function deleteContract(id) {
    return apiRequest(`/api/contracts/${id}`, 'DELETE');
  }

  async function renewContract(id, data) {
    return apiRequest(`/api/contracts/${id}/renew`, 'POST', data);
  }

  // ============================================================
  // ATTACHMENTS
  // ============================================================

  async function getContractAttachments(contractId) {
    return apiRequest(`/api/contracts/${contractId}/attachments`);
  }

  async function uploadContractAttachment(contractId, file) {
    const formData = new FormData();
    formData.append('file', file);
    return apiRequest(`/api/contracts/${contractId}/attachments`, 'POST', formData);
  }

  async function downloadContractAttachment(attachmentId) {
    return apiBlobRequest(`/api/contracts/attachments/${attachmentId}/download`);
  }

  async function deleteContractAttachment(attachmentId) {
    return apiRequest(`/api/contracts/attachments/${attachmentId}`, 'DELETE');
  }

  // ============================================================
  // ANALYTICS
  // ============================================================

  async function getAnalyticsSummary() {
    return apiRequest('/api/contracts/analytics/summary');
  }

  async function getExpiringContracts(days) {
    return apiRequest(`/api/contracts/analytics/expiring?days=${days}`);
  }

  async function getSpendByCategory() {
    return apiRequest('/api/contracts/analytics/spend-by-category');
  }

  async function getSpendByProperty() {
    return apiRequest('/api/contracts/analytics/spend-by-property');
  }

  async function getSpendOverTime() {
    return apiRequest('/api/contracts/analytics/spend-over-time');
  }

  async function getContractHistory(contractId) {
    return apiRequest(`/api/contracts/${contractId}/history`);
  }

  // ============================================================
  // ENHANCED ANALYTICS
  // ============================================================

  async function getServiceMatrix() {
    return apiRequest('/api/contracts/analytics/service-matrix');
  }

  async function getCompleteness() {
    return apiRequest('/api/contracts/analytics/completeness');
  }

  async function getServiceCoverage() {
    return apiRequest('/api/contracts/analytics/service-coverage');
  }

  async function getVendorDependency() {
    return apiRequest('/api/contracts/analytics/vendor-dependency');
  }

  async function getRenewalPipeline() {
    return apiRequest('/api/contracts/analytics/renewal-pipeline');
  }

  async function getCostBenchmarks() {
    return apiRequest('/api/contracts/analytics/cost-benchmarks');
  }

  // ============================================================
  // EMAIL
  // ============================================================

  async function sendExpiryReminders() {
    return apiRequest('/api/contracts/notifications/expiry-reminders', 'POST');
  }

  // ============================================================
  // EXPOSE ALL FUNCTIONS TO API OBJECT
  // ============================================================

  async function verifyAuth(token) {
    return apiRequest('/api/auth/verify', 'GET', null, token);
  }

  async function getCurrentUser(token) {
    return apiRequest('/api/auth/me', 'GET', null, token);
  }

  API.login = login;
  API.loginWithDomo = loginWithDomo;
  API.verifyAuth = verifyAuth;
  API.getCurrentUser = getCurrentUser;

  API.getAllProjects = getAllProjects;
  API.getAllPersons = getAllPersons;

  API.getAllCategories = getAllCategories;
  API.getCategoryById = getCategoryById;
  API.createCategory = createCategory;
  API.updateCategory = updateCategory;
  API.deleteCategory = deleteCategory;

  API.getAllVendors = getAllVendors;
  API.getVendorById = getVendorById;
  API.createVendor = createVendor;
  API.updateVendor = updateVendor;
  API.deleteVendor = deleteVendor;

  API.getAllContracts = getAllContracts;
  API.getContractById = getContractById;
  API.createContract = createContract;
  API.updateContract = updateContract;
  API.deleteContract = deleteContract;
  API.renewContract = renewContract;

  API.getContractAttachments = getContractAttachments;
  API.uploadContractAttachment = uploadContractAttachment;
  API.downloadContractAttachment = downloadContractAttachment;
  API.deleteContractAttachment = deleteContractAttachment;

  API.getAnalyticsSummary = getAnalyticsSummary;
  API.getExpiringContracts = getExpiringContracts;
  API.getSpendByCategory = getSpendByCategory;
  API.getSpendByProperty = getSpendByProperty;
  API.getSpendOverTime = getSpendOverTime;
  API.getContractHistory = getContractHistory;

  API.getServiceMatrix = getServiceMatrix;
  API.getCompleteness = getCompleteness;
  API.getServiceCoverage = getServiceCoverage;
  API.getVendorDependency = getVendorDependency;
  API.getRenewalPipeline = getRenewalPipeline;
  API.getCostBenchmarks = getCostBenchmarks;

  API.sendExpiryReminders = sendExpiryReminders;

  // PM Management (Admin only, shared with T12 backend)
  API.getAdminUsers = function() { return apiRequest('/api/t12/admin/users'); };
  API.getPropertyManagers = function() { return apiRequest('/api/t12/admin/property-managers'); };
  API.addPropertyManager = function(name, email) { return apiRequest('/api/t12/admin/property-managers', 'POST', { name: name, email: email }); };
  API.changeUserRole = function(userId, role) { return apiRequest('/api/t12/admin/users/' + userId + '/role', 'PUT', { role: role }); };
  API.getAssignments = function() { return apiRequest('/api/t12/admin/assignments'); };
  API.assignProperty = function(propertyName, data) { return apiRequest('/api/t12/admin/assignments/' + encodeURIComponent(propertyName), 'PUT', data); };
  API.removeAssignment = function(propertyName) { return apiRequest('/api/t12/admin/assignments/' + encodeURIComponent(propertyName), 'DELETE'); };

  const globalScope = typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : typeof self !== 'undefined' ? self : this;

  globalScope.API = API;

  Object.keys(API).forEach(key => {
    globalScope[key] = API[key];
  });

})(typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : typeof self !== 'undefined' ? self : this);
