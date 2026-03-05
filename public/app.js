const dashboardView = document.getElementById('dashboardView');
const editorView = document.getElementById('editorView');
const recordsView = document.getElementById('recordsView');
const startNewMomHeroBtn = document.getElementById('startNewMomHeroBtn');
const backToDashboardBtn = document.getElementById('backToDashboardBtn');
const newMomBtn = document.getElementById('newMomBtn');
const openRecordsBtnDashboard = document.getElementById('openRecordsBtnDashboard');
const openRecordsBtnEditor = document.getElementById('openRecordsBtnEditor');
const recordsBackDashboardBtn = document.getElementById('recordsBackDashboardBtn');
const recordsRefreshBtn = document.getElementById('recordsRefreshBtn');
const recordsNewMomBtn = document.getElementById('recordsNewMomBtn');
const recordsSearchInput = document.getElementById('recordsSearchInput');
const recordsStatus = document.getElementById('recordsStatus');
const recordsTableBody = document.querySelector('#recordsTable tbody');
const recordExportModal = document.getElementById('recordExportModal');
const recordExportMeta = document.getElementById('recordExportMeta');
const recordOptGeneratePdf = document.getElementById('recordOptGeneratePdf');
const recordOptPrintPdf = document.getElementById('recordOptPrintPdf');
const recordOptSendEmail = document.getElementById('recordOptSendEmail');
const recordEmailFields = document.getElementById('recordEmailFields');
const recordEmailTo = document.getElementById('recordEmailTo');
const recordEmailCc = document.getElementById('recordEmailCc');
const recordEmailSubject = document.getElementById('recordEmailSubject');
const recordEmailBody = document.getElementById('recordEmailBody');
const confirmRecordExportBtn = document.getElementById('confirmRecordExportBtn');
const cancelRecordExportBtn = document.getElementById('cancelRecordExportBtn');

const dashboardSearchInput = document.getElementById('dashboardSearchInput');
const dashboardRecentProjects = document.getElementById('dashboardRecentProjects');
const zohoListStatus = document.getElementById('zohoListStatus');

const projectEntryHint = document.getElementById('projectEntryHint');
const zohoProjectPickerRow = document.getElementById('zohoProjectPickerRow');
const zohoProjectSelect = document.getElementById('zohoProjectSelect');
const zohoProjectMeta = document.getElementById('zohoProjectMeta');
const zohoMetaOwner = document.getElementById('zohoMetaOwner');
const zohoMetaStatus = document.getElementById('zohoMetaStatus');
const zohoMetaUpdated = document.getElementById('zohoMetaUpdated');
const zohoMetaStatusWrap = document.getElementById('zohoMetaStatusWrap');
const projectNameInput = document.getElementById('projectName');
const projectNoInput = document.getElementById('projectNoWorkOrderNo');
const clientNameInput = document.getElementById('clientName');

const kpiTotalMom = document.getElementById('kpiTotalMom');
const kpiZohoMode = document.getElementById('kpiZohoMode');
const kpiEmailMode = document.getElementById('kpiEmailMode');

const momForm = document.getElementById('momForm');
const addAgendaRowBtn = document.getElementById('addAgendaRowBtn');
const addTaskRowBtn = document.getElementById('addTaskRowBtn');
const addAttendeeRowBtn = document.getElementById('addAttendeeRowBtn');
const agendaTableBody = document.querySelector('#agendaTable tbody');
const taskTableBody = document.querySelector('#taskTable tbody');
const attendeeTableBody = document.querySelector('#attendeeTable tbody');
const toast = document.getElementById('toast');
const authenticityLinePreview = document.getElementById('authenticityLinePreview');

const projectSourceModal = document.getElementById('projectSourceModal');
const chooseManualBtn = document.getElementById('chooseManualBtn');
const chooseZohoBtn = document.getElementById('chooseZohoBtn');
const closeSourceModal = document.getElementById('closeSourceModal');

const projectModal = document.getElementById('projectModal');
const projectSearchInput = document.getElementById('projectSearchInput');
const searchProjectBtn = document.getElementById('searchProjectBtn');
const projectResults = document.getElementById('projectResults');
const closeProjectModal = document.getElementById('closeProjectModal');

const deliveryModal = document.getElementById('deliveryModal');
const optGeneratePdf = document.getElementById('optGeneratePdf');
const optPrintPdf = document.getElementById('optPrintPdf');
const optSendEmail = document.getElementById('optSendEmail');
const emailFields = document.getElementById('emailFields');
const confirmSubmitBtn = document.getElementById('confirmSubmitBtn');
const cancelDelivery = document.getElementById('cancelDelivery');

const TOTAL_MOM_KEY = 'mom_total_submitted_count';
const DASHBOARD_RECENT_LIMIT = 8;
const STAGE_TONE_CLASSES = [
  'tone-default',
  'tone-active',
  'tone-planning',
  'tone-review',
  'tone-hold',
  'tone-delay',
  'tone-complete',
  'tone-cancelled'
];
let dashboardSearchTimer;
let activeProjectSource = 'zoho';
let cachedZohoProjects = [];
let allZohoProjects = [];
let currentProjectUsers = [];
let isProjectUsersLoading = false;
let projectUsersRequestSeq = 0;
const projectUsersCache = new Map();
let currentProjectTasks = [];
let isProjectTasksLoading = false;
let projectTasksRequestSeq = 0;
const projectTasksCache = new Map();
let momGeneratedBy = 'ETPL AI_M.O.M System';
let recordsSearchTimer;
let recordsCache = [];
let activeRecordExport = null;

function showToast(message, type = 'success') {
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 2800);
}

function setView(viewName) {
  dashboardView.classList.toggle('view-active', viewName === 'dashboard');
  editorView.classList.toggle('view-active', viewName === 'editor');
  recordsView.classList.toggle('view-active', viewName === 'records');
}

function getSubmittedCount() {
  return Number.parseInt(localStorage.getItem(TOTAL_MOM_KEY) || '0', 10);
}

function incrementSubmittedCount() {
  const nextCount = getSubmittedCount() + 1;
  localStorage.setItem(TOTAL_MOM_KEY, String(nextCount));
  refreshDashboardStats();
}

function refreshDashboardStats() {
  kpiTotalMom.textContent = String(getSubmittedCount());
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatAuthLineTimestamp(date) {
  const year = date.getFullYear();
  const month = pad2(date.getMonth() + 1);
  const day = pad2(date.getDate());
  const hours = pad2(date.getHours());
  const minutes = pad2(date.getMinutes());
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function generatePreviewDocId(date) {
  return `M.O.M-DRAFT-${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}-${pad2(date.getHours())}${pad2(date.getMinutes())}`;
}

function renderAuthenticityLine(authenticity = null) {
  if (!authenticityLinePreview) {
    return;
  }

  if (authenticity && authenticity.line) {
    authenticityLinePreview.textContent = authenticity.line;
    return;
  }

  const now = new Date();
  const previewLine = `Document ID: ${generatePreviewDocId(now)} | Generated: ${formatAuthLineTimestamp(now)} | Generated by: ${momGeneratedBy}`;
  authenticityLinePreview.textContent = previewLine;
}

async function refreshHealthStatus() {
  try {
    const response = await fetch('/api/health');
    const data = await response.json();

    if (!response.ok || !data.success) {
      throw new Error('Health check failed');
    }

    kpiZohoMode.textContent = data.zohoMode === 'mock' ? 'Mock Mode' : 'Live Mode';
    momGeneratedBy = String(data.generatedBy || 'ETPL AI_M.O.M System').trim() || 'ETPL AI_M.O.M System';
    renderAuthenticityLine();
    if (data.emailMode === 'outlook-draft') {
      kpiEmailMode.textContent = 'Outlook Draft';
    } else {
      kpiEmailMode.textContent = data.emailEnabled ? 'Enabled' : 'Disabled';
    }
  } catch (_error) {
    kpiZohoMode.textContent = 'Unavailable';
    kpiEmailMode.textContent = 'Unavailable';
  }
}

function parseDateToMs(value) {
  if (!value) {
    return 0;
  }

  if (typeof value === 'number') {
    return value > 1e12 ? value : value * 1000;
  }

  const raw = String(value).trim();
  if (!raw) {
    return 0;
  }

  if (/^\d+$/.test(raw)) {
    const numeric = Number.parseInt(raw, 10);
    return numeric > 1e12 ? numeric : numeric * 1000;
  }

  const normalized = raw.replace(/([+-]\d{2})(\d{2})$/, '$1:$2');
  const parsed = Date.parse(normalized);
  if (!Number.isNaN(parsed)) {
    return parsed;
  }

  const fallback = Date.parse(raw);
  return Number.isNaN(fallback) ? 0 : fallback;
}

function getProjectRecencyMs(project) {
  return parseDateToMs(project.updatedAt) || parseDateToMs(project.createdAt);
}

function formatProjectDate(project) {
  const ms = getProjectRecencyMs(project);
  if (!ms) {
    return 'Date unavailable';
  }
  return new Date(ms).toLocaleDateString(undefined, {
    year: 'numeric',
    month: 'short',
    day: '2-digit'
  });
}

function sortProjectsByRecent(projects) {
  return [...projects].sort((a, b) => getProjectRecencyMs(b) - getProjectRecencyMs(a));
}

function getProjectDisplayName(project) {
  if (project.name && project.name.trim()) {
    return project.name.trim();
  }
  if (project.projectNumber && project.projectNumber.trim()) {
    return project.projectNumber.trim();
  }
  return 'Untitled Project';
}

function getProjectOwner(project) {
  return String(project.ownerName || '').trim() || 'Not synced';
}

function getProjectStage(project) {
  return String(project.stage || '').trim() || 'Not synced';
}

function normalizeStageText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatStageLabel(value) {
  const text = String(value || '').replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!text) {
    return 'Not Synced';
  }
  return text
    .split(' ')
    .map((word) => {
      if (!word) {
        return word;
      }
      if (word === word.toUpperCase() && word.length <= 3) {
        return word;
      }
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

function getStageLabel(project) {
  return formatStageLabel(getProjectStage(project));
}

function getStageToneClass(project) {
  const stage = normalizeStageText(getProjectStage(project));
  if (!stage || stage === 'not synced') {
    return 'tone-default';
  }
  if (/cancel|cancelled|canceled|reject|drop|abandon|terminate/.test(stage)) {
    return 'tone-cancelled';
  }
  if (/complete|completed|closed|done|deliver|handover/.test(stage)) {
    return 'tone-complete';
  }
  if (/hold|on hold|blocked|pause|paused|stuck/.test(stage)) {
    return 'tone-hold';
  }
  if (/review|approval|client review|pending client|verify|validation|qa/.test(stage)) {
    return 'tone-review';
  }
  if (/plan|planning|draft|kickoff|kick off|initiat|scoping/.test(stage)) {
    return 'tone-planning';
  }
  if (/delay|delayed|overdue|risk/.test(stage)) {
    return 'tone-delay';
  }
  if (/active|progress|execution|running|ongoing|wip/.test(stage)) {
    return 'tone-active';
  }
  return 'tone-default';
}

function renderZohoProjectMeta(project) {
  if (!zohoProjectMeta || !zohoMetaOwner || !zohoMetaStatus || !zohoMetaUpdated || !zohoMetaStatusWrap) {
    return;
  }

  if (!project) {
    zohoProjectMeta.classList.add('hidden');
    zohoMetaOwner.textContent = '-';
    zohoMetaStatus.textContent = 'Not Synced';
    zohoMetaUpdated.textContent = '-';
    STAGE_TONE_CLASSES.forEach((toneClass) => zohoMetaStatusWrap.classList.remove(toneClass));
    zohoMetaStatusWrap.classList.add('tone-default');
    return;
  }

  zohoProjectMeta.classList.remove('hidden');
  zohoMetaOwner.textContent = getProjectOwner(project);
  zohoMetaStatus.textContent = getStageLabel(project);
  zohoMetaUpdated.textContent = formatProjectDate(project);
  STAGE_TONE_CLASSES.forEach((toneClass) => zohoMetaStatusWrap.classList.remove(toneClass));
  zohoMetaStatusWrap.classList.add(getStageToneClass(project));
}

function normalizeRefValue(value) {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value).trim();
}

function isLikelyZohoProjectId(value) {
  return /^\d{8,}$/.test(normalizeRefValue(value));
}

function getProjectUserRef(project) {
  const id = normalizeRefValue(project?.id);
  if (isLikelyZohoProjectId(id)) {
    return id;
  }

  const projectNumber = normalizeRefValue(project?.projectNumber);
  const name = normalizeRefValue(project?.name);

  const mapped = allZohoProjects.find((candidate) => {
    if (!isLikelyZohoProjectId(candidate?.id)) {
      return false;
    }
    const sameNumber =
      projectNumber &&
      normalizeRefValue(candidate.projectNumber).toLowerCase() === projectNumber.toLowerCase();
    const sameName = name && normalizeRefValue(candidate.name).toLowerCase() === name.toLowerCase();
    return sameNumber || sameName;
  });

  if (mapped?.id) {
    return normalizeRefValue(mapped.id);
  }

  return id || projectNumber || name;
}

function setProjectSource(mode) {
  activeProjectSource = mode;

  const isManual = mode === 'manual';
  projectNameInput.readOnly = !isManual;
  projectNoInput.readOnly = !isManual;
  clientNameInput.readOnly = false;
  zohoProjectPickerRow.classList.toggle('hidden', isManual);

  if (isManual) {
    projectEntryHint.textContent = 'Project details source: Manual Entry';
    projectEntryHint.classList.add('manual');
    clearProjectFields();
    currentProjectUsers = [];
    isProjectUsersLoading = false;
    projectUsersRequestSeq += 1;
    projectUsersCache.clear();
    refreshAttendeeUserDropdowns();
    currentProjectTasks = [];
    isProjectTasksLoading = false;
    projectTasksRequestSeq += 1;
    projectTasksCache.clear();
    refreshTaskDropdowns();
    zohoProjectSelect.value = '';
    renderZohoProjectMeta(null);
  } else {
    projectEntryHint.textContent = 'Project details source: Zoho Projects';
    projectEntryHint.classList.remove('manual');
    if (zohoProjectSelect.value) {
      const currentProject = cachedZohoProjects.find((item) => item._key === zohoProjectSelect.value);
      renderZohoProjectMeta(currentProject || null);
    } else {
      renderZohoProjectMeta(null);
    }
    refreshTaskDropdowns();
  }
}

function buildUserOptionLabel(user) {
  const name = user.displayName || 'Unknown User';
  if (user.email) {
    return `${name} (${user.email})`;
  }
  return name;
}

function isElegrowUser(user) {
  const email = String(user?.email || '').trim().toLowerCase();
  return email.endsWith('@elegrow.com');
}

function filterElegrowUsers(users) {
  const unique = new Map();
  for (const user of users) {
    if (!isElegrowUser(user)) {
      continue;
    }
    const key = `${String(user.displayName || '').trim().toLowerCase()}::${String(user.email || '')
      .trim()
      .toLowerCase()}`;
    if (!unique.has(key)) {
      unique.set(key, user);
    }
  }
  return Array.from(unique.values());
}

function populateAttendeeUserSelect(selectEl, selectedValue = '') {
  if (!selectEl) {
    return;
  }

  const existingValue = selectedValue || selectEl.value || '';
  selectEl.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  if (isProjectUsersLoading) {
    placeholder.textContent = 'Loading Zoho project users...';
  } else if (currentProjectUsers.length) {
    placeholder.textContent = 'Select Elegrow attendee';
  } else if (activeProjectSource === 'manual') {
    placeholder.textContent = 'Switch to Zoho project to load attendees';
  } else {
    placeholder.textContent = 'No @elegrow.com users found for selected project';
  }
  selectEl.appendChild(placeholder);

  for (const user of currentProjectUsers) {
    const option = document.createElement('option');
    option.value = user.displayName;
    option.textContent = buildUserOptionLabel(user);
    selectEl.appendChild(option);
  }

  if (existingValue) {
    const hasValue = currentProjectUsers.some((user) => user.displayName === existingValue);
    if (!hasValue) {
      const custom = document.createElement('option');
      custom.value = existingValue;
      custom.textContent = `${existingValue} (Not in project users)`;
      selectEl.appendChild(custom);
    }
    selectEl.value = existingValue;
  }

  selectEl.disabled = isProjectUsersLoading || currentProjectUsers.length === 0;
}

function refreshAttendeeUserDropdowns() {
  const selects = attendeeTableBody.querySelectorAll('.attendee-elegrow-name');
  selects.forEach((selectEl) => {
    const currentValue = selectEl.value || '';
    populateAttendeeUserSelect(selectEl, currentValue);
  });
}

function buildTaskOptionLabel(task) {
  const name = String(task?.name || '').trim() || 'Unnamed Task';
  const status = String(task?.status || '').trim();
  const list = String(task?.taskListName || '').trim();
  const parts = [];
  if (status) {
    parts.push(`Status: ${status}`);
  }
  if (list) {
    parts.push(`List: ${list}`);
  }
  return parts.length ? `${name} (${parts.join(' | ')})` : name;
}

function populateTaskSelect(selectEl, selectedTaskId = '', selectedTaskName = '') {
  if (!selectEl) {
    return;
  }

  const existingTaskId = selectedTaskId || selectEl.value || '';
  const existingTaskName = selectedTaskName || '';
  selectEl.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.setAttribute('data-task-name', '');
  if (isProjectTasksLoading) {
    placeholder.textContent = 'Loading project tasks from Zoho...';
  } else if (currentProjectTasks.length) {
    placeholder.textContent = 'Select project task';
  } else if (activeProjectSource === 'manual') {
    placeholder.textContent = 'Switch to Zoho project to load tasks';
  } else {
    placeholder.textContent = 'No tasks found for selected project';
  }
  selectEl.appendChild(placeholder);

  for (const task of currentProjectTasks) {
    const option = document.createElement('option');
    option.value = String(task.id || task.name || '').trim();
    option.textContent = buildTaskOptionLabel(task);
    option.setAttribute('data-task-name', String(task.name || '').trim());
    selectEl.appendChild(option);
  }

  if (existingTaskId) {
    const hasValue = [...selectEl.options].some((option) => option.value === existingTaskId);
    if (!hasValue) {
      const custom = document.createElement('option');
      custom.value = existingTaskId;
      custom.setAttribute('data-task-name', existingTaskName || existingTaskId);
      custom.textContent = `${existingTaskName || existingTaskId} (Not in synced project tasks)`;
      selectEl.appendChild(custom);
    }
    selectEl.value = existingTaskId;
  }

  selectEl.disabled = isProjectTasksLoading || currentProjectTasks.length === 0;
}

function refreshTaskDropdowns() {
  const selects = taskTableBody.querySelectorAll('.task-name-select');
  selects.forEach((selectEl) => {
    const currentValue = selectEl.value || '';
    const selectedOption = selectEl.selectedOptions?.[0] || null;
    const currentName = selectedOption?.getAttribute('data-task-name') || '';
    populateTaskSelect(selectEl, currentValue, currentName);
  });
}

function nextIndexFromTable(tbodySelector) {
  const tbody = document.querySelector(tbodySelector);
  return tbody.children.length + 1;
}

function addAgendaRow(row = {}) {
  const tr = document.createElement('tr');
  const index = nextIndexFromTable('#agendaTable tbody');

  tr.innerHTML = `
    <td><input type="text" value="${row.srNo || index}" class="agenda-sr" /></td>
    <td><textarea rows="2" class="agenda-topic">${row.agenda || ''}</textarea></td>
    <td><textarea rows="2" class="agenda-action">${row.actionPlan || ''}</textarea></td>
    <td><input type="text" value="${row.responsibility || ''}" class="agenda-responsibility" /></td>
    <td><button type="button" class="btn btn-light remove-row">Remove</button></td>
  `;

  tr.querySelector('.remove-row').addEventListener('click', () => {
    tr.remove();
    reindexRows();
  });

  agendaTableBody.appendChild(tr);
}

function addTaskRow(row = {}) {
  const tr = document.createElement('tr');
  const index = nextIndexFromTable('#taskTable tbody');

  tr.innerHTML = `
    <td><input type="text" value="${row.srNo || index}" class="task-sr" /></td>
    <td><select class="task-name-select"></select></td>
    <td><input type="text" value="${row.quantityDescription || ''}" class="task-quantity-description" /></td>
    <td><input type="text" value="${row.remarks || ''}" class="task-remarks" /></td>
    <td>
      <select class="task-mom-status">
        <option value="">Select</option>
        <option value="Completed">Completed</option>
        <option value="Not Completed">Not Completed</option>
        <option value="Partially Completed">Partially Completed</option>
      </select>
    </td>
    <td><button type="button" class="btn btn-light remove-row">Remove</button></td>
  `;

  tr.querySelector('.remove-row').addEventListener('click', () => {
    tr.remove();
    reindexRows();
  });

  taskTableBody.appendChild(tr);
  const taskSelect = tr.querySelector('.task-name-select');
  populateTaskSelect(taskSelect, row.taskId || '', row.taskName || '');
  const taskStatus = tr.querySelector('.task-mom-status');
  taskStatus.value = row.status || '';
}

function addAttendeeRow(row = {}) {
  const tr = document.createElement('tr');
  const index = nextIndexFromTable('#attendeeTable tbody');

  tr.innerHTML = `
    <td><input type="text" value="${row.srNo || index}" class="attendee-sr" /></td>
    <td><select class="attendee-elegrow-name"></select></td>
    <td><input type="text" value="${row.clientName || ''}" class="attendee-client-name" /></td>
    <td><button type="button" class="btn btn-light remove-row">Remove</button></td>
  `;

  tr.querySelector('.remove-row').addEventListener('click', () => {
    tr.remove();
    reindexRows();
  });

  attendeeTableBody.appendChild(tr);
  populateAttendeeUserSelect(tr.querySelector('.attendee-elegrow-name'), row.elegrowName || '');
}

function reindexRows() {
  [...agendaTableBody.querySelectorAll('tr')].forEach((row, i) => {
    const sr = row.querySelector('.agenda-sr');
    if (sr && !sr.value.trim()) {
      sr.value = String(i + 1);
    }
  });

  [...attendeeTableBody.querySelectorAll('tr')].forEach((row, i) => {
    const sr = row.querySelector('.attendee-sr');
    if (sr && !sr.value.trim()) {
      sr.value = String(i + 1);
    }
  });

  [...taskTableBody.querySelectorAll('tr')].forEach((row, i) => {
    const sr = row.querySelector('.task-sr');
    if (sr && !sr.value.trim()) {
      sr.value = String(i + 1);
    }
  });
}

function resetRows() {
  agendaTableBody.innerHTML = '';
  taskTableBody.innerHTML = '';
  attendeeTableBody.innerHTML = '';
  addAgendaRow();
  addTaskRow();
  addAttendeeRow();
}

function clearProjectFields() {
  projectNameInput.value = '';
  projectNoInput.value = '';
  clientNameInput.value = '';
}

function resetFormForNewSheet() {
  momForm.reset();
  clearProjectFields();
  currentProjectUsers = [];
  isProjectUsersLoading = false;
  projectUsersRequestSeq += 1;
  refreshAttendeeUserDropdowns();
  currentProjectTasks = [];
  isProjectTasksLoading = false;
  projectTasksRequestSeq += 1;
  refreshTaskDropdowns();
  renderZohoProjectMeta(null);
  resetRows();
  document.getElementById('organizationAddress').value =
    '302, Sangini Aspire, Beside Sanskruti Township Near Pal RTO, Pal-Hajira Road, Pal Gam, Surat, Gujarat - 395009';
  renderAuthenticityLine();
}

function fillProjectFields(project) {
  projectNameInput.value = project.name || '';

  const projectNoValue = project.projectNumber || project.name || '';
  const projectNoWorkOrder = [projectNoValue, project.workOrderNo].filter(Boolean).join(' / ');
  projectNoInput.value = projectNoWorkOrder;
  clientNameInput.value = '';
  renderZohoProjectMeta(project);
}

function getProjectKey(project, index = 0) {
  return String(project.id || project.projectNumber || `${project.name || 'project'}-${index}`);
}

function populateZohoProjectSelect(projects) {
  cachedZohoProjects = sortProjectsByRecent(projects).map((project, index) => ({
    ...project,
    _key: getProjectKey(project, index),
    _userRef: getProjectUserRef(project)
  }));

  const currentValue = zohoProjectSelect.value;
  zohoProjectSelect.innerHTML = '<option value=\"\">Select a Zoho project</option>';

  for (const project of cachedZohoProjects) {
    const option = document.createElement('option');
    option.value = project._key;
    const name = getProjectDisplayName(project);
    const stageLabel = getStageLabel(project);
    option.textContent = `${name} | ${stageLabel}`;
    zohoProjectSelect.appendChild(option);
  }

  if (currentValue && cachedZohoProjects.some((project) => project._key === currentValue)) {
    zohoProjectSelect.value = currentValue;
  }
}

function applyZohoProjectByKey(projectKey) {
  const project = cachedZohoProjects.find((item) => item._key === projectKey);
  if (!project) {
    currentProjectUsers = [];
    isProjectUsersLoading = false;
    projectUsersRequestSeq += 1;
    refreshAttendeeUserDropdowns();
    currentProjectTasks = [];
    isProjectTasksLoading = false;
    projectTasksRequestSeq += 1;
    refreshTaskDropdowns();
    renderZohoProjectMeta(null);
    return;
  }
  fillProjectFields(project);
  syncProjectUsersForProject(project).catch((error) => {
    showToast(error.message || 'Failed to sync project users.', 'error');
  });
  syncProjectTasksForProject(project).catch((error) => {
    showToast(error.message || 'Failed to sync project tasks.', 'error');
  });
}

async function fetchProjects(query = '') {
  const response = await fetch(`/api/zoho/projects?query=${encodeURIComponent(query)}`);
  const data = await response.json();

  if (!response.ok || !data.success) {
    throw new Error(data.message || 'Unable to fetch projects from Zoho.');
  }

  return data.projects || [];
}

async function fetchProjectUsers(projectId) {
  const response = await fetch(`/api/zoho/projects/${encodeURIComponent(projectId)}/users`);
  const data = await response.json();

  if (!response.ok || !data.success) {
    throw new Error(data.message || 'Unable to fetch Zoho project users.');
  }

  return Array.isArray(data.users) ? data.users : [];
}

async function fetchProjectTasks(projectId) {
  const response = await fetch(`/api/zoho/projects/${encodeURIComponent(projectId)}/tasks`);
  const data = await response.json();

  if (!response.ok || !data.success) {
    throw new Error(data.message || 'Unable to fetch Zoho project tasks.');
  }

  return Array.isArray(data.tasks) ? data.tasks : [];
}

async function syncProjectUsersForProject(project) {
  const requestSeq = ++projectUsersRequestSeq;
  const candidateRefs = [
    project?._userRef,
    project?.id,
    project?.projectNumber,
    project?.name
  ]
    .map((value) => normalizeRefValue(value))
    .filter(Boolean);
  const uniqueRefs = [...new Set(candidateRefs)];
  const projectRef = uniqueRefs[0] || '';
  if (!projectRef || projectRef === '[object Object]') {
    currentProjectUsers = [];
    isProjectUsersLoading = false;
    refreshAttendeeUserDropdowns();
    return;
  }

  if (projectUsersCache.has(projectRef)) {
    if (requestSeq !== projectUsersRequestSeq) {
      return;
    }
    currentProjectUsers = projectUsersCache.get(projectRef);
    isProjectUsersLoading = false;
    refreshAttendeeUserDropdowns();
    return;
  }

  isProjectUsersLoading = true;
  currentProjectUsers = [];
  refreshAttendeeUserDropdowns();

  try {
    let users = [];
    let lastError = null;

    for (const ref of uniqueRefs) {
      try {
        users = filterElegrowUsers(await fetchProjectUsers(ref));
        lastError = null;
        break;
      } catch (error) {
        lastError = error;
      }
    }

    if (lastError) {
      throw lastError;
    }

    if (requestSeq !== projectUsersRequestSeq) {
      return;
    }
    projectUsersCache.set(projectRef, users);
    currentProjectUsers = users;
    isProjectUsersLoading = false;
    refreshAttendeeUserDropdowns();
  } catch (error) {
    if (requestSeq !== projectUsersRequestSeq) {
      return;
    }
    currentProjectUsers = [];
    isProjectUsersLoading = false;
    refreshAttendeeUserDropdowns();
    const message = String(error?.message || '').toLowerCase();
    if (message.includes('expected pattern') || message.includes('given url is wrong')) {
      showToast('Project loaded, but Zoho user list could not be resolved for this project.', 'error');
      return;
    }
    throw error;
  }
}

async function syncProjectTasksForProject(project) {
  const requestSeq = ++projectTasksRequestSeq;
  const candidateRefs = [
    project?._userRef,
    project?.id,
    project?.projectNumber,
    project?.name
  ]
    .map((value) => normalizeRefValue(value))
    .filter(Boolean);
  const uniqueRefs = [...new Set(candidateRefs)];
  const projectRef = uniqueRefs[0] || '';
  if (!projectRef || projectRef === '[object Object]') {
    currentProjectTasks = [];
    isProjectTasksLoading = false;
    refreshTaskDropdowns();
    return;
  }

  if (projectTasksCache.has(projectRef)) {
    if (requestSeq !== projectTasksRequestSeq) {
      return;
    }
    currentProjectTasks = projectTasksCache.get(projectRef);
    isProjectTasksLoading = false;
    refreshTaskDropdowns();
    return;
  }

  isProjectTasksLoading = true;
  currentProjectTasks = [];
  refreshTaskDropdowns();

  try {
    let tasks = [];
    let lastError = null;

    for (const ref of uniqueRefs) {
      try {
        tasks = await fetchProjectTasks(ref);
        lastError = null;
        break;
      } catch (error) {
        lastError = error;
      }
    }

    if (lastError) {
      throw lastError;
    }

    if (requestSeq !== projectTasksRequestSeq) {
      return;
    }

    const cleanedTasks = tasks
      .map((task) => ({
        id: String(task.id || '').trim(),
        name: String(task.name || '').trim(),
        status: String(task.status || '').trim(),
        taskListName: String(task.taskListName || '').trim()
      }))
      .filter((task) => task.id || task.name);

    projectTasksCache.set(projectRef, cleanedTasks);
    currentProjectTasks = cleanedTasks;
    isProjectTasksLoading = false;
    refreshTaskDropdowns();
  } catch (error) {
    if (requestSeq !== projectTasksRequestSeq) {
      return;
    }
    currentProjectTasks = [];
    isProjectTasksLoading = false;
    refreshTaskDropdowns();
    const message = String(error?.message || '').toLowerCase();
    if (message.includes('expected pattern') || message.includes('given url is wrong')) {
      showToast('Project loaded, but Zoho task list could not be resolved for this project.', 'error');
      return;
    }
    throw error;
  }
}

function renderProjectResults(projects) {
  projectResults.innerHTML = '';

  if (!projects.length) {
    const li = document.createElement('li');
    li.className = 'result-item';
    li.textContent = 'No projects found.';
    projectResults.appendChild(li);
    return;
  }

  for (const project of projects) {
    const li = document.createElement('li');
    li.className = 'result-item';
    const name = getProjectDisplayName(project);
    const dateText = formatProjectDate(project);
    const stageClass = getStageToneClass(project);
    const stageLabel = getStageLabel(project);

    li.innerHTML = `
      <strong>${name}</strong><br />
      <small>
        Owner: ${getProjectOwner(project)} |
        Status: <span class="result-stage-badge ${stageClass}">${stageLabel}</span> |
        Updated: ${dateText}
      </small>
    `;

    li.addEventListener('click', () => {
      setProjectSource('zoho');
      const matched = cachedZohoProjects.find((item) => {
        const sameId = String(item.id || '') && String(item.id || '') === String(project.id || '');
        const sameNumberAndName =
          String(item.projectNumber || '') === String(project.projectNumber || '') &&
          String(item.name || '') === String(project.name || '');
        return (
          sameId ||
          sameNumberAndName
        );
      });
      zohoProjectSelect.value = matched?._key || '';
      const selectedProject = matched || project;
      fillProjectFields(selectedProject);
      syncProjectUsersForProject(selectedProject).catch((error) => {
        showToast(error.message || 'Failed to sync project users.', 'error');
      });
      syncProjectTasksForProject(selectedProject).catch((error) => {
        showToast(error.message || 'Failed to sync project tasks.', 'error');
      });
      projectModal.close();
      showToast('Project details loaded from Zoho.');
    });

    projectResults.appendChild(li);
  }
}

function renderDashboardProjects(projects, query = '') {
  dashboardRecentProjects.innerHTML = '';

  if (!projects.length) {
    zohoListStatus.textContent = query
      ? 'No recent projects matched your search.'
      : 'No recent Zoho projects found.';
    return;
  }

  zohoListStatus.textContent = query
    ? `Showing ${projects.length} recent match(es)`
    : `Showing ${projects.length} most recently added/modified project(s)`;

  for (const project of projects) {
    const card = document.createElement('article');
    const stageClass = getStageToneClass(project);
    card.className = `recent-project-card ${stageClass}`;

    const name = getProjectDisplayName(project);
    const updatedLabel = formatProjectDate(project);
    const stageLabel = getStageLabel(project);

    card.innerHTML = `
      <div class="recent-project-head">
        <h3>${name}</h3>
        <span class="recent-date">${updatedLabel}</span>
      </div>
      <div class="recent-meta-grid">
        <div class="recent-meta-pill">
          <span>Owner</span>
          <strong>${getProjectOwner(project)}</strong>
        </div>
        <div class="recent-meta-pill stage ${stageClass}">
          <span>Status</span>
          <strong class="stage-value"><i class="stage-dot"></i>${stageLabel}</strong>
        </div>
      </div>
      <button type="button" class="btn btn-light recent-use-btn">Use For New M.O.M</button>
    `;

    card.querySelector('.recent-use-btn').addEventListener('click', () => {
      setProjectSource('zoho');
      const match = cachedZohoProjects.find((item) => {
        return (
          (String(item.id || '') && String(item.id || '') === String(project.id || '')) ||
          (String(item.name || '') === String(project.name || '') &&
            String(item.projectNumber || '') === String(project.projectNumber || ''))
        );
      });
      if (match) {
        zohoProjectSelect.value = match._key;
      }
      const selectedProject = match || project;
      fillProjectFields(selectedProject);
      syncProjectUsersForProject(selectedProject).catch((error) => {
        showToast(error.message || 'Failed to sync project users.', 'error');
      });
      syncProjectTasksForProject(selectedProject).catch((error) => {
        showToast(error.message || 'Failed to sync project tasks.', 'error');
      });
      setView('editor');
      showToast('Recent Zoho project selected for new M.O.M.');
    });

    dashboardRecentProjects.appendChild(card);
  }
}

async function loadDashboardProjects() {
  try {
    zohoListStatus.textContent = 'Syncing Zoho projects...';
    allZohoProjects = await fetchProjects('');
    populateZohoProjectSelect(allZohoProjects);
    applyDashboardFilter();
  } catch (error) {
    dashboardRecentProjects.innerHTML = '';
    zohoListStatus.textContent = error.message;
    showToast(error.message, 'error');
  }
}

function applyDashboardFilter() {
  const query = dashboardSearchInput.value.trim().toLowerCase();
  const filtered = query
    ? allZohoProjects.filter((project) =>
        [getProjectDisplayName(project), project.projectNumber || '', project.ownerName || '', project.stage || '']
          .join(' ')
          .toLowerCase()
          .includes(query)
      )
    : allZohoProjects;

  const recentSubset = sortProjectsByRecent(filtered).slice(0, DASHBOARD_RECENT_LIMIT);
  renderDashboardProjects(recentSubset, query);
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatRecordDateTime(value) {
  const ms = Date.parse(String(value || ''));
  if (Number.isNaN(ms)) {
    return '-';
  }
  return new Date(ms).toLocaleString(undefined, {
    year: 'numeric',
    month: 'short',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function getRecordDaysLeft(record) {
  const expiresAtMs = Date.parse(String(record?.expiresAt || ''));
  if (Number.isNaN(expiresAtMs)) {
    return '-';
  }

  const diffMs = expiresAtMs - Date.now();
  if (diffMs <= 0) {
    return '0';
  }

  const days = Math.ceil(diffMs / (24 * 60 * 60 * 1000));
  return String(days);
}

function buildOutlookDraftUrlForRecord(record, options, pdfAbsoluteUrl) {
  const to = String(options.emailTo || '').trim();
  const cc = String(options.emailCc || '').trim();
  const subject =
    String(options.emailSubject || '').trim() || `Minutes of Meeting - ${record.projectName || 'Project'}`;
  const customBody = String(options.emailBody || '').trim();

  const defaultBody = [
    'Dear Team,',
    '',
    `Please find the Minutes of Meeting for "${record.projectName || '-'}" below.`,
    `PDF Link: ${pdfAbsoluteUrl}`,
    '',
    'Please attach the generated PDF manually before sending.',
    '',
    'Regards,',
    'M.O.M App'
  ].join('\n');

  const body = customBody
    ? `${customBody}\n\nPDF Link: ${pdfAbsoluteUrl}\n\nPlease attach the generated PDF manually before sending.`
    : defaultBody;

  const params = new URLSearchParams();
  if (to) {
    params.set('to', to);
  }
  if (cc) {
    params.set('cc', cc);
  }
  params.set('subject', subject);
  params.set('body', body);

  return `https://outlook.office.com/mail/deeplink/compose?${params.toString()}`;
}

function getRecordOutputBadges(record) {
  const output = record.output || {};
  const badges = [];
  if (output.generatePdf) {
    badges.push('<span class="record-badge">PDF</span>');
  }
  if (output.printPdf) {
    badges.push('<span class="record-badge">Print</span>');
  }
  if (output.sendEmail) {
    badges.push('<span class="record-badge">Email</span>');
  }
  return badges.length ? badges.join(' ') : '<span class="record-badge">-</span>';
}

function renderRecordsTable(records) {
  activeRecordExport = null;
  recordsCache = Array.isArray(records) ? records : [];
  recordsTableBody.innerHTML = '';

  if (!recordsCache.length) {
    recordsStatus.textContent = 'No M.O.M records found.';
    return;
  }

  recordsStatus.textContent = `${recordsCache.length} record(s) found.`;

  for (const record of recordsCache) {
    const tr = document.createElement('tr');
    const pdfUrl = String(record.pdfUrl || '').trim();
    const daysLeft = getRecordDaysLeft(record);
    const pdfLink = pdfUrl
      ? `<a class="record-pdf-link" href="${escapeHtml(pdfUrl)}" target="_blank" rel="noopener">Open PDF</a>`
      : '-';

    tr.innerHTML = `
      <td>${escapeHtml(formatRecordDateTime(record.createdAt))}</td>
      <td>${escapeHtml(record.documentId || '-')}</td>
      <td>${escapeHtml(record.projectName || '-')}</td>
      <td>${escapeHtml(record.meetingTitle || '-')}</td>
      <td>${escapeHtml(record.meetingDate || '-')}</td>
      <td>${escapeHtml(daysLeft)}</td>
      <td class="record-output-cell">${getRecordOutputBadges(record)}</td>
      <td>${pdfLink}</td>
      <td class="record-actions-cell">
        <button type="button" class="btn btn-light record-export-btn">Re-export</button>
        <button type="button" class="btn btn-light record-delete-btn">Delete</button>
      </td>
    `;

    tr.querySelector('.record-export-btn')?.addEventListener('click', () => {
      openRecordExportModal(record);
    });

    tr.querySelector('.record-delete-btn')?.addEventListener('click', () => {
      deleteRecordById(record.id);
    });

    recordsTableBody.appendChild(tr);
  }
}

async function loadRecords(query = '', { silent = false } = {}) {
  try {
    if (!silent) {
      recordsStatus.textContent = 'Loading records...';
    }
    const response = await fetch(`/api/mom/records?query=${encodeURIComponent(query)}`);
    const data = await response.json();

    if (!response.ok || !data.success) {
      throw new Error(data.message || 'Failed to load records.');
    }

    renderRecordsTable(Array.isArray(data.records) ? data.records : []);
  } catch (error) {
    recordsCache = [];
    activeRecordExport = null;
    recordsTableBody.innerHTML = '';
    recordsStatus.textContent = error.message || 'Failed to load records.';
    if (!silent) {
      showToast(error.message || 'Failed to load records.', 'error');
    }
  }
}

function getMeetingTypes() {
  const checked = [...document.querySelectorAll('input[name="meetingType"]:checked')];
  return checked.map((item) => item.value);
}

function collectAgendaRows() {
  return [...agendaTableBody.querySelectorAll('tr')].map((row, index) => {
    return {
      srNo: row.querySelector('.agenda-sr')?.value || String(index + 1),
      agenda: row.querySelector('.agenda-topic')?.value || '',
      actionPlan: row.querySelector('.agenda-action')?.value || '',
      responsibility: row.querySelector('.agenda-responsibility')?.value || ''
    };
  });
}

function collectTaskRows() {
  return [...taskTableBody.querySelectorAll('tr')].map((row, index) => {
    const taskSelect = row.querySelector('.task-name-select');
    const selectedOption = taskSelect?.selectedOptions?.[0] || null;
    const taskId = taskSelect?.value || '';
    const taskName =
      selectedOption?.getAttribute('data-task-name') ||
      selectedOption?.textContent ||
      '';

    return {
      srNo: row.querySelector('.task-sr')?.value || String(index + 1),
      taskId,
      taskName,
      quantityDescription: row.querySelector('.task-quantity-description')?.value || '',
      remarks: row.querySelector('.task-remarks')?.value || '',
      status: row.querySelector('.task-mom-status')?.value || ''
    };
  });
}

function collectAttendeeRows() {
  return [...attendeeTableBody.querySelectorAll('tr')].map((row, index) => {
    return {
      srNo: row.querySelector('.attendee-sr')?.value || String(index + 1),
      elegrowName: row.querySelector('.attendee-elegrow-name')?.value || '',
      clientName: row.querySelector('.attendee-client-name')?.value || ''
    };
  });
}

function collectMomPayload() {
  return {
    meetingTitle: document.getElementById('meetingTitle').value,
    projectName: projectNameInput.value,
    projectNoWorkOrderNo: projectNoInput.value,
    projectOwner: activeProjectSource === 'zoho' ? String(zohoMetaOwner?.textContent || '').trim() : '',
    projectStatus: activeProjectSource === 'zoho' ? String(zohoMetaStatus?.textContent || '').trim() : '',
    projectUpdated: activeProjectSource === 'zoho' ? String(zohoMetaUpdated?.textContent || '').trim() : '',
    clientName: clientNameInput.value,
    meetingDate: document.getElementById('meetingDate').value,
    meetingTime: document.getElementById('meetingTime').value,
    entryTime: document.getElementById('entryTime').value,
    exitTime: document.getElementById('exitTime').value,
    meetingLocation: document.getElementById('meetingLocation').value,
    meetingCalledBy: document.getElementById('meetingCalledBy').value,
    meetingType: getMeetingTypes(),
    meetingTypeOther: document.getElementById('meetingTypeOther').value,
    facilitatorRepresentative: document.getElementById('facilitatorRepresentative').value,
    elegrowRepresentative: document.getElementById('elegrowRepresentative').value,
    clientRepresentative: document.getElementById('clientRepresentative').value,
    agendaRows: collectAgendaRows(),
    taskRows: collectTaskRows(),
    attendeeRows: collectAttendeeRows(),
    organizationAddress: document.getElementById('organizationAddress').value
  };
}

function collectSubmitOptions() {
  return {
    generatePdf: optGeneratePdf.checked,
    printPdf: optPrintPdf.checked,
    sendEmail: optSendEmail.checked,
    emailTo: document.getElementById('emailTo').value,
    emailCc: document.getElementById('emailCc').value,
    emailSubject: document.getElementById('emailSubject').value,
    emailBody: document.getElementById('emailBody').value
  };
}

function toggleEmailFields() {
  emailFields.classList.toggle('hidden', !optSendEmail.checked);
}

function toggleRecordEmailFields() {
  recordEmailFields.classList.toggle('hidden', !recordOptSendEmail.checked);
}

function getRecordById(recordId) {
  const id = String(recordId || '').trim();
  if (!id) {
    return null;
  }
  return recordsCache.find((record) => String(record.id || '') === id) || null;
}

function openRecordExportModal(record) {
  if (!record) {
    showToast('Record not found for export.', 'error');
    return;
  }

  activeRecordExport = record;
  recordOptGeneratePdf.checked = true;
  recordOptPrintPdf.checked = false;
  recordOptSendEmail.checked = false;
  recordEmailTo.value = '';
  recordEmailCc.value = '';
  recordEmailSubject.value = `Minutes of Meeting - ${record.projectName || 'Project'}`;
  recordEmailBody.value = '';
  recordExportMeta.textContent = `Document ID: ${record.documentId || '-'} | Project: ${record.projectName || '-'}`;
  toggleRecordEmailFields();

  if (typeof recordExportModal.showModal === 'function') {
    recordExportModal.showModal();
  }
}

async function deleteRecordById(recordId) {
  const record = getRecordById(recordId);
  if (!record) {
    showToast('Record not found.', 'error');
    return;
  }

  const ok = window.confirm(`Delete record ${record.documentId || record.id}? This action cannot be undone.`);
  if (!ok) {
    return;
  }

  try {
    const response = await fetch(`/api/mom/records/${encodeURIComponent(record.id)}`, {
      method: 'DELETE'
    });
    const data = await response.json();
    if (!response.ok || !data.success) {
      throw new Error(data.message || 'Failed to delete record.');
    }
    showToast('Record deleted successfully.');
    loadRecords(recordsSearchInput.value.trim(), { silent: true });
  } catch (error) {
    showToast(error.message || 'Failed to delete record.', 'error');
  }
}

function collectRecordExportOptions() {
  return {
    generatePdf: recordOptGeneratePdf.checked,
    printPdf: recordOptPrintPdf.checked,
    sendEmail: recordOptSendEmail.checked,
    emailTo: recordEmailTo.value,
    emailCc: recordEmailCc.value,
    emailSubject: recordEmailSubject.value,
    emailBody: recordEmailBody.value
  };
}

function printPdfFromUrl(url, existingWindow = null) {
  const fullUrl = new URL(url, window.location.origin).toString();
  const printWindow = existingWindow || window.open('about:blank', '_blank');
  if (!printWindow) {
    showToast('Popup blocked. Please allow popups to print.', 'error');
    return false;
  }

  printWindow.location.href = fullUrl;
  printWindow.addEventListener('load', () => {
    printWindow.focus();
    printWindow.print();
  });
  return true;
}

function closeAllDialogs() {
  if (projectSourceModal.open) {
    projectSourceModal.close();
  }
  if (projectModal.open) {
    projectModal.close();
  }
  if (deliveryModal.open) {
    deliveryModal.close();
  }
  if (recordExportModal.open) {
    recordExportModal.close();
  }
}

function openProjectSourcePicker() {
  setView('editor');
  resetFormForNewSheet();
  setProjectSource('zoho');
  projectSearchInput.value = '';
  projectResults.innerHTML = '';

  if (!allZohoProjects.length) {
    fetchProjects('')
      .then((projects) => {
        allZohoProjects = projects;
        populateZohoProjectSelect(allZohoProjects);
      })
      .catch((error) => {
        showToast(error.message || 'Unable to sync Zoho projects.', 'error');
      });
  }
}

async function openZohoProjectPicker() {
  setProjectSource('zoho');
  if (projectSourceModal.open) {
    projectSourceModal.close();
  }

  try {
    if (!allZohoProjects.length) {
      allZohoProjects = await fetchProjects('');
      populateZohoProjectSelect(allZohoProjects);
    }
    renderProjectResults(allZohoProjects);
    showToast('Zoho projects synced. Select one from the Project dropdown.');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

startNewMomHeroBtn.addEventListener('click', openProjectSourcePicker);
newMomBtn.addEventListener('click', openProjectSourcePicker);
openRecordsBtnDashboard.addEventListener('click', () => {
  closeAllDialogs();
  setView('records');
  loadRecords(recordsSearchInput.value.trim());
});
openRecordsBtnEditor.addEventListener('click', () => {
  closeAllDialogs();
  setView('records');
  loadRecords(recordsSearchInput.value.trim());
});
recordsBackDashboardBtn.addEventListener('click', () => {
  closeAllDialogs();
  setView('dashboard');
});
recordsRefreshBtn.addEventListener('click', () => {
  loadRecords(recordsSearchInput.value.trim());
});
recordsNewMomBtn.addEventListener('click', openProjectSourcePicker);

chooseManualBtn.addEventListener('click', () => {
  setProjectSource('manual');
  if (projectSourceModal.open) {
    projectSourceModal.close();
  }
  showToast('Manual project entry enabled.');
});

chooseZohoBtn.addEventListener('click', openZohoProjectPicker);

closeSourceModal.addEventListener('click', () => {
  if (projectSourceModal.open) {
    projectSourceModal.close();
  }
});

zohoProjectSelect.addEventListener('change', () => {
  applyZohoProjectByKey(zohoProjectSelect.value);
});

backToDashboardBtn.addEventListener('click', () => {
  closeAllDialogs();
  setView('dashboard');
});

dashboardSearchInput.addEventListener('input', () => {
  clearTimeout(dashboardSearchTimer);
  dashboardSearchTimer = setTimeout(() => {
    applyDashboardFilter();
  }, 260);
});

recordsSearchInput.addEventListener('input', () => {
  clearTimeout(recordsSearchTimer);
  recordsSearchTimer = setTimeout(() => {
    loadRecords(recordsSearchInput.value.trim(), { silent: true });
  }, 260);
});

searchProjectBtn.addEventListener('click', async () => {
  try {
    const query = projectSearchInput.value.trim();
    const projects = await fetchProjects(query);
    renderProjectResults(projects);
  } catch (error) {
    showToast(error.message, 'error');
  }
});

closeProjectModal.addEventListener('click', () => {
  if (projectModal.open) {
    projectModal.close();
  }
});

addAgendaRowBtn.addEventListener('click', () => addAgendaRow());
addTaskRowBtn.addEventListener('click', () => addTaskRow());
addAttendeeRowBtn.addEventListener('click', () => addAttendeeRow());

momForm.addEventListener('submit', (event) => {
  event.preventDefault();

  if (activeProjectSource === 'zoho' && !zohoProjectSelect.value) {
    showToast('Please select a Zoho project from the dropdown.', 'error');
    return;
  }

  if (!projectNameInput.value.trim()) {
    if (activeProjectSource === 'manual') {
      showToast('Please enter project name manually.', 'error');
    } else {
      showToast('Please select a project from Zoho or choose Manual Entry.', 'error');
    }
    return;
  }

  if (typeof deliveryModal.showModal === 'function') {
    deliveryModal.showModal();
  }
});

optSendEmail.addEventListener('change', toggleEmailFields);
cancelDelivery.addEventListener('click', () => {
  if (deliveryModal.open) {
    deliveryModal.close();
  }
});

recordOptSendEmail.addEventListener('change', toggleRecordEmailFields);
cancelRecordExportBtn.addEventListener('click', () => {
  if (recordExportModal.open) {
    recordExportModal.close();
  }
});

confirmRecordExportBtn.addEventListener('click', () => {
  if (!activeRecordExport) {
    showToast('Please select a record first.', 'error');
    return;
  }

  const options = collectRecordExportOptions();
  if (!options.generatePdf && !options.printPdf && !options.sendEmail) {
    showToast('Please select at least one export option.', 'error');
    return;
  }

  const record = activeRecordExport;
  const pdfAbsoluteUrl =
    String(record.pdfAbsoluteUrl || '').trim() ||
    (record.pdfUrl ? new URL(record.pdfUrl, window.location.origin).toString() : '');

  if (!pdfAbsoluteUrl) {
    showToast('PDF URL is unavailable for this record.', 'error');
    return;
  }

  const needsPdfWindow = Boolean(options.generatePdf || options.printPdf || options.sendEmail);
  const preopenedPdfWindow = needsPdfWindow ? window.open('about:blank', '_blank') : null;
  const preopenedOutlookWindow = options.sendEmail ? window.open('about:blank', '_blank') : null;

  try {
    let pdfOpened = false;

    if ((options.generatePdf || options.sendEmail) && preopenedPdfWindow) {
      preopenedPdfWindow.location.href = pdfAbsoluteUrl;
      pdfOpened = true;
    } else if (options.generatePdf) {
      window.open(pdfAbsoluteUrl, '_blank', 'noopener');
      pdfOpened = true;
    }

    if (options.printPdf) {
      const printed = printPdfFromUrl(pdfAbsoluteUrl, preopenedPdfWindow && !pdfOpened ? preopenedPdfWindow : null);
      pdfOpened = pdfOpened || printed;
    }

    if (options.sendEmail) {
      const outlookUrl = buildOutlookDraftUrlForRecord(record, options, pdfAbsoluteUrl);
      if (preopenedOutlookWindow) {
        preopenedOutlookWindow.location.href = outlookUrl;
      } else {
        const win = window.open(outlookUrl, '_blank', 'noopener');
        if (!win) {
          showToast('Popup blocked for Outlook draft. Please allow popups and retry.', 'error');
        }
      }
      if (pdfOpened) {
        showToast('Record re-exported. Outlook draft opened with same PDF link.');
      } else {
        showToast('Outlook draft opened. Please attach PDF before sending.');
      }
    } else {
      showToast('Record export completed.');
    }

    if (recordExportModal.open) {
      recordExportModal.close();
    }
  } catch (error) {
    if (preopenedPdfWindow) {
      preopenedPdfWindow.close();
    }
    if (preopenedOutlookWindow) {
      preopenedOutlookWindow.close();
    }
    showToast(error.message || 'Record export failed.', 'error');
  }
});

confirmSubmitBtn.addEventListener('click', async () => {
  const mom = collectMomPayload();
  const options = collectSubmitOptions();

  if (!options.generatePdf && !options.printPdf && !options.sendEmail) {
    showToast('Please select at least one submit option.', 'error');
    return;
  }

  confirmSubmitBtn.disabled = true;
  confirmSubmitBtn.textContent = 'Submitting...';

  const needsPdfWindow = Boolean(options.generatePdf || options.printPdf || options.sendEmail);
  const preopenedPdfWindow = needsPdfWindow ? window.open('about:blank', '_blank') : null;
  const preopenedOutlookWindow = options.sendEmail ? window.open('about:blank', '_blank') : null;

  try {
    const response = await fetch('/api/mom/submit', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ mom, options })
    });

    const data = await response.json();

    if (!response.ok || !data.success) {
      const message = data.errors?.join(' ') || data.message || 'Submission failed.';
      throw new Error(message);
    }

    const result = data.result || {};
    const pdfUrl = result.pdfUrl;

    if (deliveryModal.open) {
      deliveryModal.close();
    }
    renderAuthenticityLine(result.authenticity || null);
    showToast('M.O.M submitted successfully.');
    incrementSubmittedCount();

    const pdfAbsoluteUrl = String(result.pdfAbsoluteUrl || '').trim() || (pdfUrl ? new URL(pdfUrl, window.location.origin).toString() : '');
    let pdfOpened = false;

    if (pdfAbsoluteUrl && (options.generatePdf || options.sendEmail) && preopenedPdfWindow) {
      preopenedPdfWindow.location.href = pdfAbsoluteUrl;
      pdfOpened = true;
    } else if (pdfAbsoluteUrl && options.generatePdf) {
      window.location.href = pdfAbsoluteUrl;
      pdfOpened = true;
    }

    if (pdfAbsoluteUrl && options.printPdf) {
      const printed = printPdfFromUrl(pdfAbsoluteUrl, preopenedPdfWindow && !pdfOpened ? preopenedPdfWindow : null);
      pdfOpened = pdfOpened || printed;
    }

    if (options.sendEmail) {
      const emailDraft = result.emailDraft || {};
      const outlookUrl = String(emailDraft.outlookComposeUrl || '').trim();
      if (!outlookUrl) {
        showToast('Outlook draft URL unavailable from server.', 'error');
      } else if (preopenedOutlookWindow) {
        preopenedOutlookWindow.location.href = outlookUrl;
      } else {
        const outlookWindow = window.open(outlookUrl, '_blank', 'noopener');
        if (!outlookWindow) {
          showToast('Popup blocked for Outlook draft. Please allow popups and retry.', 'error');
        }
      }

      if (pdfOpened) {
        showToast('Outlook draft opened. Attach the generated PDF and send.');
      } else {
        showToast('Outlook draft opened. Generate/download PDF and attach before sending.');
      }
    }

    await loadRecords('', { silent: true });
    setView('records');
  } catch (error) {
    if (preopenedPdfWindow) {
      preopenedPdfWindow.close();
    }
    if (preopenedOutlookWindow) {
      preopenedOutlookWindow.close();
    }
    showToast(error.message, 'error');
  } finally {
    confirmSubmitBtn.disabled = false;
    confirmSubmitBtn.textContent = 'Confirm Submit';
  }
});

setProjectSource('zoho');
resetRows();
renderAuthenticityLine();
toggleEmailFields();
toggleRecordEmailFields();
refreshDashboardStats();
refreshHealthStatus();
setView('dashboard');
loadDashboardProjects('');
loadRecords('', { silent: true });
