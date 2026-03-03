const dashboardView = document.getElementById('dashboardView');
const editorView = document.getElementById('editorView');
const startNewMomHeroBtn = document.getElementById('startNewMomHeroBtn');
const backToDashboardBtn = document.getElementById('backToDashboardBtn');
const newMomBtn = document.getElementById('newMomBtn');

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
let momGeneratedBy = 'M.O.M System';

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
    momGeneratedBy = String(data.generatedBy || 'M.O.M System').trim() || 'M.O.M System';
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

function printPdfFromUrl(url) {
  const fullUrl = new URL(url, window.location.origin).toString();
  const printWindow = window.open(fullUrl, '_blank');
  if (!printWindow) {
    showToast('Popup blocked. Please allow popups to print.', 'error');
    return;
  }

  printWindow.addEventListener('load', () => {
    printWindow.focus();
    printWindow.print();
  });
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

confirmSubmitBtn.addEventListener('click', async () => {
  const mom = collectMomPayload();
  const options = collectSubmitOptions();

  if (!options.generatePdf && !options.printPdf && !options.sendEmail) {
    showToast('Please select at least one submit option.', 'error');
    return;
  }

  confirmSubmitBtn.disabled = true;
  confirmSubmitBtn.textContent = 'Submitting...';

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

    let pdfOpened = false;

    if (pdfUrl && (options.generatePdf || options.sendEmail)) {
      window.open(pdfUrl, '_blank');
      pdfOpened = true;
    }

    if (pdfUrl && options.printPdf) {
      printPdfFromUrl(pdfUrl);
    }

    if (options.sendEmail) {
      const emailDraft = result.emailDraft || {};
      const outlookUrl = String(emailDraft.outlookComposeUrl || '').trim();
      const fallbackMailto = String(emailDraft.mailtoUrl || '').trim();
      const draftWindow = outlookUrl ? window.open(outlookUrl, '_blank') : null;

      if (!draftWindow && fallbackMailto) {
        window.location.href = fallbackMailto;
      }

      if (pdfOpened) {
        showToast('Outlook draft opened. Attach the opened PDF and send when ready.');
      } else {
        showToast('Outlook draft opened. Please generate/download PDF and attach before sending.');
      }
    }
  } catch (error) {
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
refreshDashboardStats();
refreshHealthStatus();
setView('dashboard');
loadDashboardProjects('');
