const dashboardView = document.getElementById('dashboardView');
const editorView = document.getElementById('editorView');
const navDashboardBtn = document.getElementById('navDashboardBtn');
const navEditorBtn = document.getElementById('navEditorBtn');
const startNewMomBtn = document.getElementById('startNewMomBtn');
const startNewMomHeroBtn = document.getElementById('startNewMomHeroBtn');
const backToDashboardBtn = document.getElementById('backToDashboardBtn');
const newMomBtn = document.getElementById('newMomBtn');

const dashboardSearchInput = document.getElementById('dashboardSearchInput');
const dashboardRecentProjects = document.getElementById('dashboardRecentProjects');
const zohoListStatus = document.getElementById('zohoListStatus');

const projectEntryHint = document.getElementById('projectEntryHint');
const zohoProjectPickerRow = document.getElementById('zohoProjectPickerRow');
const zohoProjectSelect = document.getElementById('zohoProjectSelect');
const projectNameInput = document.getElementById('projectName');
const projectNoInput = document.getElementById('projectNoWorkOrderNo');
const clientNameInput = document.getElementById('clientName');

const kpiTotalMom = document.getElementById('kpiTotalMom');
const kpiZohoMode = document.getElementById('kpiZohoMode');
const kpiEmailMode = document.getElementById('kpiEmailMode');

const momForm = document.getElementById('momForm');
const addAgendaRowBtn = document.getElementById('addAgendaRowBtn');
const addAttendeeRowBtn = document.getElementById('addAttendeeRowBtn');
const agendaTableBody = document.querySelector('#agendaTable tbody');
const attendeeTableBody = document.querySelector('#attendeeTable tbody');
const toast = document.getElementById('toast');

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
let dashboardSearchTimer;
let activeProjectSource = 'zoho';
let cachedZohoProjects = [];
let allZohoProjects = [];
let currentProjectUsers = [];
let isProjectUsersLoading = false;
let projectUsersRequestSeq = 0;
const projectUsersCache = new Map();

function showToast(message, type = 'success') {
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 2800);
}

function setActiveNav(viewName) {
  navDashboardBtn.classList.toggle('nav-chip-active', viewName === 'dashboard');
  navEditorBtn.classList.toggle('nav-chip-active', viewName === 'editor');
}

function setView(viewName) {
  dashboardView.classList.toggle('view-active', viewName === 'dashboard');
  editorView.classList.toggle('view-active', viewName === 'editor');
  setActiveNav(viewName);
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

async function refreshHealthStatus() {
  try {
    const response = await fetch('/api/health');
    const data = await response.json();

    if (!response.ok || !data.success) {
      throw new Error('Health check failed');
    }

    kpiZohoMode.textContent = data.zohoMode === 'mock' ? 'Mock Mode' : 'Live Mode';
    kpiEmailMode.textContent = data.emailEnabled ? 'Enabled' : 'Disabled';
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
  clientNameInput.readOnly = !isManual;
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
    zohoProjectSelect.value = '';
  } else {
    projectEntryHint.textContent = 'Project details source: Zoho Projects';
    projectEntryHint.classList.remove('manual');
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
}

function resetRows() {
  agendaTableBody.innerHTML = '';
  attendeeTableBody.innerHTML = '';
  addAgendaRow();
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
  resetRows();
  document.getElementById('organizationAddress').value =
    '302, Sangini Aspire, Beside Sanskruti Township Near Pal RTO, Pal-Hajira Road, Pal Gam, Surat, Gujarat - 395009';
}

function fillProjectFields(project) {
  projectNameInput.value = project.name || '';

  const projectNoValue = project.projectNumber || project.name || '';
  const projectNoWorkOrder = [projectNoValue, project.workOrderNo].filter(Boolean).join(' / ');
  projectNoInput.value = projectNoWorkOrder;
  clientNameInput.value = project.clientName || '';
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
    const client = project.clientName ? ` | ${project.clientName}` : '';
    option.textContent = `${name}${client}`;
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
    return;
  }
  fillProjectFields(project);
  syncProjectUsersForProject(project).catch((error) => {
    showToast(error.message || 'Failed to sync project users.', 'error');
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

    li.innerHTML = `
      <strong>${name}</strong><br />
      <small>Client: ${project.clientName || '-'} | Updated: ${dateText}</small>
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
    card.className = 'recent-project-card';

    const name = getProjectDisplayName(project);
    const updatedLabel = formatProjectDate(project);

    card.innerHTML = `
      <div class="recent-project-head">
        <h3>${name}</h3>
        <span class="recent-date">${updatedLabel}</span>
      </div>
      <p class="recent-client">Client: ${project.clientName || '-'}</p>
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
        [getProjectDisplayName(project), project.clientName || '', project.projectNumber || '']
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
    minutes: document.getElementById('minutes').value,
    agendaRows: collectAgendaRows(),
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
  projectSearchInput.value = '';
  projectResults.innerHTML = '';

  if (typeof projectSourceModal.showModal === 'function') {
    projectSourceModal.showModal();
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

startNewMomBtn.addEventListener('click', openProjectSourcePicker);
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

navDashboardBtn.addEventListener('click', () => {
  setView('dashboard');
});

navEditorBtn.addEventListener('click', () => {
  setView('editor');
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

  if (options.sendEmail && !options.emailTo.trim()) {
    showToast('Please provide recipient email in To field.', 'error');
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
    showToast('M.O.M submitted successfully.');
    incrementSubmittedCount();

    if (pdfUrl && options.generatePdf) {
      window.open(pdfUrl, '_blank');
    }

    if (pdfUrl && options.printPdf) {
      printPdfFromUrl(pdfUrl);
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
toggleEmailFields();
refreshDashboardStats();
refreshHealthStatus();
setView('dashboard');
loadDashboardProjects('');
