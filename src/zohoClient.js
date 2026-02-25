const axios = require('axios');
const config = require('./config');

const mockProjects = [
  {
    id: '1001',
    name: 'Corporate Website Revamp',
    projectNumber: 'PRJ-2026-001',
    clientName: 'Acme Corporation',
    workOrderNo: 'WO-ACME-42'
  },
  {
    id: '1002',
    name: 'Mobile App Rollout',
    projectNumber: 'PRJ-2026-002',
    clientName: 'BlueWave Logistics',
    workOrderNo: 'WO-BLUE-88'
  }
];

const mockProjectUsers = {
  '1001': [
    { id: 'u101', displayName: 'Rakesh Patel', email: 'rakesh@elegrow.com', role: 'Project Manager' },
    { id: 'u102', displayName: 'Nidhi Shah', email: 'nidhi@elegrow.com', role: 'Engineer' }
  ],
  '1002': [
    { id: 'u201', displayName: 'Mohana Vamsee', email: 'mohana@elegrow.com', role: 'Engineer' },
    { id: 'u202', displayName: 'Arpit Jain', email: 'arpit@elegrow.com', role: 'Coordinator' }
  ]
};

let currentAccessToken = config.zoho.accessToken || '';
let refreshInFlight = null;

function asText(value) {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value);
  }
  if (typeof value === 'object') {
    const keys = ['id', 'id_string', 'project_id', 'value', 'text', 'name'];
    for (const key of keys) {
      const nested = asText(value[key]);
      if (nested) {
        return nested;
      }
    }
  }
  return '';
}

function extractProjectId(project) {
  const candidates = [];
  const directKeys = ['id_string', 'id', 'project_id', 'projectId', 'projectid'];
  for (const key of directKeys) {
    const rawValue = project?.[key];
    if (rawValue === undefined || rawValue === null) {
      continue;
    }

    if (typeof rawValue === 'string' || typeof rawValue === 'number') {
      const value = String(rawValue).trim();
      if (value) {
        candidates.push(value);
      }
      continue;
    }

    if (typeof rawValue === 'object') {
      const nestedKeys = ['id', 'id_string', 'project_id', 'value', 'text'];
      for (const nestedKey of nestedKeys) {
        const nestedValue = asText(rawValue[nestedKey]).trim();
        if (nestedValue) {
          candidates.push(nestedValue);
        }
      }
    }
  }

  const selfUrl =
    asText(project?.link?.self?.url) ||
    asText(project?.links?.self?.href) ||
    asText(project?.self?.url) ||
    '';
  if (selfUrl) {
    const match = selfUrl.match(/\/projects\/([^/?#]+)/i);
    if (match && match[1]) {
      candidates.push(match[1]);
    }
  }

  if (!candidates.length) {
    return '';
  }

  const uniqueCandidates = [...new Set(candidates)];
  const numericCandidate = uniqueCandidates.find((id) => /^\d{8,}$/.test(id));
  if (numericCandidate) {
    return numericCandidate;
  }

  const idLikeCandidate = uniqueCandidates.find((id) => /^[a-z0-9-]{6,}$/i.test(id));
  return idLikeCandidate || '';
}

function readCustomValue(project, aliases) {
  if (!project || typeof project !== 'object') {
    return '';
  }

  const normalizedAliases = aliases.map((name) => String(name).toLowerCase());

  if (Array.isArray(project.custom_fields)) {
    for (const field of project.custom_fields) {
      const key = String(field?.label || field?.name || '').toLowerCase();
      if (normalizedAliases.some((alias) => key.includes(alias))) {
        return field?.value || field?.text || '';
      }
    }
  }

  for (const [key, value] of Object.entries(project)) {
    const normalizedKey = key.toLowerCase();
    if (normalizedAliases.some((alias) => normalizedKey.includes(alias))) {
      if (typeof value === 'string' || typeof value === 'number') {
        return String(value);
      }
    }
  }

  return '';
}

function normalizeProject(project) {
  const id = extractProjectId(project);
  const name = asText(project.name || project.project_name);
  const projectNumber =
    readCustomValue(project, ['project number', 'project no']) ||
    asText(project.project_no) ||
    asText(project.project_number) ||
    '';
  const workOrderNo =
    readCustomValue(project, ['work order', 'wo number', 'work order no']) ||
    asText(project.work_order_no) ||
    '';
  const clientName =
    readCustomValue(project, ['client']) ||
    asText(project.client_name) ||
    asText(project.client?.name) ||
    asText(project.customer_name) ||
    asText(project.customer?.name) ||
    asText(project.owner_name) ||
    '';
  const updatedAt =
    asText(project.last_updated_time) ||
    asText(project.updated_time) ||
    asText(project.last_update_time) ||
    asText(project.modified_time) ||
    asText(project.updated_at) ||
    '';
  const createdAt =
    asText(project.created_time) ||
    asText(project.created_at) ||
    asText(project.create_time) ||
    '';

  return {
    id: asText(id),
    name: asText(name),
    projectNumber: asText(projectNumber),
    clientName: asText(clientName),
    workOrderNo: asText(workOrderNo),
    updatedAt: asText(updatedAt),
    createdAt: asText(createdAt)
  };
}

function collectProjectList(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  const possibleKeys = [
    'projects',
    'project',
    'data',
    'result',
    'response',
    'items',
    'records'
  ];

  for (const key of possibleKeys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
    if (payload[key] && typeof payload[key] === 'object') {
      const nested = collectProjectList(payload[key]);
      if (nested.length > 0) {
        return nested;
      }
    }
  }

  return [];
}

function hasRefreshCredentials() {
  return (
    Boolean(config.zoho.refreshToken) &&
    Boolean(config.zoho.clientId) &&
    Boolean(config.zoho.clientSecret)
  );
}

function isAxiosStatus(error, status) {
  return axios.isAxiosError(error) && error.response?.status === status;
}

async function refreshZohoAccessToken() {
  if (!hasRefreshCredentials()) {
    throw new Error(
      'Missing Zoho refresh credentials. Set ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET.'
    );
  }

  if (refreshInFlight) {
    return refreshInFlight;
  }

  refreshInFlight = (async () => {
    const base = String(config.zoho.accountsBaseUrl || '').replace(/\/+$/, '');
    const url = `${base}/oauth/v2/token`;
    const body = new URLSearchParams({
      refresh_token: config.zoho.refreshToken,
      client_id: config.zoho.clientId,
      client_secret: config.zoho.clientSecret,
      grant_type: 'refresh_token'
    });

    const response = await axios.post(url, body.toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      timeout: 15000
    });

    const token = response?.data?.access_token;
    if (!token) {
      throw new Error('Zoho refresh succeeded but no access_token was returned.');
    }

    currentAccessToken = token;
    return token;
  })();

  try {
    return await refreshInFlight;
  } finally {
    refreshInFlight = null;
  }
}

function toZohoErrorMessage(error) {
  if (!axios.isAxiosError(error)) {
    return error?.message || 'Unknown Zoho API error';
  }

  const status = error.response?.status;
  const data = error.response?.data;
  const apiMessage =
    data?.message ||
    data?.error?.message ||
    data?.error?.code ||
    data?.code ||
    '';

  if (status === 401) {
    const refreshHint = hasRefreshCredentials()
      ? ' Auto-refresh is configured; verify refresh token/client credentials and accounts domain.'
      : ' Configure refresh credentials to auto-recover token expiry.';
    return (
      'Zoho authentication failed (401). Access token is invalid/expired, ' +
      'or token data-center does not match API domain (zoho.in vs zoho.com).' +
      refreshHint
    );
  }

  if (status === 403) {
    return 'Zoho access denied (403). OAuth scopes or portal permissions are insufficient.';
  }

  if (status) {
    return `Zoho API request failed (${status})${apiMessage ? `: ${apiMessage}` : ''}`;
  }

  return error.message || 'Zoho API request failed';
}

function ensureZohoCredentials() {
  if (!config.zoho.portalId) {
    throw new Error('Missing Zoho credentials. Configure ZOHO_PORTAL_ID in .env');
  }

  if (!currentAccessToken) {
    if (hasRefreshCredentials()) {
      return refreshZohoAccessToken();
    }
    throw new Error(
      'Missing Zoho access token. Configure ZOHO_ACCESS_TOKEN or refresh credentials in .env'
    );
  }

  return Promise.resolve(currentAccessToken);
}

function buildZohoProjectsBasePath() {
  const customEndpoint = String(config.zoho.projectsEndpoint || '').trim();
  if (customEndpoint) {
    if (customEndpoint.includes('/portal/') && customEndpoint.includes('/projects')) {
      return customEndpoint.replace(/\/+$/, '').replace(/\/projects\/?$/, '/projects');
    }

    if (customEndpoint.includes('/portal/')) {
      return `${customEndpoint.replace(/\/+$/, '')}/projects`;
    }

    return `${customEndpoint.replace(/\/+$/, '')}/portal/${config.zoho.portalId}/projects`;
  }

  return `${config.zoho.baseUrl.replace(/\/+$/, '')}/portal/${config.zoho.portalId}/projects`;
}

async function requestZohoGet(url, params = {}, allowRetry = true) {
  try {
    return await axios.get(url, {
      headers: {
        Authorization: `Zoho-oauthtoken ${currentAccessToken}`
      },
      params,
      timeout: 15000
    });
  } catch (error) {
    if (allowRetry && isAxiosStatus(error, 401) && hasRefreshCredentials()) {
      await refreshZohoAccessToken();
      return requestZohoGet(url, params, false);
    }
    throw error;
  }
}

function normalizeUser(user) {
  const id = asText(user.id || user.user_id || user.zpuid || user.uid);
  const firstName = asText(user.first_name || user.firstname);
  const lastName = asText(user.last_name || user.lastname);
  const fullName =
    asText(user.display_name) ||
    asText(user.name) ||
    asText(user.full_name) ||
    `${firstName} ${lastName}`.trim() ||
    asText(user.email) ||
    '';
  const email = asText(user.email || user.email_id || user.mail || user.primary_email);
  const role = asText(user.role || user.role_name || user.user_role);

  return {
    id: asText(id),
    displayName: asText(fullName),
    email: asText(email),
    role: asText(role)
  };
}

function collectUserList(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  const keys = ['users', 'user', 'data', 'result', 'response', 'records', 'items'];
  for (const key of keys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
    if (payload[key] && typeof payload[key] === 'object') {
      const nested = collectUserList(payload[key]);
      if (nested.length > 0) {
        return nested;
      }
    }
  }

  const nestedObjects = Object.values(payload).filter(
    (value) => value && typeof value === 'object' && !Array.isArray(value)
  );
  if (nestedObjects.length) {
    const likelyUsers = nestedObjects.filter((value) => {
      return Boolean(
        value.user_id ||
          value.zpuid ||
          value.uid ||
          value.email ||
          value.email_id ||
          value.display_name ||
          value.name
      );
    });
    if (likelyUsers.length) {
      return likelyUsers;
    }
  }

  return [];
}

function filterUsersToOrgDomain(users) {
  const configured = String(config.zoho.organizationUserEmailDomain || '')
    .trim()
    .toLowerCase()
    .replace(/^@/, '');

  if (!configured) {
    return users;
  }

  return users.filter((user) => {
    const email = String(user?.email || '').trim().toLowerCase();
    return email.endsWith(`@${configured}`);
  });
}

async function fetchFromZoho(query) {
  await ensureZohoCredentials();
  const endpoint = `${buildZohoProjectsBasePath()}/`;

  const range = 200;
  let index = 1;
  const maxPages = 50;
  const uniqueMap = new Map();

  try {
    for (let page = 0; page < maxPages; page += 1) {
      const params = {
        index,
        range
      };
      if (query) {
        params.search_term = query;
      }

      const response = await requestZohoGet(endpoint, params, true);
      const pageProjects = collectProjectList(response.data).map(normalizeProject);
      if (!pageProjects.length) {
        break;
      }

      let addedInThisPage = 0;
      for (const project of pageProjects) {
        const key = `${project.id}::${project.projectNumber}::${project.name}`;
        if (!uniqueMap.has(key)) {
          uniqueMap.set(key, project);
          addedInThisPage += 1;
        }
      }

      // Stop if server repeats the same page or returns fewer records than requested range.
      if (addedInThisPage === 0 || pageProjects.length < range) {
        break;
      }

      index += 1;
    }
  } catch (error) {
    throw new Error(toZohoErrorMessage(error));
  }

  const projects = Array.from(uniqueMap.values());

  if (!query) {
    return projects;
  }

  const needle = query.toLowerCase();
  return projects.filter((project) => {
    return [
      project.id,
      project.name,
      project.projectNumber,
      project.clientName,
      project.workOrderNo
    ]
      .join(' ')
      .toLowerCase()
      .includes(needle);
  });
}

async function fetchProjectUsersFromZoho(projectId) {
  await ensureZohoCredentials();
  const projectBase = buildZohoProjectsBasePath();
  const endpoints = [
    `${projectBase}/${projectId}/users/`,
    `${projectBase}/${projectId}/users`
  ];

  let lastError = null;
  for (const endpoint of endpoints) {
    try {
      const response = await requestZohoGet(endpoint, {}, true);
      const users = collectUserList(response.data).map(normalizeUser);
      const unique = new Map();
      for (const user of users) {
        if (!user.displayName) {
          continue;
        }
        const key = `${user.id}::${user.email}::${user.displayName}`;
        if (!unique.has(key)) {
          unique.set(key, user);
        }
      }
      return Array.from(unique.values());
    } catch (error) {
      lastError = error;
      // try fallback endpoint for 404 only
      if (!isAxiosStatus(error, 404)) {
        throw new Error(toZohoErrorMessage(error));
      }
    }
  }

  throw new Error(toZohoErrorMessage(lastError));
}

async function resolveProjectReferenceToId(projectRef) {
  const raw = String(projectRef || '').trim();
  if (!raw) {
    return '';
  }

  if (/^\d{8,}$/.test(raw)) {
    return raw;
  }

  let matches = await fetchFromZoho(raw);
  const needle = raw.toLowerCase();
  const exactByRef = (projects) =>
    projects.find((project) => {
      return [project.id, project.projectNumber, project.name]
        .map((value) => String(value || '').toLowerCase())
        .includes(needle);
    });

  let exact = exactByRef(matches);

  if (!matches.length || !exact) {
    const allProjects = await fetchFromZoho('');
    const contains = allProjects.filter((project) => {
      return [project.id, project.projectNumber, project.name]
        .join(' ')
        .toLowerCase()
        .includes(needle);
    });
    matches = contains.length ? contains : allProjects;
    exact = exactByRef(matches);
  }

  const selected = exact || matches[0];
  const resolved = String(selected?.id || '').trim();
  return resolved || raw;
}

async function getProjects(query = '') {
  if (config.zoho.useMock) {
    if (!query) {
      return mockProjects;
    }

    const needle = query.toLowerCase();
    return mockProjects.filter((project) => {
      return [
        project.id,
        project.name,
        project.projectNumber,
        project.clientName,
        project.workOrderNo
      ]
        .join(' ')
        .toLowerCase()
        .includes(needle);
    });
  }

  return fetchFromZoho(query);
}

async function getProjectUsers(projectId) {
  const id = String(projectId || '').trim();
  if (!id) {
    throw new Error('Project ID is required to fetch Zoho project users.');
  }

  if (config.zoho.useMock) {
    return filterUsersToOrgDomain(mockProjectUsers[id] || []);
  }

  const resolvedId = await resolveProjectReferenceToId(id);
  const users = await fetchProjectUsersFromZoho(resolvedId);
  return filterUsersToOrgDomain(users);
}

module.exports = {
  getProjects,
  getProjectUsers
};
