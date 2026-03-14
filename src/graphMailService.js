const fs = require('fs');
const axios = require('axios');

let cachedToken = {
  value: '',
  expiresAt: 0
};

function isGraphDraftConfigured(config) {
  const graph = config?.graph || {};
  return Boolean(
    graph.enabled &&
      graph.tenantId &&
      graph.clientId &&
      graph.clientSecret &&
      graph.mailboxUser
  );
}

function normalizeBaseUrl(url) {
  return String(url || '').replace(/\/+$/, '');
}

function toBase64(value) {
  return String(value || '')
    .replace(/-/g, '+')
    .replace(/_/g, '/')
    .padEnd(Math.ceil(String(value || '').length / 4) * 4, '=');
}

function maskValue(value, start = 4, end = 4) {
  const input = String(value || '').trim();
  if (!input) {
    return '';
  }
  if (input.length <= start + end) {
    return `${input.slice(0, 1)}***${input.slice(-1)}`;
  }
  return `${input.slice(0, start)}***${input.slice(-end)}`;
}

function parseAddressList(value) {
  return String(value || '')
    .split(/[;,]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function buildRecipients(addresses = []) {
  return addresses.map((address) => ({
    emailAddress: { address }
  }));
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function textToHtmlParagraphs(text) {
  const safe = escapeHtml(text);
  return safe
    .split(/\r?\n\r?\n/)
    .map((segment) => `<p>${segment.replace(/\r?\n/g, '<br>')}</p>`)
    .join('');
}

function buildGraphBodyHtml({ bodyText = '', pdfUrl = '' }) {
  const safePdfUrl = escapeHtml(pdfUrl);
  const clickableLink = safePdfUrl
    ? `<p><a href="${safePdfUrl}" target="_blank" rel="noopener noreferrer">${safePdfUrl}</a></p>`
    : '';
  return `${textToHtmlParagraphs(bodyText)}${clickableLink}`;
}

function resolveGraphErrorMessage(error) {
  const status = error?.response?.status;
  const graphMessage = error?.response?.data?.error?.message;
  if (status && graphMessage) {
    return `Microsoft Graph request failed (${status}): ${graphMessage}`;
  }
  if (status) {
    return `Microsoft Graph request failed (${status}).`;
  }
  return error?.message || 'Microsoft Graph request failed.';
}

function parseGraphError(error) {
  const status = error?.response?.status || null;
  const graphCode = String(error?.response?.data?.error?.code || '').trim();
  const graphMessage = String(error?.response?.data?.error?.message || '').trim();
  const requestId = String(error?.response?.headers?.['request-id'] || error?.response?.headers?.['x-ms-request-id'] || '').trim();
  return {
    status,
    code: graphCode || '',
    message: graphMessage || error?.message || 'Microsoft Graph request failed.',
    requestId: requestId || ''
  };
}

function decodeJwtClaims(token = '') {
  const parts = String(token || '').split('.');
  if (parts.length < 2) {
    return {};
  }
  try {
    const payload = Buffer.from(toBase64(parts[1]), 'base64').toString('utf8');
    const parsed = JSON.parse(payload);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_error) {
    return {};
  }
}

async function getGraphAccessToken(config) {
  const now = Date.now();
  if (cachedToken.value && cachedToken.expiresAt - 60_000 > now) {
    return cachedToken.value;
  }

  const graph = config.graph;
  const authorityHost = normalizeBaseUrl(graph.authorityHost);
  const tokenUrl = `${authorityHost}/${encodeURIComponent(graph.tenantId)}/oauth2/v2.0/token`;
  const form = new URLSearchParams();
  form.set('client_id', graph.clientId);
  form.set('client_secret', graph.clientSecret);
  form.set('scope', graph.scope || 'https://graph.microsoft.com/.default');
  form.set('grant_type', 'client_credentials');

  const response = await axios.post(tokenUrl, form.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    timeout: 30000
  });

  const token = String(response?.data?.access_token || '').trim();
  const expiresIn = Number.parseInt(String(response?.data?.expires_in || '3600'), 10);
  if (!token) {
    throw new Error('Microsoft Graph token response did not include access_token.');
  }

  cachedToken = {
    value: token,
    expiresAt: now + Math.max(300, Number.isNaN(expiresIn) ? 3600 : expiresIn) * 1000
  };

  return token;
}

async function createGraphDraft(config, payload) {
  const token = await getGraphAccessToken(config);
  const graphBaseUrl = normalizeBaseUrl(config.graph.baseUrl || 'https://graph.microsoft.com/v1.0');
  const mailboxUser = encodeURIComponent(config.graph.mailboxUser);

  const toRecipients = buildRecipients(parseAddressList(payload.to));
  const ccRecipients = buildRecipients(parseAddressList(payload.cc));
  const requestBody = {
    subject: String(payload.subject || '').trim(),
    body: {
      contentType: 'HTML',
      content: String(payload.bodyHtml || '').trim()
    },
    toRecipients,
    ccRecipients
  };

  let draftResponse;
  try {
    draftResponse = await axios.post(
      `${graphBaseUrl}/users/${mailboxUser}/messages`,
      requestBody,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        timeout: 45000
      }
    );
  } catch (error) {
    throw new Error(resolveGraphErrorMessage(error));
  }

  const draftId = String(draftResponse?.data?.id || '').trim();
  const webLink = String(draftResponse?.data?.webLink || '').trim();

  if (!draftId) {
    throw new Error('Microsoft Graph draft creation succeeded but draft ID is missing.');
  }

  let attachmentNote = '';
  if (payload.pdfFilePath && fs.existsSync(payload.pdfFilePath)) {
    try {
      const bytes = fs.readFileSync(payload.pdfFilePath);
      const contentBytes = bytes.toString('base64');
      const attachmentPayload = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: String(payload.pdfFileName || 'Minutes-of-Meeting.pdf'),
        contentType: 'application/pdf',
        contentBytes
      };

      await axios.post(
        `${graphBaseUrl}/users/${mailboxUser}/messages/${encodeURIComponent(draftId)}/attachments`,
        attachmentPayload,
        {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          timeout: 45000
        }
      );
    } catch (error) {
      attachmentNote = resolveGraphErrorMessage(error);
    }
  } else {
    attachmentNote = 'Generated PDF file was not found on server; attachment step skipped.';
  }

  return {
    mode: 'graph-draft',
    draftId,
    outlookDraftWebUrl: webLink,
    attachmentAutoSupported: true,
    attachmentAdded: !attachmentNote,
    attachmentNote
  };
}

async function runGraphDiagnostics(config, options = {}) {
  const writeProbe = String(options.writeProbe || '').toLowerCase() === 'true' || options.writeProbe === true;
  const graph = config?.graph || {};
  const graphBaseUrl = normalizeBaseUrl(graph.baseUrl || 'https://graph.microsoft.com/v1.0');
  const mailboxUserRaw = String(graph.mailboxUser || '').trim();
  const mailboxUser = encodeURIComponent(mailboxUserRaw);
  const diagnostics = {
    timestamp: new Date().toISOString(),
    configuration: {
      enabled: Boolean(graph.enabled),
      isConfigured: isGraphDraftConfigured(config),
      authorityHost: normalizeBaseUrl(graph.authorityHost || ''),
      baseUrl: graphBaseUrl,
      scope: String(graph.scope || '').trim(),
      tenantIdConfigured: Boolean(graph.tenantId),
      tenantIdPreview: maskValue(graph.tenantId),
      clientIdConfigured: Boolean(graph.clientId),
      clientIdPreview: maskValue(graph.clientId),
      clientSecretConfigured: Boolean(graph.clientSecret),
      mailboxUserConfigured: Boolean(mailboxUserRaw),
      mailboxUser: mailboxUserRaw
    },
    token: {
      ok: false,
      expiresAt: '',
      audience: '',
      tenant: '',
      appId: '',
      roles: [],
      scopes: [],
      error: null
    },
    permissionCheck: {
      ok: false,
      expected: ['Mail.ReadWrite'],
      foundRoles: [],
      foundScopes: []
    },
    mailboxCheck: {
      ok: false,
      mailboxUser: mailboxUserRaw,
      userId: '',
      displayName: '',
      mail: '',
      userPrincipalName: '',
      error: null
    },
    draftProbe: {
      performed: writeProbe,
      ok: false,
      createdMessageId: '',
      deleted: false,
      error: null
    },
    hints: [],
    summary: {
      status: 'needs_attention',
      readyForServerDrafts: false
    }
  };

  if (!diagnostics.configuration.enabled) {
    diagnostics.hints.push('Set MS_GRAPH_ENABLED=true in deployment environment variables.');
  }
  if (!diagnostics.configuration.tenantIdConfigured) {
    diagnostics.hints.push('Set MS_GRAPH_TENANT_ID to your Microsoft Entra tenant ID.');
  }
  if (!diagnostics.configuration.clientIdConfigured) {
    diagnostics.hints.push('Set MS_GRAPH_CLIENT_ID to your app registration client ID.');
  }
  if (!diagnostics.configuration.clientSecretConfigured) {
    diagnostics.hints.push('Set MS_GRAPH_CLIENT_SECRET to an active app secret value.');
  }
  if (!diagnostics.configuration.mailboxUserConfigured) {
    diagnostics.hints.push('Set MS_GRAPH_MAILBOX_USER to the mailbox address used for draft creation.');
  }

  if (!diagnostics.configuration.isConfigured) {
    return diagnostics;
  }

  let token = '';
  try {
    token = await getGraphAccessToken(config);
    const claims = decodeJwtClaims(token);
    const expEpoch = Number.parseInt(String(claims.exp || '0'), 10);
    const roles = Array.isArray(claims.roles) ? claims.roles.map((entry) => String(entry)) : [];
    const scopes = String(claims.scp || '')
      .split(/\s+/)
      .map((entry) => entry.trim())
      .filter(Boolean);

    diagnostics.token.ok = true;
    diagnostics.token.expiresAt = Number.isFinite(expEpoch) && expEpoch > 0 ? new Date(expEpoch * 1000).toISOString() : '';
    diagnostics.token.audience = String(claims.aud || '');
    diagnostics.token.tenant = String(claims.tid || '');
    diagnostics.token.appId = String(claims.appid || '');
    diagnostics.token.roles = roles;
    diagnostics.token.scopes = scopes;
    diagnostics.permissionCheck.foundRoles = roles;
    diagnostics.permissionCheck.foundScopes = scopes;

    diagnostics.permissionCheck.ok =
      roles.includes('Mail.ReadWrite') || scopes.includes('Mail.ReadWrite') || scopes.includes('Mail.ReadWrite.Shared');

    if (!diagnostics.permissionCheck.ok) {
      diagnostics.hints.push(
        'Grant Microsoft Graph Mail.ReadWrite permission to the app registration and provide admin consent.'
      );
    }
  } catch (error) {
    const parsed = parseGraphError(error);
    diagnostics.token.error = parsed;
    diagnostics.hints.push(
      'Token request failed. Verify tenant/client credentials and ensure app secret is current in Render env vars.'
    );
    if (parsed.status === 401) {
      diagnostics.hints.push('401 from token endpoint: client secret may be invalid/expired or tenant/client mismatch.');
    }
    return diagnostics;
  }

  try {
    const mailboxResponse = await axios.get(`${graphBaseUrl}/users/${mailboxUser}?$select=id,displayName,mail,userPrincipalName`, {
      headers: {
        Authorization: `Bearer ${token}`
      },
      timeout: 30000
    });

    diagnostics.mailboxCheck.ok = true;
    diagnostics.mailboxCheck.userId = String(mailboxResponse?.data?.id || '');
    diagnostics.mailboxCheck.displayName = String(mailboxResponse?.data?.displayName || '');
    diagnostics.mailboxCheck.mail = String(mailboxResponse?.data?.mail || '');
    diagnostics.mailboxCheck.userPrincipalName = String(mailboxResponse?.data?.userPrincipalName || '');
  } catch (error) {
    const parsed = parseGraphError(error);
    diagnostics.mailboxCheck.error = parsed;
    diagnostics.hints.push('Mailbox lookup failed. Ensure MS_GRAPH_MAILBOX_USER exists and app has permission to access it.');
    if (parsed.status === 403) {
      diagnostics.hints.push('403 indicates missing Graph API application permissions or missing admin consent.');
    }
    return diagnostics;
  }

  if (writeProbe) {
    let probeId = '';
    try {
      const probeResponse = await axios.post(
        `${graphBaseUrl}/users/${mailboxUser}/messages`,
        {
          subject: `M.O.M Graph Diagnostics Probe ${new Date().toISOString()}`,
          body: {
            contentType: 'Text',
            content: 'Diagnostics probe message. This message will be deleted automatically.'
          }
        },
        {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          timeout: 30000
        }
      );
      probeId = String(probeResponse?.data?.id || '').trim();
      diagnostics.draftProbe.createdMessageId = probeId;
      diagnostics.draftProbe.ok = Boolean(probeId);

      if (probeId) {
        await axios.delete(`${graphBaseUrl}/users/${mailboxUser}/messages/${encodeURIComponent(probeId)}`, {
          headers: {
            Authorization: `Bearer ${token}`
          },
          timeout: 30000
        });
        diagnostics.draftProbe.deleted = true;
      }
    } catch (error) {
      diagnostics.draftProbe.ok = false;
      diagnostics.draftProbe.error = parseGraphError(error);
      diagnostics.hints.push('Draft probe failed. Confirm Mail.ReadWrite application permission and mailbox access.');
      return diagnostics;
    }
  }

  diagnostics.summary.readyForServerDrafts = Boolean(
    diagnostics.configuration.isConfigured &&
      diagnostics.token.ok &&
      diagnostics.permissionCheck.ok &&
      diagnostics.mailboxCheck.ok &&
      (!writeProbe || diagnostics.draftProbe.ok)
  );
  diagnostics.summary.status = diagnostics.summary.readyForServerDrafts ? 'healthy' : 'needs_attention';

  if (!diagnostics.summary.readyForServerDrafts && diagnostics.hints.length === 0) {
    diagnostics.hints.push('Graph diagnostics completed with issues. Review token, permission, and mailbox sections.');
  }

  return diagnostics;
}

module.exports = {
  isGraphDraftConfigured,
  buildGraphBodyHtml,
  createGraphDraft,
  runGraphDiagnostics
};
